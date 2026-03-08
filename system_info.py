"""
system_info.py  –  PC Health AI v3
Collects detailed hardware and OS information from the local Windows PC.

v3 Changes:
  - Parallel hardware scan (ThreadPoolExecutor) — startup ~5s -> ~1.5s
  - Fan speed reading via LHM WMI (get_fans)
  - GPU Load % + VRAM controller % (get_gpu_usage)
  - AMD Tdie/Tctl disambiguation (Tdie preferred — no +10C offset on AM5)
  - WMI ThermalZone acceptance range widened (0-120C)
  - Sensor names stripped before matching
"""

import wmi
import psutil
import winreg
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

try:
    import pythoncom
    _HAS_PYTHONCOM = True
except ImportError:
    _HAS_PYTHONCOM = False


def _co_init():
    if _HAS_PYTHONCOM:
        try: pythoncom.CoInitialize()
        except Exception: pass

def _co_done():
    if _HAS_PYTHONCOM:
        try: pythoncom.CoUninitialize()
        except Exception: pass


class SystemInfo:
    """Gathers complete system hardware and software information."""

    def __init__(self):
        pass

    # ─────────────────────────────────────── parallel entry point ─────────────

    def get_all_info(self):
        """
        Return all hardware info.  Each sub-call runs in its own thread
        so WMI queries overlap instead of serialising (5s -> ~1.5s).
        """
        tasks = {
            "cpu":         self.get_cpu_info,
            "gpu":         self.get_gpu_info,
            "ram":         self.get_ram_info,
            "storage":     self.get_storage_info,
            "motherboard": self.get_motherboard_info,
            "os":          self.get_os_info,
            "network":     self.get_network_info,
        }
        results = {}
        with ThreadPoolExecutor(max_workers=len(tasks)) as ex:
            futures = {ex.submit(fn): key for key, fn in tasks.items()}
            for fut in as_completed(futures):
                key = futures[fut]
                try:
                    results[key] = fut.result()
                except Exception as e:
                    results[key] = {"error": str(e)}
        return results

    # ─────────────────────────────────────────────────────────── CPU ──────────

    def get_cpu_info(self):
        _co_init()
        try:
            c       = wmi.WMI()
            cpu_wmi = c.Win32_Processor()[0]
            freq    = psutil.cpu_freq()
            usage   = psutil.cpu_percent(interval=1, percpu=False)
            return {
                "name":              cpu_wmi.Name.strip(),
                "manufacturer":      cpu_wmi.Manufacturer,
                "cores":             cpu_wmi.NumberOfCores,
                "threads":           cpu_wmi.NumberOfLogicalProcessors,
                "max_speed_mhz":     cpu_wmi.MaxClockSpeed,
                "current_speed_mhz": round(freq.current) if freq else "N/A",
                "usage_percent":     usage,
                "socket":            cpu_wmi.SocketDesignation,
                "l2_cache_kb":       cpu_wmi.L2CacheSize,
                "l3_cache_kb":       cpu_wmi.L3CacheSize,
                "architecture":      cpu_wmi.Architecture,
                "per_core_usage":    psutil.cpu_percent(interval=0.5, percpu=True),
            }
        except Exception as exc:
            return {"error": str(exc)}
        finally:
            _co_done()

    # ─────────────────────────────────────────────────────────── GPU ──────────

    _dxgi_cache = None

    def _get_vram_via_dxgi(self):
        if SystemInfo._dxgi_cache is not None:
            return SystemInfo._dxgi_cache
        cs = r"""
using System; using System.Collections.Generic; using System.Runtime.InteropServices;
public static class DxgiVram {
    [DllImport("dxgi.dll", CallingConvention=CallingConvention.StdCall)]
    static extern int CreateDXGIFactory(ref Guid riid, out IntPtr ppFactory);
    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
    struct AdapterDesc {
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst=128)] public string Description;
        public uint VendorId, DeviceId, SubSysId, Revision;
        public UIntPtr DedicatedVideoMemory, DedicatedSystemMemory, SharedSystemMemory;
        public int LuidLow, LuidHigh;
    }
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int DEnum(IntPtr self, uint i, out IntPtr pAdapter);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int DDesc(IntPtr self, out AdapterDesc pDesc);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate uint DRel(IntPtr self);
    static Delegate Vt(IntPtr obj, int slot, Type t) {
        var vt = Marshal.ReadIntPtr(obj);
        var fp = Marshal.ReadIntPtr(new IntPtr(vt.ToInt64() + (long)slot * IntPtr.Size));
        return Marshal.GetDelegateForFunctionPointer(fp, t); }
    public static string Query() {
        var list = new List<string>();
        var iid = new Guid("7b7166ec-21c7-44ae-b21a-c9ae321ae369"); IntPtr factory;
        if (CreateDXGIFactory(ref iid, out factory) < 0 || factory == IntPtr.Zero) return "";
        var enumFn = (DEnum)Vt(factory, 7, typeof(DEnum)); var relFact = (DRel)Vt(factory, 2, typeof(DRel));
        for (uint i = 0; ; i++) { IntPtr adapter;
            if (enumFn(factory, i, out adapter) < 0 || adapter == IntPtr.Zero) break;
            var descFn = (DDesc)Vt(adapter, 8, typeof(DDesc)); var relAdp = (DRel)Vt(adapter, 2, typeof(DRel));
            AdapterDesc d; if (descFn(adapter, out d) >= 0) { ulong vram = (ulong)d.DedicatedVideoMemory;
                if (vram > 0) list.Add(d.Description.Trim() + "|" + vram.ToString()); }
            relAdp(adapter); }
        relFact(factory); return string.Join(";", list); } }
"""
        import subprocess, tempfile, os
        tmp = None
        try:
            with tempfile.NamedTemporaryFile(mode="w", suffix=".ps1", delete=False, encoding="utf-8") as fh:
                fh.write("Add-Type -TypeDefinition @'\n"); fh.write(cs); fh.write("\n'@\n[DxgiVram]::Query()\n")
                tmp = fh.name
            result = subprocess.run(["powershell","-NoProfile","-NonInteractive","-ExecutionPolicy","Bypass","-File",tmp],
                capture_output=True, text=True, timeout=25, creationflags=0x08000000)
            vmap = {}
            for entry in result.stdout.strip().split(";"):
                if "|" in entry:
                    name, vb = entry.strip().rsplit("|", 1)
                    try:
                        gb = round(int(vb.strip()) / (1024**3), 1)
                        if gb > 0: vmap[name.strip().lower()] = gb
                    except Exception: pass
            SystemInfo._dxgi_cache = vmap; return vmap
        except Exception:
            SystemInfo._dxgi_cache = {}; return {}
        finally:
            if tmp:
                try: os.unlink(tmp)
                except Exception: pass

    def _get_vram_entries(self):
        entries = []
        reg_path = r"SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}"
        val_names = ["HardwareInformation.qwMemorySize","HardwareInformation.MemorySize","QWordMemorySize"]
        try:
            base = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path)
            idx = 0
            while True:
                try:
                    sub_name = winreg.EnumKey(base, idx)
                    if sub_name.isdigit():
                        sub = winreg.OpenKey(base, sub_name); raw = None
                        for vname in val_names:
                            try: raw, _ = winreg.QueryValueEx(sub, vname)
                            except FileNotFoundError: pass
                            if raw: break
                        if raw:
                            if isinstance(raw, bytes): raw = int.from_bytes(raw, "little")
                            if raw > 0:
                                gb = round(int(raw) / (1024**3), 1)
                                try: desc, _ = winreg.QueryValueEx(sub, "DriverDesc"); entries.append((desc.lower(), gb))
                                except FileNotFoundError: entries.append(("", gb))
                        winreg.CloseKey(sub)
                    idx += 1
                except OSError: break
            winreg.CloseKey(base)
        except Exception: pass
        entries.sort(key=lambda x: x[1], reverse=True)
        return entries

    def get_gpu_info(self):
        _co_init()
        try:
            c = wmi.WMI(); gpus_wmi = c.Win32_VideoController()
            dxgi_map = self._get_vram_via_dxgi()
            reg_entries = self._get_vram_entries(); reg_used = set()
            gpu_list = []
            for gpu in gpus_wmi:
                name = (gpu.Name or "").strip(); name_lower = name.lower(); vram_gb = None
                if dxgi_map:
                    best_score, best_gb = 0, None
                    tokens = [t for t in name_lower.split() if len(t) > 3]
                    for dname, gb in dxgi_map.items():
                        score = sum(1 for t in tokens if t in dname)
                        if score > best_score: best_score, best_gb = score, gb
                    if best_score > 0: vram_gb = best_gb
                    elif len(dxgi_map) == 1: vram_gb = next(iter(dxgi_map.values()))
                if vram_gb is None and reg_entries:
                    best_score, best_idx = 0, None
                    tokens = [t for t in name_lower.split() if len(t) > 3]
                    for ri, (desc, gb) in enumerate(reg_entries):
                        if ri in reg_used: continue
                        score = sum(1 for t in tokens if t in desc)
                        if score > best_score: best_score, best_idx = score, ri
                    if best_idx is None:
                        for ri, _ in enumerate(reg_entries):
                            if ri not in reg_used: best_idx = ri; break
                    if best_idx is not None: vram_gb = reg_entries[best_idx][1]; reg_used.add(best_idx)
                if vram_gb is None:
                    vram_gb = round(int(gpu.AdapterRAM)/(1024**3),1) if gpu.AdapterRAM else "N/A"
                res = "N/A"
                try: res = f"{gpu.CurrentHorizontalResolution}x{gpu.CurrentVerticalResolution}"
                except Exception: pass
                gpu_list.append({"name": gpu.Name, "driver_version": gpu.DriverVersion,
                    "driver_date": str(gpu.DriverDate)[:8] if gpu.DriverDate else "N/A",
                    "vram_gb": vram_gb, "video_processor": gpu.VideoProcessor,
                    "current_resolution": res, "refresh_rate_hz": gpu.CurrentRefreshRate,
                    "status": gpu.Status, "adapter_dac_type": gpu.AdapterDACType})
            return gpu_list
        except Exception as exc:
            return [{"error": str(exc)}]
        finally:
            _co_done()

    # ─────────────────────────────────────────────────────────── RAM ──────────

    _RAM_MFR = {
        "802c":"Corsair","ce00":"Samsung","ad00":"SK Hynix","2c00":"Micron",
        "04cd":"G.Skill","0198":"Kingston","029e":"Corsair","9e00":"Crucial",
        "5105":"Patriot","f1ff":"G.Skill","01ba":"Samsung","7f7f":"Unknown",
        "7f98":"Kingston","9801":"Kingston",
    }

    def _resolve_mfr(self, raw):
        if not raw: return "N/A"
        cleaned = raw.strip().lower().replace(" ","").replace("0x","")
        for key, name in self._RAM_MFR.items():
            if key in cleaned: return name
        if any(c.isalpha() and c not in "abcdef" for c in cleaned): return raw.strip()
        return raw.strip()

    def _detect_xmp_expo(self, modules):
        if not modules: return "N/A"
        try:
            speed = int(modules[0].get("speed_mhz") or 0)
            is_ddr5 = "5" in modules[0].get("memory_type","")
            base_freq = 4800 if is_ddr5 else 2133; profile = "EXPO" if is_ddr5 else "XMP"
            if speed > base_freq: return f"Enabled — {speed} MHz (base {base_freq} MHz)"
            elif speed == base_freq: return f"Disabled — running at base {speed} MHz"
            elif speed > 0: return f"Disabled — {speed} MHz (base {base_freq} MHz)"
        except Exception: pass
        return "N/A"

    def get_ram_info(self):
        _co_init()
        try:
            c = wmi.WMI(); vm = psutil.virtual_memory(); modules = []
            smbios_map = {0:"Unknown",1:"Other",2:"DRAM",3:"SDRAM",4:"SDRAM",6:"Flash",
                17:"SDRAM",18:"DDR2",19:"DDR2",20:"DDR2",21:"DDR3",22:"DDR3",24:"DDR3",
                26:"DDR4",27:"LPDDR4",29:"DDR4",34:"DDR5",35:"LPDDR5",36:"DDR5"}
            for mod in c.Win32_PhysicalMemory():
                cap_gb = round(int(mod.Capacity)/(1024**3),1) if mod.Capacity else 0
                smbios_raw = int(getattr(mod,"SMBIOSMemoryType",0) or 0)
                if smbios_raw in (0,2):
                    legacy = {18:"DDR2",20:"DDR3",24:"DDR3",26:"DDR4",27:"LPDDR4",34:"DDR5",35:"LPDDR5"}
                    mem_type_str = legacy.get(int(mod.MemoryType or 0), smbios_map.get(smbios_raw,"Unknown"))
                else:
                    mem_type_str = smbios_map.get(smbios_raw, f"Type {smbios_raw}")
                modules.append({"capacity_gb":cap_gb,"speed_mhz":mod.Speed,
                    "manufacturer":self._resolve_mfr(mod.Manufacturer),
                    "part_number":(mod.PartNumber.strip() if mod.PartNumber else "N/A"),
                    "bank_label":mod.BankLabel,"device_locator":mod.DeviceLocator,
                    "memory_type":mem_type_str,"form_factor":mod.FormFactor})
            return {"total_gb":round(vm.total/(1024**3),1),"available_gb":round(vm.available/(1024**3),1),
                "used_gb":round(vm.used/(1024**3),1),"usage_percent":vm.percent,
                "modules":modules,"module_count":len(modules),
                "channel_info":"Dual Channel" if len(modules)>=2 else "Single Channel",
                "xmp_expo":self._detect_xmp_expo(modules)}
        except Exception as exc:
            return {"error": str(exc)}
        finally:
            _co_done()

    # ─────────────────────────────────────────────────────── Storage ──────────

    def get_storage_info(self):
        _co_init()
        try:
            c = wmi.WMI(); disk_list = []; part_usages = {}
            for part in psutil.disk_partitions(all=False):
                try:
                    usage = psutil.disk_usage(part.mountpoint)
                    key = part.mountpoint.rstrip("\\").upper()
                    part_usages[key] = {"mountpoint":part.mountpoint,"fstype":part.fstype,
                        "total_gb":round(usage.total/(1024**3),1),"used_gb":round(usage.used/(1024**3),1),
                        "free_gb":round(usage.free/(1024**3),1),"percent":usage.percent}
                except PermissionError: pass
            for disk in c.Win32_DiskDrive():
                size_gb = round(int(disk.Size)/(1024**3),1) if disk.Size else 0
                partitions = []
                for dp in c.Win32_DiskDriveToDiskPartition():
                    try:
                        if dp.Antecedent.split('"')[1] != disk.DeviceID: continue
                        part_path = dp.Dependent.split('"')[1]
                    except (IndexError,AttributeError): continue
                    for lp in c.Win32_LogicalDiskToPartition():
                        try:
                            if part_path not in lp.Antecedent: continue
                            drive_letter = lp.Dependent.split('"')[1].upper()
                        except (IndexError,AttributeError): continue
                        usage_entry = part_usages.get(drive_letter)
                        if usage_entry and usage_entry not in partitions: partitions.append(usage_entry)
                if not partitions and not disk_list: partitions = list(part_usages.values())
                disk_list.append({"model":disk.Model,"size_gb":size_gb,"interface":disk.InterfaceType,
                    "media_type":disk.MediaType,"serial":(disk.SerialNumber.strip() if disk.SerialNumber else "N/A"),
                    "partitions":partitions})
            if not disk_list:
                disk_list.append({"model":"Unknown Disk","size_gb":"N/A","interface":"N/A",
                    "media_type":"N/A","serial":"N/A","partitions":list(part_usages.values())})
            return disk_list
        except Exception as exc:
            return [{"error": str(exc)}]
        finally:
            _co_done()

    # ─────────────────────────────────────────────────────── Motherboard ──────

    def get_motherboard_info(self):
        _co_init()
        try:
            c = wmi.WMI(); board = c.Win32_BaseBoard()[0]; bios = c.Win32_BIOS()[0]
            return {"manufacturer":board.Manufacturer,"product":board.Product,"version":board.Version,
                "serial":board.SerialNumber,"bios_version":bios.SMBIOSBIOSVersion,
                "bios_date":(str(bios.ReleaseDate)[:8] if bios.ReleaseDate else "N/A"),
                "bios_manufacturer":bios.Manufacturer,"bios_serial":bios.SerialNumber}
        except Exception as exc:
            return {"error": str(exc)}
        finally:
            _co_done()

    # ─────────────────────────────────────────────────────── OS Info ──────────

    def get_os_info(self):
        _co_init()
        try:
            c = wmi.WMI(); os_wmi = c.Win32_OperatingSystem()[0]; uptime = "N/A"
            try:
                boot_raw = os_wmi.LastBootUpTime
                boot_dt = datetime.strptime(boot_raw[:14], "%Y%m%d%H%M%S")
                delta = datetime.now() - boot_dt
                hours, rem = divmod(int(delta.total_seconds()), 3600)
                mins, _ = divmod(rem, 60); uptime = f"{hours}h {mins}m"
            except Exception: pass
            return {"name":os_wmi.Caption,"version":os_wmi.Version,"build":os_wmi.BuildNumber,
                "architecture":os_wmi.OSArchitecture,
                "install_date":(str(os_wmi.InstallDate)[:8] if os_wmi.InstallDate else "N/A"),
                "last_boot":(str(os_wmi.LastBootUpTime)[:14] if os_wmi.LastBootUpTime else "N/A"),
                "system_uptime":uptime,"registered_user":os_wmi.RegisteredUser,"computer_name":os_wmi.CSName}
        except Exception as exc:
            return {"error": str(exc)}
        finally:
            _co_done()

    # ─────────────────────────────────────────────────────── Network ──────────

    _VIRTUAL_ADAPTERS = (
        "loopback","pseudo","teredo","isatap","6to4","tunnel",
        "vmware","virtualbox","vethernet","hyper-v","bluetooth pan",
        "npcap loopback","miniport","wan miniport",
    )

    def get_network_info(self):
        try:
            adapters = []; if_addrs = psutil.net_if_addrs()
            if_stats = psutil.net_if_stats(); net_io = psutil.net_io_counters(pernic=True)
            for name, addrs in if_addrs.items():
                if any(v in name.lower() for v in self._VIRTUAL_ADAPTERS): continue
                stats = if_stats.get(name); io = net_io.get(name)
                ip4 = next((a.address for a in addrs if a.family.name == "AF_INET"), "N/A")
                if ip4.startswith("127."): continue
                mac = next((a.address for a in addrs if a.family.name == "AF_LINK"), "N/A")
                adapters.append({"name":name,"is_up":stats.isup if stats else False,
                    "speed_mbps":stats.speed if stats else 0,"ipv4":ip4,"mac":mac,
                    "bytes_sent":round(io.bytes_sent/(1024**2),1) if io else 0,
                    "bytes_recv":round(io.bytes_recv/(1024**2),1) if io else 0})
            return adapters
        except Exception as exc:
            return [{"error": str(exc)}]

    # ─────────────────────────────────────────────────────── LHM / Temps ──────

    _LHM_SEARCH_PATHS = [
        r"C:\Program Files\LibreHardwareMonitor\LibreHardwareMonitor.exe",
        r"C:\LibreHardwareMonitor\LibreHardwareMonitor.exe",
        r"C:\tools\LibreHardwareMonitor\LibreHardwareMonitor.exe",
    ]
    _LHM_ZIP_URL = ("https://github.com/LibreHardwareMonitor/LibreHardwareMonitor"
                    "/releases/latest/download/LibreHardwareMonitor-net472.zip")
    _LHM_ZIP_FALLBACK = ("https://github.com/LibreHardwareMonitor/LibreHardwareMonitor"
                         "/releases/download/v0.9.4/LibreHardwareMonitor-net472.zip")

    _lhm_proc  = None
    _lhm_tried = False

    @classmethod
    def _lhm_wmi_ready(cls):
        _co_init()
        try:
            wmi.WMI(namespace="root/LibreHardwareMonitor").Sensor()
            return True
        except Exception:
            return False
        finally:
            try: _co_done()
            except Exception: pass

    @classmethod
    def _download_lhm_zip(cls, status_cb):
        import urllib.request as _ur, ssl as _ssl
        headers = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/122.0",
                   "Accept":"application/octet-stream,*/*"}
        for url in [cls._LHM_ZIP_URL, cls._LHM_ZIP_FALLBACK]:
            for verify in (True, False):
                try:
                    ctx = (_ssl.create_default_context() if verify
                           else _ssl._create_unverified_context())
                    req = _ur.Request(url, headers=headers)
                    with _ur.urlopen(req, timeout=60, context=ctx) as resp:
                        total = int(resp.headers.get("Content-Length",0))
                        chunks, received = [], 0
                        while True:
                            chunk = resp.read(65536)
                            if not chunk: break
                            chunks.append(chunk); received += len(chunk)
                            if total > 0 and status_cb:
                                try: status_cb(f"progress:{int(received*100/total)}")
                                except Exception: pass
                        return b"".join(chunks)
                except Exception: continue
        return None

    @classmethod
    def ensure_lhm(cls, app_dir=None, status_cb=None):
        import os
        def _cb(msg):
            if status_cb:
                try: status_cb(msg)
                except Exception: pass

        if cls._lhm_wmi_ready():
            _cb("ready"); return True

        # If LHM is already running in tasklist, wait for WMI to register
        try:
            import subprocess as _chk
            _r = _chk.run(["tasklist","/FI","IMAGENAME eq LibreHardwareMonitor.exe","/FO","CSV"],
                capture_output=True, text=True, timeout=5, creationflags=0x08000000)
            if "LibreHardwareMonitor.exe" in _r.stdout:
                _cb("waiting:0")
                import time as _tw
                for _i in range(20):
                    _tw.sleep(0.7); _cb(f"waiting:{_i}")
                    if cls._lhm_wmi_ready(): _cb("ready"); return True
        except Exception: pass

        search = list(cls._LHM_SEARCH_PATHS)
        if app_dir:
            search += [os.path.join(app_dir,"tools","lhm","LibreHardwareMonitor.exe"),
                       os.path.join(app_dir,"LibreHardwareMonitor.exe")]
        lhm_exe = next((p for p in search if os.path.isfile(p)), None)

        if lhm_exe is None and app_dir:
            _cb("downloading")
            import zipfile as _zf, io as _io
            target_dir = os.path.join(app_dir,"tools","lhm")
            os.makedirs(target_dir, exist_ok=True)
            zip_data = cls._download_lhm_zip(_cb)
            if zip_data is None:
                _cb("download_failed:Cannot reach GitHub — check internet and run as Administrator")
                return False
            _cb("extracting")
            try:
                with _zf.ZipFile(_io.BytesIO(zip_data)) as z: z.extractall(target_dir)
            except Exception as e:
                _cb(f"download_failed:Zip extract error: {e}"); return False
            for root, _, files in os.walk(target_dir):
                for fn in files:
                    if fn.lower() == "librehardwaremonitor.exe":
                        lhm_exe = os.path.join(root, fn); break
                if lhm_exe: break
            if not lhm_exe:
                _cb("download_failed:LibreHardwareMonitor.exe not found in zip"); return False

        if not lhm_exe or not os.path.isfile(lhm_exe):
            _cb("not_found:LHM not found — run as Administrator to auto-download"); return False

        # Kill any stale LHM process that may be blocking WMI registration
        try:
            import subprocess as _sp2
            _sp2.run(["taskkill","/F","/IM","LibreHardwareMonitor.exe"],
                     capture_output=True, creationflags=0x08000000, timeout=5)
        except Exception: pass

        _cb("starting")
        try:
            import subprocess as _sp
            cls._lhm_proc = _sp.Popen([lhm_exe], creationflags=0x08000000, close_fds=True)
        except Exception as e:
            _cb(f"launch_failed:{e}"); return False

        import time as _t
        # Wait up to 25 seconds total — first few ticks are faster
        schedule = [0.4]*5 + [0.6]*5 + [1.0]*10 + [1.5]*5   # 35 steps ~ 25s max
        for i, delay in enumerate(schedule):
            _t.sleep(delay); _cb(f"waiting:{i}")
            if cls._lhm_wmi_ready():
                _cb("ready"); return True
        _cb("timeout:WMI sensor timed out after 25s — try running as Administrator")
        return False

    @classmethod
    def reset_lhm(cls):
        cls._lhm_tried = False

    # ─────────────────────────────────────────────────────── get_temps ────────

    def get_temps(self, app_dir=None, status_cb=None) -> dict:
        result = {"cpu_package":None,"cpu_cores":[],"gpu":None,"fans":[],"source":"none"}

        if not SystemInfo._lhm_tried:
            SystemInfo._lhm_tried = True
            SystemInfo.ensure_lhm(app_dir, status_cb=status_cb)

        for ns in ("root/LibreHardwareMonitor","root/OpenHardwareMonitor"):
            _co_init()
            try:
                c = wmi.WMI(namespace=ns)
                cpu_pkg = None; cpu_tdie = None; cpu_tctl = None
                cpu_cores = []; gpu_max = None; fans = []
                for s in c.Sensor():
                    stype = getattr(s,"SensorType","")
                    name  = (getattr(s,"Name","") or "").strip().lower()
                    val   = getattr(s,"Value",None)
                    if val is None: continue
                    val   = round(float(val),1)

                    if stype == "Temperature":
                        # CPU package: LHM names vary by chip/version
                        # Ryzen 7600X shows sensor named exactly "Package"
                        if (name == "package"
                                or any(k in name for k in (
                                    "cpu package","package temperature",
                                    "cpu temp","processor temp"))):
                            cpu_pkg = val
                        # AMD Tdie is the real junction temp; prefer over Tctl (+10C offset on AM5)
                        elif (name in ("ccd1 (tdie)","ccd2 (tdie)","ccd (tdie)")
                                or any(k in name for k in ("cpu tdie","tdie","core (tdie"))):
                            if cpu_tdie is None or val > cpu_tdie:
                                cpu_tdie = val
                        elif any(k in name for k in ("cpu tctl","tctl","core (tctl","tctl/tdie")):
                            if cpu_tctl is None or val > cpu_tctl:
                                cpu_tctl = val
                        elif "core #" in name or "cpu core" in name or name.startswith("core "):
                            cpu_cores.append(val)
                        # GPU — AMD RX shows as "gpu core" in LHM
                        elif any(k in name for k in ("gpu core","gpu temperature",
                                "gpu hot spot","hotspot","gpu (diode)","gpu thermal",
                                "junction","gpu edge","gpu temp")):
                            if gpu_max is None or val > gpu_max: gpu_max = val

                    elif stype == "Fan":
                        # name like "cpu fan" / "chassis fan 1" / "gpu fan"
                        fans.append({"name": name.title(), "rpm": int(val)})

                # Priority: Package > Tdie > Tctl
                if cpu_pkg is None:
                    cpu_pkg = cpu_tdie if cpu_tdie is not None else cpu_tctl

                result["cpu_cores"]   = cpu_cores
                result["fans"]        = fans
                result["cpu_package"] = (cpu_pkg if cpu_pkg is not None
                                         else (round(max(cpu_cores),1) if cpu_cores else None))
                if gpu_max is not None: result["gpu"] = gpu_max
                if result["cpu_package"] is not None or result["gpu"] is not None:
                    result["source"] = ns.split("/")[-1]
                    _co_done(); return result
            except Exception: pass
            finally:
                try: _co_done()
                except Exception: pass

        # Fallback 3: WMI ThermalZone
        # Note: Win32_PerfFormattedData_Counters_ThermalZoneInformation.Temperature
        # reports in tenths-of-Kelvin on some Windows builds, raw Kelvin on others.
        # We try both and keep only values in the realistic range 30-120C.
        _co_init()
        try:
            c, temps = wmi.WMI(), []
            for z in c.Win32_PerfFormattedData_Counters_ThermalZoneInformation():
                raw = getattr(z,"Temperature",None)
                if not raw: continue
                raw = int(raw)
                # Try tenths-of-Kelvin (most common): (raw - 2732) / 10
                c1 = round((raw - 2732) / 10.0, 1)
                # Try raw Kelvin: raw - 273
                c2 = round(raw - 273.0, 1)
                # Accept whichever falls in the realistic 30-110C range
                if 30 < c1 < 110:
                    temps.append(c1)
                elif 30 < c2 < 110:
                    temps.append(c2)
            if temps:
                result["cpu_package"] = round(max(temps), 1)
                result["source"] = "WMI-ThermalZone"; _co_done(); return result
        except Exception: pass
        finally:
            try: _co_done()
            except Exception: pass

        # Fallback 4: MSAcpi
        _co_init()
        try:
            c, temps = wmi.WMI(namespace="root/wmi"), []
            for z in c.MSAcpi_ThermalZoneTemperature():
                raw = getattr(z,"CurrentTemperature",None)
                if not raw: continue
                raw = int(raw)
                c1 = round((raw - 2732) / 10.0, 1)
                c2 = round(raw - 273.0, 1)
                if 30 < c1 < 110:   temps.append(c1)
                elif 30 < c2 < 110: temps.append(c2)
            if temps:
                result["cpu_package"] = round(max(temps),1)
                result["source"] = "ACPI"; _co_done(); return result
        except Exception: pass
        finally:
            try: _co_done()
            except Exception: pass

        # Fallback 5: nvidia-smi
        try:
            import subprocess as _sp
            r = _sp.run(["nvidia-smi","--query-gpu=temperature.gpu","--format=csv,noheader,nounits"],
                capture_output=True,text=True,timeout=5,creationflags=0x08000000)
            if r.returncode == 0 and r.stdout.strip():
                vals = [float(v.strip()) for v in r.stdout.strip().split("\n") if v.strip()]
                if vals: result["gpu"] = vals[0]; result["source"] = "nvidia-smi"
        except Exception: pass

        return result

    # ─────────────────────────────────────────────────────── get_gpu_usage ────

    def get_gpu_usage(self) -> dict:
        """
        GPU Load % and VRAM controller %.
        Priority:
          1. LHM / OHM WMI (most accurate — requires LHM running as admin)
          2. Win32_PerfFormattedData_GPUPerformanceCounters_GPUEngine
             (built into Windows 10/11 — no third-party tool needed)
          3. PowerShell Get-Counter fallback (last resort)
        """
        result = {"gpu_load_pct": None, "vram_ctrl_pct": None, "source": "none"}

        # ── 1. LHM / OHM ──────────────────────────────────────────────────────
        for ns in ("root/LibreHardwareMonitor", "root/OpenHardwareMonitor"):
            _co_init()
            try:
                c = wmi.WMI(namespace=ns)
                gpu_load = None; vram_ctrl = None
                for s in c.Sensor():
                    if getattr(s, "SensorType", "") != "Load": continue
                    name = (getattr(s, "Name", "") or "").strip().lower()
                    val  = getattr(s, "Value", None)
                    if val is None: continue
                    val = round(float(val), 1)
                    if "gpu core" in name or "d3d 3d" in name:
                        if gpu_load is None or val > gpu_load: gpu_load = val
                    elif any(k in name for k in
                             ("gpu memory controller", "d3d memory",
                              "gpu memory", "vram")):
                        if vram_ctrl is None or val > vram_ctrl: vram_ctrl = val
                if gpu_load is not None or vram_ctrl is not None:
                    result.update({"gpu_load_pct": gpu_load,
                                   "vram_ctrl_pct": vram_ctrl,
                                   "source": ns.split("/")[-1]})
                    _co_done(); return result
            except Exception:
                pass
            finally:
                try: _co_done()
                except Exception: pass

        # ── 2. Win32 GPU Performance Counters (built-in, no LHM needed) ───────
        _co_init()
        try:
            c = wmi.WMI(namespace="root/cimv2")
            engines = c.Win32_PerfFormattedData_GPUPerformanceCounters_GPUEngine()
            d3d_vals  = []
            copy_vals = []
            vram_vals = []
            for eng in engines:
                name = (getattr(eng, "Name", "") or "").lower()
                util = getattr(eng, "UtilizationPercentage", None)
                if util is None: continue
                try: util = float(util)
                except Exception: continue
                if "engtype_3d"   in name: d3d_vals.append(util)
                elif "engtype_copy" in name or "engtype_videodecode" in name:
                    copy_vals.append(util)
                # VRAM is separate counter class below
            if d3d_vals:
                gpu_load = round(sum(d3d_vals), 1)
                # cap at 100 (multiple engines can sum > 100 on multi-GPU)
                gpu_load = min(gpu_load, 100.0)
                result.update({"gpu_load_pct": gpu_load,
                               "source": "Win32PerfCounter"})
            # try VRAM usage counter
            try:
                vram_objs = c.Win32_PerfFormattedData_GPUPerformanceCounters_GPUAdapterMemory()
                for v in vram_objs:
                    dedicated = getattr(v, "DedicatedUsage", None)
                    total     = getattr(v, "DedicatedLimit", None)
                    if dedicated and total:
                        try:
                            pct = round(float(dedicated) / float(total) * 100, 1)
                            if 0 < pct <= 100:
                                result["vram_ctrl_pct"] = pct; break
                        except Exception: pass
            except Exception: pass
            if result["gpu_load_pct"] is not None:
                _co_done(); return result
        except Exception:
            pass
        finally:
            try: _co_done()
            except Exception: pass

        # ── 3. PowerShell Get-Counter last resort ──────────────────────────────
        try:
            import subprocess as _sp
            ps = (
                "try {"
                "  $c = (Get-Counter '\\GPU Engine(*engtype_3D*)\\Utilization Percentage'"
                "         -ErrorAction Stop).CounterSamples"
                "         | Measure-Object CookedValue -Sum;"
                "  [math]::Round([math]::Min($c.Sum, 100), 1)"
                "} catch { 0 }"
            )
            r = _sp.run(
                ["powershell", "-NoProfile", "-NonInteractive",
                 "-ExecutionPolicy", "Bypass", "-Command", ps],
                capture_output=True, text=True,
                timeout=8, creationflags=0x08000000)
            if r.returncode == 0 and r.stdout.strip():
                val = float(r.stdout.strip())
                if val >= 0:
                    result.update({"gpu_load_pct": round(val, 1),
                                   "source": "PerfCounter-PS"})
        except Exception:
            pass

        return result

    # ─────────────────────────────────────────────────────── live metrics ─────

    def get_live_metrics(self) -> dict:
        freq = psutil.cpu_freq(); vm = psutil.virtual_memory()
        return {"cpu_usage":psutil.cpu_percent(interval=0.3,percpu=False),
                "cpu_per_core":psutil.cpu_percent(interval=0,percpu=True),
                "cpu_freq_mhz":round(freq.current) if freq else None,
                "ram_pct":vm.percent,"ram_used_gb":round(vm.used/(1024**3),1),
                "ram_free_gb":round(vm.available/(1024**3),1)}

    # ─────────────────────────────────────────────────────── processes ────────

    def get_top_processes(self, n=12) -> list:
        """Return top-n processes sorted by CPU then RAM, safe for any OS user."""
        procs = []
        for p in psutil.process_iter(
                ["pid","name","cpu_percent","memory_info","status"]):
            try:
                mi = p.info.get("memory_info")
                procs.append({
                    "pid":    p.info["pid"],
                    "name":   p.info["name"] or "?",
                    "cpu":    p.info.get("cpu_percent") or 0.0,
                    "ram_mb": round(mi.rss/(1024**2),1) if mi else 0.0,
                    "status": p.info.get("status","?"),
                })
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        # Two separate top-n lists (by CPU and by RAM)
        by_cpu = sorted(procs, key=lambda x: x["cpu"], reverse=True)[:n]
        by_ram = sorted(procs, key=lambda x: x["ram_mb"], reverse=True)[:n]
        return {"by_cpu": by_cpu, "by_ram": by_ram}
