"""
diagnosis_engine.py
Runs a battery of PC health checks and returns structured findings.
Each finding is a dict:
  {
    id           : str              – unique key
    category     : str              – group label
    icon         : str              – emoji
    title        : str              – short headline
    status       : "ok"|"info"|"warning"|"critical"
    detail       : str              – explanation + advice
    action_label : str | None       – button text
    action_type  : "url"|"winget"|"command"|"store"|None
    action_value : str | None       – URL, package ID, or command
  }
"""

import subprocess
import json
import re
import winreg
from datetime import datetime

CREATE_NO_WINDOW = 0x08000000


# ─────────────────────────────────────────────────── low-level helpers ────────

def _run(cmd, timeout=30):
    try:
        r = subprocess.run(cmd, capture_output=True, text=True,
                           encoding="utf-8", errors="replace",
                           creationflags=CREATE_NO_WINDOW, timeout=timeout)
        return r.stdout, r.stderr, r.returncode
    except subprocess.TimeoutExpired:
        return "", "timeout", -1
    except Exception as e:
        return "", str(e), -1


def _run_ps(script, timeout=20):
    return _run(["powershell", "-ExecutionPolicy", "Bypass",
                 "-NonInteractive", "-Command", script], timeout=timeout)


def _reg_exists(hive, path):
    try:
        winreg.CloseKey(winreg.OpenKey(hive, path))
        return True
    except Exception:
        return False


def _reg_get(hive, path, value):
    try:
        key  = winreg.OpenKey(hive, path)
        data, _ = winreg.QueryValueEx(key, value)
        winreg.CloseKey(key)
        return data
    except Exception:
        return None


def _ok(id_, cat, icon, title, detail):
    return {"id": id_, "category": cat, "icon": icon, "title": title,
            "status": "ok", "detail": detail,
            "action_label": None, "action_type": None, "action_value": None}


def _finding(id_, cat, icon, title, status, detail,
             action_label=None, action_type=None, action_value=None):
    return {"id": id_, "category": cat, "icon": icon, "title": title,
            "status": status, "detail": detail,
            "action_label": action_label,
            "action_type": action_type,
            "action_value": action_value}


# ══════════════════════════════════════════════════════════════════════════════
class DiagnosisEngine:
    """Runs all PC health checks. Thread-safe (no shared mutable state)."""

    # ordered list of (display-label, method)
    CHECKS = [
        ("Xbox / Gaming Services",     "_check_gaming_services"),
        ("Visual C++ Runtimes",        "_check_vcredist"),
        (".NET Runtime",               "_check_dotnet"),
        ("DirectX",                    "_check_directx"),
        ("Disk Health (SMART)",        "_check_disk_health"),
        ("Virtual Memory / Page File", "_check_page_file"),
        ("Power Plan",                 "_check_power_plan"),
        ("Windows Defender",           "_check_defender"),
        ("Pending Reboot",             "_check_pending_reboot"),
        ("GPU Driver Age",             "_check_gpu_driver_age"),
        ("Startup Programs",           "_check_startup_items"),
        ("Windows Audio",              "_check_audio_service"),
        ("Windows Update Service",     "_check_wu_service"),
        ("Fast Startup",               "_check_fast_startup"),
        ("Recent System Errors",       "_check_event_errors"),
    ]

    def run_all(self, progress_cb=None):
        """
        Run every check.
        progress_cb(fraction: float, label: str) called between checks.
        Returns list of finding dicts, sorted: critical → warning → info → ok.
        """
        results = []
        total   = len(self.CHECKS)
        for i, (label, method_name) in enumerate(self.CHECKS):
            if progress_cb:
                progress_cb(i / total, f"Checking {label}…")
            try:
                found = getattr(self, method_name)()
                if found is None:
                    continue
                if isinstance(found, list):
                    results.extend(found)
                else:
                    results.append(found)
            except Exception:
                pass   # never crash the whole scan for one bad check

        if progress_cb:
            progress_cb(1.0, f"Done — {len(results)} finding(s)")

        order = {"critical": 0, "warning": 1, "info": 2, "ok": 3}
        results.sort(key=lambda f: order.get(f["status"], 9))
        return results

    # ─────────────────────────────────────────────────────── Xbox / Gaming ────

    def _check_gaming_services(self):
        findings = []

        # ── Gaming Services package (required for Xbox Game Pass) ─────────────
        # Method 1: winget
        out, _, rc = _run(
            ["winget", "list", "--id", "Microsoft.GamingServices",
             "--accept-source-agreements"], timeout=30)
        gaming_installed = rc == 0 and "GamingServices" in out

        # Method 2: AppxPackage fallback (catches MS Store installs winget may miss)
        if not gaming_installed:
            ps_gs, _, _ = _run_ps(
                "Get-AppxPackage -AllUsers -Name 'Microsoft.GamingServices'"
                " -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name",
                timeout=15)
            if ps_gs.strip():
                gaming_installed = True

        # Method 3: Service registry key
        if not gaming_installed:
            gaming_installed = _reg_exists(
                winreg.HKEY_LOCAL_MACHINE,
                r"SYSTEM\CurrentControlSet\Services\GamingServices")

        if not gaming_installed:
            findings.append(_finding(
                "gaming_services_missing", "Gaming", "🎮",
                "Microsoft Gaming Services Not Installed",
                "critical",
                "Gaming Services is required for Xbox Game Pass, Game Bar, and most "
                "Xbox PC games. Many titles will fail to launch without it. "
                "Install it from the Microsoft Store.",
                "Install Gaming Services",
                "store",
                "ms-windows-store://pdp/?productid=9MV0B5HZVK9Z",
            ))
        else:
            # Check if the background service is actually running
            svc_out, _, _ = _run_ps(
                "(Get-Service GamingServices -ErrorAction SilentlyContinue).Status",
                timeout=10)
            status = svc_out.strip()
            if status and "Running" not in status:
                findings.append(_finding(
                    "gaming_services_stopped", "Gaming", "🎮",
                    "Gaming Services Installed But Not Running",
                    "warning",
                    f"The GamingServices service reports '{status}'. "
                    "This can cause Game Pass titles to fail to launch or update. "
                    "Use the official Xbox PC App troubleshooter to repair it.",
                    "Open Xbox Troubleshooter",
                    "url",
                    "https://support.xbox.com/en-US/help/games-apps/game-setup-and-play/xbox-app-for-pc-troubleshooter",
                ))

        # ── Xbox Identity Provider — try multiple detection methods ─────────
        xip_found = False

        # Method 1: Get-AppxPackage (most reliable, needs no elevation)
        xip_out, _, _ = _run_ps(
            "Get-AppxPackage -AllUsers -Name 'Microsoft.XboxIdentityProvider'"
            " -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Version",
            timeout=15)
        if xip_out.strip():
            xip_found = True

        # Method 2: winget list as fallback
        if not xip_found:
            wg_out, _, wg_rc = _run(
                ["winget", "list", "--id", "Microsoft.XboxIdentityProvider",
                 "--accept-source-agreements"], timeout=20)
            if wg_rc == 0 and "XboxIdentityProvider" in wg_out:
                xip_found = True

        # Method 3: Check registry for the package family
        if not xip_found:
            try:
                import winreg as _wr
                _xip_path = (
                    r"SOFTWARE\Classes\Local Settings\Software\Microsoft\Windows"
                    r"\CurrentVersion\AppModel\Repository\Packages"
                )
                key = _wr.OpenKey(_wr.HKEY_LOCAL_MACHINE, _xip_path)
                i = 0
                while True:
                    try:
                        sub = _wr.EnumKey(key, i)
                        if "XboxIdentityProvider" in sub:
                            xip_found = True
                            break
                        i += 1
                    except OSError:
                        break
                _wr.CloseKey(key)
            except Exception:
                pass

        if not xip_found:
            findings.append(_finding(
                "xbox_identity_missing", "Gaming", "🎮",
                "Xbox Identity Provider Missing",
                "warning",
                "Xbox Identity Provider is required to sign into Xbox Live and launch "
                "Xbox / Game Pass games. Without it, games may refuse to start or "
                "display sign-in errors.",
                "Get from Microsoft Store",
                "store",
                "ms-windows-store://pdp/?productid=9WZDNCRD1HKW",
            ))

        # ── Core Xbox Live services ───────────────────────────────────────────
        for svc, friendly in [("XblAuthManager", "Xbox Live Auth Manager"),
                               ("XblGameSave",    "Xbox Live Game Save")]:
            out2, _, _ = _run(["sc", "query", svc], timeout=10)
            if "RUNNING" not in out2 and "STOPPED" not in out2:
                findings.append(_finding(
                    f"svc_{svc.lower()}", "Gaming", "🎮",
                    f"{friendly} Service Missing",
                    "warning",
                    f"The '{svc}' service is not present on this system. "
                    "This can prevent Xbox Live sign-in, cloud saves, and "
                    "achievements from working. Reinstalling Gaming Services usually fixes it.",
                    "Repair Gaming Services",
                    "store",
                    "ms-windows-store://pdp/?productid=9MV0B5HZVK9Z",
                ))

        if not findings:
            return _ok("gaming_services_ok", "Gaming", "🎮",
                       "Xbox / Gaming Services",
                       "Gaming Services, Xbox Identity Provider, and Xbox Live services "
                       "are all installed and healthy.")
        return findings

    # ──────────────────────────────────────────────────────── Runtimes ────────

    def _check_vcredist(self):
        """Visual C++ 2015-2022 Redistributable (x64) — needed by almost every game."""
        installed = False
        for hive, path in [
            (winreg.HKEY_LOCAL_MACHINE,
             r"SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64"),
            (winreg.HKEY_LOCAL_MACHINE,
             r"SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\x64"),
        ]:
            val = _reg_get(hive, path, "Installed")
            if val == 1:
                installed = True
                break

        if not installed:
            return _finding(
                "vcredist_missing", "Runtimes", "📦",
                "Visual C++ 2015–2022 Redistributable (x64) Missing",
                "critical",
                "The VC++ Redistributable is required by the vast majority of modern games, "
                "applications, and device drivers. Programs will crash or refuse to start "
                "without it. Download and run the official installer from Microsoft.",
                "Download VC++ Redistributable",
                "url",
                "https://aka.ms/vs/17/release/vc_redist.x64.exe",
            )
        return _ok("vcredist_ok", "Runtimes", "📦",
                   "Visual C++ Redistributable",
                   "Visual C++ 2015–2022 Redistributable (x64) is installed.")

    def _check_dotnet(self):
        """.NET runtime — detected via registry (reliable even without dotnet in PATH)."""
        versions = []

        # ── Method 1: Registry — SharedFramework installed versions (most reliable) ──
        reg_roots = [
            r"SOFTWARE\dotnet\Setup\InstalledVersions\x64\sharedhost",
            r"SOFTWARE\dotnet\Setup\InstalledVersions\x86\sharedhost",
        ]
        # Also enumerate the runtimes key
        for arch in ("x64", "x86"):
            try:
                key = winreg.OpenKey(
                    winreg.HKEY_LOCAL_MACHINE,
                    rf"SOFTWARE\dotnet\Setup\InstalledVersions\{arch}\sharedfx"
                    r"\Microsoft.NETCore.App")
                i = 0
                while True:
                    try:
                        ver = winreg.EnumKey(key, i)
                        versions.append(ver); i += 1
                    except OSError:
                        break
                winreg.CloseKey(key)
            except Exception:
                pass

        # ── Method 2: dotnet CLI (works if it happens to be in PATH) ─────────
        if not versions:
            out, _, rc = _run(["dotnet", "--list-runtimes"], timeout=10)
            if rc == 0 and out.strip():
                for line in out.splitlines():
                    parts = line.split()
                    if len(parts) >= 2 and "NETCore" in parts[0]:
                        versions.append(parts[1])

        # ── Method 3: PowerShell Get-Item on known install paths ─────────────
        if not versions:
            ps_out, _, _ = _run_ps(
                "Get-ChildItem 'C:\\Program Files\\dotnet\\shared"
                "\\Microsoft.NETCore.App' -ErrorAction SilentlyContinue"
                " | Select-Object -ExpandProperty Name", timeout=10)
            for line in ps_out.splitlines():
                v = line.strip()
                if v and re.match(r"\d+\.\d+", v):
                    versions.append(v)

        # ── Deduplicate and sort ───────────────────────────────────────────────
        seen = set(); clean = []
        for v in versions:
            if v not in seen:
                seen.add(v); clean.append(v)
        versions = sorted(clean, key=lambda v: [int(x) for x in re.findall(r"\d+", v)][:3])

        if not versions:
            return _finding(
                "dotnet_missing", "Runtimes", "📦",
                ".NET Runtime Not Detected",
                "info",
                "No .NET runtime was found via registry or file system. "
                "Many modern Windows apps and some games require .NET 6, 8, or 10. "
                "If you don't use apps that need it this can be ignored; otherwise "
                "install the latest LTS runtime from Microsoft.",
                "Download .NET Runtime",
                "url",
                "https://dotnet.microsoft.com/en-us/download/dotnet",
            )

        has_modern = any(
            v.startswith(("6.", "7.", "8.", "9.", "10.")) for v in versions)

        if not has_modern:
            latest = versions[-1]
            return _finding(
                "dotnet_old", "Runtimes", "📦",
                f".NET Runtime May Be Outdated (highest: {latest})",
                "info",
                f"The highest .NET runtime found is {latest}. Recent applications may "
                "require .NET 8 or .NET 10 (LTS). Consider installing the latest version.",
                "Download Latest .NET",
                "url",
                "https://dotnet.microsoft.com/en-us/download/dotnet",
            )

        return _ok("dotnet_ok", "Runtimes", "📦",
                   f".NET Runtime ({versions[-1]})",
                   f"Modern .NET runtime installed. "
                   f"Versions found: {', '.join(versions[-4:])}")

    def _check_directx(self):
        """DirectX version from registry."""
        version = _reg_get(winreg.HKEY_LOCAL_MACHINE,
                           r"SOFTWARE\Microsoft\DirectX", "Version")
        if not version:
            return _finding(
                "directx_unknown", "Runtimes", "🎲",
                "DirectX Version Could Not Be Read",
                "info",
                "DirectX registry entry not found. Run dxdiag from the Start menu "
                "to verify your DirectX installation.",
                "Run DxDiag",
                "command",
                "dxdiag",
            )
        parts = version.split(".")
        minor = int(parts[1]) if len(parts) > 1 else 0
        dx_ver = "12" if minor >= 9 else ("11" if minor >= 8 else str(minor))
        return _ok("directx_ok", "Runtimes", "🎲",
                   f"DirectX {dx_ver}",
                   f"DirectX {dx_ver} is installed (registry version: {version}).")

    # ─────────────────────────────────────────────────────── Storage ──────────

    def _check_disk_health(self):
        """SMART status of all physical disks via WMIC."""
        out, _, _ = _run(
            ["wmic", "diskdrive", "get", "Caption,Status"], timeout=20)
        findings = []
        lines = [l.strip() for l in out.splitlines() if l.strip()
                 and not l.strip().startswith("Caption")]
        for line in lines:
            parts = line.rsplit(None, 1)
            if len(parts) < 2:
                continue
            model, status = parts[0].strip(), parts[1].strip()
            if status.lower() not in ("ok",):
                findings.append(_finding(
                    f"disk_{model[:20].replace(' ', '_')}", "Storage", "💿",
                    f"Disk Warning: {model[:45]}",
                    "critical",
                    f"SMART status reported as '{status}'. This disk may be experiencing "
                    "hardware faults. Back up your data immediately and run a full "
                    "disk health check using CrystalDiskInfo.",
                    "Download CrystalDiskInfo",
                    "url",
                    "https://crystalmark.info/en/software/crystaldiskinfo/",
                ))
        if not findings:
            return _ok("disk_health_ok", "Storage", "💿",
                       "Disk Health (SMART)",
                       "All physical disks are reporting a healthy SMART status.")
        return findings

    # ─────────────────────────────────────────────────────── Memory ───────────

    def _check_page_file(self):
        """Check Windows virtual memory / page file."""
        val = _reg_get(
            winreg.HKEY_LOCAL_MACHINE,
            r"SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management",
            "PagingFiles")
        if not val or (isinstance(val, (list, tuple)) and not any(val)):
            return _finding(
                "pagefile_disabled", "Memory", "💾",
                "Page File (Virtual Memory) Is Disabled",
                "warning",
                "The Windows page file is disabled. This can cause system instability, "
                "application crashes, and 'out of memory' errors — especially in games "
                "that need more RAM than is physically available.",
                "Open Virtual Memory Settings",
                "command",
                "SystemPropertiesPerformance.exe",
            )
        return _ok("pagefile_ok", "Memory", "💾",
                   "Page File (Virtual Memory)",
                   "Virtual memory is configured and enabled.")

    # ──────────────────────────────────────────────────── Performance ─────────

    def _check_power_plan(self):
        """Active Windows power plan."""
        out, _, _ = _run(["powercfg", "/getactivescheme"], timeout=10)
        line = out.strip().lower()
        if "ultimate" in line or "e9a42b02" in line:
            return _ok("power_plan_ok", "Performance", "⚡",
                       "Power Plan: Ultimate Performance",
                       "Ultimate Performance plan is active — maximum clock speeds for "
                       "gaming and workstation use.")
        if "high performance" in line or "8c5e7fda" in line:
            return _ok("power_plan_ok", "Performance", "⚡",
                       "Power Plan: High Performance",
                       "High Performance plan is active — optimal for gaming.")
        m   = re.search(r'\((.+)\)', out.strip())
        plan = m.group(1) if m else out.strip()[-40:]
        return _finding(
            "power_plan_suboptimal", "Performance", "⚡",
            f"Power Plan: '{plan}' (Not Optimal for Gaming)",
            "warning",
            f"The current plan '{plan}' may throttle CPU and GPU clocks to save power. "
            "Switching to High Performance removes those limits and can improve game "
            "frame rates and loading times.",
            "Switch to High Performance",
            "command",
            "powercfg /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c",
        )

    def _check_startup_items(self):
        """Count registry startup entries and warn if excessive."""
        items = []
        for hive, path in [
            (winreg.HKEY_CURRENT_USER,
             r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run"),
            (winreg.HKEY_LOCAL_MACHINE,
             r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run"),
            (winreg.HKEY_LOCAL_MACHINE,
             r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run"),
        ]:
            try:
                key = winreg.OpenKey(hive, path)
                i = 0
                while True:
                    try:
                        name, _, _ = winreg.EnumValue(key, i)
                        items.append(name); i += 1
                    except OSError:
                        break
                winreg.CloseKey(key)
            except Exception:
                pass

        n = len(items)
        if n > 15:
            return _finding(
                "startup_excessive", "Performance", "🚀",
                f"Excessive Startup Programs ({n} items)",
                "warning",
                f"Found {n} programs configured to start with Windows. "
                "This slows boot time and keeps background processes consuming RAM and CPU "
                "even when you're not using those apps.",
                "Open Startup Settings",
                "command",
                "start ms-settings:startupapps",
            )
        if n > 8:
            return _finding(
                "startup_many", "Performance", "🚀",
                f"Startup Programs: {n} items",
                "info",
                f"{n} apps start with Windows. Review them in Task Manager "
                "and disable any you don't need immediately on boot.",
                "Open Startup Settings",
                "command",
                "start ms-settings:startupapps",
            )
        return _ok("startup_ok", "Performance", "🚀",
                   f"Startup Programs ({n} items)",
                   f"{n} startup items — reasonable, boot time should be unaffected.")

    # ─────────────────────────────────────────────────────── Security ─────────

    def _check_defender(self):
        """Windows Defender real-time protection status."""
        out, _, _ = _run_ps(
            "(Get-MpComputerStatus -ErrorAction SilentlyContinue)"
            ".RealTimeProtectionEnabled",
            timeout=15)
        val = out.strip().lower()
        if val == "false":
            return _finding(
                "defender_disabled", "Security", "🛡",
                "Windows Defender Real-Time Protection Is Off",
                "critical",
                "Real-time virus and threat protection is disabled. Your PC is at risk "
                "from malware. Enable it unless a third-party antivirus (e.g. Bitdefender, "
                "ESET) is actively managing protection.",
                "Open Windows Security",
                "command",
                "start windowsdefender:",
            )
        if val == "true":
            return _ok("defender_ok", "Security", "🛡",
                       "Windows Defender",
                       "Real-time protection is active — your system is protected.")
        return None  # 3rd-party AV or undetermined — skip

    # ─────────────────────────────────────────────────────── System ───────────

    def _check_pending_reboot(self):
        """Check registry flags that indicate a Windows restart is required."""
        reboot_keys = [
            (winreg.HKEY_LOCAL_MACHINE,
             r"SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate"
             r"\Auto Update\RebootRequired"),
            (winreg.HKEY_LOCAL_MACHINE,
             r"SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing"
             r"\RebootPending"),
        ]
        for hive, path in reboot_keys:
            if _reg_exists(hive, path):
                return _finding(
                    "pending_reboot", "System", "🔄",
                    "Windows Restart Is Required",
                    "warning",
                    "A pending restart was detected (likely from a recent Windows Update "
                    "or driver installation). Some changes won't take effect and new "
                    "updates can't install until the restart is completed.",
                    "Open Windows Update",
                    "command",
                    "start ms-settings:windowsupdate",
                )
        return _ok("pending_reboot_ok", "System", "🔄",
                   "Pending Restart",
                   "No pending restart detected — system is up to date.")

    def _check_audio_service(self):
        """Windows Audio service state."""
        out, _, _ = _run(["sc", "query", "AudioSrv"], timeout=10)
        if "RUNNING" not in out:
            return _finding(
                "audio_stopped", "System", "🔊",
                "Windows Audio Service Is Not Running",
                "critical",
                "The Windows Audio service is stopped. You will have no sound in any "
                "application or game. This can happen after a failed update or service "
                "corruption. Restart the service to restore audio.",
                "Restart Audio Service",
                "command",
                "net start AudioSrv",
            )
        return _ok("audio_ok", "System", "🔊",
                   "Windows Audio Service",
                   "Windows Audio is running normally.")

    def _check_wu_service(self):
        """Windows Update service state."""
        out, _, _ = _run(["sc", "query", "wuauserv"], timeout=10)
        if "RUNNING" not in out:
            return _finding(
                "wu_stopped", "System", "🪟",
                "Windows Update Service Is Stopped",
                "info",
                "The Windows Update service is not running. It normally starts on demand, "
                "but if updates are consistently failing this may be the reason.",
                "Start Windows Update Service",
                "command",
                "net start wuauserv",
            )
        return _ok("wu_ok", "System", "🪟",
                   "Windows Update Service",
                   "Windows Update service is running.")

    def _check_fast_startup(self):
        """Fast Startup can cause USB / dual-boot issues."""
        val = _reg_get(
            winreg.HKEY_LOCAL_MACHINE,
            r"SYSTEM\CurrentControlSet\Control\Session Manager\Power",
            "HiberbootEnabled")
        if val == 1:
            return _finding(
                "fast_startup_on", "System", "⚡",
                "Fast Startup Is Enabled",
                "info",
                "Fast Startup saves a partial hibernate state to speed up boot. "
                "This is generally fine but can occasionally cause USB devices not to "
                "re-initialise correctly, BIOS/UEFI settings not to apply, or issues "
                "with dual-boot Linux setups.",
                "Open Power & Sleep Settings",
                "command",
                "start ms-settings:powersleep",
            )
        return _ok("fast_startup_ok", "System", "⚡",
                   "Fast Startup",
                   "Fast Startup is disabled — full hardware reset on every boot.")

    def _check_event_errors(self):
        """Critical/Error entries in the System event log in the last 24 hours."""
        ps = (
            "$cutoff=(Get-Date).AddHours(-24); "
            "Get-EventLog -LogName System -EntryType Error,Warning "
            "-After $cutoff -ErrorAction SilentlyContinue "
            "| Select-Object -First 20 Source,EntryType,Message "
            "| ConvertTo-Json -ErrorAction SilentlyContinue"
        )
        out, _, rc = _run_ps(ps, timeout=25)
        if rc != 0 or not out.strip() or out.strip() in ("null", ""):
            return _ok("event_ok", "System", "📋",
                       "System Event Log",
                       "No errors or warnings in the Windows System log in the last 24 hours.")
        try:
            data = json.loads(out.strip())
            if isinstance(data, dict):
                data = [data]
            if not data:
                return _ok("event_ok", "System", "📋",
                           "System Event Log",
                           "No errors in the Windows System log in the last 24 hours.")
            errors   = [e for e in data if e.get("EntryType") in ("Error",   "1")]
            warnings = [e for e in data if e.get("EntryType") in ("Warning", "2")]
            sources  = list({e.get("Source", "?") for e in data})[:6]
            severity = "critical" if len(errors) >= 5 else "warning" if errors else "info"
            return _finding(
                "event_errors", "System", "📋",
                f"System Log: {len(errors)} Error(s), {len(warnings)} Warning(s) in 24h",
                severity,
                f"Found {len(errors)} error(s) and {len(warnings)} warning(s) in the "
                f"Windows System Event Log in the last 24 hours. "
                f"Sources: {', '.join(sources)}. "
                "Open Event Viewer for full details.",
                "Open Event Viewer",
                "command",
                "eventvwr.msc",
            )
        except Exception:
            return None

    def _check_gpu_driver_age(self):
        """Warn if the primary GPU driver is more than 6 months old."""
        try:
            import pythoncom
            import wmi as _wmi
            pythoncom.CoInitialize()
            c = _wmi.WMI()
            for g in c.Win32_VideoController():
                name = g.Name or ""
                if not any(k in name for k in ("AMD", "Radeon", "NVIDIA",
                                               "GeForce", "Intel Arc")):
                    continue
                date_str = str(g.DriverDate or "")[:8]
                pythoncom.CoUninitialize()
                if len(date_str) == 8:
                    try:
                        driver_date = datetime.strptime(date_str, "%Y%m%d")
                        age         = (datetime.now() - driver_date).days
                        months      = age // 30
                        if age > 180:
                            # Pick the right download URL per brand
                            if "AMD" in name or "Radeon" in name:
                                lbl = "Update AMD Driver"
                                act = "winget"
                                val = "AMD.AdrenalinSoftware"
                            elif "NVIDIA" in name or "GeForce" in name:
                                lbl = "Download NVIDIA Driver"
                                act = "url"
                                val = "https://www.nvidia.com/download/index.aspx"
                            else:
                                lbl = "Download Intel Arc Driver"
                                act = "url"
                                val = ("https://www.intel.com/content/www/us/en/products"
                                       "/docs/discrete-gpus/arc/software.html")
                            return _finding(
                                "gpu_driver_old", "Drivers", "🎮",
                                f"GPU Driver Is {months} Month(s) Old",
                                "warning",
                                f"{name} — driver dated "
                                f"{driver_date.strftime('%B %Y')} ({age} days ago). "
                                "Newer drivers often include game-specific optimisations, "
                                "bug fixes, and performance improvements.",
                                lbl, act, val,
                            )
                        return _ok("gpu_driver_ok", "Drivers", "🎮",
                                   "GPU Driver",
                                   f"{name} — driver from "
                                   f"{driver_date.strftime('%B %Y')} ({age} days old). "
                                   "Up to date.")
                    except Exception:
                        pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        except Exception:
            pass
        return None
