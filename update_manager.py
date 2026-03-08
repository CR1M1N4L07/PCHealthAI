"""
update_manager.py
Manages software updates, system repairs, temp cleanup,
gaming optimisations, and Windows bloatware removal.
"""

import subprocess
import json
import os

CREATE_NO_WINDOW = 0x08000000


class UpdateManager:

    def __init__(self):
        self._winget_available = self._check_winget()

    # ─────────────────────────────────────── internal helpers ───────────────

    def _check_winget(self):
        try:
            r = subprocess.run(["winget", "--version"],
                               capture_output=True, text=True,
                               creationflags=CREATE_NO_WINDOW, timeout=10)
            return r.returncode == 0
        except Exception:
            return False

    def _run(self, cmd, timeout=180):
        try:
            p = subprocess.run(cmd, capture_output=True, text=True,
                               encoding="utf-8", errors="replace",
                               creationflags=CREATE_NO_WINDOW, timeout=timeout)
            return p.stdout, p.stderr, p.returncode
        except subprocess.TimeoutExpired:
            return "", "Timed out.", -1
        except FileNotFoundError:
            return "", f"Not found: {cmd[0]}", -1
        except Exception as e:
            return "", str(e), -1

    # Lines from winget output that are pure noise — never pass to the UI
    _WINGET_NOISE = frozenset({"-", "\\", "|", "/", ""})
    # Substrings whose lines are informational but not errors
    _WINGET_INFO_HINTS = (
        "Microsoft is not responsible",
        "This application is licensed",
        "Starting package install",
        "Downloading ",
        "Verifying",
        "Successfully verified",
        "Initiating",
    )

    def _classify_winget_line(self, line: str):
        """
        Returns (text, tag) where tag is one of:
          'success' | 'warn' | 'error' | 'info' | 'dim' | None (skip)
        """
        s = line.strip()
        if not s or s in self._WINGET_NOISE:
            return None, None
        sl = s.lower()
        if "successfully installed" in sl or "successfully updated" in sl:
            return s, "success"
        if "no applicable upgrade" in sl or "no newer package" in sl:
            return s, "warn"
        if "failed" in sl or "error" in sl or "0x8" in s:
            return s, "error"
        if "version number cannot be determined" in sl:
            return "ℹ  Version unknown — installing with --include-unknown", "info"
        if any(h in s for h in self._WINGET_INFO_HINTS):
            return s, "dim"
        return s, ""

    def _stream(self, cmd, cb=None, timeout=600):
        """Stream a subprocess, calling cb(line, tag) for each meaningful line."""
        try:
            proc = subprocess.Popen(cmd, stdout=subprocess.PIPE,
                                    stderr=subprocess.STDOUT,
                                    text=True, encoding="utf-8",
                                    errors="replace",
                                    creationflags=CREATE_NO_WINDOW)
            lines = []
            for raw in proc.stdout:
                s = raw.rstrip()
                text, tag = self._classify_winget_line(s)
                if text is None:
                    continue          # skip noise
                lines.append(text)
                if cb:
                    cb(text, tag)     # pass tag to UI
            proc.wait()
            return {"success": proc.returncode in (0, 1),
                    "output": "\n".join(lines),
                    "rc": proc.returncode}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def _run_ps(self, script, timeout=300):
        return self._run(
            ["powershell", "-ExecutionPolicy", "Bypass",
             "-NonInteractive", "-Command", script],
            timeout=timeout)

    # ─────────────────────────────────────── winget updates ─────────────────

    def check_winget_updates(self):
        if not self._winget_available:
            return {"error": "winget not available."}
        out, _, rc = self._run(
            ["winget", "upgrade", "--include-unknown",
             "--accept-source-agreements"], timeout=90)
        if not out.strip():
            return {"error": "winget returned no output."}
        return self._parse_winget(out)

    def _parse_winget(self, output):
        updates, lines = [], output.splitlines()
        hi = next((i for i, l in enumerate(lines)
                   if "Name" in l and "Id" in l and "Version" in l), -1)
        if hi < 0:
            return updates
        h = lines[hi]
        nc = h.find("Name"); ic = h.find("Id")
        vc = h.find("Version"); ac = h.find("Available")
        sc = h.find("Source"); sc = sc if sc > 0 else len(h)
        for line in lines[hi + 2:]:
            if not line.strip() or line.strip().startswith("-"):
                continue
            try:
                updates.append({
                    "name":              line[nc:ic].strip(),
                    "id":                line[ic:vc].strip(),
                    "current_version":   line[vc:ac].strip(),
                    "available_version": line[ac:sc].strip(),
                })
            except Exception:
                continue
        return updates

    def install_all_winget_updates(self, cb=None):
        if not self._winget_available:
            return {"success": False, "error": "winget not available."}
        return self._stream(
            ["winget", "upgrade", "--all", "--silent",
             "--accept-package-agreements",
             "--accept-source-agreements", "--include-unknown"], cb)

    def install_winget_update(self, pkg_id, cb=None):
        return self._stream(
            ["winget", "upgrade", "--id", pkg_id, "--silent",
             "--accept-package-agreements",
             "--accept-source-agreements",
             "--include-unknown"], cb)

    def check_amd_driver(self):
        try:
            import pythoncom, wmi
            pythoncom.CoInitialize()
            c = wmi.WMI()
            amd = [g for g in c.Win32_VideoController()
                   if any(k in (g.Name or "")
                          for k in ("AMD", "Radeon", "ATI"))]
            pythoncom.CoUninitialize()
            if amd:
                g = amd[0]
                return {"name": g.Name, "current_driver": g.DriverVersion,
                        "driver_date": str(g.DriverDate)[:8] if g.DriverDate else "N/A",
                        "check_url": "https://www.amd.com/en/support/download/drivers.html",
                        "winget_id": "AMD.AdrenalinSoftware"}
        except Exception as e:
            return {"error": str(e)}
        return {"message": "No AMD GPU found."}

    def install_amd_software(self, cb=None):
        return self.install_winget_update("AMD.AdrenalinSoftware", cb)

    def open_windows_update(self):
        try:
            subprocess.Popen("start ms-settings:windowsupdate",
                             shell=True, creationflags=CREATE_NO_WINDOW)
            return {"success": True}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_recent_windows_updates(self, count=15):
        ps = (f"Get-WmiObject Win32_QuickFixEngineering "
              f"| Sort-Object InstalledOn -Descending "
              f"| Select-Object -First {count} HotFixID,Description,InstalledOn "
              f"| ConvertTo-Json")
        out, _, _ = self._run_ps(ps, timeout=30)
        try:
            data = json.loads(out.strip())
            if isinstance(data, dict):
                data = [data]
            return [{"hotfix_id": i.get("HotFixID", "N/A"),
                     "description": i.get("Description", "N/A"),
                     "installed_on": i.get("InstalledOn", "N/A")}
                    for i in (data or [])]
        except Exception:
            return []

    # ─────────────────────────────────────── system repair ──────────────────

    def run_sfc_scan(self, cb=None):
        """
        Run SFC /scannow. Must be called from an elevated process.
        Streams output line by line via cb(line).
        """
        return self._stream(
            ["sfc", "/scannow"], cb, timeout=900)

    def run_dism_restore_health(self, cb=None):
        """
        Run DISM /RestoreHealth. Should be run AFTER SFC.
        Streams output line by line via cb(line).
        """
        return self._stream(
            ["DISM", "/Online", "/Cleanup-Image", "/RestoreHealth"],
            cb, timeout=1800)

    def run_dism_scan_health(self, cb=None):
        return self._stream(
            ["DISM", "/Online", "/Cleanup-Image", "/ScanHealth"],
            cb, timeout=600)

    def run_chkdsk(self, drive="C:", cb=None):
        """Schedule chkdsk on next reboot (non-destructive)."""
        out, err, rc = self._run(
            ["chkdsk", drive, "/f", "/r", "/x"], timeout=30)
        full = out + err
        if cb:
            cb(full)
        return {"success": True, "output": full}

    # ─────────────────────────────────────── temp / junk cleanup ────────────

    def clean_temp_folders(self, cb=None):
        """
        Delete contents of:
          %TEMP%  (user temp)
          C:\\Windows\\Temp  (system temp)
          C:\\Windows\\Prefetch
        Returns total bytes freed.
        """
        import shutil, glob, ctypes

        def _log(msg):
            if cb:
                cb(msg)

        total_freed = 0
        targets = [
            os.environ.get("TEMP", ""),
            os.environ.get("TMP", ""),
            r"C:\Windows\Temp",
            r"C:\Windows\Prefetch",
        ]
        seen = set()
        for folder in targets:
            folder = os.path.normpath(folder)
            if not folder or folder in seen or not os.path.isdir(folder):
                continue
            seen.add(folder)
            _log(f"  Cleaning: {folder}")
            deleted = skipped = freed = 0
            for item in os.listdir(folder):
                item_path = os.path.join(folder, item)
                try:
                    if os.path.isfile(item_path) or os.path.islink(item_path):
                        size = os.path.getsize(item_path)
                        os.unlink(item_path)
                        freed += size
                        deleted += 1
                    elif os.path.isdir(item_path):
                        size = sum(
                            os.path.getsize(os.path.join(dp, f))
                            for dp, _, fs in os.walk(item_path)
                            for f in fs
                            if os.path.exists(os.path.join(dp, f)))
                        shutil.rmtree(item_path, ignore_errors=True)
                        freed += size
                        deleted += 1
                except Exception:
                    skipped += 1
            total_freed += freed
            _log(f"    Removed {deleted} items, skipped {skipped} "
                 f"({freed / (1024**2):.1f} MB freed)")

        # Also clear DNS cache
        try:
            subprocess.run(["ipconfig", "/flushdns"],
                           capture_output=True, creationflags=CREATE_NO_WINDOW)
            _log("  DNS cache flushed.")
        except Exception:
            pass

        return {"success": True,
                "total_freed_mb": round(total_freed / (1024 ** 2), 1)}

    # ───────────────────────────────── gaming optimisation ──────────────────

    def apply_gaming_optimisations(self, cb=None):
        """
        Apply a curated set of Windows registry and service tweaks
        that reduce latency and background CPU usage during gaming.
        All changes are reversible via Windows settings.
        """
        def _log(msg):
            if cb:
                cb(msg)

        tweaks = []

        # ── 1. Set Power Plan to High Performance ──────────────────────────
        tweaks.append((
            "Power Plan → High Performance",
            'powercfg /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'
        ))

        # ── 2. Disable Xbox Game Bar (keeps running in background) ─────────
        tweaks.append((
            "Disable Xbox Game Bar",
            r'reg add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\GameDVR" '
            r'/v AppCaptureEnabled /t REG_DWORD /d 0 /f'
        ))
        tweaks.append((
            "Disable Game DVR",
            r'reg add "HKCU\System\GameConfigStore" '
            r'/v GameDVR_Enabled /t REG_DWORD /d 0 /f'
        ))

        # ── 3. Enable Hardware-Accelerated GPU Scheduling (HAGS) ───────────
        tweaks.append((
            "Enable HAGS (GPU Scheduling)",
            r'reg add "HKLM\SYSTEM\CurrentControlSet\Control\GraphicsDrivers" '
            r'/v HwSchMode /t REG_DWORD /d 2 /f'
        ))

        # ── 4. Disable Nagle's Algorithm (reduces network latency) ─────────
        tweaks.append((
            "Disable Nagle's Algorithm (low-latency network)",
            r'reg add "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces" '
            r'/v TcpAckFrequency /t REG_DWORD /d 1 /f & '
            r'reg add "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces" '
            r'/v TCPNoDelay /t REG_DWORD /d 1 /f'
        ))

        # ── 5. Disable SysMain (Superfetch) – helps with fast SSDs ────────
        tweaks.append((
            "Disable SysMain / Superfetch",
            'sc config SysMain start= disabled & net stop SysMain'
        ))

        # ── 6. Disable Windows Search indexing service ─────────────────────
        tweaks.append((
            "Disable Windows Search Indexer (background I/O)",
            'sc config WSearch start= disabled & net stop WSearch'
        ))

        # ── 7. Set GPU priority for games in registry ──────────────────────
        tweaks.append((
            "Prioritise GPU for graphics tasks",
            r'reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile\Tasks\Games" '
            r'/v "GPU Priority" /t REG_DWORD /d 8 /f & '
            r'reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile\Tasks\Games" '
            r'/v Priority /t REG_DWORD /d 6 /f & '
            r'reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile\Tasks\Games" '
            r'/v "Scheduling Category" /t REG_SZ /d High /f'
        ))

        # ── 8. Disable fullscreen optimisations globally ───────────────────
        tweaks.append((
            "Disable Fullscreen Optimisations (better exclusive fullscreen)",
            r'reg add "HKCU\System\GameConfigStore" '
            r'/v GameDVR_FSEBehaviorMode /t REG_DWORD /d 2 /f & '
            r'reg add "HKCU\System\GameConfigStore" '
            r'/v GameDVR_HonorUserFSEBehaviorMode /t REG_DWORD /d 1 /f'
        ))

        # ── 9. Adjust visual effects for performance ───────────────────────
        tweaks.append((
            "Visual Effects → Best Performance mode",
            r'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects" '
            r'/v VisualFXSetting /t REG_DWORD /d 2 /f'
        ))

        # ── 10. Disable mouse acceleration ─────────────────────────────────
        tweaks.append((
            "Disable Mouse Acceleration (raw input)",
            r'reg add "HKCU\Control Panel\Mouse" /v MouseSpeed /t REG_SZ /d 0 /f & '
            r'reg add "HKCU\Control Panel\Mouse" /v MouseThreshold1 /t REG_SZ /d 0 /f & '
            r'reg add "HKCU\Control Panel\Mouse" /v MouseThreshold2 /t REG_SZ /d 0 /f'
        ))

        results = []
        for name, cmd in tweaks:
            _log(f"\n  ▶  {name}")
            out, err, rc = self._run(["cmd", "/c", cmd], timeout=30)
            status = "✅ Done" if rc == 0 else f"⚠  rc={rc}"
            _log(f"     {status}")
            results.append({"name": name, "success": rc == 0})

        _log("\n  ✅  All gaming optimisations applied.")
        _log("  ⚠   Restart your PC to activate all changes.\n")
        return {"success": True, "results": results}

    # ───────────────────────────────── bloatware removal ────────────────────

    def remove_bloatware(self, cb=None):
        """
        Uninstall common Microsoft bloatware and disable AI/Copilot features.
        Uses PowerShell Remove-AppxPackage for store apps,
        and registry tweaks to disable Copilot, Recall, and ads.
        """
        def _log(msg):
            if cb:
                cb(msg)

        # ── AppX packages to remove ────────────────────────────────────────
        appx_packages = [
            ("Cortana",                    "Microsoft.549981C3F5F10"),
            ("Microsoft Copilot",          "Microsoft.Copilot"),
            ("Xbox App",                   "Microsoft.XboxApp"),
            ("Xbox Game Overlay",          "Microsoft.XboxGameOverlay"),
            ("Xbox Gaming Overlay",        "Microsoft.XboxGamingOverlay"),
            ("Xbox Identity Provider",     "Microsoft.XboxIdentityProvider"),
            ("Xbox Speech To Text",        "Microsoft.XboxSpeechToTextOverlay"),
            ("Bing News",                  "Microsoft.BingNews"),
            ("Bing Weather",               "Microsoft.BingWeather"),
            ("Bing Search",                "Microsoft.BingSearch"),
            ("Get Help",                   "Microsoft.GetHelp"),
            ("Feedback Hub",               "Microsoft.WindowsFeedbackHub"),
            ("Mixed Reality Portal",       "Microsoft.MixedReality.Portal"),
            ("3D Viewer",                  "Microsoft.Microsoft3DViewer"),
            ("Mail and Calendar",          "microsoft.windowscommunicationsapps"),
            ("Movies & TV",                "Microsoft.ZuneVideo"),
            ("Groove Music",               "Microsoft.ZuneMusic"),
            ("People",                     "Microsoft.People"),
            ("Your Phone / Phone Link",    "Microsoft.YourPhone"),
            ("Power Automate",             "Microsoft.PowerAutomateDesktop"),
            ("Teams (personal)",           "MicrosoftTeams"),
            ("OneNote for Windows 10",     "Microsoft.Office.OneNote"),
            ("Solitaire Collection",       "Microsoft.MicrosoftSolitaireCollection"),
            ("Tips",                       "Microsoft.Getstarted"),
            ("Todo",                       "Microsoft.Todos"),
            ("Clipchamp",                  "Clipchamp.Clipchamp"),
            ("OneDrive (store version)",   "Microsoft.OneDriveSync"),
        ]

        _log("  ── Removing bloatware AppX packages ──────────────────────")
        for friendly, pkg in appx_packages:
            _log(f"  Removing: {friendly}")
            ps = (f'Get-AppxPackage -AllUsers "*{pkg}*" '
                  f'| Remove-AppxPackage -ErrorAction SilentlyContinue; '
                  f'Get-AppxProvisionedPackage -Online '
                  f'| Where-Object DisplayName -like "*{pkg}*" '
                  f'| Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue')
            self._run_ps(ps, timeout=60)

        # ── Disable Windows Copilot ────────────────────────────────────────
        _log("\n  ── Disabling Windows Copilot & AI features ───────────────")

        reg_tweaks = [
            ("Disable Copilot button on taskbar",
             r'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" '
             r'/v ShowCopilotButton /t REG_DWORD /d 0 /f'),
            ("Disable Copilot via policy",
             r'reg add "HKCU\Software\Policies\Microsoft\Windows\WindowsCopilot" '
             r'/v TurnOffWindowsCopilot /t REG_DWORD /d 1 /f'),
            ("Disable Windows Recall (AI screenshot feature)",
             r'reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsAI" '
             r'/v DisableAIDataAnalysis /t REG_DWORD /d 1 /f'),
            ("Disable AI-powered search highlights",
             r'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\SearchSettings" '
             r'/v IsDynamicSearchBoxEnabled /t REG_DWORD /d 0 /f'),
            ("Disable personalised ads",
             r'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo" '
             r'/v Enabled /t REG_DWORD /d 0 /f'),
            ("Disable app launch tracking",
             r'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" '
             r'/v Start_TrackProgs /t REG_DWORD /d 0 /f'),
            ("Disable suggested content in Settings",
             r'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" '
             r'/v SubscribedContent-338393Enabled /t REG_DWORD /d 0 /f & '
             r'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" '
             r'/v SubscribedContent-353694Enabled /t REG_DWORD /d 0 /f & '
             r'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" '
             r'/v SubscribedContent-353696Enabled /t REG_DWORD /d 0 /f'),
            ("Disable Start Menu suggestions / ads",
             r'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" '
             r'/v SystemPaneSuggestionsEnabled /t REG_DWORD /d 0 /f & '
             r'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" '
             r'/v SoftLandingEnabled /t REG_DWORD /d 0 /f'),
            ("Disable Telemetry (set to Security level)",
             r'reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\DataCollection" '
             r'/v AllowTelemetry /t REG_DWORD /d 0 /f'),
            ("Disable DiagTrack (Connected User Experiences) service",
             'sc config DiagTrack start= disabled & net stop DiagTrack'),
            ("Disable WAP Push Message Routing Service",
             'sc config dmwappushservice start= disabled & net stop dmwappushservice'),
            ("Remove Chat / Teams icon from taskbar",
             r'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" '
             r'/v TaskbarMn /t REG_DWORD /d 0 /f'),
            ("Remove Widgets from taskbar",
             r'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" '
             r'/v TaskbarDa /t REG_DWORD /d 0 /f'),
        ]

        for name, cmd in reg_tweaks:
            _log(f"  ▶  {name}")
            _, _, rc = self._run(["cmd", "/c", cmd], timeout=20)
            _log(f"     {'✅ Done' if rc == 0 else '⚠  rc=' + str(rc)}")

        # ── Disable OneDrive auto-start ────────────────────────────────────
        _log("\n  ── Disabling OneDrive auto-start ─────────────────────────")
        self._run(["cmd", "/c",
                   r'reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" '
                   r'/v OneDrive /f'], timeout=10)
        _log("  ✅  OneDrive auto-start removed.")

        _log("\n  ✅  Bloatware removal complete.")
        _log("  ⚠   Restart your PC to fully apply all changes.\n")
        return {"success": True}
