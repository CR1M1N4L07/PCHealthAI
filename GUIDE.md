# PC Health AI – Setup Guide

---

## What This App Does

| Feature | Details |
|---|---|
| **System Scanner** | Reads CPU, GPU, RAM, Storage, Motherboard, Network |
| **GPU VRAM Detection** | Special code bypasses Windows' 4 GB VRAM cap for accurate readings |
| **Update Centre** | Checks and installs app updates via winget, GPU drivers, Windows Updates |
| **Diagnosis & Repair** | Runs SFC, DISM, cleans temp files, flushes RAM |
| **Optimise** | Gaming tweaks and Windows debloat tools |

---

## STEP 1 – Install Python

1. Open your browser and go to: **https://www.python.org/downloads/**
2. Click the big yellow **"Download Python 3.12.x"** button
3. Run the downloaded `.exe`
4. On the FIRST screen, **tick the checkbox** that says **"Add Python to PATH"** (this is critical)
5. Click **"Install Now"**
6. Wait for it to finish, then click **Close**

**To verify it worked:**
- Press `Win + R`, type `cmd`, press Enter
- Type `python --version` and press Enter
- You should see something like `Python 3.12.4`

---

## STEP 2 – Copy the Files

Copy all files you received into a folder on your PC, for example:

```
C:\PCHealthAI\
  ├── main.py
  ├── system_info.py
  ├── update_manager.py
  ├── diagnosis_engine.py
  ├── config.json
  ├── requirements.txt
  ├── setup.bat
  └── start.bat
```

---

## STEP 3 – Run Setup (One Time Only)

1. Go to your `PCHealthAI` folder in File Explorer
2. **Right-click** `setup.bat`
3. Select **"Run as administrator"**
4. A black window opens and installs everything automatically
5. When you see **"Setup Complete!"** — press any key to close

> This takes 1–3 minutes depending on your internet speed.
> You only need to do this ONCE.

---

## STEP 4 – Launch the App

1. Go to your `PCHealthAI` folder in File Explorer
2. **Double-click** `start.bat`
3. The app opens with a dark window titled **"PC Health AI"**

---

## STEP 5 – Using the App

### System Scan
- The app **auto-scans on startup** — you'll see all your hardware in a few seconds
- Click **🔍 Scan System** any time to refresh

### Navigation (left sidebar)

| Section | What it shows |
|---|---|
| **🏠 Dashboard** | Quick cards for every component at a glance |
| **💻 Hardware** | CPU · GPU · RAM · Storage detail tabs |
| **🌐 Network** | All network adapters, IP, speeds, DNS changer |
| **📋 Processes** | Top CPU & RAM processes, live updated |
| **🔧 Maintenance** | Updates · Diagnosis · System Repair tools |
| **⚡ Optimise** | Gaming tweaks & Windows debloat |

### Update Centre (Maintenance tab)

| Button | What it does |
|---|---|
| 🔍 Scan Updates | Checks for available app updates via winget |
| 📦 Update Selected | Installs only the apps you tick |
| 📦 Update All | Installs all available updates silently |
| 🎮 AMD Driver | Updates AMD Adrenalin Software |
| 🪟 Windows Update | Opens Windows Update settings |

### DNS Changer (Dashboard or Network tab)

Choose a DNS preset from the dropdown and click **✅ Apply**. Options include Cloudflare, Google, Quad9, OpenDNS, and Cloudflare Gaming. Use **↩ Reset** to go back to automatic (DHCP).

---

## Everyday Use

After the first setup, just **double-click `start.bat`** to open the app. No installation needed again.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `python is not recognized` | Re-install Python with "Add to PATH" ticked |
| App opens then closes | Right-click `start.bat` → Run as administrator |
| GPU shows wrong VRAM | Run as Administrator; restart and scan again |
| winget not found | Open Microsoft Store, search "App Installer", install it |
| WMI errors on scan | Run `start.bat` as Administrator |
| Package errors | Double-click `fix_packages.bat` |

---

## Administrator Rights

Running as Administrator unlocks full functionality:
- Accurate SMART drive health data
- DNS changes applied to all adapters
- SFC and DISM system repair
- Gaming tweaks that require registry/service changes
- Deeper RAM flush

Right-click `start.bat` → **Run as administrator**, or click **🔑 Relaunch as Admin** in the sidebar.
