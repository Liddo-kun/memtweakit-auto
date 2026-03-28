# MemTweakIt Automation

An app for managing and automatically applying memory timings profiles using MemTweakIt. Can be used for applying memory timings at startup for systems that don't support some BIOS-level memory timings.

## Requirements

- Windows 10/11 (64-bit)
- [MemTweakIt](https://rog.asus.com/articles/overclocking/how-to-use-memtweakit/) by ASUS ROG (download from your motherboard's support page on [asus.com](https://www.asus.com) under Drivers & Tools > Utilities)
- Must run as Administrator (MemTweakIt loads a kernel driver)

## Setup

Place all of these in the same folder:

```
apply_timings.exe       <- this tool
MemTweakIt.exe          <- from MemTweakIt download
cpuidsdk.dll            <- from MemTweakIt download
cpuidsdk64.dll          <- from MemTweakIt download
UpdateHelper.dll        <- from MemTweakIt download
```

## Usage

### First Run

Double-click `apply_timings.exe` (run as admin). Since no `timings.ini` exists yet, it will:

1. Launch MemTweakIt
2. Read all your current memory timing values
3. Save them to `timings.ini`
4. Tell you to edit the file and run again

### Editing Timings

Open `timings.ini` in any text editor. It looks like this:

```ini
[Settings]
; Path to MemTweakIt.exe (defaults to same directory as apply_timings.exe)
; Path = C:\path\to\MemTweakIt.exe
StartDelay = 3
; Action: apply = click Apply only, ok = click OK (apply+close)
Action = ok

[Timings]
CAS# Latency (CL) = 52
RAS# to CAS# Delay (tRCD) = 49
tRFC = 670
; ... etc
```

- Change values to your desired timings
- Comment out (`;`) or delete any timing you don't want to change
- Values must exactly match what appears in MemTweakIt's dropdowns

### Applying Timings

Double-click `apply_timings.exe` (run as admin). It will:

1. Minimize itself and MemTweakIt to the background
2. Set each timing to the value in `timings.ini`
3. Click Apply + OK to commit changes and close

### Apply at Startup

To automatically apply timings every time you log in:

1. Open Task Scheduler (`taskschd.msc`)
2. Create a new task (not a basic task)
3. **General tab**: Check "Run with highest privileges"
4. **Trigger**: "At log on" for your user
5. **Action**: "Start a program", browse to `apply_timings.exe`
6. **Conditions**: Uncheck "Start only if on AC power" if on a desktop
7. **Settings**: Check "Run only when user is logged on"

### Re-discovering Timings

If you change timings in BIOS or want to capture current values without overwriting your config:

```
apply_timings.exe --discover
```

This writes to `timings_discovered.ini` (your `timings.ini` is not touched).

### Command Line Options

```
apply_timings.exe                      Apply timings from timings.ini
apply_timings.exe --discover           Discover current values, write timings_discovered.ini
apply_timings.exe --ini custom.ini     Use a custom INI file
```

## Building from Source

Requires Visual Studio Build Tools (MSVC). From a Developer Command Prompt:

```
cl /O2 /W3 /MT /D_UNICODE /DUNICODE apply_timings.c /Fe:apply_timings.exe user32.lib kernel32.lib shell32.lib
```

Single file, no external dependencies. Static CRT (`/MT`) so no runtime DLLs needed.

## How It Works

The tool automates MemTweakIt's GUI via Win32 `SendMessage`/`PostMessage`:

1. Launches MemTweakIt and finds its window by PID
2. Minimizes both windows to run in the background
3. Switches between tabs by clicking tab headers (cross-process `VirtualAllocEx` for tab rect coordinates)
4. Matches each combo box to its label using spatial proximity (label's right edge nearest to combo's left edge at the same Y)
5. For dropdown combos: navigates with `VK_DOWN`/`VK_UP` key messages via `PostMessage`
6. For typable combos: selects all text and types the new value via `WM_CHAR`
7. Verifies each value took by reading it back
8. Clicks Apply (commits to hardware) then OK (closes)
