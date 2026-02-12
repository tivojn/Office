# AppleScript Patterns — Excel Orchestration & Recovery (macOS)

AppleScript serves two roles in the xlsx-design-agent workflow:

1. **Orchestration** — Open files in Excel, run VBA macros, save, close, export
2. **Recovery** — Fix Excel when it's stuck, unresponsive, or in a modal state

All file creation and editing is done by **openpyxl** (no Excel needed). AppleScript enters the picture only when you need Excel to be running — to view the finished workbook, run VBA macros, or recover from errors.

## When to Use AppleScript

| Situation | Use AppleScript? | Alternative |
|-----------|-----------------|-------------|
| Open finished workbook in Excel | **YES** | `subprocess.run(['open', path])` also works |
| Run a VBA macro | **YES** (`do Visual Basic`) | — |
| Save the workbook in Excel | **YES** | — |
| Close the workbook in Excel | **YES** | — |
| Export PDF via VBA | **YES** (run VBA macro) | — |
| Excel stuck in modal dialog | **YES** (recovery) | — |
| Excel not frontmost / lost focus | **YES** (recovery) | — |
| Write cell values | **NO** | openpyxl |
| Format cells | **NO** | openpyxl |
| Create/style charts | **NO** | openpyxl |
| Insert images | **NO** | openpyxl |
| Set borders | **NO** | openpyxl |
| Conditional formatting | **NO** | openpyxl |

---

## Orchestration Patterns

### 1. Open a File in Excel

```bash
# Simple — works for .xlsx and .xlsm
open "/path/to/workbook.xlsx"
```

Or via AppleScript for more control:

```bash
osascript -e 'tell application "Microsoft Excel" to open "/path/to/workbook.xlsx"'
```

### 2. Save the Active Workbook

```bash
osascript -e 'tell application "Microsoft Excel" to save active workbook'
```

### 3. Close the Active Workbook

```bash
# Close with save
osascript -e 'tell application "Microsoft Excel" to close active workbook saving yes'

# Close without save
osascript -e 'tell application "Microsoft Excel" to close active workbook saving no'
```

### 4. Quit Excel

```bash
osascript -e 'tell application "Microsoft Excel" to quit saving yes'
```

### 5. Run a Named VBA Macro

```bash
osascript -e 'tell application "Microsoft Excel" to do Visual Basic "Call MacroName"'
```

With arguments:
```bash
osascript -e 'tell application "Microsoft Excel" to do Visual Basic "Call MacroName(\"arg1\", \"arg2\")"'
```

### 6. Run Inline VBA Code

```bash
osascript <<'EOF'
tell application "Microsoft Excel"
    do Visual Basic "
        Sub TempMacro()
            ActiveWorkbook.RefreshAll
            Application.CalculateFull
        End Sub
        Call TempMacro
    "
end tell
EOF
```

### 7. Export PDF via VBA

```bash
osascript -e 'tell application "Microsoft Excel" to do Visual Basic "Call ExportPdf"'
```

Or inline:
```bash
osascript <<'EOF'
tell application "Microsoft Excel"
    do Visual Basic "
        ActiveSheet.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=\"/Users/user/output.pdf\", _
            Quality:=xlQualityStandard
    "
end tell
EOF
```

### 8. Set Calculation Mode

```bash
# Manual calculation (faster for bulk operations)
osascript -e 'tell application "Microsoft Excel" to do Visual Basic "Application.Calculation = xlCalculationManual"'

# Automatic calculation (normal mode)
osascript -e 'tell application "Microsoft Excel" to do Visual Basic "Application.Calculation = xlCalculationAutomatic"'
```

### 9. Force Recalculation

```bash
osascript -e 'tell application "Microsoft Excel" to do Visual Basic "Application.CalculateFull"'
```

### 10. Activate Excel & Bring to Front

```bash
osascript -e 'tell application "Microsoft Excel" to activate'
```

### 11. Check if Excel is Running

```bash
osascript -e 'tell application "System Events" to (name of processes) contains "Microsoft Excel"'
```

### 12. Launch Excel if Not Running

```bash
osascript -e 'tell application "Microsoft Excel" to launch'
```

### 13. Get Active Workbook Path

```bash
osascript -e 'tell application "Microsoft Excel" to return full name of active workbook'
```

### 14. List Open Workbooks

```bash
osascript -e 'tell application "Microsoft Excel" to return name of every workbook'
```

---

## Python Helper Functions

Use these in your Python scripts to call AppleScript from Python:

```python
import subprocess
import time

def open_in_excel(path):
    """Open a workbook in Excel."""
    subprocess.run(['open', path], check=True)
    time.sleep(2)  # Give Excel time to open the file

def save_workbook():
    """Save the active workbook in Excel."""
    subprocess.run([
        'osascript', '-e',
        'tell application "Microsoft Excel" to save active workbook'
    ], capture_output=True, check=True)

def close_workbook(save=True):
    """Close the active workbook."""
    saving = "yes" if save else "no"
    subprocess.run([
        'osascript', '-e',
        f'tell application "Microsoft Excel" to close active workbook saving {saving}'
    ], capture_output=True, check=True)

def quit_excel(save=True):
    """Quit Excel."""
    saving = "yes" if save else "no"
    subprocess.run([
        'osascript', '-e',
        f'tell application "Microsoft Excel" to quit saving {saving}'
    ], capture_output=True, check=True)

def run_vba_macro(macro_name, *args):
    """Run a named VBA macro via AppleScript."""
    if args:
        arg_str = ', '.join(f'"{a}"' if isinstance(a, str) else str(a) for a in args)
        vba_call = f'Call {macro_name}({arg_str})'
    else:
        vba_call = f'Call {macro_name}'

    result = subprocess.run([
        'osascript', '-e',
        f'tell application "Microsoft Excel" to do Visual Basic "{vba_call}"'
    ], capture_output=True, text=True)

    if result.returncode != 0:
        raise RuntimeError(f"VBA macro failed: {result.stderr}")
    return result.stdout.strip()

def run_vba_inline(code):
    """Run inline VBA code via AppleScript."""
    # Escape double quotes in the VBA code
    escaped = code.replace('"', '\\"')
    result = subprocess.run([
        'osascript', '-e',
        f'tell application "Microsoft Excel" to do Visual Basic "{escaped}"'
    ], capture_output=True, text=True)

    if result.returncode != 0:
        raise RuntimeError(f"VBA inline failed: {result.stderr}")
    return result.stdout.strip()

def recalculate():
    """Force full recalculation in Excel."""
    subprocess.run([
        'osascript', '-e',
        'tell application "Microsoft Excel" to do Visual Basic "Application.CalculateFull"'
    ], capture_output=True, check=True)
```

---

## Recovery Patterns

Use these when Excel becomes stuck, unresponsive, or enters a modal state.

### Close Modal Dialog (Safe)

```bash
osascript <<'EOF'
tell application "System Events"
    tell process "Microsoft Excel"
        if exists (window 1) then
            keystroke (ASCII character 27)
            delay 0.5
        end if
    end tell
end tell
EOF
```

### Close All Dialogs & Reset Focus

```bash
osascript <<'EOF'
tell application "Microsoft Excel"
    activate
    delay 1
end tell
tell application "System Events"
    tell process "Microsoft Excel"
        repeat 3 times
            keystroke (ASCII character 27)
            delay 0.3
        end repeat
    end tell
end tell
EOF
```

### Force Save via AppleScript

```bash
osascript -e 'tell application "Microsoft Excel" to save active workbook'
```

### Close and Reopen Workbook

```bash
osascript <<'EOF'
tell application "Microsoft Excel"
    set wb to active workbook
    set wbPath to full name of wb
    close wb saving yes
    delay 1
    open wbPath
end tell
EOF
```

### Recovery Workflow (Python)

```python
import subprocess, time

def recover_excel():
    """Try to recover Excel from a stuck state."""
    # Step 1: Activate
    subprocess.run(['osascript', '-e',
        'tell application "Microsoft Excel" to activate'], capture_output=True)
    time.sleep(1)

    # Step 2: Dismiss any dialogs
    subprocess.run(['osascript', '-e', '''
        tell application "System Events"
            tell process "Microsoft Excel"
                repeat 3 times
                    keystroke (ASCII character 27)
                    delay 0.3
                end repeat
            end tell
        end tell
    '''], capture_output=True)
    time.sleep(1)

    # Step 3: Verify Excel is responsive
    result = subprocess.run(['osascript', '-e',
        'tell application "Microsoft Excel" to return name of active workbook'],
        capture_output=True, text=True, timeout=5)

    if result.returncode == 0:
        return True  # Excel is responsive
    else:
        return False  # Excel may need to be restarted
```

---

## Hard Rules

1. **Never use AppleScript for data writes.** Use openpyxl.
2. **Never use AppleScript for formatting.** Use openpyxl.
3. **Never use AppleScript for chart operations.** Use openpyxl.
4. **Never use UI scripting (keystroke/menu clicks) for normal operations.** Only for recovery.
5. **Always use openpyxl first.** AppleScript is for opening files in Excel and running VBA macros.
6. **Keep AppleScript scope minimal.** One action per call.
