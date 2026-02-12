# AppleScript IPC — Complete Capability Reference

## Table of Contents

1. [Dual-Engine Architecture](#dual-engine-architecture)
2. [Known Quirks](#known-quirks)
3. [Presentation Management](#presentation-management)
4. [Slide Operations](#slide-operations)
5. [Reading Shapes](#reading-shapes)
6. [Modifying Text — LIVE](#modifying-text--live)
7. [Font Properties — LIVE](#font-properties--live)
8. [Paragraph Alignment — LIVE](#paragraph-alignment--live)
9. [Text Frame Properties — LIVE](#text-frame-properties--live)
10. [Shape Position & Size — LIVE](#shape-position--size--live)
11. [Shape Rotation — LIVE](#shape-rotation--live)
12. [Adding Shapes — LIVE](#adding-shapes--live)
13. [Shape Fill & Line — LIVE](#shape-fill--line--live)
14. [Shadow Effects — LIVE](#shadow-effects--live)
15. [Shape Z-Order — LIVE](#shape-z-order--live)
16. [Shape Visibility — LIVE](#shape-visibility--live)
17. [Shape Naming — LIVE](#shape-naming--live)
18. [Duplicate & Delete — LIVE](#duplicate--delete--live)
19. [Speaker Notes — LIVE](#speaker-notes--live)
20. [Slide Background Color — LIVE](#slide-background-color--live)
21. [Selection Info](#selection-info)
22. [Comprehensive Slide Reader](#comprehensive-slide-reader)
23. [Known Limitations](#known-limitations)
24. [Unit System](#unit-system)
25. [Decision Matrix](#decision-matrix)

## Dual-Engine Architecture

You have **two engines** for manipulating PowerPoint:

- **python-pptx** (file-based): Bulk creation, complex formatting (gradients, corner radius, letter spacing), images, charts, tables, lxml XML manipulation.
- **AppleScript IPC** (live editing): Real-time text edits, font changes, position/size, fill colors, z-order, visibility, rotation, shadows, speaker notes, slide management — all reflected instantly.

### The Golden Workflow
```
1. python-pptx  →  Create/rebuild slides (heavy lifting, images, charts)
2. AppleScript  →  Open the file in PowerPoint
3. AppleScript  →  Navigate, verify, make live tweaks
4. AppleScript  →  Save when done
```

For **edit-only** tasks on an already-open presentation:
```
1. AppleScript  →  Read current state (shapes, text, positions)
2. AppleScript  →  Make targeted edits live
3. AppleScript  →  Save
   (No python-pptx needed! No file reload!)
```

## Known Quirks

| Issue | Workaround |
|-------|------------|
| `top of shape` throws access error (-10003) | Avoid bare `top` — use `set t to top of s` with caution or read via python-pptx |
| `saving no` syntax varies by version | Wrap in `try` block |
| `make new picture` fails | Use python-pptx `add_picture()` instead |
| Font color setting unreliable | Use python-pptx for color changes |
| Navigation timing | Add `delay` between operations |

## Presentation Management

```applescript
-- Open a file
tell application "Microsoft PowerPoint"
    activate
    open POSIX file "/path/to/file.pptx"
end tell

-- Save
tell application "Microsoft PowerPoint"
    save active presentation
end tell

-- Export as PDF
tell application "Microsoft PowerPoint"
    save active presentation in ((path to downloads folder as text) & "output.pdf") as save as PDF
end tell

-- Close without saving
tell application "Microsoft PowerPoint"
    close active presentation saving no
end tell

-- Get presentation info
tell application "Microsoft PowerPoint"
    set pName to name of active presentation
    set sCount to count of slides of active presentation
    set sWidth to width of page setup of active presentation
    set sHeight to height of page setup of active presentation
end tell

-- Close and reopen (for refreshing after python-pptx edits)
tell application "Microsoft PowerPoint"
    activate
    try
        close active presentation saving no
    end try
    delay 0.5
    open POSIX file "/path/to/file.pptx"
end tell
```

## Slide Operations

```applescript
-- Navigate to slide
tell application "Microsoft PowerPoint"
    set theView to view of active window
    go to slide theView number 3
end tell

-- Add a new blank slide
tell application "Microsoft PowerPoint"
    set newSlide to make new slide at end of active presentation
end tell

-- Add slide at specific position
tell application "Microsoft PowerPoint"
    set newSlide to make new slide at before slide 3 of active presentation
end tell

-- Delete a slide
tell application "Microsoft PowerPoint"
    delete slide 5 of active presentation
end tell

-- Duplicate a slide
tell application "Microsoft PowerPoint"
    tell active presentation
        set dupSlide to duplicate slide 1
    end tell
end tell

-- Reorder slides (move slide 10 before slide 2)
tell application "Microsoft PowerPoint"
    tell active presentation
        move slide 10 to before slide 2
    end tell
end tell

-- Get slide count
tell application "Microsoft PowerPoint"
    return count of slides of active presentation
end tell

-- Get slide layout name
tell application "Microsoft PowerPoint"
    tell slide 1 of active presentation
        set ln to name of slide layout of it
    end tell
end tell
```

## Slideshow Control

```applescript
-- Start slideshow from beginning
tell application "Microsoft PowerPoint"
    activate
    run slide show slide show settings of active presentation
end tell

-- FALLBACK: If "run slide show" fails with Parameter error (-50),
-- use System Events to click the Slide Show menu.
-- This is MORE RELIABLE, especially when restarting a slideshow
-- or when PowerPoint is in an unexpected state.
tell application "Microsoft PowerPoint"
    activate
    delay 0.3
end tell
tell application "System Events"
    tell process "Microsoft PowerPoint"
        click menu item "Play from Start" of menu "Slide Show" of menu bar 1
    end tell
end tell

-- Exit current slideshow
tell application "Microsoft PowerPoint"
    try
        set theShow to slide show window 1
        tell theShow to exit slide show
    end try
end tell
```

> **Known issue:** `run slide show` can fail with "Parameter error (-50)" if a slideshow was recently exited or PowerPoint is in a transitional state. Always prefer the System Events menu fallback for reliability. Wrap `run slide show` in a `try` block when using it.

## Reading Shapes

```applescript
-- List all shapes on a slide with names and text
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set shapeCount to count of shapes
        set output to ""
        repeat with i from 1 to shapeCount
            set s to shape i
            set sName to name of s
            set sLeft to left position of s
            set sW to width of s
            set sH to height of s
            set output to output & i & ": [" & sName & "] left=" & sLeft & " size(" & sW & "x" & sH & ")"
            try
                if has text frame of s then
                    set theText to content of text range of text frame of s
                    if length of theText > 0 then
                        set output to output & " → \"" & theText & "\""
                    end if
                end if
            end try
            set output to output & return
        end repeat
        return output
    end tell
end tell

-- Access shape by name
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set s to shape "TextBox 4"
        set t to content of text range of text frame of s
    end tell
end tell
```

## Modifying Text — LIVE

```applescript
-- Replace full text of a shape
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set s to shape "TextBox 6"
        set content of text range of text frame of s to "New text here!"
    end tell
end tell

-- Modify a specific paragraph
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set s to shape "TextBox 9"
        set tf to text frame of s
        set tr to text range of tf
        set pCount to count of paragraphs of tr
        -- Modify paragraph 2 only
        set content of paragraph 2 of tr to "Modified line 2"
    end tell
end tell

-- Batch text update (multiple shapes in one call)
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set content of text range of text frame of shape "TextBox 4" to "New Title"
        set content of text range of text frame of shape "TextBox 6" to "New Subtitle"
    end tell
    save active presentation
end tell
```

## Font Properties — LIVE

```applescript
-- Read font properties
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set s to shape "TextBox 4"
        set tr to text range of text frame of s
        set f to font of tr
        set fs to font size of f    -- e.g., 30.0
        set fb to bold of f         -- e.g., true
    end tell
end tell

-- Change font size
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set f to font of text range of text frame of shape "TextBox 4"
        set font size of f to 24
    end tell
end tell

-- Change bold
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set f to font of text range of text frame of shape "TextBox 4"
        set bold of f to true
    end tell
end tell

-- Change italic
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set f to font of text range of text frame of shape "TextBox 4"
        set italic of f to true
    end tell
end tell

-- Change underline
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set f to font of text range of text frame of shape "TextBox 4"
        set underline of f to true
    end tell
end tell

-- Change font name
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set f to font of text range of text frame of shape "TextBox 4"
        set font name of f to "Montserrat"
    end tell
end tell

-- Note: font COLOR setting via AppleScript has quirks.
-- For reliable color changes, use python-pptx instead.
```

## Paragraph Alignment — LIVE

```applescript
-- Read alignment
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set s to shape "TextBox 4"
        set p1 to paragraph 1 of text range of text frame of s
        set al to alignment of paragraph format of p1
        -- Returns: paragraph align left / center / right
    end tell
end tell

-- Set alignment
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set p1 to paragraph 1 of text range of text frame of shape "TextBox 4"
        set alignment of paragraph format of p1 to paragraph align center
        -- Options: paragraph align left, paragraph align center,
        --          paragraph align right, paragraph align justify
    end tell
end tell
```

## Text Frame Properties — LIVE

```applescript
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set s to shape "TextBox 4"
        set tf to text frame of s
        set wa to word wrap of tf          -- true/false
        set an to auto size of tf          -- auto size enum
        set ml to margin left of tf        -- in points
        set mt to margin top of tf
        set mr to margin right of tf
        set mb to margin bottom of tf
    end tell
end tell

-- Set text frame properties
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set tf to text frame of shape "TextBox 4"
        set word wrap of tf to true
        set margin left of tf to 10
        set margin top of tf to 5
    end tell
end tell
```

## Shape Position & Size — LIVE

```applescript
-- Read position and size
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set s to shape "TextBox 6"
        set l to left position of s
        set w to width of s
        set h to height of s
    end tell
end tell

-- Move a shape
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set s to shape "TextBox 6"
        set left position of s to 200
    end tell
end tell

-- Resize a shape
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set s to shape "TextBox 6"
        set width of s to 500
        set height of s to 100
    end tell
end tell
```

> **Important:** AppleScript positions are in **points** (1 inch = 72 points), not EMUs.
> Conversion: `EMU = points * 12700`, `points = EMU / 12700`.

## Shape Rotation — LIVE

```applescript
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set s to shape "TextBox 4"
        set rotation of s to 15    -- degrees
        set rotation of s to 0     -- reset
    end tell
end tell
```

## Adding Shapes — LIVE

```applescript
-- Add a rectangle (appears instantly!)
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set sh to make new shape at end with properties ¬
            {left position:100, top:100, width:200, height:100}
        -- Set fill color (RGB packed as long int: R*65536 + G*256 + B)
        set fore color of fill format of sh to 200 * 65536 + 100 * 256 + 50
        -- Set transparency (0.0 = opaque, 1.0 = fully transparent)
        set transparency of fill format of sh to 0.3
    end tell
end tell

-- Add a text box
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set tb to make new text box at end with properties ¬
            {left position:100, top:100, width:300, height:50}
        set content of text range of text frame of tb to "Live text!"
        set font size of font of text range of text frame of tb to 18
        set bold of font of text range of text frame of tb to true
    end tell
end tell
```

## Shape Fill & Line — LIVE

```applescript
-- Set fill color (RGB as long int: R * 65536 + G * 256 + B)
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set sh to shape "Shape_1"
        set fore color of fill format of sh to 26 * 65536 + 51 * 256 + 88
        set transparency of fill format of sh to 0.2
    end tell
end tell

-- Set line/border
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set sh to shape "Shape_1"
        tell line format of sh
            set fore color to 255 * 65536 + 0 * 256 + 0
            set weight to 2.0
        end tell
    end tell
end tell
```

> **Color encoding:** `colorValue = Red * 65536 + Green * 256 + Blue`
> Example: `#C9A84C` = `201 * 65536 + 168 * 256 + 76` = `13215820`
>
> Common colors:
> - White `#FFFFFF`: `16777215`
> - Black `#000000`: `0`
> - Red `#FF0000`: `16711680`
> - Gold `#C9A84C`: `13215820`
> - Blue `#3B82F6`: `3899126`

## Shadow Effects — LIVE

```applescript
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set sh to shape "Shape_1"
        tell shadow format of sh
            set visible to true
            set blur to 5
            set offset x to 3
            set offset y to 3
        end tell
    end tell
end tell
```

## Shape Z-Order — LIVE

```applescript
-- Bring shape to front
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set s to shape "TextBox 4"
        set z order of s to bring to front
        -- Options: bring to front, send to back, bring forward, send backward
    end tell
end tell
```

## Shape Visibility — LIVE

```applescript
-- Hide a shape
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set visible of shape "TextBox 4" to false
    end tell
end tell

-- Show a shape
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set visible of shape "TextBox 4" to true
    end tell
end tell
```

## Shape Naming — LIVE

```applescript
-- Rename a shape (useful for scripting later)
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set name of shape 3 to "MyCustomName"
    end tell
end tell
```

## Duplicate & Delete — LIVE

```applescript
-- Duplicate a shape (creates copy on same slide)
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set s to shape "TextBox 4"
        set dup to duplicate s
        set left position of dup to 400
    end tell
end tell

-- Delete a shape
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        delete shape "TextBox 36"
        -- Or by index:
        delete shape (count of shapes)
    end tell
end tell
```

## Speaker Notes — LIVE

```applescript
-- Write speaker notes
tell application "Microsoft PowerPoint"
    tell slide 1 of active presentation
        set np to notes page of it
        set tr to text range of text frame of shape 2 of np
        set content of tr to "Remember to mention the budget breakdown here."
    end tell
end tell

-- Read speaker notes
tell application "Microsoft PowerPoint"
    tell slide 1 of active presentation
        set np to notes page of it
        set noteText to content of text range of text frame of shape 2 of np
    end tell
end tell
```

## Slide Background Color — LIVE

```applescript
tell application "Microsoft PowerPoint"
    tell slide 1 of active presentation
        set fore color of fill format of background of it to 11 * 65536 + 29 * 256 + 58
    end tell
end tell
```

## Selection Info

```applescript
tell application "Microsoft PowerPoint"
    set sel to selection of active window
    set selType to selection type of sel
    -- Returns: selection type slides, selection type shapes, etc.
end tell
```

## Comprehensive Slide Reader

Copy-paste ready script to audit all slides:

```applescript
tell application "Microsoft PowerPoint"
    set output to ""
    set sCount to count of slides of active presentation
    repeat with slideIdx from 1 to sCount
        set output to output & return & "=== Slide " & slideIdx & " ===" & return
        tell slide slideIdx of active presentation
            set shapeCount to count of shapes
            set output to output & "Shapes: " & shapeCount & return
            repeat with i from 1 to shapeCount
                set s to shape i
                set sName to name of s
                set sLeft to left position of s
                set sW to width of s
                set sH to height of s
                set output to output & "  " & i & ": [" & sName & "] left=" & sLeft & " size(" & sW & "x" & sH & ")"
                try
                    if has text frame of s then
                        set theText to content of text range of text frame of s
                        if length of theText > 0 then
                            if length of theText > 60 then
                                set theText to text 1 thru 60 of theText
                            end if
                            set output to output & " → \"" & theText & "\""
                        end if
                    end if
                end try
                set output to output & return
            end repeat
        end tell
    end repeat
    return output
end tell
```

## Known Limitations

| Capability | Status | Workaround |
|-----------|--------|------------|
| Insert picture from file | Not supported (`make new picture` fails) | Use python-pptx `add_picture()` then reload |
| Set font color | Unreliable (`fore color of font color` syntax issues) | Use python-pptx for color changes |
| Copy shape across slides | Not supported (`copy` not understood) | Duplicate on same slide + set props, or python-pptx |
| Gradient backgrounds | Not exposed in AS dictionary | Use python-pptx + lxml |
| Letter spacing (tracking) | Not exposed | Use python-pptx + lxml `rPr.set('spc', '300')` |
| Corner radius on rounded rects | Not exposed | Use python-pptx + lxml |
| Table cell-level styling | Partial — text works, styling limited | Use python-pptx for tables |
| Chart creation/editing | Not exposed | Use python-pptx chart API |
| Slide transitions | Read-only (duration) | Use python-pptx for transitions |
| `top of shape` property | Throws -10003 on some versions | Read via python-pptx or use workaround |

## Unit System

```
AppleScript uses POINTS (1 inch = 72 points)
python-pptx uses EMUs (1 inch = 914400 EMUs)

Conversion:
  EMU = points * 12700
  points = EMU / 12700

Slide dimensions in points: 720 x 540
Slide dimensions in EMUs: 12192000 x 6858000
```

## Decision Matrix

| User Request | Engine | Why |
|-------------|--------|-----|
| "Change the title on slide 3" | **AppleScript** | Simple text edit, instant |
| "Make the subtitle bigger" | **AppleScript** | Font size change, live |
| "Move that card to the right" | **AppleScript** | Position change, live |
| "Add speaker notes to all slides" | **AppleScript** | Notes access, live |
| "Delete slide 8" | **AppleScript** | Quick slide management |
| "Reorder slides 3 and 7" | **AppleScript** | Live slide reordering |
| "What's on slide 4?" | **AppleScript** | Quick read, no file I/O |
| "Duplicate slide 2" | **AppleScript** | Live duplication |
| "Add a simple colored box" | **AppleScript** | Basic shape, live |
| "Make the title bold and centered" | **AppleScript** | Font + alignment, live |
| "Hide that shape" | **AppleScript** | Visibility toggle, live |
| "Bring that shape to front" | **AppleScript** | Z-order, live |
| "Rename shape for scripting" | **AppleScript** | Shape naming, live |
| "Add a shadow to the card" | **AppleScript** | Shadow effects, live |
| "Set the border to 2pt red" | **AppleScript** | Line format, live |
| "Rotate the arrow 45 degrees" | **AppleScript** | Rotation, live |
| "Set slide background to dark blue" | **AppleScript** | Solid bg color, live |
| "Start slideshow" | **AppleScript** | System Events menu fallback if `run slide show` fails |
| "Build me a 10-slide deck" | **python-pptx** | Bulk creation |
| "Add a background image to slide 1" | **python-pptx** | Image insertion |
| "Change the color scheme" | **python-pptx** | Gradient/fill/transparency via lxml |
| "Add a chart showing Q1-Q4 revenue" | **python-pptx** | Chart API |
| "Add letter spacing to all titles" | **python-pptx** | lxml XML manipulation |
| "Redesign slide 5 completely" | **python-pptx** | Complex rebuild |
| "Add a rounded rectangle with gradient" | **python-pptx** | Gradient + corner radius need lxml |
| "Add a table with styled cells" | **python-pptx** | Table API + cell styling |
