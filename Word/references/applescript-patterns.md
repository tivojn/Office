# AppleScript IPC — Complete Capability Reference for Microsoft Word

## Table of Contents

1. [Dual-Engine Architecture](#dual-engine-architecture)
2. [Known Quirks](#known-quirks)
3. [Document Management](#document-management)
4. [Navigation](#navigation)
5. [Reading Content](#reading-content)
6. [Modifying Text — LIVE](#modifying-text--live)
7. [Font Properties — LIVE](#font-properties--live)
8. [Paragraph Formatting — LIVE](#paragraph-formatting--live)
9. [Find and Replace — LIVE](#find-and-replace--live)
10. [Field Updates (TOC, Page Numbers)](#field-updates-toc-page-numbers)
11. [Table Operations — LIVE](#table-operations--live)
12. [Selection Operations](#selection-operations)
13. [Bookmarks](#bookmarks)
14. [Print and Export](#print-and-export)
15. [View Settings](#view-settings)
16. [Comprehensive Document Reader](#comprehensive-document-reader)
17. [Known Limitations](#known-limitations)
18. [Unit System](#unit-system)
19. [Task Classification Matrix](#task-classification-matrix)
20. [Decision Matrix](#decision-matrix)

## Dual-Engine Architecture

You have **two engines** for manipulating Word documents:

- **python-docx** (file-based): Bulk creation, paragraphs, runs, styles, tables, images, headers/footers, sections, page setup, lxml XML manipulation. Deterministic, headless, cross-platform.
- **AppleScript IPC** (live editing): Text edits, font changes, find/replace, field updates (TOC, cross-references, page numbers), PDF export, print settings — all reflected instantly.

### The Golden Workflow
```
1. python-docx  →  Create/rebuild document (heavy lifting, styles, tables, images)
2. AppleScript  →  Open the file in Word
3. AppleScript  →  Update all fields (TOC, page numbers)
4. AppleScript  →  Verify, make live tweaks
5. AppleScript  →  Save (and optionally export PDF)
```

For **edit-only** tasks on an already-open document:
```
1. AppleScript  →  Read current content
2. AppleScript  →  Make targeted edits live
3. AppleScript  →  Update fields if needed
4. AppleScript  →  Save
   (No python-docx needed! No file reload!)
```

For **finalization** tasks:
```
1. AppleScript  →  Open document
2. AppleScript  →  Update all fields
3. AppleScript  →  Export to PDF
4. AppleScript  →  Save
   (Smallest possible AppleScript scope!)
```

## Known Quirks

| Issue | Workaround |
|-------|------------|
| `content of text range` may include field codes | Filter or use `text object` |
| `insert picture` may not preserve positioning | Use python-docx `add_picture()` |
| Font color setting can be unreliable | Use python-docx for color changes |
| Large document reads can time out | Read section by section |
| AppleScript execution blocked by permission dialogs | Grant Automation permissions in System Preferences |
| `update table of contents` may fail silently | Use `update all fields in` instead |

## Document Management

```applescript
-- Open a file
tell application "Microsoft Word"
    activate
    open "/path/to/file.docx"
end tell

-- Save
tell application "Microsoft Word"
    save active document
end tell

-- Save As (new name)
tell application "Microsoft Word"
    save as active document file name "/path/to/new_file.docx"
end tell

-- Close without saving
tell application "Microsoft Word"
    close active document saving no
end tell

-- Close and reopen (after python-docx edits)
tell application "Microsoft Word"
    activate
    try
        close active document saving no
    end try
    delay 0.5
    open "/path/to/file.docx"
end tell

-- Get document info
tell application "Microsoft Word"
    set docName to name of active document
    set docPath to full name of active document
    set pageCount to count of pages of active document
    set paraCount to count of paragraphs of active document
end tell

-- Check if document is open
tell application "Microsoft Word"
    set docCount to count of documents
    if docCount > 0 then
        set docName to name of active document
    end if
end tell
```

## Navigation

```applescript
-- Go to a specific page
tell application "Microsoft Word"
    set myRange to go to what page number which go to number 3 of active document
end tell

-- Go to beginning of document
tell application "Microsoft Word"
    set view type of view of active window to normal view
    home key move selection start of story extend by moving
end tell

-- Go to end of document
tell application "Microsoft Word"
    end key move selection end of story extend by moving
end tell

-- Go to a bookmark
tell application "Microsoft Word"
    set myRange to go to what bookmark of active document bookmark name "MyBookmark"
end tell
```

## Reading Content

```applescript
-- Read full document text
tell application "Microsoft Word"
    set docText to content of text object of active document
end tell

-- Read specific paragraph
tell application "Microsoft Word"
    set paraText to content of text object of paragraph 3 of active document
end tell

-- Read paragraph count
tell application "Microsoft Word"
    set pCount to count of paragraphs of active document
end tell

-- Read a range of paragraphs
tell application "Microsoft Word"
    set output to ""
    repeat with i from 1 to 10
        set pText to content of text object of paragraph i of active document
        set output to output & i & ": " & pText & return
    end repeat
    return output
end tell

-- Read text with formatting info
tell application "Microsoft Word"
    set p to paragraph 1 of active document
    set pText to content of text object of p
    set fName to name of font object of text object of p
    set fSize to font size of font object of text object of p
    set fBold to bold of font object of text object of p
end tell
```

## Modifying Text — LIVE

```applescript
-- Replace text of a specific paragraph
tell application "Microsoft Word"
    set content of text object of paragraph 3 of active document to "New paragraph text here."
end tell

-- Insert text at beginning of document
tell application "Microsoft Word"
    set myRange to create range active document start 0 end 0
    insert text "Inserted at beginning. " at myRange
end tell

-- Insert text at end of document
tell application "Microsoft Word"
    set docEnd to end of content of text object of active document
    set myRange to create range active document start docEnd end docEnd
    insert text "Appended text." at myRange
end tell

-- Insert paragraph break
tell application "Microsoft Word"
    set myRange to create range active document start 0 end 0
    insert paragraph at myRange
end tell
```

## Font Properties — LIVE

```applescript
-- Change font of a paragraph
tell application "Microsoft Word"
    set f to font object of text object of paragraph 1 of active document
    set name of f to "Georgia"
    set font size of f to 14
    set bold of f to true
    set italic of f to false
end tell

-- Change font color (RGB packed as long int)
tell application "Microsoft Word"
    set f to font object of text object of paragraph 1 of active document
    -- Color = R * 65536 + G * 256 + B
    set color index of f to 0  -- reset to auto first
    -- For specific color, use font color property:
    -- Note: font color via AppleScript can be unreliable
    -- Prefer python-docx for color changes
end tell

-- Change underline
tell application "Microsoft Word"
    set f to font object of text object of paragraph 1 of active document
    set underline of f to underline single
    -- Options: underline none, underline single, underline double, underline dotted
end tell

-- Read font properties
tell application "Microsoft Word"
    set f to font object of text object of paragraph 1 of active document
    set fName to name of f
    set fSize to font size of f
    set fBold to bold of f
    set fItalic to italic of f
    return "Font: " & fName & " " & fSize & "pt, Bold=" & fBold & " Italic=" & fItalic
end tell
```

## Paragraph Formatting — LIVE

```applescript
-- Set paragraph alignment
tell application "Microsoft Word"
    set alignment of paragraph format of paragraph 1 of active document to align paragraph center
    -- Options: align paragraph left, align paragraph center,
    --          align paragraph right, align paragraph justify
end tell

-- Set paragraph spacing
tell application "Microsoft Word"
    set pf to paragraph format of paragraph 1 of active document
    set space before of pf to 12  -- points
    set space after of pf to 6
end tell

-- Set line spacing
tell application "Microsoft Word"
    set pf to paragraph format of paragraph 1 of active document
    set line spacing of pf to 18  -- points (exact)
    set line spacing rule of pf to line space exactly
    -- Options: line space single, line space 1pt5, line space double,
    --          line space exactly, line space at least, line space multiple
end tell

-- Set indentation
tell application "Microsoft Word"
    set pf to paragraph format of paragraph 1 of active document
    set left indent of pf to 36  -- points (0.5 inch)
    set right indent of pf to 36
    set first line indent of pf to 18  -- points
end tell
```

## Find and Replace — LIVE

```applescript
-- Simple find and replace (all occurrences)
tell application "Microsoft Word"
    tell find object of text object of active document
        clear formatting
        set content to "old text"
        tell replacement of it
            clear formatting
            set content to "new text"
        end tell
        execute find replace replace all
    end tell
end tell

-- Find and replace with formatting
tell application "Microsoft Word"
    tell find object of text object of active document
        clear formatting
        set content to "Important"
        tell replacement of it
            clear formatting
            set content to "IMPORTANT"
            set bold of font object to true
        end tell
        execute find replace replace all
    end tell
end tell

-- Find text (without replacing)
tell application "Microsoft Word"
    tell find object of selection
        clear formatting
        set content to "search term"
        set forward to true
        set wrap to find continue
        execute find
    end tell
end tell
```

## Field Updates (TOC, Page Numbers)

This is a **core use case** for AppleScript — python-docx cannot update fields.

```applescript
-- Update ALL fields in document (TOC, page numbers, cross-refs, dates)
tell application "Microsoft Word"
    -- Select all content
    set myRange to text object of active document
    -- Update fields in the range
    update fields in myRange
end tell

-- Update Table of Contents specifically
tell application "Microsoft Word"
    tell active document
        set tocCount to count of tables of contents
        if tocCount > 0 then
            update table of contents (table of contents 1)
        end if
    end tell
end tell

-- Update just page fields
tell application "Microsoft Word"
    set myRange to text object of active document
    update fields in myRange
end tell

-- Refresh after python-docx edits (close, reopen, update fields)
tell application "Microsoft Word"
    activate
    set docPath to full name of active document
    close active document saving no
    delay 0.5
    open docPath
    delay 1
    set myRange to text object of active document
    update fields in myRange
    save active document
end tell
```

## Table Operations — LIVE

```applescript
-- Read table contents
tell application "Microsoft Word"
    set t to table 1 of active document
    set rowCount to count of rows of t
    set colCount to count of columns of t
    set output to "Table: " & rowCount & " rows x " & colCount & " cols" & return
    repeat with r from 1 to rowCount
        repeat with c from 1 to colCount
            set cellText to content of text object of cell r column c of t
            set output to output & "[" & r & "," & c & "] " & cellText & " | "
        end repeat
        set output to output & return
    end repeat
    return output
end tell

-- Modify table cell text
tell application "Microsoft Word"
    set t to table 1 of active document
    set content of text object of cell 2 column 3 of t to "New Value"
end tell

-- Set cell font
tell application "Microsoft Word"
    set t to table 1 of active document
    set f to font object of text object of cell 1 column 1 of t
    set bold of f to true
    set font size of f to 12
end tell

-- Add a row to existing table
tell application "Microsoft Word"
    set t to table 1 of active document
    -- Add row at end
    set newRow to make new row at end of t
end tell
```

## Selection Operations

```applescript
-- Select all
tell application "Microsoft Word"
    select all active document
end tell

-- Select a specific paragraph
tell application "Microsoft Word"
    select text object of paragraph 5 of active document
end tell

-- Get current selection text
tell application "Microsoft Word"
    set selText to content of text object of selection
end tell

-- Type at current cursor position
tell application "Microsoft Word"
    type text selection text "Hello World"
end tell
```

## Bookmarks

```applescript
-- Add a bookmark at selection
tell application "Microsoft Word"
    set myRange to text object of selection
    make new bookmark at active document with properties {name:"MyBookmark", bookmark range:myRange}
end tell

-- Go to bookmark
tell application "Microsoft Word"
    set myRange to go to what bookmark of active document bookmark name "MyBookmark"
end tell

-- List all bookmarks
tell application "Microsoft Word"
    set bCount to count of bookmarks of active document
    set output to ""
    repeat with i from 1 to bCount
        set bName to name of bookmark i of active document
        set output to output & i & ": " & bName & return
    end repeat
    return output
end tell
```

## Print and Export

```applescript
-- Export to PDF
tell application "Microsoft Word"
    save as active document file name "/Users/me/Desktop/output.pdf" file format format PDF
end tell

-- Print document
tell application "Microsoft Word"
    print out active document
end tell

-- Print specific pages
tell application "Microsoft Word"
    print out active document pages "1-3"
end tell

-- Print settings
tell application "Microsoft Word"
    print out active document copies 2 pages "1-5"
end tell
```

## View Settings

```applescript
-- Switch to Print Layout view
tell application "Microsoft Word"
    set view type of view of active window to page view
    -- Options: normal view, page view, outline view, master view, web view
end tell

-- Set zoom
tell application "Microsoft Word"
    set percentage of zoom of active window to 125
end tell

-- Show/hide formatting marks
tell application "Microsoft Word"
    set show all of view of active window to true   -- show all marks
    set show all of view of active window to false  -- hide marks
end tell

-- Show/hide ruler
tell application "Microsoft Word"
    set display rulers of active window to true
end tell
```

## Comprehensive Document Reader

Copy-paste ready script to audit a document:

```applescript
tell application "Microsoft Word"
    set output to ""
    set docName to name of active document
    set output to output & "Document: " & docName & return

    -- Page count
    set pCount to count of pages of active document
    set output to output & "Pages: " & pCount & return

    -- Paragraph count
    set paraCount to count of paragraphs of active document
    set output to output & "Paragraphs: " & paraCount & return

    -- Table count
    set tblCount to count of tables of active document
    set output to output & "Tables: " & tblCount & return

    -- First 30 paragraphs (or all if fewer)
    set maxPara to paraCount
    if maxPara > 30 then set maxPara to 30
    set output to output & return & "=== Content (first " & maxPara & " paragraphs) ===" & return

    repeat with i from 1 to maxPara
        set pText to content of text object of paragraph i of active document
        -- Trim to 80 chars
        if length of pText > 80 then
            set pText to text 1 thru 80 of pText & "..."
        end if
        -- Get style name
        try
            set sName to name of style of paragraph i of active document
        on error
            set sName to "?"
        end try
        set output to output & i & " [" & sName & "]: " & pText & return
    end repeat

    -- Table summary
    if tblCount > 0 then
        set output to output & return & "=== Tables ===" & return
        repeat with t from 1 to tblCount
            set rowCount to count of rows of table t of active document
            set colCount to count of columns of table t of active document
            set output to output & "Table " & t & ": " & rowCount & "x" & colCount & return
        end repeat
    end if

    return output
end tell
```

## Known Limitations

| Capability | Status | Workaround |
|-----------|--------|------------|
| Insert picture from file | Works but positioning limited | Use python-docx `add_picture()` for precise positioning |
| Set font color reliably | Unreliable in some versions | Use python-docx for color changes |
| Set paragraph shading/borders | Not fully exposed | Use python-docx + lxml |
| Modify styles globally | Partial | Use python-docx `doc.styles` |
| Set page margins | Limited | Use python-docx section properties |
| Set page borders | Not exposed | Use python-docx + lxml |
| Add headers/footers | Partial (text only) | Use python-docx for complex headers |
| Create/modify sections | Limited | Use python-docx |
| Advanced table formatting | Cell text/font only | Use python-docx + lxml for borders/shading |
| Gradient/pattern fills | Not available | Use python-docx + lxml |
| Column layout | Not exposed | Use python-docx + lxml |
| Watermarks | Not directly supported | Use python-docx + lxml |
| Update TOC | **Works well** | Primary use case for AppleScript |
| Update all fields | **Works well** | Primary use case for AppleScript |
| PDF export | **Works well** | Primary use case for AppleScript |
| Find/Replace | **Works well** | Efficient for batch text changes |
| Live text editing | **Works well** | Fast for targeted changes |

## Unit System

```
AppleScript for Word uses POINTS (1 inch = 72 points)
python-docx uses EMU internally (1 inch = 914400 EMUs)
python-docx exposes Inches(), Pt(), Cm() helpers
OOXML XML uses DXA/twips for many attributes (1 inch = 1440 twips)

Conversion:
  EMU = points * 12700
  points = EMU / 12700
  twips = inches * 1440
  twips = points * 20

Page dimensions in points:
  Letter: 612 x 792
  A4: 595 x 842
```

## Task Classification Matrix

For every task, classify it as python-docx or AppleScript:

| Task | Tool | Justification |
|------|------|---------------|
| Create new document from scratch | **python-docx** | File-level, no Word needed |
| Fill template with text | **python-docx** | Deterministic, headless |
| Add/modify paragraphs and runs | **python-docx** | Core text editing |
| Insert/modify tables | **python-docx** | Better API, precise control |
| Insert inline images | **python-docx** | `add_picture()` reliable |
| Set styles (heading, body, etc.) | **python-docx** | Style API, global changes |
| Add/edit headers and footers | **python-docx** | Section-level, precise |
| Set page margins/orientation | **python-docx** | Section properties |
| Add paragraph borders/shading | **python-docx** | lxml for complex formatting |
| Add page borders | **python-docx** | lxml section properties |
| Add TOC placeholder | **python-docx** | Field code insertion |
| Batch document generation | **python-docx** | Headless, repeatable |
| Cross-platform execution | **python-docx** | No Word dependency |
| **Update fields (TOC, page numbers)** | **AppleScript** | Requires Word's engine |
| **Update Table of Contents** | **AppleScript** | Requires Word's engine |
| **Export to PDF** | **AppleScript** | Word's rendering engine |
| Change text in open document | **AppleScript** | Instant, no reload |
| Change font properties live | **AppleScript** | Instant feedback |
| Find and replace text | **AppleScript** | Efficient batch operation |
| Navigate to sections | **AppleScript** | Cursor/view control |
| Read document content | **AppleScript** | Quick inspection |
| Print document | **AppleScript** | Requires Word |
| View/zoom settings | **AppleScript** | Application-level |
| Handle content controls | **AppleScript** | Requires Word |

### Split Tasks

If a task needs both engines:
1. **python-docx first** — create/modify content, styles, layout
2. **AppleScript second** — open, update fields, verify, export

Example: "Create a report with TOC and export to PDF"
1. python-docx: Create document, add content, add TOC placeholder
2. AppleScript: Open file, update TOC field, export PDF, save

## Decision Matrix

| User Request | Engine | Why |
|-------------|--------|-----|
| "Create a quarterly report" | **python-docx** | Full document creation |
| "Build a proposal with cover page" | **python-docx** | Complex layout, images |
| "Add a table of financial data" | **python-docx** | Table API with styling |
| "Insert company logo in header" | **python-docx** | Image in header |
| "Change margins to 0.75 inches" | **python-docx** | Section properties |
| "Add page borders" | **python-docx** | lxml page borders |
| "Style all headings dark blue" | **python-docx** | Modify style objects |
| "Generate 50 form letters" | **python-docx** | Batch, headless |
| "Update the table of contents" | **AppleScript** | Field update |
| "Export to PDF" | **AppleScript** | Word rendering |
| "Change 'Q3' to 'Q4' everywhere" | **AppleScript** | Find/replace |
| "Make the title bigger" | **AppleScript** | Quick font change |
| "What's on page 3?" | **AppleScript** | Quick read |
| "Fix the typo in paragraph 5" | **AppleScript** | Targeted text edit |
| "Print 3 copies" | **AppleScript** | Print control |
| "Switch to Print Layout view" | **AppleScript** | View settings |
| "Redesign the entire document" | **python-docx** | Complex rebuild |
| "Add styled callout boxes" | **python-docx** | lxml borders/shading |
| "Create report then export PDF" | **Both** | python-docx builds, AppleScript finalizes |
