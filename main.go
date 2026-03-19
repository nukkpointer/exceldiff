package main

import (
	"fmt"
	"os"
	"strings"

	"github.com/xuri/excelize/v2"
)

const (
	// fillTypePattern is the excelize fill type for solid pattern fills.
	fillTypePattern = "pattern"

	// fillPatternSolid is the excelize pattern index for a solid (flat colour) fill.
	fillPatternSolid = 1

	// diffHighlightColor is the hex colour used to highlight differing cells.
	diffHighlightColor = "FFFF00" // yellow

	// commentAuthor is the author name attached to all diff comments.
	commentAuthor = "exceldiff"

	// outputFile is the name of the generated diff spreadsheet.
	outputFile = "diff.xlsx"

	// expectedArgs is the total number of os.Args entries expected (program + 2 files).
	expectedArgs = 3
)

// fontProps holds the font properties of a cell that are relevant for diffing.
type fontProps struct {
	bold   bool
	italic bool
	size   float64
	name   string
}

// yellowFill is the fill style applied to cells that differ between documents.
var yellowFill = excelize.Fill{
	Type:    fillTypePattern,
	Pattern: fillPatternSolid,
	Color:   []string{diffHighlightColor},
}

// fatal prints an error message to stderr and exits the program.
// Only called from main.
func fatal(format string, args ...any) {
	_, _ = fmt.Fprintf(os.Stderr, format+"\n", args...)
	os.Exit(1)
}

// openFile opens an xlsx file and returns it, or an error on failure.
func openFile(path string) (*excelize.File, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf("opening %s: %w", path, err)
	}
	return f, nil
}

// getFont returns the font properties of a cell in the given sheet.
// Returns a zero-value fontProps if the cell has no font style set.
func getFont(f *excelize.File, sheet, cell string) (fontProps, error) {
	styleID, err := f.GetCellStyle(sheet, cell)
	if err != nil {
		return fontProps{}, fmt.Errorf("getting cell style for %s: %w", cell, err)
	}
	style, err := f.GetStyle(styleID)
	if err != nil {
		return fontProps{}, fmt.Errorf("getting style %d: %w", styleID, err)
	}
	defaultFont, _ := f.GetDefaultFont()
	if style == nil || style.Font == nil {
		return fontProps{name: defaultFont}, nil
	}
	name := style.Font.Family
	if name == "" {
		name = defaultFont
	}
	return fontProps{
		bold:   style.Font.Bold,
		italic: style.Font.Italic,
		size:   style.Font.Size,
		name:   name,
	}, nil
}

// styleDiffs compares two fontProps and returns a human-readable description
// of each property that differs between them.
func styleDiffs(a, b fontProps) []string {
	var diffs []string
	if a.bold != b.bold {
		diffs = append(diffs, fmt.Sprintf("bold: %v → %v", a.bold, b.bold))
	}
	if a.italic != b.italic {
		diffs = append(diffs, fmt.Sprintf("italic: %v → %v", a.italic, b.italic))
	}
	if a.size != b.size && a.size != 0 && b.size != 0 {
		diffs = append(diffs, fmt.Sprintf("font size: %v → %v", a.size, b.size))
	}
	if a.name != b.name {
		diffs = append(diffs, fmt.Sprintf("font name: %q → %q", a.name, b.name))
	}
	return diffs
}

// buildComment constructs the hover comment text for a differing cell,
// describing value changes (added/deleted/changed) and style changes.
func buildComment(cell1, cell2 string, fontChanges []string) string {
	var parts []string
	if cell1 != cell2 {
		switch {
		case cell1 == "":
			parts = append(parts, fmt.Sprintf("Added: %q (present in file2, missing in file1)", cell2))
		case cell2 == "":
			parts = append(parts, fmt.Sprintf("Deleted: %q (present in file1, missing in file2)", cell1))
		default:
			parts = append(parts, fmt.Sprintf("Changed: %q → %q", cell1, cell2))
		}
	}
	if len(fontChanges) > 0 {
		parts = append(parts, "Style changed: "+strings.Join(fontChanges, ", "))
	}
	return strings.Join(parts, "\n")
}

// markCell applies a yellow background and attaches a hover comment to a cell
// in the diff file to indicate it differs between the two source documents.
func markCell(diff *excelize.File, sheetName, cellName, comment string) error {
	existingID, err := diff.GetCellStyle(sheetName, cellName)
	if err != nil {
		return fmt.Errorf("getting existing style for %s: %w", cellName, err)
	}
	style, err := diff.GetStyle(existingID)
	if err != nil {
		return fmt.Errorf("getting style %d: %w", existingID, err)
	}
	if style == nil {
		style = &excelize.Style{}
	}
	style.Fill = yellowFill
	styleID, err := diff.NewStyle(style)
	if err != nil {
		return fmt.Errorf("creating style: %w", err)
	}
	if err := diff.SetCellStyle(sheetName, cellName, cellName, styleID); err != nil {
		return fmt.Errorf("setting style for %s: %w", cellName, err)
	}
	if err := diff.AddComment(sheetName, excelize.Comment{
		Cell:   cellName,
		Author: commentAuthor,
		Paragraph: []excelize.RichTextRun{
			{Text: comment},
		},
	}); err != nil {
		return fmt.Errorf("adding comment for %s: %w", cellName, err)
	}
	return nil
}

// diffCell compares the value and font style of a single cell between f1 and f2.
// f2HasSheet must be false when the sheet does not exist in f2, in which case
// the font lookup on f2 is skipped and the cell is treated as value-only deleted.
// If any difference is found, the cell is marked in the diff file.
// Returns true if a difference was found.
func diffCell(f1, f2, diff *excelize.File, sheetName, cellName, cell1, cell2 string, f2HasSheet bool) (bool, error) {
	font1, err := getFont(f1, sheetName, cellName)
	if err != nil {
		return false, fmt.Errorf("reading font from file1 at %s: %w", cellName, err)
	}
	font2 := fontProps{}
	if f2HasSheet {
		font2, err = getFont(f2, sheetName, cellName)
		if err != nil {
			return false, fmt.Errorf("reading font from file2 at %s: %w", cellName, err)
		}
	}
	fontChanges := styleDiffs(font1, font2)
	if cell1 != cell2 || len(fontChanges) > 0 {
		if err := markCell(diff, sheetName, cellName, buildComment(cell1, cell2, fontChanges)); err != nil {
			return false, fmt.Errorf("marking cell %s: %w", cellName, err)
		}
		return true, nil
	}
	return false, nil
}

// diffRow compares all cells in a single row between f1 and f2 and marks
// any differences in the diff file.
// f2HasSheet must be false when the sheet does not exist in f2.
// Returns true if any difference was found in the row.
func diffRow(f1, f2, diff *excelize.File, sheetName string, row1, row2 []string, rowIdx int, f2HasSheet bool) (bool, error) {
	numCols := len(row1)
	if len(row2) > numCols {
		numCols = len(row2)
	}

	hasDiff := false
	for colIdx := 0; colIdx < numCols; colIdx++ {
		var cell1, cell2 string
		if colIdx < len(row1) {
			cell1 = row1[colIdx]
		}
		if colIdx < len(row2) {
			cell2 = row2[colIdx]
		}

		cellName, err := excelize.CoordinatesToCellName(colIdx+1, rowIdx+1)
		if err != nil {
			return false, fmt.Errorf("converting coordinates (%d, %d): %w", colIdx+1, rowIdx+1, err)
		}

		changed, err := diffCell(f1, f2, diff, sheetName, cellName, cell1, cell2, f2HasSheet)
		if err != nil {
			return false, fmt.Errorf("diffing cell %s: %w", cellName, err)
		}
		if changed {
			hasDiff = true
		}
	}
	return hasDiff, nil
}

// diffSheet compares all rows in a single sheet between f1 and f2 and marks
// any differences in the corresponding sheet of the diff file.
// If the sheet does not exist in f2, all cells from f1 are treated as deleted.
// Returns true if any difference was found in the sheet.
func diffSheet(f1, f2, diff *excelize.File, sheetName string) (bool, error) {
	rows1, err := f1.GetRows(sheetName)
	if err != nil {
		return false, fmt.Errorf("reading sheet %s from file1: %w", sheetName, err)
	}

	var rows2 [][]string
	sheetIdx, err := f2.GetSheetIndex(sheetName)
	if err != nil {
		return false, fmt.Errorf("looking up sheet %s in file2: %w", sheetName, err)
	}
	f2HasSheet := sheetIdx >= 0
	if f2HasSheet {
		rows2, err = f2.GetRows(sheetName)
		if err != nil {
			return false, fmt.Errorf("reading sheet %s from file2: %w", sheetName, err)
		}
	}

	numRows := len(rows1)
	if len(rows2) > numRows {
		numRows = len(rows2)
	}

	hasDiff := false
	for rowIdx := 0; rowIdx < numRows; rowIdx++ {
		var row1, row2 []string
		if rowIdx < len(rows1) {
			row1 = rows1[rowIdx]
		}
		if rowIdx < len(rows2) {
			row2 = rows2[rowIdx]
		}
		changed, err := diffRow(f1, f2, diff, sheetName, row1, row2, rowIdx, f2HasSheet)
		if err != nil {
			return false, fmt.Errorf("diffing row %d in sheet %q: %w", rowIdx+1, sheetName, err)
		}
		if changed {
			hasDiff = true
		}
	}
	return hasDiff, nil
}

// diffFiles iterates all sheets in f1 and compares them against f2,
// writing highlighted differences into the diff file.
// Returns true if any difference was found across all sheets.
func diffFiles(f1, f2, diff *excelize.File) (bool, error) {
	hasDiff := false
	for _, sheetName := range f1.GetSheetList() {
		changed, err := diffSheet(f1, f2, diff, sheetName)
		if err != nil {
			return false, fmt.Errorf("diffing sheet %q: %w", sheetName, err)
		}
		if changed {
			hasDiff = true
		}
	}
	return hasDiff, nil
}

// main parses command-line arguments, opens the input files, runs the diff,
// and writes the result to diff.xlsx.
func main() {
	if len(os.Args) != expectedArgs {
		fatal("Usage: exceldiff file1.xlsx file2.xlsx")
	}

	f1, err := openFile(os.Args[1])
	if err != nil {
		fatal("Error: %v", err)
	}
	defer func() {
		if err := f1.Close(); err != nil {
			fatal("Error closing file1: %v", err)
		}
	}()

	f2, err := openFile(os.Args[2])
	if err != nil {
		fatal("Error: %v", err)
	}
	defer func() {
		if err := f2.Close(); err != nil {
			fatal("Error closing file2: %v", err)
		}
	}()

	diff, err := openFile(os.Args[2])
	if err != nil {
		fatal("Error: %v", err)
	}
	defer func() {
		if err := diff.Close(); err != nil {
			fatal("Error closing diff file: %v", err)
		}
	}()

	hasDiff, err := diffFiles(f1, f2, diff)
	if err != nil {
		fatal("Error: %v", err)
	}

	if !hasDiff {
		fmt.Println("Files are identical.")
		os.Exit(0)
	}

	if err := diff.SaveAs(outputFile); err != nil {
		fatal("Error saving %s: %v", outputFile, err)
	}

	fmt.Printf("%s created successfully\n", outputFile)
	os.Exit(1)
}
