package main

import (
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

// newFile creates a blank in-memory excelize file for use in tests.
func newFile() *excelize.File {
	return excelize.NewFile()
}

// setCellBold sets the bold font property on a cell in the given file.
func setCellBold(t *testing.T, f *excelize.File, sheet, cell string, bold bool) {
	t.Helper()
	styleID, err := f.NewStyle(&excelize.Style{Font: &excelize.Font{Bold: bold}})
	require.NoError(t, err)
	require.NoError(t, f.SetCellStyle(sheet, cell, cell, styleID))
}

// isHighlighted reports whether the given cell in the diff file has a yellow fill.
func isHighlighted(t *testing.T, f *excelize.File, sheet, cell string) bool {
	t.Helper()
	styleID, err := f.GetCellStyle(sheet, cell)
	require.NoError(t, err)
	style, err := f.GetStyle(styleID)
	require.NoError(t, err)
	return style != nil &&
		style.Fill.Type == fillTypePattern &&
		len(style.Fill.Color) > 0 &&
		style.Fill.Color[0] == diffHighlightColor
}

// commentText returns the concatenated text of all paragraphs in the first
// comment on the given cell, or an empty string if no comment exists.
func commentText(t *testing.T, f *excelize.File, sheet, cell string) string {
	t.Helper()
	comments, err := f.GetComments(sheet)
	require.NoError(t, err)
	for _, c := range comments {
		if c.Cell == cell {
			text := ""
			for _, p := range c.Paragraph {
				text += p.Text
			}
			return text
		}
	}
	return ""
}

// TestStyleDiffs_identical verifies that no diffs are reported for equal fontProps.
func TestStyleDiffs_identical(t *testing.T) {
	fp := fontProps{bold: true, italic: false, size: 12, name: "Arial"}
	assert.Empty(t, styleDiffs(fp, fp))
}

// TestStyleDiffs_bold verifies that a bold change is detected.
func TestStyleDiffs_bold(t *testing.T) {
	diffs := styleDiffs(fontProps{bold: false}, fontProps{bold: true})
	assert.Equal(t, []string{"bold: false → true"}, diffs)
}

// TestStyleDiffs_italic verifies that an italic change is detected.
func TestStyleDiffs_italic(t *testing.T) {
	diffs := styleDiffs(fontProps{italic: false}, fontProps{italic: true})
	assert.Equal(t, []string{"italic: false → true"}, diffs)
}

// TestStyleDiffs_fontSize verifies that a font size change is detected.
func TestStyleDiffs_fontSize(t *testing.T) {
	diffs := styleDiffs(fontProps{size: 11}, fontProps{size: 14})
	assert.Equal(t, []string{"font size: 11 → 14"}, diffs)
}

// TestStyleDiffs_fontName verifies that a font name change is detected.
func TestStyleDiffs_fontName(t *testing.T) {
	diffs := styleDiffs(fontProps{name: "Arial"}, fontProps{name: "Helvetica"})
	assert.Equal(t, []string{`font name: "Arial" → "Helvetica"`}, diffs)
}

// TestStyleDiffs_multiple verifies that multiple simultaneous changes are all reported.
func TestStyleDiffs_multiple(t *testing.T) {
	a := fontProps{bold: false, italic: false, size: 11, name: "Arial"}
	b := fontProps{bold: true, italic: true, size: 14, name: "Helvetica"}
	assert.Len(t, styleDiffs(a, b), 4)
}

// TestBuildComment_changed verifies the comment when a cell value is changed.
func TestBuildComment_changed(t *testing.T) {
	assert.Equal(t, `Changed: "old" → "new"`, buildComment("old", "new", nil))
}

// TestBuildComment_added verifies the comment when a cell is present in file2 but not file1.
func TestBuildComment_added(t *testing.T) {
	assert.Equal(t, `Added: "new" (present in file2, missing in file1)`, buildComment("", "new", nil))
}

// TestBuildComment_deleted verifies the comment when a cell is present in file1 but not file2.
func TestBuildComment_deleted(t *testing.T) {
	assert.Equal(t, `Deleted: "old" (present in file1, missing in file2)`, buildComment("old", "", nil))
}

// TestBuildComment_styleOnly verifies the comment when only style differs (value is the same).
func TestBuildComment_styleOnly(t *testing.T) {
	assert.Equal(t, "Style changed: bold: false → true", buildComment("val", "val", []string{"bold: false → true"}))
}

// TestBuildComment_valueAndStyle verifies that both value and style changes appear in the comment.
func TestBuildComment_valueAndStyle(t *testing.T) {
	want := "Changed: \"old\" → \"new\"\nStyle changed: bold: false → true"
	assert.Equal(t, want, buildComment("old", "new", []string{"bold: false → true"}))
}

// TestDiffCell_identical verifies that identical cells are not highlighted.
func TestDiffCell_identical(t *testing.T) {
	f1, f2, diff := newFile(), newFile(), newFile()
	require.NoError(t, f1.SetCellValue("Sheet1", "A1", "hello"))
	require.NoError(t, f2.SetCellValue("Sheet1", "A1", "hello"))

	hasDiff, err := diffCell(f1, f2, diff, "Sheet1", "A1", "hello", "hello", true)
	assert.NoError(t, err)
	assert.False(t, hasDiff)
	assert.False(t, isHighlighted(t, diff, "Sheet1", "A1"))
}

// TestDiffCell_valueChanged verifies that a changed value is highlighted with the correct comment.
func TestDiffCell_valueChanged(t *testing.T) {
	f1, f2, diff := newFile(), newFile(), newFile()

	hasDiff, err := diffCell(f1, f2, diff, "Sheet1", "A1", "old", "new", true)
	assert.NoError(t, err)
	assert.True(t, hasDiff)
	assert.True(t, isHighlighted(t, diff, "Sheet1", "A1"))
	assert.Equal(t, `Changed: "old" → "new"`, commentText(t, diff, "Sheet1", "A1"))
}

// TestDiffCell_valueAdded verifies that a cell present only in file2 is highlighted as added.
func TestDiffCell_valueAdded(t *testing.T) {
	f1, f2, diff := newFile(), newFile(), newFile()

	hasDiff, err := diffCell(f1, f2, diff, "Sheet1", "A1", "", "new", true)
	assert.NoError(t, err)
	assert.True(t, hasDiff)
	assert.Equal(t, `Added: "new" (present in file2, missing in file1)`, commentText(t, diff, "Sheet1", "A1"))
}

// TestDiffCell_valueDeleted verifies that a cell present only in file1 is highlighted as deleted.
func TestDiffCell_valueDeleted(t *testing.T) {
	f1, f2, diff := newFile(), newFile(), newFile()

	hasDiff, err := diffCell(f1, f2, diff, "Sheet1", "A1", "old", "", true)
	assert.NoError(t, err)
	assert.True(t, hasDiff)
	assert.Equal(t, `Deleted: "old" (present in file1, missing in file2)`, commentText(t, diff, "Sheet1", "A1"))
}

// TestDiffCell_styleChanged verifies that a style-only change is highlighted.
func TestDiffCell_styleChanged(t *testing.T) {
	f1, f2, diff := newFile(), newFile(), newFile()
	setCellBold(t, f1, "Sheet1", "A1", false)
	setCellBold(t, f2, "Sheet1", "A1", true)

	hasDiff, err := diffCell(f1, f2, diff, "Sheet1", "A1", "same", "same", true)
	assert.NoError(t, err)
	assert.True(t, hasDiff)
	assert.True(t, isHighlighted(t, diff, "Sheet1", "A1"))
}

// TestDiffRow_noDiff verifies that identical rows produce no highlights.
func TestDiffRow_noDiff(t *testing.T) {
	f1, f2, diff := newFile(), newFile(), newFile()
	row := []string{"a", "b", "c"}

	hasDiff, err := diffRow(f1, f2, diff, "Sheet1", row, row, 0, true)
	assert.NoError(t, err)
	assert.False(t, hasDiff)
}

// TestDiffRow_extraColumnInFile2 verifies that an extra column in file2 is detected.
func TestDiffRow_extraColumnInFile2(t *testing.T) {
	f1, f2, diff := newFile(), newFile(), newFile()

	hasDiff, err := diffRow(f1, f2, diff, "Sheet1", []string{"a"}, []string{"a", "b"}, 0, true)
	assert.NoError(t, err)
	assert.True(t, hasDiff)
	assert.True(t, isHighlighted(t, diff, "Sheet1", "B1"))
}

// TestDiffSheet_missingSheetInFile2 verifies that all cells are marked as deleted
// when the sheet does not exist in file2.
func TestDiffSheet_missingSheetInFile2(t *testing.T) {
	f1, f2, diff := newFile(), newFile(), newFile()
	require.NoError(t, f1.SetCellValue("Sheet1", "A1", "hello"))
	require.NoError(t, f2.SetSheetName("Sheet1", "Other"))

	hasDiff, err := diffSheet(f1, f2, diff, "Sheet1")
	assert.NoError(t, err)
	assert.True(t, hasDiff)
	assert.True(t, isHighlighted(t, diff, "Sheet1", "A1"))
}

// TestDiffFiles_identical verifies that diffFiles returns false for identical files.
func TestDiffFiles_identical(t *testing.T) {
	f1, f2, diff := newFile(), newFile(), newFile()
	require.NoError(t, f1.SetCellValue("Sheet1", "A1", "hello"))
	require.NoError(t, f2.SetCellValue("Sheet1", "A1", "hello"))

	hasDiff, err := diffFiles(f1, f2, diff)
	assert.NoError(t, err)
	assert.False(t, hasDiff)
}

// TestDiffFiles_withDiff verifies that diffFiles returns true when files differ.
func TestDiffFiles_withDiff(t *testing.T) {
	f1, f2, diff := newFile(), newFile(), newFile()
	require.NoError(t, f1.SetCellValue("Sheet1", "A1", "hello"))
	require.NoError(t, f2.SetCellValue("Sheet1", "A1", "world"))

	hasDiff, err := diffFiles(f1, f2, diff)
	assert.NoError(t, err)
	assert.True(t, hasDiff)
}
