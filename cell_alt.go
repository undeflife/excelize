package excelize

// A monkey-patched version of cell.go that provides a function to get formatted rich text for the cell from shardString
// and apply global stylesheet if there is one associated to it.

import (
	"strconv"
	"strings"
)

// getValueFrom return a value from a column/row cell, this function is
// intended to be used with for range on rows an argument with the spreadsheet
// opened file.
func (c *xlsxC) getRichValueFrom(f *File, d *xlsxSST, raw bool) (RichText, error) {
	switch c.T {
	case "b":
		b, err := c.getCellBool(f, raw)
		if err != nil {
			return RichText{}, err
		}
		return newRichText(b), nil
	case "d":
		date, err := c.getCellDate(f, raw)
		if err != nil {
			return RichText{}, err
		}
		return newRichText(date), nil
	case "s":
		if c.V != "" {
			xlsxSI, _ := strconv.Atoi(strings.TrimSpace(c.V))
			//NOTE: Only when create shared string on the fly
			//if _, ok := f.tempFiles.Load(defaultXMLPathSharedStrings); ok {
			//	return f.formattedValue(&xlsxC{S: c.S, V: f.getFromStringItem(xlsxSI)}, raw, CellTypeSharedString)
			//}
			d.mu.Lock()
			defer d.mu.Unlock()
			if len(d.SI) > xlsxSI {
				shared := d.SI[xlsxSI]
				if shared.R == nil {
					val, err := f.formattedRichValue(&xlsxC{S: c.S, V: shared.String()}, raw, CellTypeSharedString)
					return val, err
				} else {
					return RichText{f.getCellRichText(&shared)}, nil
				}
			}
		}
		val, err := f.formattedValue(c, raw, CellTypeSharedString)
		return newRichText(val), err
	case "inlineStr":
		if c.IS != nil {
			val, err := f.formattedValue(&xlsxC{S: c.S, V: c.IS.String()}, raw, CellTypeInlineString)
			return newRichText(val), err
		}
		val, err := f.formattedValue(c, raw, CellTypeInlineString)
		return newRichText(val), err
	default:
		if isNum, precision, decimal := isNumeric(c.V); isNum && !raw {
			if precision > 15 {
				c.V = strconv.FormatFloat(decimal, 'G', 15, 64)
			} else {
				c.V = strconv.FormatFloat(decimal, 'f', -1, 64)
			}
		}
		val, err := f.formattedValue(c, raw, CellTypeNumber)
		return newRichText(val), err
	}
}

// formattedValue provides a function to returns a value after formatted. If
// it is possible to apply a format to the cell value, it will do so, if not
// then an error will be returned, along with the raw value of the cell.
func (f *File) formattedRichValue(c *xlsxC, raw bool, cellType CellType) (RichText, error) {
	if raw || c.S == 0 {
		return newRichText(c.V), nil
	}
	styleSheet, err := f.stylesReader()
	if err != nil {
		return newRichText(c.V), err
	}
	if styleSheet.CellXfs == nil {
		return newRichText(c.V), err
	}
	if c.S >= len(styleSheet.CellXfs.Xf) || c.S < 0 {
		return newRichText(c.V), err
	}
	var numFmtID int
	if styleSheet.CellXfs.Xf[c.S].NumFmtID != nil {
		numFmtID = *styleSheet.CellXfs.Xf[c.S].NumFmtID
	}
	var font *xlsxFont
	if styleSheet.CellXfs.Xf[c.S].FontID != nil {
		font = styleSheet.Fonts.Font[*styleSheet.CellXfs.Xf[c.S].FontID]
		// replace color rgb 
		if font.Color != nil && !f.isWindowColor(font.Color) {
			font.Color.RGB = f.getThemeColor(font.Color)
		}
	}

	date1904 := false
	wb, err := f.workbookReader()
	if err != nil {
		return newRichText(c.V).apply(font), err
	}
	if wb != nil && wb.WorkbookPr != nil {
		date1904 = wb.WorkbookPr.Date1904
	}
	if fmtCode, ok := styleSheet.getCustomNumFmtCode(numFmtID); ok {
		return newRichText(format(c.V, fmtCode, date1904, cellType, f.options)).apply(font), err
	}
	return newRichText(c.V).apply(font), err
}

// getCellRichText returns rich text of cell by given string item that will apply theme color to Font int RichTextRun
func (f *File) getCellRichText(si *xlsxSI) (runs []RichTextRun) {
	for _, v := range si.R {
		run := RichTextRun{
			Text: v.T.Val,
		}
		if v.RPr != nil {
			run.Font = newFont(v.RPr)
			if run.Font.Color == "" && !f.isWindowColor(v.RPr.Color) {
				run.Font.Color = f.getThemeColor(v.RPr.Color)
			}
		}
		runs = append(runs, run)
	}
	return
}
