package excelize

// A monkey-patched version of rows.go that provides a function to streaming deserialize sheetData and treat all values
// as rich text.

import (
	"encoding/xml"
	"os"
)

type RichText struct {
	Runs []RichTextRun
}

// rowRichIterator defined runtime use field for the worksheet row SAX parser.
type rowRichIterator struct {
	err              error
	inElement        string
	cellCol, cellRow int
	cells            []RichText
}

func newRichText(val string) RichText {
	return RichText{
		[]RichTextRun{
			{Text: val},
		},
	}
}

func (r RichText) apply(font *xlsxFont) RichText {
	if font == nil {
		return r
	}
	for i := 0; i < len(r.Runs); i++ {
		run := &r.Runs[i]
		if run.Font == nil {
			run.Font = new(Font)
		}
		if font.B != nil {
			run.Font.Bold = true
		}
		if font.I != nil {
			run.Font.Italic = true
		}
		if font.Strike != nil {
			run.Font.Strike = true
		}
		if font.U != nil && font.U.Val != nil {
			run.Font.Underline = "single"
		}
		if font.Color != nil && font.Color.RGB != "" {
			run.Font.Color = font.Color.RGB
		}
	}
	return r
}

func (r RichText) isEmpty() bool {
	return r.Runs == nil || (r.Runs[0].Text == "" && r.Runs[0].Font == nil)
}

// Values return the current row's column values. This fetches the worksheet
// data as a stream, returns each cell in a row as is, and will not skip empty
// rows in the tail of the worksheet.
func (rows *AltRows) Values() ([]RichText, error) {
	if rows.curRow > rows.seekRow {
		return nil, nil
	}
	var rowIterator rowRichIterator
	var token xml.Token
	if rows.sst, rowIterator.err = rows.f.sharedStringsReader(); rowIterator.err != nil {
		return nil, rowIterator.err
	}
	for {
		if rows.token != nil {
			token = rows.token
		} else if token, _ = rows.decoder.Token(); token == nil {
			break
		}
		switch xmlElement := token.(type) {
		case xml.StartElement:
			rowIterator.inElement = xmlElement.Name.Local
			if rowIterator.inElement == "row" {
				rowNum := 0
				if rowNum, rowIterator.err = attrValToInt("r", xmlElement.Attr); rowNum != 0 {
					rows.curRow = rowNum
				} else if rows.token == nil {
					rows.curRow++
				}
				rows.token = token
				rows.seekRowOpts = extractRowOpts(xmlElement.Attr)
				if rows.curRow > rows.seekRow {
					rows.token = nil
					return rowIterator.cells, rowIterator.err
				}
			}
			if rows.rowRichHandler(&rowIterator, &xmlElement, rows.rawCellValue); rowIterator.err != nil {
				rows.token = nil
				return rowIterator.cells, rowIterator.err
			}
			rows.token = nil
		case xml.EndElement:
			if xmlElement.Name.Local == "worksheet" {
				return rowIterator.cells, rowIterator.err
			}
		}
	}
	return rowIterator.cells, rowIterator.err
}

// rowXMLHandler parse the row XML element of the worksheet.
func (rows *AltRows) rowRichHandler(rowIterator *rowRichIterator, xmlElement *xml.StartElement, raw bool) {
	if rowIterator.inElement == "c" {
		rowIterator.cellCol++
		colCell := xlsxC{}
		_ = rows.decoder.DecodeElement(&colCell, xmlElement)
		if colCell.R != "" {
			if rowIterator.cellCol, _, rowIterator.err = CellNameToCoordinates(colCell.R); rowIterator.err != nil {
				return
			}
		}
		blank := rowIterator.cellCol - len(rowIterator.cells)
		if val, _ := colCell.getRichValueFrom(rows.f, rows.sst, raw); !val.isEmpty() || colCell.F != nil {
			rowIterator.cells = append(appendSpaceToRich(blank, rowIterator.cells), val)
		}
	}
}

// appendSpace append blank characters to slice by given length and source slice.
func appendSpaceToRich(l int, s []RichText) []RichText {
	for i := 1; i < l; i++ {
		s = append(s, RichText{[]RichTextRun{
			{
				Text: "",
			},
		}})
	}
	return s
}

type AltRows struct {
	err                     error
	curRow, seekRow         int
	needClose, rawCellValue bool
	sheet                   string
	f                       *File
	tempFile                *os.File
	sst                     *xlsxSST
	decoder                 *xml.Decoder
	token                   xml.Token
	curRowOpts, seekRowOpts RowOpts
	links                   map[string]string
}

func (f *File) AltRows(sheet string) (*AltRows, error) {
	if err := checkSheetName(sheet); err != nil {
		return nil, err
	}
	name, ok := f.getSheetXMLPath(sheet)
	if !ok {
		return nil, ErrSheetNotExist{sheet}
	}
	if worksheet, ok := f.Sheet.Load(name); ok && worksheet != nil {
		ws := worksheet.(*xlsxWorksheet)
		ws.mu.Lock()
		defer ws.mu.Unlock()
		// Flush data
		output, _ := xml.Marshal(ws)
		f.saveFileList(name, f.replaceNameSpaceBytes(name, output))
	}
	var err error
	rows := AltRows{f: f, sheet: name}
	rows.needClose, rows.decoder, rows.tempFile, err = f.xmlDecoder(name)
	rows.links = map[string]string{}
	return &rows, err
}

func (rows *AltRows) Next() bool {
	rows.seekRow++
	if rows.curRow >= rows.seekRow {
		rows.curRowOpts = rows.seekRowOpts
		return true
	}
	for {
		token, _ := rows.decoder.Token()
		if token == nil {
			return false
		}
		switch xmlElement := token.(type) {
		case xml.StartElement:
			if xmlElement.Name.Local == "row" {
				rows.curRow++
				if rowNum, _ := attrValToInt("r", xmlElement.Attr); rowNum != 0 {
					rows.curRow = rowNum
				}
				rows.token = token
				rows.curRowOpts = extractRowOpts(xmlElement.Attr)
				return true
			} else if xmlElement.Name.Local == "hyperlink" {
				var ref, id string
				for _, attr := range xmlElement.Attr {
					if attr.Name.Local == "ref" && ref == "" {
						ref = attr.Value
					}
					if attr.Name.Local == "id" && id == "" {
						id = attr.Value
					}
				}
				if ref != "" && id != "" {
					rows.links[ref] = id
				}
			}
		case xml.EndElement:
			if xmlElement.Name.Local == "worksheet" {
				return false
			}
		}
	}
}

func (rows *AltRows) GetHyperlinks() map[string]string {
	return rows.links
}
