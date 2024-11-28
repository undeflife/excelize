package excelize

import "encoding/xml"

// A monkey-patched	version of sheet.go that provides a function to get all external links in the worksheet,
// that enables a more flexible way to deal with hyperlinks instead of get hyperlink cell by cell inefficiently.

// Hyperlinks returns a map contains all external links in the worksheet by giving sheet name while cell ref as key, link
// url as value
func (f *File) Hyperlinks(sheet string) (map[string]string, error) {
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
		output, _ := xml.Marshal(ws)
		f.saveFileList(name, f.replaceNameSpaceBytes(name, output))
	}
	_, xmlDecoder, _, err := f.xmlDecoder(name)
	if err != nil {
		return nil, err
	}
	links := make(map[string]string)
	for {
		token, _ := xmlDecoder.Token()
		if token == nil {
			return links, nil
		}
		switch xmlElement := token.(type) {

		case xml.StartElement:
			if xmlElement.Name.Local == "sheetData" {
				skipElement(xmlDecoder, xmlElement)
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
					links[ref] = id
				}
			}
		case xml.EndElement:
			if xmlElement.Name.Local == "hyperlinks" {
				return links, nil
			}
		}
	}
}

func skipElement(decoder *xml.Decoder, se xml.StartElement) error {
	// We will keep reading tokens until we encounter the end of this element
	for {
		t, err := decoder.Token()
		if err != nil {
			return err
		}

		switch t := t.(type) {
		case xml.StartElement:
			// Recursively skip nested elements
			if err := skipElement(decoder, t); err != nil {
				return err
			}
		case xml.EndElement:
			if t.Name.Local == se.Name.Local {
				return nil // We found the end of the current element
			}
		}
	}
}
