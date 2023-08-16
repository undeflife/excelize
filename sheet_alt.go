package excelize

// A monkey-patched	version of sheet.go that provides a function to get all external links in the worksheet,
// that enables a more flexible way to deal with hyperlinks instead of get hyperlink cell by cell inefficiently.

// Hyperlinks returns a map contains all external links in the worksheet by giving sheet name while cell ref as key, link
// url as value
func (f *File) Hyperlinks(sheet string) (map[string]string, error) {
	links := make(map[string]string)
	ws, err := f.workSheetReader(sheet)
	if err != nil {
		return links, err
	}
	if ws.Hyperlinks != nil {
		for _, link := range ws.Hyperlinks.Hyperlink {
			if link.RID != "" {
				links[link.Ref] = f.getSheetRelationshipsTargetByID(sheet, link.RID)
			}
		}
	}
	return links, nil
}
