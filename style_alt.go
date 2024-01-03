package excelize

import "strings"

func (f *File) isWindowColor(clr *xlsxColor) bool {
	if clr == nil || clr.Theme == nil || f.Theme == nil {
		return false
	}
	return *clr.Theme == 1
}

// getThemeColorIgnoreFew Color Scheme but dt1 and lt1 will be ignored
func (f *File) getThemeColorIgnoreFew(clr *xlsxColor) string {
	var RGB string
	if clr == nil || f.Theme == nil {
		return RGB
	}
	if clrScheme := f.Theme.ThemeElements.ClrScheme; clr.Theme != nil {
		if val, ok := map[int]*string{
			2: clrScheme.Lt2.SrgbClr.Val,
			3: clrScheme.Dk2.SrgbClr.Val,
			4: clrScheme.Accent1.SrgbClr.Val,
			5: clrScheme.Accent2.SrgbClr.Val,
			6: clrScheme.Accent3.SrgbClr.Val,
			7: clrScheme.Accent4.SrgbClr.Val,
			8: clrScheme.Accent5.SrgbClr.Val,
			9: clrScheme.Accent6.SrgbClr.Val,
		}[*clr.Theme]; ok && val != nil {
			return strings.TrimPrefix(ThemeColor(*val, clr.Tint), "FF")
		}
	}
	if len(clr.RGB) == 6 {
		return clr.RGB
	}
	if len(clr.RGB) == 8 {
		return strings.TrimPrefix(clr.RGB, "FF")
	}
	if f.Styles.Colors != nil && f.Styles.Colors.IndexedColors != nil && clr.Indexed < len(f.Styles.Colors.IndexedColors.RgbColor) {
		return strings.TrimPrefix(ThemeColor(strings.TrimPrefix(f.Styles.Colors.IndexedColors.RgbColor[clr.Indexed].RGB, "FF"), clr.Tint), "FF")
	}
	if clr.Indexed < len(IndexedColorMapping) {
		return strings.TrimPrefix(ThemeColor(IndexedColorMapping[clr.Indexed], clr.Tint), "FF")
	}
	return RGB
}