package excelize

import "strings"

// getColorScheme according to specification L.4.3.2.3 Color Scheme
// Dark 1 (dk1) – This represents a dark color, usually defined as a system text color
// Light 1 (lt1) – This represents a light color, usually defined as the system window color
// Dark 2 (dk2) – This represents a second dark color for use
// Light 2 (lt2) – This represents a second light color for use
// Accents 1 through 6 (accent1 through accent6) – These are six colors which can be used as
// accent colors in the theme
// Hyperlink (hlink) – The color of hyperlinks
// Followed Hyperlink (folHlink) – The color of a followed hyperlink
func (f *File) getColorScheme(clr *xlsxColor) string {
	var RGB string
	if clr == nil || f.Theme == nil {
		return RGB
	}
	if clrScheme := f.Theme.ThemeElements.ClrScheme; clr.Theme != nil {
		if val, ok := map[int]*string{
			0: clrScheme.Lt1.SrgbClr.Val,
			1: clrScheme.Dk1.SrgbClr.Val,
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
