package excelize

import "strings"

func (f *File) isWindowColor(clr *xlsxColor) bool {
	if clr == nil || clr.Theme == nil || f.Theme == nil {
		return false
	}
	return *clr.Theme == 1
}

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
	var rgb string
	if clr == nil || f.Theme == nil {
		return rgb
	}
	if len(clr.RGB) == 6 {
		return clr.RGB
	}
	if len(clr.RGB) == 8 {
		return strings.TrimPrefix(clr.RGB, "FF")
	}
	switch *clr.Theme {
	case 0:
		if f.Theme.ThemeElements.ClrScheme.Dk1.SysClr != nil {
			rgb = f.Theme.ThemeElements.ClrScheme.Dk1.SysClr.LastClr
		}
	case 1:
		if f.Theme.ThemeElements.ClrScheme.Lt1.SysClr != nil {
			rgb = f.Theme.ThemeElements.ClrScheme.Lt1.SysClr.LastClr
		}
	case 2:
		rgb = *f.Theme.ThemeElements.ClrScheme.Dk2.SrgbClr.Val
	case 3:
		rgb = *f.Theme.ThemeElements.ClrScheme.Lt2.SrgbClr.Val
	case 4:
		rgb = *f.Theme.ThemeElements.ClrScheme.Accent1.SrgbClr.Val
	case 5:
		rgb = *f.Theme.ThemeElements.ClrScheme.Accent2.SrgbClr.Val
	case 6:
		rgb = *f.Theme.ThemeElements.ClrScheme.Accent3.SrgbClr.Val
	case 7:
		rgb = *f.Theme.ThemeElements.ClrScheme.Accent4.SrgbClr.Val
	case 8:
		rgb = *f.Theme.ThemeElements.ClrScheme.Accent5.SrgbClr.Val
	case 9:
		rgb = *f.Theme.ThemeElements.ClrScheme.Accent6.SrgbClr.Val
	case 10:
		rgb = *f.Theme.ThemeElements.ClrScheme.Hlink.SrgbClr.Val
	case 11:
		rgb = *f.Theme.ThemeElements.ClrScheme.FolHlink.SrgbClr.Val
	}
	return strings.TrimPrefix(ThemeColor(rgb, clr.Tint), "FF")
}
