package xls

import (
	"fmt"
	"math"
	"strconv"
	"strings"
	"time"
)

// Ranger interface for multi-row ranges
type Ranger interface {
	FirstRow() uint16
	LastRow() uint16
}

// CellRange represents a range of cells in multi rows
type CellRange struct {
	FirstRowB uint16
	LastRowB  uint16
	FristColB uint16
	LastColB  uint16
}

func (c *CellRange) FirstRow() uint16 {
	return c.FirstRowB
}

func (c *CellRange) LastRow() uint16 {
	return c.LastRowB
}

func (c *CellRange) FirstCol() uint16 {
	return c.FristColB
}

func (c *CellRange) LastCol() uint16 {
	return c.LastColB
}

// HyperLink represents a hyperlink cell
type HyperLink struct {
	CellRange
	Description      string
	TextMark         string
	TargetFrame      string
	Url              string
	ShortedFilePath  string
	ExtendedFilePath string
	IsUrl            bool
}

func (h *HyperLink) String(wb *WorkBook) []string {
	res := make([]string, h.LastColB-h.FristColB+1)
	var str string
	if h.IsUrl {
		str = fmt.Sprintf("%s(%s)", h.Description, h.Url)
	} else {
		str = h.ExtendedFilePath
	}

	for i := uint16(0); i < h.LastColB-h.FristColB+1; i++ {
		res[i] = str
	}
	return res
}

// content type
type contentHandler interface {
	String(*WorkBook) []string
	FirstCol() uint16
	LastCol() uint16
}

type Col struct {
	RowB      uint16
	FirstColB uint16
}

type Coler interface {
	Row() uint16
}

func (c *Col) Row() uint16 {
	return c.RowB
}

func (c *Col) FirstCol() uint16 {
	return c.FirstColB
}

func (c *Col) LastCol() uint16 {
	return c.FirstColB
}

func (c *Col) String(wb *WorkBook) []string {
	return []string{"default"}
}

type XfRk struct {
	Index uint16
	Rk    RK
}

func (xf *XfRk) String(wb *WorkBook) string {
	idx := int(xf.Index)
	if len(wb.Xfs) > idx {
		if wb.Xfs[idx] == nil {
			return xf.Rk.String()
		}
		fNo := wb.Xfs[idx].formatNo()
		if fNo >= 164 { // user defined format
			if formatter := wb.Formats[fNo]; formatter != nil {
				formatterLower := strings.ToLower(formatter.str)
				if formatterLower == "general" ||
					strings.Contains(formatter.str, "#") ||
					strings.Contains(formatter.str, ".00") ||
					strings.Contains(formatterLower, "m/y") ||
					strings.Contains(formatterLower, "d/y") ||
					strings.Contains(formatterLower, "m.y") ||
					strings.Contains(formatterLower, "d.y") ||
					strings.Contains(formatterLower, "h:") ||
					strings.Contains(formatterLower, "д.г") {
					//If format contains # or .00 then this is a number
					return xf.Rk.String()
				} else {
					i, f, isFloat := xf.Rk.number()
					if !isFloat {
						f = float64(i)
					}
					t := timeFromExcelTime(f, wb.dateMode == 1)
					return formatExcelTime(t, formatter.str)
				}
			}
			// see http://www.openoffice.org/sc/excelfileformat.pdf Page #174
		} else if 14 <= fNo && fNo <= 17 || fNo == 22 || 27 <= fNo && fNo <= 36 || 50 <= fNo && fNo <= 58 { // jp. date format
			i, f, isFloat := xf.Rk.number()
			if !isFloat {
				f = float64(i)
			}
			t := timeFromExcelTime(f, wb.dateMode == 1)
			return t.Format(time.RFC3339) //TODO it should be international
		}
	}
	return xf.Rk.String()
}

type RK uint32

func (rk RK) number() (intNum int64, floatNum float64, isFloat bool) {
	multiplied := rk & 1
	isInt := rk & 2
	val := int32(rk) >> 2
	if isInt == 0 {
		isFloat = true
		floatNum = math.Float64frombits(uint64(val) << 34)
		if multiplied != 0 {
			floatNum = floatNum / 100
		}
		return
	}
	if multiplied != 0 {
		isFloat = true
		floatNum = float64(val) / 100
		return
	}
	return int64(val), 0, false
}

func (rk RK) String() string {
	i, f, isFloat := rk.number()
	if isFloat {
		return strconv.FormatFloat(f, 'f', -1, 64)
	}
	return strconv.FormatInt(i, 10)
}

var ErrIsInt = fmt.Errorf("is int")

func (rk RK) Float() (float64, error) {
	_, f, isFloat := rk.number()
	if !isFloat {
		return 0, ErrIsInt
	}
	return f, nil
}

type MulrkCol struct {
	Col
	Xfrks    []XfRk
	LastColB uint16
}

func (c *MulrkCol) LastCol() uint16 {
	return c.LastColB
}

func (c *MulrkCol) String(wb *WorkBook) []string {
	var res = make([]string, len(c.Xfrks))
	for i := 0; i < len(c.Xfrks); i++ {
		xfrk := c.Xfrks[i]
		res[i] = xfrk.String(wb)
	}
	return res
}

type MulBlankCol struct {
	Col
	Xfs      []uint16
	LastColB uint16
}

func (c *MulBlankCol) LastCol() uint16 {
	return c.LastColB
}

func (c *MulBlankCol) String(wb *WorkBook) []string {
	return make([]string, len(c.Xfs))
}

type NumberCol struct {
	Col
	Index uint16
	Float float64
}

func (c *NumberCol) String(wb *WorkBook) []string {
	if int(c.Index) < len(wb.Xfs) && wb.Xfs[c.Index] != nil {
		if fNo := wb.Xfs[c.Index].formatNo(); fNo != 0 {
			t := timeFromExcelTime(c.Float, wb.dateMode == 1)
			if fmtObj := wb.Formats[fNo]; fmtObj != nil {
				return []string{formatExcelTime(t, fmtObj.str)}
			}
		}
	}
	return []string{strconv.FormatFloat(c.Float, 'f', -1, 64)}
}

// formatExcelTime formats time according to Excel format string
// Supports common Excel date/time format patterns
func formatExcelTime(t time.Time, format string) string {
	// Convert Excel format to Go format
	goFormat := convertExcelFormatToGo(format)
	return t.Format(goFormat)
}

// convertExcelFormatToGo converts Excel date/time format to Go format
func convertExcelFormatToGo(excelFormat string) string {
	// Map common Excel format patterns to Go format
	result := excelFormat

	// Handle minute patterns first (before month) - only when preceded by :
	// Use a placeholder to avoid collision with month
	if strings.Contains(result, ":") {
		result = strings.ReplaceAll(result, ":mm", ":###MIN###")
		result = strings.ReplaceAll(result, ":m", ":###MIN1###")
	}

	// Year patterns (case insensitive handled by lower conversion)
	result = strings.ReplaceAll(result, "yyyy", "2006")
	result = strings.ReplaceAll(result, "YYYY", "2006")
	result = strings.ReplaceAll(result, "yy", "06")
	result = strings.ReplaceAll(result, "YY", "06")

	// Month patterns
	result = strings.ReplaceAll(result, "mmmm", "January")
	result = strings.ReplaceAll(result, "MMMM", "January")
	result = strings.ReplaceAll(result, "mmm", "Jan")
	result = strings.ReplaceAll(result, "MMM", "Jan")
	result = strings.ReplaceAll(result, "mm", "01")
	result = strings.ReplaceAll(result, "MM", "01")
	result = strings.ReplaceAll(result, "m", "1")
	result = strings.ReplaceAll(result, "M", "1")

	// Day patterns
	result = strings.ReplaceAll(result, "dd", "02")
	result = strings.ReplaceAll(result, "DD", "02")
	result = strings.ReplaceAll(result, "d", "2")
	result = strings.ReplaceAll(result, "D", "2")

	// Hour patterns
	result = strings.ReplaceAll(result, "hh", "15")
	result = strings.ReplaceAll(result, "HH", "15")
	result = strings.ReplaceAll(result, "h", "3")
	result = strings.ReplaceAll(result, "H", "3")

	// Restore minute patterns
	result = strings.ReplaceAll(result, ":###MIN###", ":04")
	result = strings.ReplaceAll(result, ":###MIN1###", ":4")

	// Second patterns
	result = strings.ReplaceAll(result, "ss", "05")
	result = strings.ReplaceAll(result, "SS", "05")
	result = strings.ReplaceAll(result, "s", "5")
	result = strings.ReplaceAll(result, "S", "5")

	// AM/PM
	result = strings.ReplaceAll(result, "AM/PM", "PM")
	result = strings.ReplaceAll(result, "am/pm", "pm")

	return result
}

type FormulaStringCol struct {
	Col
	RenderedValue string
}

func (c *FormulaStringCol) String(wb *WorkBook) []string {
	return []string{c.RenderedValue}
}

//str, err = wb.get_string(buf_item, size)
//wb.sst[offset_pre] = wb.sst[offset_pre] + str

type FormulaCol struct {
	Header struct {
		Col
		IndexXf uint16
		Result  [8]byte
		Flags   uint16
		_       uint32
	}
	Bts []byte
}

func (c *FormulaCol) String(wb *WorkBook) []string {
	return []string{"FormulaCol"}
}

type RkCol struct {
	Col
	Xfrk XfRk
}

func (c *RkCol) String(wb *WorkBook) []string {
	return []string{c.Xfrk.String(wb)}
}

type LabelsstCol struct {
	Col
	Xf  uint16
	Sst uint32
}

func (c *LabelsstCol) String(wb *WorkBook) []string {
	if int(c.Sst) < len(wb.sst) {
		return []string{wb.sst[int(c.Sst)]}
	}
	return []string{""}
}

type labelCol struct {
	BlankCol
	Str string
}

func (c *labelCol) String(wb *WorkBook) []string {
	return []string{c.Str}
}

type BlankCol struct {
	Col
	Xf uint16
}

func (c *BlankCol) String(wb *WorkBook) []string {
	return []string{""}
}
