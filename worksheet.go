package xls

import (
	"bytes"
	"encoding/binary"
	"io"
	"unicode/utf16"
)

// Row info structure
type rowInfo struct {
	Index    uint16
	Fcell    uint16
	Lcell    uint16
	Height   uint16
	Notused  uint16
	Notused2 uint16
	Flags    uint32
}

// Row represents the data of one row
type Row struct {
	wb   *WorkBook
	info *rowInfo
	cols map[uint16]contentHandler
}

// Col Get the Nth Col from the Row, if has not, return nil.
// Suggest use Has function to test it.
func (r *Row) Col(i int) string {
	serial := uint16(i)
	if ch, ok := r.cols[serial]; ok {
		strs := ch.String(r.wb)
		return strs[0]
	} else {
		for _, v := range r.cols {
			if v.FirstCol() <= serial && v.LastCol() >= serial {
				strs := v.String(r.wb)
				return strs[serial-v.FirstCol()]
			}
		}
	}
	return ""
}

// ColExact Get the Nth Col from the Row, if has not, return nil.
// For merged cells value is returned for first cell only
func (r *Row) ColExact(i int) string {
	serial := uint16(i)
	if ch, ok := r.cols[serial]; ok {
		strs := ch.String(r.wb)
		return strs[0]
	}
	return ""
}

// LastCol Get the number of Last Col of the Row.
func (r *Row) LastCol() int {
	return int(r.info.Lcell)
}

// FirstCol Get the number of First Col of the Row.
func (r *Row) FirstCol() int {
	return int(r.info.Fcell)
}

type TWorkSheetVisibility byte

const (
	WorkSheetVisible    TWorkSheetVisibility = 0
	WorkSheetHidden     TWorkSheetVisibility = 1
	WorkSheetVeryHidden TWorkSheetVisibility = 2
)

// WorkSheet in one WorkBook
type WorkSheet struct {
	bs         *boundsheet
	wb         *WorkBook
	Name       string
	Selected   bool
	Visibility TWorkSheetVisibility
	rows       map[uint16]*Row
	//NOTICE: this is the max row number of the sheet, so it should be count -1
	MaxRow      uint16
	parsed      bool
	rightToLeft bool
}

func (w *WorkSheet) Row(i int) *Row {
	row := w.rows[uint16(i)]
	if row != nil {
		row.wb = w.wb
	}
	return row
}

func (w *WorkSheet) parse(buf io.ReadSeeker) {
	w.rows = make(map[uint16]*Row)
	b := new(bof)
	var bof_pre *bof
	var col_pre interface{}
	for {
		if err := binary.Read(buf, binary.LittleEndian, b); err == nil {
			bof_pre, col_pre = w.parseBof(buf, b, bof_pre, col_pre)
			if b.Id == 0xa {
				break
			}
		} else {
			break
		}
	}
	w.parsed = true
}

func (w *WorkSheet) parseBof(buf io.ReadSeeker, b *bof, pre *bof, col_pre interface{}) (*bof, interface{}) {
	var col interface{}
	var bts = make([]byte, b.Size)
	binary.Read(buf, binary.LittleEndian, bts)
	buf = bytes.NewReader(bts)
	switch b.Id {
	// case 0x0E5: //MERGEDCELLS
	// ws.mergedCells(buf)
	case 0x23E: // WINDOW2
		var sheetOptions, firstVisibleRow, firstVisibleColumn uint16
		binary.Read(buf, binary.LittleEndian, &sheetOptions)
		binary.Read(buf, binary.LittleEndian, &firstVisibleRow)    // not valuable
		binary.Read(buf, binary.LittleEndian, &firstVisibleColumn) // not valuable
		//buf.Seek(int64(b.Size)-2*3, 1)
		w.rightToLeft = (sheetOptions & 0x40) != 0
		w.Selected = (sheetOptions & 0x400) != 0
	case 0x208: //ROW
		r := new(rowInfo)
		binary.Read(buf, binary.LittleEndian, r)
		w.addRow(r)
	case 0x0BD: //MULRK
		mc := new(MulrkCol)
		size := (b.Size - 6) / 6
		binary.Read(buf, binary.LittleEndian, &mc.Col)
		mc.Xfrks = make([]XfRk, size)
		for i := uint16(0); i < size; i++ {
			binary.Read(buf, binary.LittleEndian, &mc.Xfrks[i])
		}
		binary.Read(buf, binary.LittleEndian, &mc.LastColB)
		col = mc
	case 0x0BE: //MULBLANK
		mc := new(MulBlankCol)
		size := (b.Size - 6) / 2
		binary.Read(buf, binary.LittleEndian, &mc.Col)
		mc.Xfs = make([]uint16, size)
		for i := uint16(0); i < size; i++ {
			binary.Read(buf, binary.LittleEndian, &mc.Xfs[i])
		}
		binary.Read(buf, binary.LittleEndian, &mc.LastColB)
		col = mc
	case 0x203: //NUMBER
		col = new(NumberCol)
		binary.Read(buf, binary.LittleEndian, col)
	case 0x06: //FORMULA
		c := new(FormulaCol)
		binary.Read(buf, binary.LittleEndian, &c.Header)
		c.Bts = make([]byte, b.Size-20)
		binary.Read(buf, binary.LittleEndian, &c.Bts)
		col = c
	case 0x207: //STRING = FORMULA-VALUE is expected right after FORMULA
		if ch, ok := col_pre.(*FormulaCol); ok {
			c := new(FormulaStringCol)
			c.Col = ch.Header.Col
			var cStringLen uint16
			binary.Read(buf, binary.LittleEndian, &cStringLen)
			str, err := w.wb.get_string(buf, cStringLen)
			if nil == err {
				c.RenderedValue = str
			}
			col = c
		}
	case 0x27e: //RK
		col = new(RkCol)
		binary.Read(buf, binary.LittleEndian, col)
	case 0xFD: //LABELSST
		col = new(LabelsstCol)
		binary.Read(buf, binary.LittleEndian, col)
	case 0x204:
		c := new(labelCol)
		binary.Read(buf, binary.LittleEndian, &c.BlankCol)
		var count uint16
		binary.Read(buf, binary.LittleEndian, &count)
		c.Str, _ = w.wb.get_string(buf, count)
		col = c
	case 0x201: //BLANK
		col = new(BlankCol)
		binary.Read(buf, binary.LittleEndian, col)
	case 0x1b8: //HYPERLINK
		var hy HyperLink
		binary.Read(buf, binary.LittleEndian, &hy.CellRange)
		buf.Seek(20, 1)
		var flag uint32
		binary.Read(buf, binary.LittleEndian, &flag)
		var count uint32

		if flag&0x14 != 0 {
			binary.Read(buf, binary.LittleEndian, &count)
			hy.Description = b.utf16String(buf, count)
		}
		if flag&0x80 != 0 {
			binary.Read(buf, binary.LittleEndian, &count)
			hy.TargetFrame = b.utf16String(buf, count)
		}
		if flag&0x1 != 0 {
			var guid [2]uint64
			binary.Read(buf, binary.BigEndian, &guid)
			if guid[0] == 0xE0C9EA79F9BACE11 && guid[1] == 0x8C8200AA004BA90B { //URL
				hy.IsUrl = true
				binary.Read(buf, binary.LittleEndian, &count)
				hy.Url = b.utf16String(buf, count/2)
			} else if guid[0] == 0x303000000000000 && guid[1] == 0xC000000000000046 { //URL{
				var upCount uint16
				binary.Read(buf, binary.LittleEndian, &upCount)
				binary.Read(buf, binary.LittleEndian, &count)
				bts := make([]byte, count)
				binary.Read(buf, binary.LittleEndian, &bts)
				hy.ShortedFilePath = string(bts)
				buf.Seek(24, 1)
				binary.Read(buf, binary.LittleEndian, &count)
				if count > 0 {
					binary.Read(buf, binary.LittleEndian, &count)
					buf.Seek(2, 1)
					hy.ExtendedFilePath = b.utf16String(buf, count/2+1)
				}
			}
		}
		if flag&0x8 != 0 {
			binary.Read(buf, binary.LittleEndian, &count)
			var bts = make([]uint16, count)
			binary.Read(buf, binary.LittleEndian, &bts)
			runes := utf16.Decode(bts[:len(bts)-1])
			hy.TextMark = string(runes)
		}

		w.addRange(&hy.CellRange, &hy)
	case 0x809:
		buf.Seek(int64(b.Size), 1)
	case 0xa:
	default:
		// log.Printf("Unknow %X,%d\n", b.Id, b.Size)
		buf.Seek(int64(b.Size), 1)
	}
	if col != nil {
		w.add(col)
	}
	return b, col
}

func (w *WorkSheet) add(content interface{}) {
	if ch, ok := content.(contentHandler); ok {
		if col, ok := content.(Coler); ok {
			w.addCell(col, ch)
		}
	}

}

func (w *WorkSheet) addCell(col Coler, ch contentHandler) {
	w.addContent(col.Row(), ch)
}

func (w *WorkSheet) addRange(rang Ranger, ch contentHandler) {

	for i := rang.FirstRow(); i <= rang.LastRow(); i++ {
		w.addContent(i, ch)
	}
}

func (w *WorkSheet) addContent(row_num uint16, ch contentHandler) {
	var row *Row
	var ok bool
	if row, ok = w.rows[row_num]; !ok {
		info := new(rowInfo)
		info.Index = row_num
		row = w.addRow(info)
	}
	if row.info.Lcell < ch.LastCol() {
		row.info.Lcell = ch.LastCol()
	}
	row.cols[ch.FirstCol()] = ch
}

func (w *WorkSheet) addRow(info *rowInfo) (row *Row) {
	if info.Index > w.MaxRow {
		w.MaxRow = info.Index
	}
	var ok bool
	if row, ok = w.rows[info.Index]; ok {
		row.info = info
	} else {
		row = &Row{info: info, cols: make(map[uint16]contentHandler)}
		w.rows[info.Index] = row
	}
	return
}
