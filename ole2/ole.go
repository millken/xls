package ole2

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"io"
	"unicode/utf16"
)

const ENDOFCHAIN = uint32(0xFFFFFFFE) //-2
const FREESECT = uint32(0xFFFFFFFF)   // -1

// Header represents the OLE2 compound file header
type Header struct {
	Id        [2]uint32
	Clid      [4]uint32
	Verminor  uint16
	Verdll    uint16
	Byteorder uint16
	Lsectorb  uint16
	Lssectorb uint16
	_         uint16
	_         uint64

	Cfat     uint32 //Total number of sectors used for the sector allocation table
	Dirstart uint32 //SecID of first sector of the directory stream

	_ uint32

	Sectorcutoff uint32 //Minimum size of a standard stream
	Sfatstart    uint32 //SecID of first sector of the short-sector allocation table
	Csfat        uint32 //Total number of sectors used for the short-sector allocation table
	Difstart     uint32 //SecID of first sector of the master sector allocation table
	Cdif         uint32 //Total number of sectors used for the master sector allocation table
	Msat         [109]uint32
}

// File types in OLE2 directory
const (
	EMPTY       = iota
	USERSTORAGE = iota
	USERSTREAM  = iota
	LOCKBYTES   = iota
	PROPERTY    = iota
	ROOT        = iota
)

// File represents an entry in the OLE2 directory
type File struct {
	NameBts   [32]uint16
	Bsize     uint16
	Type      byte
	Flag      byte
	Left      uint32
	Right     uint32
	Child     uint32
	Guid      [8]uint16
	Userflags uint32
	Time      [2]uint64
	Sstart    uint32
	Size      uint32
	Proptype  uint32
}

func (d *File) Name() string {
	runes := utf16.Decode(d.NameBts[:d.Bsize/2-1])
	return string(runes)
}

// Sector represents a sector in OLE2 file
type Sector []byte

func (s *Sector) Uint32(bit uint32) uint32 {
	return binary.LittleEndian.Uint32((*s)[bit : bit+4])
}

func (s *Sector) NextSid(size uint32) uint32 {
	return s.Uint32(size - 4)
}

func (s *Sector) MsatValues(size uint32) []uint32 {
	return s.values(size, int(size/4-1))
}

func (s *Sector) AllValues(size uint32) []uint32 {
	return s.values(size, int(size/4))
}

func (s *Sector) values(size uint32, length int) []uint32 {
	var res = make([]uint32, length)
	buf := bytes.NewBuffer((*s))
	binary.Read(buf, binary.LittleEndian, res)
	return res
}

// Ole represents an OLE2 compound file
type Ole struct {
	header   *Header
	Lsector  uint32
	Lssector uint32
	SecID    []uint32
	SSecID   []uint32
	Files    []File
	reader   io.ReadSeeker
}

// Open opens an OLE2 compound file
func Open(reader io.ReadSeeker, charset string) (ole *Ole, err error) {
	var header *Header
	var hbts = make([]byte, 512)
	if _, err = reader.Read(hbts); err != nil {
		return nil, err
	}
	if header, err = parseHeader(hbts); err == nil {
		ole = new(Ole)
		ole.reader = reader
		ole.header = header
		ole.Lsector = 512
		ole.Lssector = 64
		err = ole.readMSAT()
		return ole, err
	}

	return nil, err
}

func parseHeader(bts []byte) (*Header, error) {
	buf := bytes.NewBuffer(bts)
	header := new(Header)
	binary.Read(buf, binary.LittleEndian, header)
	if header.Id[0] != 0xE011CFD0 || header.Id[1] != 0xE11AB1A1 || header.Byteorder != 0xFFFE {
		return nil, fmt.Errorf("not an excel file")
	}
	return header, nil
}

func (o *Ole) ListDir() (dir []*File, err error) {
	sector := o.stream_read(o.header.Dirstart, 0)
	dir = make([]*File, 0)
	for {
		d := new(File)
		err = binary.Read(sector, binary.LittleEndian, d)
		if err == nil && d.Type != EMPTY {
			dir = append(dir, d)
		} else {
			break
		}
	}
	if err == io.EOF && dir != nil {
		return dir, nil
	}
	return
}

func (o *Ole) OpenFile(file *File, root *File) io.ReadSeeker {
	if file.Size < o.header.Sectorcutoff {
		return o.short_stream_read(file.Sstart, file.Size, root.Sstart)
	} else {
		return o.stream_read(file.Sstart, file.Size)
	}
}

// Read MSAT
func (o *Ole) readMSAT() error {
	count := uint32(109)
	if o.header.Cfat < 109 {
		count = o.header.Cfat
	}

	for i := uint32(0); i < count; i++ {
		if sector, err := o.sector_read(o.header.Msat[i]); err == nil {
			sids := sector.AllValues(o.Lsector)
			o.SecID = append(o.SecID, sids...)
		} else {
			return err
		}
	}

	for sid := o.header.Difstart; sid != ENDOFCHAIN; {
		if sector, err := o.sector_read(sid); err == nil {
			sids := sector.MsatValues(o.Lsector)

			for _, sid := range sids {
				if sector, err := o.sector_read(sid); err == nil {
					sids := sector.AllValues(o.Lsector)
					o.SecID = append(o.SecID, sids...)
				} else {
					return err
				}
			}

			sid = sector.NextSid(o.Lsector)
		} else {
			return err
		}
	}

	for i := uint32(0); i < o.header.Csfat; i++ {
		sid := o.header.Sfatstart

		if sid != ENDOFCHAIN {
			if sector, err := o.sector_read(sid); err == nil {
				sids := sector.MsatValues(o.Lsector)
				o.SSecID = append(o.SSecID, sids...)
				sid = sector.NextSid(o.Lsector)
			} else {
				return err
			}
		}
	}
	return nil
}

// StreamReader reads streams from OLE2 file
type StreamReader struct {
	sat              []uint32
	start            uint32
	reader           io.ReadSeeker
	offset_of_sector uint32
	offset_in_sector uint32
	size_sector      uint32
	size             int64
	offset           int64
	sector_pos       func(uint32, uint32) uint32
}

func (r *StreamReader) Read(p []byte) (n int, err error) {
	if r.offset_of_sector == ENDOFCHAIN {
		return 0, io.EOF
	}
	if len(p) == 0 {
		return 0, nil
	}
	pos := r.sector_pos(r.offset_of_sector, r.size_sector) + r.offset_in_sector
	if _, err = r.reader.Seek(int64(pos), io.SeekStart); err != nil {
		return 0, err
	}
	readed := uint32(0)
	remaining := uint32(len(p))

	for remaining > 0 {
		available := r.size_sector - r.offset_in_sector
		if remaining <= available {
			n, err = r.reader.Read(p[readed:])
			if n > 0 {
				r.offset_in_sector += uint32(n)
				r.offset += int64(n)
			}
			if err != nil {
				return int(readed) + n, err
			}
			return int(readed) + n, nil
		}

		// Read available bytes in current sector
		var bytesRead int
		bytesRead, err = r.reader.Read(p[readed : readed+available])
		if bytesRead > 0 {
			readed += uint32(bytesRead)
			remaining -= uint32(bytesRead)
			r.offset += int64(bytesRead)
		}
		if err != nil {
			return int(readed), err
		}

		// Move to next sector
		r.offset_in_sector = 0
		if r.offset_of_sector >= uint32(len(r.sat)) {
			return int(readed), io.EOF
		}
		r.offset_of_sector = r.sat[r.offset_of_sector]
		if r.offset_of_sector == ENDOFCHAIN {
			return int(readed), io.EOF
		}
		pos = r.sector_pos(r.offset_of_sector, r.size_sector)
		if _, err = r.reader.Seek(int64(pos), io.SeekStart); err != nil {
			return int(readed), err
		}
	}

	return int(readed), nil
}

func (r *StreamReader) Seek(offset int64, whence int) (int64, error) {
	var newOffset int64

	switch whence {
	case io.SeekStart:
		newOffset = offset
	case io.SeekCurrent:
		newOffset = r.offset + offset
	case io.SeekEnd:
		newOffset = r.size + offset
	default:
		return r.offset, io.ErrUnexpectedEOF
	}

	if newOffset < 0 {
		return r.offset, io.ErrUnexpectedEOF
	}
	if newOffset > r.size {
		return r.offset, io.EOF
	}

	// Reset to start and seek forward
	r.offset_of_sector = r.start
	r.offset_in_sector = 0
	r.offset = 0

	remaining := newOffset
	for remaining >= int64(r.size_sector) {
		if r.offset_of_sector >= uint32(len(r.sat)) {
			return r.offset, io.EOF
		}
		r.offset_of_sector = r.sat[r.offset_of_sector]
		if r.offset_of_sector == ENDOFCHAIN {
			return r.offset, io.EOF
		}
		remaining -= int64(r.size_sector)
	}

	r.offset_in_sector = uint32(remaining)
	r.offset = newOffset

	return newOffset, nil
}

func (o *Ole) stream_read(sid uint32, size uint32) *StreamReader {
	return &StreamReader{o.SecID, sid, o.reader, sid, 0, o.Lsector, int64(size), 0, sector_pos}
}

func (o *Ole) short_stream_read(sid uint32, size uint32, startSecId uint32) *StreamReader {
	ssatReader := &StreamReader{o.SecID, startSecId, o.reader, sid, 0, o.Lsector, int64(uint32(len(o.SSecID)) * o.Lssector), 0, sector_pos}
	return &StreamReader{o.SSecID, sid, ssatReader, sid, 0, o.Lssector, int64(size), 0, short_sector_pos}
}

func (o *Ole) sector_read(sid uint32) (Sector, error) {
	return o.sector_read_internal(sid, o.Lsector)
}

func (o *Ole) short_sector_read(sid uint32) (Sector, error) {
	return o.sector_read_internal(sid, o.Lssector)
}

func (o *Ole) sector_read_internal(sid, size uint32) (Sector, error) {
	pos := sector_pos(sid, size)
	if _, err := o.reader.Seek(int64(pos), 0); err != nil {
		return nil, err
	}
	bts := make([]byte, size)
	if _, err := o.reader.Read(bts); err != nil {
		return nil, err
	}
	return Sector(bts), nil
}

func sector_pos(sid uint32, size uint32) uint32 {
	return 512 + sid*size
}

func short_sector_pos(sid uint32, size uint32) uint32 {
	return sid * size
}
