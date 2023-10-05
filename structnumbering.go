/*
Copyright (c) 2023 Fumiama Minamoto (源文雨)

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU Affero General Public License as published
by the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU Affero General Public License for more details.

You should have received a copy of the GNU Affero General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/

package docx

import (
	"encoding/xml"
	"fmt"
	"io"
	"log"
)

type CommonAttrVal struct {
	Val string `xml:"w:val,attr"`
}

func (c *CommonAttrVal) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	err = commonSetAttrVal(c, d, start)
	return
}

type Numbering struct {
	XMLName xml.Name `xml:"w:numbering"`
	XMLW    string   `xml:"xmlns:w,attr"`

	// この XML 指定は正しく働いた
	AbstractNums *[]AbstractNum `xml:"w:abstractNum",omitempty`
	Nums         *[]Num         `xml:"w:num",omitempty`
}

type AbstractNum struct {
	XMLName       xml.Name `xml:"w:abstractNum",omitempty`
	AbstractNumID string   `xml:"w:abstractNumId,attr"`
	Lvl           *[]Lvl   `xml:"w:lvl",omitempty"`
	// Lvl            *map[int]*Lvl   `xml:"-",omitempty"`
	NSID           *NSID           `xml:"w:nsid",omitempty`
	MultiLevelType *MultiLevelType `xml:"w:multiLevelType",omitempty`
	Tmpl           *Tmpl           `xml:"w:tmpl",omitempty`
}

type Num struct {
	XMLName        xml.Name `xml:"w:num",omitempty`
	NumID          string   `xml:"w:numId,attr"`
	*AbstractNumID `xml:"w:abstractNumId",omitempty`
	// AbstractNumID *AbstractNumID `xml:"w:abstractNumId",omitempty`
}

type AbstractNumID struct {
	XMLName xml.Name `xml:"w:abstractNumId",omitempty`
	// CommonAttrVal *CommonAttrVal
	*CommonAttrVal
}

type NSID struct {
	XMLName xml.Name `xml:"w:nsid",omitempty`
	// CommonAttrVal *CommonAttrVal
	*CommonAttrVal
}

type MultiLevelType struct {
	XMLName xml.Name `xml:"w:multiLevelType",omitempty`
	// CommonAttrVal *CommonAttrVal
	*CommonAttrVal
}

type Tmpl struct {
	XMLName xml.Name `xml:"w:tmpl",omitempty`
	// CommonAttrVal *CommonAttrVal
	*CommonAttrVal
}

type Lvl struct {
	XMLName   xml.Name `xml:"w:lvl",omitempty`
	ILvl      int      `xml:"w:ilvl,attr"`
	Tplc      string   `xml:"w:tplc,attr"`
	Tentative string   `xml:"w:tentative,attr"`

	Start   *Start                 `xml:"w:start",omitempty`
	NumFmt  *NumFmt                `xml:"w:numFmt",omitempty`
	LvlText *LvlText               `xml:"w:lvlText",omitempty`
	LvlJc   *LvlJc                 `xml:"w:lvlJc",omitempty`
	Ppr     *[]ParagraphProperties `xml:"w:pPr",omitempty"`
}

type Start struct {
	XMLName xml.Name `xml:"w:start",omitempty`
	// CommonAttrVal *CommonAttrVal
	*CommonAttrVal
}

type NumFmt struct {
	XMLName xml.Name `xml:"w:numFmt",omitempty`
	// CommonAttrVal *CommonAttrVal
	*CommonAttrVal
}

type LvlText struct {
	XMLName xml.Name `xml:"w:lvlText",omitempty`
	// CommonAttrVal *CommonAttrVal
	*CommonAttrVal
}

type LvlJc struct {
	XMLName xml.Name `xml:"w:lvlJc",omitempty`
	// CommonAttrVal *CommonAttrVal
	*CommonAttrVal
}

func (n *Numbering) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "abstractNum":
				var an AbstractNum
				err = d.DecodeElement(&an, &tt)
				if err != nil {
					return err
				}
				if n.AbstractNums == nil {
					n.AbstractNums = &[]AbstractNum{}
				}
				*n.AbstractNums = append(*n.AbstractNums, an)
			case "num":
				var num Num
				err = d.DecodeElement(&num, &tt)
				if err != nil {
					return err
				}
				if n.Nums == nil {
					n.Nums = &[]Num{}
				}
				*n.Nums = append(*n.Nums, num)
			default:
				// ignore other attributes
			}
		}
	}

	return nil
}

// AbstractNum UnmarshalXML ...
func (a *AbstractNum) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "abstractNumId":
			a.AbstractNumID = attr.Value
		default:
			// ignore other attributes
		}
	}

	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "lvl":
				l := Lvl{}
				err = d.DecodeElement(&l, &tt)
				if a.Lvl == nil {
					a.Lvl = &[]Lvl{}
				}
				lvlList := paddingListIfNeeded(a.Lvl, l.ILvl).([]Lvl)
				a.Lvl = &lvlList
				(*a.Lvl)[l.ILvl] = l
			case "nsid":
				n := NewNSID()
				err = d.DecodeElement(n, &tt)
				(*a).NSID = n
			case "multiLevelType":
				m := NewMultiLevelType()
				err = d.DecodeElement(m, &tt)
				(*a).MultiLevelType = m
			case "tmpl":
				t := NewTmpl()
				err = d.DecodeElement(t, &tt)
				(*a).Tmpl = t
			}
			if err != nil {
				return err
			}
		}
	}
	// log.Println("a.Lvl: ", *a.Lvl)
	// a.Lvl の各要素の実態を表示
	// ここまでは正確に処理されていた。
	// for _, v := range *a.Lvl {
	// 	log.Println("v.LvlText.Val: ", v.LvlText.Val)
	// }

	return
}

// Lvl UnmarshalXML ...
func (l *Lvl) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "ilvl":
			l.ILvl, err = GetInt(attr.Value)
		case "tplc":
			l.Tplc = attr.Value
		case "tentative":
			l.Tentative = attr.Value
		default:
			// ignore other attributes
		}
	}

	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "pPr":
				var p ParagraphProperties
				err = d.DecodeElement(&p, &tt)
				if err != nil {
					return err
				}
				if l.Ppr == nil {
					l.Ppr = &[]ParagraphProperties{}
				}
				*l.Ppr = append(*l.Ppr, p)
			case "start":
				s := NewStart()
				err = d.DecodeElement(s, &tt)
				if err != nil {
					return err
				}
				l.Start = s
			case "numFmt":
				n := NewNumFmt()
				err = d.DecodeElement(n, &tt)
				if err != nil {
					return err
				}
				l.NumFmt = n
			case "lvlText":
				lt := NewLvlText()
				err = d.DecodeElement(lt, &tt)
				if err != nil {
					return err
				}
				// log.Println("lt: ", lt.Val)
				l.LvlText = lt
			case "lvlJc":
				lj := NewLvlJc()
				err = d.DecodeElement(lj, &tt)
				if err != nil {
					return err
				}
				l.LvlJc = lj
			default:
				// ignore other attributes

			}
		}
	}

	return
}

func (n *Num) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "numId":
			n.NumID = attr.Value
		default:
			// ignore other attributes
		}
	}

	// attr ではなく要素の値を取得する
	for {
		t, err := d.Token()

		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "abstractNumId":
				an := NewAbstractNumID()
				err = d.DecodeElement(an, &tt)
				if err != nil {
					return err
				}
				n.AbstractNumID = an
			default:
				// ignore other attributes
			}
		}
	}
	return
}

func commonSetAttrVal(a *CommonAttrVal, d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "val":
			if a.Val != "" {
				log.Printf("commonSetAttrVal: %s is already set", attr.Name.Local)
			}
			a.Val = attr.Value
		default:
			// ignore other attributes
		}
	}

	_, err = d.Token()
	return
}

func NewAbstractNumID() *AbstractNumID {
	return &AbstractNumID{CommonAttrVal: &CommonAttrVal{}}
}

func NewNSID() *NSID {
	return &NSID{CommonAttrVal: &CommonAttrVal{}}
}

func NewMultiLevelType() *MultiLevelType {
	return &MultiLevelType{CommonAttrVal: &CommonAttrVal{}}
}

func NewTmpl() *Tmpl {
	return &Tmpl{CommonAttrVal: &CommonAttrVal{}}
}

func NewStart() *Start {
	return &Start{CommonAttrVal: &CommonAttrVal{}}
}

func NewNumFmt() *NumFmt {
	return &NumFmt{CommonAttrVal: &CommonAttrVal{}}
}

func NewLvlText() *LvlText {
	return &LvlText{CommonAttrVal: &CommonAttrVal{}}
}

func NewLvlJc() *LvlJc {
	return &LvlJc{CommonAttrVal: &CommonAttrVal{}}
}

func paddingListIfNeeded(list interface{}, index int) interface{} {
	switch v := list.(type) {
	case *[]Lvl:
		if index < len(*v) {
			return v
		}
		return append(*v, make([]Lvl, index-len(*v)+1)...)
	default:
		panic(fmt.Sprintf("paddingListIfNeeded: unsupported type %T", v))
	}
}
