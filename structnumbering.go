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
	"io"
)

type Numbering struct {
	XMLName      xml.Name       `xml:"numbering"`
	AbstructNums *[]AbstractNum `xml:"abstructNmum",omitempty`
}

type AbstractNum struct {
	XMLName        xml.Name `xml:"abstructNmum",omitempty`
	AbstractNumID  string   `xml:"w:abstractNumId,attr"`
	Lvl            *[]Lvl   `xml:"lvl"`
	NSID           *NSID
	MultiLevelType *MultiLevelType `xml:"multiLevelType",omitempty`
	Tmpl           *Tmpl           `xml:"tmpl",omitempty`
}

type NSID struct {
	XMLName xml.Name `xml:"nsid",omitempty`
	Val     string   `xml:"w:val,attr"`
}

type MultiLevelType struct {
	XMLName xml.Name `xml:"multiLevelType",omitempty`
	Val     string   `xml:"w:val,attr"`
}

type Tmpl struct {
	XMLName xml.Name `xml:"tmpl",omitempty`
	Val     string   `xml:"w:val,attr"`
}

type Lvl struct {
	XMLName xml.Name               `xml:"lvl",omitempty`
	ILvl    string                 `xml:"w:ilvl,attr"`
	Tplc    string                 `xml:"w:tplc,attr"`
	Start   *Start                 `xml:"start",omitempty`
	NumFmt  *NumFmt                `xml:"numFmt",omitempty`
	LvlText *LvlText               `xml:"lvlText",omitempty`
	LvlJc   *LvlJc                 `xml:"lvlJc",omitempty`
	Ppr     *[]ParagraphProperties `xml:"pPr",omitempty`
}

type Start struct {
	XMLName xml.Name `xml:"start",omitempty`
	Val     string   `xml:"w:val,attr"`
}

type NumFmt struct {
	XMLName xml.Name `xml:"numFmt",omitempty`
	Val     string   `xml:"w:val,attr"`
}

type LvlText struct {
	XMLName xml.Name `xml:"lvlText",omitempty`
	Val     string   `xml:"w:val,attr"`
}

type LvlJc struct {
	XMLName xml.Name `xml:"lvlJc",omitempty`
	Val     string   `xml:"w:val,attr"`
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
			if tt.Name.Local == "abstractNum" {
				var an AbstractNum
				err = d.DecodeElement(&an, &tt)
				if err != nil {
					return err
				}
				if n.AbstructNums == nil {
					n.AbstructNums = &[]AbstractNum{}
				}
				*n.AbstructNums = append(*n.AbstructNums, an)
			}
		}
	}

	return nil
}

// AbstructNum UnmarshalXML ...
func (a *AbstractNum) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "abstractNumId":
			a.AbstractNumID = attr.Value
		case "nsid":
			a.NSID.Val = attr.Value
		case "multiLevelType":
			a.MultiLevelType.Val = attr.Value
		case "tmpl":
			a.Tmpl.Val = attr.Value
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
			if tt.Name.Local == "lvl" {
				var l Lvl
				err = d.DecodeElement(&l, &tt)
				if err != nil {
					return err
				}
				if a.Lvl == nil {
					a.Lvl = &[]Lvl{}
				}
				*a.Lvl = append(*a.Lvl, l)
			}
		}
	}

	return
}

// Lvl UnmarshalXML ...
func (l *Lvl) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "ilvl":
			l.ILvl = attr.Value
		case "tplc":
			l.Tplc = attr.Value
		case "start":
			l.Start.Val = attr.Value
		case "numFmt":
			l.NumFmt.Val = attr.Value
		case "lvlText":
			l.LvlText.Val = attr.Value
		case "lvlJc":
			l.LvlJc.Val = attr.Value
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
			if tt.Name.Local == "pPr" {
				var p ParagraphProperties
				err = d.DecodeElement(&p, &tt)
				if err != nil {
					return err
				}
				if l.Ppr == nil {
					l.Ppr = &[]ParagraphProperties{}
				}
				*l.Ppr = append(*l.Ppr, p)
			}
		}
	}

	return
}
