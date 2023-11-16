/*
   Copyright (c) 2020 gingfrederik
   Copyright (c) 2021 Gonzalo Fernandez-Victorio
   Copyright (c) 2021 Basement Crowd Ltd (https://www.basementcrowd.com)
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
	"log"
)

type StructuredDocumentTag struct {
	XMLName xml.Name `xml:"w:sdt"`

	SdtPr      *StructuredDocumentTagProperties    `xml:"w:sdtPr,omitempty"`
	SdtEndPr   *StructuredDocumentTagEndProperties `xml:"w:sdtEndPr,omitempty"`
	SdtContent *StructuredDocumentTagContent       `xml:"w:sdtContent,omitempty"`
}

type StructuredDocumentTagProperties struct {
	XMLName xml.Name `xml:"w:sdtPr"`

	Rpr        *RunProperties      `xml:"w:rPr,omitempty"`
	DocPartObj *DocumentPartObject `xml:"w:docPartObj,omitempty"`
	ID         *TagID              `xml:"w:id,omitempty"`
}

type TagID struct {
	XMLName xml.Name `xml:"w:id"`

	Val string `xml:"w:val,attr,omitempty"`
}

type DocumentPartObject struct {
	XMLName xml.Name `xml:"w:docPartObj"`

	DocumentPartGallery *DocumentPartGallery `xml:"w:docPartGallery,omitempty"`
	DocumentPartUnique  *DocumentPartUnique  `xml:"w:docPartUnique,omitempty"`
}
type DocumentPartGallery struct {
	XMLName xml.Name `xml:"w:docPartGallery"`

	Val string `xml:"w:val,attr,omitempty"`
}
type DocumentPartUnique struct {
	XMLName xml.Name `xml:"w:docPartUnique"`

	Val string `xml:"w:val,attr,omitempty"`
}

type StructuredDocumentTagEndProperties struct {
	XMLName xml.Name       `xml:"w:sdtEndPr"`
	Rpr     *RunProperties `xml:"w:rPr,omitempty"`
}

func (s *StructuredDocumentTagEndProperties) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			log.Println("error", err)
			return err
		}

		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "rPr":
				s.Rpr = &RunProperties{}
				err = d.DecodeElement(s.Rpr, &tt)
				if err != nil {
					return err
				}
			default:
				err = d.Skip()
				if err != nil {
					return err
				}
			}
		}
	}

	return nil
}

type StructuredDocumentTagContent struct {
	XMLName xml.Name `xml:"w:sdtContent"`

	Paragraphs *[]*Paragraph `xml:"w:p,omitempty"`

	Runs *[]*Run `xml:"w:r,omitempty"`

	Tables *[]*Table `xml:"w:tbl,omitempty"`
}

// UnmarshalXML unmarshals
func (sdt *StructuredDocumentTag) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			// log.Println("sdt EOF")
			break
		}
		if err != nil {
			return err
		}

		if t == nil {
			break
		}

		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "sdtPr":
				sdt.SdtPr = &StructuredDocumentTagProperties{}
				err = d.DecodeElement(sdt.SdtPr, &tt)
				if err != nil {
					return err
				}
			case "sdtEndPr":
				sdt.SdtEndPr = &StructuredDocumentTagEndProperties{}
				err = d.DecodeElement(sdt.SdtEndPr, &tt)
				if err != nil {
					return err
				}
			case "sdtContent":
				sdt.SdtContent = &StructuredDocumentTagContent{}
				err = d.DecodeElement(sdt.SdtContent, &tt)
				if err != nil {
					return err
				}
			default:
				err = d.Skip()
				if err != nil {
					return err
				}
			}
		}
	}

	return nil
}

func (sdtp *StructuredDocumentTagProperties) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if t == nil {
			break
		}

		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "rPr":
				sdtp.Rpr = &RunProperties{}
				err = d.DecodeElement(sdtp.Rpr, &tt)
				if err != nil {
					return err
				}
			case "docPartObj":
				sdtp.DocPartObj = &DocumentPartObject{}
				err = d.DecodeElement(sdtp.DocPartObj, &tt)
				if err != nil {
					return err
				}
			case "id":
				sdtp.ID = &TagID{}
				err = d.DecodeElement(sdtp.ID, &tt)
				if err != nil {
					return err
				}
				// log.Println("id", sdtp.ID.Val)
			default:
				err = d.Skip()
				if err != nil {
					return err
				}
			}
		}
	}

	return nil
}

func (sdtp *DocumentPartObject) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if t == nil {
			break
		}

		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "docPartGallery":
				sdtp.DocumentPartGallery = &DocumentPartGallery{}
				err = d.DecodeElement(sdtp.DocumentPartGallery, &tt)
				if err != nil {
					return err
				}
			case "docPartUnique":
				sdtp.DocumentPartUnique = &DocumentPartUnique{}
				err = d.DecodeElement(sdtp.DocumentPartUnique, &tt)
				if err != nil {
					return err
				}
			default:
				err = d.Skip()
				if err != nil {
					return err
				}
			}
		}
	}

	return nil
}

func (sdtc *StructuredDocumentTagContent) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if t == nil {
			break
		}

		// log.Println("sdtc", t)

		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "p":
				p := &Paragraph{}
				err = d.DecodeElement(p, &tt)
				if err != nil {
					return err
				}
				if sdtc.Paragraphs == nil {
					sdtc.Paragraphs = &[]*Paragraph{}
				}
				*sdtc.Paragraphs = append(*sdtc.Paragraphs, p)
			case "r":
				r := &Run{}
				err = d.DecodeElement(r, &tt)
				if err != nil {
					return err
				}
				if sdtc.Runs == nil {
					sdtc.Runs = &[]*Run{}
				}
				*sdtc.Runs = append(*sdtc.Runs, r)
			case "tbl":
				tbl := &Table{}
				err = d.DecodeElement(tbl, &tt)
				if err != nil {
					return err
				}
				if sdtc.Tables == nil {
					sdtc.Tables = &[]*Table{}
				}
				*sdtc.Tables = append(*sdtc.Tables, tbl)
			default:
			}
		}
	}

	return nil
}

func (i *TagID) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, a := range start.Attr {
		switch a.Name.Local {
		case "val":
			i.Val = a.Value
		}
	}

	for {
		_, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			log.Println("error", err)
			return err
		}
	}

	return nil
}

func (dpg *DocumentPartGallery) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, a := range start.Attr {
		switch a.Name.Local {
		case "val":
			dpg.Val = a.Value
		}
	}

	for {
		_, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			log.Println("error", err)
			return err
		}
	}

	return nil
}

func (dpu *DocumentPartUnique) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, a := range start.Attr {
		switch a.Name.Local {
		case "val":
			dpu.Val = a.Value
		}
	}

	for {
		_, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			log.Println("error", err)
			return err
		}
	}

	return nil
}
