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
	"strings"
)

// Hyperlink element contains links
type Hyperlink struct {
	XMLName xml.Name `xml:"w:hyperlink,omitempty"`
	ID      string   `xml:"r:id,attr"`
	Anchor  string   `xml:"w:anchor,attr,omitempty"` // anchor is used for internal links
	Runs    *[]*Run  `xml:"w:r,omitempty"`
}

// UnmarshalXML ...
func (r *Hyperlink) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if tt, ok := t.(xml.StartElement); ok {
			if tt.Name.Local == "r" {
				run := Run{}
				err = d.DecodeElement(&run, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				if r.Runs == nil {
					r.Runs = &[]*Run{&run}
				} else {
					*r.Runs = append(*r.Runs, &run)
				}
				continue
			}
			err = d.Skip() // skip unsupported tags
			if err != nil {
				return err
			}
		}
	}
	return nil
}

type BookmarkStart struct {
	XMLName xml.Name `xml:"w:bookmarkStart,omitempty"`
	ID      string   `xml:"w:id,attr"`
	Name    string   `xml:"w:name,attr,omitempty"`
	Run     Run      `xml:"w:r,omitempty"`
}

func (b *BookmarkStart) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// log.Println("BookmarkStart", start.Attr)
	for _, attr := range start.Attr {
		if attr.Name.Local == "id" {
			b.ID = attr.Value
		}
		if attr.Name.Local == "name" {
			b.Name = attr.Value
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
			if tt.Name.Local == "r" {
				err = d.DecodeElement(&b.Run, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				continue
			}
			err = d.Skip() // skip unsupported tags
			if err != nil {
				return err
			}
		}
	}
	return nil
}

type BookmarkEnd struct {
	XMLName xml.Name `xml:"w:bookmarkEnd,omitempty"`
	ID      string   `xml:"w:id,attr"`
}

func (b *BookmarkEnd) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// log.Println("BookmarkEnd", start.Attr)
	for _, attr := range start.Attr {
		if attr.Name.Local == "id" {
			b.ID = attr.Value
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

		if _, ok := t.(xml.EndElement); ok {
			break
		}
	}
	return nil
}
