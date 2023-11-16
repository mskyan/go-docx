package docx

import (
	"encoding/xml"
	"io"
)

// sample source
// <w:sectPr w:rsidR="00A75538" w:rsidRPr="00A75538" w:rsidSect="00976E94">
// <w:footerReference w:type="default" r:id="rId12"/>
// <w:pgSz w:w="11906" w:h="16838" w:code="9"/>
// <w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992" w:gutter="0"/>
// <w:cols w:space="425"/>
// <w:docGrid w:type="linesAndChars" w:linePitch="386"/>
// </w:sectPr>

type SectPr struct {
	XMLName         xml.Name         `xml:"w:sectPr"`
	RsidR           string           `xml:"w:rsidR,attr,omitempty"`
	RsidRPr         string           `xml:"w:rsidRPr,attr,omitempty"`
	RsidSect        string           `xml:"w:rsidSect,attr,omitempty"`
	FooterReference *FooterReference `xml:"w:footerReference,omitempty"`
	PgSz            *PgSz            `xml:"w:pgSz,omitempty"`
	PgMar           *PgMar           `xml:"w:pgMar,omitempty"`
	Cols            *Cols            `xml:"w:cols,omitempty"`
	DocGrid         *DocGrid         `xml:"w:docGrid,omitempty"`
}

type FooterReference struct {
	XMLName xml.Name `xml:"w:footerReference"`
	Type    string   `xml:"w:type,attr,omitempty"`
	Id      string   `xml:"r:id,attr,omitempty"`
}

type PgSz struct {
	XMLName xml.Name `xml:"w:pgSz"`
	W       string   `xml:"w:w,attr,omitempty"`
	H       string   `xml:"w:h,attr,omitempty"`
	Code    string   `xml:"w:code,attr,omitempty"`
}

type PgMar struct {
	XMLName xml.Name `xml:"w:pgMar"`
	Top     string   `xml:"w:top,attr,omitempty"`
	Right   string   `xml:"w:right,attr,omitempty"`
	Bottom  string   `xml:"w:bottom,attr,omitempty"`
	Left    string   `xml:"w:left,attr,omitempty"`
	Header  string   `xml:"w:header,attr,omitempty"`
	Footer  string   `xml:"w:footer,attr,omitempty"`
	Gutter  string   `xml:"w:gutter,attr,omitempty"`
}

type Cols struct {
	XMLName xml.Name `xml:"w:cols"`
	Space   string   `xml:"w:space,attr,omitempty"`
}

type DocGrid struct {
	XMLName   xml.Name `xml:"w:docGrid"`
	Type      string   `xml:"w:type,attr,omitempty"`
	LinePitch string   `xml:"w:linePitch,attr,omitempty"`
}

// Unmarshals ...

func (s *SectPr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "rsidR":
			s.RsidR = attr.Value
		case "rsidRPr":
			s.RsidRPr = attr.Value
		case "rsidSect":
			s.RsidSect = attr.Value
		}
	}
	for {
		t, _ := d.Token()
		if t == nil {
			return nil
		}
		switch se := t.(type) {
		case xml.StartElement:
			switch se.Name.Local {
			case "footerReference":
				var v FooterReference
				if err := d.DecodeElement(&v, &se); err != nil {
					return err
				}
				s.FooterReference = &v
			case "pgSz":
				var v PgSz
				if err := d.DecodeElement(&v, &se); err != nil {
					return err
				}
				s.PgSz = &v
			case "pgMar":
				var v PgMar
				if err := d.DecodeElement(&v, &se); err != nil {
					return err
				}
				s.PgMar = &v
			case "cols":
				var v Cols
				if err := d.DecodeElement(&v, &se); err != nil {
					return err
				}
				s.Cols = &v
			case "docGrid":
				var v DocGrid
				if err := d.DecodeElement(&v, &se); err != nil {
					return err
				}
				s.DocGrid = &v
			}
		}
	}
}

func (f *FooterReference) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "type":
			f.Type = attr.Value
		case "id":
			f.Id = attr.Value
		}
	}

	for {
		_, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
	}

	return nil
}

func (p *PgSz) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "w":
			p.W = attr.Value
		case "h":
			p.H = attr.Value
		case "code":
			p.Code = attr.Value
		}
	}

	for {
		_, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
	}

	return nil
}

func (p *PgMar) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "top":
			p.Top = attr.Value
		case "right":
			p.Right = attr.Value
		case "bottom":
			p.Bottom = attr.Value
		case "left":
			p.Left = attr.Value
		case "header":
			p.Header = attr.Value
		case "footer":
			p.Footer = attr.Value
		case "gutter":
			p.Gutter = attr.Value
		}
	}

	for {
		_, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
	}

	return nil
}

func (c *Cols) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "space":
			c.Space = attr.Value
		}
	}

	for {
		_, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
	}

	return nil
}

func (dg *DocGrid) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "type":
			dg.Type = attr.Value
		case "linePitch":
			dg.LinePitch = attr.Value
		}
	}

	// skip to end of element
	for {
		_, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
	}

	return nil
}
