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
	"reflect"
	"strconv"
	"strings"
)

// ParagraphProperties <w:pPr>
// Properties in this struct are defined in structeffects.go
type ParagraphProperties struct {
	XMLName         xml.Name `xml:"w:pPr,omitempty"`
	Tabs            *Tabs
	Spacing         *Spacing
	Ind             *Ind
	Justification   *Justification
	Shade           *Shade
	Kern            *Kern
	Style           *Style
	TextAlignment   *TextAlignment
	AdjustRightInd  *AdjustRightInd
	SnapToGrid      *SnapToGrid
	Kinsoku         *Kinsoku
	OverflowPunct   *OverflowPunct
	NumPr           *NumPr
	KeepNext        *KeepNext
	KeepLines       *KeepLines
	WidowControl    *WidowControl
	PageBreakBefore *PageBreakBefore
	SectPr          *SectPr
	PBDR            *PBDR

	RunProperties *RunProperties
}

// UnmarshalXML ...
func (p *ParagraphProperties) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			log.Println("UnmarshalXML ParagraphProperties error:", err)
			return err
		}
		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "tabs":
				var value Tabs
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				p.Tabs = &value
			case "spacing":
				var value Spacing
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				p.Spacing = &value
			case "ind":
				var value Ind
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				p.Ind = &value
			case "jc":
				p.Justification = &Justification{Val: getAtt(tt.Attr, "val")}
			case "shd":
				var value Shade
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				p.Shade = &value
			case "kern":
				var value Kern
				v := getAtt(tt.Attr, "val")
				if v == "" {
					continue
				}
				value.Val, err = GetInt64(v)
				if err != nil {
					return err
				}
				p.Kern = &value
			case "rPr":
				var value RunProperties
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				p.RunProperties = &value
			case "pStyle":
				p.Style = &Style{Val: getAtt(tt.Attr, "val")}
			case "textAlignment":
				p.TextAlignment = &TextAlignment{Val: getAtt(tt.Attr, "val")}
			case "adjustRightInd":
				var value AdjustRightInd
				v := getAtt(tt.Attr, "val")
				if v == "" {
					continue
				}
				value.Val, err = GetInt(v)
				if err != nil {
					return err
				}
				p.AdjustRightInd = &value
			case "snapToGrid":
				var value SnapToGrid
				v := getAtt(tt.Attr, "val")
				if v == "" {
					continue
				}
				value.Val, err = GetInt(v)
				if err != nil {
					return err
				}
				p.SnapToGrid = &value
			case "kinsoku":
				var value Kinsoku
				v := getAtt(tt.Attr, "val")
				if v == "" {
					continue
				}
				value.Val, err = GetInt(v)
				if err != nil {
					return err
				}
				p.Kinsoku = &value
			case "overflowPunct":
				var value OverflowPunct
				v := getAtt(tt.Attr, "val")
				if v == "" {
					continue
				}
				value.Val, err = GetInt(v)
				if err != nil {
					return err
				}
				p.OverflowPunct = &value
			case "numPr":
				var value NumPr
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				p.NumPr = &value
			case "keepNext":
				var value KeepNext
				v := getAtt(tt.Attr, "val")
				if v == "" {
					v = "0"
				}
				value.Val = v
				p.KeepNext = &value
			case "keepLines":
				// ここで、値を取得しておく。
				// decoder に渡さないことで
				// 閉じタグを別に造らない。
				var value KeepLines
				v := getAtt(tt.Attr, "val")
				if v == "" {
					v = "0"
				}
				value.Val = v
				p.KeepLines = &value
			case "widowControl":
				var value WidowControl
				v := getAtt(tt.Attr, "val")
				if v == "" {
					v = "0"
				}
				value.Val = v
				p.WidowControl = &value
			case "sectPr":
				var value SectPr
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				p.SectPr = &value
			case "pageBreakBefore":
				var value PageBreakBefore
				v := getAtt(tt.Attr, "val")
				if v == "" {
					v = "0"
				}
				value.Val = v
				p.PageBreakBefore = &value
			case "pBdr":
				var value PBDR
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				p.PBDR = &value
			default:
				// 取り損ねた値を log に表示
				log.Println("UnmarshalXML ParagraphProperties unsupported, skip:", tt.Name.Local)

				err = d.Skip() // skip unsupported tags
				if err != nil {
					return err
				}
				continue
			}
		}

		// consume end tag
		// if _, ok := t.(xml.EndElement); ok {
		// 	break
		// }
	}
	return nil
}

// Paragraph <w:p>
type Paragraph struct {
	XMLName xml.Name `xml:"w:p,omitempty"`

	ParaId       string `xml:"w14:paraId,attr,omitempty"`
	RsidR        string `xml:"w:rsidR,attr,omitempty"`
	RsidRPr      string `xml:"w:rsidRPr,attr,omitempty"`
	RsidRDefault string `xml:"w:rsidRDefault,attr,omitempty"`
	RsidP        string `xml:"w:rsidP,attr,omitempty"`
	TextId       string `xml:"w14:textId,attr,omitempty"`

	Hyperlink     *[]*Hyperlink     `xml:"w:hyperlink,omitempty"`     // 0 or more
	BookmarkStart *[]*BookmarkStart `xml:"w:bookmarkStart,omitempty"` // 0 or more
	BookmarkEnd   *[]*BookmarkEnd   `xml:"w:bookmarkEnd,omitempty"`   // 0 or more

	StructuredDocumentTag *[]*StructuredDocumentTag `xml:"w:sdt,omitempty"` // 0 or more

	KeepLines       *KeepLines       `xml:"w:keepLines,omitempty"`
	WidowControl    *WidowControl    `xml:"w:widowControl,omitempty"`
	PageBreakBefore *PageBreakBefore `xml:"w:pageBreakBefore,omitempty"`
	PBDR            *PBDR            `xml:"w:pBdr,omitempty"`
	Ind             *Ind             `xml:"w:ind,omitempty"`
	Spacing         *Spacing         `xml:"w:spacing,omitempty"`
	Shd             *Shade           `xml:"w:shd,omitempty"`

	Properties *ParagraphProperties
	Children   []interface{}

	file *Docx
}

func (p *Paragraph) String() string {
	// rPr は Children の中には、含まれず、p.Properties に含まれる
	// 並びとしては、連動しているので、Children に他と同様に含めた方が
	// 表示処理上は。都合が良い。
	//
	// numberting.xml のを内容を読み取る必要がある。
	// p.file *Docx から struct 加工済み numbering.xml を取得できるようにする。

	// if (*p.file).Numbering.XMLName.Local == "" {
	// 	log.Println("Paragraph.Numbering is not set, skip")
	// } else {
	// 	log.Println("Paragraph.Numbering is set, proceed")
	// }

	// NUmID の Val は int と定義。structnumbering 側もいずれ合わせる。

	sb := strings.Builder{}
	for _, c := range p.Children {
		switch o := c.(type) {
		case *Hyperlink:
			id := o.ID
			// there are multiple Run in Hyperlink
			// range o.Runs
			for _, r := range *o.Runs {
				text := r.InstrText
				link, err := p.file.ReferTarget(id)
				sb.WriteString("[")
				sb.WriteString(text)
				sb.WriteString("](")
				if err != nil {
					sb.WriteString(id)
				} else {
					sb.WriteString(link)
				}
				sb.WriteByte(')')
			}
		case *Run:
			for _, c := range o.Children {
				switch x := c.(type) {
				case *Text:
					sb.WriteString(x.Text)
				case *Tab:
					sb.WriteByte('\t')
				case *BarterRabbet:
					sb.WriteByte('\n')
				case *Drawing:
					if x.Inline != nil {
						sb.WriteString(x.Inline.String())
						continue
					}
					if x.Anchor != nil {
						sb.WriteString(x.Anchor.String())
						continue
					}
				}
			}
		// pPr
		case *ParagraphProperties:
			// NumPr と KeepNext の値があるかどうか。
			if o.NumPr != nil {
				numbered := getNumberedString(p, o.NumPr)
				sb.WriteString(numbered)
			}
			if o.KeepNext != nil {
				// log.Println("ParagraphProperties.KeepNext is set, proceed")
				// TODO
			}

			continue
		default:
			continue
		}
	}
	return sb.String()
}

// Numbering ...
// 一旦変数に格納すれば済むもの。
// これとは別に、「値の入力⇒定義に基づいた変換」という関数自体を id と結び付けて格納しておく map を用意する。

type abstractNumFunc func(args *[]int) string

var numbering *Numbering
var abstractNums *[]AbstractNum
var nums *[]*Num

var numIDToAbstractNumIDMap = make(map[int]int)

var abstractNumIDToAbstractNumMap = make(map[int]*AbstractNum)

// key is NumPr.NumId, value is func
var abstractNumToFuncMap = make(map[int]map[int]*abstractNumFunc)

// abstractNumId キーごとに、また ilvl 階層ごとに、カウント値を格納する。
var lvlCountListMap = make(map[int]*[]int)

// 本編では連続出現 ID が同じ限りはカウントで、
// そうでない場合はリセットされる。
// 間に連続判定関数須。
var numIDCountMap = make(map[int]int)
var previousNumID int = -1

// リセットは要らないかもしれない。
func countNumID(numID int) int {
	numIDCountMap[numID]++
	// if numID == previousNumID {
	// 	numIDCountMap[numID]++
	// } else {
	// 	// zero start
	// 	numIDCountMap[numID] = 0
	// }
	// previousNumID = numID
	return numIDCountMap[numID] - 1
}

func getNumberedString(p *Paragraph, numPr *NumPr) string {
	// func getNumberedString(p *Paragraph, ilvl, numId int) string {
	if numbering == nil {
		numbering = &p.file.Numbering
		abstractNums = numbering.AbstractNums
		nums = numbering.Nums
		generateNumIDToAbstractNumIDMap()
		generateAbstractNumIDToAbstractNumMap()
		generateAbstractNumToFuncMap()
	}

	// NumPr から NumID を取得
	numID := numPr.NumID.Val
	numIDCount := countNumID(numID)
	// log.Println("numIDCount:", numIDCount)

	// 対象 AbstructNumID を取得
	abstractNumID := numIDToAbstractNumIDMap[numID]
	// abstractNum := abstractNumIDToAbstractNumMap[abstractNumID]
	iLvl := numPr.Ilvl.Val
	// log.Println("iLvl:", iLvl)

	f := abstractNumToFuncMap[abstractNumID][iLvl]
	// log.Println("f:", *f)

	iLvlCountList := getILvlCountList(abstractNumID, iLvl, numIDCount == 0)
	// if iLvlCountList == nil {
	// 	log.Println("iLvlCountList is nil")
	// } else {
	// 	log.Println("iLvlCountList:", *iLvlCountList)
	// }

	var s string
	if f != nil {
		s = (*f)(iLvlCountList)
	} else {
		log.Println("abstractNumID:", abstractNumID, "iLvl:", iLvl, "f is nil")
		s = "＠"
	}

	// log.Println("s:", s)

	return s
}

func generateNumIDToAbstractNumIDMap() {
	for _, num := range *nums {
		numID, err := GetInt(num.NumID)
		if err != nil {
			continue
		}
		abstractNumID, err := GetInt(num.AbstractNumID.Val)
		if err != nil {
			continue
		}
		numIDToAbstractNumIDMap[numID] = abstractNumID
	}
}

func generateAbstractNumIDToAbstractNumMap() {
	for _, abstractNum := range *abstractNums {
		abstractNumID, err := GetInt(abstractNum.AbstractNumID)
		if err != nil {
			continue
		}

		// ここでは被りなし
		// if abstractNumIDToAbstractNumMap[abstractNumID] != nil {
		// 	log.Println("abstractNumIDToAbstractNumMap[abstractNumID] is not nil")
		// }

		// log.Println("abstractNumID:", abstractNumID, "abstractNum:", abstractNum)

		abstractNumIDToAbstractNumMap[abstractNumID] = &abstractNum
	}
}

func generateAbstractNumToFuncMap() {
	numOfFunc := 0
	for _, abstractNum := range *abstractNums {
		abstractNumID, err := GetInt(abstractNum.AbstractNumID)
		if err != nil {
			continue
		}
		iLvlKeys := make([]int, 0, len(*abstractNum.Lvl))
		for _, lvl := range *abstractNum.Lvl {
			iLvlKeys = append(iLvlKeys, lvl.ILvl)
		}
		// log.Println("iLvlKeys:", iLvlKeys)

		abstractNumToFuncMap[abstractNumID] = make(map[int]*abstractNumFunc)
		for _, iLvl := range iLvlKeys {
			// log.Println("iLvl:", iLvl)

			// ★関数実行の以外の場では、正確な値が出ている。
			lvlText := (*abstractNum.Lvl)[iLvl].LvlText.Val
			// log.Println("lvlText:", lvlText)

			// f := getAbstractNumFunc(abstractNumID, iLvl)
			abstractNumToFuncMap[abstractNumID][iLvl] = getAbstractNumFunc(abstractNumID, iLvl, lvlText)
			// abstractNumToFuncMap[abstractNumID][iLvl] = getAbstractNumFunc(abstractNumID, iLvl)
			// abstractNumToFuncMap[abstractNumID][iLvl] = &f

			// log.Println("abstractNumID:", abstractNumID, "iLvl:", iLvl, "f:", f)

			numOfFunc++
		}
	}

	// log.Println("numOfFunc:", numOfFunc)
}

func getILvlCountList(abstractNumID, iLvl int, isRestart bool) *[]int {
	countList := lvlCountListMap[abstractNumID]
	abstractNum := abstractNumIDToAbstractNumMap[abstractNumID]
	// log.Println("abstractNum:", abstractNum)
	start, err := GetInt((*abstractNum.Lvl)[iLvl].Start.Val)
	if err != nil {
		// log.Println("GetInt error:", err)
		start = 0
	}

	if countList == nil || isRestart {
		countList = &[]int{start}
		lvlCountListMap[abstractNumID] = countList
	} else {
		if len(*countList) == iLvl {
			// if len(*countList) < iLvl {
			*countList = append(*countList, start)
		} else {
			(*countList)[iLvl]++
		}
	}
	// log.Println("here.")
	return countList
}

func getAbstractNumFunc(abstractNumID int,
	iLvl int, lvlText string) *abstractNumFunc {
	// log.Println("abstractNumID:", abstractNumID, "iLvl:", iLvl)
	abstractNum := abstractNumIDToAbstractNumMap[abstractNumID]
	numFmt := (*abstractNum.Lvl)[iLvl].NumFmt.Val
	// lvlText := (*abstractNum.Lvl)[iLvl].LvlText.Val
	// log.Println("lvlText:", lvlText)

	f := func(iLvlCountList *[]int) string {
		list := *iLvlCountList
		var lvlTextReplaced string
		// list から順に値を取り出す。
		for _, val := range list {
			n := GetFormatNumber(val, numFmt)

			// lvlText から %i を探し、n に置き換える。
			// "%i" がなければ何もしない。
			// iStr := "%" + string(i+1)

			// int to string
			iStr := strconv.Itoa(iLvl + 1)
			// log.Println("iStr:", iStr)

			// log.Println("lvlText(before):", lvlText)

			lvlTextReplaced = strings.Replace(lvlText, "%"+iStr, n, -1)
			// log.Println(
			// "val:", val,
			// "n:", n,
			// "lvlTextReplaced:", lvlTextReplaced)
		}

		return lvlTextReplaced
		// return lvlText
	}

	// cast?
	return (*abstractNumFunc)(&f)
}

func GetFormatNumber(val int, formatType string) string {
	switch formatType {
	case "bullet":
		return ""
	case "decimal":
		return strconv.Itoa(val)
	case "decimalFullWidth":
		return getDecimalFullWidth(val)
	case "decimalEnclosedParen":
		return "(" + strconv.Itoa(val) + ")"
	case "decimalEnclosedCircle":
		return decimalEnclosedCircleMap[val]
	case "irohaFullWidth":
		return irohaFullWidthMap[val]
	default:
		// TODO
		return "[unknown format type]: " + formatType + " " + strconv.Itoa(val)
	}
}

// 半角数値を日本語全角数値へ変換する関数
func getDecimalFullWidth(val int) string {
	// 全角数値の文字列を定義
	fullWidthNums := []string{"０", "１", "２", "３", "４", "５", "６", "７", "８", "９"}

	// val を文字列に変換
	strVal := strconv.Itoa(val)

	// 各桁の数値を全角数値に変換
	var result string
	for _, c := range strVal {
		if '0' <= c && c <= '9' {
			result += fullWidthNums[c-'0']
		} else {
			result += string(c)
		}
	}

	return result
}

// 1: ①, 2: ②, 3: ③, 4: ④, 5: ⑤, 6: ⑥, 7: ⑦, 8: ⑧, 9: ⑨, 10: ⑩
var decimalEnclosedCircleMap = map[int]string{
	1:  "①",
	2:  "②",
	3:  "③",
	4:  "④",
	5:  "⑤",
	6:  "⑥",
	7:  "⑦",
	8:  "⑧",
	9:  "⑨",
	10: "⑩",
}

// func getEnclosedCircle(val int) string {
// 	return decimalEnclosedCircleMap[val]
// }

// 1: イ, 2: ロ, 3: ハ, 4: ニ, 5: ホ, 6: ヘ, 7: ト, 8: チ, 9: リ, 10: ヌ, 11: ル, 12: ヲ, 13: ワ, 14: カ, 15: ヨ, 16: タ, 17: レ, 18: ソ, 19: ツ, 20: ネ, 21: ナ, 22: ラ, 23: ム, 24: ウ, 25: ヰ, 26: ノ, 27: オ, 28: ク, 29: ヤ, 30: マ, 31: ケ, 32: フ, 33: コ, 34: エ, 35: テ, 36: ア, 37: サ, 38: キ, 39: ユ, 40: メ, 41: ミ, 42: シ, 43: ヱ, 44: ヒ, 45: モ, 46: セ, 47: ス, 48: ン, 49: ゛, 50: ゜, 51: ヴ, 52: ヽ, 53: ヾ
var irohaFullWidthMap = map[int]string{
	1:  "イ",
	2:  "ロ",
	3:  "ハ",
	4:  "ニ",
	5:  "ホ",
	6:  "ヘ",
	7:  "ト",
	8:  "チ",
	9:  "リ",
	10: "ヌ",
	11: "ル",
	12: "ヲ",
}

// UnmarshalXML ...
func (p *Paragraph) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// 生成された document.xml の p にちゃんと attr の値群は入っている。
	for _, attr := range start.Attr {
		// log.Println("Paragraph UnmarshalXML attr:", attr.Name.Local, attr.Value)
		switch attr.Name.Local {
		case "paraId":
			p.ParaId = attr.Value
		case "rsidR":
			p.RsidR = attr.Value
		case "rsidRPr":
			p.RsidRPr = attr.Value
		case "rsidRDefault":
			p.RsidRDefault = attr.Value
		case "rsidP":
			p.RsidP = attr.Value
		case "textId":
			p.TextId = attr.Value
		default:
			// ignore other attributes
		}
	}
	children := make([]interface{}, 0, 64)
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
		if tt, ok := t.(xml.StartElement); ok {
			// log.Println("Paragraph UnmarshalXML:", tt.Name.Local)
			var elem interface{}
			switch tt.Name.Local {
			case "pageBreakBefore":
				var value PageBreakBefore
				v := getAtt(tt.Attr, "val")
				if v == "" {
					v = "0"
				}
				p.PageBreakBefore = &value
				elem = &value
			case "keepLines":
				var value KeepLines
				v := getAtt(tt.Attr, "val")
				if v == "" {
					v = "0"
				}
				value.Val = v
				p.KeepLines = &value
				elem = &value
			case "hyperlink":
				// log.Println("hyperlink")
				var value Hyperlink
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				if p.Hyperlink == nil {
					p.Hyperlink = &[]*Hyperlink{&value}
				} else {
					*p.Hyperlink = append(*p.Hyperlink, &value)
				}
				elem = &value
			case "bookmarkStart":
				var value BookmarkStart
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}

				if p.BookmarkStart == nil {
					p.BookmarkStart = &[]*BookmarkStart{&value}
				} else {
					*p.BookmarkStart = append(*p.BookmarkStart, &value)
				}
				elem = &value
			case "bookmarkEnd":
				var value BookmarkEnd
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}

				if p.BookmarkEnd == nil {
					p.BookmarkEnd = &[]*BookmarkEnd{&value}
				} else {
					*p.BookmarkEnd = append(*p.BookmarkEnd, &value)
				}
				elem = &value
			case "sdt":
				var value StructuredDocumentTag
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}

				if p.StructuredDocumentTag == nil {
					p.StructuredDocumentTag = &[]*StructuredDocumentTag{&value}
				} else {
					*p.StructuredDocumentTag = append(*p.StructuredDocumentTag, &value)
				}

				elem = &value
				// log.Println("sdt added in paragraph")
			case "r":
				var value Run
				value.file = p.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				elem = &value
			case "rPr":
				var value RunProperties
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				elem = &value
			case "pPr":
				var value ParagraphProperties
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				// p.Properties = &value

				// 重複するのでひとまず省く
				elem = &value
			default:
				log.Println("UnmarshalXML Paragraph unsupported, skip:", tt.Name.Local)
				err = d.Skip() // skip unsupported tags
				if err != nil {
					return err
				}
				continue
			}
			children = append(children, elem)
		}
	}
	p.Children = children
	return nil
}

// KeepElements keep named elems amd removes others
//
// names: *docx.Hyperlink *docx.Run *docx.RunProperties
func (p *Paragraph) KeepElements(name ...string) {
	items := make([]interface{}, 0, len(p.Children))
	namemap := make(map[string]struct{}, len(name)*2)
	for _, n := range name {
		namemap[n] = struct{}{}
	}
	for _, item := range p.Children {
		_, ok := namemap[reflect.ValueOf(item).Type().String()]
		if ok {
			items = append(items, item)
		}
	}
	p.Children = items
}

// DropCanvas drops all canvases in paragraph
func (p *Paragraph) DropCanvas() {
	for _, pc := range p.Children {
		if r, ok := pc.(*Run); ok {
			nrc := make([]interface{}, 0, len(r.Children))
			for _, rc := range r.Children {
				if d, ok := rc.(*Drawing); ok {
					if d.Inline != nil && d.Inline.Graphic != nil && d.Inline.Graphic.GraphicData != nil {
						if d.Inline.Graphic.GraphicData.Canvas != nil {
							continue
						}
					}
					if d.Anchor != nil && d.Anchor.Graphic != nil && d.Anchor.Graphic.GraphicData != nil {
						if d.Anchor.Graphic.GraphicData.Canvas != nil {
							continue
						}
					}
				}
				nrc = append(nrc, rc)
			}
			r.Children = nrc
		}
	}
}

// DropShape drops all shapes in paragraph
func (p *Paragraph) DropShape() {
	for _, pc := range p.Children {
		if r, ok := pc.(*Run); ok {
			nrc := make([]interface{}, 0, len(r.Children))
			for _, rc := range r.Children {
				if d, ok := rc.(*Drawing); ok {
					if d.Inline != nil && d.Inline.Graphic != nil && d.Inline.Graphic.GraphicData != nil {
						if d.Inline.Graphic.GraphicData.Shape != nil {
							continue
						}
					}
					if d.Anchor != nil && d.Anchor.Graphic != nil && d.Anchor.Graphic.GraphicData != nil {
						if d.Anchor.Graphic.GraphicData.Shape != nil {
							continue
						}
					}
				}
				nrc = append(nrc, rc)
			}
			r.Children = nrc
		}
	}
}

// DropGroup drops all groups in paragraph
func (p *Paragraph) DropGroup() {
	for _, pc := range p.Children {
		if r, ok := pc.(*Run); ok {
			nrc := make([]interface{}, 0, len(r.Children))
			for _, rc := range r.Children {
				if d, ok := rc.(*Drawing); ok {
					if d.Inline != nil && d.Inline.Graphic != nil && d.Inline.Graphic.GraphicData != nil {
						if d.Inline.Graphic.GraphicData.Group != nil {
							continue
						}
					}
					if d.Anchor != nil && d.Anchor.Graphic != nil && d.Anchor.Graphic.GraphicData != nil {
						if d.Anchor.Graphic.GraphicData.Group != nil {
							continue
						}
					}
				}
				nrc = append(nrc, rc)
			}
			r.Children = nrc
		}
	}
}

// DropShapeAndCanvas drops all shapes and canvases in paragraph
func (p *Paragraph) DropShapeAndCanvas() {
	for _, pc := range p.Children {
		if r, ok := pc.(*Run); ok {
			nrc := make([]interface{}, 0, len(r.Children))
			for _, rc := range r.Children {
				if d, ok := rc.(*Drawing); ok {
					if d.Inline != nil && d.Inline.Graphic != nil && d.Inline.Graphic.GraphicData != nil {
						if d.Inline.Graphic.GraphicData.Shape != nil || d.Inline.Graphic.GraphicData.Canvas != nil {
							continue
						}
					}
					if d.Anchor != nil && d.Anchor.Graphic != nil && d.Anchor.Graphic.GraphicData != nil {
						if d.Anchor.Graphic.GraphicData.Shape != nil || d.Anchor.Graphic.GraphicData.Canvas != nil {
							continue
						}
					}
				}
				nrc = append(nrc, rc)
			}
			r.Children = nrc
		}
	}
}

// DropShapeAndCanvasAndGroup drops all shapes, canvases and groups in paragraph
func (p *Paragraph) DropShapeAndCanvasAndGroup() {
	for _, pc := range p.Children {
		if r, ok := pc.(*Run); ok {
			nrc := make([]interface{}, 0, len(r.Children))
			for _, rc := range r.Children {
				if d, ok := rc.(*Drawing); ok {
					if d.Inline != nil && d.Inline.Graphic != nil && d.Inline.Graphic.GraphicData != nil {
						if d.Inline.Graphic.GraphicData.Shape != nil || d.Inline.Graphic.GraphicData.Canvas != nil || d.Inline.Graphic.GraphicData.Group != nil {
							continue
						}
					}
					if d.Anchor != nil && d.Anchor.Graphic != nil && d.Anchor.Graphic.GraphicData != nil {
						if d.Anchor.Graphic.GraphicData.Shape != nil || d.Anchor.Graphic.GraphicData.Canvas != nil || d.Anchor.Graphic.GraphicData.Group != nil {
							continue
						}
					}
				}
				nrc = append(nrc, rc)
			}
			r.Children = nrc
		}
	}
}

// DropNilPicture drops all drawings with nil picture in paragraph
func (p *Paragraph) DropNilPicture() {
	for _, pc := range p.Children {
		if r, ok := pc.(*Run); ok {
			nrc := make([]interface{}, 0, len(r.Children))
			for _, rc := range r.Children {
				if d, ok := rc.(*Drawing); ok {
					if d.Inline == nil && d.Anchor == nil {
						continue
					}
					if (d.Inline != nil && d.Inline.Graphic == nil) || (d.Anchor != nil && d.Anchor.Graphic == nil) {
						continue
					}
					if d.Inline != nil && d.Inline.Graphic != nil && d.Inline.Graphic.GraphicData == nil {
						continue
					}
					if d.Anchor != nil && d.Anchor.Graphic != nil && d.Anchor.Graphic.GraphicData == nil {
						continue
					}
					if d.Inline != nil && d.Inline.Graphic != nil && d.Inline.Graphic.GraphicData != nil {
						if d.Inline.Graphic.GraphicData.Pic == nil {
							continue
						}
					}
					if d.Anchor != nil && d.Anchor.Graphic != nil && d.Anchor.Graphic.GraphicData != nil {
						if d.Anchor.Graphic.GraphicData.Pic == nil {
							continue
						}
					}
				}
				nrc = append(nrc, rc)
			}
			r.Children = nrc
		}
	}
}
