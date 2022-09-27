//go:generate goversioninfo -icon=QRGene.ico
package main

import (
	"bytes"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"runtime"
	"strings"
	"time"

	"github.com/signintech/gopdf"
	qrcode "github.com/skip2/go-qrcode"
	"github.com/tealeg/xlsx"
)

var excelpath, bodyFont, indexFont string

func NumofSheets() int {
	file, _ := xlsx.OpenFile(excelpath)
	return len(file.Sheets)
}

func MaxnumofRow(NumofSheet int) int {
	file, _ := xlsx.OpenFile(excelpath)
	return file.Sheets[NumofSheet].MaxRow
}

func NameofSheet(NumofSheet int) string {
	file, _ := xlsx.OpenFile(excelpath)
	return file.Sheets[NumofSheet].Name
}

func VCard(numofSheet, numofRow int) (string, string) {
	data, _ := xlsx.FileToSlice(excelpath)
	// 名前とふりがなの分割
	FN := strings.Split(strings.Replace(data[numofSheet][numofRow][0], "　", " ", -1), " ") //全角スーペスを半角スペースに
	Yomi := strings.Split(strings.Replace(data[numofSheet][numofRow][1], "　", " ", -1), " ")
	// vCard用データの作成
	vObject := "\nBEGIN:VCARD\nVERSION:3.0\n" +
		"FN:" + FN[0] + FN[1] + "\n" +
		"N:" + FN[0] + ";" + FN[1] + ";;;\n" +
		"X-PHONETIC-FIRST-NAME:" + Yomi[0] + "\n" +
		"X-PHONETIC-LAST-NAME:" + Yomi[1] + "\n" +
		"ORG:" + data[numofSheet][numofRow][2] + "\n" +
		"TITLE:" + data[numofSheet][numofRow][3] + "\n" +
		"EMAIL:" + data[numofSheet][numofRow][4] + "\n" +
		"TEL;TYPE=CELL:" + data[numofSheet][numofRow][5] + "\n" +
		"END:VCARD\n"
	// フルネームで表記したい場合
	Name := FN[0] + FN[1]
	// 役職で表記したい場合
	// Name := FN[0] + data[numofSheet][numofRow][2]
	return Name, vObject
}

func Cm(length float64) float64 {
	//    Convenience constructor for pt in centimeters
	return length * 72 / 2.54
}
func IsExists(filename string) bool {
	_, err := os.Stat(filename)
	return err == nil
}
func makePdf() {
	var err error
	ostype := runtime.GOOS
	// 平成明朝
	//bodyFont := "./fonts/DFHsm3003.ttf"
	//indexFont := "./fonts/DFHsm7003.ttf"
	if true == IsExists("./fonts/HMProN001.ttf") && true == IsExists("./fonts/HMProN004.ttf") {
		bodyFont = "./fonts/HMProN001.ttf"
		indexFont = "./fonts/HMProN004.ttf"
		fmt.Println("ヒラギノ明朝で作成されます。")
	} else if true == IsExists("./fonts/bodyFont.ttf") && true == IsExists("./fonts/indexFont.ttf") {
		bodyFont = "./fonts/bodyFont.ttf"
		indexFont = "./fonts/indexFont.ttf"
		fmt.Println("bodyFont..で作成されます。")
	} else if ostype == "windows" {
		// 游明朝フォント（Windows Dafault）
		bodyFont = "C:/Windows/fonts/yuminl.ttf"
		indexFont = "C:/Windows/fonts/yumindb.ttf"
		fmt.Println("OSは" + ostype + "です。游明朝フォントで作成されます。")
	} else {
		fmt.Println("Fonts not found.")
	}
	pdf := gopdf.GoPdf{}
	pdf.Start(gopdf.Config{PageSize: gopdf.Rect{H: 595.28, W: 841.89}}) //595.28pt, 841.89pt = A4
	err = pdf.AddTTFFont("indexFont", indexFont)
	if err != nil {
		log.Print(err.Error())
		return
	}
	err = pdf.AddTTFFont("bodyFont", bodyFont)
	if err != nil {
		log.Print(err.Error())
		return
	}
	for numofSheet := 0; numofSheet < NumofSheets(); numofSheet++ {
		// シート毎に改頁
		pdf.AddPage()
		// フォントの設定
		pdf.SetFont("indexFont", "", 20)
		pdf.SetX(Cm(1.25))
		pdf.SetY(Cm(1 + 0.6))
		pdf.Cell(nil, NameofSheet(numofSheet))
		// 日付の印刷用の処理
		// 右寄せを行うためにテキスト幅を計算
		date := time.Now().Format("2006年1月2日")
		pdf.SetFont("indexFont", "", 12)
		textw, _ := pdf.MeasureTextWidth(date + "作成")
		rmargin := textw * 2.54 / 72 //cmに変換
		//
		fmt.Println(textw)
		fmt.Println(rmargin)
		pdf.SetX(Cm(29.7 - rmargin - 1.15))
		pdf.SetY(Cm(1.3))
		pdf.Cell(nil, date+"作成")

		// 個別の出力（横長に4x8個出力）
		for i := 1; i < MaxnumofRow(numofSheet); i++ {
			Name, vObject := VCard(numofSheet, i)
			png, _ := qrcode.Encode(vObject, qrcode.Medium, 150)
			imgHolder, _ := gopdf.ImageHolderByReader(bytes.NewReader(png))
			pdf.SetFont("bodyFont", "", 11)
			if i <= 8 {
				pdf.SetX(Cm(1.3) + Cm(3.5)*float64((i-1)%8))
				pdf.SetY(Cm(2.65 + 0.5))
				pdf.ImageByHolder(imgHolder, Cm(1.15)+Cm(3.5)*float64((i-1)%8), Cm(3+0.45), nil)
				pdf.Cell(nil, Name)
			} else if i <= 16 {
				pdf.SetX(Cm(1.3) + Cm(3.5)*float64((i-1)%8))
				pdf.SetY(Cm(7.15 + 0.5))

				pdf.ImageByHolder(imgHolder, Cm(1.15)+Cm(3.5)*float64((i-1)%8), Cm(7.5+0.45), nil)
				pdf.Cell(nil, Name)
			} else if i <= 24 {
				pdf.SetX(Cm(1.3) + Cm(3.5)*float64((i-1)%8))
				pdf.SetY(Cm(11.65 + 0.5))
				pdf.ImageByHolder(imgHolder, Cm(1.15)+Cm(3.5)*float64((i-1)%8), Cm(12+0.45), nil)
				pdf.Cell(nil, Name)
			} else if i <= 32 {
				pdf.SetX(Cm(1.3) + Cm(3.5)*float64((i-1)%8))
				pdf.SetY(Cm(16.15 + 0.5))
				pdf.ImageByHolder(imgHolder, Cm(1.15)+Cm(3.5)*float64((i-1)%8), Cm(16.5+0.45), nil)
				pdf.Cell(nil, Name)
			}
		}
	}
	rep := regexp.MustCompile(`.xlsx`)
	pdfbasename := filepath.Base(rep.ReplaceAllString(excelpath, ""))
	pdf.WritePdf(pdfbasename + ".pdf")
	fmt.Println("「" + pdfbasename + ".pdf」が作成されました。")
}

func main() {
	if len(os.Args) < 2 {
		fmt.Println("ERROR: エクセルファイルを指定してください。")
		os.Exit(1)
	}
	excelpath = os.Args[1]
	fmt.Println("「" + filepath.Base(excelpath) + "」を処理しています。")
	makePdf() // (シート番号)

}
