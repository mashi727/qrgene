package main

import (
	"fmt"
	"log"
	"strconv"
	"strings"

	"github.com/signintech/gopdf"
	qrcode "github.com/skip2/go-qrcode"
	"github.com/tealeg/xlsx"
)

const (
	filepath = "20201215.xlsx"
	tempdir  = "./qr_temp"
)

func NumofSheets() int {
	file, _ := xlsx.OpenFile(filepath)
	return len(file.Sheets)
}

func MaxnumofRow(NumofSheet int) int {
	file, _ := xlsx.OpenFile(filepath)
	return file.Sheets[NumofSheet].MaxRow
}

func NameofSheet(NumofSheet int) string {
	file, _ := xlsx.OpenFile(filepath)
	return file.Sheets[NumofSheet].Name
}

func VCard(numofSheet, numofRow int) (string, string, string) {
	data, _ := xlsx.FileToSlice(filepath)
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
	No_Name := strconv.Itoa(numofRow) + "_" + FN[0] + data[numofSheet][numofRow][3]
	Name := FN[0] + data[numofSheet][numofRow][3]
	return No_Name, Name, vObject
}

func main() {
	No_Name, Name, vObject := VCard(2, 1)
	fmt.Println(No_Name, vObject)
	qrcode.WriteFile(vObject, qrcode.Medium, 128, No_Name+".png")
	// gopdf関連
	pdf := gopdf.GoPdf{}
	pdf.Start(gopdf.Config{PageSize: gopdf.Rect{H: 595.28, W: 841.89}}) //595.28, 841.89 = A4
	pdf.AddPage()
	var err error
	err = pdf.AddTTFFont("HiraMin0001", "./ヒラギノ明朝-ProN001.ttf")
	if err != nil {
		log.Print(err.Error())
		return
	}
	err = pdf.AddTTFFont("HiraMin0004", "./ヒラギノ明朝-ProN004.ttf")
	if err != nil {
		log.Print(err.Error())
		return
	}
	err = pdf.SetFont("HiraMin0004", "", 20)
	if err != nil {
		log.Print(err.Error())
		return
	}
	pdf.SetX(20)
	pdf.SetY(20)
	pdf.Cell(nil, NameofSheet(2))
	err = pdf.SetFont("HiraMin0001", "", 12)
	if err != nil {
		log.Print(err.Error())
		return
	}
	textw, _ := pdf.MeasureTextWidth(Name)
	fmt.Println(textw)
	pdf.SetX(25 + 48 - textw/2)
	pdf.SetY(57)
	pdf.Cell(nil, Name)

	pdf.Image(No_Name+".png", 25, 70, nil)
	pdf.WritePdf("hello.pdf")
	/*
		for i := 1; i < MaxnumofRow(2); i++ {
			Name, vObject := VCard(2, i) // (シート番号,行数（Row）)
			fmt.Println(Name, vObject)
		}

			var png []byte
			png, err := qrcode.Encode(vObject, qrcode.Medium, 256)
	*/
}
