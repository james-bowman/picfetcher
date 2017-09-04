package main

import (
	"log"
	"github.com/tealeg/xlsx"
	"github.com/360EntSecGroup-Skylar/excelize"
	"net/http"
	"bytes"
	"io/ioutil"
	"fmt"
	"sync"
	_ "image/jpeg"
)

func main() {
	rows := make(chan *xlsx.Row, 5)
	enrichedRows := make(chan []byte, 5)
	eof := make(chan bool)

	excelFileName := "input.xlsx"
    xlFile, err := xlsx.OpenFile(excelFileName)
    if err != nil {
        log.Fatal(err)
	}
	
	go Write(enrichedRows, eof)
	
	var fetcherGroup sync.WaitGroup
	for i:=0; i < 10; i++ {
		go Fetch(rows, enrichedRows, &fetcherGroup)
	}

    for _, sheet := range xlFile.Sheets {
		for _, row := range sheet.Rows {
			rows <- row
		}
	}

	close(rows)
	fetcherGroup.Wait()
	close(enrichedRows)
	<- eof
}

func Fetch(in chan *xlsx.Row, out chan []byte, group *sync.WaitGroup) {
	group.Add(1)
	for row := range in {
		url := row.Cells[2].String()
		//log.Printf("Fetching URL %s\n", url)
		
		resp, _ := http.Get(url)
		
		buf := new(bytes.Buffer)
		buf.ReadFrom(resp.Body)
		
		out <- buf.Bytes()
		resp.Body.Close()
	}
	
	group.Done()
}

func Write(in chan []byte, eof chan bool) {

	xlsx := excelize.NewFile()
    // Create a new sheet.
	row := 1
	for img := range in {

		if err := ioutil.WriteFile("./tmp.jpg", img, 0644); err != nil {
			log.Printf("Error writing image: %v\n", err)
		}

		cell := fmt.Sprintf("A%d", row)
		if err := xlsx.AddPicture("Sheet1", cell, "./tmp.jpg", ""); err != nil {
        	fmt.Println(err)
		}	
		row++
		if row % 10 == 1 {
			fmt.Printf(".")
		}
		if row % 500 == 1 {
			fmt.Printf("\n")
			log.Printf("Saving %d rows\n", row)
			if err := xlsx.SaveAs("./output.xlsx"); err != nil {
				fmt.Printf("Error saving xls: %v\n", err)
			}
			xlsx = nil
			xlsx, err = excelize.OpenFile("./output.xlsx")
			if err != nil {
				fmt.Printf("Error reopening xlsx: %v\n", err)
			}
		}
	}
	if err := xlsx.Save(); err != nil {
		fmt.Printf("Error saving xls: %v\n", err)
	}	
	eof <- true
}
