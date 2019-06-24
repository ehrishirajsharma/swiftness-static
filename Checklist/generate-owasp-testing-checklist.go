// Fetch and Parse from https://github.com/tanprathan/OWASP-Testing-Checklist

package main

import (
	"encoding/json"
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"net/http"
	"os"
	"strings"

	"github.com/google/uuid"
	"github.com/tealeg/xlsx"
)

type xlsxFolder struct {
	ID     string      `json:"id"`
	Title  string      `json:"title"`
	Checks []xlsxCheck `json:"checklist"`
}

type xlsxCheck struct {
	ID      string `json:"id"`
	Title   string `json:"title"`
	Content string `json:"content"`
}

func main() {
	excelFileName, err := ioutil.TempFile("", "excel-otg.xlsx")
	if err != nil {
		log.Fatal(err)
	}
	name := excelFileName.Name()
	defer os.Remove(name)

	resp, err := http.Get("https://github.com/tanprathan/OWASP-Testing-Checklist/blob/master/OWASPv4_Checklist.xlsx?raw=true")
	if err != nil {
		log.Fatal(err)
	}
	defer resp.Body.Close()

	io.Copy(excelFileName, resp.Body)
	excelFileName.Close()

	xlFile, err := xlsx.OpenFile(name)
	if err != nil {
		log.Fatal(err)
	}
	sheet := xlFile.Sheets[0]

	var item *xlsxFolder
	list := []*xlsxFolder{}

	for i := 0; i < len(sheet.Rows); i++ {
		// If the second cell exists and contains "Test Name" then it's an header
		if len(sheet.Rows[i].Cells) > 2 {
			if strings.EqualFold(sheet.Rows[i].Cells[1].String(), "Test Name") {
				if item != nil {
					list = append(list, item)
				}
				item = &xlsxFolder{ID: uuid.New().String(), Title: sheet.Rows[i].Cells[0].String()}
			}
		}

		if len(sheet.Rows[i].Cells) > 4 {
			if strings.EqualFold(sheet.Rows[i].Cells[4].String(), "Not Started") {
				var title string
				if sheet.Rows[i].Cells[0].String() == "" {
					title = fmt.Sprintf("%s", sheet.Rows[i].Cells[1].String())
				} else {
					title = fmt.Sprintf("%s [%s]", sheet.Rows[i].Cells[1].String(), sheet.Rows[i].Cells[0].String())
				}

				content := strings.Replace(sheet.Rows[i].Cells[2].String(), "\n", "<br>", -1)
				item.Checks = append(item.Checks, xlsxCheck{ID: uuid.New().String(), Title: title, Content: fmt.Sprintf("<p>%s</p>", content)})
			}
		}
	}

	data, err := json.Marshal(&list)
	if err != nil {
		log.Fatal(err)
	}

	final := fmt.Sprintf("{\"targets\":[],\"libraries\":[{\"id\":\"%s\",\"title\": \"OWASP Testing Checklist\",\"folders\":%s}],\"templates\":[],\"payloads\":[],\"messages\": {\"showDeleteConfirmation\": true}}", uuid.New().String(), string(data))
	fmt.Println(final)
