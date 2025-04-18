package main

import (
	"fmt"
	"html/template"
	"log"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

func main() {
	file, err := excelize.OpenFile("order1.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	rows, err := file.GetRows("Form Responses 1")
	if err != nil {
		log.Fatal(err)
	}

	type Count struct {
		TotalCount map[string]int
	}
	totalCount := make(map[string]int)
	Header := []string{}
	CompanyName := 0
	EmailId := 0
	Phonenumber := 0
	TimeStamp := 0
	var tc Count

	for i, row := range rows {
		// if i > 2 {
		// break
		// }
		fmt.Println(i)
		type Company struct {
			CompanyName string
			EmailId     string
			PhoneNumber string
			Items       map[string]string
		}
		var MyData Company
		MyData.Items = make(map[string]string)
		if i == 0 {
			for j, col := range row {
				Header = append(Header, col)
				if strings.Contains(col, "Name of the company") {
					CompanyName = j
				} else if strings.Contains(col, "Mobile Number") {
					Phonenumber = j
				} else if strings.Contains(col, "Email ID") {
					EmailId = j
				} else if strings.Contains(col, "Timestamp") {
					TimeStamp = j
				} else {
					continue
				}

				// fmt.Print(col, "\t")
			}

		} else {
			for k, col := range row {

				if k == CompanyName {
					fmt.Println(col)
					currentTime := time.Now()

					// Format the time using the reference time "Mon Jan 2 15:04:05 -0700 MST 2006"
					formattedTime := currentTime.Format("2006-01-02") // Example: YYYY-MM-DD HH:MM:SS

					MyData.CompanyName = col + "_" + formattedTime

				} else if k == Phonenumber {
					MyData.PhoneNumber = col

				} else if k == EmailId {
					MyData.EmailId = col
				} else if k == TimeStamp {
					continue
				} else {
					if col != "" {
						boxCount := strings.Split(col, " ")

						cnt, ok := totalCount[Header[k]]
						// fmt.Println(boxCount, ok)
						if ok {
							intCnt, err := strconv.Atoi(boxCount[0])
							if err == nil {
								totalCount[Header[k]] = cnt + intCnt

							}

						} else {
							intCnt, err := strconv.Atoi(boxCount[0])
							// fmt.Println(intCnt)
							if err == nil {
								totalCount[Header[k]] = intCnt
								// fmt.Println(totalCount[Header[k]])

							}

						}
						// fmt.Print(col, k, "\t")
						MyData.Items[Header[k]] = col
					}
				}
			}

			var tmplFile = "invoice.html"

			tmpl, err := template.New(tmplFile).ParseFiles(tmplFile)
			if err != nil {
				panic(err)
			}
			var f *os.File
			f, err = os.Create("./Customers/" + strings.TrimSpace(MyData.CompanyName) + ".html")
			if err != nil {
				panic(err)
			}
			err = tmpl.Execute(f, MyData)
			if err != nil {
				panic(err)
			}
			err = f.Close()
			if err != nil {
				panic(err)
			}

		}
	}
	tc.TotalCount = totalCount
	var tmplFilec = "count.html"
	tmplc, err := template.New(tmplFilec).ParseFiles(tmplFilec)
	if err != nil {
		panic(err)
	}
	var fc *os.File
	fc, err = os.Create("./Customers/" + "TotalCount.html")
	if err != nil {
		panic(err)
	}
	err = tmplc.Execute(fc, tc)
	if err != nil {
		panic(err)
	}
	err = fc.Close()
	if err != nil {
		panic(err)
	}
	// fmt.Println(totalCount)
}
