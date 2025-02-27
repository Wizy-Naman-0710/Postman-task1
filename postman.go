package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"math"
	"regexp"
	"strconv"
	"os"
	"time"
	"github.com/xuri/excelize/v2"
)

func round(num float64) int {
    return int(num + math.Copysign(0.5, num))
}

func toFixed(num float64, precision int) float64 {
    output := math.Pow(10, float64(precision))
    return float64(round(num * output)) / output
}

type Ranker struct {
	Emplid 	string	`json:"emplid"`
	Marks 	float64	`json:"marks"`
}

type TotalError struct {
    ErrorIndex     string  `json:"error_index"`
    ExpectedTotal  float64 `json:"expected_total"`
    ActualTotal    float64 `json:"actual_total"`
}

func main() {
	start := time.Now()
	file_name := "data.xlsx"
	sheet_name := "CSF111_202425_01_GradeBook"

	f, err := excelize.OpenFile(file_name)
	if err != nil {
		fmt.Println(err)
		return
	}

	defer func() {
		elapsed := time.Since(start)
		fmt.Printf("Task took %s\n", elapsed)

		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()


	rows, err := f.GetRows(sheet_name)
	if err != nil {
		fmt.Println(err)
		return
	}

	// intializing flags for go CLI 
	exportToJSON := flag.String("export", "", "You can type json to export the content in json format")
	roomNum := flag.Int("room", -1, "You can get the data for specific room numbers")

	flag.Parse()
	

	// declaring all the Variables 
	exam_title := []string{"Quiz", "Midsem", "Lab Test", "Weekly Lab", "Pre-compre", "Compre", "Total"}
	var first, second, third Ranker
	averages := make(map[string]float64)
	count := 0
	errorCount := 0
	branchwiseAverage := make(map[string]float64)
	branchwiseStudents := make(map[string]int)
	rankerList := make(map[string]Ranker)
	errorList := make(map[int]TotalError)
	
	for _, row := range rows[1:] {

		currentRoomNum, _ := strconv.Atoi(row[1])
		if *roomNum != -1 && currentRoomNum != *roomNum {
			continue
		}


		// calculating pre compre marks and averages 
		var pre_compre float64 = 0
		for idx, col := range row[4:8] {
			marks, _ := strconv.ParseFloat(col, 64)
			averages[exam_title[idx]] += marks
			pre_compre += marks
		}

		compre_in_excel, _ := strconv.ParseFloat(row[9], 64)
		total_in_excel, _ := strconv.ParseFloat(row[10], 64)
		totalCalculated := toFixed((pre_compre+compre_in_excel), 3)

		averages[exam_title[4]] += pre_compre
		averages[exam_title[5]] += compre_in_excel
		averages[exam_title[6]] += totalCalculated

		// reporting discrepancy in the report and fixing it
		if totalCalculated != total_in_excel {
			row[10] = strconv.FormatFloat(totalCalculated, 'f', 2, 64)
			currentTotalError := TotalError{ErrorIndex: row[0], ExpectedTotal: totalCalculated, ActualTotal: total_in_excel}
			errorCount++
			errorList[errorCount] = currentTotalError
		}

		// branch-wise total average
		re := regexp.MustCompile(`[A-Z]\d`)
		campus_id := re.FindAllString(row[3], -1)

		for _, id := range campus_id {
			_, prs := branchwiseAverage[id]
			if prs {
				branchwiseAverage[id] += totalCalculated
				branchwiseStudents[id]++
			} else {
				branchwiseAverage[id] = totalCalculated
				branchwiseStudents[id] = 1
			}
		}
		count++

		// ranking 1st, 2nd and 3rd O(n)
		if totalCalculated > first.Marks {
			first.Emplid = row[2]
			first.Marks = totalCalculated
		}else if totalCalculated > second.Marks {
			second.Emplid = row[2]
			second.Marks = totalCalculated
		} else if totalCalculated > third.Marks {
			third.Emplid = row[2]
			third.Marks = totalCalculated
		}

	}

	for key, value := range averages {
		averages[key] = toFixed(value/float64(count), 3)
	}

	for key, value := range branchwiseAverage {
		branchwiseAverage[key] = toFixed(value / float64(branchwiseStudents[key]), 2)
	}

	rankerList["1st"] = first
	rankerList["2nd"] = second
	rankerList["3rd"] = third

	
	//Printing the summary of the data
	fmt.Printf("The discrepancy found in the table: \n")
	for _, value := range errorList {
		fmt.Printf("There is a error at Serial No. %s \t Expected Total Marks: %f \t Given: %f \n", value.ErrorIndex, value.ExpectedTotal, value.ActualTotal)
	}
	fmt.Println()
	
	fmt.Printf("The averages of each components: \n")
	for key, value := range averages {
		fmt.Println(key, value)
	}
	fmt.Println()

	fmt.Printf("The Total Averages Batch-wise: \n")
	for key, value := range branchwiseAverage {
		fmt.Println(key, value)
	}
	fmt.Println()

	fmt.Printf("Class topper are : \n")
	fmt.Printf("1st \t Emplid %s \t Marks %f \n", first.Emplid, first.Marks)
	fmt.Printf("2nd \t Emplid %s \t Marks %f \n", second.Emplid, second.Marks)
	fmt.Printf("3rd \t Emplid %s \t Marks %f \n", third.Emplid, third.Marks)
	fmt.Println()


	// exporting the data is JSON Format
	if *exportToJSON == "json" {
		data := map[string]interface{}{
			"errors": 				errorList, 
			"average": 				averages, 
			"branchwise-average": 	branchwiseAverage, 
			"rankers": 				rankerList,
		}


		jsonData, _ := json.MarshalIndent(data, "", "	")
		err = os.WriteFile("report.json", jsonData, 0644)
		if err != nil {
			fmt.Println("Error writing JSON to file:", err)
		} else {
			fmt.Println("Data successfully exported to report.json")
		}
	}
}
