package main

import (
	"flag"
	"fmt"
	"os"
	"github.com/tealeg/xlsx"
	"strings"
	"encoding/json"
)

func main() {
	// Get open file path & fields to use
	filePath, inputFields := getArgs()

	if filePath == "" || inputFields == "" {
		fmt.Println("Syntax: excel_to_json --file<EXCEL FILE> --fields=<FIELDS TO EXTRACT>\n Example: excel_to_json --file=test.xlsx --fields=id,name,timestamp")
		os.Exit(1)
	}

	fields := getFields(inputFields)

	// Read the excel file
	xlFile, err := xlsx.OpenFile(filePath)
	if err != nil {
		fmt.Println(err)
		os.Exit(0)
	}

	// Open the first sheet
	sheet := xlFile.Sheets[0]

	// Read the first line of the sheet to get the headers
	headers := getHeaders(sheet)

	// Create a slice with the headers as indexes
	indices := getHeaderIndices(fields, headers)
	var jsonArray []map[string]string

	// Loop through all the rows
	for _, row := range sheet.Rows {
		// Skip the first row, contains labels
		if row == sheet.Rows[0] {
			continue
		}
		rowMap := make(map[string]string)

		// Loop through all cells
		for key, cell := range row.Cells {

			// If the current cell matches our fields argument
			if inIntSlice(key, indices) {
				cellString, _ := cell.String()
				header := getHeader(headers, key)
				rowMap[header] = cellString

			}

		}

		// If there are items that matched
		if rowMap != nil {
			// Store in slice of rows
			jsonArray = append(jsonArray, rowMap)
		}
	}

	// Get JSON string from slice
	jsonAsString := mapToReadableJson(jsonArray)

	// Write the string to a file
	writeFilePath := filePath + ".json"
	file, err := os.Create(writeFilePath)
	if err != nil {
		fmt.Println(err)
		os.Exit(0)
	}

	file.WriteString(jsonAsString)
	file.Sync()
	file.Close()

}

/**
 * return a slice with indexes for the headers
 * @param fields	[]string	The fields we accept
 * @param headers	[]string	The headers of the excel file
 * @returns []int			The slice with the indexes
 */
func getHeaderIndices(fields []string, headers []string) []int {
	var indices []int
	for key, header := range headers {
		if inSlice(header, fields) {
			indices = append(indices, key)
		}
	}
	return indices
}

/**
 * Get headers from excel file
 * @param sheet	*xlsx.Sheet	The excel sheet
 * @returns []string		A slice with all the headers
 */
func getHeaders(sheet *xlsx.Sheet) []string {
	var headers []string
	for _, cell := range sheet.Rows[0].Cells {
		text, _ := cell.String()
		headers = append(headers, text)
	}

	return headers
}

/**
 * Get all filepath & fields command line arguments
 * @returns string, string	The arguments
 */
func getArgs() (string, string) {
	var filePath = flag.String("file", "", "Bestand dat je wilt openen")
	var fields = flag.String("fields", "", "De velden die je wilt hebben")
	flag.Parse()
	return *filePath, *fields
}

/**
 * Parse the --fields argument
 * @returns	[]string	A slice of all fields
 */
func getFields(fields string) []string {
	return strings.Split(fields, ",")
}

/**
 * Check if a value exists in a slice
 * @param	needle		int	What to search for
 * @param	haystack	int[]	The slice to search in
 * @returns	bool			Whether the value exists or not
 */
func inIntSlice(needle int, haystack []int) bool {
	for _, value := range haystack {
		if value == needle {
			return true
		}
	}

	return false
}

/**
 * Check if a value exists in a slice
 * @param	needle		string	What to search for
 * @param	haystack	string[]	The slice to search in
 * @returns	bool			Whether the value exists or not
 */
func inSlice(needle string, haystack []string) bool {
	for _, value := range haystack {
		if value == needle {
			return true
		}
	}

	return false
}

/**
 * Get header from the slice of arrays by key
 * @param	headers	[]string	The  slice of headers to search in
 * @param	index	int		Index of the header
 * @returns	string			The header
 */
func getHeader(headers []string, index int) string {
	return headers[index]
}

/**
 * Turn our slice of maps with data into a readable json string
 * @param mapArg	[]map[string]string	The map to convert
 * @returns		string			The indented json string
 */
func mapToReadableJson(mapArg []map[string]string) string {
	jsonMarshalled, _ := json.MarshalIndent(mapArg, "", "    ")
	jsonAsString := string(jsonMarshalled)
	return jsonAsString
}
