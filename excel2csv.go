package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"github.com/tealeg/xlsx"
	"io"
	"log"
	"os"
	"strings"
)

type csvOptSetter func(*csv.Writer)

// Структура параметров эспорта
type excelData struct {
	FileName  string    // Имя файл xlsx
	SheetName string    // Название листа Excel
	Start     string    // Адрес стартовой ячейки листа
	End       string    // Адрес конечной ячейки листа
}

func main() {
	var (
		outFile   = flag.String("o", "-", "имя CSV-файла, если не будет указано, содержимое будет выведено только на экран")
		sheetName = flag.String("s", "", "Имя листа для экспорта")
		diapazon  = flag.String("d", "", "Диапазон ячеек (A1:A2)")
		delimiter = flag.String("r", "\t", "Разделитель ячеек в строке")
	)
	flag.Usage = func() {
		fmt.Fprintf(os.Stderr, `
%s переводит файл Excel (XLSX) в CSV-файл согласно указанным параметрам.
Использование:
	%s [ключи] <file.xlsx>
`, "excel2csv", "excel2csv")
		flag.PrintDefaults()
	}

	flag.Parse()
	if flag.NArg() != 1 {
		flag.Usage()
		os.Exit(1)
	}
	out := os.Stdout
	if !(*outFile == "" || *outFile == "-") {
		var err error
		if out, err = os.Create(*outFile); err != nil {
			log.Fatal(err)
		}
	}
	defer func() {
		if closeErr := out.Close(); closeErr != nil {
			log.Fatal(closeErr)
		}
	}()

	getParts := strings.Split(*diapazon, ":")
	var start, end string
	if len(getParts) > 1 {
		start = getParts[0]
		end = getParts[1]
	}

	eData := excelData{
		FileName:  flag.Arg(0),
		SheetName: *sheetName,
		Start:     start,
		End:       end,
	}

	csvOpts := func(cw *csv.Writer) { cw.Comma = ([]rune(*delimiter))[0] }

	err := xlsx2csv(out, eData, csvOpts)
	if err != nil {
		log.Fatal(err)
	}
}

// Обрабатываем файл XLSX
// w — объект, реализующий io.Writer, например, файл или стандартный выход
// eData — данные для экспорта
// csvOpts — параметры записи CSV
func xlsx2csv(w io.Writer, eData excelData, csvOpts csvOptSetter) error {
	xlFile, err := xlsx.OpenFile(eData.FileName)
	if err != nil {
		return err
	}

	if eData.SheetName == "" {
		err := exportSheet(xlFile.Sheets[0], w, eData, csvOpts)
		if err != nil {
			log.Println(err)
		}
	} else {
		for i, sheet := range xlFile.Sheets {
			if sheet.Name == strings.TrimSpace(eData.SheetName) {
				err := exportSheet(xlFile.Sheets[i], w, eData, csvOpts)
				if err != nil {
					log.Println(err)
				}
			}
		}
	}
	return nil
}

func exportSheet(sheet *xlsx.Sheet, w io.Writer, eData excelData, csvOpts csvOptSetter) error {
	cw := csv.NewWriter(w)
	if csvOpts != nil {
		csvOpts(cw)
	}

	var vals []string
	if eData.Start == "" || eData.End == "" {
		for _, row := range sheet.Rows {
			if row != nil {
				vals = vals[:0]
				for _, cell := range row.Cells {
					str, err := cell.FormattedValue()
					if err != nil {
						vals = append(vals, err.Error())
					}
					vals = append(vals, fmt.Sprintf("%s", str))
				}
			}
			cw.Write(vals)
		}
	} else {
		startCol, startRow, err := xlsx.GetCoordsFromCellIDString(eData.Start)
		if err != nil {
			return err
		}
		endCol, endRow, err := xlsx.GetCoordsFromCellIDString(eData.End)
		if err != nil {
			return err
		}
		for ri, row := range sheet.Rows {
			if row != nil && ri >= startRow && ri <= endRow {
				vals = vals[:0]
				for ci, cell := range row.Cells {
					if ci >= startCol && ci <= endCol {
						str := cell.Value
						//if err != nil {
						//	vals = append(vals, err.Error())
						//}
						vals = append(vals, fmt.Sprintf("%s", str))
					}
				}
				cw.Write(vals)
			}
		}
	}
	cw.Flush()
	return cw.Error()
}
