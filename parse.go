package hexcel

import (
	"github.com/360EntSecGroup-Skylar/excelize"
	"io"
)

// 解析Excel返回二维数组
func GetRows(r io.Reader) (rows [][]string, err error) {
	file, err := excelize.OpenReader(r)
	if err != nil {
		return
	}
	sheet := file.GetSheetName(1)
	rows = file.GetRows(sheet)
	return
}
