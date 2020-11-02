package hexcel

import (
	"bytes"
	"github.com/360EntSecGroup-Skylar/excelize"
	"reflect"
	"strconv"
)

// 5*26列 ，用于根据int序号获取Excel列坐标
var COL = []string{
	"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
	"AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ",
	"BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ",
	"CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ",
	"DA", "DB", "DC", "DD", "DE", "DF", "DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ"}

// Render 替换Excel模板中占位字符返回新的Excel文件   注意：模板中占位符不要有空格  ${xxx} 。 为保证数据格式显示一致，在创建模板时建议所有数据类型都用string
// @params tpl Excel模板文件
// @params sheet Sheet名称，默认：Sheet1
// @params kv 默认渲染参数，k为占位符，v为数据，数据类型支持字符串、数组（单列多行）、二维数组（多列多行）等
// 多行时会从占位行开始向后插入行，多列时不会插入列，只会按顺序设置值。 制作模板时注意排版。
func Render(tpl []byte, sheet string, kv map[string]interface{}) (fileBytes []byte, err error) {
	if sheet == "" {
		sheet = "Sheet1"
	}
	if kv == nil {
		return tpl, nil
	}
	f, err := excelize.OpenReader(bytes.NewReader(tpl))
	if err != nil {
		return
	}
	for k, v := range kv {
		xys := f.SearchSheet(sheet, "${"+k+"}") // 搜索占位符坐标
		if xys == nil {
			continue
		}
		for _, axis := range xys {
			switch reflect.TypeOf(v).Kind() {
			case reflect.Slice:
				// 为数组时，使用当前行后新增行的方式渲染多行数据
				setRowValue(f, sheet, axis, v)
			default:
				f.SetCellValue(sheet, axis, v)
			}
		}
	}
	buffer, err := f.WriteToBuffer()
	if err != nil {
		return
	}
	fileBytes = buffer.Bytes()
	return
}

// xyAxis 截取坐标中的行号
func xyAxis(axis string) (col string, row int) {
	row, err := strconv.Atoi(axis[1:])
	if err == nil {
		col = axis[:1]
		return
	}
	row, err = strconv.Atoi(axis[2:])
	if err == nil {
		col = axis[:2]
		return
	}
	row, err = strconv.Atoi(axis[3:])
	if err == nil {
		col = axis[:3]
		return
	}
	row, err = strconv.Atoi(axis[4:])
	if err == nil {
		col = axis[:4]
		return
	}
	panic(err)
}

// setRowValue 循环设置单元格行
// f Excel文件
// sheet Sheet名
// axis 开始位置单元格，从此单元格循环插入
// slice 行单元格数据数组
func setRowValue(f *excelize.File, sheet, axis string, slice interface{}) {
	values := reflect.ValueOf(slice)
	for i := 0; i < values.Len(); i++ {
		value := values.Index(i)
		col, row := xyAxis(axis)
		nAxis := col + strconv.Itoa(row+i)
		if i > 0 {
			f.DuplicateRow(sheet, row+i-1)
			f.SetRowHeight(sheet, row+i, f.GetRowHeight(sheet, row))
			f.SetCellStyle(sheet, nAxis, nAxis, f.GetCellStyle(sheet, axis))
		}
		setColSlice(f, sheet, nAxis, value.Interface())
	}
}

// setColSlice 循环横向设置单元格值
// slice 如果是数组，则横向循环设置值，否则直接设置当前单元格值
func setColSlice(f *excelize.File, sheet, axis string, slice interface{}) {
	// 非数组直接设置单元格值
	if reflect.TypeOf(slice).Kind() != reflect.Slice {
		f.SetCellValue(sheet, axis, slice)
		return
	}

	// 数组横向循环输出
	values := reflect.ValueOf(slice)

	col, row := xyAxis(axis)

	// 列转为数字下标
	nCol := 0
	for i := range COL {
		if COL[i] == col {
			nCol = i
			break
		}
	}
	nRow := strconv.Itoa(row)
	for i := 0; i < values.Len(); i++ {
		value := values.Index(i)
		nAxis := COL[i+nCol] + nRow
		f.SetCellValue(sheet, nAxis, value)
	}
}
