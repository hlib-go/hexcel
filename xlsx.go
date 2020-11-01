package hxlsx

import (
	"bytes"
	"github.com/360EntSecGroup-Skylar/excelize"
)

// Template 替换Excel模板中占位字符返回新的Excel文件   注意：模板中占位符不要有空格  ${xxx}
func Template(tpl []byte, sheet string, kv map[string]interface{}) (fileBytes []byte, err error) {
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
		xys := f.SearchSheet(sheet, "${"+k+"}")
		if xys == nil {
			continue
		}
		for _, xy := range xys {
			f.SetCellValue(sheet, xy, v)
		}
	}
	buffer, err := f.WriteToBuffer()
	if err != nil {
		return
	}
	fileBytes = buffer.Bytes()
	return
}
