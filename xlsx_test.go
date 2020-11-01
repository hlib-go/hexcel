package hxlsx

import (
	"io/ioutil"
	"testing"
)

func TestTemplate(t *testing.T) {
	tpl, err := ioutil.ReadFile("./tpl.xlsx")
	if err != nil {
		t.Error(err)
		return
	}
	kv := make(map[string]interface{})
	kv["orgName"] = "测试机构"

	fileBytes, err := Template(tpl, "", kv)
	if err != nil {
		t.Error(err)
		return
	}
	ioutil.WriteFile("./test.xlsx", fileBytes, 0666)
}
