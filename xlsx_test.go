package hxlsx

import (
	"io/ioutil"
	"strconv"
	"testing"
)

func TestRender(t *testing.T) {
	tpl, err := ioutil.ReadFile("./txls/tpl.xlsx")
	if err != nil {
		t.Error(err)
		return
	}
	kv := make(map[string]interface{})
	kv["orgName"] = "测试机构"
	kv["codes"] = [][]string{{"1234567890", "123456"}, {"1234567891", "123452"}}
	kv["num"] = 3

	fileBytes, err := Render(tpl, "", kv)
	if err != nil {
		t.Error(err)
		return
	}
	ioutil.WriteFile("./txls/test.xlsx", fileBytes, 0666)
	t.Log("success......")
}

func TestB(t *testing.T) {
	i, err := strconv.ParseInt("a", 64, 10)
	if err != nil {
		t.Error(err)
		return
	}
	t.Log(i)
}
