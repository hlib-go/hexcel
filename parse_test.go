package hexcel

import (
	"bytes"
	"encoding/json"
	"io/ioutil"
	"testing"
)

func TestGetRows(t *testing.T) {
	tpl, err := ioutil.ReadFile("./txls/import-template.xlsx")
	if err != nil {
		t.Error(err)
		return
	}
	rows, err := GetRows(bytes.NewReader(tpl))
	if err != nil {
		t.Error(err)
		return
	}
	rbytes, _ := json.Marshal(rows)
	t.Log(string(rbytes))
}
