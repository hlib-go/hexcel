# hxlsx

#### 介绍

Excel 表格

#### 使用

1. 渲染Excel模板文件
```go
hxlsx.Template(tpl []byte, sheet string, kv map[string]interface{}) (fileBytes []byte, err error) 
```
