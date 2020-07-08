### 91.6 解決辦法：使用SaveAs方法保存.xlsx後，再次打開提示: 文件損壞,後綴名錯誤（格式錯誤）

**問題描述：** 舊宏文件裏使用`SaveAs`方法，保存爲.xls文件，當改成保存爲.xlsx文件後，再次打開保存後的文件時，提示文件名後綴錯誤，無法打開宏文件生成的文件。

**解決辦法：** 修改`SaveAs`方法中的參數，將FileFormat的參數設置成如下：

```vb
FileFormat:=xlOpenXMLWorkbook
```
