Attribute VB_Name = "DEV"
Option Explicit
Private Sub test()
    Dim scriptEngine As New ASF
    Dim progIdx As Long
    Dim tmpResult As Variant
    With scriptEngine
        .AppAccess = True
        .OverrideCollMethods = True
        .verbose = True
'        .EnableCallTrace = True
        progIdx = .Compile("  prototype.COM.ListRow asDictionary() {" & _
                            "    let headers = this.parent.listcolumns;" & _
                            "    let values = this.range.value2;" & _
                            "    let result = {};" & _
                            "    for (let i = 1, i <= headers.count, i+=1) {" & _
                            "        result.set(headers.item(i).name, values[1][i]);" & _
                            "    };" & _
                            "    return result;" & _
                            " };" & _
                            " let myData = $1.Sheets(1).ListObjects('Demo_Table').ListRows.map(" & _
                            "     fun(row) {" & _
                            "         return row.asDictionary().Get('ip_address');" & _
                            "      }" & _
                            " ); print(myData[5])")
        .Run progIdx, ThisWorkbook
'        Debug.Print .GetCallStackTrace
    End With
End Sub
