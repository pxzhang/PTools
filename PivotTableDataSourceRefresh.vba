Sub AdjustPivotDataRange()
'PURPOSE: Automatically readjust a Pivot Table's data source range
'SOURCE: www.TheSpreadsheetGuru.com/The-Code-Vault

Dim Data_sht As Worksheet
Dim Pivot_sht As Worksheet
Dim StartPoint As Range
Dim DataRange As Range
Dim PivotName As String
Dim NewRange As String
Dim pt As PivotTable

'Set Variables Equal to Data Sheet and Pivot Sheet
  Set Data_sht = ThisWorkbook.Worksheets("Data")
  Set Pivot_sht = ThisWorkbook.Worksheets("PivotSheet")
  
'Dynamically Retrieve Range Address of Data
  Set StartPoint = Data_sht.Range("B1")
  Set DataRange = Data_sht.Range(StartPoint, StartPoint.SpecialCells(xlLastCell))
  
  NewRange = Data_sht.Name & "!" & _
    DataRange.Address(ReferenceStyle:=xlR1C1)

'Make sure every column in data set has a heading and is not blank (error prevention)
  If WorksheetFunction.CountBlank(DataRange.Rows(1)) > 0 Then
    MsgBox "至少一个字段名为空，" & vbNewLine _
      & "请修复后重试!.", vbCritical, "Column Heading Missing!"
    Exit Sub
  End If
  
'Refresh pivot tables one by one.
  For Each pt In Pivot_sht.PivotTables
    pt.ChangePivotCache _
      ThisWorkbook.PivotCaches.Create( _
      SourceType:=xlDatabase, _
      SourceData:=NewRange)
    pt.RefreshTable
  Next pt
'Success alertness.
  MsgBox "数据更新成功!"

End Sub
