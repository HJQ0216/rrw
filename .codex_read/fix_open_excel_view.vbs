Option Explicit

Dim excelApp, workbook, worksheet, window, targetPath, found
If WScript.Arguments.Count < 1 Then
  WScript.Echo "ERROR: missing target path"
  WScript.Quit 1
End If
targetPath = WScript.Arguments(0)
found = False

On Error Resume Next
Set excelApp = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
  WScript.Echo "ERROR: no running Excel instance"
  WScript.Quit 2
End If
On Error GoTo 0

For Each workbook In excelApp.Workbooks
  If LCase(workbook.FullName) = LCase(targetPath) Then
    found = True
    Set worksheet = workbook.Worksheets(1)
    worksheet.Activate
    Set window = workbook.Windows(1)
    window.FreezePanes = False
    window.SplitColumn = 2
    window.SplitRow = 7
    window.FreezePanes = True
    worksheet.Range("C8").Select
    window.ScrollColumn = 3
    window.ScrollRow = 8
    workbook.Save
    WScript.Echo "FIXED: " & workbook.FullName
    Exit For
  End If
Next

If Not found Then
  WScript.Echo "ERROR: target workbook is not open"
  WScript.Quit 3
End If
