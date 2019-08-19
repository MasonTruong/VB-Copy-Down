Sub CopyDownToEmpty

Dim rangeSelect As Range
Dim destCol AS Range
Dim getFirstNum
Dim getNextNum
Dim responseError

'ask for cel range selection
Set rangeSelect = Nothing
On Error Resume Next
Set rangeSelect = Application.InputBox(prompt:="Range", Title:="Column Copy", Type:=8)
On Error Goto 0
'catch error msg on the instance cancel is selected
    If rangeSelect Is Nothing Then
        responseError = MsgBox(prompt:="No range selected!", Buttons:=vbCritical, Title:="Range Error!")
        Exit Sub
    End If

'loop through selected range. Retrieves filled cells and copies that cell value
'to the next empty cell down. Stops at the end of selected range
For Each rw In rangeSelect
    getFirstNum = rw.value
    getNextNum = rw.Offset(rowOffset:=1).Value

    If getNextNum = Empty Then
        rw.Offset(rowOffset:=1) = getFirstNum
    End If
Next rw

End Sub