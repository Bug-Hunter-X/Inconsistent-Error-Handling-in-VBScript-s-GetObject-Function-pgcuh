Function GetObject(path)
  On Error Resume Next
  Set obj = GetObject(path)
  If Err.Number <> 0 Then
    Err.Clear
    Set obj = Nothing
  End If
  Set GetObject = obj
End Function

'Example Usage:
Set obj = GetObject("C:\\Windows\\System32\\notepad.exe")
if not obj is nothing then
 msgbox "Object found"
else
 msgbox "Object not found"
end if