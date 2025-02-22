Function GetObjectSafe(path)
  Dim obj, errNum
  On Error Resume Next
  Set obj = GetObject(path)
  errNum = Err.Number
  On Error GoTo 0
  If errNum <> 0 Then
    'Handle specific error codes if needed
    Select Case errNum
      Case 429:
        MsgBox "The system cannot find the file specified."
      Case Else:
        MsgBox "An error occurred: " & errNum
    End Select
    Set obj = Nothing
  End If
  Set GetObjectSafe = obj
End Function

'Example Usage:
Set obj = GetObjectSafe("C:\\Windows\\System32\\notepad.exe")
if not obj is nothing then
 msgbox "Object found"
else
 msgbox "Object not found"
end if