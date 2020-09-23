Attribute VB_Name = "Module1"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Sub Movement(Obj As Object, Tl As Integer)
  If Tl = 1 Then
    ReleaseCapture
       SendMessage Obj.hwnd, &HA1, 2, 0
  End If
End Sub

