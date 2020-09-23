Attribute VB_Name = "TrapMouse"

Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long

Public Sub DisableTrap(CurForm As Form)
Dim erg As Long
Dim NewRect As RECT
With NewRect
  .Left = 0&
  .Top = 0&
  .Right = Screen.Width / Screen.TwipsPerPixelX
  .Bottom = Screen.Height / Screen.TwipsPerPixelY
End With
erg& = ClipCursor(NewRect)
End Sub

Public Sub EnableTrap(CurForm As Form)
Dim X As Long, Y As Long, erg As Long
Dim NewRect As RECT
X& = Screen.TwipsPerPixelX
Y& = Screen.TwipsPerPixelY
'of the form
With NewRect
.Left = CurForm.Left / X&
.Top = CurForm.Top / Y&
.Right = .Left + CurForm.Width / X&
.Bottom = .Top + CurForm.Height / Y&
End With
erg& = ClipCursor(NewRect)
End Sub
