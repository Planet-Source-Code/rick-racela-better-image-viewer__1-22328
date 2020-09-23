Attribute VB_Name = "FormTrix"


Sub ClearAllTextBoxes(frm As Form)
' example of call       clearalltextboxes form1
Dim Control
    For Each Control In frm.Controls
    If TypeOf Control Is TextBox Then Control.Text = ""
    
    Next Control

End Sub
Function app_path() As String

Dim X
    X = App.Path
        If Right$(X, 1) <> "\" Then X = X + "\"
    app_path = UCase$(X)
   
End Function

 Sub centerform(frmIN As Form)
  
'       usage: CenterForm(FormName)
'           Just pass the form name and the form will
'           always be centered on the screen no matter
'           what the screen resolution.

Dim Itop As Integer, iLeft As Integer

If frmIN.WindowState <> 0 Then Exit Sub
Itop = (Screen.Height - frmIN.Height) / 2
iLeft = (Screen.Width - frmIN.Width) / 2
frmIN.Move iLeft, Itop

End Sub
