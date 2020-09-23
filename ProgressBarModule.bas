Attribute VB_Name = "ProgressBarModule"
Public MeterValue As Single

Public Function LoadMeter(CurForm As Form, BarColor As Long)
On Error GoTo Error
'MeterPos
'CurForm.MeterPos.Caption = "0%"
'CurForm.MeterPos.Left = 0
'CurForm.MeterPos.Width = CurForm.MeterBox.ScaleWidth
'Meter
CurForm.Meter.Left = -1
CurForm.Meter.Width = 0
CurForm.Meter.Height = CurForm.MeterBox.ScaleHeight
CurForm.Meter.BackColor = BarColor
CurForm.Meter.BorderColor = BarColor
Exit Function
Error:
MsgBox err.Description
End Function

Public Function SetMeter(SetNum As Single, CurForm As Form)
On Error GoTo Error
Dim MeterNum As Single
'Checks if SetNum for Certain Values
SetNum = round(SetNum)
Select Case SetNum
Case Is <= -3
    SetNum = 0
Case -1
    If CurForm.Meter.Width = 1 Then
        SetNum = round((CurForm.Meter.Width / (CurForm.MeterBox.ScaleWidth / 100)))
    Else
        SetNum = round((CurForm.Meter.Width / (CurForm.MeterBox.ScaleWidth / 100))) + 1
    End If
    If SetNum >= 100 Then
        SetNum = 100
        End If
Case -2
    SetNum = round((CurForm.Meter.Width / (CurForm.MeterBox.ScaleWidth / 100))) - 1
    If SetNum <= 0 Then
        SetNum = 0
        End If
Case Is > 100
    Exit Function
End Select

'Sets the width
MeterNum = round(CurForm.MeterBox.ScaleWidth / 100)
CurForm.Meter.Width = round(MeterNum) * round(SetNum)
'CurForm.MeterPos.Caption = round(SetNum) & "%"
Exit Function
Error:
MsgBox err.Description
End Function

Public Function GetMeter(CurForm As Form)
If CurForm.Meter.Width = 1 Then
    MeterValue = round((CurForm.Meter.Width / (CurForm.MeterBox.ScaleWidth / 100))) - 1
Else
    MeterValue = round((CurForm.Meter.Width / (CurForm.MeterBox.ScaleWidth / 100)))
End If
End Function
Public Function round(X)

    round = (Int(100 * (X + 0.005))) / 100

End Function
