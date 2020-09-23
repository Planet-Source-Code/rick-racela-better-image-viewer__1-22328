VERSION 5.00
Begin VB.Form ImageBrowser 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   5535
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Set Walllpaper"
      Height          =   255
      Left            =   5160
      TabIndex        =   13
      Top             =   4680
      Width           =   1215
   End
   Begin VB.PictureBox MeterBox 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   2760
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   237
      TabIndex        =   9
      Top             =   3600
      Width           =   3615
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   3615
      End
      Begin VB.Shape Meter 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   135
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   255
      Left            =   5160
      TabIndex        =   10
      Top             =   5040
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "All"
      Height          =   375
      Index           =   3
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bmp"
      Height          =   315
      Index           =   2
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4560
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Gif"
      Height          =   315
      Index           =   1
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Jpg"
      Height          =   315
      Index           =   0
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   615
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   120
      Pattern         =   "*.jpg;*.gif;*.bmp"
      TabIndex        =   2
      Top             =   2520
      Width           =   2475
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00FFFFFF&
      Height          =   1440
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2475
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2475
   End
   Begin VB.Shape Shape1 
      Height          =   5535
      Left            =   0
      Top             =   0
      Width           =   6615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      X1              =   120
      X2              =   6480
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5880
      TabIndex        =   14
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Display 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   4080
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dbl Click Image To Enlarge"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   3360
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Filters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Image WorkSpace 
      Height          =   1815
      Left            =   2880
      Top             =   360
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label lbl_filename 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Image Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4080
      TabIndex        =   3
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   5610
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6660
   End
End
Attribute VB_Name = "ImageBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'====================================
'
'   Used to change the systems Wallpaper
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
   (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
   
Const SPIF_UPDATEINIFILE = &H1
Const SPI_SETDESKWALLPAPER = 20
Const SPIF_SENDWININICHANGE = &H2
'
'====================================


Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Command2_Click()
    Dim l, a$
    
'   No matter what the extension type ( jpg, bmp, gif )
'   Save the file to the root directory with a .Bmp extension
    SavePicture WorkSpace.Picture, "c:\Wallpaper.bmp"
 
'   Set the Systems Wallpaper to the previously saved file
    a$ = "c:\Wallpaper.bmp"
    l = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, a$, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
End Sub

Private Sub Dir1_Change()
   File1.Path = Dir1.Path
   If File1.ListCount <= 0 Then
   Call ScalePicture(2500, 500, 500, Display, "")
   Exit Sub
   Else
File1.ListIndex = 0
End If

End Sub

Private Sub Display_DblClick()
DisableTrap ImageBrowser
ImageDisplay.Show
End Sub

Private Sub Drive1_Change()
On Error GoTo err
Dir1.Path = Drive1.Drive
lbl_filename.Caption = ""
Exit Sub
err:
    MsgBox "The device could not be opened.", vbCritical, "Error"
    Exit Sub
End Sub

Private Sub File1_Click()

If File1.ListIndex = -1 Then Exit Sub Else _
lbl_filename.Caption = File1.FileName
If Right$(File1.Path, 1) <> "\" Then
picpath = File1.Path & "\" & File1.FileName
Else
picpath = File1.Path & lbl_filename.Caption
End If
r = GetFileSize(picpath)

If totalBytes >= "1048576" Then totalBytes = totalBytes / 50    ' make the delay a little shorter
' display and update progress bar
   SetMeter 1, ImageBrowser
For i = 1 To 100
Label3.Caption = ""
For X = 1 To totalBytes
Next X
   SetMeter -1, ImageBrowser
    Next

Call ScalePicture(2500, 3500, 500, Display, picpath)
For X = 1 To totalBytes
Next X
r = round(totalBytes)
Label3.Caption = "File size " & r & " Kbs"
SetMeter 0, ImageBrowser


End Sub

Sub ScalePicture(ImageSize As Integer, Column As Integer, Row As Integer, ctl As Image, picpath As String)
    Dim ScaleRate As Double
    Display.Visible = False
    ' Loads picture into WorkSpace ImageBox
    WorkSpace.Picture = LoadPicture(picpath)
    
    ' When image is smaller than the ImageSize its ok
        If ImageSize > WorkSpace.Height And ImageSize > WorkSpace.Width Then
            Display.Stretch = False
            Display = WorkSpace
        Else
        
        '   If image is larger than the display, we need to calculate
        '   its aspect ratio (ScaleRate) and set the displays
        '   Stretch property to True

        If WorkSpace.Width > WorkSpace.Height Then
            ScaleRate = Abs(WorkSpace.Width / WorkSpace.Height)
                Display.Width = ImageSize
                Display.Height = ImageSize / ScaleRate
        Else
            ScaleRate = Abs(WorkSpace.Height / WorkSpace.Width)
                Display.Height = ImageSize
                Display.Width = ImageSize / ScaleRate
        End If
                Display.Stretch = True
        End If
    
    ' Assign the picture to Dispaly's ImageBox
    
        Display.BorderStyle = 0
        Display.Left = Column + (ImageSize - ctl.Width) / 2
        Display.Top = Row + (ImageSize - ctl.Height) / 2
        Display = WorkSpace
        Display.Visible = True
End Sub

Private Sub Form_Load()
On Error GoTo Error
Image1.Picture = LoadPicture(app_path & "strand.jpg")
Dir1.Path = "c:\My documents"
If File1.ListIndex < -1 Then Exit Sub Else _
File1.ListIndex = 0

LoadMeter Me, RGB(0, 0, 128)

Error:


    Select Case err.Number
        Case 0
        ' No Errors Found
            Resume Next

    Case 76
    ' Path Not Found
        MsgBox "Error " & err.Number & "     " & err.Description & "     C:\Graphics", vbCritical, "Error"
        MsgBox "Defaulting to   " & Chr$(34) & "ROOT Directory" & Chr$(34), vbCritical
        Dir1.Path = "c:\"
            Resume Next

    End Select
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
EnableTrap ImageBrowser
End Sub

Private Sub Form_Unload(cancel As Integer)

DisableTrap ImageBrowser
Unload Me
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
EnableTrap ImageBrowser
Label4.ForeColor = &HFF0000
End Sub

Private Sub Label4_Click()
DisableTrap ImageBrowser
Label4.ForeColor = &HFF0000
About.Show
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFF&
End Sub

Private Sub Option1_Click(Index As Integer)

    Select Case Index
        Case 0
            File1.Pattern = "*.jpg"
        Case 1
            File1.Pattern = "*.gif"
        Case 2
            File1.Pattern = "*.bmp"
        Case 3
            File1.Pattern = "*.jpg;*.gif;*.bmp"
    End Select
    
End Sub
