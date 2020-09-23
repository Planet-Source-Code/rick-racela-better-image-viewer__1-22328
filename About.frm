VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   3135
      Left            =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Unload Me
DisableTrap Me
EnableTrap ImageBrowser
End Sub

Private Sub Form_Load()
centerform Me
EnableTrap Me
Image1.Stretch = True
Image1.Picture = LoadPicture(app_path & "bestwater.jpg")

Label1(0).Caption = App.ProductName
Label1(1).Caption = "Version " & App.Major & "." & App.Minor & "   Beta  " & App.Revision
Label1(2).Caption = App.Comments
Label1(3).Caption = "Copyright " & App.CompanyName
End Sub
