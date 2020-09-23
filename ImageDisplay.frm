VERSION 5.00
Begin VB.Form ImageDisplay 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image WorkSpace 
      Height          =   1815
      Left            =   2760
      Top             =   2640
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Image Display 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "ImageDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub ScalePicture(ImageSize As Integer, Column As Integer, Row As Integer, Display As Image, picpath As String)
    Dim ScaleRate As Double
    ' Loads picture into WorkSpace ImageBox
    WorkSpace.Picture = LoadPicture(picpath)
    ' When image is smaller than the ImageSize its ok

        If WorkSpace.Height < ImageSize Then
            ScaleRate = Abs(WorkSpace.Height / WorkSpace.Width)
                WorkSpace.Height = ImageSize
                WorkSpace.Width = ImageSize / ScaleRate
        End If
            
        If WorkSpace.Width < ImageSize Then
            ScaleRate = Abs(WorkSpace.Width / WorkSpace.Height)
                WorkSpace.Width = ImageSize
                WorkSpace.Height = ImageSize / ScaleRate
        End If

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
    
    ' Assign the picture to destination ImageBox
    
        Display.BorderStyle = 1
        Display.Left = Column
        Display.Top = Row
        Display = WorkSpace
        ImageDisplay.Width = Display.Width
        ImageDisplay.Height = Display.Height
    
    centerform Me
    EnableTrap Me
    Display.Visible = True
End Sub


Private Sub Display_Click()
'   Enlarged picture no longer needed, reclaim memory space
        Set Display.Picture = Nothing
        
DisableTrap ImageBrowser
Unload Me
End Sub


Private Sub Form_Load()

'
    WorkSpace.Picture = LoadPicture("")
    Display.Picture = LoadPicture("")

    centerform Me
    Call ScalePicture(8000, 0, 0, Display, picpath)
    
End Sub

