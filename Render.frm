VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "3d Render"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9240
   Icon            =   "Render.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5340
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3705
      Width           =   3435
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   210
      Left            =   4080
      Max             =   10
      Min             =   1
      TabIndex        =   7
      Top             =   540
      Value           =   5
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   5500
      TabIndex        =   6
      Top             =   480
      Width           =   3000
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   5500
      TabIndex        =   5
      Top             =   90
      Width           =   3000
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   5500
      Pattern         =   "*.bmp;*.jpg*"
      TabIndex        =   4
      Top             =   2150
      Width           =   3000
   End
   Begin VB.CommandButton cmdRender 
      Caption         =   "&Render"
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   1095
      Width           =   900
   End
   Begin VB.PictureBox Render1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      DrawWidth       =   3
      FillStyle       =   0  'Solid
      Height          =   4770
      Left            =   0
      ScaleHeight     =   318
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   613
      TabIndex        =   2
      Top             =   4035
      Visible         =   0   'False
      Width           =   9200
   End
   Begin VB.PictureBox Picmap 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      DrawWidth       =   3
      FillStyle       =   0  'Solid
      Height          =   3840
      Left            =   -15
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   1
      Top             =   0
      Width           =   3840
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save as"
      Height          =   255
      Left            =   4200
      Picture         =   "Render.frx":030A
      TabIndex        =   0
      Top             =   3705
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "< Height ratio >"
      Height          =   315
      Index           =   0
      Left            =   4080
      TabIndex        =   8
      Top             =   150
      Width           =   1260
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
'*************************

Static Function Log10(X1)
    Log10 = Log(X1) / Log(10#) ' Log function to work out depth of vision
End Function

Private Sub cmdRender_Click()
startTime = 0
Render1.Visible = True
Render1.Cls
MyISO = 0
Me.AutoRedraw = True

For GenC = 1 To 3839 Step UserDetail
   For GenC2 = 1 To 3072 Step UserDetail
      ColInfo(GenC) = Picmap.Point(GenC / UserDetail, (GenC2) / UserDetail) ' get colour of original image
      MyHeight = ColInfo(GenC) And 255 ' get height information from colour of pixel
      MyISO = ((Log10(GenC2 / UserDetail)) * 1.2) 'adjust for depth of vision, further away = darker, nearer = lighter
      MyHeight = MyHeight * (0.5 + (MyISO * (2.5 * UserHeight))) ' calculate height
      Render1.Line (((GenC) / 5), (GenC2 + (950 - MyHeight)) / UserDetail)- _
                   (((GenC) / 5), (GenC2 + (1050 - MyHeight)) / UserDetail), ColInfo(GenC) 'draw line
     
   Next GenC2
    MyISO = 0
Next GenC

Me.AutoRedraw = False
MsgBox "Rendered"
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrorHandler
FileName = Text1.Text
SavePicture Render1.Image, FileName
MsgBox "Saved as " & FileName
Exit Sub
ErrorHandler:
End Sub

Private Sub Dir1_Change()

File1.Path = Dir1.Path

End Sub

Private Sub Drive1_Change()
On Error GoTo ErrorHandler
Dir1.Path = Drive1.Drive
Exit Sub
ErrorHandler:
End Sub

Private Sub File1_Click()
On Error GoTo ErrorHandler
FileName = Dir1.Path & "\" & File1.FileName
Set Picmap.Picture = LoadPicture(FileName)

Exit Sub
ErrorHandler:
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler
UserHeight = HScroll1.Value / 5
UserDetail = 12 ' sets level of detail, be careful if changing this, may cause strange effects to image
FillCol1 = 0
Dir1.Path = App.Path
Text1.Text = App.Path & "\Render1.bmp"
FileName = Dir1.Path & "\mtstHelens2.jpg"
Set Picmap.Picture = LoadPicture(FileName)
Exit Sub
ErrorHandler:
End Sub

Private Sub HScroll1_Change()
UserHeight = HScroll1.Value / 5
End Sub

