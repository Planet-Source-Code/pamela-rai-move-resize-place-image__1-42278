VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   446
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   571
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Edit controls"
      Height          =   6615
      Left            =   6120
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":4CAE
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1485
   End
   Begin VB.Image Image3 
      Height          =   210
      Left            =   1320
      MousePointer    =   8  'Size NW SE
      Picture         =   "Form1.frx":55A7
      Top             =   6480
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   3  'Dot
      Height          =   255
      Left            =   120
      Top             =   6360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   855
      Index           =   1
      Left            =   120
      Picture         =   "Form1.frx":57EA
      Stretch         =   -1  'True
      Top             =   840
      Width           =   930
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public holdx
Public holdy
Public activeindex
Dim oldX As Long, _
    oldY As Long, _
    isMoving As Boolean
    







Private Sub Command1_Click()

End Sub

Private Sub Form_Click()
Shape1.Visible = False
Image3.Visible = False
End Sub

Private Sub Form_Load()
'MsgBox Screen.TwipsPerPixelX

End Sub

Private Sub Image2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
X = X / 15
Y = Y / 15
Shape1.Visible = False
Image3.Visible = False
Image2(activeindex).BorderStyle = 1
  

    oldX = X
    oldY = Y
    isMoving = True

End Sub

Private Sub Image2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    X = X / 15
Y = Y / 15
     Image2(activeindex).MousePointer = 15
    activeindex = Index
    
    
    If isMoving Then
 
        Image2(Index).Top = Image2(Index).Top - (oldY - Y)
        Image2(Index).Left = Image2(Index).Left - (oldX - X)

    End If
End Sub

Private Sub Image2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    X = X / 15
Y = Y / 15
    Shape1.Visible = True
    Image3.Visible = True
    Shape1.Top = Image2(activeindex).Top - 8
    Shape1.Left = Image2(activeindex).Left - 8
    Shape1.Height = Image2(activeindex).Height + 16
    Shape1.Width = Image2(activeindex).Width + 16
    Image3.Top = Shape1.Top + (Shape1.Height - 8)
    Image3.Left = Shape1.Left + (Shape1.Width - 8)
    isMoving = False
Image2(activeindex).BorderStyle = 0
Image2(activeindex).MousePointer = 0
End Sub




Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    X = X / 15
Y = Y / 15
    oldX = X
 
    oldY = Y
    isMoving = True
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    X = X / 15
Y = Y / 15
    If isMoving Then
 
        Image3.Top = Image3.Top - (oldY - Y)
        Image3.Left = Image3.Left - (oldX - X)
Shape1.Width = Image3.Left - (Shape1.Left - 8)
Shape1.Height = Image3.Top - (Shape1.Top - 8)
    End If

End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
X = (X / 15)
Y = (Y / 15)
isMoving = False
Image2(activeindex).Stretch = True


Image2(activeindex).Height = Shape1.Height - 16
Image2(activeindex).Width = Shape1.Width - 16
'
End Sub

