VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000015&
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3840
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6120
      Top             =   2280
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   240
      Max             =   200
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3000
      Value           =   100
      Width           =   3015
   End
   Begin VB.CheckBox Check1 
      Caption         =   "EruLabel(0).BackStyle = Opaque"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   3360
      Width           =   3015
   End
   Begin Project1.EruLabel EruLabel1 
      Height          =   600
      Index           =   2
      Left            =   720
      Top             =   1440
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1058
      AutoSize        =   -1  'True
      BackStyle       =   1
      Caption         =   "Form1.frx":698A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   15
      PaddingBottom   =   5
      PaddingLeft     =   5
      PaddingRight    =   5
      PaddingTop      =   5
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Label1(0) (yeah, a normal one)"
      ForeColor       =   &H8000000F&
      Height          =   375
      Index           =   0
      Left            =   2880
      MousePointer    =   15  'Size All
      TabIndex        =   1
      Top             =   1080
      Width           =   2655
   End
   Begin Project1.EruLabel EruLabel1 
      Height          =   1455
      Index           =   1
      Left            =   5040
      Top             =   360
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   2566
      AutoSize        =   -1  'True
      BackColor       =   12648447
      BackStyle       =   1
      BorderColor     =   8438015
      BorderRadius    =   15
      BorderWidth     =   1
      Caption         =   "Form1.frx":69CC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16576
      MousePointer    =   15
      PaddingBottom   =   10
      PaddingLeft     =   10
      PaddingRight    =   10
      PaddingTop      =   10
      WordWrap        =   -1  'True
   End
   Begin Project1.EruLabel EruLabel1 
      Height          =   930
      Index           =   0
      Left            =   240
      Top             =   240
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   1640
      AutoSize        =   -1  'True
      BackColor       =   -2147483635
      BackStyle       =   1
      BorderColor     =   -2147483627
      BorderRadius    =   19
      BorderStyle     =   2
      Caption         =   "Form1.frx":6B0A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483634
      MarginBottom    =   2
      MarginLeft      =   5
      MarginRight     =   4
      MarginTop       =   3
      MousePointer    =   15
      Opacity         =   0,5
      PaddingBottom   =   10
      PaddingLeft     =   10
      PaddingRight    =   10
      PaddingTop      =   10
   End
   Begin Project1.EruLabel EruLabel1 
      Height          =   2100
      Index           =   3
      Left            =   3360
      Top             =   1560
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   3704
      Alignment       =   2
      AutoSize        =   -1  'True
      BackStyle       =   1
      BorderRadius    =   30
      Caption         =   "Form1.frx":6B66
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   15
      PaddingBottom   =   10
      PaddingLeft     =   10
      PaddingRight    =   10
      PaddingTop      =   10
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OldX As Single, OldY As Single

Private Sub Check1_Click()
    If EruLabel1(0).BackStyle <> Check1.Value Then EruLabel1(0).BackStyle = Check1.Value
End Sub

Private Sub EruLabel1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And vbLeftButton Then
        ' EruLabel uses the same ScaleMode it's container uses! Regular Label does not!
        OldX = X
        OldY = Y
        EruLabel1(Index).ZOrder
    End If
End Sub

Private Sub EruLabel1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And vbLeftButton Then
        ' EruLabel uses the same ScaleMode it's container uses! Regular Label does not!
        EruLabel1(Index).Move EruLabel1(Index).Left + X - OldX, EruLabel1(Index).Top + Y - OldY
    End If
End Sub

Private Sub HScroll1_Change()
    EruLabel1(0).Opacity = HScroll1.Value / HScroll1.Max
    Me.Caption = "Opacity: " & EruLabel1(0).Opacity
End Sub

Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And vbLeftButton Then
        ' EruLabel uses the same ScaleMode it's container uses! Regular Label does not!
        OldX = Me.ScaleX(X, vbTwips, Me.ScaleMode)
        OldY = Me.ScaleY(Y, vbTwips, Me.ScaleMode)
        Label1(Index).ZOrder
    End If
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And vbLeftButton Then
        ' EruLabel uses the same ScaleMode it's container uses! Regular Label does not!
        Label1(Index).Move Label1(Index).Left + Me.ScaleX(X, vbTwips, Me.ScaleMode) - OldX, _
            Label1(Index).Top + Me.ScaleY(Y, vbTwips, Me.ScaleMode) - OldY
    End If
End Sub

Private Sub Form_Load()
    Check1.Value = EruLabel1(0).BackStyle
    HScroll1.Value = EruLabel1(0).Opacity * HScroll1.Max
End Sub

Private Sub Timer1_Timer()
    Static Counter As Long
    Counter = Counter + 1
    If Counter = 50 Then Counter = 0: EruLabel1(2).Caption = EruLabel1(2).Caption & vbNewLine
    EruLabel1(2).Caption = EruLabel1(2).Caption & "."
End Sub
