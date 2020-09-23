VERSION 5.00
Begin VB.Form frmResistor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Resistor Tool"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4500
   Icon            =   "Resistor.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pbColorSelect 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   3
      Left            =   2700
      Picture         =   "Resistor.frx":34CA
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   4
      Top             =   960
      Width           =   240
   End
   Begin VB.PictureBox pbColorSelect 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1500
      Index           =   2
      Left            =   1980
      Picture         =   "Resistor.frx":3E6C
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   3
      Top             =   960
      Width           =   240
   End
   Begin VB.PictureBox pbColorSelect 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1500
      Index           =   1
      Left            =   1620
      Picture         =   "Resistor.frx":516E
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   2
      Top             =   960
      Width           =   240
   End
   Begin VB.PictureBox pbColorSelect 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1500
      Index           =   0
      Left            =   1260
      Picture         =   "Resistor.frx":6470
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      Top             =   960
      Width           =   240
   End
   Begin VB.Label lblResult 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   405
      TabIndex        =   0
      Top             =   60
      Width           =   3690
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   1
      Left            =   180
      Top             =   555
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   0
      Left            =   3420
      Top             =   555
      Width           =   855
   End
   Begin VB.Shape Band 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   3
      Left            =   2700
      Top             =   315
      Width           =   255
   End
   Begin VB.Shape Band 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   2
      Left            =   1980
      Top             =   315
      Width           =   255
   End
   Begin VB.Shape Band 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   1620
      Top             =   315
      Width           =   255
   End
   Begin VB.Shape Band 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   0
      Left            =   1260
      Top             =   315
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080C0D0&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   1020
      Shape           =   4  'Rounded Rectangle
      Top             =   315
      Width           =   2415
   End
End
Attribute VB_Name = "frmResistor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim V(3) As Long

Private Sub Form_Load()
    V(0) = 2
    V(1) = 2
    V(2) = 2
    V(3) = 1
    calc
End Sub

Private Sub calc()
Dim ohms As Single
Dim tolerance As String
Dim prefix As String
    ohms = (V(0) * 10 + V(1)) * 10 ^ V(2)
    If ohms >= 1000 Then
        ohms = ohms / 1000
        prefix = "K"
    End If
    If ohms >= 1000 Then
        ohms = ohms / 1000
        prefix = "M"
    End If
    tolerance = " ± " & Trim$(Mid$(" 1 2 51020", 1 + V(3) * 2, 2)) & "%"
    lblResult.Caption = Str$(ohms) & prefix & "Ohms" & tolerance
End Sub

Private Sub pbColorSelect_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    V(Index) = Int(Y / 10)
    If Index = 3 And V(Index) = 4 Then
        Band(Index).BackColor = Shape1.BackColor
    Else
        Band(Index).BackColor = pbColorSelect(Index).Point(X, Y)
    End If
    calc
End Sub
