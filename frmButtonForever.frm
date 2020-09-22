VERSION 5.00
Begin VB.Form frmButtonForever 
   Caption         =   "Owner Draw ActiveX Button"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   4830
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   167
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   322
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   2460
      TabIndex        =   1
      Top             =   180
      Width           =   2235
   End
   Begin prjButtonForever.Command Command1 
      Height          =   360
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   635
      Caption         =   "Forever Blue"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientEndColor=   16750093
   End
   Begin prjButtonForever.Command Command1 
      Height          =   360
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   600
      Width           =   1200
      _ExtentX        =   2884
      _ExtentY        =   873
      Caption         =   "Custom"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientEndColor=   255
      GradientStartColor=   65535
   End
   Begin prjButtonForever.Command Command1 
      Height          =   480
      Index           =   3
      Left            =   180
      TabIndex        =   3
      Top             =   1020
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Caption         =   "Go"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientEndColor=   16750093
   End
   Begin prjButtonForever.Command Command1 
      Height          =   840
      Index           =   5
      Left            =   180
      TabIndex        =   4
      Top             =   1560
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   1482
      Caption         =   "&Custom"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientEndColor=   255
      GradientStartColor=   65535
      ForeColor       =   16777215
   End
End
Attribute VB_Name = "frmButtonForever"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Showevent(Idx As Integer, Nam As String)
   List1.AddItem "Button " & Format(Idx, "00 ") & Nam
   List1.ListIndex = List1.ListCount - 1
End Sub

Private Sub Command1_Click(Index As Integer)
   Showevent Index, "Click"
End Sub

Private Sub Command1_DblClick(Index As Integer)
   Showevent Index, "DblClick"
End Sub

Private Sub Command1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Showevent Index, "MouseDown"
End Sub

Private Sub Command1_MouseEnter(Index As Integer)
   Showevent Index, "MouseEnter"
End Sub

Private Sub Command1_MouseExit(Index As Integer)
   Showevent Index, "MouseExit"
End Sub

Private Sub Command1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Showevent Index, "MouseUp"
End Sub


