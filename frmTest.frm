VERSION 5.00
Object = "*\AIDK_IconButton.vbp"
Begin VB.Form frmTest 
   Caption         =   "IconButton Test"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin IDK_IconButton.IconButton IconButton7 
      Height          =   360
      Left            =   2640
      TabIndex        =   10
      Top             =   3960
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   635
      Enabled         =   0   'False
      Picture         =   "frmTest.frx":0000
   End
   Begin IDK_IconButton.IconButton IconButton6 
      Height          =   600
      Left            =   2520
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2400
      Width           =   960
      _ExtentX        =   1058
      _ExtentY        =   1058
      Picture         =   "frmTest.frx":015A
   End
   Begin IDK_IconButton.IconButton IconButton5 
      Height          =   360
      Left            =   2640
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1440
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   635
      Picture         =   "frmTest.frx":0474
      ButtonStyle     =   1
   End
   Begin IDK_IconButton.IconButton IconButton4 
      Height          =   600
      Left            =   3240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   240
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   1058
      Picture         =   "frmTest.frx":05CE
      ButtonStyle     =   1
   End
   Begin IDK_IconButton.IconButton IconButton3 
      Height          =   600
      Left            =   2520
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   1058
      Picture         =   "frmTest.frx":08E8
      ButtonStyle     =   1
   End
   Begin IDK_IconButton.IconButton IconButton2 
      Height          =   360
      Left            =   3120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1440
      Width           =   480
      _ExtentX        =   635
      _ExtentY        =   635
      Picture         =   "frmTest.frx":0C02
      ButtonStyle     =   1
   End
   Begin IDK_IconButton.IconButton IconButton1 
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "click here"
      Top             =   3360
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   635
      Picture         =   "frmTest.frx":0D5C
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disabled button:"
      Height          =   195
      Left            =   840
      TabIndex        =   11
      Top             =   4080
      Width           =   1155
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Standard button with the value set:"
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Standard popup button with a 32 x 32 icon:"
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "XP buttons with 16 x 16 icons. The value is toggled when clicked so it will function like a toolbar button:"
      Height          =   975
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XP buttons with 32 x 32 icons:"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   2145
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub IconButton1_Click()
    If IconButton1.Value = ibPressed Then
        IconButton1.Value = ibUnpressed
    Else
        IconButton1.Value = ibPressed
    End If
End Sub

Private Sub IconButton2_Click()
    If IconButton2.Value = ibPressed Then
        IconButton2.Value = ibUnpressed
    Else
        IconButton2.Value = ibPressed
    End If
End Sub

Private Sub IconButton5_Click()
    If IconButton5.Value = ibPressed Then
        IconButton5.Value = ibUnpressed
    Else
        IconButton5.Value = ibPressed
    End If
End Sub
