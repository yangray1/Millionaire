VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Who Wants to be a Millionaire"
   ClientHeight    =   8250
   ClientLeft      =   3990
   ClientTop       =   3210
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   10410
   Begin VB.CommandButton cmdLeave 
      Caption         =   "Leave with money"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5880
      TabIndex        =   28
      Top             =   7680
      Width           =   2055
   End
   Begin VB.Label lblWinOrNot 
      Caption         =   "1"
      Height          =   255
      Left            =   8640
      TabIndex        =   33
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   3360
      X2              =   3360
      Y1              =   0
      Y2              =   8280
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      Height          =   1455
      Left            =   4440
      Top             =   4320
      Width           =   4935
   End
   Begin VB.Label lblD 
      BackStyle       =   0  'Transparent
      Caption         =   "D:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   6960
      TabIndex        =   32
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label lblB 
      BackStyle       =   0  'Transparent
      Caption         =   "B:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   6960
      TabIndex        =   31
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label lblC 
      BackStyle       =   0  'Transparent
      Caption         =   "C:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   3480
      TabIndex        =   30
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label lblA 
      BackStyle       =   0  'Transparent
      Caption         =   "A:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   3480
      TabIndex        =   29
      Top             =   6120
      Width           =   375
   End
   Begin VB.Line linCross 
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      Index           =   4
      Visible         =   0   'False
      X1              =   7080
      X2              =   10200
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line linCross 
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      Index           =   3
      Visible         =   0   'False
      X1              =   3720
      X2              =   6840
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line linCross 
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      Index           =   2
      Visible         =   0   'False
      X1              =   7080
      X2              =   10200
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line linCross 
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      Index           =   1
      Visible         =   0   'False
      X1              =   3720
      X2              =   6840
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Image imgHold 
      Height          =   375
      Index           =   2
      Left            =   600
      Picture         =   "FinalProjectMain.frx":0000
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgHold 
      Height          =   375
      Index           =   12
      Left            =   6600
      Picture         =   "FinalProjectMain.frx":7C30
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgHold 
      Height          =   375
      Index           =   11
      Left            =   6000
      Picture         =   "FinalProjectMain.frx":F320
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgHold 
      Height          =   375
      Index           =   10
      Left            =   5400
      Picture         =   "FinalProjectMain.frx":15A0A
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgHold 
      Height          =   375
      Index           =   9
      Left            =   4800
      Picture         =   "FinalProjectMain.frx":1C7F5
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgHold 
      Height          =   375
      Index           =   8
      Left            =   4200
      Picture         =   "FinalProjectMain.frx":242F5
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgHold 
      Height          =   375
      Index           =   7
      Left            =   3600
      Picture         =   "FinalProjectMain.frx":2B88A
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgHold 
      Height          =   375
      Index           =   6
      Left            =   3000
      Picture         =   "FinalProjectMain.frx":33726
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgHold 
      Height          =   375
      Index           =   5
      Left            =   2400
      Picture         =   "FinalProjectMain.frx":3A356
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgHold 
      Height          =   375
      Index           =   4
      Left            =   1800
      Picture         =   "FinalProjectMain.frx":40D05
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgHold 
      Height          =   375
      Index           =   3
      Left            =   1200
      Picture         =   "FinalProjectMain.frx":47618
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgHold 
      Height          =   375
      Index           =   1
      Left            =   0
      Picture         =   "FinalProjectMain.frx":4EC6A
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgMain 
      Height          =   3135
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   840
      Width           =   5535
   End
   Begin VB.Label lblAns 
      Height          =   495
      Index           =   5
      Left            =   6720
      TabIndex        =   27
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblAskAudience 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2280
      TabIndex        =   26
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblHelp3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ask the Audience"
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
      Height          =   495
      Left            =   2280
      TabIndex        =   25
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblHelp2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Call a friend"
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
      Height          =   375
      Left            =   1320
      TabIndex        =   24
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblHelp1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fifty fifty chance"
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
      Height          =   375
      Left            =   240
      TabIndex        =   23
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblTipCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Tip:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3480
      TabIndex        =   22
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblTip 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   735
      Left            =   4080
      TabIndex        =   21
      Top             =   120
      Width           =   5895
   End
   Begin VB.Line Line23 
      BorderWidth     =   3
      X1              =   6480
      X2              =   6480
      Y1              =   6960
      Y2              =   7440
   End
   Begin VB.Line Line22 
      BorderWidth     =   3
      X1              =   3960
      X2              =   3960
      Y1              =   6960
      Y2              =   7440
   End
   Begin VB.Line Line21 
      BorderWidth     =   3
      X1              =   3960
      X2              =   6480
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line20 
      BorderWidth     =   3
      X1              =   3960
      X2              =   6480
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Label lblAns 
      Alignment       =   2  'Center
      Caption         =   "aa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   3960
      TabIndex        =   20
      Top             =   6960
      Width           =   2535
   End
   Begin VB.Label lblAns 
      Alignment       =   2  'Center
      Caption         =   "aa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   7440
      TabIndex        =   19
      Top             =   6960
      Width           =   2535
   End
   Begin VB.Line Line19 
      BorderWidth     =   7
      X1              =   7440
      X2              =   9960
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line18 
      BorderWidth     =   7
      X1              =   7440
      X2              =   9960
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line17 
      BorderWidth     =   7
      X1              =   7440
      X2              =   7440
      Y1              =   6960
      Y2              =   7440
   End
   Begin VB.Line Line16 
      BorderWidth     =   7
      X1              =   9960
      X2              =   9960
      Y1              =   6960
      Y2              =   7440
   End
   Begin VB.Label lblQuestion 
      Alignment       =   2  'Center
      Caption         =   "Question"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4440
      TabIndex        =   18
      Top             =   4320
      Width           =   4935
   End
   Begin VB.Line Line15 
      BorderWidth     =   3
      X1              =   9960
      X2              =   9960
      Y1              =   6120
      Y2              =   6600
   End
   Begin VB.Line Line14 
      BorderWidth     =   3
      X1              =   7440
      X2              =   7440
      Y1              =   6120
      Y2              =   6600
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      X1              =   7440
      X2              =   9960
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line12 
      BorderWidth     =   3
      X1              =   7440
      X2              =   9960
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label lblAns 
      Alignment       =   2  'Center
      Caption         =   "aa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   7440
      TabIndex        =   17
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Line Line11 
      BorderWidth     =   3
      X1              =   6480
      X2              =   6480
      Y1              =   6120
      Y2              =   6600
   End
   Begin VB.Line Line10 
      BorderWidth     =   3
      X1              =   3960
      X2              =   3960
      Y1              =   6120
      Y2              =   6600
   End
   Begin VB.Line Line9 
      BorderWidth     =   3
      X1              =   3960
      X2              =   6480
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line8 
      BorderWidth     =   3
      X1              =   3960
      X2              =   6480
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label lblAns 
      Alignment       =   2  'Center
      Caption         =   "aa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3960
      TabIndex        =   16
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   2280
      Picture         =   "FinalProjectMain.frx":550B6
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
   Begin VB.Image imgTwo 
      Height          =   495
      Left            =   1320
      Picture         =   "FinalProjectMain.frx":58399
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblFiftyChance 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "50:50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   240
      Width           =   975
   End
   Begin VB.Shape shpHelp 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   120
      Shape           =   2  'Oval
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "$1 MILLION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   15
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "$500 000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "$250 000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "$125 000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "$64 000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "$32 000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   10
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "$16 000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   8
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "$8 000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "$4 000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   6
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "$2 000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   5
      Top             =   5400
      Width           =   2775
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "$1 000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   4
      Top             =   5880
      Width           =   2775
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "$500"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   6360
      Width           =   2775
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "$300"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   2
      Top             =   6840
      Width           =   2775
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "$200"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   7320
      Width           =   2775
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "$100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   7800
      Width           =   2775
   End
   Begin VB.Shape shpHelp2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1200
      Shape           =   2  'Oval
      Top             =   120
      Width           =   975
   End
   Begin VB.Image imgBackGround1 
      Height          =   8295
      Left            =   0
      Picture         =   "FinalProjectMain.frx":58B9E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3375
   End
   Begin VB.Image imgBackground2 
      Height          =   8295
      Left            =   3360
      Picture         =   "FinalProjectMain.frx":5AEDC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7095
   End
   Begin VB.Menu frmFile 
      Caption         =   "File"
      Begin VB.Menu cmdNewGame 
         Caption         =   "New Game"
      End
      Begin VB.Menu cmdExit 
         Caption         =   "Exit Game"
      End
   End
   Begin VB.Menu cmdAbout 
      Caption         =   "About"
      Begin VB.Menu cmdHowToPlay 
         Caption         =   "How to play"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Title: Who Wants to be a Millionaire.
'Author: Raymond Yang
'Date:Wednesday June 3rd, 2015
'Files: FinalProject.bas, FinalProject.vbp, FinalProjectAskAudience.frm,FinalProjectAskAudience.frx,
'       FinalProjectCall.frm,FinalProjectCall.frx,FinalProjectCheck.frm,FinalProjectCheck.frx,FinalProjectHowToPlay.frm,
'       FinalProjectHowToPlay.frx,FinalProjectMain.frm,FinalProjectMain.frx,FinalProjectTitle.frm,FinalProjectTitle.frx,
'       Questions.txt,Tip.txt
'Purpose: This application is used to simulate the game 'Who wants to be a millionaire'.
Option Explicit
Dim Question(1 To 100) As String
Dim A(1 To 100) As String
Dim B(1 To 100) As String
Dim C(1 To 100) As String
Dim D(1 To 100) As String
Dim Answer(1 To 100) As String
Dim Used(1 To 100) As Boolean
Dim tip(1 To 100) As String

Private Sub cmdExit_Click()
    EndProgram
End Sub

  

Private Sub Image1_Click()

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
End Sub



Private Sub cmdHowToPlay_Click()
    frmHowToPlay.Show vbModal
End Sub

Private Sub cmdLeave_Click()
    Dim Temp As Integer
    
    Temp = MsgBox("Are you sure you want to leave with the money?", vbInformation + vbYesNo, "LEAVE")
    If Temp = vbYes Then
        lblWinOrNot.Caption = 2
        Load frmCheck
        
        frmCheck.Show vbModeless
        
    End If
End Sub


Private Sub cmdNewGame_Click()
    Unload frmMain
    frmMain.Show vbModeless
End Sub

Private Sub Form_Load()
    
    Dim k As Integer

    Randomize
    
    'Sets up the game with a question,a tip and a picture.
    
    ReadFile Used(), A(), B(), C(), D(), Answer(), Question(), tip()
    GetQuestion frmMain, Used(), 1, Question(), A(), B(), C(), D(), Answer()
    NewTip frmMain, tip()
    lblPrize(1).BackStyle = 1
    
    GetPicture frmMain
    cmdLeave.Enabled = False
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearMouseMove frmMain
End Sub

P


Private Sub imgBackGround1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearMouseMove frmMain
End Sub

Private Sub imgBackground2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearMouseMove frmMain
End Sub

Private Sub imgMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearMouseMove frmMain
End Sub

Private Sub imgTwo_Click()
    Dim Temp As Integer
    
    Temp = MsgBox("Call a friend?", vbQuestion + vbYesNo, "CONFIRM")
    If Temp = vbYes Then
    
        imgTwo.Enabled = False
        lblHelp2.Enabled = False
        frmCallAFriend.Show vbModal
    End If
End Sub

Private Sub imgTwo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearMouseMove frmMain
    shpHelp2.BackColor = vbYellow
    lblHelp2.Visible = True
End Sub

Private Sub Label4_Click()

End Sub

Private Sub lblAns_Click(Index As Integer)
    
    'Declare the variables.
    
    Dim Temp As Integer
    Dim Temp2 As String
    Dim NextHighlight As Integer
    Dim Final As Integer
    Dim k As Integer
    Dim Difficulty As Integer
    Dim Valid As Boolean
    
    'Verify if the answer is their final answer.
    
    Final = MsgBox("Is this your final answer?", vbInformation + vbYesNo, "FINAL ANSWER")
    If Final = vbYes Then
    
        Temp = Val(lblAns(5))
        
        'Determine if the answer is correct.
        
        If Index = Temp Then
            MsgBox "Correct!", vbInformation, "WINNER"
            
            'Determine if the person answered all 15 questions.
            
            If lblPrize(15).BackStyle = 1 Then
                lblWinOrNot.Caption = 3
                Load frmCheck
                frmCheck.Show vbModeless
            Else
                k = 0
                Do
                    Valid = True
                    k = k + 1
                    If lblPrize(k).BackStyle = 1 Then
                    
                    Valid = False
                    NextHighlight = k + 1
                    
                    'Determine the difficulty of the next question.
                    
                    Select Case (k + 1)
                        Case 15
                            Difficulty = 4
                        Case Is > 10
                            Difficulty = 3
                        Case Is > 5
                            Difficulty = 2
                        Case Else
                            Difficulty = 1
                    End Select
                    End If
                    
                    'Change the colour of the 5th,10th and 15th money prize label.
                    
                    If (k = 5) Or (k = 10) Or (k = 15) Then
                        lblPrize(k).ForeColor = vbWhite
                        
                    End If
                    lblPrize(k).BackStyle = 0
                    If k = 15 Then
                        Valid = False
                    End If
                Loop While Valid
                
                
                If NextHighlight = 5 Or NextHighlight = 10 Or NextHighlight = 15 Then
                    lblPrize(NextHighlight).ForeColor = vbBlack
                End If
                
                'Sets up a new question,tip and picture.
                
                lblPrize(NextHighlight).BackStyle = 1
                cmdLeave.Enabled = True
                NewTip frmMain, tip()
        
                GetQuestion frmMain, Used(), Difficulty, Question(), A(), B(), C(), D(), Answer()
   
                GetPicture frmMain
                
                'Clears any visible crossouts from the use of the 50-50 lifeline.
                
                For k = 1 To 4
                    linCross(k).Visible = False
                    lblAns(k).Enabled = True
                Next k
            End If
        Else
            Temp2 = lblAns(Temp).Caption
            MsgBox "Sorry, the correct answer was " & Temp2 & ". You lose.", vbInformation, "LOST"

            frmCheck.Show vbModeless
            
        End If
        
        
        
    End If
    
    
End Sub

Private Sub lblAns_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearMouseMove frmMain
    lblAns(Index).BackColor = vbYellow
    
End Sub

Private Sub lblAskAudience_Click()
    Dim Temp As Integer
    
    'Verify if the person wants to ask the audience.
    
    Temp = MsgBox("Ask the audience?", vbQuestion + vbYesNo, "CONFIRM")
    If Temp = vbYes Then
        lblAskAudience.Enabled = False
        lblHelp3.Enabled = False
        frmAskAudience.Show vbModal
    End If
End Sub

Private Sub lblAskAudience_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearMouseMove frmMain
    lblHelp3.Visible = True
End Sub

Private Sub lblD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearMouseMove frmMain
End Sub
Private Sub lblB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearMouseMove frmMain
End Sub
Private Sub lblC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearMouseMove frmMain
End Sub
Private Sub lblA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearMouseMove frmMain
End Sub

Private Sub lblFiftyChance_Click()
    
    'Declare the variables.
    
    Dim Temp As Integer
    Dim Temp2 As Integer
    Dim OldNumber As Integer
    Dim Number As Integer
    Dim Valid As Boolean
    Dim k As Integer
    
    'Verify if the person wants to use fifty fifty chance.
    
    Temp = MsgBox("Use fifty fifty chance?", vbQuestion + vbYesNo, "CONFIRM")

    If Temp = vbYes Then
        Number = 0
        OldNumber = 0
        
        Temp2 = Val(lblAns(5).Caption)
        For k = 1 To 2
            
            'Generate two random question numbers to cross out.
            
            Do
                Valid = True
                Number = Int(Rnd * (4)) + 1
                If Number = Temp2 Or OldNumber = Number Then
                    Valid = False
                End If
                
            Loop While Not Valid
            
            'Cross out two options and make them unchooseable.
            
            lblAns(Number).Enabled = False
            linCross(Number).Visible = True
            OldNumber = Number
        Next k
        lblFiftyChance.Enabled = False
        lblHelp1.Enabled = False
    End If
    
    
            
End Sub

Private Sub lblFiftyChance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearMouseMove frmMain
    shpHelp.BackColor = vbYellow
    lblFiftyChance.ForeColor = vbRed
    lblHelp1.Visible = True
End Sub



Private Sub lblHelp1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearMouseMove frmMain
End Sub

Private Sub lblHelp2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearMouseMove frmMain
End Sub

Private Sub lblHelp3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearMouseMove frmMain
End Sub

Private Sub lblPrize_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearMouseMove frmMain
End Sub

Private Sub lblQuestion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearMouseMove frmMain
End Sub

Private Sub lblTip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearMouseMove frmMain
End Sub

Private Sub lblTipCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearMouseMove frmMain
End Sub
