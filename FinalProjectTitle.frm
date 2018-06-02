VERSION 5.00
Begin VB.Form frmTitle 
   BackColor       =   &H00000000&
   Caption         =   "Who Wants to be a Millionaire"
   ClientHeight    =   7365
   ClientLeft      =   4230
   ClientTop       =   2745
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   9105
   Begin VB.CommandButton cmdHowToPlay 
      Caption         =   "How to play"
      Height          =   735
      Left            =   1440
      TabIndex        =   4
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   735
      Left            =   480
      TabIndex        =   3
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   735
      Left            =   6960
      TabIndex        =   2
      Top             =   6240
      Width           =   2055
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play Now"
      Height          =   735
      Left            =   3600
      TabIndex        =   1
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   3015
      Left            =   2880
      Picture         =   "FinalProjectTitle.frx":0000
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Who Wants to be a Millionaire"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   1680
      TabIndex        =   0
      Top             =   600
      Width           =   5655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      X1              =   1560
      X2              =   1080
      Y1              =   360
      Y2              =   1080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      X1              =   1080
      X2              =   1680
      Y1              =   1080
      Y2              =   1920
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      X1              =   1560
      X2              =   7320
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      X1              =   7320
      X2              =   8040
      Y1              =   360
      Y2              =   1080
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      X1              =   8040
      X2              =   7440
      Y1              =   1080
      Y2              =   1920
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      X1              =   1680
      X2              =   7440
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      X1              =   1080
      X2              =   120
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      X1              =   8880
      X2              =   8040
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Image Image1 
      Height          =   7215
      Left            =   0
      Picture         =   "FinalProjectTitle.frx":1483D
      Stretch         =   -1  'True
      Top             =   360
      Width           =   9135
   End
End
Attribute VB_Name = "frmTitle"
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



Private Sub Command1_Click()

End Sub

Private Sub cmdAbout_Click()
    MsgBox "Copyright By Raymond Yang", vbInformation, "ABOUT"
End Sub

Private Sub cmdExit_Click()
    EndProgram
End Sub

Private Sub cmdHowToPlay_Click()
    frmHowToPlay.Show vbModal
End Sub

Private Sub cmdPlay_Click()
    frmMain.Show vbModeless
    Unload frmTitle
    
    
End Sub

