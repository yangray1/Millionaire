VERSION 5.00
Begin VB.Form frmAskAudience 
   Caption         =   "Who Wants to be a Millionaire- Ask the audience"
   ClientHeight    =   6495
   ClientLeft      =   4200
   ClientTop       =   2910
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   6150
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Back to game"
      Height          =   615
      Left            =   4560
      TabIndex        =   4
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   0
      Picture         =   "FinalProjectAskAudience.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Poll"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label lblD 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   4800
      Width           =   3495
   End
   Begin VB.Label lblC 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   4080
      Width           =   3495
   End
   Begin VB.Label lblB 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   3360
      Width           =   3495
   End
   Begin VB.Label lblA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   2640
      Width           =   3495
   End
End
Attribute VB_Name = "frmAskAudience"
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

Private Sub cmdMenu_Click()
    Unload frmAskAudience
    
End Sub

Private Sub Form_Load()
    
    'Declare the variables.
    
    Dim A As Integer
    Dim B As Integer
    Dim C As Integer
    Dim D As Integer
    Dim k As Integer
    
    Randomize
    
    'Generate 4 random numbers and make sure they add up to 100.
    
    A = Int(Rnd * (97)) + 1
    If A = 97 Then
        B = 1
        C = 1
        D = 1
    Else
        B = Int(Rnd * ((98 - A))) + 1
        If A + B = 98 Then
            C = 1
            D = 1
        Else
            C = Int(Rnd * ((99 - B - A))) + 1
            If A + B + C = 99 Then
                D = 1
            Else
                D = 100 - A - B - C
            End If
        End If
    End If
    
    'Display the results.
    
    lblA.Caption = "A: " & A & "%"
    lblB.Caption = "B: " & B & "%"
    lblC.Caption = "C: " & C & "%"
    lblD.Caption = "D: " & D & "%"
End Sub

Private Sub Label2_Click()
End Sub

Private Sub Label4_Click()

End Sub
