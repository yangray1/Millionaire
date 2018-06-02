VERSION 5.00
Begin VB.Form frmCallAFriend 
   Caption         =   "Who wants to be a Millionaire- call a friend"
   ClientHeight    =   6525
   ClientLeft      =   5070
   ClientTop       =   3210
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   7095
   Begin VB.CommandButton cmdBackToGame 
      Caption         =   "Back to the game"
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   3855
      Left            =   120
      Picture         =   "FinalProjectCall.frx":0000
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   480
      Picture         =   "FinalProjectCall.frx":5C29
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label lblDisplay 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   2280
      TabIndex        =   0
      Top             =   1920
      Width           =   4575
   End
End
Attribute VB_Name = "frmCallAFriend"
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

Private Sub cmdBackToGame_Click()
    Unload frmCallAFriend
End Sub



Private Sub Form_Load()
    
    'Declare the variables.
    
    Dim Name As String
    Dim Number As Integer
    Dim Msg As String
    Dim Valid As Boolean
    
    Randomize
    
    'Validate the name of the person being called.
    
    Do
        Name = InputBox$("Who do you wish to call?:", "CALLING...")
        If Trim$(Name) = "" Then
             MsgBox "Error. Please enter the person you wish to call.", vbCritical, "ERROR"
        End If
    Loop While Trim$(Name) = ""
    
    Valid = True
    
    'Check if the person has chose the ones that are not crossed out from the fifty fifty chance.
    
    Do
        
        Number = Int(Rnd * (4)) + 1
        If frmMain!linCross(Number).Visible <> True Then
            Valid = False
        End If
    Loop While Valid
    
    'Display the message.
    
    Msg = "Hi, this is " & Name & ". I heard you're going for the 1 million dollars,"
    
    Msg = Msg & "but you're stuck on this question so you decided to call me."
    Msg = Msg & " I think the answer is "
    
    Select Case Number
        Case 1
            Msg = Msg & "A"
        Case 2
            Msg = Msg & "B"
        Case 3
            Msg = Msg & "C"
        Case 4
            Msg = Msg & "D"
    End Select
    
    lblDisplay.Caption = Msg
        
   
        
    
    
End Sub

