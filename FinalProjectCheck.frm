VERSION 5.00
Begin VB.Form frmCheck 
   Caption         =   "Check"
   ClientHeight    =   5325
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   8880
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit game"
      Height          =   855
      Left            =   6240
      TabIndex        =   6
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play again"
      Height          =   855
      Left            =   3240
      TabIndex        =   5
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton cmdTitle 
      Caption         =   "Go back to title page"
      Height          =   855
      Left            =   480
      TabIndex        =   4
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "2008/01/01"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label lblMoney2 
      Alignment       =   2  'Center
      Caption         =   "------------------------Money-------------------------"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1920
      Width           =   6375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "You"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   5
      Height          =   495
      Left            =   7080
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblMoney 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "You"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   3975
      Left            =   120
      Picture         =   "FinalProjectCheck.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   8655
   End
End
Attribute VB_Name = "frmCheck"
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

Private Sub cmdExit_Click()
    EndProgram
End Sub

Private Sub cmdPlay_Click()
    frmCheck.Hide
    frmMain.Show vbModeless
    Unload frmCheck
End Sub

Private Sub cmdTitle_Click()
    frmCheck.Hide
    frmTitle.Show vbModeless
    Unload frmCheck
End Sub

Private Sub Form_Load()
    Dim Money As String
    Dim k As Integer
    Dim Valid As Boolean
    
    Valid = True
    Do
        k = k + 1
        
        'Determine what the prize was when the player lost/won.
        
        If frmMain!lblPrize(k).BackStyle = 1 Then
            If frmMain!lblWinOrNot.Caption = 2 Then
                Money = frmMain!lblPrize(k - 1).Caption
                
                
            ElseIf frmMain!lblWinOrNot.Caption = 1 Then
            
                'Determine the prize.
                
                If (k > 10) Then
                    Money = frmMain!lblPrize(10).Caption
                ElseIf (k > 5) Then
                    Money = frmMain!lblPrize(5).Caption
                Else
                    Money = "$0.00"
                End If
            Else
                Money = "$1 000 000"
            End If
            Valid = False
           
            frmMain!lblPrize(k).BackStyle = 0
        End If
            
                
        
    Loop While Valid = True
    
    'Display the results in the check.
    
    lblMoney.Caption = Money
    lblMoney2.Caption = "----------------" & Money & " dollars ------------------------"
    
    
    Unload frmMain
    
   
    
    
    
    
    
    
    
End Sub

    
Private Sub txtWinOrNot_Change()

End Sub

