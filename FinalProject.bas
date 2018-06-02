Attribute VB_Name = "Module1"
'Title: Who Wants to be a Millionaire.
'Author: Raymond Yang
'Date:Wednesday June 3rd, 2015
'Files: FinalProject.bas, FinalProject.vbp, FinalProjectAskAudience.frm,FinalProjectAskAudience.frx,
'       FinalProjectCall.frm,FinalProjectCall.frx,FinalProjectCheck.frm,FinalProjectCheck.frx,FinalProjectHowToPlay.frm,
'       FinalProjectHowToPlay.frx,FinalProjectMain.frm,FinalProjectMain.frx,FinalProjectTitle.frm,FinalProjectTitle.frx,
'       Questions.txt,Tip.txt
'Purpose: This application is used to simulate the game 'Who wants to be a millionaire'.
Option Explicit

'This procedure asks if the person wants to end the program.

Sub EndProgram()
    Dim Temp As Integer
    Dim DType As Integer
    
    DType = vbQuestion + vbYesNo
    
    Temp = MsgBox("Do you really want to exit?", DType, "CONFIRM")
    
    If Temp = vbYes Then
        End
    End If
    
        
End Sub

'This procedure gets a new tip.

Sub NewTip(frmGeneral As Form, tip() As String)

    Dim tipMsg As String
    Dim Number As Integer
    Dim k As Integer
    Dim Valid As Boolean
   
    Number = GetNumber(6, 1)
    
    tipMsg = tip(Number)
    
    frmGeneral!lblTip.Caption = tipMsg
    
    
End Sub

'This procedure gets a question.

Sub GetQuestion(frmGeneral As Form, Used() As Boolean, Difficulty As Integer, Question() As String, A() As String, B() As String, C() As String, D() As String, Answer() As String)
   
    Dim Valid As Boolean
    Dim Number As Integer
    Dim High As Integer
    Dim Low As Integer
    
    
    
    Valid = True
   
    'Determine the difficulty of the question.
    
    If Difficulty = 1 Then
        Low = 1
        High = 10
        
    ElseIf Difficulty = 2 Then
        Low = 11
        High = 20
    ElseIf Difficulty = 3 Then
        Low = 21
        High = 30
    Else
        Low = 31
        High = 35
    End If
    
    Valid = True
    
    'Check if the question has been used.
    
    Do
        
        Number = GetNumber(High, Low)
        If Used(Number) = False Then
            Valid = False
        End If
        
    Loop While Valid = True
    
    frmMain!lblQuestion.Caption = Question(Number)
    frmMain!lblAns(1).Caption = A(Number)
    frmMain!lblAns(2).Caption = B(Number)
    frmMain!lblAns(3).Caption = C(Number)
    frmMain!lblAns(4).Caption = D(Number)
    frmMain!lblAns(5).Caption = Answer(Number)
    Used(Number) = True
    
    
    
    
 
            
End Sub

'This function returns a random number.

Function GetNumber(High As Integer, Low As Integer)

   
    GetNumber = Int(Rnd * (High - Low + 1)) + Low
    
End Function

'This procedure gets a picture of the host asking questions.

Sub GetPicture(frmGeneral As Form)
    
    Dim Number As Integer
    
    Number = Int(Rnd * (12)) + 1
    
    frmGeneral!imgMain.Picture = frmGeneral!imgHold(Number)
    
    
    
    
End Sub

'This procedure clears any mousemove events.

Sub ClearMouseMove(frmGeneral As Form)
    Dim k As Integer

    For k = 1 To 4
        frmGeneral!lblAns(k).BackColor = &H8000000F
    Next k
    
    frmGeneral!lblFiftyChance.ForeColor = vbWhite
    frmGeneral!lblFiftyChance.BackColor = vbBlack
    frmGeneral!shpHelp2.BackColor = vbWhite
    frmGeneral!lblHelp2.Visible = False
    frmGeneral!lblHelp1.Visible = False
    frmGeneral!lblHelp3.Visible = False
    frmGeneral!shpHelp.BackColor = vbBlack
    frmGeneral!lblFiftyChance.ForeColor = vbWhite
End Sub

'This procedure reads the file containing data nessesary for this game.

Sub ReadFile(Used() As Boolean, A() As String, B() As String, C() As String, D() As String, Answer() As String, Question() As String, tip() As String)

    Dim k As Integer
    
    
    k = 0
    Open App.Path & "\questions.txt" For Input As #1
        Do While Not EOF(1)
            k = k + 1
            Input #1, Question(k), A(k), B(k), C(k), D(k), Answer(k)
            Used(k) = False
    
        Loop
    Close #1
    k = 0
    Open App.Path & "\tip.txt" For Input As #1
        Do While Not EOF(1)
            k = k + 1
            Input #1, tip(k)
        Loop
    Close #1
    
    
    
End Sub










