VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Evaluator"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   413
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frEvalSec 
      Caption         =   "Evaluate For One &Second"
      Enabled         =   0   'False
      Height          =   1035
      Left            =   120
      TabIndex        =   16
      Top             =   4380
      Width           =   4995
      Begin VB.CommandButton cmdGoSec 
         Caption         =   "&Go"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label lblEvalSec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Evaluate the Expression for an entire second, and return its value."
         Enabled         =   0   'False
         Height          =   390
         Left            =   1620
         TabIndex        =   18
         Top             =   240
         Width           =   3180
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblEvalSecResult 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Result:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1620
         TabIndex        =   19
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblEvalSecResultVal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         Height          =   195
         Left            =   2220
         TabIndex        =   20
         Top             =   720
         Width           =   45
      End
   End
   Begin VB.Frame frEvalMult 
      Caption         =   "Evaluate &Multiple (Not Recommended For Default Expression)"
      Enabled         =   0   'False
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   2700
      Width           =   4995
      Begin VB.TextBox txtEvalMultTimes 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Text            =   "1000"
         Top             =   1140
         Width           =   1395
      End
      Begin VB.CommandButton cmdGoMult 
         Caption         =   "G&o"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label lblEvalMultResult 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Result:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1620
         TabIndex        =   14
         Top             =   1260
         Width           =   495
      End
      Begin VB.Label lblEvalMultResultVal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         Height          =   195
         Left            =   2220
         TabIndex        =   15
         Top             =   1260
         Width           =   45
      End
      Begin VB.Label lblEvalMult 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Evaluate the Expression multiple times and return its value."
         Enabled         =   0   'False
         Height          =   390
         Left            =   1620
         TabIndex        =   11
         Top             =   360
         Width           =   3180
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblEvalMultTimes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Evaluate &Count:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   900
         Width           =   1140
      End
   End
   Begin VB.PictureBox pAlign 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   5250
      TabIndex        =   21
      Top             =   5520
      Width           =   5250
      Begin VB.Frame frHover 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   780
         TabIndex        =   25
         Top             =   360
         Width           =   2235
         Begin VB.Label lblAuthEMailVal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Linguar_Amadala@hotmail.com"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   0
            MousePointer    =   10  'Up Arrow
            TabIndex        =   26
            Top             =   0
            Width           =   2235
         End
      End
      Begin VB.Timer tHover 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   3360
         Top             =   60
      End
      Begin VB.Label lblAuthEMail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblAuthorName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Linguar Amadala (Allen Copeland)"
         Height          =   195
         Left            =   780
         TabIndex        =   23
         Top             =   120
         Width           =   2385
      End
      Begin VB.Label lblAuthor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Author:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   510
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   30
         X2              =   5250
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   5220
         Y1              =   15
         Y2              =   15
      End
   End
   Begin VB.Frame frEvalSingle 
      Caption         =   "E&valuate Once"
      Enabled         =   0   'False
      Height          =   1035
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   4995
      Begin VB.CommandButton cmdGoOnce 
         Caption         =   "&Go"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label lblEvalOnceResultVal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         Height          =   195
         Left            =   2220
         TabIndex        =   8
         Top             =   720
         Width           =   45
      End
      Begin VB.Label lblEvalOnceResult 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Result:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1620
         TabIndex        =   7
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblEvalOnce 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Evaluate the Expression once and return its value"
         Enabled         =   0   'False
         Height          =   390
         Left            =   1620
         TabIndex        =   6
         Top             =   240
         Width           =   3180
         WordWrap        =   -1  'True
      End
   End
   Begin VB.TextBox txtExpression 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "frmMain.frx":000C
      Top             =   360
      Width           =   4995
   End
   Begin VB.Label lblStatusText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Error: Empty Expression"
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   1260
      Width           =   4095
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1260
      Width           =   495
   End
   Begin VB.Label lblExpression 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Expression:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   810
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_essExpressions As EXPR_SEM_Expressions
Private m_esiIntermediateInfo As EXPR_Sem_Intermediate
Private m_lngMultTimes As Long

Private Sub cmdGoMult_Click()
    Dim m_lngLoop As Long
    Dim m_varResult As Variant
    Dim m_sngTimeStart As Single
    Dim m_sngTimeTaken As Single
    m_sngTimeStart = Timer
    For m_lngLoop = 1 To m_lngMultTimes
        m_varResult = GroupEvaluate(m_essExpressions, m_esiIntermediateInfo)
    Next
    m_sngTimeTaken = Timer - m_sngTimeStart
    lblEvalMultResultVal.Caption = FormatVal(m_varResult) & " Time Taken: " & (m_sngTimeTaken) * 1000 & "ms"
End Sub

Private Sub cmdGoOnce_Click()
    lblEvalOnceResultVal.Caption = FormatVal(GroupEvaluate(m_essExpressions, m_esiIntermediateInfo))
End Sub

Private Sub cmdGoSec_Click()
    Dim m_sngStart As Single
    Dim m_varResult As Variant
    Dim m_lngTimes As Long
    m_sngStart = Timer
    Do Until Timer - m_sngStart >= 1
        m_varResult = GroupEvaluate(m_essExpressions, m_esiIntermediateInfo)
        m_lngTimes = m_lngTimes + 1
    Loop
    lblEvalSecResultVal.Caption = FormatVal(m_varResult) & " Times: " & m_lngTimes
End Sub

Private Sub Form_Load()
    txtExpression_Change
    txtEvalMultTimes_Change
End Sub

Private Sub lblAuthEMailVal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tHover.Enabled = True
End Sub

Private Sub tHover_Timer()
    If Not MousedOver(frHover.hWnd) Then
        With lblAuthEMailVal
            If .Font.Underline Then _
                .Font.Underline = False
            If (Not (.ForeColor = vbBlue)) Then _
                .ForeColor = vbBlue
        End With
        tHover.Enabled = False
    Else
        With lblAuthEMailVal
            If (Not (.ForeColor = vbYellow)) Then _
                .ForeColor = vbRed
            If Not .Font.Underline Then _
                .Font.Underline = True
        End With
    End If
End Sub

Private Sub txtEvalMultTimes_Change()
    If IsNumeric(txtEvalMultTimes.Text) Then
        m_lngMultTimes = txtEvalMultTimes.Text
        If Not cmdGoMult.Enabled Then _
            cmdGoMult.Enabled = True
    Else
        If cmdGoMult.Enabled Then _
            cmdGoMult.Enabled = False
    End If
End Sub

Private Sub txtEvalMultTimes_KeyPress(KeyAscii As Integer)
    Select Case Chr(KeyAscii)
        Case "0" To "9", Chr(8)
        Case Else
            KeyAscii = 0
            Beep
    End Select
End Sub

Private Sub txtExpression_Change()
    Dim m_esiIntermediate As EXPR_Sem_Intermediate
    Dim m_essExprs As EXPR_SEM_Expressions
    Dim m_expExp As EXPR_Sem_Expression
    Dim m_lngIndex As Long
    m_essExprs = modInterpret.GroupParse(txtExpression.Text, m_esiIntermediate)
    If m_esiIntermediate.Exceptions.Count > 0 Then
        EnableChange False
        lblStatusText = GetExceptionStr(m_esiIntermediate, m_esiIntermediate.Exceptions.Items(0))
    ElseIf txtExpression.Text = vbNullString Then
        EnableChange False
        lblStatusText = "Error: empty expression"
    Else
        EnableChange True
        m_esiIntermediateInfo = m_esiIntermediate
        m_essExpressions = m_essExprs
        lblStatusText = "OK"
    End If
End Sub

Private Sub EnableChange(Enabled As Boolean)
    frEvalSingle.Enabled = Enabled
    lblEvalOnce.Enabled = Enabled
    lblEvalOnceResult.Enabled = Enabled
    lblEvalOnceResultVal.Enabled = Enabled
    cmdGoOnce.Enabled = Enabled
    frEvalMult.Enabled = Enabled
    cmdGoMult.Enabled = Enabled
    lblEvalMult.Enabled = Enabled
    lblEvalMultResult.Enabled = Enabled
    lblEvalMultResultVal.Enabled = Enabled
    lblEvalMultTimes.Enabled = Enabled
    txtEvalMultTimes.Enabled = Enabled
    frEvalSec.Enabled = Enabled
    lblEvalSec.Enabled = Enabled
    lblEvalSecResult.Enabled = Enabled
    lblEvalSecResultVal.Enabled = Enabled
    cmdGoSec.Enabled = Enabled
    If Enabled Then
        txtExpression.BackColor = vbWindowBackground
        txtExpression.ForeColor = vbWindowText
    Else
        txtExpression.BackColor = vbRed
        txtExpression.ForeColor = vbWhite
    End If
End Sub
