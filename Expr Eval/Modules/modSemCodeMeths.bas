Attribute VB_Name = "modSemCodeMeths"
Option Explicit
'//Author: Linguar Amadala (Allen Copeland)

Public Sub ThrowException(ByRef Intermediate As EXPR_Sem_Intermediate, ByVal Position As Long, ByVal Reason As String, ByVal ErrorCode As EXPR_SEM_Intermediate_EXP_Type)
    With Intermediate.Exceptions
        If .Count = 0 Then
            ReDim .Items(.Count)
        Else
            ReDim Preserve .Items(.Count)
        End If
        With .Items(.Count)
            .Reason = AddString(Intermediate.ConstTable, Reason, False)
            .Error = ErrorCode
            .Position = Position
        End With
        .Count = .Count + 1
    End With
End Sub

Public Function GetToken(ByRef Index As Long, ByRef l As EXPR_Lex_Tokens) As EXPR_Lex_Token
    GetToken = l.Tokens(Index)
End Function
