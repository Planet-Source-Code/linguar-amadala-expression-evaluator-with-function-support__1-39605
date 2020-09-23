Attribute VB_Name = "modLexStructs"
Option Explicit
'//Author: Linguar Amadala (Allen Copeland)
Public Enum EXPR_Lex_TokenType
    EXPR_L_TT_Keyword = 1
    EXPR_L_TT_Operator = 2
    EXPR_L_TT_WhiteSpace = 4
    EXPR_L_TT_Comment = 8
    EXPR_L_TT_Identifier = 16
    EXPR_L_TT_String = 64
    EXPR_L_TT_Constant = 128
    EXPR_L_TT_UnknownChar = 256
    EXPR_L_TT_LineFeed = 512
End Enum
Public Type EXPR_Lex_Token
    Value As String
    Position As Long
    Length As Long
    Type As Integer
End Type
Public Type EXPR_Lex_MethResult
    Successful As Boolean
    Token As EXPR_Lex_Token
End Type
Public Type EXPR_Lex_MethInput
    Script As String
    ScriptLength As Long
    LoopPosition As Long
End Type
Public Type EXPR_Lex_Tokens
    Count As Long
    Tokens() As EXPR_Lex_Token
End Type

Public Sub AddToken(Tokens As EXPR_Lex_Tokens, Token As EXPR_Lex_Token)
    With Tokens
        If .Count = 0 Then
            ReDim .Tokens(0 To .Count)
        Else
            ReDim Preserve .Tokens(0 To .Count + 1)
        End If
        .Tokens(.Count) = Token
        .Count = .Count + 1
    End With
End Sub
