Attribute VB_Name = "modLexMeths"
Option Explicit
'//Author: Linguar Amadala (Allen Copeland)

Public Function IsOperator(ByRef Char As String) As Boolean
    Select Case Char
        '//Select the character in the comarison
        Case m_gl_CST_strOprAnd, m_gl_CST_strOprApostrophe, m_gl_CST_strOprBackSlash, m_gl_CST_strOprComma, m_gl_CST_strOprEquals, m_gl_CST_strOprForwardSlash, m_gl_CST_strOprLeftBracket, m_gl_CST_strOprRightBracket, m_gl_CST_strOprPlus, m_gl_CST_strOprMinus, m_gl_CST_strOprPercent, m_gl_CST_strOprQuote, m_gl_CST_strOprTimes, m_gl_CST_strOprUnderscore, m_gl_CST_strOprPeriod, m_gl_CST_strOprConditional, m_gl_CST_strOprGT, m_gl_CST_strOprLT, m_gl_CST_strOprCarrot
            '//If it's any one of these operators, then...
            IsOperator = True
                '//Return true
    End Select '//Char
End Function

Public Function IsKeyword(ByRef Expression As String) As Boolean
    Select Case LCase(Expression)
        '//Select the lower case version of the expression into the comparison
        Case m_gl_CST_strKwdNot, m_gl_CST_strKwdOr, m_gl_CST_strKwdAnd, m_gl_CST_strKwdXOr, m_gl_CST_strKwdRemainder
            '//if it's any one of these keywords, then...
            IsKeyword = True
                '//return true
    End Select
End Function

Public Function LexString(ByRef Args As EXPR_Lex_MethInput) As EXPR_Lex_MethResult
    Dim m_lngPosition As Long
        '//Current position
    Dim m_strChar As String
        '//Current character
    Dim m_lmrResult As EXPR_Lex_MethResult
        '//Result struct
Try:
    '//Try block
    On Error GoTo Catch
        '//Error handling
    m_strChar = Mid(Args.Script, Args.LoopPosition, 1)
        '//Obtain first character
    If (Not (m_strChar = m_gl_CST_strOprQuote)) Then
        '//If it's not a quote, then...
        GoTo Finally
            '//We're done, it's not a string
    End If '//(Not (m_strChar = m_gl_CST_strOprQuote))
    For m_lngPosition = Args.LoopPosition + 1 To Args.ScriptLength
        '//Loop through the expression
        m_strChar = Mid(Args.Script, m_lngPosition, 1)
        If (m_strChar = m_gl_CST_strOprQuote) Then
            m_lngPosition = m_lngPosition + 1
                '//Increment the position, if the next character is
                '//another quote, it is to be ignored.
            m_strChar = Mid(Args.Script, m_lngPosition, 1)
                '//Obtain the next character
            If (Not (m_strChar = m_gl_CST_strOprQuote)) Then
                '//If the Second character isn't another quote, then...
                m_lngPosition = m_lngPosition - 1
                    '//decrement position
                GoTo Success
                    '//Assume Success
            End If
        ElseIf ((m_strChar = vbCr) Or (m_strChar = vbLf)) Then
            '//If the line ends before the end quote is hit...
            GoTo Finally
                '//Assume failure...
        End If '//(m_strChar = m_gl_CST_strOprQuote)
    Next '//[m_lngPosition]
    GoTo Finally
Success:
    m_lmrResult.Successful = True
        '//Indicate success
    With m_lmrResult.Token
        '//Select the results token's namespace
        .Position = Args.LoopPosition
            '//Set the token's position to the starting location
        .Length = (m_lngPosition - (Args.LoopPosition - 1))
            '//set the token's length
        .Value = Mid(Args.Script, .Position, .Length)
            '//Set the value
        .Type = EXPR_L_TT_String
            '//Indicate the token is a string
    End With '//m_lmrResult.Token
        '//Deselect the namespace
    GoTo Finally
        '//Go to the exit
Catch:
    
Finally:
    LexString = m_lmrResult
        '//Return the result
End Function

Public Function LexWord(ByRef Args As EXPR_Lex_MethInput) As EXPR_Lex_MethResult
    Dim m_lngPosition As Long
        '//Current loop position
    Dim m_lmrResult As EXPR_Lex_MethResult
        '//The procedure result
    Dim m_strChar As String
        '//Current character
    Dim m_lngIdxStart As Long
        '//The start index offset
Try:
    On Error GoTo Catch:
    m_strChar = Mid(Args.Script, Args.LoopPosition, 1)
        '//Get the first character
    If (Not IsAlpha(m_strChar)) Then
        '//If it's not an alphabetic character, then...
        If (m_strChar = m_gl_CST_strOprConditional) Then
            '//If the character is the conditional start character, then...
            m_strChar = Mid(Args.Script, Args.LoopPosition + 1, 1)
                '//Obtain the next character after the conditional char
            If (Not IsAlpha(m_strChar)) Then
                '//If it's not an alphabetic character, then...
                GoTo Finally
                    '//Exit, we're done
            End If '//(Not IsAlpha(m_strChar))
            m_lngIdxStart = 2
                '//Set the index start to 2, so it starts at the
                '//third character
        Else
            '//Otherwise
            GoTo Finally
                '//Exit it's not a word
        End If '//(m_strChar = m_gl_CST_strOprConditional)
    Else
        '//Otherwise
        m_lngIdxStart = 1
            '//Set the index shift to 1 to start at char 2
    End If '//(Not IsAlpha(m_strChar))
    For m_lngPosition = Args.LoopPosition + m_lngIdxStart To Args.ScriptLength
        '//Loop through the expression
        m_strChar = Mid(Args.Script, m_lngPosition, 1)
            '//Select the active character
        If (((IsOperator(m_strChar) And (Not (m_strChar = m_gl_CST_strOprUnderscore)))) Or IsWhitespace(m_strChar)) Then
            '//If the active character is an operator
            '//and the operator isn't an underscore,
            '//or it's a whitespace character, then...
            m_lngPosition = m_lngPosition - 1
                '//Decrement the position
            GoTo Success
                '//Success
        End If '//(((IsOperator(m_strChar) And (Not (m_strChar = m_gl_CST_strOprUnderscore)))) Or IsWhitespace(m_strChar))
    Next '//m_lngPosition
        '//Continue loop
    GoTo Finally
        '//Exit procedure
Success:
    m_lmrResult.Successful = True
        '//Indicate success
    With m_lmrResult.Token
        '//Select the results token's namespace
        .Position = Args.LoopPosition
            '//Set the position
        .Length = (m_lngPosition - (Args.LoopPosition - 1))
            '//Set the length
        .Value = Mid(Args.Script, .Position, .Length)
            '//Set the value
        If IsKeyword(.Value) Then
            '//If the value is a keyword, then...
            .Type = EXPR_L_TT_Keyword
                '//Indicate it's a keyword
        Else
            '//Otherwise...
            .Type = EXPR_L_TT_Identifier
                '//Indicate that it's an identifier
        End If '//IsKeyword(.Value)
    End With '//m_lmrResult.Token
    GoTo Finally
        '//Exit procedure
Catch:
    '//Error handler, just exits...
    
Finally:
    '//Try end
    LexWord = m_lmrResult
        '//Return the result
End Function

Public Function LexConst(ByRef Args As EXPR_Lex_MethInput) As EXPR_Lex_MethResult
    Dim m_lngPosition As Long
    Dim m_lmrResult As EXPR_Lex_MethResult
    Dim m_strChar As String
    Dim m_booHex As Boolean
    Dim m_booOct As Boolean
Try:
    On Error GoTo Catch:
    m_strChar = Mid(Args.Script, Args.LoopPosition, 1)
    If m_strChar = m_gl_CST_strOprAnd Then
        m_lngPosition = Args.LoopPosition + 1
        m_strChar = Mid(Args.Script, m_lngPosition, 1)
        If LCase(m_strChar) = "h" Then
            m_booHex = True
        ElseIf LCase(m_strChar) = "o" Then
            m_booOct = True
        Else
            GoTo Finally
        End If
        m_lngPosition = m_lngPosition + 1
        m_strChar = Mid(Args.Script, m_lngPosition, 1)
    ElseIf Not IsNumeric(m_strChar) Then
        GoTo Finally
    End If
    If m_booHex Then
        If Not IsHexadecimal(m_strChar) Then
            GoTo Finally
        End If
        For m_lngPosition = Args.LoopPosition + 3 To Args.ScriptLength
            m_strChar = Mid(Args.Script, m_lngPosition, 1)
            If IsOperator(m_strChar) Then
                If m_strChar = m_gl_CST_strOprAnd Then
                    GoTo Success
                Else
                    m_lngPosition = m_lngPosition - 1
                    GoTo Success
                End If
            ElseIf IsWhitespace(m_strChar) Then
                m_lngPosition = m_lngPosition - 1
                GoTo Success
            ElseIf Not IsHexadecimal(m_strChar) Then
                GoTo Finally
            End If
        Next
    ElseIf m_booOct Then
        If Not IsOctal(m_strChar) Then
            GoTo Finally
        End If
        For m_lngPosition = Args.LoopPosition + 3 To Args.ScriptLength
            m_strChar = Mid(Args.Script, m_lngPosition, 1)
            If IsOperator(m_strChar) Then
                If m_strChar = m_gl_CST_strOprAnd Then
                    GoTo Success
                Else
                    m_lngPosition = m_lngPosition - 1
                    GoTo Success
                End If
            ElseIf IsWhitespace(m_strChar) Then
                m_lngPosition = m_lngPosition - 1
                GoTo Success
            ElseIf Not IsOctal(m_strChar) Then
                GoTo Finally
            End If
        Next
    Else
        For m_lngPosition = (Args.LoopPosition + 1) To Args.ScriptLength
            m_strChar = Mid(Args.Script, m_lngPosition, 1)
            If (IsOperator(m_strChar) Or IsWhitespace(m_strChar)) And Not m_strChar = "." Then
                m_lngPosition = m_lngPosition - 1
                GoTo Success
            ElseIf Not IsNumeric(m_strChar) And Not m_strChar = "." Then
                GoTo Finally
            End If
        Next
    End If
    GoTo Finally
Success:
    m_lmrResult.Successful = True
    With m_lmrResult.Token
        .Position = Args.LoopPosition
        .Length = (m_lngPosition - (Args.LoopPosition - 1))
        .Value = Mid(Args.Script, .Position, .Length)
        If (m_booHex Or m_booOct) Then
            If (Right(.Value, 1) = m_gl_CST_strOprAnd) Then
                .Value = Left(.Value, .Length - 1)
            End If
        End If
        .Type = EXPR_L_TT_Constant
    End With
    GoTo Finally
Catch:
    
Finally:
    LexConst = m_lmrResult
End Function

Public Function LexComment(ByRef Args As EXPR_Lex_MethInput) As EXPR_Lex_MethResult
    Dim m_lngPosition As Long
        '//Loop Position
    Dim m_lmrResult As EXPR_Lex_MethResult
        '//Procedure result
    Dim m_strChar As String
        '//Current character.
Try:
    On Error GoTo Catch:
        '//Error handler
    m_strChar = Mid(Args.Script, Args.LoopPosition, 1)
        '//Get the first character
    If (Not (m_strChar = "'")) Then
        '//If the first character isn't a "'" then...
        GoTo Finally
            '//Exit, it's not a comment
    End If '//(Not (m_strChar = "'"))
    For m_lngPosition = Args.LoopPosition + 1 To Args.ScriptLength
        '//Loop through the expression
        m_strChar = Mid(Args.Script, m_lngPosition, 1)
            '//Get the active character
        If (m_strChar = m_gl_CST_strOprUnderscore) Then
            '//If the character is an underscore, then...
            m_lngPosition = m_lngPosition + 1
                '//Increment the position..
            m_strChar = Mid(Args.Script, m_lngPosition, 1)
                '//Get the current character...
            If (Not ((m_strChar = vbCr) Or (m_strChar = vbLf))) Then
                '//If the character isn't a carrage return or a line-
                '//feed, then...
                m_lngPosition = m_lngPosition - 1
                    '//Decrement the position
            Else
                '//Otherwise...
                m_lngPosition = m_lngPosition + 1
                    '//Increment the position...
                m_strChar = Mid(Args.Script, m_lngPosition, 1)
                    '//Get the current character...
                If (Not ((m_strChar = vbCr) Or (m_strChar = vbLf))) Then
                    '//If the current character isn't a carrage return
                    '//or a line-feed, then...
                    m_lngPosition = m_lngPosition - 1
                        '//Decrement the position.
                End If '//(Not ((m_strChar = vbCr) Or (m_strChar = vbLf)))
            End If '//(Not ((m_strChar = vbCr) Or (m_strChar = vbLf)))
        ElseIf m_strChar = vbCr Or m_strChar = vbLf Then
            '//If the current character is a carrage return or a line-
            '//feed, then...
            m_lngPosition = m_lngPosition - 1
                '//Decrement the position.
            GoTo Success
                '//Assume success
        End If '//(m_strChar = m_gl_CST_strOprUnderscore)
    Next '//[m_lngPosition]
    GoTo Finally
        '//Exit, it's not valid
Success:
    '//Success block
    m_lmrResult.Successful = True
        '//Indicate success.
    With m_lmrResult.Token
        '//Select the results token's namespace
        .Position = Args.LoopPosition
            '//set the position
        .Length = (m_lngPosition - (Args.LoopPosition - 1))
            '//Set the length
        .Value = Mid(Args.Script, .Position, .Length)
            '//Set the value
        .Type = EXPR_L_TT_Comment
            '//Indicate that it's a comment
        If ((InStr(1, .Value, vbCr) <> 0) Or (InStr(1, .Value, vbLf) <> 0)) Then
            '//If the value has a Carrage return or a line-feed, then...
            .Type = .Type Or EXPR_L_TT_LineFeed
                '//Indicate that it's also a linefeed or has a cr[/]lf
        End If '//((InStr(1, .Value, vbCr) <> 0) Or (InStr(1, .Value, vbLf) <> 0))
    End With '//m_lmrResult.Token
    GoTo Finally
Catch:
    
Finally:
    LexComment = m_lmrResult
        '//Return
End Function

Public Function LexWhitespace(ByRef Args As EXPR_Lex_MethInput) As EXPR_Lex_MethResult
    Dim m_lngPosition As Long
        '//Loop position
    Dim m_strChar As String
        '//Current character
    Dim m_lmrResult As EXPR_Lex_MethResult
        '//Procedure Result
    Dim m_booLineFeed As Boolean
    Dim m_booContinue As Boolean
Try:
    On Error GoTo Catch:
        '//Error handler
    m_strChar = Mid(Args.Script, Args.LoopPosition, 1)
        '//Obtain the first character
    If (Not IsWhitespace(m_strChar)) Then
        '//If the character isn't a whitespace character, then...
        GoTo Finally
            '//It's not whitespace.
    End If '//(Not IsWhitespace(m_strChar))
    For m_lngPosition = Args.LoopPosition + 1 To Args.ScriptLength
        '//Loop through the expression
        m_strChar = Mid(Args.Script, m_lngPosition, 1)
            '//Get the active character
        If (Not IsWhitespace(m_strChar)) Then
            '//If the character isn't whitespace, then...
            If (m_strChar = m_gl_CST_strOprUnderscore) Then
                '//If the character is an underscore, then...
                m_lngPosition = m_lngPosition + 1
                    '//increment the position.
                m_strChar = Mid(Args.Script, m_lngPosition, 1)
                    '//Get the active character...
                m_booContinue = True
                If (Not (m_strChar = vbCr Or m_strChar = vbLf)) Then
                    '//If the character isn't whitespace, then...
                    m_lngPosition = m_lngPosition - 2
                        '//Decrement the position
                    GoTo Success
                        '//Success
                End If '//(Not IsWhitespace(m_strChar))
            Else
                '//Otherwise
                m_lngPosition = m_lngPosition - 1
                    '//Decrement the position
                GoTo Success
                    '//Success
            End If '//(m_strChar = m_gl_CST_strOprUnderscore)
        ElseIf m_strChar = vbCr Or m_strChar = vbLf Then
            If m_booContinue Then
                m_booContinue = False
            Else
                m_booLineFeed = True
            End If
        End If '//(Not IsWhitespace(m_strChar))
    Next
    GoTo Success
        '//We're successful, this is because the lex proc adds a
        '//carrage return and line-feed combination.
Success:
    m_lmrResult.Successful = True
        '//Indicate success
    With m_lmrResult.Token
        '//Select the resluts token's namespace
        .Position = Args.LoopPosition
            '//Set the position
        .Length = (m_lngPosition - (Args.LoopPosition - 1))
            '//Set the length
        .Value = Mid(Args.Script, .Position, .Length)
            '//Set the value
        .Type = EXPR_L_TT_WhiteSpace
            '//Indicate that it's whitespace
        If m_booLineFeed Then
            '//If it has a carrage return or a line-feed, then...
            .Type = .Type Or EXPR_L_TT_LineFeed
                '//Indicate that this is also a line-feed
        End If '//((InStr(1, .Value, vbCr) <> 0) Or (InStr(1, .Value, vbLf) <> 0))
    End With '//m_lmrResult.Token
    GoTo Finally
Catch:
    
Finally:
    LexWhitespace = m_lmrResult
        '//Return
End Function

Public Function LexOperator(ByRef Args As EXPR_Lex_MethInput) As EXPR_Lex_MethResult
    Dim m_lmrResult As EXPR_Lex_MethResult
        '//Procedure result
    Dim m_strChar As String
        '//Active Character
Try:
    On Error GoTo Catch:
    m_strChar = Mid(Args.Script, Args.LoopPosition, 1)
        '//Get the first character
    If ((IsOperator(m_strChar)) And (Not (m_strChar = m_gl_CST_strOprQuote))) Then
        '//If it is an operator, and it's not a quote (because strings
        '//use it)
        With m_lmrResult
            '//Select the reuslt's namespace
            .Successful = True
                '//Indicate success
            With .Token
                '//Select the token's namespace
                .Position = Args.LoopPosition
                    '//Set the position
                .Length = 1
                    '//Set the length
                .Value = m_strChar
                    '//Set the value
                .Type = EXPR_L_TT_Operator
                    '//Indicate it's an operator
            End With '//.Token
        End With '//m_lmrResult
    End If '//((IsOperator(m_strChar)) And (Not (m_strChar = m_gl_CST_strOprQuote)))
    GoTo Finally
Catch:
    
Finally:
    LexOperator = m_lmrResult
        '//Return
End Function

Public Function LexProc(ByVal Expression As String) As EXPR_Lex_Tokens
    Dim m_ltsTokens As EXPR_Lex_Tokens
        '//Result
    Dim m_lmrResult As EXPR_Lex_MethResult
        '//Each sub-procedure's result
    Dim m_lmiInput As EXPR_Lex_MethInput
        '//Input to the sub-procedures
    m_lmiInput.Script = Expression & vbCrLf
        '//Set the script of the input arg
    m_lmiInput.ScriptLength = Len(m_lmiInput.Script)
        '//Set the script length
    For m_lmiInput.LoopPosition = 1 To Len(Expression)
        '//Loop through the length of the expression
        m_lmrResult = LexWord(m_lmiInput)
            '//Return the result from lexword
        If Not m_lmrResult.Successful Then _
            m_lmrResult = LexComment(m_lmiInput)
            '//If lex word failed, then...
                '//Get the result from LexComment
        If Not m_lmrResult.Successful Then _
            m_lmrResult = LexConst(m_lmiInput)
            '//If lex comment failed, then...
                '//Get the result from lexconst
        If Not m_lmrResult.Successful Then _
            m_lmrResult = LexOperator(m_lmiInput)
            '//If lexconst failed, then...
                '//Get the result from LexOperator
        If Not m_lmrResult.Successful Then _
            m_lmrResult = LexString(m_lmiInput)
            '//If LexOperator failed, then...
                '//Get the result from LexString
        If Not m_lmrResult.Successful Then _
            m_lmrResult = LexWhitespace(m_lmiInput)
        If (Not m_lmrResult.Successful) Then
            '//If LexWhitespace failed, then...
            With m_lmrResult.Token
                '//Select the results token's namespace...
                .Type = EXPR_L_TT_UnknownChar
                    '//Indicate unknown character
                .Position = m_lmiInput.LoopPosition
                    '//Set the position
                .Value = Mid(m_lmiInput.Script, m_lmiInput.LoopPosition, 1)
                    '//Set the value
                .Length = 1
                    '//Set the length
                AddToken m_ltsTokens, m_lmrResult.Token
                    'Add the token.
            End With '//m_lmrResult.Token
        Else
            '//Otherwise...
            AddToken m_ltsTokens, m_lmrResult.Token
                '//Add the token
            m_lmiInput.LoopPosition = m_lmiInput.LoopPosition + (m_lmrResult.Token.Length - 1)
                '//Increment the position
        End If '//(Not m_lmrResult.Successful)
    Next '//[m_lmiInput.LoopPosition]
    LexProc = m_ltsTokens
        '//Return
End Function

Public Function IsAlpha(ByRef Char As String) As Boolean
    Select Case Char
        '//Select the character
        Case "a" To "z", "A" To "Z"
            '//If it's any of the characters
            IsAlpha = True
                '//Return true
    End Select '//Char
End Function

Public Function IsWhitespace(ByRef Char As String) As Boolean
    Select Case Char
        '//Select the character
        Case " ", vbCr, vbLf, vbTab
            '//If it's one of the following characters, then...
            IsWhitespace = True
                '//Return true
    End Select '//Char
End Function

Public Function IsHexadecimal(ByRef Char As String) As Boolean
    Select Case Char
        '//Select the character
        Case "a" To "f", "A" To "F", "0" To "9"
            '//If it's any one of these characters/numbers, then...
            IsHexadecimal = True
                '//Return True
    End Select '//Char
End Function

Public Function IsOctal(ByRef Char As String) As Boolean
    Select Case Char
        '//Select the character
        Case "0" To "7"
            '//If it's one of these numbers, then...
            IsOctal = True
                '//Return true
    End Select '//Char
End Function

Public Function TokenIs(Stream As EXPR_Lex_Tokens, Position As Long, ByVal TokenType As Long, ByVal Value As String) As Boolean
    Dim m_staMembers() As Variant
    Dim m_varMember As Variant
    With Stream.Tokens(Position)
        If Value = vbNullString Then
            TokenIs = (TokenType = .Type)
        Else
            If InStr(1, Value, "|") <> 0 Then
                m_staMembers = Split(LCase(Value), "|")
                For Each m_varMember In m_staMembers
                    If m_varMember = LCase(.Value) Then
                        TokenIs = True
                    End If
                Next
            Else
                TokenIs = ((TokenType = .Type) And (LCase(Value) = LCase(.Value)))
            End If
        End If
    End With
End Function

