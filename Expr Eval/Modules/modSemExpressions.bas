Attribute VB_Name = "modSemExpressions"
    Option Explicit
'//Author: Linguar Amadala (Allen Copeland)

Private Function GetOpPrecedence(ByVal Operator As EXPR_SEM_ExprSection_SubOperation) As EXPR_SEM_ExprSection_Op_Precedence
    '//Get Operator Precedence
    '//Author: Linguar Amadala (Allen Copeland)
    '//Purpose: To return an operator's Precedence for sorting purposes
    Select Case Operator
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_InEquality
            GetOpPrecedence = EXPR_SEM_ExprSection_Op_Precedence.EXPR_SEM_ES_OP_P_BoolInequality
                '//Addition and subtraction
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Addition, EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Subtraction
            GetOpPrecedence = EXPR_SEM_ExprSection_Op_Precedence.EXPR_SEM_ES_OP_P_AddSubt
                '//Addition and subtraction
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_AddTo, EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_StrAppend
            GetOpPrecedence = EXPR_SEM_ExprSection_Op_Precedence.EXPR_SEM_ES_OP_P_Append
                '//Append
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_And
            GetOpPrecedence = EXPR_SEM_ExprSection_Op_Precedence.EXPR_SEM_ES_OP_P_StrConcatination
                '//String concatination
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_BinaryAnd
            GetOpPrecedence = EXPR_SEM_ExprSection_Op_Precedence.EXPR_SEM_ES_OP_P_LogAnd
                '//Logical And
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Equals
            GetOpPrecedence = EXPR_SEM_ExprSection_Op_Precedence.EXPR_SEM_ES_OP_P_Append
                '//Append...
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_BinaryEquals
            GetOpPrecedence = EXPR_SEM_ExprSection_Op_Precedence.EXPR_SEM_ES_OP_P_BoolEquality
                '//Equality
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_BinaryOr
            GetOpPrecedence = EXPR_SEM_ExprSection_Op_Precedence.EXPR_SEM_ES_OP_P_LogOr
                '//Logical Or
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Division, EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Multiplication
            GetOpPrecedence = EXPR_SEM_ExprSection_Op_Precedence.EXPR_SEM_ES_OP_P_MultDivide
                '//Multiplication and Division
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_GreaterThan
            GetOpPrecedence = EXPR_SEM_ExprSection_Op_Precedence.EXPR_SEM_ES_OP_P_BoolGreaterThan
                '//Greater than
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_GreaterThanOrEqualTo
            GetOpPrecedence = EXPR_SEM_ExprSection_Op_Precedence.EXPR_SEM_ES_OP_P_BoolGreaterThanOrEqualTo
                '//Greater than or equal to
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Increment
            GetOpPrecedence = EXPR_SEM_ExprSection_Op_Precedence.EXPR_SEM_ES_OP_P_Append
                '//Increment value
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_IntDivision
            GetOpPrecedence = EXPR_SEM_ExprSection_Op_Precedence.EXPR_SEM_ES_OP_P_IntegerDivide
                '//Integer Division
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_LessThan
            GetOpPrecedence = EXPR_SEM_ExprSection_Op_Precedence.EXPR_SEM_ES_OP_P_BoolLessThan
                '//Less than
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_LessThanOrEqualTo
            GetOpPrecedence = EXPR_SEM_ExprSection_Op_Precedence.EXPR_SEM_ES_OP_P_BoolLessThanOrEqualTo
                '//Less than or Equal to
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Negate
            GetOpPrecedence = EXPR_SEM_ExprSection_Op_Precedence.EXPR_SEM_ES_OP_P_Negation
                '//Unary Negate
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_PowerOf
            GetOpPrecedence = EXPR_SEM_ExprSection_Op_Precedence.EXPR_SEM_ES_OP_P_Exponentation
                '//Exponentation (^)
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_XOr
            GetOpPrecedence = EXPR_SEM_ExprSection_Op_Precedence.EXPR_SEM_ES_OP_P_LogXor
                '//Logical Exclusive Or
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Modulus
            GetOpPrecedence = EXPR_SEM_ExprSection_Op_Precedence.EXPR_SEM_ES_OP_P_Modulus
    End Select
End Function

Private Function GetPrevItem(ByRef Expression As EXPR_Sem_Expression, ByRef Section As EXPR_Sem_Expr_Section, ByVal CurrentLevel As EXPR_SEM_ExprSection_Op_Precedence, ByRef Start As Integer, ByVal Limit As Integer) As Integer
    Dim m_lngLoop As Long
    Dim m_mbmMember As EXPR_Sem_Expr_Member
    Dim m_lngLevel As EXPR_SEM_ExprSection_Op_Precedence
    Dim m_booFound As Boolean
    For m_lngLevel = CurrentLevel To EXPR_SEM_ES_OP_P_Exponentation
        '//Loop from the current level to the exponentation level
        For m_lngLoop = Start To Limit Step -1
            '//Loop from the start to the limit, backwards, since we're getting the
            '//previous item
            m_mbmMember = Section.Members.Items(m_lngLoop)
                '//Obtain the member at the current index
            If m_mbmMember.Type = EXPR_S_E_M_T_Operator Then
                '//If the member's type is an operator, then...
                If (GetOpPrecedence(CInt(Section.Operators.Items(m_mbmMember.Index).Operation)) = m_lngLevel) Then
                    '//If the operator precedence for the active member is equal to the
                    '//current level, then...
                    m_booFound = True
                        '//We've found our next item...
                    Exit For
                        '//Exit.
                End If '//(GetOpPrecedence(CInt(Section.Operators.Items(m_mbmMember.Index).Operation)) = m_lngLevel)
            End If '//If m_mbmMember.Type = EXPR_S_E_M_T_Operator Then
        Next '//[m_lngLoop]
        If m_booFound Then _
            Exit For
            '//If we've found the current item, then...
                '//Exit
    Next '//[m_lngLevel]
    If m_booFound Then
        '//If we've found our item, then...
        GetPrevItem = m_lngLoop
            '//Return the last active index
    Else
        '//Otherwise
        GetPrevItem = -1
            '//Return failure
    End If
End Function

Private Function GetNextItem(ByRef Expression As EXPR_Sem_Expression, ByRef Section As EXPR_Sem_Expr_Section, ByVal CurrentLevel As EXPR_SEM_ExprSection_Op_Precedence, ByRef Start As Integer, ByVal Limit As Integer) As Integer
    Dim m_lngLoop As Long
        '//Loop variable
    Dim m_mbmMember As EXPR_Sem_Expr_Member
        '//Semantic Expression Member
    Dim m_lngLevel As EXPR_SEM_ExprSection_Op_Precedence
        '//Operator Precedence level
    Dim m_booFound As Boolean
        '//Whether or not a member was found
    Dim m_booBack As Boolean
    Dim m_lngResult As Long
    For m_lngLevel = CurrentLevel To EXPR_SEM_ES_OP_P_Exponentation
        '//Loop from the current level to the exponentation level
        For m_lngLoop = Start To Limit
            '//Loop from the start to the limit, forwards, since we're getting the
            '//next item
            m_mbmMember = Section.Members.Items(m_lngLoop)
                '//Obtain the member at the current index
            If (m_mbmMember.Type = EXPR_S_E_M_T_Operator) Then
                '//If the member's type is an operator, then...
                If (GetOpPrecedence(CInt(Section.Operators.Items(m_mbmMember.Index).Operation)) = m_lngLevel) Then
                    '//If the operator precedence for the active member is equal to the
                    '//current level, then...
                    If GetOpPrecedence(CInt(Section.Operators.Items(m_mbmMember.Index).Operation)) = EXPR_SEM_ES_OP_P_MultDivide Then
                        m_booBack = True
                            '//Indicate that we must find the last operator in this group, since it goes in the reverse order implied.
                    End If
                    m_lngResult = m_lngLoop
                        '//Indicate the result is the current index
                    m_booFound = True
                        '//We've found our next item...
                    Exit For
                        '//Exit.
                End If '//GetOpPrecedence(CInt(Section.Operators.Items(m_mbmMember.Index).Operation)) = m_lngLevel
            End If '//(m_mbmMember.Type = EXPR_S_E_M_T_Operator)
        Next '//[m_lngLoop]
        If m_booFound Then _
            Exit For
            '//If we've found the current item, then...
                '//Exit
    Next '//[m_lngLevel]
    If m_booBack Then
        For m_lngLoop = m_lngResult To Limit
            m_mbmMember = Section.Members.Items(m_lngLoop)
            If m_mbmMember.Type = EXPR_S_E_M_T_Operator Then
                If GetOpPrecedence(Section.Operators.Items(m_mbmMember.Index).Operation) = GetOpPrecedence(Section.Operators.Items(Section.Members.Items(m_lngResult).Index).Operation) Then
                    m_lngResult = m_lngLoop
                Else
                    Exit For
                End If
            End If
        Next
    End If
    If m_booFound Then
        '//If we've found our item, then...
        GetNextItem = m_lngResult
            '//Return the last active index
    Else
        GetNextItem = -1
            '//Return failure
    End If
End Function

Private Sub AppendOperator(ByRef Intermediate As EXPR_Sem_Intermediate, ByVal Expression As String, ByRef Section As EXPR_Sem_Expr_Section, ByRef Position As Long)
    If Expression = "-" Then
        '//If the expression is a hyphen, then...
        If (Section.Operators.Items(Section.Operators.Count - 1).Operation <> EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Subtraction) Then
            '//If the last operator is anything but subtraction, then...
            AddOperator "neg", Section
                '//Add a negate operator
            Exit Sub
                '//Exit
        Else
            '//Otherwise...
            Section.Operators.Items(Section.Operators.Count - 1).Operation = EXPR_SEM_ES_SO_Addition
                '//Change the subtraction operator to addition, '--' = '+'
            Exit Sub
                '//Exit
        End If '//(Section.Operators.Items(Section.Operators.Count - 1).Operation <> EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Subtraction)
    End If
    With Section.Operators
        '//Select the Sections Operator's Namespace
        With .Items(.Count - 1)
            '//Select the item at the last index's namespace
            Select Case .Operation
                '//Select the operation value into comparison case
                Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_GreaterThan
                    '//If it's greater than, then...
                    If Expression = "=" Then
                        '//If the expression is equals, then...
                        .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_GreaterThanOrEqualTo
                            '//Append the operator to be Greaterthan or Equal to
                    ElseIf Expression = "<" Then
                        .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_InEquality
                    Else
                        '//Otherwise...
                        GoTo ErrUnexpected
                            '//Throw an exception
                    End If
                Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_LessThan
                    '//If it's less than, then...
                    If Expression = "=" Then
                        '//If the expression is equals, then...
                        .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_LessThanOrEqualTo
                            '//Make the operation Less than or Equal to
                    ElseIf Expression = ">" Then
                        .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_InEquality
                    Else
                        '//Otherwise
                        GoTo ErrUnexpected
                            '//Throw an exception
                    End If
                Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Equals
                    '//If the operation is equals, then...
                    If Expression = ">" Then
                        '//If the expression is greater than, then...
                        .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_GreaterThanOrEqualTo
                            '//Append to Greater Than or Equal To
                    ElseIf Expression = "<" Then
                        '//Otherwise, if the Expression is less than, then...
                        .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_LessThanOrEqualTo
                            '//Append to Less Than or Equal To
                    ElseIf Expression = "&" Then
                        '//If the expression is ampersand then...
                        .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_StrAppend
                            '//String append
                    ElseIf Expression = "+" Then
                        '//If the expression is addition, then...
                        .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_AddTo
                            '//Change to add to
                    ElseIf Expression = "=" Then
                        '//If the expression is Equals, then...
                        .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_BinaryEquals
                            '//Change the operation to Logical Comparison
                    Else
                        GoTo ErrUnexpected
                            '//Throw an exception
                    End If
                Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_And '//&
                    If Expression = "=" Then
                        '//If the expression is equals, then...
                        .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_StrAppend
                            '//StrAppend
                    Else
                        GoTo ErrUnexpected
                            '//Throw an exception
                    End If
                Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Addition
                    '//If the last member is addition, then...
                    If Expression = "=" Then
                        '//If the expression is equals, then...
                        .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_AddTo
                            '//Set it to add-to
                    ElseIf Expression = "+" Then
                        .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Increment
                            '//Increment
                    Else
                        '//Otherwise...
                        GoTo ErrUnexpected
                            '//Throw an exception
                    End If
                Case Else
                    GoTo ErrUnexpected
                        '//Throw an exception... This Operator is non appendable
            End Select
        End With
    End With
    GoTo Finally
ErrUnexpected:
    ThrowException Intermediate, Position, "'" & Expression & "'", EXPR_SEM_I_EXP_T_Unexpected
Finally:
    Exit Sub
End Sub

Private Function AddESCIArg(ByRef ArgList As EXPR_Sem_Expr_ComplexID_ArgList, ByRef Arg As EXPR_Sem_Expr_ComplexID_Arg) As Long
    With ArgList
        '//Select the argument list namespace
        If .Count = 0 Then
            '//If the count is zero, then...
            ReDim .Items(.Count)
                '//Redimension the array
        Else
            ReDim Preserve .Items(.Count)
                '//Redimension the array while keeping the old data stored.
        End If
        .Items(.Count) = Arg
            '//Let the last item to the argument
        AddESCIArg = .Count
            '//Return the index of the new item
        .Count = .Count + 1
            '//Increment the structure's count
    End With
End Function

Private Function AddComplexIdentifier(ByRef Section As EXPR_Sem_Expr_Section, Expression As EXPR_Sem_Expression, ByRef Identifier As EXPR_Sem_Expr_ComplexID, ByVal AddNode As Boolean) As Long
    Dim m_cinMember As EXPR_Sem_Expr_Member
        '//Expression Section Member
    With Expression.ComplexIdentifiers
        '//Select the expression's complex identifiers
        If .Count = 0 Then
            '//If the count is zero, then...
            ReDim .Items(.Count)
                '//Redimension the array
        Else
            ReDim Preserve .Items(.Count)
                '//Redimension and preserve the array
        End If
        If AddNode Then
            '//If we're to add the node...
            With m_cinMember
                '//Select the expressions sections member's namespace
                .Type = EXPR_S_E_M_T_ComplexIdentifier
                    '//Notify that it's a complex identifier
                .Index = Expression.ComplexIdentifiers.Count
                    '//Set the index to the complex identifier count
            End With
        End If
        AddComplexIdentifier = .Count
            '//Return the index of the new complex identifier
        .Items(.Count) = Identifier
            '//Set the last complex identifier to the argument identifier
        .Count = .Count + 1
            '//Increment count
    End With '//Expression.ComplexIdentifiers
    If AddNode Then
        '//If we're to add the member, then...
        AddExprSectNode Section, m_cinMember
            '//Add the node to the expression section's member list
            '//Single member arguments don't want the complex identifier added to the
            '//expression section
    End If '//[AddNode]
End Function

Private Sub AddExprSectNode(ByRef Section As EXPR_Sem_Expr_Section, ByRef Member As EXPR_Sem_Expr_Member)
    With Section.Members
        '//Select the section's members' namespace
        If (.Count = 0) Then
            '//If there are no members, then...
            ReDim .Items(.Count)
                '//Redimension the array
        Else
            '//If there are members, then...
            ReDim Preserve .Items(.Count)
                '//Redimension, while preserving existing data, the array
        End If '//(.Count = 0 )
        .Items(.Count) = Member
            '//Change the last member to the passed member
        .Count = .Count + 1
            '//Increment the count
    End With
End Sub

Private Function AddExprSection(ByRef Expression As EXPR_Sem_Expression, ByRef Section As EXPR_Sem_Expr_Section) As Long
    With Expression.Subsections
        '//Select the expression's subsections' namespace
        If .Count = 0 Then
            '//If there are no members, then...
            ReDim .Items(.Count)
                '//Initialize the array
        Else
            '//If there are members, then...
            ReDim Preserve .Items(.Count)
                '//Redimension, while preserving the data, the array to include the new
                '//member
        End If
        AddExprSection = .Count
            '//Return the index
        .Items(.Count) = Section
            '//Change the last member to the passed section
        .Count = .Count + 1
            '//Increment the count
    End With
End Function

Private Sub AddSubSection(ByRef Expression As EXPR_Sem_Expression, ByRef MainSection As EXPR_Sem_Expr_Section, ByRef SubSection As EXPR_Sem_Expr_Section)
    Dim m_cinMember As EXPR_Sem_Expr_Member
        '//Expression Section Member
    With Expression.Subsections
        '//Select the expressions subsections' namespace
        If (.Count = 0) Then
            '//If there are no sub-sections, then...
            ReDim .Items(.Count)
                '//Initialize the array
        Else
            '//If there are subsections, then...
            ReDim Preserve .Items(.Count)
                '//Redimension, while preserving the old data, the array to
                '//include the new member
        End If '//(.Count=0)
        With m_cinMember
            '//Select the section member's namespace
            .Type = EXPR_Sem_ExprSection_Member_Type.EXPR_S_E_M_T_SubExpression
                '//Set its type to indicate it's a subsection
            .Index = Expression.Subsections.Count
                '//Indicate its index to be the current last member
        End With '//m_cinMember
        .Items(.Count) = SubSection
            '//Change the new item to the subsection passed
        .Count = .Count + 1
            '//Increment the number of subsections
    End With
    AddExprSectNode MainSection, m_cinMember
        '//Add the expression section node
End Sub

Private Sub AddCIMember(ByRef ComplexIdentifier As EXPR_Sem_Expr_ComplexID, ByRef Member As EXPR_Sem_Expr_ComplexID_Member)
    '//Add Complex Identifier Member
        '//Used to bulid the complex ID structure
        '//eg. A.B.C would be 5 members, 3 string identifiers, and two member accesses
    With ComplexIdentifier
        '//Select the complexidentifier's namespace
        If .Count = 0 Then
            '//If the complex identifier is blank, then...
            ReDim .Items(.Count)
                '//Initialize the complexidentifier's array
        Else
            ReDim Preserve .Items(.Count)
                '//Redimension, while preserving previous complexid members, the array
                '//to include the new member
        End If
        .Items(.Count) = Member
            '//Change the new member to the passed member
        .Count = .Count + 1
            '//Increment the counts
    End With
End Sub

Private Sub AddCIDString(Value As String, ByRef ComplexIdentifier As EXPR_Sem_Expr_ComplexID, ByRef Intermediate As EXPR_Sem_Intermediate)
        '//Add Complex Identifier String Identifier
    Dim m_cimMember As EXPR_Sem_Expr_ComplexID_Member
    With m_cimMember
        .Index = AddString(Intermediate.ConstTable, Value, False)
            '//Change the member's Index to the index of the newly added (or existing)
            '//item
        .Type = EXPR_Sem_Expr_ComplexID_Member_Type.EXPR_S_E_CI_M_T_Identifier
            '//Indicate that it's a string identifier member
    End With '//m_cimMember
    AddCIMember ComplexIdentifier, m_cimMember
        '//Add the Complex Identifier Member to the Complex Identifier
End Sub

Private Sub AddCIMembAccess(ByRef ComplexIdentifier As EXPR_Sem_Expr_ComplexID)
    Dim m_cimMember As EXPR_Sem_Expr_ComplexID_Member
        '//Complex Identifier Member var
    With m_cimMember
        '//Select the ComplexID Member's namespace
        .Index = -1
            '//Indicate that it doesn't have an index
        .Type = EXPR_S_E_CI_M_T_SubMemberItem
            '//Indicate that it's a node to a sub-item on the item before..
    End With '//m_cimMember
    AddCIMember ComplexIdentifier, m_cimMember
        '//Add the complex identifier member to the complex identifier
End Sub

Private Sub AddCIArgList(ByRef ComplexIdentifier As EXPR_Sem_Expr_ComplexID, ByRef ArgList As EXPR_Sem_Expr_ComplexID_ArgList, ByRef Expression As EXPR_Sem_Expression)
    Dim m_cimMember As EXPR_Sem_Expr_ComplexID_Member
        '//Complex Identifier Member var
    With m_cimMember
        .Index = AddArgList(Expression, ArgList)
            '//Add the argument list and indicate that this member is owner of it.
        .Type = EXPR_Sem_Expr_ComplexID_Member_Type.EXPR_S_E_CI_M_T_ArgumentList
            '//Indicate the type is an argument list
    End With
    AddCIMember ComplexIdentifier, m_cimMember
        '//Add the complex identifier member to the complex identifier
End Sub

Private Function AddArgList(ByRef Expression As EXPR_Sem_Expression, ByRef ArgList As EXPR_Sem_Expr_ComplexID_ArgList) As Long
    With Expression.ArgumentLists
        '//Select the expression argumentlists' namespace
        If (.Count = 0) Then
            '//If there are no argument lists, then...
            ReDim .Items(.Count)
                '//Initialize the array
        Else
            '//If there are argument lists, then...
            ReDim Preserve .Items(.Count)
                '//Redimension, while preserving existing data, the array to include
                '//the indicated argument list
        End If '//(.Count = 0)
        .Items(.Count) = ArgList
            '//Change the new member to the indicated argument list
        AddArgList = .Count
            '//Return the index of the new member
        .Count = .Count + 1 '//++;
            '//Increment the count
    End With
End Function

Private Function AddOperator(ByVal Opr As String, ByRef Section As EXPR_Sem_Expr_Section)
    Dim m_esmMember As EXPR_Sem_Expr_Member
        '//Expression Section Member var
    With Section.Operators
        '//Select the sections operators' namespace
        If (.Count = 0) Then
            '//If there are no members, then...
            ReDim .Items(.Count)
                '//Initialize the array
        Else
            ReDim Preserve .Items(.Count)
                '//Redimension the array, while preserving existing members, to include
                '//space for the new member
        End If '//(.Count = 0)
        With .Items(.Count)
            '//Select the Last member's namespace
            Select Case LCase$(Opr)
                '//Select the lower case version of the operator passed
                Case "-"
                    '//If it's a minus/hyphen
                    .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Subtraction
                        '//Indicate that the operation is subtraction
                Case "+"
                    '//If it's a plus
                    .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Addition
                        '//Indicate that the operation is addition
                Case "^"
                    '//If it's a carrot
                    .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_PowerOf
                        '//Indicate that the operation is exponentation
                Case "*"
                    '//If it's an asterisk
                    .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Multiplication
                        '//Indicate that the operation is multiplication
                Case "/"
                    '//If it's a forwardslash
                    .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Division
                        '//Indicate that the operation is division
                Case "\"
                    '//If it's a backslash
                    .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_IntDivision
                        '//Indicate that the operation is integer division
                Case ">"
                    '//If it's a greater than sign
                    .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_GreaterThan
                        '//Indicate that the operation is greater than
                Case "<"
                    '//If it's a less than sign
                    .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_LessThan
                        '//Indicate that the operation is less than
                Case "="
                    '//If it's an equals
                    .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Equals
                        '//Indicate that the operation is equals
                Case m_gl_CST_strKwdOr
                    '//If it's the keyword or
                    .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_BinaryOr
                        '//Indicate that the operation is binary or
                Case m_gl_CST_strKwdNot
                    '//If it's the keyword not
                    .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_BinaryNot
                        '//Indicate that the operation is binary not
                Case m_gl_CST_strKwdAnd
                    '//If it's the keyword and
                    .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_BinaryAnd
                        '//Indicate that the operation is binary and
                Case m_gl_CST_strKwdRemainder
                    '//If it's the keyword modulus (mod)
                    .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Modulus
                        '//Indicate that the operation is modulus
                Case "&"
                    '//If it's an ampersand
                    .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_And '//string
                    '//Indicate that the operation is concatination
                Case m_gl_CST_strKwdXOr
                    '//If it's the keyword eXclusive Or (XOr)
                    .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_XOr
                        '//Indicate that the operation is Binary Exclusive Or
                Case "neg"
                    '//If it's a negate
                    .Operation = EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Negate
                        '//Indicate that the operation is negate
            End Select '//LCase$(Opr)
        End With '//.Items(.Count)
        With m_esmMember
            '//Select the namespace of the expressions sections' member
            .Type = EXPR_Sem_ExprSection_Member_Type.EXPR_S_E_M_T_Operator
                '//Indicate the type is that of an operator
            .Index = Section.Operators.Count
                '//Change the index to the operators count
        End With '//m_esmMember
        .Count = .Count + 1
            '//Increment the numer of operators
    End With
    AddExprSectNode Section, m_esmMember
        '//Add the node to the section
End Function


Private Function AddSectConstant(ByRef Section As EXPR_Sem_Expr_Section, ByRef Expression As EXPR_Sem_Expression, ByVal Value As Variant, ByRef Intermediate As EXPR_Sem_Intermediate, Optional ByVal AddNode As Boolean = True) As Long
    Dim m_esmMember As EXPR_Sem_Expr_Member
        '//Expression member
    Dim m_lngIndex As Long
        '//Index var
    With Expression.Constants
        '//Select the expression constants namespace
        If .Count = 0 Then
            '//If there are no members, then...
            ReDim .Items(.Count)
                '//Initialize the array
        Else
            '//If there are...
            ReDim Preserve .Items(.Count)
                '//Resize the array and preserve previous data for the new member
        End If '//(.Count=0)
        On Error GoTo ErrDouble
            '//If we error, add a double, most of these mod values should work
        With .Items(.Count)
            If IsNumeric(Value) Then
                Value = CDec(Value)
                '.Type = EXPR_Sem_Expr_Const_Type.EXPR_S_E_C_T_Number
                Select Case True
                    '//Heh, I'm lazy, Rather then 'if then'
                    Case CBool((Value - (Value Mod 256)) = 0)
                        '//If it is a byte
                        m_lngIndex = AddByte(Intermediate.ConstTable, Value)
                            '//Add the byte and set the index to the position in the
                            '//byte array
                        .SubType = EXPR_S_E_C_T_Byte
                            '//Indicate that it's a byte
                    Case CBool((Value - (Value Mod 256 ^ 2)) = 0)
                        m_lngIndex = AddInteger(Intermediate.ConstTable, Value)
                            '//Add the integer and set the index to the position in the
                            '//integer array
                        .SubType = EXPR_S_E_C_T_Integer
                            '//Indicate that it's an integer
                    Case CBool((Value \ 256 \ 256 - ((Value \ 256 \ 256) Mod 256 ^ 2)) = 0)
                        If InStr(1, CStr(Value), ".") = 0 Then
                            '//If there is no decimal then
                            m_lngIndex = AddLong(Intermediate.ConstTable, Value)
                                '//Add the long value and set the index to the position
                                '//in the long array
                            .SubType = EXPR_S_E_C_T_Long
                                '//Indicate that it's a long value
                        Else
                            '//If it has a decimal, then...
                            If (CDbl(CSng(Value)) <> CDbl(Value)) Then
                                '//If the accuracy between the double value and the
                                '//single value is different, then...
                                GoTo AddDbl
                                    '//Add a double value
                            Else
                                '//If it's the same...
                                m_lngIndex = AddSingle(Intermediate.ConstTable, Value)
                                    '//Add the single value and set the index to the
                                    '//position in the single array
                                .SubType = EXPR_S_E_C_T_Single
                                    '//Indicate that it's a single value
                            End If
                        End If
                    Case CBool((Value \ 256 \ 256 \ 256 \ 256 \ 256 \ 256 - ((Value \ 256 \ 256 \ 256 \ 256 \ 256 \ 256) Mod 256 ^ 2)) = 0)
                        '//This probably will error
AddDbl:
                        m_lngIndex = AddDouble(Intermediate.ConstTable, Value)
                            '//Add the double value and set the index to the position
                            '//in the double array
                        .SubType = EXPR_S_E_C_T_Double
                            '//Indicate the datatype is a double value
                    Case Else
                        
                End Select
            Else
                .SubType = EXPR_Sem_Expr_Const_Type.EXPR_S_E_C_T_String
                    '//Indicate the datatype is that of a string
                m_lngIndex = AddString(Intermediate.ConstTable, RemoveDoubleQuotes(CStr(Right(Left(Value, Len(Value) - 1), Len(Value) - 2))), True)
                    '//Add the string and set the index to the position in the
                    '//string array
            End If
            .Index = m_lngIndex
                '//Change the constant index to the index of the new (or existing) member
        End With '//.Items(.Count)
        m_esmMember.Type = EXPR_Sem_ExprSection_Member_Type.EXPR_S_E_M_T_Constant
            '//Indicate the expression section member is a constant
        m_esmMember.Index = .Count
            '//Put its index to the new member's index
        AddSectConstant = .Count
            '//Return the index of the new member
        .Count = .Count + 1
            '//Increment the count
    End With '//Expression.Constants
    If AddNode Then _
        AddExprSectNode Section, m_esmMember
        '//If we're to add the node, then...
            '//Add it to the nodes...
        '//This is optional for single member arguments, hence why addnode is used
        '//if we're not to add it, it's hanging there and referenced merely by the
        '//argument.
    GoTo Finally
ErrDouble:
    GoTo AddDbl
Finally:
    Exit Function
End Function

Public Function ParseExpression(ByRef Intermediate As EXPR_Sem_Intermediate, ByRef Stream As EXPR_Lex_Tokens, ByRef Start As Long, Optional SubExpression As Boolean = False) As EXPR_Sem_Expression
    Dim m_emeExpression As EXPR_Sem_Expression
        '//Expression result
    m_emeExpression.FirstSection = GetExpressionSection(Intermediate, Stream, Start, m_emeExpression, SubExpression)
        '//Obtain the first expression which goes down the line blocking the expression
        '//up as it goes.
    ParseExpression = m_emeExpression
        '//Return the expression
End Function

Private Function GetComplexIdentifier(ByRef Intermediate As EXPR_Sem_Intermediate, ByRef Stream As EXPR_Lex_Tokens, ByRef Section As EXPR_Sem_Expr_Section, ByRef Start As Long, ByRef Expression As EXPR_Sem_Expression) As EXPR_Sem_Expr_ComplexID
    Dim m_ltoTok As EXPR_Lex_Token
        '//Current token
    Dim m_eciIdentifier As EXPR_Sem_Expr_ComplexID
        '//Complex identifier
    Dim m_lngPosition As Long
        '//Current position
    m_lngPosition = Start
        '//Put the position to the start
    Do
        '//Loop
        m_ltoTok = Stream.Tokens(m_lngPosition)
            '//Obtain the active token
        Select Case m_ltoTok.Type
            '//Select the token's type into memory for comparison
            Case EXPR_Lex_TokenType.EXPR_L_TT_Keyword
                '//If it's a keyword, then...
                If m_eciIdentifier.Count = 0 Then
                    '//If the identifier is empty, then...
                    '//Should there be an exception here?
                    '//Probably not, since this procedure isn't called unless a valid
                    '//identifier is encountered
                    Exit Do
                        '//Exit, we're not to do anything
                Else
                    Exit Do
                End If
            Case EXPR_Lex_TokenType.EXPR_L_TT_Identifier
                '//If it's an identifier
                AddCIDString m_ltoTok.Value, m_eciIdentifier, Intermediate
                    '//Add the identifier to the complex identifier
            Case EXPR_Lex_TokenType.EXPR_L_TT_Operator
                '//If we're working with an operator, then...
                Select Case m_ltoTok.Value
                    '//Select the token's value into comparison
                    Case "("
                        '//If we're working with a left bracket
                        m_lngPosition = m_lngPosition + 1
                            '//Increment the position
                        AddCIArgList m_eciIdentifier, GetArgumentList(Intermediate, Stream, Section, Expression, m_lngPosition), Expression
                            '//Obtain the argument list following and add it to
                            '//the complex identifier
                        If (Not (m_lngPosition > Stream.Count)) Then
                            If (Not (TokenIs(Stream, m_lngPosition, EXPR_L_TT_Operator, ")"))) Then
                                '//If the next token isn't a end brace, then we have a problem...
                                ThrowException Intermediate, m_lngPosition, "expression", EXPR_SEM_I_EXP_T_Expected
                                    '//Indicate the error
                            End If '//(Not (TokenIs(Stream, m_lngPosition, EXPR_L_TT_Operator, ")")))
                        Else
                            ThrowException Intermediate, m_lngPosition, "expression", EXPR_SEM_I_EXP_T_Expected
                        End If '//(Not (m_lngPosition > Stream.Count))
                    Case "."
                        '//If it's a period, then...
                        AddCIMembAccess m_eciIdentifier
                            '//Add a member access
                    Case Else
                        '//Otherwise...
                        m_lngPosition = m_lngPosition - 1
                            '//Decrement, we're not supposed to continue
                        Exit Do
                            '//Exit
                End Select '//m_ltoTok.Value
            Case EXPR_Lex_TokenType.EXPR_L_TT_WhiteSpace
                '//If we're working with whitespace
                m_lngPosition = m_lngPosition + 1
                    '//Increment
                If (Not (m_lngPosition >= Stream.Count)) Then
                    '//If we're at or over the number of items, then...
                    m_ltoTok = GetToken(m_lngPosition, Stream)
                        '//Obtain the token at the current index...
                    If (TokenIs(Stream, m_lngPosition, EXPR_L_TT_Operator, "(")) Then
                        '//If the token after the
                        m_lngPosition = m_lngPosition - 1
                            '//Decrement, we don't want to skip the arg list
                    Else
                        '//Otherwise...
                        m_lngPosition = m_lngPosition - 1
                            '//Decrement, Complex identifiers don't have spaces
                        Exit Do
                            '//Exit
                    End If '//(TokenIs(Stream, m_lngPosition, EXPR_L_TT_Operator, "("))
                End If '//(Not (m_lngPosition >= Stream.Count))
            Case EXPR_Lex_TokenType.EXPR_L_TT_WhiteSpace Or EXPR_Lex_TokenType.EXPR_L_TT_LineFeed
                '//If it's the end of the line....
                m_lngPosition = m_lngPosition - 1
                    '//It's the end of the line...
                Exit Do
                    '//Exit
        End Select '//m_ltoTok.Type
        m_lngPosition = m_lngPosition + 1
            '//Increment
    Loop Until m_lngPosition >= Stream.Count
        '//If the position is greater or equal to the stream count
    GetComplexIdentifier = m_eciIdentifier
        '//Return the complex identifier
    Start = m_lngPosition
        '//Change the start
End Function

Private Function GetArgumentList(ByRef Intermediate As EXPR_Sem_Intermediate, ByRef Stream As EXPR_Lex_Tokens, ByRef Section As EXPR_Sem_Expr_Section, ByRef Expression As EXPR_Sem_Expression, ByRef Start As Long) As EXPR_Sem_Expr_ComplexID_ArgList
    Dim m_lngPosition As Long
        '//Current position
    Dim m_ltoTok As EXPR_Lex_Token
        '//Current token
    Dim m_lasArgs As EXPR_Sem_Expr_ComplexID_ArgList
        '//Argument list result
    Dim m_booBegin As Boolean
        '//Beginning
    m_booBegin = True
        '//We're beginning
    m_lngPosition = Start
        '//Set the initial position
    Do
        '//Loop
        m_ltoTok = Stream.Tokens(m_lngPosition)
            '//Obtain the active token
        Select Case m_ltoTok.Type
            '//Select the token's type
            Case EXPR_Lex_TokenType.EXPR_L_TT_Operator
                '//If it is an operator, then...
                Select Case m_ltoTok.Value
                    '//Select the token's value
                    Case ")"
                        '//If it's the end of the list, then...
                        Exit Do
                            '//Exit
                    Case ","
                        '//Next argument...
                        If (m_booBegin And (m_lngPosition = Start)) Then
                            '//If we're just beginning and the position is the start, then...
                            m_booBegin = False
                                '//We're no longer beginning
                            AddESCIArg m_lasArgs, GetArgument(Intermediate, Expression, Section, Stream, m_lngPosition)
                                '//Add the first arg (empty)
                            m_lngPosition = m_lngPosition + 1
                                '//Increment position
                            AddESCIArg m_lasArgs, GetArgument(Intermediate, Expression, Section, Stream, m_lngPosition)
                                '//Add the next argument
                        Else
                            '//Otherwise
                            m_lngPosition = m_lngPosition + 1
                                '//Increment
                            m_booBegin = False
                                '//We're not beginning, in case we were still
                            AddESCIArg m_lasArgs, GetArgument(Intermediate, Expression, Section, Stream, m_lngPosition)
                                '//Add the next argument
                        End If '//(m_booBegin And (m_lngPosition = Start))
                    Case "."
                        '//If we encounter a period, then they're using a with statement, so this is valid
                        AddESCIArg m_lasArgs, GetArgument(Intermediate, Expression, Section, Stream, m_lngPosition)
                            '//Add the argument
                    Case "-", "("
                        '//If we encounter negate or sub expression begin, then...
                        If (m_booBegin) Then
                            '//If we're still beginning, then...
                            m_booBegin = False
                                '//We're no longer beginning
                            AddESCIArg m_lasArgs, GetArgument(Intermediate, Expression, Section, Stream, m_lngPosition)
                                '//Add the next argument
                        End If '//(m_booBegin)
                End Select '//m_ltoTok.Value
            Case EXPR_Lex_TokenType.EXPR_L_TT_Keyword
                '//If it's a keyword, then...
                Select Case LCase(m_ltoTok.Value)
                    '//Select the token's value
                    Case m_gl_CST_strKwdNot
                        '//If it's a not keyword then...
                        If (m_booBegin) Then
                            '//If we're beginning, then...
                            m_booBegin = False
                                '//Indicate we're not
                            AddESCIArg m_lasArgs, GetArgument(Intermediate, Expression, Section, Stream, m_lngPosition)
                                '//Add the arg
                        End If '//(m_booBegin)
                End Select '//LCase(m_ltoTok.Value)
            Case EXPR_Lex_TokenType.EXPR_L_TT_Identifier, EXPR_L_TT_String, EXPR_L_TT_Constant
                '//If it's an identifier, string or constant
                AddESCIArg m_lasArgs, GetArgument(Intermediate, Expression, Section, Stream, m_lngPosition)
                    '//Add the argument
            Case EXPR_Lex_TokenType.EXPR_L_TT_WhiteSpace Or EXPR_Lex_TokenType.EXPR_L_TT_LineFeed
                '//If it's the end of the line
                ThrowException Intermediate, m_lngPosition, m_gl_CST_strExc_LstSepOrEBR, EXPR_SEM_I_EXP_T_Expected
                    '//Throw an exception, they should have more here
                Exit Do
                    '//Exit loop
            Case EXPR_Lex_TokenType.EXPR_L_TT_WhiteSpace
                '//If it's whitespace, then...
                If (m_booBegin) Then
                    '//If we're beginning, then...
                    AddESCIArg m_lasArgs, GetArgument(Intermediate, Expression, Section, Stream, m_lngPosition)
                        '//Add argument
                    m_booBegin = False
                        '//We're not beginning
                End If '//(m_booBegin)
        End Select '//m_ltoTok.Type
        m_lngPosition = m_lngPosition + 1
    Loop Until m_lngPosition >= Stream.Count '//Loop until the last item
    With m_lasArgs
        '//Select the argument list's namespace
        If (.Count > 0) Then
            '//If the count is greater then zero, then...
            With .Items(.Count - 1)
                '//Select the last member's namespace
                If (.Type = EXPR_Sem_ExprSection_CI_DA_Type.EXPR_SEM_CI_DA_T_Empty) Then
                    '//If it's empty, then...
                    ThrowException Intermediate, m_lngPosition, "expression", EXPR_SEM_I_EXP_T_Expected
                        '//Throw an exception, there should be an expression
                End If '//(.Type = EXPR_Sem_ExprSection_CI_DA_Type.EXPR_SEM_CI_DA_T_Empty)
            End With '//.Items(.Count - 1)
        End If '//(.Count > 0)
    End With '//m_lasArgs
    Start = m_lngPosition
        '//Update the position
    GetArgumentList = m_lasArgs
        '//Return
End Function

Private Function GetExpressionSection(ByRef Intermediate As EXPR_Sem_Intermediate, ByRef Stream As EXPR_Lex_Tokens, ByRef Start As Long, ByRef Expression As EXPR_Sem_Expression, Optional ByVal SubSection As Boolean = True) As EXPR_Sem_Expr_Section
    Dim m_lngPosition As Long
    Dim m_ltoTok As EXPR_Lex_Token
    Dim m_mesSec As EXPR_Sem_Expr_Section
    Dim m_lngBracketDepth As Long
    m_lngPosition = Start
    If Stream.Count = 0 Then Exit Function
    Do
        If m_lngPosition >= Stream.Count Then Exit Do
        m_ltoTok = Stream.Tokens(m_lngPosition)
        Select Case m_ltoTok.Type
            Case EXPR_Lex_TokenType.EXPR_L_TT_Operator
                Select Case m_ltoTok.Value
                    Case "("
                        If m_mesSec.Members.Count > 0 Then
                            If m_mesSec.Members.Items(m_mesSec.Members.Count - 1).Type = EXPR_Sem_ExprSection_Member_Type.EXPR_S_E_M_T_Constant Then
                                ThrowException Intermediate, m_lngPosition, m_gl_CST_strExc_EndStatement, EXPR_SEM_I_EXP_T_Expected
                                Exit Function
                            End If
                        End If
                        m_lngPosition = m_lngPosition + 1
                        AddSubSection Expression, m_mesSec, GetExpressionSection(Intermediate, Stream, m_lngPosition, Expression)
                        If (Not (m_lngPosition >= Stream.Count)) Then
                            '//If the position is in range, then...
                            If Not TokenIs(Stream, m_lngPosition, EXPR_L_TT_Operator, ")") Then
                                ThrowException Intermediate, m_lngPosition - 1, m_gl_CST_strExc_EBR, EXPR_SEM_I_EXP_T_Expected
                            End If
                        Else
                            '//If it's not in range
                            ThrowException Intermediate, m_lngPosition - 1, m_gl_CST_strExc_EBR, EXPR_SEM_I_EXP_T_Expected
                            
                        End If '//(Not (m_lngPosition >= Stream.Count))
                    Case ")", ","
                        If ((m_ltoTok.Value = ")") And (Not (SubSection))) Then
                            ThrowException Intermediate, m_lngPosition, m_gl_CST_strExc_EBR, EXPR_SEM_I_EXP_T_Unexpected
                            Exit Do
                        End If
                        'm_lngPosition = m_lngPosition - 1
                        Exit Do
                    Case "+", "-", "^", "/", "\", "*", ">", "<", "=", "&"
                        If m_mesSec.Members.Count > 0 Then
                            If m_mesSec.Members.Items(m_mesSec.Members.Count - 1).Type = EXPR_Sem_ExprSection_Member_Type.EXPR_S_E_M_T_Operator Then
                                AppendOperator Intermediate, m_ltoTok.Value, m_mesSec, m_lngPosition
                            Else
                                AddOperator m_ltoTok.Value, m_mesSec
                            End If
                        ElseIf m_ltoTok.Value = "-" Then
                            AddOperator "neg", m_mesSec
                        Else
                            ThrowException Intermediate, m_lngPosition, "operator", EXPR_SEM_I_EXP_T_Unexpected
                        End If
                    Case "."
                        GoTo GetIdentifier
                    Case Else
                        
                End Select
            Case EXPR_Lex_TokenType.EXPR_L_TT_String, EXPR_L_TT_Constant
                With m_mesSec.Members
                    If .Count > 0 Then
                        If .Items(.Count - 1).Type = EXPR_S_E_M_T_Operator Then
                            AddSectConstant m_mesSec, Expression, m_ltoTok.Value, Intermediate
                        Else
                            If SubSection Then
                                ThrowException Intermediate, m_lngPosition, m_gl_CST_strExc_LstSepOrEBR, EXPR_SEM_I_EXP_T_Expected
                            Else
                                ThrowException Intermediate, m_lngPosition, m_gl_CST_strExc_EndStatement, EXPR_SEM_I_EXP_T_Expected
                            End If
                        End If
                    Else
                        AddSectConstant m_mesSec, Expression, m_ltoTok.Value, Intermediate
                    End If
                End With
            Case EXPR_Lex_TokenType.EXPR_L_TT_Keyword
                Select Case LCase$(m_ltoTok.Value)
                    Case m_gl_CST_strKwdXOr, m_gl_CST_strKwdNot, m_gl_CST_strKwdAnd, m_gl_CST_strKwdOr, m_gl_CST_strKwdRemainder
                        AddOperator m_ltoTok.Value, m_mesSec
                    Case Else
                        Exit Do
                End Select
            Case EXPR_Lex_TokenType.EXPR_L_TT_Identifier
GetIdentifier:
                With m_mesSec.Members
                    If .Count > 0 Then
                        'Debug.Print .Items(.Count - 1).Type
                        If .Items(.Count - 1).Type = EXPR_S_E_M_T_Operator Then
                            AddComplexIdentifier m_mesSec, Expression, GetComplexIdentifier(Intermediate, Stream, m_mesSec, m_lngPosition, Expression), True
                        Else
                            If SubSection Then
                                ThrowException Intermediate, m_lngPosition, m_gl_CST_strExc_LstSepOrEBR, EXPR_SEM_I_EXP_T_Expected
                            Else
                                ThrowException Intermediate, m_lngPosition, m_gl_CST_strExc_EndStatement, EXPR_SEM_I_EXP_T_Expected
                            End If
                        End If
                    Else
                        AddComplexIdentifier m_mesSec, Expression, GetComplexIdentifier(Intermediate, Stream, m_mesSec, m_lngPosition, Expression), True
                    End If
                End With
            Case EXPR_Lex_TokenType.EXPR_L_TT_WhiteSpace
                '//Ignore
            Case EXPR_Lex_TokenType.EXPR_L_TT_WhiteSpace Or EXPR_Lex_TokenType.EXPR_L_TT_LineFeed
                If SubSection Then
                    ThrowException Intermediate, m_lngPosition, m_gl_CST_strExc_EBR, EXPR_SEM_I_EXP_T_Expected
                End If
                Exit Do
        End Select
        m_lngPosition = m_lngPosition + 1
    Loop Until m_lngPosition >= Stream.Count
    If Not m_mesSec.Members.Count = 0 Then
        With m_mesSec.Members
            If .Items(.Count - 1).Type = EXPR_S_E_M_T_Operator Then
                If Not m_mesSec.Operators.Items(.Items(.Count - 1).Index).Operation = EXPR_SEM_ES_SO_Increment Then
                    ThrowException Intermediate, m_lngPosition, "expression", EXPR_SEM_I_EXP_T_Expected
                End If
            End If
        End With
    End If
    Start = m_lngPosition
    ReorderExpSect Expression, Intermediate, m_lngPosition, m_mesSec
    GetExpressionSection = m_mesSec
End Function

Private Function GetArgument(ByRef Intermediate As EXPR_Sem_Intermediate, ByRef Expression As EXPR_Sem_Expression, ByRef Section As EXPR_Sem_Expr_Section, ByRef Stream As EXPR_Lex_Tokens, ByRef Start As Long) As EXPR_Sem_Expr_ComplexID_Arg
    Dim m_lngPosition As Long
    Dim m_ltoTok As EXPR_Lex_Token
    Dim m_lngNum As Long
    Dim m_lngBraceDepth As Long
    Dim m_lngType As EXPR_Sem_ExprSection_CI_DA_Type
    Dim m_lngMemberStart As Long
    Dim m_lngMemberCount As Long
    Dim m_booFirstOpr As Boolean
    Dim m_elaArg As EXPR_Sem_Expr_ComplexID_Arg
    m_lngPosition = Start
    m_lngType = -1
    Do
        m_ltoTok = GetToken(m_lngPosition, Stream)
        'Debug.Print m_ltoTok.Value
        Select Case m_ltoTok.Type
            Case EXPR_Lex_TokenType.EXPR_L_TT_WhiteSpace
                '//Ignore... sometimes
                If m_lngType = EXPR_SEM_CI_DA_T_ComplexIdentifier Then
                    m_lngPosition = m_lngPosition + 1
                    If m_lngPosition < Stream.Count Then
                        If Not TokenIs(Stream, m_lngPosition, EXPR_L_TT_Operator, "(") Then
                            m_lngMemberCount = m_lngMemberCount + 1
                        End If
                        m_lngPosition = m_lngPosition - 1
                    End If
                Else
                    
                End If
            Case EXPR_Lex_TokenType.EXPR_L_TT_Identifier
                If m_lngType = -1 Then
                    m_lngType = EXPR_SEM_CI_DA_T_ComplexIdentifier
                    m_lngMemberCount = m_lngMemberCount + 1
                    m_lngMemberStart = m_lngPosition
                End If
            Case EXPR_Lex_TokenType.EXPR_L_TT_Keyword
                If (Not (m_lngType = EXPR_SEM_CI_DA_T_ComplexIdentifier)) Then
                    m_lngMemberCount = m_lngMemberCount + 1
                    m_booFirstOpr = True
                Else
                    m_lngMemberCount = m_lngMemberCount + 1
                End If
            Case EXPR_Lex_TokenType.EXPR_L_TT_Constant, EXPR_L_TT_String
                If m_lngMemberCount = 0 Then
                    m_lngType = EXPR_SEM_CI_DA_T_Constant
                End If
                m_lngMemberCount = m_lngMemberCount + 1
                m_lngMemberStart = m_lngPosition
            Case EXPR_Lex_TokenType.EXPR_L_TT_Operator
                If m_lngMemberCount = 1 Then
                    If m_booFirstOpr = True Then
                        m_lngMemberCount = m_lngMemberCount + 1
                    End If
                    Select Case m_ltoTok.Value
                        Case "("
                            m_lngPosition = m_lngPosition + 1
                            Do
                                m_ltoTok = GetToken(m_lngPosition, Stream)
                                Select Case m_ltoTok.Type
                                    Case EXPR_L_TT_Operator
                                        Select Case m_ltoTok.Value
                                            Case "("
                                                m_lngBraceDepth = m_lngBraceDepth + 1
                                            Case ")"
                                                If m_lngBraceDepth = 0 Then
                                                    Exit Do
                                                Else
                                                    m_lngBraceDepth = m_lngBraceDepth - 1
                                                End If
                                        End Select
                                End Select
                                m_lngPosition = m_lngPosition + 1
                            Loop Until m_lngPosition >= Stream.Count
                        Case ",", ")"
                            m_lngPosition = m_lngPosition - 1
                            Exit Do
                        Case Else
                            m_lngMemberCount = m_lngMemberCount + 1
                    End Select
                Else
                    Select Case m_ltoTok.Value
                        Case "."
                            m_lngMemberCount = m_lngMemberCount + 1
                            m_lngMemberStart = m_lngPosition
                            m_lngType = EXPR_SEM_CI_DA_T_ComplexIdentifier
                        Case "-", "("
                            m_lngMemberCount = 2
                            Exit Do
                        Case ")", ","
                            Exit Do
                    End Select
                End If
            Case EXPR_Lex_TokenType.EXPR_L_TT_WhiteSpace Or EXPR_Lex_TokenType.EXPR_L_TT_LineFeed
                ThrowException Intermediate, m_lngPosition, m_gl_CST_strExc_LstSepOrEBR, EXPR_SEM_I_EXP_T_Expected
                Exit Do
        End Select
        m_lngPosition = m_lngPosition + 1
    Loop Until m_lngPosition >= Stream.Count
    If m_lngMemberCount > 1 Then
        m_elaArg.Type = EXPR_Sem_ExprSection_CI_DA_Type.EXPR_SEM_CI_DA_T_SubSection
        m_lngPosition = Start
        m_elaArg.Index = AddExprSection(Expression, GetExpressionSection(Intermediate, Stream, m_lngPosition, Expression))
        m_lngPosition = m_lngPosition - 1
        Start = m_lngPosition
    ElseIf m_lngMemberCount = 0 Then
        m_elaArg.Type = EXPR_Sem_ExprSection_CI_DA_Type.EXPR_SEM_CI_DA_T_Empty
        If m_lngPosition = Start Then
            Start = m_lngPosition - 1
        End If
    Else
        If m_lngType = EXPR_SEM_CI_DA_T_ComplexIdentifier Then
            m_elaArg.Type = EXPR_Sem_ExprSection_CI_DA_Type.EXPR_SEM_CI_DA_T_ComplexIdentifier
            m_elaArg.Index = AddComplexIdentifier(Section, Expression, GetComplexIdentifier(Intermediate, Stream, Section, m_lngMemberStart, Expression), False)
        ElseIf m_lngType = EXPR_SEM_CI_DA_T_Constant Then
            m_elaArg.Type = EXPR_Sem_ExprSection_CI_DA_Type.EXPR_SEM_CI_DA_T_Constant
            m_elaArg.Index = AddSectConstant(Section, Expression, Stream.Tokens(m_lngMemberStart).Value, Intermediate, False)
        End If
        Start = m_lngPosition
    End If
    GetArgument = m_elaArg
End Function

Private Sub ReorderExpSect(ByRef Expression As EXPR_Sem_Expression, Intermediate As EXPR_Sem_Intermediate, ByRef Position As Long, ByRef Section As EXPR_Sem_Expr_Section)
    Dim m_colStack As Collection
    Dim m_nodNode As EXPR_Sem_Expr_Member
    Dim m_intNode As Integer
    Dim m_lngNodType As Long
    Dim m_oprOperator As EXPR_Sem_Expr_Operator
    If Section.Members.Count = 0 Then
        ThrowException Intermediate, Position, "expression", EXPR_SEM_I_EXP_T_Expected
        Exit Sub
    End If
    m_intNode = GetPrevItem(Expression, Section, EXPR_SEM_ES_OP_P_Append, Section.Members.Count - 1, 0)
    If Not m_intNode = -1 Then
        m_lngNodType = Section.Operators.Items(Section.Members.Items(m_intNode).Index).Operation
        Section.FirstMember = m_intNode
        ProcessOperator m_lngNodType, Section, Expression, m_intNode, 0, Section.Members.Count - 1
    End If
End Sub

Public Function RemoveDoubleQuotes(Expression As String) As String
    RemoveDoubleQuotes = Replace(Expression, """""", """")
End Function

Private Function ProcessOperator(ByVal OperatorType As EXPR_SEM_ExprSection_SubOperation, ByRef Section As EXPR_Sem_Expr_Section, ByRef Expression As EXPR_Sem_Expression, ByRef Start As Integer, ByVal ALimit As Integer, ByVal BLimit As Integer, Optional ByVal HangingUnary As Boolean, Optional ByVal UnaryIdx As Integer) As Integer
    Dim m_lngIDTA As EXPR_SEM_ExprSection_OpIdxMeaningsA
    Dim m_lngIDTB As EXPR_SEM_ExprSection_OpIdxMeaningsB
    Dim m_lngIndexA As Integer
    Dim m_lngIndexB As Integer
    Select Case OperatorType
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Negate, EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_BinaryNot
            m_lngIndexA = GetNextItem(Expression, Section, GetOpPrecedence(OperatorType), Start + 1, BLimit)
            m_lngIndexB = GetPrevItem(Expression, Section, GetOpPrecedence(OperatorType), Start - 1, ALimit)
            If m_lngIndexB = -1 Then
                m_lngIDTB = EXPR_SEM_ES_OIMs_IndexB_Null
            Else
                m_lngIDTB = EXPR_SEM_ES_OIMs_IndexB_Expression
                ProcessOperator CInt(Section.Operators.Items(Section.Members.Items(m_lngIndexB).Index).Operation), Section, Expression, m_lngIndexB, ALimit, Start - 1, True, Start
            End If
            If m_lngIndexA = -1 Then
                m_lngIDTA = EXPR_SEM_ES_OIMs_IndexA_Operand
                m_lngIndexA = Start + 1
            Else
                m_lngIDTA = EXPR_SEM_ES_OIMs_IndexA_Operator
                ProcessOperator CInt(Section.Operators.Items(Section.Members.Items(m_lngIndexA).Index).Operation), Section, Expression, m_lngIndexA, Start + 1, BLimit
            End If
            Section.Operators.Items(Section.Members.Items(Start).Index).IndexA = m_lngIndexA
            Section.Operators.Items(Section.Members.Items(Start).Index).IndexB = m_lngIndexB
            Section.Operators.Items(Section.Members.Items(Start).Index).IndexAMeanings = m_lngIDTA
            Section.Operators.Items(Section.Members.Items(Start).Index).IndexBMeanings = m_lngIDTB
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Increment
            m_lngIndexA = GetPrevItem(Expression, Section, GetOpPrecedence(OperatorType), Start - 1, ALimit)
            m_lngIndexB = GetNextItem(Expression, Section, GetOpPrecedence(OperatorType), Start + 1, BLimit)
            If m_lngIndexB = -1 Then
                m_lngIDTB = EXPR_SEM_ES_OIMs_IndexB_Null
            Else
                m_lngIDTB = EXPR_SEM_ES_OIMs_IndexB_Expression
                ProcessOperator CInt(Section.Operators.Items(Section.Members.Items(m_lngIndexB).Index).Operation), Section, Expression, m_lngIndexB, ALimit, Start - 1, True, Start
            End If
            If m_lngIndexA = -1 Then
                m_lngIDTA = EXPR_SEM_ES_OIMs_IndexA_Operand
                m_lngIndexA = Start - 1
            Else
                m_lngIDTA = EXPR_SEM_ES_OIMs_IndexA_Operator
                ProcessOperator CInt(Section.Operators.Items(Section.Members.Items(m_lngIndexA).Index).Operation), Section, Expression, m_lngIndexA, Start + 1, BLimit
            End If
            Section.Operators.Items(Section.Members.Items(Start).Index).IndexA = m_lngIndexA
            Section.Operators.Items(Section.Members.Items(Start).Index).IndexB = m_lngIndexB
            Section.Operators.Items(Section.Members.Items(Start).Index).IndexAMeanings = m_lngIDTA
            Section.Operators.Items(Section.Members.Items(Start).Index).IndexBMeanings = m_lngIDTB
        Case Else
            m_lngIndexA = GetPrevItem(Expression, Section, GetOpPrecedence(OperatorType), Start - 1, ALimit)
            If m_lngIndexA = -1 Then
                m_lngIDTA = EXPR_SEM_ES_OIMs_IndexA_Operand
                m_lngIndexA = Start - 1
            Else
                m_lngIDTA = EXPR_SEM_ES_OIMs_IndexA_Operator
                ProcessOperator CInt(Section.Operators.Items(Section.Members.Items(m_lngIndexA).Index).Operation), Section, Expression, m_lngIndexA, ALimit, Start - 1
            End If
            Section.Operators.Items(Section.Members.Items(Start).Index).IndexA = m_lngIndexA
            If HangingUnary Then
                m_lngIndexB = UnaryIdx
                m_lngIDTB = EXPR_SEM_ES_OIMs_IndexB_Operand
            Else
                m_lngIndexB = GetNextItem(Expression, Section, GetOpPrecedence(OperatorType), Start + 1, BLimit)
                If m_lngIndexB = -1 Then
                    m_lngIDTB = EXPR_SEM_ES_OIMs_IndexB_Operand
                    m_lngIndexB = Start + 1
                Else
                    Select Case Section.Operators.Items(Section.Members.Items(m_lngIndexB).Index).Operation
                        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_PowerOf, EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Equals
                            m_lngIndexB = GetPrevItem(Expression, Section, GetOpPrecedence(OperatorType), BLimit, Start + 1)
                    End Select
                    m_lngIDTB = EXPR_SEM_ES_OIMs_IndexB_Operator
                    ProcessOperator CInt(Section.Operators.Items(Section.Members.Items(m_lngIndexB).Index).Operation), Section, Expression, m_lngIndexB, Start + 1, BLimit
                End If
            End If
            Section.Operators.Items(Section.Members.Items(Start).Index).IndexB = m_lngIndexB
            Section.Operators.Items(Section.Members.Items(Start).Index).IndexAMeanings = m_lngIDTA
            Section.Operators.Items(Section.Members.Items(Start).Index).IndexBMeanings = m_lngIDTB
    End Select
End Function
