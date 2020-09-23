Attribute VB_Name = "modInterpret"
Option Explicit
'//Author: Linguar Amadala (Allen Copeland)
'//Interpreter module
'//Purpose: -- To evaluate and interpret expression values/functions
Public m_gl_X As Single, m_gl_Y As Single, m_gl_Z As Single, m_gl_W As Single
    '//Base level variables, faster then user defined because of method used
Public m_colNames As Collection
    '//Names collection
Private m_varReturn As Variant
    '//Return variable used for multiple expression statements
Public m_ivsVars As EXPR_Sem_IntVars
    '//Variables member
Public Type EXPR_Sem_IntVar
    '//Interpreted variable
    Value As Variant
        '//Its value
End Type '//EXPR_Sem_IntVar
Public Type EXPR_Sem_IntVars
    Count As Integer
        '//Number of members in the array
    Items() As EXPR_Sem_IntVar
        '//Array of members
End Type '//EXPR_Sem_IntVars
Public Type EXPR_SEM_Expressions
    '//Expressions struct, used for multi-line expressions
    Count As Integer
        '//Number of expressions
    Items() As EXPR_Sem_Expression
        '//Expressions
End Type

Public Function GetExceptionStr(Intermediate As EXPR_Sem_Intermediate, Exception As EXPR_SEM_Intermediate_Exception) As String
    If Exception.Error = EXPR_SEM_I_EXP_T_Expected Then
        '//If the error was expected: %1, then...
        GetExceptionStr = Replace("Expected: %1", "%1", CStr(Intermediate.ConstTable.Strings.Items(Exception.Reason).Bytes))
            '//Return the exception string
    ElseIf Exception.Error = EXPR_SEM_I_EXP_T_Unexpected Then
        '//If the error was unexpected: %1, then...
        GetExceptionStr = Replace("Unexpected: %1", "%1", CStr(Intermediate.ConstTable.Strings.Items(Exception.Reason).Bytes))
            '//Return the exception string
    ElseIf Exception.Error = EXPR_SEM_I_EXP_T_Missing Then
        '//If the error was missing: %1, then...
        GetExceptionStr = Replace("Missing: %1", "%1", CStr(Intermediate.ConstTable.Strings.Items(Exception.Reason).Bytes))
            '//Return the exception string
    End If
End Function

Public Function Evaluate(ByRef Expression As EXPR_Sem_Expression, ByRef Intermediate As EXPR_Sem_Intermediate) As Variant
    If Intermediate.Exceptions.Count > 0 Then _
        Exit Function
        '//If there are any errors, exit
    Evaluate = EvaluateSection(Expression, Intermediate, Expression.FirstSection)
        '//Evaluate the primary section
End Function

Public Function EvaluateSection(ByRef Expression As EXPR_Sem_Expression, ByRef Intermediate As EXPR_Sem_Intermediate, ByRef Section As EXPR_Sem_Expr_Section)
    With Section
        '//Select the section's namespace
        If .Members.Count > 0 Then
            '//If members exist
            If .Members.Items(.FirstMember).Type = EXPR_S_E_M_T_Operator Then
                '//If the first member is an operator, then...
                EvaluateSection = EvalOperator(.Members.Items(.FirstMember), .FirstMember, Expression, Section, Intermediate)
                    '//Evaluate the operator
            Else
                EvaluateSection = EvalOperand(.Members.Items(.FirstMember), Expression, Section, Intermediate)
                    '//Evaluate the operand
            End If
        End If
    End With
End Function

Public Function EvalOperator(Node As EXPR_Sem_Expr_Member, ByVal NodeIndex As Integer, ByRef Expression As EXPR_Sem_Expression, ByRef Section As EXPR_Sem_Expr_Section, ByRef Intermediate As EXPR_Sem_Intermediate, Optional EvalUnary As Boolean = False, Optional EvalUnIdx As Integer, Optional ByVal EvalUnVal As Variant)
    Dim m_oprOP As EXPR_Sem_Expr_Operator
    Dim m_varValueA As Variant
    Dim m_varValueB As Variant
    Dim m_lngUnaryIdx As Long
    m_oprOP = Section.Operators.Items(Node.Index)
    If ((m_oprOP.Operation = EXPR_SEM_ES_SO_Equals) Or (m_oprOP.Operation = EXPR_SEM_ES_SO_Increment)) Then
        '//These don't evaluate
        '//Placeholder in case that changes
    Else
        '//If it's not a variable changing operator, then...
        Select Case m_oprOP.IndexAMeanings
            Case EXPR_SEM_ExprSection_OpIdxMeaningsA.EXPR_SEM_ES_OIMs_IndexA_Operator
                '//If the first member is an operator, then...
                m_varValueA = EvalOperator(Section.Members.Items(m_oprOP.IndexA), m_oprOP.IndexA, Expression, Section, Intermediate, EvalUnary, EvalUnIdx, EvalUnVal)
                    '//evaluate it, and return
            Case EXPR_SEM_ExprSection_OpIdxMeaningsA.EXPR_SEM_ES_OIMs_IndexA_Operand
                '//If the first member is an operand, then...
                m_varValueA = EvalOperand(Section.Members.Items(m_oprOP.IndexA), Expression, Section, Intermediate)
                    '//Evaluate the operand, and return
        End Select '//m_oprOP.IndexAMeanings
    End If '//((m_oprOP.Operation = EXPR_SEM_ES_SO_Equals) Or (m_oprOP.Operation = EXPR_SEM_ES_SO_Increment))
    If EvalUnary And m_oprOP.IndexB = EvalUnIdx Then
        '//If we're evaluating a unary operation, then...
        m_varValueB = EvalUnVal
            '//The second value is the unary value passed
            '//Reason:
            '//Unary operations work differently then binary operations
            '//Thus requiring us to back-track in the equasion after they're
            '//processed. To do this correctly we must pass the result of the
            '//Unary operation
    Else
        '//Otherwise...
        Select Case m_oprOP.IndexBMeanings
            Case EXPR_SEM_ExprSection_OpIdxMeaningsB.EXPR_SEM_ES_OIMs_IndexB_Operator
                '//If the second member is an operator, then...
                m_varValueB = EvalOperator(Section.Members.Items(m_oprOP.IndexB), m_oprOP.IndexB, Expression, Section, Intermediate, EvalUnary, EvalUnIdx, EvalUnVal)
                    '//Obtain the second member's value
            Case EXPR_SEM_ExprSection_OpIdxMeaningsB.EXPR_SEM_ES_OIMs_IndexB_Operand
                '//If the second member is an operand, then...
                m_varValueB = EvalOperand(Section.Members.Items(m_oprOP.IndexB), Expression, Section, Intermediate)
                    '//
            Case EXPR_SEM_ExprSection_OpIdxMeaningsB.EXPR_SEM_ES_OIMs_IndexB_Null
                m_lngUnaryIdx = -1
            Case EXPR_SEM_ExprSection_OpIdxMeaningsB.EXPR_SEM_ES_OIMs_IndexB_Expression
                '//If we're working with a unary operation, then...
                m_lngUnaryIdx = m_oprOP.IndexB
                    '//Setup the unary back-track position
        End Select
    End If
    Select Case m_oprOP.Operation
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Addition
            '//If we're to add, then...
            EvalOperator = m_varValueA + m_varValueB
                '//Return the sum of the two values
            #If DEBUGMODE And PRINTMODE Then
                Debug.Print "+"
            #End If
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Equals
            '//If we're to Set the operand, then...
            SetOperand Section.Members.Items(m_oprOP.IndexA), Expression, Section, Intermediate, m_varValueB
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_AddTo
            '//If we're to Increment the the operand by a value, then...
            SetOperand Section.Members.Items(m_oprOP.IndexA), Expression, Section, Intermediate, m_varValueB, True
                '//Set the variable
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_And
            '//If we're to Join the values together, then...
            EvalOperator = m_varValueA & m_varValueB
                '//Concatenate and return
            #If DEBUGMODE And PRINTMODE Then
                Debug.Print "&"
            #End If
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_BinaryAnd
            '//If we're to perform binary and
            EvalOperator = m_varValueA And m_varValueB
                '//Return the equal bits
            #If DEBUGMODE And PRINTMODE Then
                Debug.Print "And"
            #End If
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_BinaryEquals
            '//If we're to compare the values, then...
            EvalOperator = m_varValueA = m_varValueB
            #If DEBUGMODE And PRINTMODE Then
                Debug.Print "="
            #End If
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_BinaryNot
            '//If we're to invert the bits
            If Not m_lngUnaryIdx = -1 Then
                '//If we're to backstep, then...
                EvalOperator = EvalOperator(Section.Members.Items(m_oprOP.IndexB), m_oprOP.IndexB, Expression, Section, Intermediate, True, NodeIndex, (Not CDec(m_varValueA)))
                    '//Backstep and return
            Else
                EvalOperator = Not m_varValueA
                    '//Return the result
            End If
            #If DEBUGMODE And PRINTMODE Then
                Debug.Print "Not"
            #End If
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_BinaryOr
            '//If we're to perform Or
            EvalOperator = m_varValueA Or m_varValueB
                '//Return the bits of a, with b's bits
            #If DEBUGMODE And PRINTMODE Then
                Debug.Print "Or"
            #End If
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Division
            '//If we're to divide, then...
            EvalOperator = m_varValueA / m_varValueB
                '//Return a divided by b
            #If DEBUGMODE And PRINTMODE Then
                Debug.Print "/"
            #End If
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Modulus
            '//If we're to perform a modulus check
            EvalOperator = m_varValueA Mod m_varValueB
                '//Return the remainder of A divided by B
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_GreaterThan
            '//If we're to return whether or not one value is greater than the other, then...
            EvalOperator = m_varValueA > m_varValueB
            #If DEBUGMODE And PRINTMODE Then
                Debug.Print ">"
            #End If
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_GreaterThanOrEqualTo
            '//If we're to return whether or not one value is greater than or equal to
            '//the other, then...
            EvalOperator = m_varValueA >= m_varValueB
            #If DEBUGMODE And PRINTMODE Then
                Debug.Print ">="
            #End If
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Increment
            '//Increment the value by 1
            SetOperand Section.Members.Items(m_oprOP.IndexA), Expression, Section, Intermediate, 1, True
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_IntDivision
            '//If we're to perform integer division
            EvalOperator = m_varValueA \ m_varValueB
                '//Return the result of integer division upon a from b
            #If DEBUGMODE And PRINTMODE Then
                Debug.Print "\"
            #End If
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_InEquality
            '//If we're to compare inequality
            EvalOperator = m_varValueA <> m_varValueB
                '//Return whether or not (!(A==B))
            #If DEBUGMODE And PRINTMODE Then
                Debug.Print "<>"
            #End If
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_LessThan
            '//If we're to return whether or not one value is less than the other, then...
            EvalOperator = m_varValueA < m_varValueB
            #If DEBUGMODE And PRINTMODE Then
                Debug.Print "<"
            #End If
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_LessThanOrEqualTo
            '//If we're to return whether or not one value is less than or equal
            '//to the other, then...
            EvalOperator = (m_varValueA <= m_varValueB)
            #If DEBUGMODE And PRINTMODE Then
                Debug.Print "<="
            #End If
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Multiplication
            '//If we're to multiply, then...
            EvalOperator = m_varValueA * m_varValueB
                '//Multiply and return
            #If DEBUGMODE And PRINTMODE Then
                Debug.Print "*"
            #End If
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Negate
            '//If we're to negate a value, then...
            If Not m_lngUnaryIdx = -1 Then
                '//If we're to backtrack, then...
                EvalOperator = EvalOperator(Section.Members.Items(m_oprOP.IndexB), m_oprOP.IndexB, Expression, Section, Intermediate, True, NodeIndex, (-(m_varValueA)))
                    '//Negate, backtrack and return
            Else
                EvalOperator = -m_varValueA
                    '//Negate and return
            End If
            #If DEBUGMODE And PRINTMODE Then
                Debug.Print "Negate"
            #End If
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_PowerOf
            '//If we're to go to the power of, then...
            EvalOperator = m_varValueA ^ m_varValueB
                '//Return a to the power of b
            #If DEBUGMODE And PRINTMODE Then
                Debug.Print "^"
            #End If
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_StrAppend
                EvalOperator = m_varValueA & m_varValueB
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_Subtraction
            '//If we're to subtract, then...
            EvalOperator = m_varValueA - m_varValueB
                '//Return a minus b
            #If DEBUGMODE And PRINTMODE Then
                Debug.Print "-"
            #End If
        Case EXPR_SEM_ExprSection_SubOperation.EXPR_SEM_ES_SO_XOr
            '//If we're to perform exclusive or, then...
            EvalOperator = m_varValueA Xor m_varValueB
                '//Perform eXclusive Or
            #If DEBUGMODE And PRINTMODE Then
                Debug.Print "Xor"
            #End If
    End Select
End Function

Public Function EvalOperand(Operand As EXPR_Sem_Expr_Member, ByRef Expression As EXPR_Sem_Expression, ByRef Section As EXPR_Sem_Expr_Section, ByRef Intermediate As EXPR_Sem_Intermediate) As Variant
    Dim m_nodNode As EXPR_Sem_Expr_Member
    'm_nodNode = Section.Members.Items(Operand)
    With Operand
        '//Select the operand's namespace
        Select Case .Type
            '//Select the type for comparison
            Case EXPR_S_E_M_T_Constant
                '//If it's a constant, then...
                EvalOperand = GetCstVal(Expression.Constants.Items(.Index).SubType, Intermediate.ConstTable, Expression.Constants.Items(.Index).Index)
                    '//Return the constant value
            Case EXPR_S_E_M_T_ComplexIdentifier
                '//If it's a complex identifier, then...
                EvalOperand = ProcessOperandProcedure(Expression.ComplexIdentifiers.Items(.Index), Intermediate, Section, Expression)
                    '//Return result of the operand procedure
            Case EXPR_S_E_M_T_SubExpression
                '//If it's a sub-expression, then...
                EvalOperand = EvaluateSection(Expression, Intermediate, Expression.Subsections.Items(.Index))
                    '//Return the result of the sub-expression
        End Select
    End With
End Function

Public Function ProcessOperandProcedure(Procedure As EXPR_Sem_Expr_ComplexID, Intermediate As EXPR_Sem_Intermediate, Section As EXPR_Sem_Expr_Section, Expression As EXPR_Sem_Expression) As Variant
    Dim m_strName As String
        '//Name var
    Dim m_lngArgument As Long
        '//Arg index
    Dim m_agsArgs As EXPR_Sem_Expr_ComplexID_ArgList
        '//Arguments list
    Dim m_strPrint As String
        '//DebugPrint text
    Dim m_varResult As Variant
        '//Result
    On Error Resume Next
    Select Case Procedure.Count
        Case 0
            
        Case 1
            If GroupItemsAre(Procedure, Intermediate, Section, Expression, EXPR_Sem_Expr_ComplexID_Member_Type.EXPR_S_E_CI_M_T_Identifier) Then
                m_strName = LCase$(CStr(Intermediate.ConstTable.Strings.Items(Procedure.Items(0).Index).Bytes))
                Select Case m_strName
                    Case "vbcrlf"
                        ProcessOperandProcedure = vbCrLf
                    Case "vbcr"
                        ProcessOperandProcedure = vbCr
                    Case "vblf"
                        ProcessOperandProcedure = vbLf
                    Case "true"
                        ProcessOperandProcedure = True
                    Case "result"
                        ProcessOperandProcedure = m_varReturn
                    Case "w"
                        ProcessOperandProcedure = m_gl_W
                    Case "x"
                        ProcessOperandProcedure = m_gl_X
                    Case "y"
                        ProcessOperandProcedure = m_gl_Y
                    Case "z"
                        ProcessOperandProcedure = m_gl_Z
                    Case Else
                        If VarExists(m_strName) Then
                            ProcessOperandProcedure = VarVal(m_strName)
                        Else
                            ThrowException Intermediate, 0, "Procedure", EXPR_SEM_I_EXP_T_Missing
                        End If
                End Select
            End If
        Case 2
            If GroupItemsAre(Procedure, Intermediate, Section, Expression, EXPR_Sem_Expr_ComplexID_Member_Type.EXPR_S_E_CI_M_T_Identifier, EXPR_Sem_Expr_ComplexID_Member_Type.EXPR_S_E_CI_M_T_ArgumentList) Then
                ProcessOperandProcedure = ProcessGlobalProcedure(Procedure, Expression.ArgumentLists.Items(Procedure.Items(1).Index), Intermediate, Section, Expression, 0)
            End If
        Case 4
            If GroupItemsAre(Procedure, Intermediate, Section, Expression, EXPR_Sem_Expr_ComplexID_Member_Type.EXPR_S_E_CI_M_T_Identifier, EXPR_Sem_Expr_ComplexID_Member_Type.EXPR_S_E_CI_M_T_SubMemberItem, EXPR_Sem_Expr_ComplexID_Member_Type.EXPR_S_E_CI_M_T_Identifier, EXPR_Sem_Expr_ComplexID_Member_Type.EXPR_S_E_CI_M_T_ArgumentList) Then
                m_strName = CStr(Intermediate.ConstTable.Strings.Items(Procedure.Items(0).Index).Bytes)
                Select Case LCase$(m_strName)
                    Case "vb", "vba"
                        ProcessOperandProcedure = ProcessGlobalProcedure(Procedure, Expression.ArgumentLists.Items(Procedure.Items(3).Index), Intermediate, Section, Expression, 2)
                    Case "debug"
                        m_strName = CStr(Intermediate.ConstTable.Strings.Items(Procedure.Items(2).Index).Bytes)
                        Select Case LCase$(m_strName)
                            Case "print"
                                #If False Then
                                    m_agsArgs = Expression.ArgumentLists.Items(Procedure.Items(3).Index)
                                    For m_lngArgument = 0 To m_agsArgs.Count - 1
                                        m_varResult = GetArgValue(m_agsArgs.Items(m_lngArgument), Intermediate, Expression, Section)
                                        If Not IsNull(m_varResult) Then
                                            m_strPrint = m_strPrint & m_varResult
                                            If Not m_lngArgument = m_agsArgs.Count - 1 Then
                                                m_strPrint = m_strPrint & Space(14 - Len(CStr(m_varResult)))
                                            End If
                                        Else
                                            If Not m_lngArgument = m_agsArgs.Count - 1 Then
                                                m_strPrint = m_strPrint & Space(14)
                                            End If
                                        End If
                                    Next
                                    Debug.Print m_strPrint
                                #End If
                        End Select
                End Select
            End If
    End Select
End Function

Public Function ProcessGlobalProcedure(Procedure As EXPR_Sem_Expr_ComplexID, ArgList As EXPR_Sem_Expr_ComplexID_ArgList, Intermediate As EXPR_Sem_Intermediate, Section As EXPR_Sem_Expr_Section, Expression As EXPR_Sem_Expression, BaseIndex As Integer) As Variant
    Dim m_strName As String
        '//Name variable
    Dim m_varButtons As Variant
        '//Buttons var, for msgbox, to query whether or not it's null
    m_strName = LCase$(CStr(Intermediate.ConstTable.Strings.Items(Procedure.Items(BaseIndex).Index).Bytes))
        '//Obtain the function identifier
    On Error Resume Next
        '//On an error, ignore it.
    If ArgList.Count < 1 Then _
        Exit Function
        '//If there aren't any members, then...
            '//Exit.
    Select Case m_strName
        Case "msgbox"
            '//If we're working with a message box function
            With ArgList
                If .Count < 2 Then
                    '//If there's less then two arguments, then...
                    ProcessGlobalProcedure = MsgBox(GetArgValue(.Items(0), Intermediate, Expression, Section), , App.Title & " (Script)")
                ElseIf .Count < 3 Then
                    '//If there's less then three arguments, then...
                    ProcessGlobalProcedure = MsgBox(GetArgValue(.Items(0), Intermediate, Expression, Section), GetArgValue(.Items(1), Intermediate, Expression, Section), App.Title & " (Script)")
                        '//Call the msgbox with the prompt and buttons
                Else
                    '//If there are three or more arguments, then...
                    m_varButtons = GetArgValue(.Items(1), Intermediate, Expression, Section)
                        '//Obtain the buttons value
                    If (Not (IsNull(m_varButtons))) Then
                        '//If the buttons was omitted, then...
                        ProcessGlobalProcedure = MsgBox(GetArgValue(.Items(0), Intermediate, Expression, Section), m_varButtons, GetArgValue(.Items(2), Intermediate, Expression, Section) & " (Script)")
                            '//Process it without the buttons arg
                    Else
                        ProcessGlobalProcedure = MsgBox(GetArgValue(.Items(0), Intermediate, Expression, Section), , GetArgValue(.Items(2), Intermediate, Expression, Section) & " (Script)")
                            '//Process it with the buttons arg
                    End If '//(Not (IsNull(m_varButtons)))
                End If
            End With '//ArgList
        Case "mid"
            If ArgList.Count < 2 Then Exit Function
                '//If there aren't the required two arguments, then...
                    '//exit
            If (ArgList.Count = 2) Then
                '//If there are two args only, then...
                With ArgList
                    '//Select the arglist's namespace
                    ProcessGlobalProcedure = Mid$(CStr(GetArgValue(.Items(0), Intermediate, Expression, Section)), GetArgValue(.Items(1), Intermediate, Expression, Section))
                        '//Process the mid with two args and return
                End With '//ArgList
            ElseIf ArgList.Count > 2 Then
                With ArgList
                    '//Select the arglist's namespace
                    ProcessGlobalProcedure = Mid$(CStr(GetArgValue(.Items(0), Intermediate, Expression, Section)), GetArgValue(.Items(1), Intermediate, Expression, Section), GetArgValue(.Items(2), Intermediate, Expression, Section))
                        '//Process the mid with three args and return
                End With '//ArgList
            End If '//(ArgList.Count = 2)
        Case "left"
            '//If we're using the left function, then...
            If ArgList.Count < 2 Then _
                Exit Function
                '//If there aren't the required two arguments, then...
                    '//exit
            With ArgList
                '//Select the arglist's namespace
                ProcessGlobalProcedure = Left$(GetArgValue(.Items(0), Intermediate, Expression, Section), GetArgValue(.Items(1), Intermediate, Expression, Section))
                    '//Process the left procedure
            End With '//ArgList
        Case "right"
            '//etc beyond here
            If ArgList.Count < 2 Then Exit Function
            With ArgList
                ProcessGlobalProcedure = Right$(GetArgValue(.Items(0), Intermediate, Expression, Section), GetArgValue(.Items(1), Intermediate, Expression, Section))
            End With
        Case "sec"
            With ArgList
                ProcessGlobalProcedure = Sec(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "cosec"
            With ArgList
                ProcessGlobalProcedure = Cosec(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "cotan"
            With ArgList
                ProcessGlobalProcedure = Cotan(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "arcsin"
            With ArgList
                ProcessGlobalProcedure = Arcsin(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "arccos"
            With ArgList
                ProcessGlobalProcedure = Arccos(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "arcsec"
            With ArgList
                ProcessGlobalProcedure = Arcsec(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "arccosec"
            With ArgList
                ProcessGlobalProcedure = Arccosec(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "arccotan"
            With ArgList
                ProcessGlobalProcedure = Arccotan(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "hcos"
            With ArgList
                ProcessGlobalProcedure = HCos(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "hsin"
            With ArgList
                ProcessGlobalProcedure = HSin(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "htan"
            With ArgList
                ProcessGlobalProcedure = HTan(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "hsec"
            With ArgList
                ProcessGlobalProcedure = HSec(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "hcosec"
            With ArgList
                ProcessGlobalProcedure = HCosec(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "hcotan"
            With ArgList
                ProcessGlobalProcedure = HCotan(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "harcsin"
            With ArgList
                ProcessGlobalProcedure = HArcsin(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "harccos"
            With ArgList
                ProcessGlobalProcedure = HArccos(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "harctan"
            With ArgList
                ProcessGlobalProcedure = HArctan(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "harcsec"
            With ArgList
                ProcessGlobalProcedure = HArcsec(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "harccosec"
            With ArgList
                ProcessGlobalProcedure = HArccosec(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "harccotan"
            With ArgList
                ProcessGlobalProcedure = HArccotan(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "logn"
            With ArgList
                If .Count < 2 Then _
                    Exit Function
                ProcessGlobalProcedure = LogN(GetArgValue(.Items(0), Intermediate, Expression, Section), GetArgValue(.Items(1), Intermediate, Expression, Section))
            End With
        Case "cos"
            With ArgList
                ProcessGlobalProcedure = Cos(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "sin"
            
            With ArgList
                ProcessGlobalProcedure = Sin(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "tan"
            
            With ArgList
                ProcessGlobalProcedure = Tan(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "abs"
            With ArgList
                ProcessGlobalProcedure = Abs(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "cint"
            With ArgList
                ProcessGlobalProcedure = CInt(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "sqr"
            With ArgList
                ProcessGlobalProcedure = Sqr(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "clng"
            With ArgList
                ProcessGlobalProcedure = CLng(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "cdbl"
            With ArgList
                ProcessGlobalProcedure = CDbl(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "cbool"
            With ArgList
                ProcessGlobalProcedure = CBool(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "cbyte"
            With ArgList
                ProcessGlobalProcedure = CByte(GetArgValue(.Items(0), Intermediate, Expression, Section))
            End With
        Case "switch"
            '//Switch function
            If (ArgList.Count Mod 2 = 0) Then
                '//If the argument count is in a grouping of two
                Dim m_varExp As Variant
                    '//Evaluated argument value
                Dim m_lngArg As Long
                    '//The argument loop index
                For m_lngArg = 0 To ArgList.Count - 2 Step 2
                    '//Loop through the arglist, hitting evey other one
                    With ArgList
                        '//Select the arglist's namespace
                        m_varExp = GetArgValue(.Items(m_lngArg), Intermediate, Expression, Section)
                            '//Obtain the expressions value
                        If CBool(m_varExp) Then
                            '//If it evaluates to true, then...
                            ProcessGlobalProcedure = GetArgValue(.Items(m_lngArg + 1), Intermediate, Expression, Section)
                                '//Process the procedure
                            Exit For
                                '//Exit, we've found our value
                        End If '//CBool(m_varExp)
                    End With '//ArgList
                Next '//[m_lngArg]
            End If '//(ArgList.Count Mod 2 = 0)
        Case Else
            '//If it's undefined, then...
                '//Do nothing :o
    End Select '//m_strName
End Function

Public Function GroupItemsAre(Identifier As EXPR_Sem_Expr_ComplexID, Intermediate As EXPR_Sem_Intermediate, Section As EXPR_Sem_Expr_Section, Expression As EXPR_Sem_Expression, ParamArray Types() As Variant) As Boolean
    Dim m_cidType As EXPR_Sem_Expr_ComplexID_Member_Type
        '//Complex Identifier Member Type
    Dim m_lngIndex As Long
        '//Loop index var
    For m_lngIndex = LBound(Types) To UBound(Types)
        '//Loop through the types
        m_cidType = Types(m_lngIndex)
            '//Obtain the current member
        If (Not (Identifier.Items(m_lngIndex).Type = m_cidType)) Then
            '//If the type conflicts with the current type
            GroupItemsAre = False
                '//Indicate failure
            Exit Function
                '//Exit
        End If '//(Not (Identifier.Items(m_lngIndex).Type = m_cidType))
    Next '//[m_lngIndex]
    GroupItemsAre = True
        '//Indicate success
End Function

Public Function GetArgValue(Argument As EXPR_Sem_Expr_ComplexID_Arg, Intermediate As EXPR_Sem_Intermediate, Expression As EXPR_Sem_Expression, Section As EXPR_Sem_Expr_Section) As Variant
    Dim m_escConst As EXPR_SEM_Expression_Constant
        '//Expression constant
    Select Case Argument.Type
        '//Select the argument's type
        Case EXPR_Sem_ExprSection_CI_DA_Type.EXPR_SEM_CI_DA_T_ComplexIdentifier
            '//If it's a complex identifier, then...
            GetArgValue = ProcessOperandProcedure(Expression.ComplexIdentifiers.Items(Argument.Index), Intermediate, Section, Expression)
                '//Obtain the argument value by processing the procedure and return
        Case EXPR_Sem_ExprSection_CI_DA_Type.EXPR_SEM_CI_DA_T_Constant
            '//If it's a constant, then...
            m_escConst = Expression.Constants.Items(Argument.Index)
                '//Obtain the constant information
            GetArgValue = GetCstVal(m_escConst.SubType, Intermediate.ConstTable, m_escConst.Index)
                '//Return the constant value from the constant table
        Case EXPR_Sem_ExprSection_CI_DA_Type.EXPR_SEM_CI_DA_T_Empty
            '//If it's an empty argument
            GetArgValue = Null
                '//Return null
        Case EXPR_Sem_ExprSection_CI_DA_Type.EXPR_SEM_CI_DA_T_SubSection
            '//If it's a subexpression, then...
            GetArgValue = EvaluateSection(Expression, Intermediate, Expression.Subsections.Items(Argument.Index))
                '//Evaluate the section and return
    End Select
End Function

Public Function InterpretText(Intermediate As String) As Variant
    Dim m_essIntermediate As EXPR_Sem_Intermediate
    Dim m_emeExpression As EXPR_Sem_Expression
    Dim m_lprResult As EXPR_Lex_Tokens
    m_lprResult = modLexMeths.LexProc(Intermediate)
        '//Lexical process the text
    m_emeExpression = ParseExpression(m_essIntermediate, m_lprResult, 0, False)
        '//Semantic parse
    InterpretText = modInterpret.Evaluate(m_emeExpression, m_essIntermediate)
        '//Evaluate
End Function

Public Function FormatVal(Value As Variant)
    If IsNumeric(Value) Then
        '//If the value is numeric, then
        If (VarType(Value) = vbDouble) Then
            '//If it's a double, then...
            FormatVal = Format(Value, "###,###.############")
                '//Format it with quite a few
            If (Right(CStr(FormatVal), 1) = ".") Then
                '//If it didn't have a decimal value at the and then...
                FormatVal = Format(Value, "###,###")
                    '//Reformat
            End If '//(Right(CStr(FormatVal), 1) = ".")
        Else
            '//If it's not a double number
            FormatVal = Format(Value, "###,##0")
                '//format with no decimal
        End If '//(VarType(Value) = vbDouble)
    Else
        '//Otherwise
        FormatVal = """" & Value & """"
            '//Return with quotes
    End If
End Function

Public Function SetOperand(Operand As EXPR_Sem_Expr_Member, ByRef Expression As EXPR_Sem_Expression, ByRef Section As EXPR_Sem_Expr_Section, ByRef Intermediate As EXPR_Sem_Intermediate, Value As Variant, Optional Increment As Boolean, Optional Decrement As Boolean) As Variant
    Dim m_cidID As EXPR_Sem_Expr_ComplexID
        '//Complex identifier
    Dim m_strName As String
        '//Var Name variable
    On Error Resume Next
        '//Ignore errors
    With Operand
        '//Select the opernad's namespace
        Select Case .Type
            '//Select its type for comparison
            Case EXPR_S_E_M_T_ComplexIdentifier
                '//If it's a complex identifier, then...
                m_cidID = Expression.ComplexIdentifiers.Items(Operand.Index)
                    '//Obtain the complex identifier information
                If m_cidID.Count = 1 Then
                    '//If there is only one member, then...
                    If (m_cidID.Items(0).Type = EXPR_Sem_Expr_ComplexID_Member_Type.EXPR_S_E_CI_M_T_Identifier) Then
                        '//If the first member is a string identifier
                        m_strName = LCase(CStr(Intermediate.ConstTable.Strings.Items(m_cidID.Items(0).Index).Bytes))
                            '//Obtain the name
                        Select Case m_strName
                            '//Select the name for comparison
                            Case "w"
                                '//Internal w var
                                If Increment Then
                                    '//If we're incrementing, then...
                                    m_gl_W = m_gl_W + Value
                                        '//Set it to its value plus the other value
                                Else
                                    '//If we're not incrementing, then...
                                    m_gl_W = Value
                                        '//Set the var to the value
                                End If '//(Increment)
                            Case "x"
                                '//Internal x var
                                If Increment Then
                                    '//If we're incrementing, then...
                                    m_gl_X = m_gl_X + Value
                                        '//Set it to its value plus the other value
                                Else
                                    '//If we're not incrementing, then...
                                    m_gl_X = Value
                                        '//Set the var to the value
                                End If '//(Increment)
                            Case "y"
                                '//Internal y var
                                If Increment Then
                                    '//If we're incrementing, then...
                                    m_gl_Y = m_gl_Y + Value
                                        '//Set it to its value plus the other value
                                Else
                                    '//If we're not incrementing, then...
                                    m_gl_Y = Value
                                        '//Set the var to the value
                                End If '//(Increment)
                            Case "z"
                                '//Internal z var
                                If Increment Then
                                    '//If we're incrementing, then...
                                    m_gl_Z = m_gl_Z + Value
                                        '//Set it to its value plus the other value
                                Else
                                    '//If we're not incrementing, then...
                                    m_gl_Z = Value
                                        '//Set the var to the value
                                End If '//(Increment)
                            Case "result"
                                '//If we're working with the result var
                                '//returned by multi-expression parsing.
                                If Increment Then
                                    '//If we're to increment, then...
                                    m_varReturn = m_varReturn + Value
                                        '//set the value to its value plus the passed value
                                ElseIf Decrement Then
                                    m_varReturn = m_varReturn - Value
                                Else
                                    m_varReturn = Value
                                        '//Set the value to the passed value
                                End If '//(Increment)
                            Case Else
                                '//Other variable
                                If VarExists(m_strName) Then
                                    '//If the variable exists, then...
                                    If Increment Then
                                        '//If we're to increment the value, then...
                                        IncrementVar m_strName, Value
                                            '//Increment the variable
                                    ElseIf Decrement Then
                                        '//If we're to decrement the variable, then...
                                        IncrementVar m_strName, -Value
                                            '//Increment the variable by the negative value
                                    Else
                                        '//Otherwise...
                                        SetVarVal m_strName, Value
                                            '//Set the variable value
                                    End If '//(Increment)
                                Else
                                    '//If the variable doesn't exist, then...
                                    AddVar m_strName, Value
                                        '//Add the variable
                                End If '//(VarExists(m_strName))
                        End Select '//m_strName
                    End If '//(m_cidID.Items(0).Type = EXPR_Sem_Expr_ComplexID_Member_Type.EXPR_S_E_CI_M_T_Identifier)
                End If '//m_cidID.Count = 1
        End Select '//.Type
    End With '//Operand
End Function

Public Function VarExists(Var As String) As Boolean
    Dim m_varDummy As Variant
        '//Dummy var
    On Error GoTo Catch
        '//Error traping
    m_varDummy = m_colNames.Item(Var)
        '//Obtain the dummy value
    VarExists = True
        '//If no error, it exists...
    GoTo Finally
        '//Exit
Catch:
    '//Catch, ignore, and exit
Finally:
    '//Exit
End Function

Public Function VarVal(Name As Variant) As Variant
    VarVal = m_ivsVars.Items(m_colNames.Item(Name)).Value
        '//Return the variable's value
End Function

Public Sub IncrementVar(Name As Variant, Val As Variant)
    With m_ivsVars.Items(m_colNames.Item(Name))
        '//Select the interpreted variable's namespace
        .Value = .Value + Val
            '//increment
    End With
End Sub

Public Sub SetVarVal(Name As Variant, Val As Variant)
    m_ivsVars.Items(m_colNames.Item(Name)).Value = Val
End Sub

Public Sub AddVar(Name As Variant, Value As Variant)
    With m_ivsVars
        '//Select the internal variables' namespace
        If .Count = 0 Then
            '//If there aren't any variables, then...
            ReDim .Items(.Count)
                '//Initialize the array
        Else
            '//Otherwise...
            ReDim Preserve .Items(.Count)
                '//Redimension and preserve the array data, adding the new member
                '//to the end
        End If
        .Items(.Count).Value = Value
            '//Change the new member's value
        m_colNames.Add .Count, Name
            '//Add it to the collection
        .Count = .Count + 1
            '//Increment
    End With
End Sub

Public Function GroupParse(Text As String, Intermediate As EXPR_Sem_Intermediate) As EXPR_SEM_Expressions
    Dim m_staText() As String
        '//Text lines
    Dim m_varText As Variant
        '//Loop var
    Dim m_atsTokens As EXPR_Lex_Tokens
        '//Expression stream
    Dim m_expExpr As EXPR_Sem_Expression
        '//Active expression
    Dim m_sesExpressions As EXPR_SEM_Expressions
        '//Expressions result
    m_staText = Split(Text, vbCrLf)
        '//Obtain the lines
    For Each m_varText In m_staText
        '//Loop through the lines
        m_atsTokens = LexProc(m_varText)
            '//Obtain the tokens for the active expression
        m_expExpr = modSemExpressions.ParseExpression(Intermediate, m_atsTokens, 0)
            '//Obtain the expression for the active stream
        AddExp m_sesExpressions, m_expExpr
            '//Add the expression to the array
    Next '//[m_varText]
    GroupParse = m_sesExpressions
        '//Return
End Function

Private Sub AddExp(Expressions As EXPR_SEM_Expressions, Expression As EXPR_Sem_Expression)
    With Expressions
        '//Select the expressions' namespace
        If (.Count = 0) Then
            '//If there are no expressions, then...
            ReDim .Items(.Count)
                '//Initialize the array
        Else
            '//If there are, then...
            ReDim Preserve .Items(.Count)
                '//Redimension the array keeping old data
        End If '//(.Count = 0)
        .Items(.Count) = Expression
        .Count = .Count + 1
    End With
End Sub

Public Function GroupEvaluate(Expressions As EXPR_SEM_Expressions, Intermediate As EXPR_Sem_Intermediate) As Variant
    Dim m_lngLoop As Long
        '//Loop var
    m_varReturn = vbNullString
    For m_lngLoop = 0 To Expressions.Count - 1
        '//Loop through the expressions
        Evaluate Expressions.Items(m_lngLoop), Intermediate
            '//Evaluate the expression
    Next '//[m_lngLoop]
    GroupEvaluate = m_varReturn
        '//Return the result var
End Function
