Attribute VB_Name = "modSemStructs"
Option Explicit
'//Author: Linguar Amadala (Allen Copeland)

Public Enum EXPR_SEM_ExprSection_Op_Precedence
    EXPR_SEM_ES_OP_P_Append
        '//A += B, A &= B
    EXPR_SEM_ES_OP_P_LogXor
        '//A X0r B
    EXPR_SEM_ES_OP_P_LogOr
        '//A 0r B
    EXPR_SEM_ES_OP_P_LogAnd
        '//A And B
    EXPR_SEM_ES_OP_P_LogNot
        '//Not
    EXPR_SEM_ES_OP_P_BoolObjIs
        '//To do, not needed for an objectless system
    EXPR_SEM_ES_OP_P_BoolGreaterThanOrEqualTo
        '//A >= B
    EXPR_SEM_ES_OP_P_BoolLessThanOrEqualTo
        '//A <= B
    EXPR_SEM_ES_OP_P_BoolGreaterThan
        '//A > B
    EXPR_SEM_ES_OP_P_BoolLessThan
        '//A < B
    EXPR_SEM_ES_OP_P_BoolInequality
        '//A <> B
    EXPR_SEM_ES_OP_P_BoolEquality
        '//A == B
    EXPR_SEM_ES_OP_P_StrConcatination
        '//A & B
    EXPR_SEM_ES_OP_P_AddSubt
        '//A + B, A - B
    EXPR_SEM_ES_OP_P_Modulus
        '//A Mod B
    EXPR_SEM_ES_OP_P_IntegerDivide
        '//A \ B
    EXPR_SEM_ES_OP_P_MultDivide
        '//A * B
    EXPR_SEM_ES_OP_P_Negation
        '//-A
    EXPR_SEM_ES_OP_P_Exponentation
        '//A ^ B
End Enum

Public Enum EXPR_SEM_ExprSection_OpIdxMeaningsA
    '//First index meaning, before the operator
    EXPR_SEM_ES_OIMs_IndexA_Operand = 1
        '//It's an operand
    EXPR_SEM_ES_OIMs_IndexA_Operator = 2
        '//It's an operator
End Enum '//EXPR_SEM_ExprSection_OpIdxMeaningsA
Public Enum EXPR_SEM_ExprSection_OpIdxMeaningsB
    '//First index meaning, after the operator
    EXPR_SEM_ES_OIMs_IndexB_Operand = 1
        '//It's an operand
    EXPR_SEM_ES_OIMs_IndexB_Operator = 2
        '//It's an operator
    EXPR_SEM_ES_OIMs_IndexB_Expression = 3
        '//It's an expression, used for unary operations, to back-track
    EXPR_SEM_ES_OIMs_IndexB_Null = 4
        '//It's null, this is the end of the line
End Enum '//EXPR_SEM_ExprSection_OpIdxMeaningsB

Public Enum EXPR_SEM_ExprSection_SubOperation
    EXPR_SEM_ES_SO_Addition = 1 '//'+'
    EXPR_SEM_ES_SO_Subtraction '//'-'
    EXPR_SEM_ES_SO_Multiplication '//'*'
    EXPR_SEM_ES_SO_Division '//'/'
    EXPR_SEM_ES_SO_IntDivision '//'\'
    EXPR_SEM_ES_SO_PowerOf '//'^'
    EXPR_SEM_ES_SO_Equals '//'='
    EXPR_SEM_ES_SO_And '//'&'
    EXPR_SEM_ES_SO_XOr '//'XOr'
    EXPR_SEM_ES_SO_BinaryNot '//Not
    EXPR_SEM_ES_SO_BinaryAnd  '//'And'
    EXPR_SEM_ES_SO_BinaryOr  '//'Or'
    EXPR_SEM_ES_SO_BinaryEquals '//'=='
    EXPR_SEM_ES_SO_GreaterThan  '//'>'
    EXPR_SEM_ES_SO_LessThan  '//'<'
    EXPR_SEM_ES_SO_GreaterThanOrEqualTo '//'>='
    EXPR_SEM_ES_SO_LessThanOrEqualTo '//'<='
    EXPR_SEM_ES_SO_InEquality '<>
    EXPR_SEM_ES_SO_Negate '//"'-'(val)"
    EXPR_SEM_ES_SO_StrAppend '//'&='
    EXPR_SEM_ES_SO_AddTo '//'+='
    EXPR_SEM_ES_SO_Increment
    EXPR_SEM_ES_SO_Modulus
        '//If two operators show up, this merely negates the secondary value
End Enum '//EXPR_SEM_ExprSection_SubOperation

Public Enum EXPR_SEM_Intermediate_EXP_Type
    EXPR_SEM_I_EXP_T_Expected
        '//Expected: %1
    EXPR_SEM_I_EXP_T_Unexpected
        '//Unexpected: %1
    EXPR_SEM_I_EXP_T_Missing
        '//Missing: %1
End Enum '//EXPR_SEM_Intermediate_EXP_Type

Public Enum EXPR_Sem_ExprSection_CI_DA_Type
    EXPR_SEM_CI_DA_T_Constant
        '//Constant arg
    EXPR_SEM_CI_DA_T_ComplexIdentifier
        '//Complex identifier
    EXPR_SEM_CI_DA_T_SubSection
        '//Subsection
    EXPR_SEM_CI_DA_T_Empty
        '//An omitted argument
End Enum '//EXPR_Sem_ExprSection_CI_DA_Type

Public Enum EXPR_Sem_Expr_Const_Type
    '//EXPREval - Semantic Expression - Constant.Type
    EXPR_S_E_C_T_String
        '//String value
    EXPR_S_E_C_T_Byte
        '//Byte value
    EXPR_S_E_C_T_Integer
        '//Integer value
    EXPR_S_E_C_T_Long
        '//Long value
    EXPR_S_E_C_T_Single
        '//Single value
    EXPR_S_E_C_T_Double
        '//Double value
    EXPR_S_E_C_T_Currency
        '//Currency value
End Enum '//EXPR_Sem_Expr_Const_Type

Public Enum EXPR_Sem_ExprSection_Member_Type
    '//EXPREval - Semantic Expression - Member.Type
    EXPR_S_E_M_T_ComplexIdentifier
        '//Complex Identifier Member
    EXPR_S_E_M_T_Constant
        '//Constant Member
    EXPR_S_E_M_T_Operator
        '//Operator Member
    EXPR_S_E_M_T_SubExpression
        '//Sub-Expression
End Enum '//EXPR_Sem_ExprSection_Member_Type

Public Enum EXPR_Sem_Expr_ComplexID_Member_Type
    '//EXPREval - Semantic Expression - ComplexId.Type
    EXPR_S_E_CI_M_T_Identifier
        '//Identifier
    EXPR_S_E_CI_M_T_SubMemberItem
        '//Sub member access
    EXPR_S_E_CI_M_T_ArgumentList
        '//Argument list
End Enum '//EXPR_Sem_Expr_ComplexID_Member_Type
Public Type EXPR_UnicodeStr
    Bytes() As Byte
        '//This is so it saves the string padded with zeros
        '//useless beyond that. If you don't plan on saving the text
        '//Remove this functionality and modify it to include standard
        '//Strings.
End Type '//EXPR_UnicodeStr
Public Type EXPR_LongTable
    '//EXPR Eval - Long table
    Count As Long
        '//Num items
    Items() As Long
End Type '//EXPR_LongTable
Public Type EXPR_ByteTable
    Count As Long
        '//Num items
    Items() As Byte
        '//Bytes array
End Type '//EXPR_ByteTable
Public Type EXPR_IntegerTable
    '//EXPR Eval - Integer Table
    Count As Long
        '//Number of items
    Items() As Long
        '//Long array
End Type '//EXPR_IntegerTable
Public Type EXPR_SingleTable
    Count As Long
        '//Number of items
    Items() As Single
End Type '//EXPR_SingleTable
Public Type EXPR_DoubleTable
    Count As Long
        '//Number of items
    Items() As Single
End Type '//EXPR_DoubleTable
Public Type EXPR_CurrencyTable
    Count As Long
        '//Number of items
    Items() As Currency
        '//Currency Array
End Type '//EXPR_CurrencyTable
Public Type EXPR_StringTable
    Count As Long
        '//Number of items
    Items() As EXPR_UnicodeStr
        '//Unicode Strings array
End Type '//EXPR_StringTable
Public Type EXPR_ConstTable
    '//EXPR Eval - Constant Table
    Strings As EXPR_StringTable
        '//Strings Table
    Bytes As EXPR_ByteTable
        '//Byte Table
    Integers As EXPR_IntegerTable
        '//Integer Table
    Longs As EXPR_LongTable
        '//Long Table
    Singles As EXPR_SingleTable
        '//String table
    Doubles As EXPR_DoubleTable
        '//Double Table
    Currencies As EXPR_CurrencyTable
        '//Currency table
End Type '//EXPR_ConstTable

Public Type EXPR_Sem_Expr_Member
    '//EXPR Eval - Semantic Expression - Member
    Type As Byte 'EXPR_Sem_ExprSection_Member_Type
        '//The type of the member
    Index As Integer
        '//Its index to the member types
End Type '//EXPR_Sem_Expr_Member
Public Type EXPR_Sem_Expr_Members
    '//EXPR Eval - Semantic Expression - Members
    Count As Integer
        '//Number of members
    Items() As EXPR_Sem_Expr_Member
        '//Members array
End Type '//EXPR_Sem_Expr_Members

Public Type EXPR_Sem_Expr_Operator
    '//EXPR Eval - Semantic Expression - Operator
    IndexAMeanings As Byte 'EXPR_SEM_ExprSection_OpIdxMeaningsA
        '//Previous member Information
    IndexBMeanings As Byte 'EXPR_SEM_ExprSection_OpIdxMeaningsB
        '//Next member information
    Operation As Byte 'EXPR_SEM_ExprSection_SubOperation
        '//Operation to perform
    IndexA As Integer
        '//First member index
    IndexB As Integer
        '//Prior member index
End Type '//EXPR_Sem_Expr_Operator
Public Type EXPR_Sem_Expr_Operators
    '//EXPR Eval - Semantic Expression Operators list
    Count As Integer
        '//Number of operators
    Items() As EXPR_Sem_Expr_Operator
        '//Array of operators
End Type '//EXPR_Sem_Expr_Operators

Public Type EXPR_Sem_Expr_Section
    '//EXPR Eval - Semantic Expression - Expression Section
    FirstMember As Integer
        '//First member in the section
    Members As EXPR_Sem_Expr_Members
        '//Members array
    Operators As EXPR_Sem_Expr_Operators
        '//Operators array
End Type '//EXPR_Sem_Expr_Section

Public Type EXPR_Sem_Expr_ComplexID_Arg
    '//EXPR Eval - Semantic Expression - Complex Identifier - Argument
    Type            As Byte 'EXPR_Sem_ExprSection_CI_DA_Type
        '//The type of the argument
    Index           As Integer
        '//The index of the sub type
End Type
Public Type EXPR_Sem_Expr_ComplexID_ArgList
    '//EXPR Eval - Semantic Expression - Complex Identifier - Argument List
    Count           As Byte
        '//The number of arguments, 255 max
    Items()         As EXPR_Sem_Expr_ComplexID_Arg
        '//The Arg array
End Type '//EXPR_Sem_Expr_ComplexID_ArgList

Public Type EXPR_Sem_Expr_ComplexID_ArgLists
    Count As Integer
    Items() As EXPR_Sem_Expr_ComplexID_ArgList
End Type

Public Type EXPR_Sem_Expr_ComplexID_Member
    '//EXPR Eval - Semantic Expression - Complex Identifier - Member
    Type As Byte 'EXPR_Sem_Expr_ComplexID_Member_Type
        '//The member type
    Index As Long
        '//The sub structure index
End Type '//EXPR_Sem_Expr_ComplexID_Member
Public Type EXPR_Sem_Expr_ComplexID
    '//EXPR Eval - Semantic Expression - Complex Identifier
    Count As Integer
        '//Number of items
    Items() As EXPR_Sem_Expr_ComplexID_Member
        '//Items array
End Type '//EXPR_Sem_Expr_ComplexID
Public Type EXPR_Sem_Expr_ComplexIDs
    '//EXPR Eval - Semantic Expression - Complex Identifiers Array
    Count As Integer
        '//Number of items
    Items() As EXPR_Sem_Expr_ComplexID
        '//Items array
End Type '//EXPR_Sem_Expr_ComplexIDs

Public Type EXPR_Sem_Expr_Sections
    '//EXPR Eval - Semantic Expression - Expression Sections
    Count As Integer
        '//Number of sections
    Items() As EXPR_Sem_Expr_Section
        '//Sections array
End Type

Public Type EXPR_SEM_Expression_Constant
    '//EXPR Eval - Semantic Expression - Constant
    SubType         As Byte 'EXPR_Sem_Expr_Const_Type
        '//Const Sub type
    Index           As Long
        '//Sub Structure index
End Type
Public Type EXPR_SEM_Expression_Constants
    '//EXPR Eval - Semantic Expression - Constants array
    Count As Integer
        '//Number of members
    Items() As EXPR_SEM_Expression_Constant
        '//Items array
End Type

Public Type EXPR_Sem_Expression
    '//EXPR Eval - Semantic Expression
    FirstSection As EXPR_Sem_Expr_Section
        '//First expression section
    ComplexIdentifiers As EXPR_Sem_Expr_ComplexIDs
        '//Complex Identifiers
    ArgumentLists As EXPR_Sem_Expr_ComplexID_ArgLists
        '//Argument Lists
    Constants As EXPR_SEM_Expression_Constants
        '//Constants
    Subsections As EXPR_Sem_Expr_Sections
        '//Sub Sections
End Type

Public Type EXPR_SEM_Intermediate_Exception
    '//EXPR Eval - Semantic Intermediate - Exception
    Position As Long
        '//Position the error occured at
    Error As Long
        '//Error code
    Reason As Long
        '//String reason index
End Type
Public Type EXPR_SEM_Intermediate_Exceptions
    '//EXPR Eval - Semantic Intermediate - Exceptions array
    Count As Integer
        '//Number of exceptions
    Items() As EXPR_SEM_Intermediate_Exception
        '//Exceptions array
End Type

Public Type EXPR_Sem_Intermediate
    '//EXPR Eval - Semantic Intermediate
    ConstTable As EXPR_ConstTable
        '//Constant table
    Exceptions As EXPR_SEM_Intermediate_Exceptions
        '//Exceptions or errors
End Type
