Attribute VB_Name = "modCommonMeths"
Option Explicit
'//Author: Linguar Amadala (Allen Copeland)

Public Function GetCstVal(ByVal ConstType As EXPR_Sem_Expr_Const_Type, ByRef Table As EXPR_ConstTable, ByVal Index As Long) As Variant
    With Table
        '//Select the table's namespace
        Select Case ConstType
            '//Select the constant type into the case statement
            Case EXPR_Sem_Expr_Const_Type.EXPR_S_E_C_T_Byte
                '//If it's a byte, then...
                GetCstVal = .Bytes.Items(Index)
                    '//Return a byte
            Case EXPR_Sem_Expr_Const_Type.EXPR_S_E_C_T_Currency
                '//If it's a currency value, then...
                GetCstVal = .Currencies.Items(Index)
                    '//Return a currecny
            Case EXPR_Sem_Expr_Const_Type.EXPR_S_E_C_T_Double
                '//If it's a double, then...
                GetCstVal = .Doubles.Items(Index)
                    '//Return a double value
            Case EXPR_Sem_Expr_Const_Type.EXPR_S_E_C_T_Long
                '//If it's a long value, then...
                GetCstVal = .Longs.Items(Index)
                    '//return a long value
            Case EXPR_Sem_Expr_Const_Type.EXPR_S_E_C_T_Integer
                '//If it's an integer value, then...
                GetCstVal = .Integers.Items(Index)
                    '//return an integer value
            Case EXPR_Sem_Expr_Const_Type.EXPR_S_E_C_T_Single
                '//If it's a single value, then...
                GetCstVal = .Singles.Items(Index)
                    '//return a single value
            Case EXPR_S_E_C_T_String
                '//If it's a string, then...
                GetCstVal = CStr(.Strings.Items(Index).Bytes)
                    '//Return a string
        End Select
    End With
End Function

Public Function HasFlag(ByVal Value As Long, ByVal Flag As Long) As Boolean
    HasFlag = ((Value And Flag) = Flag)
        '//Perform binary comparison, returning which bytes are the same in both values
        '//If it returns as the flag then it means all the bytes in the flag
        '//are contained in the value, thus means the flag is in the value
End Function
