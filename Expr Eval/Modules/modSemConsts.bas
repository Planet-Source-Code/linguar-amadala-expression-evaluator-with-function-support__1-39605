Attribute VB_Name = "modSemConsts"
Option Explicit
'//Author: Linguar Amadala (Allen Copeland)

Public Function AddLong(Table As EXPR_ConstTable, ByVal Val As Long) As Long
    With Table
        With .Longs
            If .Count = 0 Then
                ReDim .Items(.Count)
            Else
                ReDim Preserve .Items(.Count)
            End If
            .Items(.Count) = Val
            AddLong = .Count
            .Count = .Count + 1
        End With
    End With
End Function

Public Function AddString(ByRef ConstTable As EXPR_ConstTable, ByVal Expression As String, Optional ByVal CaseSensitive As Boolean = True) As Long
    Dim m_lngLoop As Long
    Dim m_booFound As Boolean
    Dim m_strValue As String
    Dim m_lngIndex As Long
    Dim m_strItem As String
    If CaseSensitive Then
        m_strValue = Expression
    Else
        m_strValue = LCase$(Expression)
    End If
    For m_lngLoop = 0 To ConstTable.Strings.Count - 1
        m_strItem = CStr(ConstTable.Strings.Items(m_lngLoop).Bytes)
        If Not CaseSensitive Then _
            m_strItem = LCase$(m_strItem)
        If (m_strItem = m_strValue) Then
            m_lngIndex = m_lngLoop
            m_booFound = True
            Exit For
        End If
    Next
    If (Not (m_booFound)) Then
        With ConstTable.Strings
            If .Count = 0 Then
                ReDim .Items(.Count)
            Else
                ReDim Preserve .Items(.Count)
            End If
            m_lngIndex = .Count
            .Items(.Count).Bytes = Expression
            .Count = .Count + 1
        End With
    End If
    AddString = m_lngIndex
End Function


Public Function AddInteger(Table As EXPR_ConstTable, ByVal Val As Integer) As Long
    With Table
        With .Integers
            If .Count = 0 Then
                ReDim .Items(.Count)
            Else
                ReDim Preserve .Items(.Count)
            End If
            .Items(.Count) = Val
            AddInteger = .Count
            .Count = .Count + 1
        End With
    End With
End Function

Public Function AddByte(Table As EXPR_ConstTable, ByVal Val As Byte) As Long
    With Table
        With .Bytes
            If .Count = 0 Then
                ReDim .Items(.Count)
            Else
                ReDim Preserve .Items(.Count)
            End If
            .Items(.Count) = Val
            AddByte = .Count
            .Count = .Count + 1
        End With
    End With
End Function

Public Function AddDouble(Table As EXPR_ConstTable, ByVal Val As Double) As Long
    With Table
        With .Doubles
            If .Count = 0 Then
                ReDim .Items(.Count)
            Else
                ReDim Preserve .Items(.Count)
            End If
            .Items(.Count) = Val
            AddDouble = .Count
            .Count = .Count + 1
        End With
    End With
End Function

Public Function AddSingle(Table As EXPR_ConstTable, ByVal Val As Single) As Long
    With Table
        With .Singles
            If .Count = 0 Then
                ReDim .Items(.Count)
            Else
                ReDim Preserve .Items(.Count)
            End If
            .Items(.Count) = Val
            AddSingle = .Count
            .Count = .Count + 1
        End With
    End With
End Function

Public Function AddCurrency(Table As EXPR_ConstTable, ByVal Val As Currency) As Long
    With Table
        With .Currencies
            If .Count = 0 Then
                ReDim .Items(.Count)
            Else
                ReDim Preserve .Items(.Count)
            End If
            .Items(.Count) = Val
            AddCurrency = .Count
            .Count = .Count + 1
        End With
    End With
End Function
