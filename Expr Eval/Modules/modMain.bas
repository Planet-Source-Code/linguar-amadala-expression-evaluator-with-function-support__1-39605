Attribute VB_Name = "modMain"
Option Explicit
Public Declare Function WindowFromPoint _
    Lib "user32" _
        (ByVal X As Long, _
         ByVal Y As Long) _
    As Long
Public Declare Function GetCursorPos _
    Lib "user32" _
        (Point As Point) _
    As Long
Public Type Point
    X As Long
    Y As Long
End Type
'//Author: Linguar Amadala (Allen Copeland)

Public Function MousedOver(Handle As Long) As Boolean
    Dim m_eptPoint As Point
    Dim m_lngWindow As Long
    GetCursorPos m_eptPoint
    With m_eptPoint
        m_lngWindow = WindowFromPoint(.X, .Y)
        MousedOver = CBool(m_lngWindow = Handle)
    End With
End Function
Private Sub Main()
    Set m_colNames = New Collection
    frmMain.Show
End Sub
