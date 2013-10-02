Attribute VB_Name = "modEditor"
Option Explicit

Function GetEditorIP(ByVal Index As Long) As String

    If Index > MAX_EDITORS Then Exit Function
    GetEditorIP = frmServer.EditorSocket(Index).RemoteHostIP
End Function
