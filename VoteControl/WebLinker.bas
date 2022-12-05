Attribute VB_Name = "WebLinker"
Option Explicit

Public Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
    Public Const SW_NORMAL = 1


Public Sub OpenWebsite(strWebsite As String)


    If ShellExecute(&O0, "Open", strWebsite, vbNullString, vbNullString, SW_NORMAL) < 33 Then
        ' Insert Error handling code here
    End If
End Sub

'_________Use____________
'    OpenWebsite (App.Path & "\1.htm")

