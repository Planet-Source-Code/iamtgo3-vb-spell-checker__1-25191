Attribute VB_Name = "Module1"
' This program was written and designed by George Goehring
' Date: 10/22/2000
'
' Interactive PsyberTechnology Developers Group
' CEO: George Goehring
' Developing Products to fit all your computer needs...
' ALL Developed Products Are Y2K Compliant...
' Giving you the tools to build future business today...
' For more information please visit our website for details...
' www.ipdg3.com - info@ipdg3.com - Voice: 630.236.5584
' Aurora, IL. USA

Public Const URL = "http://www.ipdg3.com"
Public Const RESUMEURL = "http://www.ipdg3.com/resume.php"
Public Const email = "info@ipdg3.com"

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

Sub Main()
    
    frmSpellChecker.Show
       
End Sub

Public Sub gotoIPDG3()
Dim Success As Long

    Success = ShellExecute(0&, vbNullString, URL, vbNullString, "C:\", SW_SHOWNORMAL)

End Sub

Public Sub gotoResume()
Dim Success As Long

    Success = ShellExecute(0&, vbNullString, RESUMEURL, vbNullString, "C:\", SW_SHOWNORMAL)

End Sub

Public Sub sendemail()
Dim Success As Long

    Success = ShellExecute(0&, vbNullString, "mailto:" & email, vbNullString, "C:\", SW_SHOWNORMAL)

End Sub


