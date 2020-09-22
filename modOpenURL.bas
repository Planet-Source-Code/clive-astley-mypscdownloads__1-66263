Attribute VB_Name = "ModOpenURL"
Option Explicit

Const SW_SHOWMAXIMIZED = 3
Const SW_SHOWMINIMIZED = 2
Const SW_SHOWDEFAULT = 10
Const SW_SHOWMINNOACTIVE = 7
Const SW_SHOWNORMAL = 1

Private Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hWnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long


'PURPOSE: Opens default browser to display URL

'RETURNS: module handle to executed application or
'Error Code ( < 32) if there is an error

'Use one of the constants in the declarations as
'the window state parameter

'can also be used to open any document associated with
'an application on the system (e.g., passing the name
'of a file with a .doc extension will open that file in Word)

Public Function OpenLocation(URL As String, _
WindowState As Long) As Long

    Dim lHWnd As Long
    Dim lAns As Long

    lAns = ShellExecute(lHWnd, "open", URL, vbNullString, _
    vbNullString, WindowState)
   
    OpenLocation = lAns

    'ALTERNATIVE: if not interested in module handle or error
    'code change return value to boolean; then the above line
    'becomes:

    'OpenLocation = (lAns > 32)

End Function

