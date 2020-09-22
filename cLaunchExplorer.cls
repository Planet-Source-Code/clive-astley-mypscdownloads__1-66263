VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLaunchExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' cLaunchExplorer
' 2004/11/19 Copyright © 2004, Larry Rebich, using the DELL8500
' 2004/11/19 larry@larryrebich.com, www.buygold.net

    Option Explicit
    DefLng A-Z
    
    Private Const SW_SHOWNORMAL As Long = 1
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
         ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Private Declare Function GetDesktopWindow Lib "user32" () As Long
    Private msPath As String
    Private mbExplorerFormat As Boolean
    '

Public Property Get bExplorerFormat() As Boolean
    bExplorerFormat = mbExplorerFormat
End Property

Public Property Let bExplorerFormat(ByVal bValue As Boolean)
    mbExplorerFormat = bValue
End Property

Public Property Get sPath() As String
    sPath = msPath
End Property

Public Property Let sPath(ByVal StrValue As String)
    msPath = StrValue
End Property

Public Function Launch(Optional vntPath As Variant, Optional vntExplorerFormat As Variant) As Boolean
' 2004/11/19 Function added by Larry Rebich using the DELL8500 while in Las Vegas, NV
' 2004/11/19 Based upon OpenURL from modConnectToWebSite.bas

    Dim lRtn As Long
    Dim sOpenHow As String
    
    On Error GoTo LaunchEH
    If Not IsMissing(vntPath) Then
        Me.sPath = vntPath
    End If
    If Not IsMissing(vntExplorerFormat) Then
        Me.bExplorerFormat = vntExplorerFormat
    End If
    
    If Me.bExplorerFormat Then
        sOpenHow = "Explore"
    Else
        sOpenHow = "Open"
    End If
    
    lRtn = RunShellExecute(sOpenHow, Me.sPath, 0&, 0&, SW_SHOWNORMAL)
    If lRtn > 32 Then       'success
        Launch = True
    End If
    Exit Function
LaunchEH:
    Err.Raise Err.Number, "cLaunchExplorer:Launch", Err.Description
End Function

Private Function RunShellExecute(sTopic As String, sFile As Variant, sParams As Variant, sDirectory As Variant, nShowCmd As Long) As Long
' 2004/11/19 Function added by Larry Rebich using the DELL8500 while in Las Vegas, NV
    RunShellExecute = ShellExecute(GetDesktopWindow(), sTopic, sFile, sParams, sDirectory, nShowCmd)

End Function

