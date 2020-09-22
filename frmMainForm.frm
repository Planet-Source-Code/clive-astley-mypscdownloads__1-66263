VERSION 5.00
Begin VB.Form frmMainForm 
   Caption         =   "My PSC Downloads"
   ClientHeight    =   6600
   ClientLeft      =   660
   ClientTop       =   1455
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6600
   ScaleWidth      =   10890
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox txtFilter 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   5520
      Width           =   4455
   End
   Begin VB.OptionButton optFilter 
      Caption         =   "All of these words"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   6000
      Width           =   1575
   End
   Begin VB.OptionButton optFilter 
      Caption         =   "Any of these words"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   7
      Top             =   6000
      Width           =   1695
   End
   Begin VB.OptionButton optFilter 
      Caption         =   "Exact phrase"
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   6
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdFilter 
      Caption         =   "Filter"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   5520
      Width           =   615
   End
   Begin VB.CommandButton cmdOpenWeb 
      Caption         =   "Open PSC Web Page"
      Height          =   375
      Left            =   8760
      TabIndex        =   4
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmdLaunchExp 
      Caption         =   "Open Code In Windows Explorer"
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   5520
      Width           =   2775
   End
   Begin VB.TextBox txtDescription 
      Height          =   4095
      Left            =   5880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1200
      Width           =   4815
   End
   Begin VB.TextBox txtTitle 
      Height          =   975
      Left            =   5880
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
   Begin VB.ListBox lstTitles 
      Height          =   5130
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strCodeLocation As String
Dim strWeb As String
Dim db As Database
Dim rs As Recordset
Dim strSQL As String

Private Sub Form_Load()
  
  strSQL = "SELECT * FROM MainTable ORDER BY Title"
  Call FillTitlesListBox(strSQL)
  
  optFilter(0).Value = True
       
End Sub

Private Sub cmdFilter_Click()
  Call Filter
End Sub

Private Sub cmdClear_Click()
  txtFilter = ""
  Call Filter
End Sub

Private Sub cmdLaunchExp_Click()
  'Launch Windows Explorer
  Dim objLaunchExplorer As New cLaunchExplorer
     
  With objLaunchExplorer
    .sPath = strCodeLocation
    .Launch
  End With

End Sub

Private Sub cmdOpenWeb_Click()
  Call OpenLocation(strWeb, vbNormal)
End Sub

Private Sub lstTitles_Click()
  
  strSQL = "SELECT * FROM MainTable WHERE ID = " & Trim(Right$(lstTitles.Text, 10)) & ";"
  Set db = OpenDatabase(strDatabasePath)
  Set rs = db.OpenRecordset(strSQL)
  
  txtTitle = rs("Title")
  txtDescription = rs("Description")
  
  strWeb = rs("Web")
  strCodeLocation = rs("FileListDir")
    
  Set rs = Nothing
  Set db = Nothing
End Sub

Private Sub Filter()
  Dim strSplit() As String
  Dim idx As Integer
  txtTitle = ""
  txtDescription = ""
  If txtFilter = "" Then
    strSQL = "SELECT * FROM MainTable ORDER BY Title"
    Call FillTitlesListBox(strSQL)
    Exit Sub
  End If

  Select Case True
    Case optFilter(0).Value 'All of these words
      strSplit = Split(txtFilter)
      strSQL = "SELECT * FROM MainTable WHERE "
      For idx = LBound(strSplit) To UBound(strSplit)
        strSQL = strSQL & "([Title]&[Description]) Like '* " & strSplit(idx) & " *'"
        If idx <> UBound(strSplit) Then
          strSQL = strSQL & " AND "
        Else
          strSQL = strSQL & ";"
        End If
      Next
      Call FillTitlesListBox(strSQL)
     
    Case optFilter(1).Value 'Any of these words
      strSplit = Split(txtFilter)
      strSQL = "SELECT * FROM MainTable WHERE "
      For idx = LBound(strSplit) To UBound(strSplit)
        strSQL = strSQL & "([Title]&[Description]) Like '* " & strSplit(idx) & " *'"
        If idx <> UBound(strSplit) Then
          strSQL = strSQL & " OR "
        Else
          strSQL = strSQL & ";"
        End If
      Next
      Call FillTitlesListBox(strSQL)
      
    Case optFilter(2).Value 'Exact phrase
      strSQL = "SELECT * FROM MainTable WHERE " & _
        "([Title]&[Description]) Like '* " & txtFilter & " *';"
      Call FillTitlesListBox(strSQL)

  End Select
End Sub

Private Sub FillTitlesListBox(strSQL)

  lstTitles.Clear

  Set db = OpenDatabase(strDatabasePath)
  Set rs = db.OpenRecordset(strSQL)
  
  If rs.RecordCount = 0 Then
    Exit Sub
  End If

   
  rs.MoveLast
  rs.MoveFirst
  
  Do Until rs.EOF
    lstTitles.AddItem rs("Title") & Space$(100) & rs("ID")
    rs.MoveNext
  Loop
  
  lstTitles.Selected(0) = True
  
  Set rs = Nothing
  Set db = Nothing
End Sub






