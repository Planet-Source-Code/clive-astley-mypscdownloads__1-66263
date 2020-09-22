VERSION 5.00
Begin VB.Form frmSplash 
   Caption         =   "My PSC Downloads"
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   ScaleHeight     =   2700
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Caption         =   "By Clive Astley"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "clive.astley@kingswoodaccounting.co.uk"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label lblWait 
      Caption         =   "Please wait whilst MyPSCDownloads is prepared"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label lblRecords 
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblImporting 
      Caption         =   "Populating the database"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim strDatabaseName As String
  Dim strFilePattern As String
  Dim strPscDirectory As String
  Dim db As Database
  Dim strSQL As String

Private Sub Form_Load()

  strDatabaseName = "MyPSCDownloads.mdb"
  strPscDirectory = "D:\MyPSCDownloadLibrary\" 'Where downloads are stored
  strFilePattern = "@PSC*.txt"

  If Right$(App.Path, 1) = "\" Then
    strDatabasePath = App.Path & strDatabaseName
  Else
    strDatabasePath = App.Path & "\" & strDatabaseName
  End If

  Set db = OpenDatabase(strDatabasePath)
  
  strSQL = "DELETE From MainTable"  'Empties the database
  db.Execute strSQL
   
  db.Close
  Call Compact  'Compacts the database to reset autonumber ID
   
  Me.Show
  'MsgBox "Emptied"
  DoEvents

  Call FindFiles(strPscDirectory, strFilePattern)

  frmMainForm.Show
  Unload Me

End Sub

'FINDFILE FUNCTION FROM ROD STEPHENS BOOK
'(my one modification to call ReadFile() is annotated below
'This extracts all the PSC Readme .txt files from the Download directory
Private Function FindFiles(ByVal start_dir As String, ByVal file_pattern As String) As String
  Dim dirs() As String
  Dim num_dirs As Long
  Dim sub_dir As String
  Dim file_name As String
  Dim i As Integer
  Dim txt As String

  file_name = Dir$(start_dir & file_pattern, vbNormal)
  Do While Len(file_name) > 0
    txt = txt & start_dir & file_name & vbCrLf
    file_name = Dir$(, vbNormal)
  Loop

  sub_dir = Dir$(start_dir & "*", vbDirectory)
  Do While Len(sub_dir) > 0
    If UCase$(sub_dir) <> "PAGEFILE.SYS" And _
      sub_dir <> "." And sub_dir <> ".." _
    Then
      sub_dir = start_dir & sub_dir
        If GetAttr(sub_dir) And vbDirectory Then
          num_dirs = num_dirs + 1
          ReDim Preserve dirs(1 To num_dirs)
          dirs(num_dirs) = sub_dir & "\"
        End If
        End If

      sub_dir = Dir$(, vbDirectory)
  Loop

  For i = 1 To num_dirs
    txt = txt & FindFiles(dirs(i), file_pattern)
    'THE NEXT THREE LINES IS MY ONLY MODIFICATION
    frmSplash.lblRecords.Caption = i
    DoEvents
    Call ReadFile(FindFiles(dirs(i), file_pattern))
  Next i

  FindFiles = txt
  
End Function

Private Sub ReadFile(ByVal strFileName As String)
  Dim strTextToCheck As String
  Dim lngPosTitle As Long
  Dim lngPosTitleCrlf As Long
  Dim lngPosDescription As Long
  Dim lngPosDescriptionCrlf As Long
  Dim lngPosHTTP As Long
  Dim lngPosHTTPcrlf As Long
  Dim lngFileListDir As Long
  Dim strWeb As String
  Dim strTitle As String
  Dim strDescription As String
  Dim strFileListDir As String
  Dim strSQL As String
  
  Set db = OpenDatabase(strDatabasePath)
  
      Dim rs As Recordset
      Set rs = db.OpenRecordset("MainTable")
  
  If strFileName <> "" Then
    strFileName = Left$(strFileName, Len(strFileName) - 2) 'Removes 2 non-displayable characters at end
    strTextToCheck = FileContents(strFileName)

    lngPosTitle = InStr(strTextToCheck, "Title")
    lngPosTitleCrlf = InStr(lngPosTitle, strTextToCheck, vbCrLf)
    strTitle = Mid$(strTextToCheck, lngPosTitle, (lngPosTitleCrlf - lngPosTitle)) & " "
    
    lngPosDescription = InStr(strTextToCheck, "Description")
    lngPosDescriptionCrlf = InStr(lngPosDescription, strTextToCheck, vbCrLf)
    strDescription = Mid$(strTextToCheck, lngPosDescription, (lngPosDescriptionCrlf - lngPosDescription))
    
    lngPosHTTP = InStr(strTextToCheck, "http")
    lngPosHTTPcrlf = InStr(lngPosHTTP, strTextToCheck, vbCrLf)
    strWeb = Mid$(strTextToCheck, lngPosHTTP, (lngPosHTTPcrlf - lngPosHTTP))
    
    lngFileListDir = InStr(strFileName, "\")
    lngFileListDir = InStr(lngFileListDir + 1, strFileName, "\")
    lngFileListDir = InStr(lngFileListDir + 1, strFileName, "\")
    strFileListDir = Left$(strFileName, lngFileListDir)
    
    rs.AddNew
    rs("Title") = strTitle
    rs("Description") = strDescription
    rs("Web") = strWeb
    rs("FileName") = strFileName
    rs("FileListDir") = strFileListDir
    rs.Update
       
  End If
    
End Sub

'THIS FUNCTION FROM ROD STEPHENS BOOK
Private Function FileContents(ByVal filename As String) As String
  Dim fnum As Integer

  On Error GoTo OpenError
  fnum = FreeFile
  Open filename For Input As fnum
  FileContents = Input$(LOF(fnum), #fnum)
  Close fnum
  Exit Function

OpenError:
    MsgBox "Error " & Format$(Err.Number) & _
      " reading file." & vbCrLf & _
        Err.Description
  Exit Function
End Function

Private Sub Compact()
Dim db_name As String
Dim temp_name As String

    db_name = strDatabasePath
    temp_name = db_name & ".temp"
    DAO.DBEngine.CompactDatabase db_name, temp_name
    Kill db_name
    Name temp_name As db_name

End Sub

