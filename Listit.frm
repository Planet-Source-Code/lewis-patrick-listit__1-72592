VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "File Check"
   ClientHeight    =   5700
   ClientLeft      =   12660
   ClientTop       =   1020
   ClientWidth     =   2325
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   2325
   Begin VB.Timer Timer1 
      Left            =   1800
      Top             =   5040
   End
   Begin VB.ListBox lstList2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4560
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   5160
      Width           =   975
   End
   Begin VB.ListBox lstList1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4560
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit
  Dim strInTable As String
  Dim strInCode As String
  Dim strWork As String
  Dim strWork1 As String
  Dim dbMyDatabase As DAO.Database
  Dim wsMyWorkspace As DAO.Workspace
  Dim rsTable As DAO.Recordset
  Dim strPassword As String
  Dim strDataBaseName As String
  Dim strTable(50) As String
  Dim strTableID(50) As String
  Dim strTablesub As Integer
  Dim intWorksub As Integer
Private Sub cmdEnd_Click()
    rsTable.Close
    dbMyDatabase.Close
    End
End Sub
Private Sub Form_Load()
    strPassword = "TestPass"
    strDataBaseName = "SpecialMDB.mdb"
    SetOnTop
    OpenDB
    BuildTable
End Sub
Private Sub lstList1_Click()
    intWorksub = lstList1.ListIndex
    If intWorksub = -1 Then Exit Sub
    If strTable(intWorksub) = "" Then
        Beep
        MsgBox "No Data in Field", vbCritical, "No Data"
        lstList1.ListIndex = -1
        Exit Sub
        End If
    Clipboard.Clear
    Clipboard.SetText strTable(intWorksub)
    Timer1.Interval = 1000
End Sub
Private Sub lstList2_Click()
    intWorksub = lstList2.ListIndex
    If lstList2.ListIndex = -1 Then Exit Sub
    If intWorksub = strTablesub Then
        lstList2.List(intWorksub) = ""
        Exit Sub
        End If
    If strTableID(intWorksub) = "" Then
        Beep
        MsgBox "No Data in Field", vbCritical, "No Data"
        lstList2.ListIndex = -1
        Exit Sub
        End If
    Clipboard.Clear
    Clipboard.SetText strTableID(intWorksub)
    Timer1.Interval = 1000
End Sub
Private Sub SetOnTop()
  Dim lR As Long
    lR = SetTopMostWindow(Form1.hwnd, True)
End Sub
    Sub OpenDB()
    Set wsMyWorkspace = DBEngine.Workspaces(0)
    strWork = App.Path + "\" + strDataBaseName
    strWork1 = "MS Access;PWD=" + strPassword
    Set dbMyDatabase = wsMyWorkspace.OpenDatabase(strWork, False, False, strWork1)
    Set rsTable = dbMyDatabase.OpenRecordset("Table", dbOpenTable)
    rsTable.Index = "PrimaryKey"
End Sub
Public Sub BuildTable()
    strTablesub = -1
    Do Until rsTable.EOF
        strWork = ""
        If Not IsNull(rsTable!Title) Then strWork = rsTable!Title
        lstList1.AddItem strWork
        strWork = ""
        If Not IsNull(rsTable!ID) Then strWork = "X"
        lstList2.AddItem strWork
        strTablesub = strTablesub + 1
        If Not IsNull(rsTable!Code) Then strTable(strTablesub) = rsTable!Code
        If Not IsNull(rsTable!ID) Then strTableID(strTablesub) = rsTable!ID
        rsTable.MoveNext
        Loop
End Sub

Private Sub Timer1_Timer()
    lstList1.ListIndex = -1
    lstList2.ListIndex = -1
    Timer1.Interval = 0
End Sub
