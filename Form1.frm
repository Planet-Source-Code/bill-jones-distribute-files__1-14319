VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Update Util"
   ClientHeight    =   7200
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataList1 
      Bindings        =   "Form1.frx":0000
      Height          =   3000
      Left            =   90
      TabIndex        =   11
      Top             =   465
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   5292
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   14737632
      Enabled         =   -1  'True
      ForeColor       =   64
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Distribution Locations"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Machine Path"
         Caption         =   "Machine Path"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "MachineName"
         Caption         =   "MachineName"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "SendTo"
         Caption         =   "SendTo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "True"
            FalseValue      =   "False"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   7
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Button          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            Button          =   -1  'True
            ColumnWidth     =   764.787
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   405
      Left            =   3840
      TabIndex        =   16
      Top             =   30
      Width           =   3330
      Begin VB.CheckBox chkForce 
         Caption         =   "Force Copy"
         Height          =   255
         Left            =   2130
         TabIndex        =   18
         Top             =   120
         Width           =   1110
      End
      Begin VB.CheckBox chkUVC 
         Caption         =   "Use Version Checking"
         Height          =   255
         Left            =   75
         TabIndex        =   17
         Top             =   120
         Width           =   1995
      End
   End
   Begin VB.CommandButton cmdOpenFolder 
      Caption         =   "Open Folder"
      Height          =   375
      Left            =   7320
      TabIndex        =   15
      Top             =   2115
      Width           =   1335
   End
   Begin VB.CommandButton cmdNewProj 
      Caption         =   "New Project"
      Height          =   255
      Left            =   7320
      TabIndex        =   14
      Top             =   120
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Form1.frx":0015
      Height          =   315
      Left            =   795
      TabIndex        =   12
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   8421504
      ForeColor       =   65535
      ListField       =   "ProjectName"
      BoundColumn     =   "ID"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   2760
      Top             =   2280
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      ConnectMode     =   16
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   0
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      ConnectMode     =   16
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   0
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtResponses 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   3210
      Left            =   4620
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   3870
      Width           =   4065
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   120
      TabIndex        =   8
      Top             =   6810
      Width           =   4395
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000C0&
      Height          =   2625
      Left            =   2280
      TabIndex        =   7
      Top             =   4185
      Width           =   2250
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00808080&
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   105
      TabIndex        =   6
      Top             =   3870
      Width           =   4395
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000C0&
      Height          =   2565
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check Selected"
      Height          =   375
      Index           =   2
      Left            =   7320
      TabIndex        =   2
      Top             =   1620
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send Selected"
      Height          =   375
      Index           =   1
      Left            =   7320
      TabIndex        =   1
      Top             =   1110
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send All"
      Height          =   375
      Index           =   0
      Left            =   7305
      TabIndex        =   0
      Top             =   675
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   7305
      TabIndex        =   4
      Top             =   3075
      Width           =   1350
   End
   Begin VB.Label Label3 
      Caption         =   "Projects:"
      Height          =   270
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   945
   End
   Begin VB.Label Label2 
      Caption         =   "Responses"
      Height          =   270
      Left            =   4680
      TabIndex        =   10
      Top             =   3555
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "Select File"
      Height          =   255
      Left            =   105
      TabIndex        =   3
      Top             =   3555
      Width           =   2100
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuProject 
      Caption         =   "&Project"
      Begin VB.Menu mnuFile_NewProj 
         Caption         =   "&New Project"
      End
      Begin VB.Menu mnuFile_DelProj 
         Caption         =   "&Delete Current Project"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTools_UseVersion 
         Caption         =   "&Use Version Checking"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuTools_Force 
         Caption         =   "&Force Copy"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuLocations 
      Caption         =   "&Location"
      Begin VB.Menu mnuTools_AddNewLoc 
         Caption         =   "&Add New Location"
      End
      Begin VB.Menu mnuTools_DelLoc 
         Caption         =   "D&elete Location"
      End
      Begin VB.Menu mnuExplore 
         Caption         =   "&Explore Folder"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTools_BatchSend 
         Caption         =   "&Include In Batch Send"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuViewHelp 
         Caption         =   "&View Help Form"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lButton As Long

Private MouseIs As Boolean
Private Function AddLocation(PathTo As String) As Boolean
    Dim tmpMachName As String
    If CurProj.CurrentID = 0 Then CurProj.CurrentID = DataCombo1.BoundText
    If Len(PathTo) = 0 Then Exit Function
    If Mid$(PathTo, 2, 1) = ":" Then
        tmpMachName = "ThisOne"
    Else
        tmpMachName = Mid$(PathTo, 3, InStr(3, PathTo, "\") - 3)
    End If
    tmpMachName = InputBox("Enter the machines name", "Enter Name", tmpMachName)
    'If tmpMachName = "" Then tmpMachName = x
    If PathTo <> "" And tmpMachName <> "" Then
        With frmMain.Adodc1.Recordset
            .AddNew
                .Fields(1) = PathTo
                .Fields(2) = tmpMachName
                .Fields(3) = True
                .Fields(4) = CurProj.CurrentID
            .UpdateBatch
        End With
        Adodc1.Recordset.Requery
        DataList1.Refresh
    End If

End Function

Private Sub Cancel_Click()
    CopyCancel = True

End Sub

Private Sub UpdateFormProj()
    Dim tbool As Boolean
        
    SourcePath = CurProj.ProjPath
    UpdateSourceDir
    
    mnuTools_UseVersion.Checked = CurProj.UseVersion
    chkUVC.Value = Abs(CInt(CurProj.UseVersion))
    Command1(2).Enabled = CurProj.UseVersion
    
    mnuTools_Force.Checked = CurProj.ForceCopy
    chkForce = Abs(CInt(CurProj.ForceCopy))
End Sub
Private Sub Adodc1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
    If ErrorNumber = 16389 Then fCancelDisplay = True
End Sub

Private Sub chkForce_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mnuTools_Force.Checked = chkForce
    CurProj.ForceCopy = CBool(chkForce)
    CurProj.SaveCurrentSettings

End Sub


Private Sub chkUVC_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mnuTools_UseVersion.Checked = chkUVC
    Command1(2).Enabled = CBool(chkUVC)
    CurProj.UseVersion = CBool(chkUVC)
    CurProj.SaveCurrentSettings

End Sub


Private Sub cmdCancel_Click()
    CopyCancel = True
End Sub

Private Sub cmdNewProj_Click()
    'new project
    mnuFile_NewProj_Click

End Sub

Private Sub cmdOpenFolder_Click()
    mnuExplore_Click
End Sub

Private Sub Command1_Click(Index As Integer)
Dim x As Boolean, RS As ADODB.Recordset, y As Integer
On Error Resume Next
    Screen.MousePointer = vbHourglass
    
    txtResponses = "**OPERATION STARTED..."
    For y = 0 To 2
        Command1(y).Enabled = Not Command1(y).Enabled
    Next
    cmdCancel.Enabled = True
    CopyCancel = False
    
    ProdVer = ""
    CurProj.ProjPath = Dir1.Path
    pFileName = File1
    If pFileName = "" Then
        'MsgBox "No file is selected.  Select a file to distribute, then try again.", vbInformation, "File error"
        For y = 0 To 2
            Command1(y).Enabled = Not Command1(y).Enabled
        Next
        cmdCancel.Enabled = False
        CopyCancel = False
        Screen.MousePointer = 0
        txtResponses = "**ERROR**" + vbCrLf + vbTab + "No file is selected." + vbCrLf + vbTab + "Select a file to distribute," + vbCrLf + vbTab + "then try again." + vbCrLf + txtResponses
        Exit Sub
    End If
    SourcePath = CurProj.ProjPath + "\" & pFileName
    
    txtResponses = vbCrLf + vbCrLf + "Distributing: " + vbCrLf & "  " + SourcePath + vbCrLf + txtResponses
    If CurProj.UseVersion Then
        DisplayVerInfo SourcePath, 0
        VFile1 = ProdVer    'Format$(pFileInfo(0).dwProductVersionMSh) & "." & Format$(pFileInfo(0).dwProductVersionMSl) & "." & Format$(pFileInfo(0).dwProductVersionLSh) & "." & Format$(pFileInfo(0).dwProductVersionLSl)
    Else
        VFile1 = "Version Check Not Enabled"
    End If
    frmMain.Caption = "Current Version = " & VFile1
    Select Case Index
        Case 0
            Set RS = Adodc1.Recordset.Clone
            RS.MoveFirst
            Do While RS.EOF = False
                'DataList1.Text = RS.Fields(1)
                frmMain.Refresh
                If RS.Fields(3) = True Then
                    txtResponses = vbCrLf + RS.Fields(1) + ":" + vbCrLf + DistribFile(RS.Fields(1) + "\" & pFileName) + txtResponses
                End If
                DoEvents
                If CopyCancel Then Exit Do
                RS.MoveNext
            Loop
            RS.Close
        Case 1
            txtResponses = vbCrLf + DataList1.Columns(0).Text + ":" + vbCrLf + DistribFile(DataList1.Columns(0).Text + "\" & pFileName) + txtResponses
        Case 2
            x = CheckFileCopied(SourcePath, DataList1.Columns(0).Text & "\" & pFileName)
            txtResponses = vbTab + GetFileDates(DataList1.Columns(0).Text & "\" & pFileName) + vbCrLf + txtResponses
            If x Then
                txtResponses = DataList1.Columns(0).Text & ":" & vbCrLf & vbTab & "File is up to date." & vbCrLf & vbTab & "Source Version: " & VFile1 & vbCrLf & vbTab & "Destination Version: " & VFile2 & vbCrLf + txtResponses
            Else
                txtResponses = DataList1.Columns(0).Text & ":" & vbCrLf & vbTab & "File is NOT up to date." & vbCrLf & vbTab & "Source Version: " & VFile1 & vbCrLf & vbTab & "Destination Version: " & VFile2 & vbCrLf + txtResponses
            End If
    End Select
    For y = 0 To 2
        Command1(y).Enabled = Not Command1(y).Enabled
    Next
    cmdCancel.Enabled = False
txtResponses = "**OPERATION COMPLETED" + vbCrLf + txtResponses
On Error GoTo 0
    Screen.MousePointer = 0

End Sub


Private Sub DataCombo1_Click(Area As Integer)
    If MouseIs Then Exit Sub
    If Area = dbcAreaList Then
        CurProj.ProjPath = Dir1.Path
        CurProj.SaveCurrentSettings
        CurProj.SetUpProjInfo DataCombo1.BoundText
        
        UpdateFormProj      ' CurProj.CurrentID
        
        Adodc1.CommandType = adCmdText
        Adodc1.RecordSource = ReQryPaths(CurProj.CurrentID)
        Adodc1.Refresh
        
        DataCombo1.Text = CurProj.ProjName
        DataCombo1.SelLength = 0
    End If
End Sub

Private Sub DataCombo1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseIs = True
End Sub


Private Sub DataCombo1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseIs = False
End Sub


Private Sub DataList1_ButtonClick(ByVal ColIndex As Integer)
    Dim temp$
    On Error GoTo Err_ButtonClick
    Select Case ColIndex
        Case 2
            'DataList1.EditActive = True
            DataList1.Columns(2).Value = Trim$(Not DataList1.Columns(2).Value)    'True Then)
            DataList1.EditActive = False
        Case 0
            temp = GetFolderPath
            If temp <> DataList1.Columns(0).Value And temp <> "" Then
                DataList1.EditActive = True
                DataList1.Columns(0) = temp
                DataList1.EditActive = False
            End If
    End Select
    
Err_ButtonClick_Exit:
    On Error GoTo 0
    Exit Sub
Err_ButtonClick:
    Select Case Err
        Case 6160
'            MsgBox "You need to add a new record first.", vbOKOnly, "No Record To Edit"
            AddLocation temp
            Resume
        Case Err
            MsgBox Error, 0, Err & " - Form1_Datalist1_ButtonClick()"
    End Select
    Resume Err_ButtonClick_Exit
    Resume
End Sub

Private Sub DataList1_Error(ByVal DataError As Integer, Response As Integer)
    Debug.Print DataError
    If DataError = 7007 Then Response = 0
    Response = 0
End Sub

Private Sub DataList1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
lButton = Button
End Sub

Private Sub DataList1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = 2 Then
        If DataList1.RowContaining(y) = DataList1.Row And DataList1.Row <> -1 Then
            'mnuTools_BatchSend.Visible = True
            'mnuProject.Visible = False
            mnuTools_BatchSend.Checked = DataList1.Columns(2).Value
            PopupMenu mnuLocations
            'mnuProject.Visible = True
            'mnuTools_BatchSend.Visible = False
        End If
    End If


End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1
    
End Sub

Private Sub File1_Click()
    Text1.Text = File1.Path & "\" & File1.Filename

End Sub

Private Sub File1_PathChange()
    Text1.Text = File1.Path
End Sub


Private Sub Form_Load()
    Dim temp$, x&, Temp2$
    Dim tbool As Boolean
    Dim DBAseConn As Connection
    
'  ***Created by
'  ***Bill Jones
'  ***Dig-its   on  09/01/2001

On Error GoTo Err_Form_Load
    
    temp = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=xxx.mdb"
    temp = ReplaceInString(temp, "xxx.mdb", App.Path & "\distlist.mdb")
        
    Adodc2.Enabled = True
    Adodc2.ConnectionString = temp
    Adodc2.CommandType = adCmdTable
    Adodc2.RecordSource = "ProjectInfo"
    Adodc2.Refresh
    
    Set CurProj.ProjectRS = frmMain.Adodc2.Recordset.Clone
    CurProj.ProjectRS.MoveFirst
    CurProj.CurrentID = CurProj.ProjectRS.Fields(0)
    CurProj.SetUpProjInfo CurProj.CurrentID
    If CurProj.ProjPath = "" Then CurProj.ProjPath = "C:\"
    DataCombo1.Text = CurProj.ProjName
    
    Adodc1.ConnectionString = temp
    Adodc1.Enabled = True
    Adodc1.CommandType = adCmdText
    
    Adodc1.RecordSource = ReQryPaths(CurProj.CurrentID)
    Adodc1.Refresh
    
    DataList1.Columns(0).Width = 3909.906
    
    UpdateFormProj      '; CurProj.CurrentID
    
Exit_Form_Load:
    On Error GoTo 0
    Exit Sub

Err_Form_Load:
    If Err = 3021 Then
        Resume Next
    End If
    MsgBox Err.Description, 0, "Form_Load "
    Resume Exit_Form_Load
    Resume

End Sub

Private Sub UpdateSourceDir()
    If SHFileExists(SourcePath) Then
        Drive1.Drive = Left$(SourcePath, 3)
        DoEvents
        Dir1.Path = SourcePath
    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
    CurProj.ProjPath = Dir1.Path
    CurProj.SaveCurrentSettings
    
    Adodc1.Enabled = False
    Set CurProj.ProjectRS = Nothing
    Set CurProj = Nothing
End Sub


Private Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub


Private Sub mnuExplore_Click()
    If CurProj.CurrentID = 0 Then CurProj.CurrentID = DataCombo1.BoundText
    ShellExecute Me.hWnd, vbNullString, Chr$(34) & DataList1.Columns(0) & Chr$(34), DataList1.Columns(1), DataList1.Columns(1), SW_SHOWNORMAL
End Sub

Private Sub mnuFile_DelProj_Click()
    Dim x%
'  ***Created by
'  ***Bill Jones
'  ***Dig-its   on  09/01/2001

On Error GoTo Err_mnuFile_DelProj_Click


    
    Adodc2.Refresh
    Adodc2.Recordset.Find "id=" & CurProj.CurrentID
    If Adodc2.Recordset.EOF Then Exit Sub
    x = MsgBox("Are you sure you want to delete the Project: " & Trim$(Adodc2.Recordset.Fields(1)) & "?", vbYesNo, "Delete Project?")
    If x = vbYes Then Adodc2.Recordset.Delete Else Exit Sub
    With Adodc1.Recordset
        If Not .EOF Then
            .MoveFirst
            Do While .EOF = False
                .Delete
                .MoveNext
            Loop
        End If
        Adodc2.Recordset.UpdateBatch
        Adodc2.Recordset.MoveFirst
        
        Set CurProj.ProjectRS = Nothing
        Set CurProj.ProjectRS = Adodc2.Recordset.Clone
        
        CurProj.SetUpProjInfo Adodc2.Recordset.Fields(0)
        UpdateFormProj
        DataCombo1.ReFill
        DataCombo1.Text = CurProj.ProjName
        
        Adodc1.CommandType = adCmdText
        Adodc1.RecordSource = ReQryPaths(Adodc2.Recordset.Fields(0))
        Adodc1.Refresh
        DataCombo1.SelLength = 0
    End With
    
Exit_mnuFile_DelProj_Click:
    On Error GoTo 0
    Exit Sub

Err_mnuFile_DelProj_Click:
    If Err = 3021 Then
        CurProj.ProjName = ""
        Resume Next
    End If
    MsgBox Err.Description, 0, "mnuFile_DelProj_Click"
    Resume Exit_mnuFile_DelProj_Click
    Resume
End Sub

Private Sub mnuFile_NewProj_Click()
    Dim temp$, x&, RS As ADODB.Recordset
    temp = InputBox("Enter New Project Name.", "New Project", "New Project")
    If temp <> "" Then
        Set RS = Adodc2.Recordset.Clone
        With RS
            .AddNew
            .Fields(1) = Trim$(temp)
            .UpdateBatch
            Set CurProj.ProjectRS = Nothing
            Set CurProj.ProjectRS = RS.Clone
            .Close
        End With
        Adodc2.Refresh  '.Recordset.Requery
        
        With CurProj.ProjectRS
            .Requery
            .MoveFirst
            Do While .Fields(1) <> temp
                .MoveNext
            Loop
            x = .Fields(0)
        End With
        CurProj.SetUpProjInfo x
        UpdateFormProj      ' x
        
        Adodc1.RecordSource = ReQryPaths(x)
        Adodc1.Refresh
        
        DataCombo1 = CurProj.ProjName
        DataCombo1.SelStart = 1
        DataCombo1.SelLength = 0
    End If
End Sub

Private Sub mnuTools_AddNewLoc_Click()
    Dim temp$
    temp = GetFolderPath
    AddLocation temp
End Sub

Private Sub mnuTools_BatchSend_Click()
    DataList1.Columns(2).Value = Not DataList1.Columns(2).Value
    mnuTools_BatchSend.Checked = DataList1.Columns(2).Value
End Sub


Private Sub mnuTools_DelLoc_Click()
    Dim x&
    x = MsgBox("Are you sure?", vbYesNo + vbQuestion, "This Is Your Final Answer!")
    On Error Resume Next
    If x = vbYes Then
        Adodc1.Recordset.Delete
    End If
    On Error GoTo 0
    DataList1.Refresh
End Sub


Private Sub mnuTools_Force_Click()
    mnuTools_Force.Checked = Not mnuTools_Force.Checked
    CurProj.ForceCopy = mnuTools_Force.Checked
    chkForce = CurProj.ForceCopy
    CurProj.SaveCurrentSettings
End Sub

Private Sub mnuTools_UseVersion_Click()
    mnuTools_UseVersion.Checked = Not mnuTools_UseVersion.Checked
    chkUVC = True
    CurProj.UseVersion = chkUVC
    Command1(2).Enabled = CurProj.UseVersion
    CurProj.SaveCurrentSettings
End Sub


Private Sub mnuViewHelp_Click()
    frmHelp.Show 1
End Sub


