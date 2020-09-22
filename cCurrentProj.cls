VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCurrentProj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public ProjPath As String
Public ForceCopy As Boolean
Public UseVersion As Boolean
Public ProjName As String

'local variable(s) to hold property value(s)
Private mvarProjectRS As Recordset 'local copy
Public CurrentID As Long

Public Function SaveCurrentSettings() As Boolean
'  ***Created by
'  ***Bill Jones
'  ***Dig-its   on  09/01/2001

On Error GoTo Err_SaveCurrentSettings_

    With mvarProjectRS
        .MoveFirst
        .Find "ID=" & CurrentID
        If .EOF Then
            'not found ?!
        Else
            '.Fields(0) = CurrentID
            .Fields(1) = ProjName
            .Fields("ProjPath") = ProjPath
            .Fields("Force") = ForceCopy
            .Fields("UseVersion") = UseVersion
            .UpdateBatch
            SaveCurrentSettings = True
        End If
    End With
    
    
Exit_SaveCurrentSettings_:
    On Error GoTo 0
    Exit Function

Err_SaveCurrentSettings_:
    If Err = 3021 Then Resume Next
    MsgBox Err.Description, 0, "SaveCurrentSettings_"
    Resume Exit_SaveCurrentSettings_
    Resume
End Function

Public Property Set ProjectRS(ByVal vData As Recordset)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ProjectRS = Form1
    Set mvarProjectRS = vData
End Property


Public Property Get ProjectRS() As Recordset
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProjectRS
    Set ProjectRS = mvarProjectRS
End Property




Public Function SetUpProjInfo(ProjID As Long) As Boolean
'  ***Created by
'  ***Bill Jones
'  ***Dig-its   on  09/01/2001

On Error GoTo Err_SetUpProjInfo_

    With mvarProjectRS
        .MoveFirst
        .Find "ID=" & ProjID
        If .EOF Then
            'not found ?!
        Else
            CurrentID = .Fields(0)
            ProjName = .Fields(1)
            ProjPath = .Fields("ProjPath")
            ForceCopy = .Fields("Force")
            UseVersion = .Fields("UseVersion")
            SetUpProjInfo = True
        End If
    End With
    
Exit_SetUpProjInfo_:
    On Error GoTo 0
    Exit Function

Err_SetUpProjInfo_:
    If Err = 3021 Then Resume Next
    If Err = 94 Then Resume Next
    MsgBox Err.Description, 0, "SetUpProjInfo_"
    Resume Exit_SetUpProjInfo_
    Resume
End Function

