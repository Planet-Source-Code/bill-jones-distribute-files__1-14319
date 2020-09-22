Attribute VB_Name = "ModFiles"
Option Explicit
Public Const FO_COPY = &H2
Public Const FO_DELETE = &H3
Public Const FO_MOVE = &H1
Public Const FO_RENAME = &H4
Public Const FOF_NOCONFIRMATION = &H10       ' Don't prompt the user.
Public Const FOF_NOERRORUI = &H400

Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Function CopyThatFile(psFromFileName As String, psToFileName As String) As Boolean
Dim shfileOP As SHFILEOPSTRUCT, ret&

    With shfileOP
        .wFunc = FO_COPY
        .hWnd = frmMain.hWnd
        .pTo = psToFileName
        .pFrom = psFromFileName
        .fFlags = FOF_NOCONFIRMATION
    End With
    'perform file operation
    ret = SHFileOperation(shfileOP)
    If ret = 0 Then CopyThatFile = True
End Function


