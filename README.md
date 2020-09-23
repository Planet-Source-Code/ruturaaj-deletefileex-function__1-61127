<div align="center">

## DeleteFileEx Function


</div>

### Description

This function can do following things ...

[1] Show default system confirmation Prompt to move File to Recycle Bin

[2] Move directly the selected File to Recycle Bin without any Prompt

[3] Show default system confirmation Prompt to remove File forever.

[4] Delete File forever without any prompt. (Same lile Kill function)
 
### More Info
 
Returns Boolean value. True if task completed without any error , or else returns False on error.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ruturaaj](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ruturaaj.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ruturaaj-deletefileex-function__1-61127/archive/master.zip)

### API Declarations

```
Private Declare Function SHFileOperation Lib "shell32.dll" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Const FO_DELETE = &amp;H3
Private Const FOF_ALLOWUNDO = &amp;H40
Private Const FOF_NOCONFIRMATION = &amp;H10
Private Type SHFILEOPSTRUCT
  hwnd As Long
  wFunc As Long
  pFrom As String
  pTo As String
  fFlags As Integer
  fAnyOperationsAborted As Long
  hNameMappings As Long
  lpszProgressTitle As String
End Type
```


### Source Code

```
Private Declare Function SHFileOperation Lib "shell32.dll" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_NOCONFIRMATION = &H10
Private Type SHFILEOPSTRUCT
  hwnd As Long
  wFunc As Long
  pFrom As String
  pTo As String
  fFlags As Integer
  fAnyOperationsAborted As Long
  hNameMappings As Long
  lpszProgressTitle As String
End Type
Public Function DeleteFileEx(lHwnd As Long, sFilePathName As String, bToRecycleBin As Boolean, Optional bConfirm As Boolean = True) As Boolean
'---------------------------------------------------------------------------------------
' Author   : Ruturaj
' Email   : mailme_friends@yahoo.com
'=======================================================================================
' Procedure : DeleteFileEx
' Type    : Function
' ReturnType : Boolean
'=======================================================================================
' Arguments : [1] lHwnd      = hWnd Property of calling Form.
'       [2] sFilePathName  = Full path of File which is to be deleted.
'       [3] bToRecycleBin  = Set to True if file is to be moved to Recycle Bin.
'       [4] bConfirm     = Optional. Default is True. Set to False if you
'                   don' want OS Confirmation prompts before
'                   performing Delete Action.
'=======================================================================================
' Purpose  : This function can do following things ...
'       [1] Show default system confirmation Prompt to move File to Recycle Bin
'       [2] Move directly the selected File to Recycle Bin without any Prompt
'       [3] Show default system confirmation Prompt to remove File forever.
'       [4] Delete File forever without any prompt. (Same lile Kill function)
'---------------------------------------------------------------------------------------
  On Error GoTo DeleteFileEx_Error
  Dim TSHStruct As SHFILEOPSTRUCT
  Dim lResult As Long
  'See if File exists ...
  If Dir(sFilePathName, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then
    'Fill the necessary Structure elements by specified values ...
    With TSHStruct
      .hwnd = lHwnd
      .pFrom = sFilePathName
      .wFunc = FO_DELETE
      'Flag settings ... the heart of this function !
      If bToRecycleBin = True Then
        If bConfirm = True Then
          .fFlags = FOF_ALLOWUNDO
        Else
          .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION
        End If
      ElseIf bToRecycleBin = False Then
        If bConfirm = False Then
          .fFlags = FOF_NOCONFIRMATION
        End If
      End If
    End With
    'It's Show-Time !
    lResult = SHFileOperation(TSHStruct)
    'SHFileOperation returns Zero if successful or non-zero if failed.
    If lResult > 0 Then
      DeleteFileEx = False
    Else
      DeleteFileEx = True
    End If
  Else
    DeleteFileEx = False
  End If
  'This will avoid empty error window to appear.
  Exit Function
DeleteFileEx_Error:
  'Show the Error Message with Error Number and its Description.
  MsgBox Err.Number & " : " & vbCrLf & vbCrLf & Err.Description, vbInformation, "Error ! (Source : DeleteFileEx)"
  'Safe Exit from DeleteFileEx
  Exit Function
End Function
```

