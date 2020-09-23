<div align="center">

## Autohighlight active control \(SDI/more than one  Form\)


</div>

### Description

This is a very simple and useful solution to highlight input controls without writting a function for each control. Only include a module with the code shown below and call SetHook at the beginning of your application and Unhook at the end. Please vote, if you think its a good solution.
 
### More Info
 
When running this progam in the IDE do not use the STOP-Button to exit the program, because the unhook function will not be executed and the IDE crashes!!!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Marcel A\. Fritsch](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/marcel-a-fritsch.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/marcel-a-fritsch-autohighlight-active-control-sdi-more-than-one-form__1-32428/archive/master.zip)





### Source Code

```
Option Explicit
' USER32 functions
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" _
 (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" _
 (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
' KERNEL32 functions
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
 (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
' CONSTANTS
Private Const WH_CALLWNDPROC = 4
Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8
' STRUCTS
Private Type CWPSTRUCT
 lParam As Long
 wParam As Long
 message As Long
 hwnd As Long
End Type
' REST
Private hHook As Long
'
Public Function SetHook()
If Not hHook Then
 hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf WndProc, App.hInstance, App.ThreadID)
End If
End Function
'
Private Function WndProc(ByVal idHook As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 Dim CWP As CWPSTRUCT
 Dim C As Control
 Dim F As Form
 Dim found As Boolean
 On Local Error Resume Next
 CopyMemory CWP, ByVal lParam, Len(CWP)
 WndProc = CallNextHookEx(hHook, idHook, wParam, ByVal lParam)
 Select Case CWP.message
 Case WM_SETFOCUS, WM_KILLFOCUS
  For Each F In Forms
  For Each C In F.Controls
   Err.Clear
   If CWP.hwnd = C.hwnd Then
    If Err.Number = 0 Then
     If CWP.message = WM_SETFOCUS Then
      If (TypeOf C Is TextBox) Or _
      (TypeOf C Is ComboBox) Then
       C.BackColor = &H80000018
      End If
     Else
      If (TypeOf C Is TextBox) Or _
      (TypeOf C Is ComboBox) Then
       C.BackColor = &H80000005
      End If
     End If
     found = True
     Exit For
    End If
   End If
  Next
  If found Then
   Exit For
  End If
  Next
 End Select
End Function
'
Public Function UnHook()
If hHook Then
 UnhookWindowsHookEx hHook
End If
End Function
```

