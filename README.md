<div align="center">

## Cool Ctrl \+ Alt \+ Del features class\!\!\!


</div>

### Description

A short but handy and useful class module

Desable/Enable Ctrl + Alt + del

Show/ Hide application from Ctrl + Alt + Del list

This is my second submission and it is quite short too so please vote for me or write some comments
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kamen](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kamen.md)
**Level**          |Beginner
**User Rating**    |4.1 (45 globes from 11 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Object Oriented Programming \(OOP\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/object-oriented-programming-oop__1-47.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kamen-cool-ctrl-alt-del-features-class__1-11084/archive/master.zip)

### API Declarations

```
Option Explicit
'Put this in class module
'
'A Class module for the Ctrl + Alt + Del features:
'Remove app from the list
'Restore app to list
'Disable/Enable Ctrl + Alt + Del
'Ex.
'
'Sub A()
'Dim C As New CtrlAltDel
'
'C.RemoveFromList 'remove your application
'     'from the list
'End Sub
'
'Send comments and ideas to kamen@sofianet.net
Private Declare Function SystemParametersInfo _
Lib "user32" Alias "SystemParametersInfoA" ( _
 ByVal uAction As Long, _
 ByVal uParam As Long, _
 lpvParam As Any, _
 ByVal fuWinIni As Long _
) As Long
Private Declare Function GetCurrentProcessId _
Lib "kernel32" () As Long
Private Declare Function RegisterServiceProcess _
Lib "kernel32" ( _
 ByVal dwProcessID As Long, _
 ByVal dwType As Long _
) As Long
Const SPI_SCREENSAVERRUNNING = 97
Const RSP_SIMPLE_SERVICE = 1
Const RSP_UNREGISTER_SERVICE = 0
Public Sub RemoveFromList()
 Dim lngProcessID As Long
 lngProcessID = GetCurrentProcessId
 Call RegisterServiceProcess(lngProcessID, RSP_SIMPLE_SERVICE)
End Sub
Public Sub RestoreToList()
 Dim lngProcessID As Long
 lngProcessID = GetCurrentProcessId()
 Call RegisterServiceProcess(lngProcessID, RSP_UNREGISTER_SERVICE)
End Sub
Public Sub Disable()
 Dim cad As Boolean
 Call SystemParametersInfo(SPI_SCREENSAVERRUNNING, _
  True, cad, 0)
End Sub
Public Sub Enable()
 Dim cad As Boolean
 Call SystemParametersInfo(SPI_SCREENSAVERRUNNING, _
  False, cad, 0)
End Sub
```


### Source Code

```
'this stays on the form
Private Sub cmdHide_Click()
 Dim C As New CtrlAltDel
 C.RemoveFromList 'this hide your application
'from the list
End Sub
```

