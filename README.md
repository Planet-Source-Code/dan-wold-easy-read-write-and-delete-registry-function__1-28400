<div align="center">

## Easy Read, Write, And Delete Registry Function


</div>

### Description

This code allows you to Read, Write, And Delete Keys from the Registry. EASY!
 
### More Info
 
BEWARE! Editing your registry may cause harm if you dont know what your doing!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dan Wold](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dan-wold.md)
**Level**          |Intermediate
**User Rating**    |5.0 (30 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[Registry](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/registry__1-36.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dan-wold-easy-read-write-and-delete-registry-function__1-28400/archive/master.zip)

### API Declarations

None!


### Source Code

```
'This was created by Dan Wold (Me)
'This code is open source, feel free to use it in Anything...
'You Dont even need to Include my name.. But it would be nice
'My Email is e_man_dan@hotmail.com If you have any questions email me
' To use this code Ill show ya below ;)
' Sorry If I go into to much detail below, Just trying to make my point ;)
'********************************************************************
'Nothing special here, just reads your Windows ProductId depending on os, gave ya NT and Windows
'So dont complain ;)
'For WinNt
'Call ReadWriteDeleteRegistry("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductId", "", 1)
'For Windows
'Call ReadWriteDeleteRegistry("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\ProductId", "", 1)
'*********************************************************************
'********************************************************************
'Write to the Registry
'Call ReadWriteDeleteRegistry("HKEY_CURRENT_USER\Software\TestKey", "Stuff to write here", 2)
'********************************************************************
'********************************************************************
'Delete From Registry
'Call ReadWriteDeleteRegistry("HKEY_CURRENT_USER\Software\TestKey", "", 3)
'********************************************************************
'********************************************************************
'Read = 1
'Write = 2
'Delete = 3
'Reads Key
'Call ReadWriteDeleteRegistry("HKEY_CURRENT_USER\Software\TestKey", "", 1)
'Writes Key
'Call ReadWriteDeleteRegistry("HKEY_CURRENT_USER\Software\TestKey", "Stuff to write", 2)
'Deletes Key
'Call ReadWriteDeleteRegistry("HKEY_CURRENT_USER\Software\TestKey", "", 3)
'********************************************************************
'ReadWriteDelete Read = 1, Write = 2, Delete = 3
Public Sub ReadWriteDeleteRegistry(RegistryKey As String, RegistryInformation As String, ReadWriteDelete As Integer)
'Error Handling In case something goes wrong
On Error GoTo ErrHandler
'Sets the Variables to be used
Dim WSHShell, RegTemp
'Starts the Wscript Object
Set WSHShell = CreateObject("WScript.Shell")
 'Checks for Read Property (Read = 1)
 If ReadWriteDelete = 1 Then
 'Reads the specified key
 RegTemp = WSHShell.RegRead(RegistryKey)
 MsgBox RegTemp
 End If
 'Checks for Write Property (Write = 2)
 If ReadWriteDelete = 2 Then
 'Writes to the registry
 WSHShell.RegWrite RegistryKey, RegistryInformation
 MsgBox Chr(34) & RegistryKey & "\" & RegistryInformation & Chr(34) & " has been written to the registry.", vbInformation, "Success"
 End If
 'Checks for Delete Property (Delete = 3)
 If ReadWriteDelete = 3 Then
 Dim MsgDeleteKey As String
 'Makes sure you really do want to delete this key
 MsgDeleteKey = MsgBox("You are about to delete: " & RegistryKey & " From Your Registry, Do you wish to continue?", vbYesNo Or vbQuestion, "Warning!")
 End If
'Checks for which buttin the user pressed (Yes Or No)
Select Case MsgDeleteKey
 'If Yes, Delete Key
 Case vbYes
 WSHShell.RegDelete (RegistryKey)
 MsgBox RegistryKey & " Has Been Deleted!", vbInformation, "Success"
 Err
 'If No, Exit the Sub
 Case vbNo
 Exit Sub
End Select
'Error Handler Label
ErrHandler:
'Checks for specific error(s)
 'This one is for a non-existant Key
 If Err.Number = (-2147024894) Then
 MsgBox "The Registry Key (" & RegistryKey & ") doesn't exist.", vbCritical, "Error - Key Not Found"
 End If
 'This one is for an Invalid Key
 If Err.Number = (-2147024893) Then
 MsgBox "The Key (" & RegistryKey & ") is invalid.", vbCritical, "Error - Invalid Key"
 End If
End Sub
```

