<div align="center">

## GetFolderPath


</div>

### Description

I've seen requests for this in the mail lists to here it is. This uses the windows scripting runtime to get the path to the requested directory. For example, as you know, the windows directory can be C:\Winnt\ , c:\windows\, etc. This code will retrieve the correct path to the directory.

'Currently written to get the Windows, System32, or Temp directory. Add others as you'd like.
 
### More Info
 
The type of folder to look for (Windows, Windows System, Temp)

No Error Handling. Make sure that you enter your own error handling methods.

The path of the directory requested

Requires the Windows Scripting Runtime. Standard dll for the Windows OS with IE 6 installed.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[SteamboatWilly](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/steamboatwilly.md)
**Level**          |Beginner
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/steamboatwilly-getfolderpath__1-42231/archive/master.zip)

### API Declarations

```
'Also requires a reference to the windows Scripting Runtime.
```


### Source Code

```
Option Explicit
Public Enum FolderType
 fldWindows = 0 'i.e. C:\WINNT\
 fldWinSystem = 1 'i.e. C:\WINNT\SYSTEM32
 fldWinTemp = 2 'i.e. C:\Temp
End Enum
'=================================================
' Function Name: GetFolderPath
' Inputs: The Special Windows Folder to get
' the path from
' Returns: string containing the desired
' directory path
'
' References: Windows Scripting Runtime
'
' Method: objFileSystem.GetSpecialFolder(1)
' Where: 1 = System Folder (ie C:\winnt\system32)
' 2 = Temporary Folder (ie c:\winnt\temp)
' 0 = Windows Folder (ie C:\winnt\)
'
'
'=================================================
Public Function GetFolderPath(FolderType As FolderType) As String
 Dim objFileSystem As Object
 Set objFileSystem = CreateObject("Scripting.FileSystemObject")
 Select Case FolderType
 Case fldWindows 'The Windows Directory
 GetFolderPath objFileSystem.GetSpecialFolder(0)
 Case fldWinSystem 'The Windows System Directory
 GetFolderPath = objFileSystem.GetSpecialFolder(1)
 Case fldWinTemp 'Windows Temp Folder
 GetFolderPath = objFileSystem.GetSpecialFolder(2)
 End Select
 Set objFileSystem = Nothing
End Function
```

