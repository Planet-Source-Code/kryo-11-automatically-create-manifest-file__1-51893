<div align="center">

## \_ Automatically Create Manifest File \_


</div>

### Description

Automatically changes controls to XP themed style in XP based OS.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[KRYO\_11](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kryo-11.md)
**Level**          |Intermediate
**User Rating**    |4.0 (16 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kryo-11-automatically-create-manifest-file__1-51893/archive/master.zip)

### API Declarations

```
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type
```


### Source Code

```
Public Function CreateManifest() As Boolean
  On Error Resume Next
  Dim EXEPath As String
  'Get The EXE Path
  EXEPath = App.Path & IIf(Right(App.Path, 1) = "\", vbNullString, "\")
  EXEPath = EXEPath & App.EXEName & IIf(LCase(Right(App.EXEName, 4)) = ".exe", ".manifest", ".exe.manifest")
  'Checks if the manifest has already been created
  If Dir(EXEPath, vbReadOnly Or vbSystem Or vbHidden) <> vbNullString Then GoTo ErrorHandler
  'Makes sure you are using windows xp
  If WinVersion = "Windows XP" Then
    Dim iFileNumber As Integer
    iFileNumber = FreeFile
    'Save the .manifest file
    Open EXEPath For Output As #iFileNumber
    Print #iFileNumber, FormatManifest
    CreateManifest = True
  Else
    Kill EXEPath
  End If
  'set the file to be hidden
  Close #iFileNumber
  SetAttr EXEPath, vbHidden Or vbSystem Or vbReadOnly Or vbArchive
ErrorHandler:
  Call InitCommonControls
End Function
'get windows version (from Microsoft.com)
Private Function WinVersion() As String
  Dim osinfo As OSVERSIONINFO
  Dim retvalue As Integer
  osinfo.dwOSVersionInfoSize = 148
  osinfo.szCSDVersion = Space$(128)
  retvalue = GetVersionExA(osinfo)
  With osinfo
    Select Case .dwPlatformId
      Case 1
        If .dwMinorVersion = 0 Then
          WinVersion = "Windows 95"
        ElseIf .dwMinorVersion = 10 Then
          WinVersion = "Windows 98"
        End If
      Case 2
        If .dwMajorVersion = 3 Then
          WinVersion = "Windows NT 3.51"
        ElseIf .dwMajorVersion = 4 Then
          WinVersion = "Windows NT 4.0"
        ElseIf .dwMajorVersion >= 5 Then
          WinVersion = "Windows XP"
        End If
      Case Else
        WinVersion = "Failed"
    End Select
End With
End Function
'Create the string for the manifest file
Private Function FormatManifest() As String
  Dim Header As String
  Header = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "yes" & Chr(34) & "?>"
  Header = Header & vbCrLf & "<assembly xmlns=" & Chr(34) & "urn:schemas-microsoft-com:asm.v1" & Chr(34) & " manifestVersion=" & Chr(34) & "1.0" & Chr(34) & ">"
  Header = Header & vbCrLf & "<assemblyIdentity"
  Header = Header & vbCrLf & "  version=" & Chr(34) & "1.0.0.0" & Chr(34)
  Header = Header & vbCrLf & "  processorArchitecture=" & Chr(34) & "X86" & Chr(34)
  Header = Header & vbCrLf & "  name=" & Chr(34) & "Microsoft.VisualBasic6.IDE" & Chr(34)
  Header = Header & vbCrLf & "  type=" & Chr(34) & "win32" & Chr(34)
  Header = Header & vbCrLf & "/>"
  Header = Header & vbCrLf & "<description>Microsoft Visual Basic 6 IDE</description>"
  Header = Header & vbCrLf & "<dependency>"
  Header = Header & vbCrLf & "  <dependentAssembly>"
  Header = Header & vbCrLf & "    <assemblyIdentity"
  Header = Header & vbCrLf & "      type=" & Chr(34) & "win32" & Chr(34)
  Header = Header & vbCrLf & "      name=" & Chr(34) & "Microsoft.Windows.Common-Controls" & Chr(34)
  Header = Header & vbCrLf & "      version=" & Chr(34) & "6.0.0.0" & Chr(34)
  Header = Header & vbCrLf & "      processorArchitecture=" & Chr(34) & "X86" & Chr(34)
  Header = Header & vbCrLf & "      publicKeyToken=" & Chr(34) & "6595b64144ccf1df" & Chr(34)
  Header = Header & vbCrLf & "      language=" & Chr(34) & "*" & Chr(34)
  Header = Header & vbCrLf & "    />"
  Header = Header & vbCrLf & "  </dependentAssembly>"
  Header = Header & vbCrLf & "</dependency>"
  Header = Header & vbCrLf & "</assembly>"
  FormatManifest = Header
End Function
```

