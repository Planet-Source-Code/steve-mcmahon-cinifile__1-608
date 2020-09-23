<div align="center">

## cIniFile


</div>

### Description

Complete access to INI files through a simple class module, which works with VB4 16,32 and VB5. This class module allows you to read/write INI values, delete values, delete sections and query whole sections through a simple inteface.
 
### More Info
 
Here is a sample of using the cIniFile class:

dim cIni as new cIniFile

with cIni

.Path = "C:\WINDOWS\SYSTEM.INI"   ' Use GetWindowsDir() call to find the correct dir

.Section = "boot"  ' Look under the section headed [boot]

.Key = "shell"     ' Search for shell=

if (ucase$(trim$(.Value)) = "EXPLORER.EXE") then ' Get the section value

msgbox "Da Shell is here",vbInformation

else

msgbox "Da Computer is too old....",vbExclamation

endif

' end with

Save the code into a file called cIniFile.cls and add it to your project. Follow the sample code in the top comment block to try it out.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Steve McMahon](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/steve-mcmahon.md)
**Level**          |Unknown
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/steve-mcmahon-cinifile__1-608/archive/master.zip)





### Source Code

```
Option Explicit
' *************************************************************************************
' Description:
' A complete class for access to Ini Files. Works in
' VB4 16 and 32 and VB5.
'
' Sample code: find out whether we are running the Windows
' 95 shell or not:
'
' dim cIni as new cIniFile
' with cIni
'  .Path = "C:\WINDOWS\SYSTEM.INI"   ' Use GetWindowsDir() call to find the correct dir
'  .Section = "boot"
'  .Key = "shell"
'  if (ucase$(trim$(.Value)) = "EXPLORER.EXE") then
'    msgbox "Da Shell is here",vbInformation
'  else
'    msgbox "Da Computer is too old..",vbExclamation
'  endif
' end with
'
' FileName: cIniFile.Cls
' Author:  Steve McMahon (Steve-McMahon@pa-consulting.com)
' Date:   30 June 1997
' *************************************************************************************
' Private variables to store the settings made:
Private m_sPath As String
Private m_sKey As String
Private m_sSection As String
Private m_sDefault As String
Private m_lLastReturnCode As Long
' Declares for cIniFile:
#If Win32 Then
  ' Profile String functions:
  Private Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
  Private Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
#Else
  ' Profile String functions:
  ' If you are developing in VB5, delete this section
  ' otherwise SetupKit gets **confused**!
  Private Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Integer
  Private Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
#End If
Property Get LastReturnCode() As Long
  ' Did the last call succeed?
  ' 0 if not!
  LastReturnCode = m_lLastReturnCode
End Property
Property Let Default(sDefault As String)
  ' What to return if something goes wrong:
  m_sDefault = sDefault
End Property
Property Get Default() As String
  ' What to return if something goes wrong:
  Default = m_sDefault
End Property
Property Let Path(sPath As String)
  ' The filename of the INI file:
  m_sPath = sPath
End Property
Property Get Path() As String
  ' The filename of the INI file:
  Path = m_sPath
End Property
Property Let Key(sKey As String)
  ' The KEY= bit to look for
  m_sKey = sKey
End Property
Property Get Key() As String
  ' The KEY= bit to look for
  Key = m_sKey
End Property
Property Let Section(sSection As String)
  ' The [SECTION] bit to look for
  m_sSection = sSection
End Property
Property Get Section() As String
  ' The [SECTION] bit to look for
  Section = m_sSection
End Property
Property Get Value() As String
  ' Get the value of the current Key within Section of Path
Dim sBuf As String
Dim iSize As String
Dim iRetCode As Integer
  sBuf = Space$(255)
  iSize = Len(sBuf)
  iRetCode = GetPrivateProfileString(m_sSection, m_sKey, m_sDefault, sBuf, iSize, m_sPath)
  If (iSize > 0) Then
    Value = Left$(sBuf, iRetCode)
  Else
    Value = ""
  End If
End Property
Property Let Value(sValue As String)
  ' Set the value of the current Key within Section of Path
Dim iPos As Integer
  ' Strip chr$(0):
  iPos = InStr(sValue, Chr$(0))
  Do While iPos <> 0
    sValue = Left$(sValue, (iPos - 1)) & Mid$(sValue, (iPos + 1))
    iPos = InStr(sValue, Chr$(0))
  Loop
  m_lLastReturnCode = WritePrivateProfileString(m_sSection, m_sKey, sValue, m_sPath)
End Property
Public Sub DeleteValue()
  ' Delete the value at Key within Section of Path
  m_lLastReturnCode = WritePrivateProfileString(m_sSection, m_sKey, 0&, m_sPath)
End Sub
Public Sub DeleteSection()
  ' Delete the Section in Path
  m_lLastReturnCode = WritePrivateProfileString(m_sSection, 0&, 0&, m_sPath)
End Sub
Property Get INISection() As String
  ' Return all the keys and values within the current
  ' section, separated by chr$(0):
Dim sBuf As String
Dim iSize As String
Dim iRetCode As Integer
  sBuf = Space$(255)
  iSize = Len(sBuf)
  iRetCode = GetPrivateProfileString(m_sSection, 0&, m_sDefault, sBuf, iSize, m_sPath)
  If (iSize > 0) Then
    INISection = Left$(sBuf, iRetCode)
  Else
    INISection = ""
  End If
End Property
Property Let INISection(sSection As String)
  ' Set one or more the keys within the current section.
  ' Keys and Values should be separated by chr$(0):
  m_lLastReturnCode = WritePrivateProfileString(m_sSection, 0&, sSection, m_sPath)
End Property
```

