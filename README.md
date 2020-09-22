<div align="center">

## RegPass


</div>

### Description

This code allows you to enter and read an encrypted password from your VB project by inserting and reading the encrypted password to and from the registry.
 
### More Info
 
Password

Create a Form (Form1) and on it put a Label (Label1), A TextBox (Text1)and 2 Buttons (Command1 and Command2).

Cut and paste the code into the form source

Unencypted Password


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Unknown
**User Rating**    |4.0 (12 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Registry](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/registry__1-36.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/regpass__1-2920/archive/master.zip)

### API Declarations

```
'Paste the following into a Module
Option Explicit
Global Const REG_SZ As Long = 1
Global Const REG_DWORD As Long = 4
Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003
Global Const ERROR_NONE = 0
Global Const ERROR_BADDB = 1
Global Const ERROR_BADKEY = 2
Global Const ERROR_CANTOPEN = 3
Global Const ERROR_CANTREAD = 4
Global Const ERROR_CANTWRITE = 5
Global Const ERROR_OUTOFMEMORY = 6
Global Const ERROR_INVALID_PARAMETER = 7
Global Const ERROR_ACCESS_DENIED = 8
Global Const ERROR_INVALID_PARAMETERS = 87
Global Const ERROR_NO_MORE_ITEMS = 259
Global Const KEY_ALL_ACCESS = &H3F
Global Const REG_OPTION_NON_VOLATILE = 0
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey& Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String)
Private Declare Function RegDeleteValue& Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String)
Function CryptIt(ToCrypt As String, CryptString As String)
 Dim PosS As Long, PosC As Long, TempString As String
 Dim Cutme As Integer
 Cutme = Len(ToCrypt)
 TempString = Space$(Cutme)
 PosC = 1
 For PosS = 1 To Len(ToCrypt)
  If PosC > Len(CryptString) Then PosC = 1
  Mid(TempString, PosS, 1) = Chr$(Asc(Mid(ToCrypt, PosS, 1)) Xor Asc(Mid(CryptString, PosC, 1)))
  If Asc(Mid(TempString, PosS, 1)) = 0 Then Mid(TempString, PosS, 1) = Mid(ToCrypt, PosS, 1)
  PosC = PosC + 1
 Next PosS
 CryptIt = TempString
End Function
Public Function DeleteKey(lPredefinedKey As Long, sKeyName As String)
' Description:
' This Function will Delete a key
'
' Syntax:
' DeleteKey Location, KeyName
'
' Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
' , HKEY_USERS
'
' KeyName is name of the key you wish to delete, it may include subkeys (example "Key1\SubKey1")
 Dim lRetVal As Long   'result of the SetValueEx function
 Dim hKey As Long   'handle of open key
 'open the specified key
 'lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
 lRetVal = RegDeleteKey(lPredefinedKey, sKeyName)
 'RegCloseKey (hKey)
End Function
Public Function DeleteValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
' Description:
' This Function will delete a value
'
' Syntax:
' DeleteValue Location, KeyName, ValueName
'
' Location must equal HKEY_CLAS
```


### Source Code

```
Private Sub Form_Load()
 Text1.Text = ""
End Sub
Private Sub Command1_Click()
 Label1.Caption = ""
 If Command1.Caption = "Set Password" Then
  Command1.Caption = "Check Password"
  Call SetPassword
 Else
  ThisPass = Text1.Text
  Call CheckPassword
 End If
End Sub
Public Sub SetPassword()
 Dim CheckPass, ThisNewPass, PassTempVar As String
 'Check and See if the password has been set
 CheckPass = QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\RegPass", "Pass")
 If CheckPass = "" Then
  'If Not Create a new registry Entry
  Ret = CreateNewKey(HKEY_LOCAL_MACHINE, "Software\RegPass")
 End If
 'Encrypt the String with a static Seed (You can Change this so long as you use the same seed in CheckPassword)
 ThisNewPass = CryptIt(Text1.Text, "ThisSeed")
 'Now Set the Password (Encrypted in the Registry)
 Ret = SetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\RegPass", "Pass", ThisNewPass, REG_SZ)
 Text1.Text = ""
 Label1.Caption = "Password Set"
End Sub
Public Sub CheckPassword()
 ' I could have used 1 variable for this but I think this is less confusing
 Dim CheckPass As String
 Dim ThisLength As Integer
 'Retrieve the Encypted Password from the Regoistry
 CheckPass = CryptIt(QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\RegPass", "Pass"), "ThisSeed")
 'For some strange reason I have to trim a trailing character...????
 ThisLength = Len(CheckPass) - 1
 CheckPass = Left(CheckPass, ThisLength)
 'Check It Against the Entered Password
 If (CheckPass = Text1.Text) Then
  Command1.Caption = "Set Password"
  Label1.Caption = ("Password Check Was Successful")
 Else
  Label1.Caption = ("Incorrect Password...Try Again (was really '" + CheckPass + "')")
  Text1.Text = ""
 End If
End Sub
Private Sub Command2_Click()
 End
End Sub
```

