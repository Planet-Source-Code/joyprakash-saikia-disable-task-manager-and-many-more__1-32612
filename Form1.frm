VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   405
      Left            =   1260
      TabIndex        =   6
      Top             =   4440
      Width           =   1665
   End
   Begin VB.Frame Options 
      BackColor       =   &H00000000&
      Caption         =   "Options to Change Ctl+alt+del dialog"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3165
      Left            =   360
      TabIndex        =   0
      Top             =   1050
      Width           =   4035
      Begin VB.CheckBox Check5 
         BackColor       =   &H00000000&
         Caption         =   "Disable Change Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   435
         Left            =   360
         TabIndex        =   5
         Top             =   2550
         Width           =   3435
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00000000&
         Caption         =   "Disable Lock Computer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   465
         Left            =   360
         TabIndex        =   4
         Top             =   2010
         Width           =   3435
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00000000&
         Caption         =   "Disable Shutdown option"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   435
         Left            =   360
         TabIndex        =   3
         Top             =   1500
         Width           =   3435
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00000000&
         Caption         =   "Disable logoff"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   435
         Left            =   360
         TabIndex        =   2
         Top             =   990
         Width           =   3435
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Disable taskmanager"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   525
         Left            =   360
         TabIndex        =   1
         Top             =   390
         Width           =   3435
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "To verify Check each option and  press CTRL+ATL+DEL  to see the Effect "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'File Name          :   form1.frm
'Function           :   To Disable Task manager and many other options on Windows NT & windows 2000
'Created By         :   Joyprakash Saikia
'Created on         :   12th March, 2002

' This Piece of Code has been tested on Windows 2000, Windows NT
' I have not able to test it on windows 9X system

' I think this Code will help you in many situations where you
' have to restrict user from Logoff and Change Password etc.

' Please vote me or Comment on this approach
Const REG_SZ = 1
Const REG_BINARY = 3
Const REG_DWORD = 4

Const HKEY_CURRENT_USER = &H80000001
' the Functions for Registry Manipulations
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
'----------------------------------------------------------------------------
'Argument       :   Handlekey, Name of the Value in side the key
'Return Value   :   String
'Function       :   To fetch the value from a key in the Registry
'Comments       :   on Success , returns the Value else empty String
    '----------------------------------------------------------------------------
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    
    lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Then
    
            strBuf = String(lDataBufSize, Chr$(0))
            'retrieve the key's value
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
    
                RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If
        ElseIf lValueType = REG_BINARY Then
            Dim strData As Integer
            'retrieve the key's value
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strData, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strData
            End If
         ElseIf lValueType = REG_DWORD Then
           
            'retrieve the key's value
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strData, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strData
            End If
            
        End If
    End If
End Function

Function GetString(hKey As Long, strPath As String, strValue As String)
'----------------------------------------------------------------------------
'Argument       :   Handlekey, path from the root , Name of the Value in side the key
'Return Value   :   String
'Function       :   To fetch the value from a key in the Registry
'Comments       :   on Success , returns the Value else empty String
'----------------------------------------------------------------------------

    Dim Ret
    'Open  key
    RegOpenKey hKey, strPath, Ret
    'Get content
    GetString = RegQueryStringValue(Ret, strValue)
    'Close the key
    RegCloseKey Ret
End Function

Sub SaveStringWORD(hKey As Long, strPath As String, strValue As String, strData As String)
'----------------------------------------------------------------------------
'Argument       :   Handlekey, Name of the Value in side the key
'Return Value   :   Nil
'Function       :   To store the value into a key in the Registry
'Comments       :   None
'----------------------------------------------------------------------------

    Dim Ret
    'Create a new key
    RegCreateKey hKey, strPath, Ret
    'Set the key's value
    RegSetValueEx Ret, strValue, 0, REG_DWORD, CLng(strData), 4
    'close the key
    RegCloseKey Ret
End Sub
Sub DelSetting(hKey As Long, strPath As String, strValue As String)
    'Not used in this form
    'you can use it to delete the current entries

    Dim Ret
    'Create a new key
    RegCreateKey hKey, strPath, Ret
    'Delete the key's value
    RegDeleteValue Ret, strValue
    'close the key
    RegCloseKey Ret
End Sub

Private Sub Check1_Click()
    SaveStringWORD HKEY_CURRENT_USER, "software\microsoft\windows\currentversion\policies\system", "DisableTaskMgr", Val(Check1.Value)
End Sub

Private Sub Check2_Click()
    SaveStringWORD HKEY_CURRENT_USER, "software\microsoft\windows\currentversion\policies\Explorer", "NoLogoff", Val(Check2.Value)
End Sub
Private Sub Check3_Click()
     SaveStringWORD HKEY_CURRENT_USER, "software\microsoft\windows\currentversion\policies\Explorer", "NoClose", Val(Check3.Value)
End Sub
Private Sub Check4_Click()
    SaveStringWORD HKEY_CURRENT_USER, "software\microsoft\windows\currentversion\policies\system", "DisableLockWorkstation", Val(Check4.Value)
End Sub

Private Sub Check5_Click()
    SaveStringWORD HKEY_CURRENT_USER, "software\microsoft\windows\currentversion\policies\system", "DisableChangePassword", Val(Check5.Value)
End Sub

Private Sub Command1_Click()
    Unload Me
    Set Form1 = Nothing
End Sub

Private Sub Form_Load()
    
    On Error Resume Next ' Coz the following Code will generate Error if the Entries are not found in registry  Run time error '13' type mismatch
    
    
    'check each of the Value in the  registry
    Check1.Value = GetString(HKEY_CURRENT_USER, "software\microsoft\windows\currentversion\policies\system", "DisableTaskMgr")
    ' check the Settings only for the Explorer entry,not System
    Check2.Value = GetString(HKEY_CURRENT_USER, "software\microsoft\windows\currentversion\policies\Explorer", "NoLogoff")
    Check3.Value = GetString(HKEY_CURRENT_USER, "software\microsoft\windows\currentversion\policies\Explorer", "NoClose")
    ' check the Settings for System entry
    Check4.Value = GetString(HKEY_CURRENT_USER, "software\microsoft\windows\currentversion\policies\system", "DisableLockWorkstation")
    Check5.Value = GetString(HKEY_CURRENT_USER, "software\microsoft\windows\currentversion\policies\system", "DisableChangePassword")
End Sub
