Attribute VB_Name = "modAccounts"
'###################################################
'I got these first strings from www.pscode.com/vb
'credit goes to the orinal author ....
'sorry i dont remember your name
'###################################################

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function GetFromINI(Section As String, Key As String, Directory As String) As String
    Dim strBuffer As String
    strBuffer = String(750, Chr(0))
    Key$ = LCase$(Key$)
    GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function


Public Sub WriteToINI(Section As String, Key As String, KeyValue As String, Directory As String)
    Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
End Sub
'End of Borrowed code
'###################################################

'###################################################
'This is pretty much self explaining these subs
'are called with relation to the ini files and ads
'###################################################
Public Sub CreateAccount(AccountID As String, ContactEmail As String, AdsShown As String)
If FileExists(AccountsPath & AccountID & ".ini") = False Then
WriteToINI "Account", "ID", AccountID, AccountsPath & AccountID & ".ini"
WriteToINI "Account", "Email", ContactEmail, AccountsPath & AccountID & ".ini"
WriteToINI "Bill", "Adshown", AdsShown, AccountsPath & AccountID & ".ini"
End If
End Sub

Public Sub UpdateBillingAccount(AccountID As String)
Dim tmpAds As String
If FileExists(AccountsPath & AccountID & ".ini") Then
tmpAds = GetFromINI("Bill", "Adshown", AccountsPath & AccountID & ".ini")
tmpAds = Val(tmpAds) + 1
WriteToINI "Bill", "Adshown", tmpAds, AccountsPath & AccountID & ".ini"
End If
End Sub

Public Sub PendingAccount(AccountID As String, ContactEmail As String, Category As String, IP As String)
If FileExists(AccountsPath & AccountID & ".ini") = False Then
WriteToINI "Account", "ID", AccountID, NewAccPath & AccountID & ".ini"
WriteToINI "Account", "Email", ContactEmail, NewAccPath & AccountID & ".ini"
WriteToINI "Site", "Category", Category, NewAccPath & AccountID & ".ini"
WriteToINI "Misc", "Time", Time, NewAccPath & AccountID & ".ini"
WriteToINI "Misc", "Date", Date, NewAccPath & AccountID & ".ini"
WriteToINI "Misc", "IP Address", IP, NewAccPath & AccountID & ".ini"
End If
End Sub
