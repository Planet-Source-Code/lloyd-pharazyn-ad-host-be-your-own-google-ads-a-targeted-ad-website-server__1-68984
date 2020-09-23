Attribute VB_Name = "modTCP"
Option Explicit

Public Ads As Long, FollowedLinks As Long
Public OpenConnections As Long

Public Sub Get_str(strData As String, Socket As Integer)
'this handles all requested ads and files
Dim strAdvert As String, strGet As String, strGet2 As String
strGet = InStr(strData, "GET ")
strGet2 = InStr(strGet + 5, strData, " ")
strAdvert = Trim(Mid(strData, strGet + 5, strGet2 - (strGet + 4)))

If Right(strAdvert, 1) = "/" Then strAdvert = Left(strAdvert, Len(strAdvert) - 1)

If Left(strAdvert, 3) = "ad=" Then
strAdvert = Replace(strAdvert, "ad=", "")
DoAdvert strAdvert, Socket
Ads = Ads + 1
Exit Sub
End If

If strAdvert = "/" Or strAdvert = "" Then strAdvert = "index.html"
Send_Data GetAd(App.Path & "\web\" & strAdvert), Socket

End Sub

Public Sub Post_str(strData As String, Socket As Integer)
On Error GoTo Hell
Dim AccountID As String, Email As String, Category As String, IP As String
'This is where you can add code for handling form posting
'if you want to have people automaticly signup to your service
AccountID = Split(strData, "SignUpName=")(1)
AccountID = Split(AccountID, "&")(0)
AccountID = Replace(AccountID, "+", " ")
Email = Split(strData, "SignUpEmail=")(1)
Email = Split(Email, "&")(0)
Category = Split(strData, "SiteAbout=")(1)
Category = Split(Category, "&")(0)
Category = Replace(Category, "+", " ")
Category = Replace(Category, "%26", "&")
IP = frmServer.Socket(Socket).RemoteHostIP
PendingAccount AccountID, Email, Category, IP
Send_Data "<center>Your Application has been sent<br>We will reply to your supplied email address (" & Email & ") when your account has been created<br>Thank you</center>", Socket
Exit Sub
Hell:
Send_Data "<center>Sorry there was a server error, your application may not have been sent<br>please try again</center>", Socket
End Sub

Public Sub Send_Data(strData As String, Socket As Integer)
frmServer.Socket(Socket).SendData strData
End Sub
