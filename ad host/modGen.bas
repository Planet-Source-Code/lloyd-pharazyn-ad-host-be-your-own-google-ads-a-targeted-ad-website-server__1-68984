Attribute VB_Name = "modGen"
Public AdsPath As String, AccountsPath As String, WebPath As String, NewAccPath As String

Public Function GetAd(Adlocation As String) As String
'this sub loads the advert and if it is missing it will load
'default ad ie: your own site one or it will display a missing message
Dim strAdvert As String

If FileExists(Adlocation) Then
    Open Adlocation For Binary As #1
    strAdvert = Input(FileLen(Adlocation), #1)
    Close #1
ElseIf FileExists(AdsPath & "default.jpg") Then
    Open AdsPath & "default.jpg" For Binary As #2
    strAdvert = Input(FileLen(AdsPath & "default.jpg"), #2)
    Close #2
Else
    strAdvert = "ad missing " & Adlocation
End If
GetAd = strAdvert
End Function

Public Function FileExists(ByVal Adlocation As String) As Boolean
'just checking to see if any needed files are here
Dim AdSize As Integer
On Error Resume Next
AdSize = Len(Dir$(Adlocation))
If Err Or AdSize = 0 Then
FileExists = False
Else
FileExists = True
End If
End Function

Public Sub DoAdvert(strData As String, Socket As Integer)
Dim AdvertAccount As String, AdCategory As String
If InStr(1, strData, "?") = 0 Then AdCategory = "": GoTo NextTodo
AdvertAccount = Split(strData, "?")(0)
AdCategory = Split(strData, "?")(1)
'######################################################
' This is where the billing info could go for the advertiser
' using the AdvertAccount string to identify the account
' i have added the code for counting how many times the
' ad has been loaded but not added the ad click yet
'######################################################
UpdateBillingAccount AdvertAccount
NextTodo:
'######################################################
' at the moment the this only loads the relating ad file
' i want to make it load an ad at random in the same category
' that will hopefully come out in the next update
'######################################################
Send_Data GetAd(AdsPath & AdCategory & ".jpg"), Socket
End Sub
