VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   Caption         =   "Ad Host Server"
   ClientHeight    =   1545
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   ScaleHeight     =   1545
   ScaleWidth      =   4290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "New Account"
      Height          =   330
      Left            =   2970
      TabIndex        =   3
      Top             =   1140
      Width           =   1230
   End
   Begin VB.Timer TmrTimeOut 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   5000
      Left            =   1515
      Top             =   1050
   End
   Begin VB.Timer TmyUpdateInfo 
      Interval        =   1000
      Left            =   1065
      Top             =   1050
   End
   Begin VB.CommandButton BtnReset 
      Caption         =   "Reset Counter"
      Height          =   330
      Left            =   2970
      TabIndex        =   2
      Top             =   525
      Width           =   1230
   End
   Begin VB.CommandButton BtnSockets 
      Caption         =   "Start Service"
      Height          =   330
      Left            =   2970
      TabIndex        =   1
      Top             =   90
      Width           =   1230
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   585
      Top             =   1050
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Request 
      Left            =   150
      Top             =   1050
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblStatus 
      Caption         =   "Label1"
      Height          =   1305
      Left            =   135
      TabIndex        =   0
      Top             =   150
      Width           =   2685
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtnReset_Click()
Ads = 0: FollowedLinks = 0
End Sub

Private Sub BtnSockets_Click()
If BtnSockets.Caption = "Start Service" Then
Request.Close
Request.LocalPort = 80
Request.Listen
BtnSockets.Caption = "Stop Service"
ElseIf BtnSockets.Caption = "Stop Service" Then
BtnSockets.Caption = "Start Service"
Request.Close
End If
End Sub



Private Sub Command1_Click()
frmAddAccount.Show
End Sub

Private Sub Form_Load()
Dim i As Integer
For i = 1 To 200
Load Socket(i)
Load TmrTimeOut(i)
Next i
BtnReset_Click
TmyUpdateInfo_Timer
AccountsPath = App.Path & "/accounts/"
AdsPath = App.Path & "/ads/"
WebPath = App.Path & "/web/"
NewAccPath = App.Path & "/pending accounts/"
End Sub

Private Sub Request_ConnectionRequest(ByVal requestID As Long)
Dim i As Integer
For i = 0 To 200
If Socket(i).State = sckClosed Then
Socket(i).Close
Socket(i).Accept (requestID)
TmrTimeOut(i).Enabled = True
OpenConnections = OpenConnections + 1
Exit Sub
End If
Next i
End Sub

Private Sub Socket_Close(Index As Integer)
OpenConnections = OpenConnections - 1
Socket(Index).Close
TmrTimeOut(Index).Enabled = False
End Sub

Private Sub Socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData As String
Dim Page As String
Socket(Index).GetData strData
If Left(strData, 3) = "GET" Then
    Get_str strData, Index
    Exit Sub
ElseIf Left(strData, 4) = "POST" Then
    Post_str strData, Index
    Exit Sub
End If

End Sub

Private Sub Socket_SendComplete(Index As Integer)
OpenConnections = OpenConnections - 1
Socket(Index).Close
TmrTimeOut(Index).Enabled = False
End Sub

Private Sub TmrTimeOut_Timer(Index As Integer)
If Socket(Index).State = sckClosed Then GoTo Hell
Socket(Index).Close
TmrTimeOut(Index).Enabled = False
OpenConnections = OpenConnections - 1
Hell:
End Sub

Private Sub TmyUpdateInfo_Timer()
lblStatus.Caption = Request.State & vbCrLf & "Adverts Requested: " & Ads & vbCrLf & "Ads Clicked: " & FollowedLinks & vbCrLf & "Open Connections: " & OpenConnections
End Sub
