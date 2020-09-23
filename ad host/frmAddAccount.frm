VERSION 5.00
Begin VB.Form frmAddAccount 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Create Account"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Save Account"
      Height          =   345
      Left            =   1215
      TabIndex        =   4
      Top             =   1485
      Width           =   1290
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   165
      TabIndex        =   3
      Top             =   1035
      Width           =   3360
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   165
      TabIndex        =   1
      Top             =   405
      Width           =   3360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Contact Email"
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   765
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Account ID"
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   840
   End
End
Attribute VB_Name = "frmAddAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CreateAccount Text1.Text, Text2.Text, 0
Unload Me
End Sub
