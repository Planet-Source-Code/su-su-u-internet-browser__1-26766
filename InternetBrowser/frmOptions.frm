VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H8000000C&
   Caption         =   "Internet Options"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   5400
      TabIndex        =   9
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "O.K."
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Top             =   840
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   480
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   4575
   End
   Begin VB.CommandButton cmdDeleteHistory 
      BackColor       =   &H80000016&
      Caption         =   "Delete History"
      Height          =   495
      Left            =   120
      MaskColor       =   &H00808080&
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000C&
      Caption         =   "Default News Site:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000C&
      Caption         =   "Default Search Site:"
      ForeColor       =   &H80000016&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000C&
      Caption         =   "Home Page:"
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDeleteHistory_Click()
On Error Resume Next
Dim Msg
    Msg = MsgBox("Are you sure you want to delete all history?", vbYesNo Or vbQuestion, "Delete?")
        If Msg = vbYes Then
            Kill App.Path & "/history.dll"
            frmMain.urlbox.Clear
        Else
            Exit Sub
        End If
End Sub

Private Sub CmdOk_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
iniPath$ = App.Path & "/web.dll"
entry$ = Text1.Text
r% = WritePrivateProfileString("Main", "Home", entry$, iniPath$)
entry$ = Text2.Text
r% = WritePrivateProfileString("Main", "SearchUrl", entry$, iniPath$)
entry$ = Text3.Text
r% = WritePrivateProfileString("Main", "NewsUrl", entry$, iniPath$)
MsgBox "Information Saved"
Unload Me
End Sub

Private Sub Form_Load()
iniPath$ = App.Path & "/Web.dll"
Text1.Text = GetFromINI("Main", "Home", iniPath$)
Text2.Text = GetFromINI("Main", "SearchUrl", iniPath$)
Text3.Text = GetFromINI("Main", "NewsUrl", iniPath$)
End Sub

