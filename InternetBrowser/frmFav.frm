VERSION 5.00
Begin VB.Form fav 
   BackColor       =   &H80000011&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton close 
      Caption         =   "Close"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   4320
      Width           =   1400
   End
   Begin VB.CommandButton remove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   4320
      Width           =   1400
   End
   Begin VB.CommandButton save 
      Caption         =   "Save"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   1400
   End
   Begin VB.ListBox favlist 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "fav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub close_Click()
Unload Me
End Sub

Private Sub favlist_DblClick()
Dim go
go = favlist.ListIndex
frmMain.urlbox.Text = favlist.list(go)
frmMain.Web.Navigate favlist.list(go)
Unload Me

End Sub

Private Sub remove_Click()
Dim remove
remove = favlist.ListIndex

favlist.RemoveItem (remove)


End Sub

Private Sub Form_Load()
Ontop Me
On Error Resume Next
    Call ReadList(favlist, "C:\windows\fav.tmp", True)
    
End Sub
Private Sub save_Click()

Call WriteList(favlist, "C:\windows\fav.tmp")
Unload Me

End Sub





