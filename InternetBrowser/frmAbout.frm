VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   Caption         =   "About MyApp"
   ClientHeight    =   3120
   ClientLeft      =   3120
   ClientTop       =   3075
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2153.479
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000018&
      Caption         =   "Written by... Daniel Ho"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1242.392
      Y2              =   1242.392
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H80000018&
      Caption         =   "Description : Functions as a Internet Web Browser"
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1050
      TabIndex        =   0
      Top             =   1200
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H80000018&
      Caption         =   "Multi-Functional Internet Browser"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   480
      Left            =   1050
      TabIndex        =   1
      Top             =   240
      Width           =   4605
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.935
      Y2              =   1697.935
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H80000018&
      Caption         =   "Version 1.00 Beta"
      Height          =   225
      Left            =   1050
      TabIndex        =   2
      Top             =   780
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Ontop Me
    
End Sub

Private Sub form_click()
Unload Me
End Sub
