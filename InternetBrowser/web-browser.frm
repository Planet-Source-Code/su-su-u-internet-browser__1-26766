VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Multi-Functional Internet Browser"
   ClientHeight    =   6810
   ClientLeft      =   1020
   ClientTop       =   1110
   ClientWidth     =   10110
   Icon            =   "web-browser.frx":0000
   LinkTopic       =   "frmMain"
   ScaleHeight     =   6810
   ScaleWidth      =   10110
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   4935
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   8705
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   5
      Top             =   510
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   794
      ButtonWidth     =   767
      ButtonHeight    =   741
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "E-Mail"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Edit Current Web Document"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "View Source"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Launch Browser Media Player"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Launch Calendar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Launch ICQ"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Launch MSN Messenger"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Avtivate System Screensaver"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   900
      ButtonWidth     =   1614
      ButtonHeight    =   847
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
            Key             =   "back"
            Object.ToolTipText     =   "Go Back a Page"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forward"
            Key             =   "forward"
            Object.ToolTipText     =   "Go Forward a Page"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            Key             =   "stop"
            Object.ToolTipText     =   "Stop the Current Task"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "refresh"
            Object.ToolTipText     =   "Refresh Current Page"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Home"
            Key             =   "home"
            Object.ToolTipText     =   "Go to Home Page"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Key             =   "search"
            Object.ToolTipText     =   "Go to the Default Search Engine"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Newspaper"
            Key             =   "newspaper"
            Object.ToolTipText     =   "Online Newspaper"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Favourites"
            Key             =   "favourites"
            Object.ToolTipText     =   "View Favourites"
            ImageIndex      =   8
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "add"
                  Text            =   "Add"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "view"
                  Text            =   "View"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "History"
            Key             =   "history"
            Object.ToolTipText     =   "View Visited Web Pages"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Key             =   "help"
            Object.ToolTipText     =   "View Help File"
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2400
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   22
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "web-browser.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "web-browser.frx":191E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "web-browser.frx":2972
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "web-browser.frx":39C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "web-browser.frx":421A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "web-browser.frx":5E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "web-browser.frx":7B02
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "web-browser.frx":9776
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1680
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "web-browser.frx":9FCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "web-browser.frx":A28E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "web-browser.frx":A552
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "web-browser.frx":B5A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "web-browser.frx":B9D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "web-browser.frx":CA2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "web-browser.frx":CF2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "web-browser.frx":FF8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "web-browser.frx":10FDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "web-browser.frx":114A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   960
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.ComboBox urlbox 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Text            =   "http://"
      Top             =   1200
      Width           =   10868
   End
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   480
      Top             =   5520
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6555
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9478
            MinWidth        =   9372
            Key             =   "status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5218
            MinWidth        =   5112
            Picture         =   "web-browser.frx":124FA
            Text            =   "Internet"
            TextSave        =   "Internet"
            Key             =   "strconnection"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2556
            MinWidth        =   2556
            TextSave        =   "6:51 PM"
            Key             =   "strtime"
         EndProperty
      EndProperty
      MousePointer    =   3
   End
   Begin MSComDlg.CommonDialog commondialog1 
      Left            =   0
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   6180
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   9757
      ExtentX         =   17210
      ExtentY         =   10901
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Address Bar:"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu itmOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu itmSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu itmSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu itmPageSetup 
         Caption         =   "Page Setup"
      End
      Begin VB.Menu itmPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu itmSeperator2 
         Caption         =   "-"
      End
      Begin VB.Menu itmWorkOffline 
         Caption         =   "Work Offline"
      End
      Begin VB.Menu itmExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu itmEditByMSWord 
         Caption         =   "Edit Web Doc Using MS Word"
         Enabled         =   0   'False
      End
      Begin VB.Menu itmCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu itmCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu itmPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu itmToobars 
         Caption         =   "Toobars"
         Begin VB.Menu itmStandard 
            Caption         =   "Standard"
            Checked         =   -1  'True
         End
         Begin VB.Menu itmAddressBar 
            Caption         =   "Address Bar"
            Checked         =   -1  'True
         End
         Begin VB.Menu itmStatusBar 
            Caption         =   "Status Bar"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu itmCalendar 
         Caption         =   "Calendar"
      End
      Begin VB.Menu itmSeperator3 
         Caption         =   "-"
      End
      Begin VB.Menu itmInternetOptions 
         Caption         =   "Internet Options"
      End
      Begin VB.Menu itmSeperator4 
         Caption         =   "-"
      End
      Begin VB.Menu itmSource 
         Caption         =   "Source"
      End
   End
   Begin VB.Menu mnuSpecial 
      Caption         =   "&Special"
      Begin VB.Menu itmEmail 
         Caption         =   "E-Mail"
      End
      Begin VB.Menu itmSeperator5 
         Caption         =   "-"
      End
      Begin VB.Menu itmMediaPlayer 
         Caption         =   "Browser Media Player"
      End
      Begin VB.Menu itmSeperator6 
         Caption         =   "-"
      End
      Begin VB.Menu itmICQ 
         Caption         =   "ICQ"
      End
      Begin VB.Menu itmMSN 
         Caption         =   "MSN Messenger"
      End
      Begin VB.Menu itmScrnSave 
         Caption         =   "Activate Screensaver"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu itmViewHelpFile 
         Caption         =   "View Help File"
      End
      Begin VB.Menu itmAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------------------
'Programmer:     Daniel Ho
'Date Written:   05/08/2001
'Purpose:        Used for demostrating how a basic
'                internet browser works and a lot
'                of advance features that IE doesn't
'                have.
'-----------------------------------------------------

Private Sub form_resize()

Web.left = 360
Web.Width = frmMain.Width - 500


End Sub

Private Sub Form_Load()
Dim starting, history
On Error Resume Next

frmAbout.Show
 iniPath$ = App.Path & "/web.dll"
 starting = GetFromINI("Main", "Home", iniPath$)
 Web.Navigate (starting)
 StatusBar1.Panels(1).Text = "Ready"
 StatusBar1.Panels(2).Text = StatusBar1.Panels(2).Text & "  -   Online"
 On Error GoTo Err
    Open App.Path & "/history.dll" For Input As #1
        Do While Not EOF(1)
            Line Input #1, history
                urlbox.AddItem history
            Loop
    Close #1
Err:
    Exit Sub
    Close #1



End Sub

Private Sub itmAbout_Click()
Dim About As New frmAbout
About.Visible = True

End Sub

Private Sub itmAddressBar_Click()
If itmAddressBar.Checked = True Then
    itmAddressBar.Checked = False
    urlbox.Visible = False
Else
    itmAddressBar.Checked = True
    urlbox.Visible = True
End If

End Sub

Private Sub itmCalendar_Click()
frmdate.Show
frmdate.MonthView.Value = Date
End Sub

Private Sub itmCopy_Click()
Dim eQuery As OLECMDF       'return value type for QueryStatusWB
On Error Resume Next
eQuery = Web.QueryStatusWB(OLECMDID_COPY)
If Err.Number = 0 Then
If eQuery And OLECMDF_ENABLED Then
Web.ExecWB OLECMDID_COPY, OLECMDEXECOPT_PROMPTUSER, "", ""      'Ok to Print?
End If
End If


End Sub

Private Sub itmEditByMSWord_Click() 'Not finished Yet


Open App.Path & "/edit.html" For Output As #1
    Print #1, Inet1.OpenURL(Web.LocationURL)
Close #1
 Shell "C:\Program Files\Microsoft Office\Office\Winword.exe"



End Sub

Private Sub itmEmail_Click()
 Dim subject, person
        person = InputBox("Enter email address", "email")
        subject = InputBox("Enter subject for email", "subject")
        Web.Navigate ("mailto:" & person & "?subject=" & subject)
End Sub

Private Sub itmExit_Click()
Unload Me
End Sub



Private Sub itmICQ_Click()
On Error Resume Next
  Shell "C:\program files\icq\icq.exe", vbNormalFocus

End Sub

Private Sub itmInternetOptions_Click()
frmOptions.Show

End Sub

Private Sub itmMediaPlayer_Click()
frmMainMediaPlayer.Show
End Sub

Private Sub itmMSN_Click()
On Error Resume Next
  Shell "C:\Program Files\Messenger\msmsgs.exe", vbNormalFocus

End Sub

Private Sub itmOpen_Click()

On Error Resume Next

commondialog1.Filter = "All Internet Files (*.htm,*.html,*.asp,*.shtml,*.js,*.dhtml) | *.htm;*.html;*.asp;*.shtml;*.js;*.dhtml | Word Documents | *.doc"
commondialog1.ShowOpen
If commondialog1.filename = "" Then
    Exit Sub
Else
    Web.Navigate (commondialog1.filename)
End If
urlbox.Text = commondialog1.filename

End Sub



Private Sub itmPageSetup_Click()
Dim eQuery As OLECMDF
On Error Resume Next
eQuery = Web.QueryStatusWB(OLECMDID_PAGESETUP)
If Err.Number = 0 Then
If eQuery And OLECMDF_ENABLED Then
Web.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_PROMPTUSER, "", ""    'Ok to Print?
Else
MsgBox "No contents of Webbrowser"
End If
End If


End Sub

Private Sub itmPrint_Click()
On Error Resume Next
Dim eQuery As OLECMDF
eQuery = Web.QueryStatusWB(OLECMDID_PRINT)
If Err.Number = 0 Then
If eQuery And OLECMDF_ENABLED Then
Web.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER, "", ""    'Ok to Print?
Else
MsgBox "Printer Not Found or No Contents in Webbrowser"
End If
End If


End Sub

Private Sub itmSaveAs_Click()
Dim buffer
On Error Resume Next

commondialog1.Filter = "Internet Files (.html) | *.html |Internet Files (.htm) | *.htm | All Files | *.*"
commondialog1.ShowSave
Open commondialog1.filename For Output As #1
buffer = Inet1.OpenURL(Web.LocationURL)
    Print #1, buffer
Close #1



End Sub


Private Sub itmScrnSave_Click()
  Call SendMessage(Me.hWnd, WM_SYSCOMMAND, SC_SCREENSAVE, _
    0&)
End Sub


Private Sub itmSource_Click()
On Error Resume Next
Open App.Path & "/source.tmp" For Output As #1
    Print #1, Inet1.OpenURL(Web.LocationURL)
Close #1
    Shell "C:\windows\notepad.exe " & App.Path & "/source.tmp", vbNormalFocus
    Kill App.Path & "/source.tmp"
End Sub

Private Sub itmStandard_Click()
If itmStandard.Checked = True Then
    itmStandard.Checked = False
    Toolbar1.Visible = False
Else
    itmStandard.Checked = True
    Toolbar1.Visible = True
End If
    

End Sub

Private Sub itmStatusBar_Click()
If itmStatusBar.Checked = True Then
    itmStatusBar.Checked = False
    StatusBar1.Visible = False
Else
    itmStatusBar.Checked = True
    StatusBar1.Visible = True
End If

End Sub

Private Sub itmViewHelpFile_Click()
Dim help
Web.Navigate (App.Path & "/help.doc")
End Sub

Private Sub itmWorkOffline_Click()
If itmWorkOffline.Checked = True Then
    itmWorkOffline.Checked = False
    Web.Offline = True
    StatusBar1.Panels(2).Text = "Internet  -   Online"
Else
    itmWorkOffline.Checked = True
    Web.Offline = True
    StatusBar1.Panels(2).Text = "Internet  -   Offline"
End If

End Sub





Private Sub Timer1_Timer()
Unload frmAbout
Me.WindowState = 2
frmMain.SetFocus
Timer1.Enabled = False
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error Resume Next

Select Case Button.Key

Case "back"
urlbox.AddItem (urlbox.Text)
Web.GoBack
Case "forward"
Web.GoForward
Case "stop"
Web.Stop
Case "refresh"
Web.Refresh
Case "home"
Dim starting
iniPath$ = App.Path & "/web.dll"
        starting = GetFromINI("Main", "Home", iniPath$)
        Web.Navigate (starting)
Case "search"
 iniPath$ = App.Path & "/web.dll"
Dim search
    search = GetFromINI("Main", "SearchUrl", iniPath$)
    Web.Navigate (search)
    urlbox.Text = search
    urlbox.AddItem (search)
Case "newspaper"
iniPath$ = App.Path & "/web.dll"
Dim newspaper
    newspaper = GetFromINI("Main", "NewsUrl", iniPath$)
    Web.Navigate (newspaper)
    urlbox.Text = newspaper
    urlbox.AddItem (newspaper)
Case "favourites"
fav.Show
Case "history"
loadhistory
Case "help"
Web.Navigate (App.Path & "/help.doc")
End Select

End Sub
Public Sub loadhistory()
On Error GoTo Err
Dim history As Integer
        Open App.Path & "/history.htm" For Output As #2
            Print #2, "<html>" & vbCrLf & "<title>History</title>" & vbCrLf & "<font size=15 face=arial color=black>History<br></br><br></br><font size=2 face=arial color=black>" & vbCrLf
                 For history = 0 To urlbox.ListCount
            Print #2, "<a href=" + urlbox.list(history) + ">" + urlbox.list(history) + "<br>"
                Next history
            Print #2, "</a><font size=1 face=arial color=black><br></br>End of history"
        Close #2
    Web.Navigate (App.Path & "/history.htm")
Err:
    Exit Sub
Web.Navigate (App.Path & "/history.htm")
End Sub


Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Key

Case "add"
Dim addfav
Dim favlist As ListBox
addfav = InputBox("Enter the Web site which you wish to add to the Favourites", "Add", urlbox.Text)
If addfav = "" Then
    Exit Sub
        Else:
        fav.favlist.AddItem (addfav)
End If
fav.Show

        
Case "view"
fav.Show

End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Select Case Button.Index
Case 1
Shell "C:\Program Files\Outlook Express\MSIMN.EXE", vbNormalFocus
Case 2

Case 3
On Error Resume Next
Open App.Path & "/source.tmp" For Output As #1
    Print #1, Inet1.OpenURL(Web.LocationURL)
Close #1
    Shell "C:\windows\notepad.exe " & App.Path & "/source.tmp", vbNormalFocus
    Kill App.Path & "/source.tmp"
Case 5
frmMainMediaPlayer.Show
Case 6
frmdate.Show
Case 7
Shell "C:\program files\icq\icq.exe", vbNormalFocus
Case 8
Shell "C:\Program Files\Messenger\msmsgs.exe", vbNormalFocus
Case 10
 Call SendMessage(Me.hWnd, WM_SYSCOMMAND, SC_SCREENSAVE, _
    0&)
End Select
Err:
    Exit Sub

End Sub

Private Sub Urlbox_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Web.Navigate (urlbox.Text)
    urlbox.AddItem (urlbox.Text)
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim history As Integer
    Open App.Path & "/history.dll" For Output As #1
        For history = 1 To urlbox.ListCount - 1
    Print #1, urlbox.list(history)
        Next history
    Close #1
End
End Sub



Private Sub Web_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)

On Error Resume Next
ProgressBar1.Max = ProgressMax
ProgressBar1.Value = Progress

End Sub

Private Sub web_downloadbegin()
StatusBar1.Panels(1).Text = "Opening Page....."
End Sub

Private Sub web_DownloadComplete()
StatusBar1.Panels(1).Text = "Download Completed..."
End Sub

Private Sub Web_NavigateComplete2(ByVal pDisp As Object, Url As Variant)
    frmMain.Caption = Web.LocationName
    StatusBar1.Panels(1).Text = Web.LocationURL
    urlbox.Text = Web.LocationURL
End Sub

Private Sub Web_DocumentComplete(ByVal pDisp As Object, Url As Variant)
    StatusBar1.Panels(1).Text = "Document Finished."
End Sub
