VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.MDIForm face 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "»—«„Ã  ”ÌÌ— «·„ƒ””«  «·Œ«’… «·‰”Œ… 6 «·Œ«’… »«·„œ«—” «·Õ—…"
   ClientHeight    =   10095
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   Icon            =   "Interface.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "Interface.frx":324A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer11 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   13800
      Top             =   5640
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   13800
      Top             =   5160
   End
   Begin VB.Timer Timer10 
      Interval        =   500
      Left            =   13320
      Top             =   1320
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   13800
      Top             =   4680
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   13800
      Top             =   4200
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   13800
      Top             =   3720
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   13800
      Top             =   3240
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   13800
      Top             =   2760
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   13800
      Top             =   2280
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   13800
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   13800
      Top             =   1320
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   13800
      OleObjectBlob   =   "Interface.frx":292C9
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   15240
      TabIndex        =   2
      Top             =   0
      Width           =   15240
      Begin MSComCtl2.DTPicker DT1 
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   0
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         Format          =   124649473
         CurrentDate     =   40612
      End
   End
   Begin MSComctlLib.StatusBar SBB1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   9810
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   13800
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Interface.frx":294FD
            Key             =   "iCategory"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Interface.frx":29897
            Key             =   "iSubCategory"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Interface.frx":29C31
            Key             =   "iSupplier"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Interface.frx":2A1CB
            Key             =   "iCustomer"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Interface.frx":2A765
            Key             =   "iUnit"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Interface.frx":2AAFF
            Key             =   "iCountry"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Interface.frx":2AE99
            Key             =   "iMony"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Interface.frx":2B233
            Key             =   "iPurchase"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Interface.frx":2B5CD
            Key             =   "iSale"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Interface.frx":2B967
            Key             =   "iPhone"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Interface.frx":2BD01
            Key             =   "iPhoneModal"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Interface.frx":2C09B
            Key             =   "iReOrder"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Interface.frx":2C435
            Key             =   "prog"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   4  'Align Right
      Height          =   9795
      Left            =   14460
      TabIndex        =   0
      Top             =   15
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   17277
      ButtonWidth     =   1270
      ButtonHeight    =   1005
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   27
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "› Õ"
            Key             =   "entrer"
            Object.ToolTipText     =   "«÷€ÿ Â‰« ·› Õ «·»—‰«„Ã"
            ImageKey        =   "iSubCategory"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "«·≈œ«—…"
            Key             =   "dirc"
            Object.ToolTipText     =   "»Ì«‰«  «·„ƒ””… Ê«·‘—ﬂ«¡ Ê«·„” Œœ„Ì‰"
            ImageKey        =   "iPurchase"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "«·‘—ﬂ«¡"
            Key             =   "part"
            Object.ToolTipText     =   "Ê÷⁄Ì… Õ”«»«  «·‘—ﬂ«¡"
            ImageKey        =   "iSale"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "«·√ﬁ”«„"
            Key             =   "classe"
            Object.ToolTipText     =   "«·√ﬁ”«„ Ê«·„Ê«œ Ê«·Ãœ«Ê· «·“„‰Ì…"
            ImageKey        =   "iCountry"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "«· ·«„Ì–"
            Key             =   "etud"
            Object.ToolTipText     =   "»Ì«‰«  Ê‰ «∆Ã Ê€Ì«»«  «· ·«„Ì– Ê«·»ÕÀ «·”—Ì⁄"
            ImageKey        =   "iCustomer"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "«·√”« –… "
            Key             =   "prof"
            Object.ToolTipText     =   "»Ì«‰«  Ê”Ã· Õ÷Ê— «·√”« –…"
            ImageKey        =   "iSupplier"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "«·’‰œÊﬁ"
            Key             =   "caiss"
            Object.ToolTipText     =   "Õ”«»«  «· ·«„Ì– Ê«·√”« –… Ê«·‘—ﬂ«¡ Ê«·„’—Ê›« "
            ImageKey        =   "iMony"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "«·„Õ«”»…"
            Key             =   "comp"
            Object.ToolTipText     =   "«·œ› — «·ÌÊ„Ì ··„Õ«”»… Ê «·„—ﬂ“ «·„«·Ì ··„ƒ””… "
            ImageKey        =   "iReOrder"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "»ÕÀ"
            Key             =   "rech"
            Object.ToolTipText     =   "⁄„·Ì«  ‘—«¡"
            ImageKey        =   "iPhoneModal"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "«·«—‘Ì›"
            Key             =   "arch"
            Object.ToolTipText     =   "«—‘Ì› «·”‰Ê«  «·œ—«”Ì…"
            ImageKey        =   "iUnit"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "»—‰«„Ã"
            Key             =   "prog"
            ImageKey        =   "iSale"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "«·»—‰«„Ã"
            Key             =   "con"
            Object.ToolTipText     =   "‰»–… ⁄‰ «·»—‰«„Ã"
            ImageKey        =   "prog"
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "€·ﬁ"
            Key             =   "ferm"
            Object.ToolTipText     =   "«÷€ÿ Â‰« ·€·ﬁ ’›Õ«  «·»—‰«„Ã"
            ImageKey        =   "iCategory"
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "≈‰Â«¡"
            Key             =   "sort"
            Object.ToolTipText     =   "«÷€ÿ Â‰« ·€·ﬁ «·»—‰«„Ã"
            ImageKey        =   "iPhoneModal"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "face"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim da As String
Dim dy As String
Dim mont As String
Dim ye As String
Dim myDate As String
Dim x As Integer
Dim yy As Integer

Private Sub MDIForm_Load()
On Error Resume Next
Call chargepanels
DT1.Value = Date
Call dater
Call cont
SBB1.Panels(13).Text = sr!eco

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
g = MsgBox(" Â·  —Ìœ «·Œ—ÊÃ „‰ «·»—‰«„Ã ‰Â«∆Ì« ", vbInformation + vbYesNo, "AGEP")
If g = vbYes Then
End
Exit Sub
Else
Cancel = 1
Exit Sub
End If

End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error Resume Next
Call unloadforms
Select Case Button.Key
Case "entrer"
If tbToolBar.Wrappable = False Then
'Call cont
'ec!ann = MDIForm1.SB1.Panels(8).Text
'ec.Update
login.Show
End If
Exit Sub
Case "classe"
If SBB1.Panels(3).Text = "1" Then
classes.Left = 15000
classes.Show
Timer1.Enabled = True
Exit Sub
End If
Case "etud"
If SBB1.Panels(4).Text = "1" Then
etudiants.Left = 15000
etudiants.Show
Timer2.Enabled = True
Exit Sub
End If
Case "prof"
If SBB1.Panels(5).Text = "1" Then
professeurs.Left = 15000
professeurs.Show
Timer3.Enabled = True
Exit Sub
End If
Case "caiss"
If SBB1.Panels(6).Text = "1" Then
caisse.Left = 15000
caisse.Show
Timer4.Enabled = True
Exit Sub
End If
Case "comp"
If SBB1.Panels(7).Text = "1" Then
comptabilite.Left = 15000
comptabilite.Show
Timer5.Enabled = True
Exit Sub
End If
Case "dirc"
If SBB1.Panels(1).Text = "1" Then
direction.Left = 15000
direction.Show
Timer6.Enabled = True
Exit Sub
End If
Case "part"
If SBB1.Panels(2).Text = "1" Then
partenaires.Left = 15000
partenaires.Show
Timer7.Enabled = True
Exit Sub
End If
Case "arch"
If SBB1.Panels(8).Text = "1" Then
archives.Left = 15000
archives.Show
Timer8.Enabled = True
Exit Sub
End If
Case "con"
contact.Left = 15000
contact.Show
Timer11.Enabled = True
Exit Sub
Case "ferm"
'Unload Me
Exit Sub
Case "sort"
g = MsgBox(" Â·  —Ìœ €·ﬁ ’›Õ«  «·»—‰«„Ã ", vbInformation + vbYesNo, "AGEP")
If g = vbYes Then
face.SBB1.Panels(1).Text = ""
face.SBB1.Panels(2).Text = ""
face.SBB1.Panels(3).Text = ""
face.SBB1.Panels(4).Text = ""
face.SBB1.Panels(5).Text = ""
face.SBB1.Panels(6).Text = ""
face.SBB1.Panels(7).Text = ""
face.SBB1.Panels(8).Text = ""
face.SBB1.Panels(11).Text = ""
face.tbToolBar.Wrappable = False
Exit Sub
Else
Exit Sub
End If

End Select
message.Visible = True
message.Label1.Visible = False
Timer9.Interval = 1
x = 0
message.Left = 15000
message.Show
Timer9.Enabled = True
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
classes.Visible = False
classes.Left = classes.Left - 2000
If classes.Left <= 300 Then
classes.Left = 100
Timer1.Enabled = False
classes.Visible = True
Call cont
If sr!dat = "Rien" Then
MsgBox "·· „ﬂ‰ „‰ œŒÊ· Â–Â «·’›Õ… Ì—ÃÏ  ÕœÌœ  «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì… √Ê·«", vbCritical
Unload classes
Exit Sub
End If
End If
classes.Visible = True
End Sub

Private Sub Timer10_Timer()
On Error Resume Next
DT1.Value = Date
Call dater

End Sub

Private Sub Timer11_Timer()
On Error Resume Next
contact.Visible = False
contact.Left = contact.Left - 2000
If contact.Left <= 300 Then
contact.Left = 100
Timer11.Enabled = False
contact.Visible = True
End If
contact.Visible = True

End Sub

Private Sub Timer2_Timer()
On Error Resume Next
etudiants.Visible = False
etudiants.Left = etudiants.Left - 2000
If etudiants.Left <= 300 Then
etudiants.Left = 100
Timer2.Enabled = False
etudiants.Visible = True
Call cont
If sr!dat = "Rien" Then
MsgBox "·· „ﬂ‰ „‰ œŒÊ· Â–Â «·’›Õ… Ì—ÃÏ  ÕœÌœ  «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì… √Ê·«", vbCritical
Unload etudiants
Exit Sub
End If
End If
etudiants.Visible = True
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
professeurs.Visible = False
professeurs.Left = professeurs.Left - 2000
If professeurs.Left <= 300 Then
professeurs.Left = 100
Timer3.Enabled = False
professeurs.Visible = True
Call cont
If sr!dat = "Rien" Then
MsgBox "·· „ﬂ‰ „‰ œŒÊ· Â–Â «·’›Õ… Ì—ÃÏ  ÕœÌœ  «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì… √Ê·«", vbCritical
Unload professeurs
Exit Sub
End If
End If
professeurs.Visible = True

End Sub
Private Sub unloadforms()
On Error Resume Next
SBB1.Panels(16).Text = ""
Unload login
Unload archives
Unload caisse
Unload classes
Unload comptabilite
Unload contact
Unload direction
Unload etudiants
Unload partenaires
Unload professeurs
End Sub

Private Sub Timer4_Timer()
On Error Resume Next
caisse.Visible = False
caisse.Left = caisse.Left - 2000
If caisse.Left <= 300 Then
caisse.Left = 100
Timer4.Enabled = False
caisse.Visible = True
Call cont
If sr!dat = "Rien" Then
MsgBox "·· „ﬂ‰ „‰ œŒÊ· Â–Â «·’›Õ… Ì—ÃÏ  ÕœÌœ  «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì… √Ê·«", vbCritical
Unload caisse
Exit Sub
End If

End If
caisse.Visible = True

End Sub

Private Sub Timer5_Timer()
On Error Resume Next
comptabilite.Visible = False
comptabilite.Left = comptabilite.Left - 2000
If comptabilite.Left <= 300 Then
comptabilite.Left = 100
Timer5.Enabled = False
comptabilite.Visible = True
Call cont
If sr!dat = "Rien" Then
MsgBox "·· „ﬂ‰ „‰ œŒÊ· Â–Â «·’›Õ… Ì—ÃÏ  ÕœÌœ  «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì… √Ê·«", vbCritical
Unload comptabilite
Exit Sub
End If
End If
comptabilite.Visible = True

End Sub

Private Sub Timer6_Timer()
On Error Resume Next
direction.Visible = False
direction.Left = direction.Left - 2000
If direction.Left <= 300 Then
direction.Left = 100
Timer6.Enabled = False
direction.Visible = True
End If
direction.Visible = True

End Sub

Private Sub Timer7_Timer()
On Error Resume Next
partenaires.Visible = False
partenaires.Left = partenaires.Left - 2000
If partenaires.Left <= 300 Then
partenaires.Left = 100
Timer7.Enabled = False
partenaires.Visible = True
Call cont
If sr!dat = "Rien" Then
MsgBox "·· „ﬂ‰ „‰ œŒÊ· Â–Â «·’›Õ… Ì—ÃÏ  ÕœÌœ  «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì… √Ê·«", vbCritical
Unload partenaires
Exit Sub
End If
End If
partenaires.Visible = True

End Sub

Private Sub Timer8_Timer()
On Error Resume Next
archives.Visible = False
archives.Left = archives.Left - 2000
If archives.Left <= 300 Then
archives.Left = 100
Timer8.Enabled = False
archives.Visible = True
Call cont
If sr!dat = "Rien" Then
MsgBox "·· „ﬂ‰ „‰ œŒÊ· Â–Â «·’›Õ… Ì—ÃÏ  ÕœÌœ  «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì… √Ê·«", vbCritical
Unload archives
Exit Sub
End If
End If
archives.Visible = True

End Sub
Private Sub chargepanels()
On Error Resume Next
SBB1.Panels(1).Width = 10
SBB1.Panels(1).Text = "1"
SBB1.Panels.Add 2
SBB1.Panels(2).Width = 10
SBB1.Panels(2).Text = "1"
SBB1.Panels.Add 3
SBB1.Panels(3).Width = 10
SBB1.Panels(3).Text = "1"
SBB1.Panels.Add 4
SBB1.Panels(4).Width = 10
SBB1.Panels(4).Text = "1"
SBB1.Panels.Add 5
SBB1.Panels(5).Width = 10
SBB1.Panels(5).Text = "1"
SBB1.Panels.Add 6
SBB1.Panels(6).Width = 10
SBB1.Panels(6).Text = "1"
SBB1.Panels.Add 7
SBB1.Panels(7).Width = 10
SBB1.Panels(7).Text = "1"
SBB1.Panels.Add 8
SBB1.Panels(8).Width = 10
SBB1.Panels(8).Text = "1"
SBB1.Panels.Add 9
SBB1.Panels(9).Width = 1300
SBB1.Panels(9).Text = ""
SBB1.Panels(9).Alignment = sbrRight
SBB1.Panels.Add 10
SBB1.Panels(10).Width = 1200
SBB1.Panels(10).Text = "«·”‰… «·œ—«”Ì…"
SBB1.Panels(10).Alignment = sbrRight
SBB1.Panels.Add 11
SBB1.Panels(11).Width = 1500
SBB1.Panels(11).Text = ""
SBB1.Panels(11).Alignment = sbrRight
SBB1.Panels.Add 12
SBB1.Panels(12).Width = 800
SBB1.Panels(12).Text = "«·„” Œœ„"
SBB1.Panels(12).Alignment = sbrRight
SBB1.Panels.Add 13
SBB1.Panels(13).Width = 4500
SBB1.Panels(13).Text = ""
SBB1.Panels(13).Alignment = sbrRight
SBB1.Panels.Add 14
SBB1.Panels(14).Width = 800
SBB1.Panels(14).Text = "«·„ƒ””…"
SBB1.Panels(14).Alignment = sbrRight
SBB1.Panels.Add 15
SBB1.Panels(15).Width = 3200
SBB1.Panels(15).Text = ""
SBB1.Panels(15).Alignment = sbrRight
SBB1.Panels.Add 16
SBB1.Panels(16).Width = 10
SBB1.Panels(16).Text = ""
SBB1.Panels(16).Alignment = sbrRight

End Sub
Private Sub dater()
On Error Resume Next
da = DT1.DayOfWeek
dy = DT1.Day
mont = DT1.Month
ye = DT1.Year
'********** Days
If da = 1 Then
da = "«·«Õœ"
ElseIf da = 2 Then
da = "«·«À‰Ì‰"
ElseIf da = 3 Then
da = "«·À·«À«¡"
ElseIf da = 4 Then
da = "«·«—»⁄«¡"
ElseIf da = 5 Then
da = "«·Œ„Ì”"
ElseIf da = 6 Then
da = "«·Ã„⁄…"
ElseIf da = 7 Then
da = "«·”» "
End If
'********** Months
If mont = 1 Then
mont = "Ì‰«Ì—"
ElseIf mont = 2 Then
mont = "›»—«Ì—"
ElseIf mont = 3 Then
mont = "„«—”"
ElseIf mont = 4 Then
mont = "«»—Ì·"
ElseIf mont = 5 Then
mont = "„«ÌÊ"
ElseIf mont = 6 Then
mont = "ÌÊ‰ÌÊ"
ElseIf mont = 7 Then
mont = "ÌÊ·ÌÊ"
ElseIf mont = 8 Then
mont = "«€”ÿ”"
ElseIf mont = 9 Then
mont = "”» „»—"
ElseIf mont = 10 Then
mont = "«ﬂ Ê»—"
ElseIf mont = 11 Then
mont = "‰Ê›„»—"
ElseIf mont = 12 Then
mont = "œÌ”„»—"
End If
'********** Date
myDate = da + " " + dy + " " + mont + " " + ye
SBB1.Panels(15).Text = myDate + " : " + Time$
'SBB1.Panels(16).Text = Time$
End Sub

Private Sub Timer9_Timer()
On Error Resume Next
If message.Left > 900 And x = 0 Then
message.Left = message.Left - 200
'x = 0
End If
If message.Left < 900 And x < 8 Then
Timer9.Interval = 250
x = x + 1
If x Mod 2 = 0 Then
message.Label1.Visible = False
Else
message.Label1.Visible = True
End If
End If
If x >= 8 Then
Timer9.Interval = 500
message.Label1.Visible = True
x = x + 1
End If
If x >= 15 Then
Unload message
Timer9.Enabled = False
End If

End Sub
