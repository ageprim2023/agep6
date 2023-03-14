VERSION 5.00
Begin VB.Form login 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Mot de passe"
   ClientHeight    =   4065
   ClientLeft      =   4035
   ClientTop       =   2205
   ClientWidth     =   2595
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "login.frx":324A
   ScaleHeight     =   4065
   ScaleWidth      =   2595
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000C0&
      Caption         =   "œŒÊ·"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      MaskColor       =   &H000000FF&
      TabIndex        =   3
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2685
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«·„” Œœ„"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   27.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂ·„… «·”—"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim b As String
Private Sub Combo1_Change()
On Error Resume Next
Call cont
Do While Not ut.EOF
If Combo1.Text = ut!uti Then
Text1.SetFocus
Exit Sub
End If
ut.MoveNext
Loop
End Sub

Private Sub Combo1_Click()
On Error Resume Next
Combo1_Change
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim ex As Integer
Dim an As String
If Combo1.Text = "" Or Text1.Text = "" Then
MsgBox "«œŒ· «”„ «·„” Œœ„ Êﬂ·„… «·”— «·„ÿ«»ﬁ… ·Â", vbCritical
Exit Sub
End If
Call cont
Do While Not ut.EOF
If Combo1.Text = ut!uti And Text1.Text = ut!mot Then
face.tbToolBar.Wrappable = True
face.SBB1.Panels(11).Text = ut!uti
face.SBB1.Panels(1).Text = ut!Dir
face.SBB1.Panels(2).Text = ut!par
face.SBB1.Panels(3).Text = ut!cla
face.SBB1.Panels(4).Text = ut!etu
face.SBB1.Panels(5).Text = ut!pro
face.SBB1.Panels(6).Text = ut!cai
face.SBB1.Panels(7).Text = ut!com
face.SBB1.Panels(8).Text = ut!Arc
face.SBB1.Panels(13).Text = sr!eco
Unload Me
Exit Sub
End If
'***********
ut.MoveNext
Loop
MsgBox "«”„ «·„” Œœ„ Ê ﬂ·„… «·”— €Ì— „ ÿ«»ﬁÌ‰", vbCritical
Text1.Text = ""
Text1.SetFocus
End Sub

Private Sub Form_Load()
'On Error Resume Next
Dim an As String
Me.Top = 300
Me.Left = 6000
'Me.vcFORMSHAPE1.ShapeIt Me
'Unload enre
Call chargec1
'Call cont
'Face.StatusBar1.Panels(4).Text = ec!ann
'Face.StatusBar1.Panels(5).Text = " «·”‰… «·œ—«”Ì…"
'Face.StatusBar1.Panels(7).Text = ec!nom
End Sub
Private Sub chargec1()
'On Error GoTo u
Call cont
Combo1.Clear
  Do While Not ut.EOF
    Combo1.AddItem ut!uti
    ut.MoveNext
  Loop
Exit Sub
u:
MsgBox "›Ì Õ«·… «” „—«— Â–Â «·„‘ﬂ·… Ì—ÃÏ «·« ’«· »«·„»—„Ã 22660920", vbExclamation
  End
  End Sub

Private Sub Label3_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Dim ex As Integer
Dim an As String
Dim am As String
If Text1.Text <> "" Then
If KeyCode = 13 Then
Command1_Click
End If
End If
End Sub

