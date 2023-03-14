VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form contact1 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3300
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6840
      OleObjectBlob   =   "contact.frx":0000
      Top             =   2400
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3015
      Index           =   0
      Left            =   120
      ScaleHeight     =   3015
      ScaleWidth      =   8895
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin VB.PictureBox Picture112 
         BorderStyle     =   0  'None
         Height          =   2775
         Index           =   2
         Left            =   6600
         Picture         =   "contact.frx":0234
         ScaleHeight     =   2775
         ScaleWidth      =   2175
         TabIndex        =   8
         Top             =   120
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "«€·«ﬁ"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   1
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·«”„   : «»Ê»ﬂ— «Õ„œÊ «·€“«·Ì "
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
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Width           =   5085
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Œ—ÌÃ Ã«„⁄… €” Ê‰ »Ì—ÃÌ ”Ì‰·ÊÌ «·”‰€«·"
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
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   3525
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   $"contact.frx":37D4
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
         Height          =   1095
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   6165
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Â« › —ﬁ„ : 22660920 - 33440920"
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
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   6165
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·»—Ìœ «·«·ﬂ —Ê‰Ì"
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
         Height          =   255
         Left            =   4440
         TabIndex        =   3
         Top             =   2160
         Width           =   1965
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " agep06@yahoo.fr "
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
         Left            =   3000
         TabIndex        =   2
         Top             =   2160
         Width           =   1965
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         FillColor       =   &H00000040&
         Height          =   2775
         Left            =   120
         Top             =   120
         Width           =   6375
      End
   End
End
Attribute VB_Name = "contact1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Top = 4000
Skin1.LoadSkin App.Path & "\green.skn"
Skin1.ApplySkin Me.hWnd
End Sub
