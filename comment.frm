VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form contact 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   9510
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   14700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9510
   ScaleWidth      =   14700
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   9255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   16325
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "„⁄·Ê„«  ⁄‰ «·»—‰«„Ã"
      TabPicture(0)   =   "comment.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ﬂÌ›Ì… «” ⁄„«· «·»—‰«„Ã"
      TabPicture(1)   =   "comment.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture1"
      Tab(1).ControlCount=   1
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8775
         Left            =   120
         ScaleHeight     =   8775
         ScaleWidth      =   14295
         TabIndex        =   5
         Top             =   360
         Width           =   14295
         Begin VB.PictureBox Picture112 
            BorderStyle     =   0  'None
            Height          =   2775
            Index           =   2
            Left            =   12000
            Picture         =   "comment.frx":0038
            ScaleHeight     =   2775
            ScaleWidth      =   2175
            TabIndex        =   19
            Top             =   5520
            Width           =   2175
         End
         Begin VB.PictureBox Picture4 
            Height          =   2775
            Left            =   3960
            Picture         =   "comment.frx":35D8
            ScaleHeight     =   2715
            ScaleWidth      =   4155
            TabIndex        =   12
            Top             =   5520
            Width           =   4215
         End
         Begin VB.PictureBox Picture3 
            Height          =   2895
            Index           =   4
            Left            =   10440
            Picture         =   "comment.frx":72C6
            ScaleHeight     =   2835
            ScaleWidth      =   3675
            TabIndex        =   11
            Top             =   2040
            Width           =   3735
         End
         Begin VB.PictureBox Picture3 
            Height          =   2775
            Index           =   3
            Left            =   120
            Picture         =   "comment.frx":A23A
            ScaleHeight     =   2715
            ScaleWidth      =   3675
            TabIndex        =   10
            Top             =   5520
            Width           =   3735
         End
         Begin VB.PictureBox Picture3 
            Height          =   2895
            Index           =   2
            Left            =   6480
            Picture         =   "comment.frx":D800
            ScaleHeight     =   2835
            ScaleWidth      =   3675
            TabIndex        =   9
            Top             =   2040
            Width           =   3735
         End
         Begin VB.PictureBox Picture3 
            Height          =   2895
            Index           =   1
            Left            =   3120
            Picture         =   "comment.frx":10D67
            ScaleHeight     =   2835
            ScaleWidth      =   3075
            TabIndex        =   8
            Top             =   2040
            Width           =   3135
         End
         Begin VB.PictureBox Picture3 
            Height          =   2895
            Index           =   0
            Left            =   120
            Picture         =   "comment.frx":13AC8
            ScaleHeight     =   2835
            ScaleWidth      =   2715
            TabIndex        =   7
            Top             =   2040
            Width           =   2775
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "pressing2012@gmail.com"
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
            Left            =   7680
            TabIndex        =   26
            Top             =   8400
            Width           =   3045
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
            Left            =   10200
            TabIndex        =   25
            Top             =   8400
            Width           =   1965
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·„»—„Ã"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   5
            Left            =   12000
            TabIndex        =   24
            Top             =   8400
            Width           =   2175
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFFFFF&
            Height          =   2775
            Left            =   8280
            Top             =   5520
            Width           =   3615
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
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
            Left            =   8280
            TabIndex        =   23
            Top             =   7920
            Width           =   3645
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
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
            Left            =   8400
            TabIndex        =   22
            Top             =   6000
            Width           =   3405
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
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
            Left            =   8400
            TabIndex        =   21
            Top             =   5640
            Width           =   3405
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   $"comment.frx":1671E
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
            Height          =   1575
            Left            =   8280
            TabIndex        =   20
            Top             =   6360
            Width           =   3645
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·‰”Œ… «·”«œ”… 09/2012"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   4
            Left            =   3960
            TabIndex        =   18
            Top             =   8400
            Width           =   4215
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·‰”Œ… «·Œ«„”… 09/2011"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   17
            Top             =   8400
            Width           =   3735
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·‰”Œ… «·—«»⁄… 09/2010"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   10440
            TabIndex        =   16
            Top             =   5040
            Width           =   3735
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·‰”Œ… «·À«·À… 01/2010"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   6480
            TabIndex        =   15
            Top             =   5040
            Width           =   3735
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·‰”Œ… «·À«‰Ì… 09/2009"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   0
            Left            =   3120
            TabIndex        =   14
            Top             =   5040
            Width           =   3135
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·‰”Œ… «·√Ê·Ï 04/2009"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   8
            Left            =   240
            TabIndex        =   13
            Top             =   5040
            Width           =   2535
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   1
            X1              =   0
            X2              =   14280
            Y1              =   5400
            Y2              =   5400
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   0
            X1              =   0
            X2              =   14280
            Y1              =   1920
            Y2              =   1920
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   $"comment.frx":167AB
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1815
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   14055
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8775
         Left            =   -74880
         ScaleHeight     =   8775
         ScaleWidth      =   14295
         TabIndex        =   1
         Top             =   360
         Width           =   14295
         Begin VB.ListBox List2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8580
            Left            =   10680
            TabIndex        =   2
            Top             =   120
            Width           =   3495
         End
         Begin ACTIVESKINLibCtl.Skin Skin1 
            Left            =   480
            OleObjectBlob   =   "comment.frx":16B58
            Top             =   480
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Left            =   8040
            TabIndex        =   4
            Top             =   6480
            Visible         =   0   'False
            Width           =   2295
         End
         Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
            Height          =   8535
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   10455
            URL             =   ""
            rate            =   1
            balance         =   0
            currentPosition =   0
            defaultFrame    =   ""
            playCount       =   1
            autoStart       =   -1  'True
            currentMarker   =   0
            invokeURLs      =   -1  'True
            baseURL         =   ""
            volume          =   50
            mute            =   0   'False
            uiMode          =   "full"
            stretchToFit    =   0   'False
            windowlessVideo =   0   'False
            enabled         =   -1  'True
            enableContextMenu=   -1  'True
            fullScreen      =   0   'False
            SAMIStyle       =   ""
            SAMILang        =   ""
            SAMIFilename    =   ""
            captioningID    =   ""
            enableErrorDialogs=   0   'False
            _cx             =   18441
            _cy             =   15055
         End
      End
   End
End
Attribute VB_Name = "contact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim taille As Double

Private Sub Form_Load()
On Error Resume Next
Me.Top = 100
Skin1.LoadSkin App.Path & "\green.skn"
Skin1.ApplySkin Me.hWnd
Call chargelist2
End Sub
Private Sub chargelist2()
On Error Resume Next
List2.Clear
List2.AddItem "ﬂÌ›Ì… ≈÷«›… »Ì«‰«  «·„ƒ””…"
List2.AddItem "ﬂÌ›Ì… ≈÷«›… »Ì«‰«  «·‘—ﬂ«¡"
List2.AddItem "ﬂÌ›Ì… ≈÷«›… »Ì«‰«  «·„” Œœ„Ì‰"
List2.AddItem "ﬂÌ›Ì… „ «»⁄… Õ”«» «·‘—ﬂ«¡"
List2.AddItem "ﬂÌ›Ì… ≈÷«›… «·√ﬁ”«„"
List2.AddItem "ﬂÌ›Ì… ≈÷«›… «·„Ê«œ"
List2.AddItem "ﬂÌ›Ì… ≈÷«›… «·Ãœ«Ê·"
List2.AddItem "ﬂÌ›Ì… ≈÷«›… «· ·«„Ì–"
List2.AddItem "ﬂÌ›Ì… ≈÷«›… ‰ﬁ«ÿ «· ·«„Ì–"
List2.AddItem "ﬂÌ›Ì… ≈÷«›… €Ì«»«  «· ·«„Ì–"
List2.AddItem "ﬂÌ›Ì… »ÕÀ ⁄«„ ⁄‰ «· ·«„Ì–"
List2.AddItem "ﬂÌ›Ì… ≈÷«›… «·√”« –…"
List2.AddItem "ﬂÌ›Ì… ≈÷«›… ”Ã· Õ÷Ê— √”« –… ” ‘"
List2.AddItem "ﬂÌ›Ì… ≈÷«›… ”Ã· Õ÷Ê— √”« –… ‰"
List2.AddItem "ﬂÌ›Ì… œ›⁄ «Ì—«œ«  «· ·«„Ì–"
List2.AddItem "ﬂÌ›Ì… œ›⁄ „” Õﬁ«  √”« –… ” ‘"
List2.AddItem "ﬂÌ›Ì… œ›⁄ „” Õﬁ«  √”« –… ‰"
List2.AddItem "ﬂÌ›Ì… œ›⁄ Ê«” ·«„ „»«·€ «·‘—ﬂ«¡"
List2.AddItem "ﬂÌ›Ì… œ›⁄ „’—Ê›«  «·„ƒ””…"
List2.AddItem "ﬂÌ›Ì… «·«Ìœ«⁄ ›Ì «·»‰ﬂ Ê«·«” ·«„ „‰Â"
List2.AddItem "ﬂÌ›Ì… „ «»⁄… Õ—ﬂ… «·’‰œÊﬁ"
List2.AddItem "ﬂÌ›Ì… „ «»⁄… ⁄„·Ì«  «·„Õ«”»…"
List2.AddItem "ﬂÌ›Ì… «· Õﬁﬁ „‰ œ›⁄ «·√Ê’«·"
List2.AddItem "ﬂÌ›Ì… „⁄—›… «·„œ›Ê⁄«  «·”‰ÊÌ…"
List2.AddItem "ﬂÌ›Ì… „⁄—›… «Ì—«œ«  «·√ﬁ”«„"
List2.AddItem "ﬂÌ›Ì… «” —Ã«⁄ «·„Õ–Ê›« "
List2.AddItem "ﬂÌ›Ì… ⁄—÷ Ê„”Õ ”‰… „‰ «·«—‘Ì›"
List2.AddItem "ﬂÌ›Ì… «” —Ã«⁄ »Ì«‰«  „‰ «·«—‘Ì›"
List2.AddItem "ﬂÌ›Ì… ⁄„· «·‰”Œ «·«Õ Ì«ÿÌ"
List2.AddItem "ﬂÌ›Ì… › Õ ”‰… œ—«”Ì… ÃœÌœ…"
End Sub

Private Sub List2_Click()
On Error Resume Next
Dim a As Double
a = List2.ListIndex
Label1.Caption = a + 1
Dim s As String
s = App.Path
WindowsMediaPlayer1.URL = s + "\" + Label1.Caption & ".AVI"
End Sub

