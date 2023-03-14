VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form direction 
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
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "«· ﬁ—Ì— «·”‰ÊÌ «·⁄«„"
      TabPicture(0)   =   "direction.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture5"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "»Ì«‰«  «·„” Œœ„Ì‰"
      TabPicture(1)   =   "direction.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture4"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "»Ì«‰«  «·‘—ﬂ«¡"
      TabPicture(2)   =   "direction.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "»Ì«‰«  «·„ÊŸ›Ì‰"
      TabPicture(3)   =   "direction.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture8"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "»Ì«‰«  «·„ƒ””…"
      TabPicture(4)   =   "direction.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Picture1"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8775
         Left            =   -74880
         ScaleHeight     =   8775
         ScaleWidth      =   14295
         TabIndex        =   68
         Top             =   360
         Width           =   14295
         Begin VB.CommandButton Command11 
            Caption         =   "Õ›Ÿ «·»Ì«‰« "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8760
            TabIndex        =   82
            Top             =   2640
            Width           =   3975
         End
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
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
            Left            =   8760
            ScrollBars      =   2  'Vertical
            TabIndex        =   81
            Top             =   1200
            Width           =   3975
         End
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
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
            Left            =   8760
            TabIndex        =   80
            Top             =   1680
            Width           =   3975
         End
         Begin VB.PictureBox Picture9 
            Height          =   2415
            Left            =   8520
            ScaleHeight     =   2355
            ScaleWidth      =   2355
            TabIndex        =   76
            Top             =   4920
            Visible         =   0   'False
            Width           =   2415
            Begin VB.Timer Timer11 
               Enabled         =   0   'False
               Interval        =   50
               Left            =   600
               Top             =   1680
            End
            Begin VB.Timer Timer12 
               Enabled         =   0   'False
               Interval        =   50
               Left            =   120
               Top             =   1680
            End
            Begin VB.Label Label27 
               Caption         =   "Label27"
               Height          =   375
               Left            =   240
               TabIndex        =   79
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label42 
               Caption         =   "Label42"
               Height          =   255
               Left            =   120
               TabIndex        =   78
               Top             =   840
               Width           =   855
            End
            Begin VB.Label Label43 
               Caption         =   "Label43"
               Height          =   375
               Left            =   120
               TabIndex        =   77
               Top             =   1200
               Width           =   1095
            End
         End
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
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
            Left            =   8760
            TabIndex        =   75
            Top             =   2160
            Width           =   3975
         End
         Begin VB.CommandButton Command12 
            Caption         =   "«·€«¡"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8040
            TabIndex        =   74
            Top             =   2640
            Width           =   615
         End
         Begin VB.ComboBox Combo3 
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
            ItemData        =   "direction.frx":008C
            Left            =   5280
            List            =   "direction.frx":0096
            Style           =   2  'Dropdown List
            TabIndex        =   73
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox Text10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
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
            Left            =   1680
            ScrollBars      =   2  'Vertical
            TabIndex        =   72
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Õ›Ÿ «·»Ì«‰« "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   240
            TabIndex        =   71
            Top             =   720
            Width           =   1335
         End
         Begin VB.ComboBox Combo8 
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
            ItemData        =   "direction.frx":00B0
            Left            =   3960
            List            =   "direction.frx":00C6
            Style           =   2  'Dropdown List
            TabIndex        =   70
            Top             =   1200
            Width           =   2775
         End
         Begin VB.ComboBox Combo4 
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
            ItemData        =   "direction.frx":0110
            Left            =   3960
            List            =   "direction.frx":011A
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   720
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.Skin Skin4 
            Left            =   8760
            OleObjectBlob   =   "direction.frx":0134
            Top             =   5400
         End
         Begin MSFlexGridLib.MSFlexGrid grd6 
            Height          =   5175
            Left            =   8040
            TabIndex        =   83
            Top             =   3120
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   9128
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            FixedCols       =   0
            BackColor       =   0
            ForeColor       =   16777215
            BackColorFixed  =   0
            ForeColorFixed  =   16777215
            BackColorBkg    =   0
            RightToLeft     =   -1  'True
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComctlLib.ProgressBar ProgressBar11 
            Height          =   375
            Left            =   8760
            TabIndex        =   84
            Top             =   720
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
         End
         Begin MSComCtl2.DTPicker DT11 
            Height          =   375
            Left            =   1680
            TabIndex        =   85
            Top             =   1200
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   126287873
            CurrentDate     =   41183
         End
         Begin MSFlexGridLib.MSFlexGrid grd12 
            Height          =   5655
            Left            =   240
            TabIndex        =   86
            Top             =   2640
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   9975
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            FixedCols       =   0
            BackColor       =   0
            ForeColor       =   16777215
            BackColorFixed  =   0
            ForeColorFixed  =   16777215
            BackColorBkg    =   0
            RightToLeft     =   -1  'True
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid grd13 
            Height          =   5535
            Left            =   3720
            TabIndex        =   87
            Top             =   2640
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   9763
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            FixedCols       =   0
            BackColor       =   0
            ForeColor       =   16777215
            BackColorFixed  =   0
            ForeColorFixed  =   16777215
            BackColorBkg    =   0
            RightToLeft     =   -1  'True
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Shape Shape6 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            Height          =   8415
            Left            =   7920
            Shape           =   4  'Rounded Rectangle
            Top             =   240
            Width           =   6255
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·Â« ›"
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
            Left            =   11760
            TabIndex        =   107
            Top             =   1680
            Width           =   2295
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«”„ «·„ÊŸ›"
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
            Left            =   12600
            TabIndex        =   106
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·ÊŸÌ›…"
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
            Index           =   7
            Left            =   11760
            TabIndex        =   105
            Top             =   2160
            Width           =   2295
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·«”„"
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
            Left            =   3840
            TabIndex        =   104
            Top             =   240
            Width           =   1455
         End
         Begin VB.Shape Shape7 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            Height          =   8415
            Left            =   120
            Top             =   120
            Width           =   7695
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Left            =   240
            TabIndex        =   103
            Top             =   240
            Width           =   4575
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·„ÊŸ›"
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
            Left            =   6240
            TabIndex        =   102
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Left            =   5400
            TabIndex        =   101
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·⁄„·Ì…"
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
            Left            =   6120
            TabIndex        =   100
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„»·€"
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
            Left            =   2400
            TabIndex        =   99
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«· «—ÌŒ"
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
            Left            =   2400
            TabIndex        =   98
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   2280
            TabIndex        =   97
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„Ã„Ê⁄ «·œ›⁄"
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
            Left            =   3480
            TabIndex        =   96
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   5040
            TabIndex        =   95
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„Ã„Ê⁄ «· ”œÌœ"
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
            Left            =   6240
            TabIndex        =   94
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   240
            TabIndex        =   93
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·—’Ìœ"
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
            Left            =   1320
            TabIndex        =   92
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "”Ã· «·œ›⁄"
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
            Left            =   1200
            TabIndex        =   91
            Top             =   2280
            Width           =   2295
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "”Ã· «· ”œÌœ"
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
            Index           =   9
            Left            =   5400
            TabIndex        =   90
            Top             =   2280
            Width           =   2295
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «·ﬂÊœ"
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
            Index           =   38
            Left            =   6240
            TabIndex        =   89
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·‘Â—"
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
            Index           =   10
            Left            =   4320
            TabIndex        =   88
            Top             =   720
            Width           =   855
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8775
         Left            =   120
         ScaleHeight     =   8775
         ScaleWidth      =   14295
         TabIndex        =   46
         Top             =   360
         Width           =   14295
         Begin VB.PictureBox Picture2 
            Height          =   6255
            Left            =   360
            ScaleHeight     =   6195
            ScaleWidth      =   3315
            TabIndex        =   52
            Top             =   1320
            Visible         =   0   'False
            Width           =   3375
            Begin VB.Timer Timer1 
               Enabled         =   0   'False
               Interval        =   50
               Left            =   360
               Top             =   2880
            End
            Begin VB.Timer Timer4 
               Enabled         =   0   'False
               Interval        =   50
               Left            =   360
               Top             =   2280
            End
            Begin MSComCtl2.DTPicker DT3 
               Height          =   375
               Left            =   240
               TabIndex        =   53
               Top             =   3840
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   661
               _Version        =   393216
               Format          =   126287873
               CurrentDate     =   71615
            End
            Begin VB.Label Label6 
               Caption         =   "Label6"
               Height          =   375
               Left            =   240
               TabIndex        =   57
               Top             =   1680
               Width           =   2295
            End
            Begin VB.Label Label5 
               Caption         =   "Label5"
               Height          =   375
               Left            =   240
               TabIndex        =   56
               Top             =   1200
               Width           =   2295
            End
            Begin VB.Label Label4 
               Caption         =   "Label4"
               Height          =   375
               Left            =   240
               TabIndex        =   55
               Top             =   720
               Width           =   2295
            End
            Begin VB.Label Label3 
               Caption         =   "Label3"
               Height          =   375
               Left            =   240
               TabIndex        =   54
               Top             =   240
               Width           =   2175
            End
         End
         Begin VB.TextBox Text 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
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
            Left            =   4920
            TabIndex        =   51
            Top             =   4200
            Width           =   2055
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Õ›Ÿ «·»Ì«‰« "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5040
            TabIndex        =   50
            Top             =   5760
            Width           =   4215
         End
         Begin VB.TextBox Text 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
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
            Left            =   4920
            TabIndex        =   49
            Top             =   3720
            Width           =   2055
         End
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
            ItemData        =   "direction.frx":0368
            Left            =   4920
            List            =   "direction.frx":037E
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   3240
            Width           =   2055
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
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
            Height          =   825
            Left            =   4920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   47
            Top             =   2310
            Width           =   3615
         End
         Begin ACTIVESKINLibCtl.Skin Skin1 
            Left            =   120
            OleObjectBlob   =   "direction.frx":03C8
            Top             =   120
         End
         Begin MSComCtl2.DTPicker DT1 
            Height          =   375
            Left            =   4920
            TabIndex        =   58
            Top             =   4680
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   126287873
            CurrentDate     =   41183
         End
         Begin MSComCtl2.DTPicker DT2 
            Height          =   375
            Left            =   4920
            TabIndex        =   59
            Top             =   5160
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   126287873
            CurrentDate     =   41547
         End
         Begin VB.Label Label331 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "„·«ÕŸ… Â«„…"
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
            Left            =   5880
            TabIndex        =   67
            Top             =   6600
            Width           =   2655
         End
         Begin VB.Label Label331 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   $"direction.frx":05FC
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
            Height          =   975
            Index           =   2
            Left            =   5040
            TabIndex        =   66
            Top             =   7320
            Width           =   4455
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì…"
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
            Index           =   54
            Left            =   7320
            TabIndex        =   65
            Top             =   4680
            Width           =   2055
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            Height          =   4335
            Left            =   4800
            Shape           =   4  'Rounded Rectangle
            Top             =   2040
            Width           =   4695
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰”»… √”« –… «·‰”»…"
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
            Left            =   6720
            TabIndex        =   64
            Top             =   4200
            Width           =   2655
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰”»… «·„ƒ””…"
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
            Left            =   6720
            TabIndex        =   63
            Top             =   3720
            Width           =   2655
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”‰… «·œ—«”Ì…"
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
            Left            =   7440
            TabIndex        =   62
            Top             =   3240
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«”„ «·„ƒ””…"
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
            Height          =   855
            Left            =   8400
            TabIndex        =   61
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «‰ Â«¡ «·”‰… «·œ—«”Ì…"
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
            Index           =   76
            Left            =   6840
            TabIndex        =   60
            Top             =   5160
            Width           =   2535
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8775
         Left            =   -74880
         ScaleHeight     =   8775
         ScaleWidth      =   14295
         TabIndex        =   43
         Top             =   360
         Width           =   14295
         Begin VB.CommandButton Command8 
            Caption         =   "”Õ»"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3600
            TabIndex        =   117
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox Text13 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
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
            IMEMode         =   3  'DISABLE
            Left            =   6720
            TabIndex        =   115
            Text            =   "5"
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox Text11 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
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
            IMEMode         =   3  'DISABLE
            Left            =   9960
            TabIndex        =   113
            Text            =   "8"
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton Command10 
            Caption         =   "”Õ»"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   110
            Top             =   1080
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.CommandButton Command9 
            Caption         =   "⁄—÷"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5640
            TabIndex        =   109
            Top             =   1080
            Width           =   2655
         End
         Begin VB.ComboBox Combo5 
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
            Left            =   8400
            Style           =   2  'Dropdown List
            TabIndex        =   108
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CommandButton Command7 
            Caption         =   "⁄—÷ «· ﬁ—Ì— «·”‰ÊÌ"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   44
            Top             =   360
            Width           =   4215
         End
         Begin MSFlexGridLib.MSFlexGrid grd3 
            Height          =   6975
            Left            =   120
            TabIndex        =   111
            Top             =   1680
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   12303
            _Version        =   393216
            FixedRows       =   0
            FixedCols       =   0
            BackColor       =   0
            ForeColor       =   16777215
            BackColorFixed  =   0
            ForeColorFixed  =   16777215
            ForeColorSel    =   8388608
            BackColorBkg    =   0
            AllowBigSelection=   0   'False
            BorderStyle     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„ÿ—Êœ"
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
            Left            =   7920
            TabIndex        =   116
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰«ÃÕ"
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
            Left            =   11400
            TabIndex        =   114
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·ﬁ”„"
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
            Index           =   25
            Left            =   9600
            TabIndex        =   112
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label26 
            Caption         =   "100"
            Height          =   375
            Left            =   240
            TabIndex        =   45
            Top             =   120
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8775
         Left            =   -74880
         ScaleHeight     =   8775
         ScaleWidth      =   14295
         TabIndex        =   14
         Top             =   360
         Width           =   14295
         Begin VB.CommandButton Command4 
            Caption         =   "Õ›Ÿ «·»Ì«‰« "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6960
            TabIndex        =   28
            Top             =   4800
            Width           =   2415
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
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
            Height          =   825
            Left            =   4920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   27
            Top             =   2310
            Width           =   3615
         End
         Begin VB.CheckBox Check2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            Caption         =   "«·— »…"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   7920
            MaskColor       =   &H00000000&
            TabIndex        =   26
            Top             =   4080
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            Caption         =   "«·— »…"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   9120
            MaskColor       =   &H00000000&
            TabIndex        =   25
            Top             =   4080
            Width           =   255
         End
         Begin VB.CheckBox Check3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            Caption         =   "«·— »…"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6720
            MaskColor       =   &H00000000&
            TabIndex        =   24
            Top             =   4080
            Width           =   255
         End
         Begin VB.CheckBox Check4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            Caption         =   "«·— »…"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   5640
            MaskColor       =   &H00000000&
            TabIndex        =   23
            Top             =   4080
            Width           =   255
         End
         Begin VB.CheckBox Check5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            Caption         =   "«·— »…"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   9120
            MaskColor       =   &H00000000&
            TabIndex        =   22
            Top             =   4440
            Width           =   255
         End
         Begin VB.CheckBox Check6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            Caption         =   "«·— »…"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   7920
            MaskColor       =   &H00000000&
            TabIndex        =   21
            Top             =   4440
            Width           =   255
         End
         Begin VB.CheckBox Check7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            Caption         =   "«·— »…"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6720
            MaskColor       =   &H00000000&
            TabIndex        =   20
            Top             =   4440
            Width           =   255
         End
         Begin VB.CheckBox Check8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            Caption         =   "«·— »…"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   5640
            MaskColor       =   &H00000000&
            TabIndex        =   19
            Top             =   4440
            Width           =   255
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
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
            IMEMode         =   3  'DISABLE
            Left            =   7080
            PasswordChar    =   "*"
            TabIndex        =   18
            Top             =   3240
            Width           =   1455
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
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
            IMEMode         =   3  'DISABLE
            Left            =   4920
            PasswordChar    =   "*"
            TabIndex        =   17
            Top             =   3240
            Width           =   1455
         End
         Begin VB.CommandButton Command5 
            Caption         =   "«·€«¡"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4920
            TabIndex        =   16
            Top             =   4800
            Width           =   615
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Õ–›"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5640
            TabIndex        =   15
            Top             =   4800
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.Skin Skin3 
            Left            =   120
            OleObjectBlob   =   "direction.frx":0692
            Top             =   120
         End
         Begin MSFlexGridLib.MSFlexGrid grd1 
            Height          =   3135
            Left            =   4920
            TabIndex        =   29
            Top             =   5280
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   5530
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            FixedCols       =   0
            BackColor       =   0
            ForeColor       =   16777215
            BackColorFixed  =   0
            ForeColorFixed  =   16777215
            BackColorBkg    =   0
            RightToLeft     =   -1  'True
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   255
            Left            =   4920
            TabIndex        =   30
            Top             =   3720
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            Height          =   6615
            Left            =   4800
            Shape           =   4  'Rounded Rectangle
            Top             =   2040
            Width           =   4695
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "’·«ÕÌ«  «·„” Œœ„"
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
            Index           =   6
            Left            =   6720
            TabIndex        =   42
            Top             =   3720
            Width           =   2655
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ﬂ·„… «·”—"
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
            Left            =   8160
            TabIndex        =   41
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«”„ «·„” Œœ„"
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
            Height          =   855
            Left            =   8400
            TabIndex        =   40
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "≈⁄«œ…"
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
            Left            =   6240
            TabIndex        =   39
            Top             =   3240
            Width           =   735
         End
         Begin VB.Label Label68 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·≈œ«—…"
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
            Index           =   0
            Left            =   8160
            TabIndex        =   38
            Top             =   4080
            Width           =   855
         End
         Begin VB.Label Label68 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·‘—ﬂ«¡"
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
            Index           =   1
            Left            =   6960
            TabIndex        =   37
            Top             =   4080
            Width           =   855
         End
         Begin VB.Label Label68 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·√ﬁ”«„"
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
            Index           =   2
            Left            =   5760
            TabIndex        =   36
            Top             =   4080
            Width           =   855
         End
         Begin VB.Label Label68 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«· ·«„Ì–"
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
            Index           =   3
            Left            =   4680
            TabIndex        =   35
            Top             =   4080
            Width           =   855
         End
         Begin VB.Label Label68 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„Õ«”»…"
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
            Index           =   4
            Left            =   5760
            TabIndex        =   34
            Top             =   4440
            Width           =   855
         End
         Begin VB.Label Label68 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·’‰œÊﬁ"
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
            Index           =   5
            Left            =   6960
            TabIndex        =   33
            Top             =   4440
            Width           =   855
         End
         Begin VB.Label Label68 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·√”« –…"
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
            Index           =   6
            Left            =   8160
            TabIndex        =   32
            Top             =   4440
            Width           =   855
         End
         Begin VB.Label Label68 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·«—‘Ì›"
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
            Index           =   7
            Left            =   4680
            TabIndex        =   31
            Top             =   4440
            Width           =   855
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8775
         Left            =   -74880
         ScaleHeight     =   8775
         ScaleWidth      =   14295
         TabIndex        =   1
         Top             =   360
         Width           =   14295
         Begin VB.CommandButton Command1 
            Caption         =   "Õ›Ÿ «·»Ì«‰« "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6240
            TabIndex        =   6
            Top             =   4680
            Width           =   3015
         End
         Begin VB.ComboBox Combo2 
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
            ItemData        =   "direction.frx":08C6
            Left            =   4920
            List            =   "direction.frx":08DC
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   3720
            Width           =   2055
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
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
            Height          =   825
            Left            =   4920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   2310
            Width           =   3615
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
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
            Left            =   4920
            TabIndex        =   3
            Top             =   4200
            Width           =   2055
         End
         Begin VB.CommandButton Command3 
            Caption         =   "«·€«¡"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4920
            TabIndex        =   2
            Top             =   4680
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.Skin Skin2 
            Left            =   120
            OleObjectBlob   =   "direction.frx":0926
            Top             =   120
         End
         Begin MSFlexGridLib.MSFlexGrid grd9 
            Height          =   3255
            Left            =   4920
            TabIndex        =   7
            Top             =   5160
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   5741
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            FixedCols       =   0
            BackColor       =   0
            ForeColor       =   16777215
            BackColorFixed  =   0
            ForeColorFixed  =   16777215
            BackColorBkg    =   0
            RightToLeft     =   -1  'True
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComctlLib.ProgressBar ProgressBar4 
            Height          =   375
            Left            =   5160
            TabIndex        =   8
            Top             =   1560
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            Height          =   6615
            Left            =   4800
            Shape           =   4  'Rounded Rectangle
            Top             =   2040
            Width           =   4695
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·Â« ›"
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
            Left            =   7080
            TabIndex        =   13
            Top             =   4200
            Width           =   2295
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰”»… «·‘—«ﬂ… «·„ «Õ…"
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
            Left            =   7080
            TabIndex        =   12
            Top             =   3720
            Width           =   2295
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰”»… «·‘—«ﬂ… «·„” Œœ„…"
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
            Left            =   7080
            TabIndex        =   11
            Top             =   3240
            Width           =   2295
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«”„ «·‘—Ìﬂ"
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
            Height          =   855
            Left            =   8640
            TabIndex        =   10
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "78.6"
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
            Left            =   4920
            TabIndex        =   9
            Top             =   3240
            Width           =   2055
         End
      End
   End
End
Attribute VB_Name = "direction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public co2 As ADODB.Connection
Public cr2 As ADODB.Recordset
Public be As ADODB.Recordset
Public ce As ADODB.Recordset
Public nn As ADODB.Recordset
Dim anes As String
Dim data As New Access.Application
Function cont2()
Set co2 = New ADODB.Connection
Set cr2 = New ADODB.Recordset
Set be = New ADODB.Recordset
Set ce = New ADODB.Recordset
Set nn = New ADODB.Recordset
co2.Provider = "microsoft.jet.oledb.4.0; jet oledb:database password=7346804"
anes = "C" + face.SBB1.Panels(9).Text
co2.ConnectionString = App.Path & "\" & anes & ".mdb"
co2.Open
cr2.Open "select*from Tcarts", co2, adOpenKeyset, adLockOptimistic
be.Open "select*from Tbulletin", co2, adOpenKeyset, adLockOptimistic
ce.Open "select*from Tcartes", co2, adOpenKeyset, adLockOptimistic
nn.Open "select*from Tnni order by aut ASC", co2, adOpenKeyset, adLockOptimistic
End Function
Public Sub chargec1()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim k As Double
Dim s As Double
Combo1.Clear
Call cont
Do While Not an.EOF
Combo1.AddItem an!ann
an.MoveNext
Loop
If sr!ann <> "" Then
Text1.Text = sr!eco
Combo1.Text = sr!ann
Text(0).Text = sr!pou
Label26.Caption = sr!pou
'Combo1.Enabled = False
End If
If sr!dat <> "Rien" Then
DT1.Value = sr!dat
'DT1.Enabled = False
DT2.Value = sr!dtf
'DT2.Enabled = False
End If
s = 0
Call cont
s = s + cl.RecordCount
s = s + mt.RecordCount
s = s + em.RecordCount
s = s + et.RecordCount
s = s + sr.RecordCount
s = s + nt.RecordCount
s = s + cf.RecordCount
s = s + cf1.RecordCount
s = s + ab.RecordCount
s = s + pr.RecordCount
s = s + ps.RecordCount
s = s + cd.RecordCount
s = s + ce.RecordCount
s = s + rc.RecordCount
s = s + ca.RecordCount
s = s + jr.RecordCount
s = s + pf.RecordCount
s = s + dp.RecordCount
s = s + pa.RecordCount
s = s + pp.RecordCount
s = s + bn.RecordCount
s = s + cr.RecordCount
s = s + an.RecordCount
s = s + ut.RecordCount
If (s > 31) Then
Combo1.Enabled = False
DT1.Enabled = False
DT2.Enabled = False
Else
Combo1.Enabled = True
DT1.Enabled = True
DT2.Enabled = True
End If

End Sub
Private Sub calculepourcentage()
On Error Resume Next
Dim a As Double
Dim k As Double
Dim j As Double
Dim i As Double
a = Label9.Caption
k = 100 - a
k = k * 10
j = -0.1
Combo2.Clear
For i = 0 To k
j = j + 0.1
MyNumber = Round(j, 1)
j = MyNumber
Combo2.AddItem j
Next i

End Sub
Private Sub Combo1_Change()
On Error Resume Next
Call cont
Do While Not an.EOF
If an!ann = Combo1.Text Then
Label3.Caption = an!an1
Label4.Caption = an!an2
DT1.Year = an!an1
DT2.Year = an!an2
Exit Sub
End If
an.MoveNext
Loop

End Sub

Private Sub Combo1_Click()
On Error Resume Next
Combo1_Change
End Sub

Private Sub Combo3_Change()
On Error Resume Next
Combo4.Visible = False
Label31(10).Visible = False
Call chargec8
If Combo3.Text = " ”œÌœ —« »" Then
Call chargec4
Combo4.Visible = True
Label31(10).Visible = True
Exit Sub
End If
'Text10.SetFocus
End Sub

Private Sub Combo3_Click()
On Error Resume Next
Combo3_Change
End Sub

Private Sub Combo4_Change()
On Error Resume Next
Text10.SetFocus

End Sub

Private Sub Combo4_Click()
On Error Resume Next
Combo4_Change
End Sub

Private Sub Combo5_Change()
On Error Resume Next
grd3.Clear
grd3.Rows = 1
End Sub

Private Sub Combo5_Click()
On Error Resume Next
Combo5_Change
End Sub

Private Sub Combo8_Change()
On Error Resume Next
Call cont
Do While Not cd.EOF
If Combo8.Text = cd!cod Then
Label42.Caption = cd!dec
Label43.Caption = cd!cas
Exit Sub
End If
If Combo8.Text = cd!dec Then
Label42.Caption = cd!cod
Label43.Caption = cd!cas
Exit Sub
End If
cd.MoveNext
Loop

End Sub

Private Sub Combo8_Click()
On Error Resume Next
Combo8_Change
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim a As Double
Dim b As Double
Text2.Text = Trim(Text2.Text)
Text3.Text = Trim(Text3.Text)
If Text2.Text = "" Then
MsgBox "«œŒ· «”„ «·‘—Ìﬂ", vbCritical
Text2.SetFocus
Exit Sub
End If
If Combo2.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— ‰”»… «·‘—«ﬂ…", vbCritical
Exit Sub
End If
If Text3.Text = "" Then
MsgBox "«œŒ· Â« › «·‘—Ìﬂ", vbCritical
Text3.SetFocus
Exit Sub
End If
Call cont
Do While Not pa.EOF
If pa!aut = Label5.Caption Then
pa!nom = Text2.Text
pa!pou = Combo2.Text
pa!tel = Text3.Text
pa.Update
ProgressBar4.Value = 0
ProgressBar4.Visible = True
Timer4.Enabled = True
Exit Sub
End If
pa.MoveNext
Loop
pa.AddNew
pa!mtr = pa!aut
pa!nom = Text2.Text
pa!pou = Combo2.Text
pa!tel = Text3.Text
pa!act = "0"
pa.Update
ProgressBar4.Value = 0
ProgressBar4.Visible = True
Timer4.Enabled = True

End Sub

Private Sub Command10_Click()
On Error GoTo u
Dim i As Double
Dim j As Double
Dim n As Double
Dim k As Double
Dim d As Double
Dim sd As Double
If Combo5.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·ﬁ”„", vbCritical
Exit Sub
End If
If grd3.Rows < 2 Then
MsgBox "·«  ÊÃœ »Ì«‰« ", vbCritical
Exit Sub
End If
FileCopy App.Path & "\report00000.xls", App.Path & "\ReportAnuelle.xls"
Command10.Enabled = False
n = grd3.Rows
Set kb = CreateObject("Excel.application")
kb.Workbooks.Open (App.Path & "\ReportAnuelle.xls")
kb.Visible = True
For i = 0 To n - 1
For j = 0 To 21
grd3.row = i
grd3.Col = j
k = 22 - j
kb.Workbooks("ReportAnuelle").Sheets(1).Cells(i + 5, k).Value = grd3.Text
Next j
Next i
kb.Workbooks("ReportAnuelle").Sheets(1).Range("L3").Value = Combo5.Text
'kb.Workbooks("fiche de presences").Sheets(1).Range("B5").Value = DT11.Value
'Workbooks("Etudiants").Close savechanges:=False
'Worksheets(1).Activate
Command10.Enabled = True
Exit Sub
u:
MsgBox "ÌÃ» «€·«ﬁ ’›Õ… «ﬂ”· «‰ ﬂ«‰  „› ÊÕ…", vbExclamation
Command10.Enabled = True

End Sub

Private Sub Command11_Click()
On Error Resume Next
Text7.Text = Trim(Text7.Text)
Text8.Text = Trim(Text8.Text)
Text9.Text = Trim(Text9.Text)
If Text7.Text = "" Then
MsgBox "«œŒ· «”„ «·„ÊŸ›", vbCritical
Text7.SetFocus
Exit Sub
End If
If Text8.Text = "" Then
MsgBox "«œŒ· —ﬁ„ «·Â« ›", vbCritical
Text8.SetFocus
Exit Sub
End If
If Text9.Text = "" Then
MsgBox "«œŒ· —ﬁ„ «·Â« ›", vbCritical
Text9.SetFocus
Exit Sub
End If
Call cont
Do While Not fc.EOF
If Label27.Caption = fc!mtr Then
fc!nom = Text7.Text
fc!tel = Text8.Text
fc!foc = Text9.Text
fc.Update
ProgressBar11.Value = 0
Timer11.Enabled = True
Exit Sub
End If
fc.MoveNext
Loop
fc.AddNew
fc!nom = Text7.Text
fc!tel = Text8.Text
fc!foc = Text9.Text
fc.Update
ProgressBar11.Value = 0
Timer11.Enabled = True
End Sub

Private Sub Command12_Click()
On Error Resume Next
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text7.SetFocus
Label27.Caption = ""
ProgressBar11.Value = 0
Timer11.Enabled = False

End Sub

Private Sub Command13_Click()
On Error Resume Next
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim e As Double
Dim au As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim dat4 As Date
Dim annes As String
Dim mca As Double
Dim mnv As Double
Text10.Text = Trim(Text10.Text)
If Label33.Caption = "" Then
MsgBox "ÌÃ»  ÕœÌœ «”„ «·„ÊŸ›", vbCritical
Exit Sub
End If
If Combo3.Text = "" Then
MsgBox "Õœœ ‰Ê⁄ «·⁄„·Ì…", vbCritical
Exit Sub
End If
If Combo4.Text = "" And Combo4.Visible = True Then
MsgBox "Õœœ «·‘Â—", vbCritical
Exit Sub
End If
If Text10.Text = "" Then
MsgBox "«œŒ· «·„»·€", vbCritical
Text10.SetFocus
Exit Sub
End If
If Combo8.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— ﬂÊœ «·⁄„·Ì…", vbCritical
Exit Sub
End If
If Combo3.Text = " ”œÌœ —« »" Then
d = Label40.Caption
If d >= 0 Then
MsgBox "ÌÃ» œ›⁄ «·„»·€ ﬁ»·  ”œÌœÂ", vbCritical
Exit Sub
End If
e = Text10.Text
d = (d * -1)
If e > d Then
MsgBox "„»·€ «· ”œÌœ ÌÃ» √‰ ÌﬂÊ‰ √’€— √Ê „”«Ê ·„»·€ «·œ›⁄", vbCritical
Exit Sub
End If
Call cont
Do While Not pfc.EOF
If pfc!mtr = Label33.Caption And pfc!moi = Combo4.Text Then
MsgBox "·ﬁœ  „  ”œÌœ —« » Â–« «·‘Â— ”«»ﬁ«", vbCritical
Exit Sub
End If
pfc.MoveNext
Loop
Else
'**** controle caisse ajouter
mca = sr!cca
mnv = Text10.Text
If mnv > mca Then
MsgBox "€Ì— „„ﬂ‰.. ·√‰Â ·« ÌÊÃœ ›Ì «·’‰œÊﬁ Õ«·Ì« Â–« «·„»·€", vbCritical
Exit Sub
End If
'**** end controle caisse
End If
Call cont
'**** controle Date Ajouter
annes = sr!ann      'annee scolaire
dat1 = sr!dat       'debut annee scolaire
dat2 = sr!dtf       'Fin annees scolaire
dat3 = DT11.Value    'Date Modifiable
dat4 = Date         'Date Machine
If dat3 < dat1 Then
MsgBox "€Ì— „„ﬂ‰.. «· «—ÌŒ «·„œŒ· ”«»ﬁ · «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì… " + annes, vbCritical + arabic
Exit Sub
End If
If dat3 > dat2 Then
MsgBox "€Ì— „„ﬂ‰.. «· «—ÌŒ «·„œŒ· „ √Œ— ⁄‰  «—ÌŒ ‰Â«Ì… «·”‰… «·œ—«”Ì… " + annes, vbCritical + arabic
Exit Sub
End If
If dat3 > dat4 Then
MsgBox "€Ì— „„ﬂ‰.. «· «—ÌŒ «·„œŒ· „ ﬁœ„ ⁄‰  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“)", vbCritical + arabic
Exit Sub
End If
'**** end controle Date
If Combo3.Text = " ”œÌœ —« »" Then
Call paysal_fonctionnaires
Else
Call paymon_fonctionnaires
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim a As Double
Dim b As Double
Dim dat1 As Date
Dim dat2 As Date
Text1.Text = Trim(Text1.Text)
If Text1.Text = "" Then
MsgBox "«œŒ· «”„ «·„ƒ””…", vbCritical
Text1.SetFocus
Exit Sub
End If
If Combo1.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·”‰… «·œ—«”Ì…", vbCritical
Exit Sub
End If
If Text(0).Text = "" Then
MsgBox "«œŒ· ‰”»… «·„ƒ””…", vbCritical
Text(0).SetFocus
Exit Sub
End If
a = DT1.Year
If a < Val(Label3.Caption) Then
MsgBox "«· «—ÌŒ «·„œŒ· ”«»ﬁ ··”‰… «·œ—«”Ì… " + Combo1.Text, vbCritical
Exit Sub
End If
If a > Val(Label4.Caption) Then
MsgBox "«· «—ÌŒ «·„œŒ· „ √Œ— ⁄‰ «·”‰… «·œ—«”Ì… " + Combo1.Text, vbCritical
Exit Sub
End If
dat1 = DT1.Value
dat2 = DT2.Value
If dat1 > dat2 Then
MsgBox " «—ÌŒ »œ«Ì… «·”‰…«·œ—«”Ì… ÌÃ» √‰ ÌﬂÊ‰ ﬁ»·  «—ÌŒ ‰Â«Ì… «·”‰…«·œ—«”Ì…", vbCritical
Exit Sub
End If
Call cont
sr!eco = Text1.Text
sr!ann = Combo1.Text
sr!pou = Text(0).Text
sr!dat = DT1.Value
sr!dtf = DT2.Value
sr.Update
'Combo1.Enabled = False
'DT1.Enabled = False
MsgBox " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ", vbInformation
face.SBB1.Panels(13).Text = Text1.Text
Label26.Caption = Text(0).Text

End Sub

Private Sub Command3_Click()
On Error Resume Next
Text2.Text = ""
Text2.SetFocus
Label5.Caption = ""
Text3.Text = ""
grd9.Visible = False
grd9.Clear
grd9.Rows = 1
Call chargegrd9
grd9.Visible = True
ProgressBar4.Value = 0
ProgressBar4.Visible = True
Timer4.Enabled = False

End Sub

Private Sub Command4_Click()
On Error Resume Next
Text4.Text = Trim(Text4.Text)
Text5.Text = Trim(Text5.Text)
Text6.Text = Trim(Text6.Text)
If Text4.Text = "" Then
MsgBox "«œŒ· «”„ «·„” Œœ„", vbCritical
Text4.SetFocus
Exit Sub
End If
If Text5.Text = "" Then
MsgBox "«œŒ· ﬂ·… «·”—", vbCritical
Text5.SetFocus
Exit Sub
End If
If Text6.Text <> Text5.Text Then
MsgBox "ﬂ·„ « «·”— €Ì— „ ÿ«»ﬁ Ì‰", vbCritical
Text6.SetFocus
Exit Sub
End If
If Check1.Value = 0 And Check2.Value = 0 And Check3.Value = 0 And Check4.Value = 0 And Check5.Value = 0 And Check6.Value = 0 And Check7.Value = 0 And Check8.Value = 0 Then
MsgBox "ÌÃ»  ÕœÌœ Ê·Ê ’·«ÕÌ… Ê«Õœ… ··„” Œœ„", vbCritical
Exit Sub
End If
Call cont
Do While Not ut.EOF
If ut!aut <> Label6.Caption And Text4.Text = ut!uti Then
MsgBox "·ﬁœ  „ ÕÃ“ Â–« «·«”„ ”«»ﬁ«", vbCritical
Exit Sub
End If
ut.MoveNext
Loop
If Label6.Caption <> "" Then
Call cont
Do While Not ut.EOF
If ut!aut = Label6.Caption Then
ut!uti = Text4.Text
ut!mot = Text5.Text
ut!Dir = Check1.Value
ut!par = Check2.Value
ut!cla = Check3.Value
ut!etu = Check4.Value
ut!pro = Check5.Value
ut!cai = Check6.Value
ut!com = Check7.Value
ut!Arc = Check8.Value
ut.Update
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
ut.MoveNext
Loop
End If
ut.AddNew
ut!uti = Text4.Text
ut!mot = Text5.Text
ut!Dir = Check1.Value
ut!par = Check2.Value
ut!cla = Check3.Value
ut!etu = Check4.Value
ut!pro = Check5.Value
ut!cai = Check6.Value
ut!com = Check7.Value
ut!Arc = Check8.Value
ut.Update
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True

End Sub

Private Sub Command5_Click()
On Error Resume Next
Text4.Text = ""
Text4.SetFocus
Label6.Caption = ""
Text5.Text = ""
Text6.Text = ""
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
Check5.Value = 0
Check6.Value = 0
Check7.Value = 0
Check8.Value = 0
grd1.Visible = False
grd1.Clear
grd1.Rows = 1
Call chargegrd1
grd1.Visible = True
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = False

End Sub

Private Sub Command6_Click()
On Error Resume Next
If grd1.Rows < 3 Then
MsgBox "Â–« ÂÊ «·„” Œœ„ «·ÊÕÌœ «·„ »ﬁÌ , ·«Ì„ﬂ‰ Õ–›Â , ≈–« √—œ  Õ–›Â ›ÌÃ» ≈÷«›… „” Œœ„ ¬Œ—", vbCritical
Text4.SetFocus
Exit Sub
End If
If Label6.Caption = "" Then
MsgBox "«÷€ÿ ⁄·Ï «”„ «·„” Œœ„ «·–Ì  —Ìœ Õ–›Â", vbCritical
Exit Sub
End If
g = MsgBox("Â·  —Ìœ «·Õ–› Õﬁ« ", vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
Call cont
Do While Not ut.EOF
If Label6.Caption = ut!aut Then
ut.Delete
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
ut.MoveNext
Loop
End If
End Sub


Private Sub Command7_Click_()
On Error Resume Next
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim e As Double
Dim f As Double
Dim j As Double
Dim dat1 As Date
Dim dat2 As Date
'**** etudiants
Call cont
Label10.Caption = cl.RecordCount
Label11.Caption = pr.RecordCount
Label14.Caption = et.RecordCount
a = et.RecordCount
b = 0
Do While Not et.EOF
If et!sex = "–ﬂ—" Then
b = b + 1
End If
If et!sex = "√‰ÀÏ" Then
c = c + 1
End If
et.MoveNext
Loop
'c = a - b
Label15.Caption = b
Label16.Caption = c
d = (b * 100) / a
MyNumber = Round(d, 0)
d = MyNumber
Label17.Caption = d
ProgressBar2.Value = d
c = (c * 100) / a
MyNumber = Round(c, 0)
c = MyNumber
Label18.Caption = c
ProgressBar3.Value = c
'**** end etudiants
'**** recettes
a = 0
b = 0
c = 0
d = 0
Call cont
Do While Not rc.EOF
a = rc!mon
b = b + a
rc.MoveNext
Loop
a = 0
e = 0
f = 0
Call cont
Do While Not ps.EOF
If ps!cas <> "p" Then
a = ps!tot
c = c + a
a = ps!prm
c = c + a
a = ps!rtr
c = c + a
End If
If ps!cas <> "m" Then
e = e + 15
Else
f = ps!nbr
e = e + f
End If
ps.MoveNext
Loop
a = b - c
d = Label26.Caption
d = (a * d) / 100
MyNumber = Round(d, 0)
d = MyNumber
d = a - d
c = c + d
a = 0
d = 0
Call cont
Do While Not dp.EOF
a = dp!mon
d = d + a
dp.MoveNext
Loop
a = b + c + d
b = (b * 100) / a
MyNumber = Round(b, 0)
b = MyNumber
Label23.Caption = b
ProgressBar10.Value = b
c = (c * 100) / a
MyNumber = Round(c, 0)
c = MyNumber
Label25.Caption = c
ProgressBar8.Value = c
d = (d * 100) / a
MyNumber = Round(d, 0)
d = MyNumber
Label24.Caption = d
ProgressBar9.Value = d
'**** end recettes
'**** ensegnements
dat1 = sr!dat
dat2 = Date
j = dat2 - dat1
f = 0
If j > 7 Then
f = j / 7
MyNumber = Round(f, 0)
f = MyNumber
End If
j = j - f
f = j * 8
e = f - e
If e < 0 Then
e = e * -1
End If
e = (e * 100) / f
MyNumber = Round(e, 0)
e = MyNumber
Label19.Caption = e
ProgressBar5.Value = e
If e > 90 Then
Label22.Caption = "„„ «“"
ElseIf e > 70 Then
Label22.Caption = "⁄«·"
ElseIf e > 50 Then
Label22.Caption = "ÃÌœ"
ElseIf e > 30 Then
Label22.Caption = "„ﬁ»Ê·"
ElseIf e > 10 Then
Label22.Caption = "÷⁄Ì›"
ElseIf e > 1 Then
Label22.Caption = "„ œ‰Ì"
Else
Label22.Caption = "„‰⁄œ„"
End If
'**** end enseignements
'**** succes and low
a = 0
b = 0
c = 0
Call cont2
Do While Not be.EOF
a = be!moy
If a >= 9 Then
b = b + 1
Else
c = c + 1
End If
be.MoveNext
Loop
a = b + c
b = (b * 100) / a
MyNumber = Round(b, 0)
b = MyNumber
Label20.Caption = b
ProgressBar6.Value = b
c = (c * 100) / a
MyNumber = Round(c, 0)
c = MyNumber
Label21.Caption = c
ProgressBar7.Value = c
'**** end succes and low
End Sub


Private Sub Command7_Click()
Dim ane As String
Dim rest As Double
Dim out As Double
rest = Text11.Text
out = Text13.Text
Call cont2
Do While Not be.EOF
If (be!moy >= rest) Then
be!obs17 = "‰«ÃÕ"
be.Update
ElseIf (be!moy <= out) Then
be!obs17 = "„ÿ—Êœ"
be.Update
Else
be!obs17 = "—«”»"
be.Update
End If
be.MoveNext
Loop
Call cont2
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
data.DoCmd.Maximize
data.DoCmd.OpenReport "Tdecision", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
data.Visible = True
Set data = Nothing

End Sub

Private Sub Command8_Click()
'On Error GoTo u
Dim i As Double
Dim j As Double
Dim n As Double
Dim k As Double
Dim d As Double
Dim sd As Double
If Combo5.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·ﬁ”„", vbCritical
Exit Sub
End If
If grd3.Rows < 2 Then
MsgBox "·«  ÊÃœ »Ì«‰« ", vbCritical
Exit Sub
End If
FileCopy App.Path & "\RIM010.xls", App.Path & "\RIM.xls"
Command8.Enabled = False
n = grd3.Rows
Set kb = CreateObject("Excel.application")
kb.Workbooks.Open (App.Path & "\RIM.xls")
kb.Visible = True
For i = 1 To n - 1
For j = 0 To 6
grd3.row = i
grd3.Col = j
k = j + 1
kb.Workbooks("RIM").Sheets(1).Cells(i + 7, k).Value = grd3.Text
Next j
Next i
kb.Workbooks("RIM").Sheets(1).Range("D3").Value = face.SBB1.Panels(9).Text
kb.Workbooks("RIM").Sheets(1).Range("B3").Value = face.SBB1.Panels(13).Text
kb.Workbooks("RIM").Sheets(1).Range("C5").Value = Combo5.Text
Command8.Enabled = True
Exit Sub
u:
MsgBox "ÌÃ» «€·«ﬁ ’›Õ… «ﬂ”· «‰ ﬂ«‰  „› ÊÕ…", vbExclamation
Command8.Enabled = True

End Sub

Private Sub Command9_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim n As Double
Dim m As Double
Dim tx As String
If Combo5.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·ﬁ”„", vbCritical
Exit Sub
End If
grd3.Visible = False
grd3.Clear
grd3.Cols = 7
grd3.Rows = 1
grd3.ColWidth(0) = 1000
grd3.ColWidth(1) = 3500
grd3.ColWidth(2) = 3500
grd3.ColWidth(3) = 1500
grd3.ColWidth(4) = 1500
grd3.ColWidth(5) = 1300
grd3.ColWidth(6) = 1300
grd3.ColAlignment(0) = 1
grd3.ColAlignment(1) = 1
grd3.ColAlignment(2) = 6
grd3.ColAlignment(3) = 3
grd3.ColAlignment(4) = 3
grd3.ColAlignment(5) = 3
grd3.ColAlignment(6) = 3
grd3.row = 0
grd3.Col = 0
grd3.Text = "Code"
grd3.Col = 1
grd3.Text = "Nom"
grd3.Col = 2
grd3.Text = "«·«”„"
grd3.Col = 3
grd3.Text = "NNI"
grd3.Col = 4
grd3.Text = "RIM"
grd3.Col = 5
grd3.Text = "LDN"
grd3.Col = 6
grd3.Text = "ADN"
i = 1
Call cont2
grd3.Rows = nn.RecordCount + 10
Do While Not nn.EOF
If nn!cla = Combo5.Text Then
grd3.row = i
grd3.Col = 0
grd3.Text = Val(nn!ser)
grd3.Col = 1
grd3.Text = nn!nof
grd3.Col = 2
grd3.Text = nn!nom
grd3.Col = 3
grd3.Text = nn!nni
grd3.Col = 4
grd3.Text = nn!rim
grd3.Col = 5
grd3.Text = nn!liu
grd3.Col = 6
grd3.Text = nn!dat
i = i + 1
End If
nn.MoveNext
Loop
grd3.Rows = i
grd3.Visible = True
End Sub

Private Sub DT1_Change()
On Error Resume Next
DT2.Value = DT1.Value + 364

End Sub

Private Sub DT1_Click()
On Error Resume Next
DT1_Change
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Top = 100
Skin1.LoadSkin App.Path & "\green.skn"
Skin1.ApplySkin Me.hWnd
Call chargegrd9
Call chargec1
Call chargegrd1
Call chargegrd6
Call chargec3
Call chargec4
Call chargec5
Combo8.Clear
DT11.Value = Date
Combo4.Visible = False
Label31(10).Visible = False
Combo3.Text = "œ›⁄ „»·€"
End Sub

Private Sub grd1_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim au As Double
i = grd1.row
j = grd1.Col
If i > 0 Then
grd1.ToolTipText = ""
If j = 1 Then
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
Check5.Value = 0
Check6.Value = 0
Check7.Value = 0
Check8.Value = 0
grd1.row = i
grd1.Col = 0
au = grd1.Text
Call cont
Do While Not ut.EOF
If ut!aut = au Then
Label6.Caption = ut!aut
Text4.Text = ut!uti
Text5.Text = ut!mot
Text6.Text = ut!mot
Check1.Value = ut!Dir
Check2.Value = ut!par
Check3.Value = ut!cla
Check4.Value = ut!etu
Check5.Value = ut!pro
Check6.Value = ut!cai
Check7.Value = ut!com
Check8.Value = ut!Arc
Exit Sub
End If
ut.MoveNext
Loop
Exit Sub
End If
grd1.ToolTipText = grd1.Text
End If
End Sub

Private Sub grd6_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
j = grd6.row
i = grd6.Col
If j > 0 Then
grd6.row = j
grd6.Col = 0
Label27.Caption = grd6.Text
If i = 4 Then
g = MsgBox("Â·  —Ìœ «·Õ–› Õﬁ« ", vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
Call cont
Do While Not fc.EOF
If Label27.Caption = fc!mtr Then
fc.Delete
ProgressBar11.Value = 0
Timer11.Enabled = True
Exit Sub
End If
fc.MoveNext
Loop
End If
Exit Sub
End If
'*** afficher
If i = 5 Then
grd6.row = j
grd6.Col = 0
Label33.Caption = grd6.Text
grd6.Col = 1
Label30.Caption = grd6.Text
grd12.Visible = False
grd13.Visible = False
Call chargegrd12_13
grd12.Visible = True
grd13.Visible = True
Exit Sub
End If
grd6.row = j
grd6.Col = 1
Text7.Text = grd6.Text
grd6.Col = 2
Text8.Text = grd6.Text
grd6.Col = 3
Text9.Text = grd6.Text
End If
End Sub

Private Sub grd9_Click()
On Error Resume Next
Dim i As Double
Dim a As Double
Dim b As Double
Dim c As Double
i = grd9.row
grd9.row = i
grd9.Col = 0
Label5.Caption = grd9.Text
grd9.Col = 1
Text2.Text = grd9.Text
grd9.Col = 2
a = grd9.Text
grd9.Col = 3
Text3.Text = grd9.Text
b = Label9.Caption
c = b - a
Label9.Caption = c
Call calculepourcentage
End Sub




Private Sub Text_Change(Index As Integer)
On Error Resume Next
Dim i As Double
Dim a As Double
Dim b As Double
Dim c As Double
i = Index
Text(i).Text = Trim(Text(i).Text)
a = 0
b = 0
c = 0
'0
If i = 0 Then
If Text(0).Text <> "" Then
a = Text(0).Text
If a > 100 Then
a = 100
Text(0).Text = a
End If
If a < 0 Then
a = 0
Text(0).Text = a
End If
'Text(1).Text = ""
c = 100 - a
Text(1).Text = c
Exit Sub
Else
Text(1).Text = ""
Exit Sub
End If
End If

'1
If i = 1 Then
If Text(1).Text <> "" Then
b = Text(1).Text
If b > 100 Then
b = 100
Text(1).Text = b
End If
If b < 0 Then
b = 0
Text(1).Text = b
End If
'Text(0).Text = ""
c = 100 - b
Text(0).Text = c
Exit Sub
Else
Text(0).Text = ""
Exit Sub
End If
End If
End Sub
Private Sub chargegrd9()
On Error Resume Next
Dim i As Double
Dim p As Double
Dim sp As Double
grd9.Clear
grd9.Cols = 4
grd9.Rows = 1
grd9.ColWidth(0) = 0
grd9.ColWidth(1) = 2400
grd9.ColWidth(2) = 600
grd9.ColWidth(3) = 1100
grd9.ColAlignment(0) = 1
grd9.ColAlignment(1) = 1
grd9.ColAlignment(2) = 1
grd9.ColAlignment(3) = 1
grd9.row = 0
grd9.Col = 1
grd9.Text = "«·‘—Ìﬂ"
grd9.Col = 2
grd9.Text = "%"
grd9.Col = 3
grd9.Text = "«·Â« ›"
i = 1
sp = 0
Call cont
grd9.Rows = pa.RecordCount + 3
Do While Not pa.EOF
grd9.row = i
grd9.Col = 0
grd9.Text = pa!aut
grd9.Col = 1
grd9.Text = pa!nom
grd9.Col = 2
grd9.Text = pa!pou
p = pa!pou
sp = sp + p
grd9.Col = 3
grd9.Text = pa!tel
i = i + 1
pa.MoveNext
Loop
Label9.Caption = sp
grd9.Rows = i
Call calculepourcentage
End Sub
Private Sub chargegrd1()
On Error Resume Next
Dim i As Double
Dim p As Double
Dim sm As String
Dim m1 As String
grd1.Clear
grd1.Cols = 4
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 2000
grd1.ColWidth(2) = 1000
grd1.ColWidth(3) = 1100
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.row = 0
grd1.Col = 1
grd1.Text = "«·„” Œœ„"
grd1.Col = 2
grd1.Text = "ﬂ·„… «·”—"
grd1.Col = 3
grd1.Text = "«·’·«ÕÌ« "
i = 1
Call cont
grd1.Rows = ut.RecordCount + 3
Do While Not ut.EOF
grd1.row = i
grd1.Col = 0
grd1.Text = ut!aut
grd1.Col = 1
grd1.Text = ut!uti
grd1.Col = 2
grd1.Text = ut!mot
grd1.CellBackColor = &HFFFFFF
sm = ""
If ut!Dir = 1 Then
m1 = "«·„ƒ””…"
sm = sm + "," + m1
End If
If ut!par = 1 Then
m1 = "«·‘—ﬂ«¡"
sm = sm + "," + m1
End If
If ut!cla = 1 Then
m1 = "«·√ﬁ”«„"
sm = sm + "," + m1
End If
If ut!etu = 1 Then
m1 = "«· ·«„Ì–"
sm = sm + "," + m1
End If
If ut!pro = 1 Then
m1 = "«·√”« –…"
sm = sm + "," + m1
End If
If ut!cai = 1 Then
m1 = "«·’‰œÊﬁ"
sm = sm + "," + m1
End If
If ut!com = 1 Then
m1 = "«·„Õ«”»…"
sm = sm + "," + m1
End If
If ut!Arc = 1 Then
m1 = "«·«—‘Ì›"
sm = sm + "," + m1
End If
grd1.Col = 3
grd1.Text = sm
grd1.CellBackColor = &HFFFFFF
i = i + 1
ut.MoveNext
Loop
grd1.Rows = i
End Sub
Private Sub chargegrd6()
On Error Resume Next
'On Error Resume Next
Dim i As Double
Dim p As Double
grd6.Clear
grd6.Cols = 6
grd6.Rows = 1
grd6.ColWidth(0) = 1100
grd6.ColWidth(1) = 1300
grd6.ColWidth(2) = 1000
grd6.ColWidth(3) = 1000
grd6.ColWidth(4) = 500
grd6.ColWidth(5) = 700
grd6.ColAlignment(0) = 1
grd6.ColAlignment(1) = 1
grd6.ColAlignment(2) = 1
grd6.ColAlignment(3) = 1
grd6.ColAlignment(4) = 1
grd6.ColAlignment(5) = 1
grd6.row = 0
grd6.Col = 0
grd6.Text = "—ﬁ„ «·„ÊŸ›"
grd6.Col = 1
grd6.Text = "«”„ «·„ÊŸ›"
grd6.Col = 2
grd6.Text = "—ﬁ„ «·Â« ›"
grd6.Col = 3
grd6.Text = "«·ÊŸÌ›…"
i = 1
p = 0
Call cont
grd6.Rows = fc.RecordCount + 3
Do While Not fc.EOF
grd6.row = i
grd6.Col = 0
grd6.Text = fc!mtr
grd6.Col = 1
grd6.Text = fc!nom
grd6.Col = 2
grd6.Text = fc!tel
grd6.Col = 3
grd6.Text = fc!foc
grd6.Col = 4
grd6.Text = "Õ–›"
grd6.Col = 5
grd6.Text = "⁄—÷"
i = i + 1
fc.MoveNext
Loop
grd6.Rows = i
End Sub
Private Sub chargegrd12_13()
On Error Resume Next
'On Error Resume Next
Dim i As Double
Dim j As Double
Dim p As Double
Dim d As Double
Dim m As Double
grd12.Clear
grd12.Cols = 3
grd12.Rows = 1
grd12.ColWidth(0) = 0
grd12.ColWidth(1) = 1300
grd12.ColWidth(2) = 1500
grd12.ColAlignment(0) = 1
grd12.ColAlignment(1) = 1
grd12.ColAlignment(2) = 1
grd12.row = 0
grd12.Col = 1
grd12.Text = "«· «—ÌŒ"
grd12.Col = 2
grd12.Text = "«·„»·€"
grd13.Clear
grd13.Cols = 4
grd13.Rows = 1
grd13.ColWidth(0) = 0
grd13.ColWidth(1) = 1300
grd13.ColWidth(2) = 700
grd13.ColWidth(3) = 1500
grd13.ColAlignment(0) = 1
grd13.ColAlignment(1) = 1
grd13.ColAlignment(2) = 1
grd13.ColAlignment(3) = 1
grd13.row = 0
grd13.Col = 1
grd13.Text = "«· «—ÌŒ"
grd13.Col = 2
grd13.Text = "«·‘Â—"
grd13.Col = 3
grd13.Text = "«·„»·€"
i = 1
j = 1
p = 0
d = 0
m = 0
Call cont
grd12.Rows = pfc.RecordCount + 3
grd13.Rows = pfc.RecordCount + 3
Do While Not pfc.EOF
m = 0
If Label33.Caption = pfc!mtr Then
If pfc!cas = "œ›⁄ „»·€" Then
grd12.row = i
grd12.Col = 0
grd12.Text = pfc!aut
grd12.Col = 1
grd12.Text = pfc!dat
grd12.Col = 2
grd12.Text = pfc!mon
m = pfc!mon
d = d + m
i = i + 1
Else
grd13.row = j
grd13.Col = 0
grd13.Text = pfc!aut
grd13.Col = 1
grd13.Text = pfc!dat
grd13.Col = 2
grd13.Text = pfc!moi
grd13.Col = 3
grd13.Text = pfc!mon
m = pfc!mon
p = p + m
j = j + 1
End If
End If
pfc.MoveNext
Loop
grd12.Rows = i
grd13.Rows = j
Label36.Caption = d
Label38.Caption = p
m = (p - d)
Label40.Caption = m
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
Dim i As Integer
Dim j As Integer
Dim n As Double
Dim k As Double
Dim vg As String
k = Index
Text(k).Text = Trim(Text(k).Text)
n = Len(Text(k).Text)
j = 0
If n = 0 And KeyAscii = 46 Then
KeyAscii = 0
End If
If KeyAscii <> 8 Then
If KeyAscii < 46 Or KeyAscii > 57 Or KeyAscii = 47 Then
KeyAscii = 0
End If
End If
For i = 1 To n
vg = Mid$(Text(k).Text, i, 1)
r = Asc(vg)
If r = 46 Then
j = i + 2
End If
If j > 2 And KeyAscii = 46 Then
KeyAscii = 0
End If
If i = j And KeyAscii <> 8 Then
KeyAscii = 0
End If
If i = j Then
i = n
End If
Next i

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii <> 8 Then
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End If

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
ProgressBar1.Value = ProgressBar1.Value + 8
If ProgressBar1.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
Text4.Text = ""
Text4.SetFocus
Label6.Caption = ""
Text5.Text = ""
Text6.Text = ""
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
Check5.Value = 0
Check6.Value = 0
Check7.Value = 0
Check8.Value = 0
grd1.Visible = False
grd1.Clear
grd1.Rows = 1
Call chargegrd1
grd1.Visible = True
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = False
End If

End Sub

Private Sub Timer11_Timer()
On Error Resume Next
ProgressBar11.Value = ProgressBar11.Value + 8
If ProgressBar11.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text7.SetFocus
Label27.Caption = ""
grd6.Visible = False
grd6.Clear
grd6.Rows = 1
Call chargegrd6
grd6.Visible = True
ProgressBar11.Value = 0
Timer11.Enabled = False
End If

End Sub

Private Sub Timer12_Timer()
On Error Resume Next
ProgressBar11.Value = ProgressBar11.Value + 8
If ProgressBar11.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
Text10.Text = ""
Text10.SetFocus
grd12.Visible = False
grd12.Clear
grd12.Rows = 1
grd13.Visible = False
grd13.Clear
grd13.Rows = 1
Call chargegrd12_13
grd12.Visible = True
grd13.Visible = True
Call chargec4
Combo8.Clear
Combo4.Visible = False
Label31(10).Visible = False
Combo3.Text = "œ›⁄ „»·€"
ProgressBar11.Value = 0
Timer12.Enabled = False
End If

End Sub

Private Sub Timer4_Timer()
On Error Resume Next
ProgressBar4.Value = ProgressBar4.Value + 8
If ProgressBar4.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
Text2.Text = ""
Text2.SetFocus
Label5.Caption = ""
Text3.Text = ""
grd9.Visible = False
grd9.Clear
grd9.Rows = 1
Call chargegrd9
grd9.Visible = True
ProgressBar4.Value = 0
ProgressBar4.Visible = True
Timer4.Enabled = False
End If

End Sub
Public Sub chargec5()
On Error Resume Next
Call cont
Combo5.Clear
  Do While Not cl.EOF
  If cl!act = "1" Then
    Combo5.AddItem cl!cla
   End If
cl.MoveNext
  Loop
End Sub
Public Sub chargec3()
On Error Resume Next
Combo3.Clear
Combo3.AddItem "œ›⁄ „»·€"
Combo3.AddItem " ”œÌœ —« »"
End Sub
Public Sub chargec4()
On Error Resume Next
Combo4.Clear
Combo4.AddItem "10"
Combo4.AddItem "11"
Combo4.AddItem "12"
Combo4.AddItem "1"
Combo4.AddItem "2"
Combo4.AddItem "3"
Combo4.AddItem "4"
Combo4.AddItem "5"
Combo4.AddItem "6"
Combo4.AddItem "7"
Combo4.AddItem "8"
Combo4.AddItem "9"
End Sub

Private Sub chargec8()
On Error Resume Next
Dim i As Double
Dim k As Integer
Call cont
k = sr!cdc
Check2.Value = k
If cd.RecordCount > 0 Then
cd.MoveFirst
End If
Combo8.Clear
Do While Not cd.EOF
'**** if salaire then depense
If Combo3.Text = " ”œÌœ —« »" Then
'**** if codes
If k = 1 Then
If cd!cas = "«·„’—Ê›« " Then
Combo8.AddItem cd!cod
End If
End If
'**** if dec
If k = 0 Then
If cd!cas = "«·„’—Ê›« " Then
Combo8.AddItem cd!dec
End If
End If
'**** if pay then caisse
Else
'**** if codes
If k = 1 Then
If cd!cas = "Õ”«» «·⁄„«·" Then
Combo8.AddItem cd!cod
End If
End If
'**** if dec
If k = 0 Then
If cd!cas = "Õ”«» «·⁄„«·" Then
Combo8.AddItem cd!dec
End If
End If
End If
'*****
cd.MoveNext
Loop
End Sub
Private Sub paymon_fonctionnaires()
On Error Resume Next
'**** if pay montants
'****** caisse
Call cont
ca.AddNew
au = ca!aun
If Val(Combo8.Text) > 0 Then
ca!cod = Combo8.Text
ca!dec = Label42.Caption
Else
ca!dec = Combo8.Text
ca!cod = Label42.Caption
End If
ca!mem = "œ›⁄ „»·€ " + Text10.Text + " ·’«·Õ «·„ÊŸ› " + Label30.Caption
ca!mon = Text10.Text
ca!cas = "Œ«—Ã"
ca!heu = Time$
ca!dat = DT11.Value
ca!ger = face.SBB1.Panels(11).Text
ca!aut = au
ca!com = Label43.Caption
ca.Update
'****** journal
c = sr!ord
jr.AddNew
jr!cre = Text10.Text
jr!deb = "0"
jr!dem = "„‰ Õ‹"
jr!com = "«·⁄„«·"
jr!dec = "œ›⁄ „»·€ " + Text10.Text + " ·’«·Õ «·„ÊŸ› " + Label30.Caption
jr!ord = c
jr!dat = DT11.Value
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
jr.AddNew
jr!cre = "0"
jr!deb = Text10.Text
jr!dem = "≈·Ï Õ‹"
jr!com = "«·’‰œÊﬁ"
jr!dec = "œ›⁄ „»·€ " + Text10.Text + " ·’«·Õ «·„ÊŸ› " + Label30.Caption
jr!ord = c
jr!dat = DT11.Value
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
'****** cascaisse
a = sr!cca
b = Text10.Text
a = a - b
sr!cca = a
sr!ord = c + 1
sr.Update
'**** end pay salaire
'**** if pay montant
pfc.AddNew
pfc!mtr = Label33.Caption
pfc!aut = au
pfc!nom = Label30.Caption
pfc!dat = DT11.Value
pfc!heu = Time$
pfc!mon = Text10.Text
pfc!cas = Combo3.Text
pfc!moi = ""
pfc!ger = face.SBB1.Panels(11).Text
pfc.Update
ProgressBar11.Value = 0
Timer12.Enabled = True

End Sub
Private Sub paysal_fonctionnaires()
On Error Resume Next
'**** if pay montants
'****** caisse
Call cont
ca.AddNew
au = ca!aun
If Val(Combo8.Text) > 0 Then
ca!cod = Combo8.Text
ca!dec = Label42.Caption
Else
ca!dec = Combo8.Text
ca!cod = Label42.Caption
End If
ca!mem = " ”œÌœ —« » " + Text10.Text + " ··„ÊŸ› " + Label30.Caption + " ·‘Â— " + Combo4.Text
ca!mon = "0"
ca!cas = "Œ«—Ã"
ca!heu = Time$
ca!dat = DT11.Value
ca!ger = face.SBB1.Panels(11).Text
ca!aut = au
ca!com = Label43.Caption
ca.Update
'****** journal
c = sr!ord
jr.AddNew
jr!cre = Text10.Text
jr!deb = "0"
jr!dem = "„‰ Õ‹"
jr!com = "«·„’—Ê›« "
jr!dec = " ”œÌœ —« » " + Text10.Text + " ··„ÊŸ› " + Label30.Caption + " ·‘Â— " + Combo4.Text
jr!ord = c
jr!dat = DT11.Value
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
jr.AddNew
jr!cre = "0"
jr!deb = Text10.Text
jr!dem = "≈·Ï Õ‹"
jr!com = "«·⁄„«·"
jr!dec = " ”œÌœ —« » " + Text10.Text + " ··„ÊŸ› " + Label30.Caption + " ·‘Â— " + Combo4.Text
jr!ord = c
jr!dat = DT11.Value
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
'****** cascaisse
sr!ord = c + 1
sr.Update
'**** end pay salaire
'**** Depenses
dp.AddNew
dp!aut = au
dp!dec = " ”œÌœ —« » «·„ÊŸ› " + Label30.Caption
dp!mon = Text10.Text
dp!dat = DT11.Value
dp!heu = Time$
dp!ger = face.SBB1.Panels(11).Text
If Val(Combo8.Text) > 0 Then
dp!com = Label42.Caption
Else
dp!com = Combo8.Text
End If
dp.Update
'**** if pay montant
pfc.AddNew
pfc!mtr = Label33.Caption
pfc!aut = au
pfc!nom = Label30.Caption
pfc!dat = DT11.Value
pfc!heu = Time$
pfc!mon = Text10.Text
pfc!cas = Combo3.Text
pfc!moi = Combo4.Text
pfc!ger = face.SBB1.Panels(11).Text
pfc.Update
ProgressBar11.Value = 0
Timer12.Enabled = True

End Sub

