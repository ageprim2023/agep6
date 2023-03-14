VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ActiveSkin.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form caisse 
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
      Height          =   9375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   16536
      _Version        =   393216
      Tabs            =   7
      Tab             =   6
      TabsPerRow      =   7
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
      TabCaption(0)   =   "«·√ﬂÊ«œ"
      TabPicture(0)   =   "caisse.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grd6"
      Tab(0).Control(1)=   "Picture1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Õ—ﬂ… «·’‰œÊﬁ"
      TabPicture(1)   =   "caisse.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture20"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Õ”«» «·»‰ﬂ"
      TabPicture(2)   =   "caisse.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture18"
      Tab(2).Control(1)=   "Picture19"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Õ”«» «·„’—Ê›« "
      TabPicture(3)   =   "caisse.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture15"
      Tab(3).Control(1)=   "Picture16"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Õ”«» «·‘—ﬂ«¡"
      TabPicture(4)   =   "caisse.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Picture17"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Õ”«» «·√”« –…"
      TabPicture(5)   =   "caisse.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Picture10"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Õ”«» «· ·«„Ì–"
      TabPicture(6)   =   "caisse.frx":00A8
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "Picture4"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      Begin VB.PictureBox Picture20 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8895
         Left            =   -74880
         ScaleHeight     =   8895
         ScaleWidth      =   14295
         TabIndex        =   233
         Top             =   360
         Width           =   14295
         Begin VB.ComboBox Combo14 
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
            ItemData        =   "caisse.frx":00C4
            Left            =   12600
            List            =   "caisse.frx":00DA
            Style           =   2  'Dropdown List
            TabIndex        =   293
            Top             =   120
            Width           =   1575
         End
         Begin VB.CommandButton Command33 
            Caption         =   "√Õœ«À „«»Ì‰ Â–Ì‰ «· «—ÌŒÌ‰"
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
            Left            =   5880
            TabIndex        =   262
            Top             =   120
            Width           =   2415
         End
         Begin VB.CommandButton Command34 
            Caption         =   "√Õœ«À «·ÌÊ„"
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
            Left            =   4680
            TabIndex        =   261
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton Command35 
            Caption         =   "√Õœ«À Â–« «·‘Â—"
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
            Left            =   3120
            TabIndex        =   260
            Top             =   120
            Width           =   1575
         End
         Begin VB.CommandButton Command36 
            Caption         =   " √Õœ«À «·”‰… «·œ—«”Ì…"
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
            Left            =   1200
            TabIndex        =   259
            Top             =   120
            Width           =   1935
         End
         Begin VB.CommandButton Command32 
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
            Left            =   120
            TabIndex        =   234
            Top             =   120
            Width           =   975
         End
         Begin MSFlexGridLib.MSFlexGrid grd23 
            Height          =   3135
            Left            =   120
            TabIndex        =   235
            Top             =   5640
            Width           =   14055
            _ExtentX        =   24791
            _ExtentY        =   5530
            _Version        =   393216
            Rows            =   8
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
         Begin MSComCtl2.DTPicker DT21 
            Height          =   375
            Left            =   10560
            TabIndex        =   236
            Top             =   120
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
            Format          =   125501441
            CurrentDate     =   41154
         End
         Begin MSComCtl2.DTPicker DT22 
            Height          =   375
            Left            =   8400
            TabIndex        =   237
            Top             =   120
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
            Format          =   125501441
            CurrentDate     =   41154
         End
         Begin MSFlexGridLib.MSFlexGrid grd22 
            Height          =   3975
            Left            =   120
            TabIndex        =   238
            Top             =   1320
            Width           =   14055
            _ExtentX        =   24791
            _ExtentY        =   7011
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
         Begin MSComctlLib.ProgressBar ProgressBar8 
            Height          =   375
            Left            =   120
            TabIndex        =   256
            Top             =   960
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label Label44 
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
            Left            =   3840
            TabIndex        =   270
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ’Ì·… «·ÌÊ„"
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
            Index           =   26
            Left            =   5040
            TabIndex        =   269
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label42 
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
            Left            =   11520
            TabIndex        =   255
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·—’Ìœ «·”«»ﬁ"
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
            Left            =   12000
            TabIndex        =   254
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "”Ã· ⁄„·Ì«  «·’‰œÊﬁ ›Ì «· «—ÌŒ √⁄·«Â Õ”» Õ”«»«  «·’‰œÊﬁ"
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
            Index           =   19
            Left            =   4440
            TabIndex        =   248
            Top             =   5280
            Width           =   5415
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "”Ã· ⁄„·Ì«  «·’‰œÊﬁ ›Ì «· «—ÌŒ √⁄·«Â"
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
            Index           =   18
            Left            =   4680
            TabIndex        =   247
            Top             =   960
            Width           =   4935
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "≈·Ï"
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
            Index           =   24
            Left            =   9120
            TabIndex        =   246
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„‰"
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
            Index           =   23
            Left            =   11160
            TabIndex        =   245
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·œ«Œ·"
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
            Index           =   22
            Left            =   10440
            TabIndex        =   244
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·Œ«—Ã"
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
            Index           =   21
            Left            =   8160
            TabIndex        =   243
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—’Ìœ «·’‰œÊﬁ «·Õ«·Ì"
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
            Index           =   20
            Left            =   1440
            TabIndex        =   242
            Top             =   600
            Width           =   2055
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
            Left            =   120
            TabIndex        =   241
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label35 
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
            Left            =   6720
            TabIndex        =   240
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label34 
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
            Left            =   9120
            TabIndex        =   239
            Top             =   600
            Width           =   1575
         End
      End
      Begin VB.PictureBox Picture19 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   -74880
         ScaleHeight     =   615
         ScaleWidth      =   14295
         TabIndex        =   210
         Top             =   360
         Width           =   14295
         Begin VB.ComboBox Combo11 
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
            ItemData        =   "caisse.frx":0124
            Left            =   11280
            List            =   "caisse.frx":013A
            Style           =   2  'Dropdown List
            TabIndex        =   214
            Top             =   120
            Width           =   2295
         End
         Begin VB.TextBox Text13 
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
            Left            =   6960
            TabIndex        =   213
            Top             =   120
            Width           =   3135
         End
         Begin VB.CommandButton Command31 
            Caption         =   "Õ›Ÿ «·⁄„·Ì…"
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
            Left            =   120
            TabIndex        =   212
            Top             =   120
            Width           =   1575
         End
         Begin VB.TextBox Text12 
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
            Left            =   1800
            TabIndex        =   211
            Top             =   120
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker DT18 
            Height          =   375
            Left            =   4920
            TabIndex        =   215
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   125829121
            CurrentDate     =   41154
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
            Index           =   56
            Left            =   12720
            TabIndex        =   219
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "»Ì«‰ «·⁄„·Ì…"
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
            Index           =   55
            Left            =   9120
            TabIndex        =   218
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label Label31 
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
            Index           =   54
            Left            =   4800
            TabIndex        =   217
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label Label31 
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
            Index           =   53
            Left            =   2640
            TabIndex        =   216
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.PictureBox Picture18 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8175
         Left            =   -74880
         ScaleHeight     =   8175
         ScaleWidth      =   14295
         TabIndex        =   201
         Top             =   1080
         Width           =   14295
         Begin VB.CommandButton Command30 
            Caption         =   "”Õ»"
            Enabled         =   0   'False
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
            Left            =   3960
            TabIndex        =   203
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton Command29 
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
            Left            =   5040
            TabIndex        =   202
            Top             =   600
            Width           =   1215
         End
         Begin MSFlexGridLib.MSFlexGrid grd20 
            Height          =   6375
            Left            =   120
            TabIndex        =   204
            Top             =   1560
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   11245
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
         Begin MSComCtl2.DTPicker DT19 
            Height          =   375
            Left            =   8160
            TabIndex        =   205
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   125501441
            CurrentDate     =   41154
         End
         Begin MSComCtl2.DTPicker DT20 
            Height          =   375
            Left            =   6360
            TabIndex        =   206
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   125501441
            CurrentDate     =   41154
         End
         Begin MSComctlLib.ProgressBar ProgressBar7 
            Height          =   255
            Left            =   3960
            TabIndex        =   207
            Top             =   1080
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
         Begin MSFlexGridLib.MSFlexGrid grd21 
            Height          =   6375
            Left            =   7200
            TabIndex        =   221
            Top             =   1560
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   11245
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
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "”Ã· «·„»«·€ «·„”ÕÊ»… „‰ «·»‰ﬂ"
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
            Index           =   17
            Left            =   120
            TabIndex        =   229
            Top             =   1080
            Width           =   4215
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "”Ã· «·„»«·€ «·„Êœ⁄… ›Ì «·»‰ﬂ"
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
            Index           =   16
            Left            =   9960
            TabIndex        =   228
            Top             =   1200
            Width           =   4215
         End
         Begin VB.Label Label33 
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
            Left            =   9480
            TabIndex        =   227
            Top             =   120
            Width           =   2775
         End
         Begin VB.Label Label32 
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
            Left            =   4440
            TabIndex        =   226
            Top             =   120
            Width           =   2775
         End
         Begin VB.Label Label30 
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
            Left            =   120
            TabIndex        =   225
            Top             =   120
            Width           =   2775
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·—’Ìœ ›Ì «·»‰ﬂ"
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
            Index           =   15
            Left            =   2280
            TabIndex        =   224
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„Ã„Ê⁄ «·„»«·€ «·„”ÕÊ»…"
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
            Index           =   14
            Left            =   6840
            TabIndex        =   223
            Top             =   120
            Width           =   2415
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„Ã„Ê⁄ «·„»«·€ «·„Êœ⁄…"
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
            Index           =   13
            Left            =   12000
            TabIndex        =   222
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„‰  «—ÌŒ"
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
            Index           =   12
            Left            =   9000
            TabIndex        =   209
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "≈·Ï"
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
            Index           =   11
            Left            =   6720
            TabIndex        =   208
            Top             =   600
            Width           =   1335
         End
      End
      Begin VB.PictureBox Picture17 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8895
         Left            =   -74880
         ScaleHeight     =   8895
         ScaleWidth      =   14295
         TabIndex        =   162
         Top             =   360
         Width           =   14295
         Begin VB.CommandButton Command28 
            Caption         =   "”Õ»"
            Enabled         =   0   'False
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
            Left            =   2040
            TabIndex        =   191
            Top             =   2760
            Width           =   1095
         End
         Begin VB.CommandButton Command27 
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
            Left            =   3240
            TabIndex        =   190
            Top             =   2760
            Width           =   1215
         End
         Begin VB.CommandButton Command26 
            Caption         =   " ÕœÌÀ"
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
            TabIndex        =   187
            Top             =   2160
            Width           =   1695
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Õ›Ÿ «·⁄„·Ì…"
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
            Left            =   2040
            TabIndex        =   179
            Top             =   2160
            Width           =   2535
         End
         Begin VB.TextBox Text11 
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
            Left            =   6840
            TabIndex        =   173
            Top             =   1680
            Width           =   1455
         End
         Begin VB.ComboBox Combo10 
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
            ItemData        =   "caisse.frx":0184
            Left            =   4680
            List            =   "caisse.frx":019A
            Style           =   2  'Dropdown List
            TabIndex        =   172
            Top             =   1680
            Width           =   1455
         End
         Begin VB.ComboBox Combo9 
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
            ItemData        =   "caisse.frx":01E4
            Left            =   4680
            List            =   "caisse.frx":01FA
            Style           =   2  'Dropdown List
            TabIndex        =   169
            Top             =   1200
            Width           =   3615
         End
         Begin MSFlexGridLib.MSFlexGrid grd16 
            Height          =   8295
            Left            =   9240
            TabIndex        =   163
            Top             =   480
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   14631
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
         Begin MSComCtl2.DTPicker DT12 
            Height          =   375
            Left            =   240
            TabIndex        =   177
            Top             =   1680
            Width           =   1695
            _ExtentX        =   2990
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
            Format          =   125829121
            CurrentDate     =   41154
         End
         Begin MSComctlLib.ProgressBar ProgressBar4 
            Height          =   375
            Left            =   4680
            TabIndex        =   180
            Top             =   2160
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
         End
         Begin MSFlexGridLib.MSFlexGrid grd17 
            Height          =   5175
            Left            =   120
            TabIndex        =   188
            Top             =   3600
            Width           =   4335
            _ExtentX        =   7646
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
         Begin MSFlexGridLib.MSFlexGrid grd18 
            Height          =   5175
            Left            =   4680
            TabIndex        =   189
            Top             =   3600
            Width           =   4335
            _ExtentX        =   7646
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
         Begin MSComCtl2.DTPicker DT13 
            Height          =   375
            Left            =   6840
            TabIndex        =   192
            Top             =   2760
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   125829121
            CurrentDate     =   41154
         End
         Begin MSComCtl2.DTPicker DT14 
            Height          =   375
            Left            =   4560
            TabIndex        =   193
            Top             =   2760
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   125829121
            CurrentDate     =   41154
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "”Ã· «” ·«„«  «·‘—Ìﬂ „‰ «·’‰œÊﬁ"
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
            Left            =   120
            TabIndex        =   199
            Top             =   3240
            Width           =   4335
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "”Ã· „œ›Ê⁄«  «·‘—Ìﬂ ··’‰œÊﬁ"
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
            Left            =   4680
            TabIndex        =   198
            Top             =   3240
            Width           =   4335
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„‰  «—ÌŒ"
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
            Left            =   7680
            TabIndex        =   195
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "≈·Ï  «—ÌŒ"
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
            Left            =   5400
            TabIndex        =   194
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "œ«∆‰ »‹ "
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
            Index           =   52
            Left            =   720
            TabIndex        =   186
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label26 
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
            TabIndex        =   185
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label25 
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
            TabIndex        =   184
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label24 
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
            Left            =   4800
            TabIndex        =   183
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„Ã„Ê⁄ «·„»«·€ «· Ì «” ·„"
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
            Index           =   51
            Left            =   2280
            TabIndex        =   182
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„Ã„Ê⁄ «·„»«·€ «· Ì œ›⁄"
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
            Index           =   50
            Left            =   6720
            TabIndex        =   181
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label31 
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
            Index           =   49
            Left            =   600
            TabIndex        =   178
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label Label31 
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
            Index           =   48
            Left            =   2760
            TabIndex        =   176
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„” ·„"
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
            Index           =   47
            Left            =   3120
            TabIndex        =   175
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·œ«›⁄"
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
            Index           =   46
            Left            =   5280
            TabIndex        =   174
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   3
            X1              =   120
            X2              =   9000
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   3
            Height          =   2535
            Left            =   120
            Top             =   120
            Width           =   8895
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "√”„«¡ «·‘—ﬂ«¡"
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
            Index           =   45
            Left            =   9840
            TabIndex        =   171
            Top             =   120
            Width           =   3735
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
            Index           =   44
            Left            =   7440
            TabIndex        =   170
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label31 
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
            Index           =   43
            Left            =   6840
            TabIndex        =   168
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label Label91 
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
            Left            =   3480
            TabIndex        =   167
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label92 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·≈”„"
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
            Left            =   6960
            TabIndex        =   166
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label95 
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
            Left            =   2520
            TabIndex        =   165
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label96 
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
            Left            =   4680
            TabIndex        =   164
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.PictureBox Picture16 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8175
         Left            =   -74880
         ScaleHeight     =   8175
         ScaleWidth      =   14295
         TabIndex        =   154
         Top             =   1080
         Width           =   14295
         Begin VB.CommandButton Command20 
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
            Left            =   5160
            TabIndex        =   156
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton Command19 
            Caption         =   "”Õ»"
            Enabled         =   0   'False
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
            Left            =   4080
            TabIndex        =   155
            Top             =   600
            Width           =   975
         End
         Begin MSFlexGridLib.MSFlexGrid grd19 
            Height          =   6975
            Left            =   120
            TabIndex        =   157
            Top             =   1080
            Width           =   14055
            _ExtentX        =   24791
            _ExtentY        =   12303
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
         Begin MSComCtl2.DTPicker DT16 
            Height          =   375
            Left            =   8040
            TabIndex        =   158
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   125829121
            CurrentDate     =   41154
         End
         Begin MSComCtl2.DTPicker DT17 
            Height          =   375
            Left            =   6240
            TabIndex        =   159
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   125829121
            CurrentDate     =   41154
         End
         Begin MSComctlLib.ProgressBar ProgressBar6 
            Height          =   375
            Left            =   4080
            TabIndex        =   200
            Top             =   120
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label Label43 
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
            Left            =   120
            TabIndex        =   266
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„Ã„Ê⁄ «·„’—Ê›« "
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
            Index           =   57
            Left            =   1800
            TabIndex        =   265
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "≈·Ï"
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
            Left            =   6600
            TabIndex        =   161
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„‰  «—ÌŒ"
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
            Left            =   8880
            TabIndex        =   160
            Top             =   600
            Width           =   1335
         End
      End
      Begin VB.PictureBox Picture15 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   -74880
         ScaleHeight     =   615
         ScaleWidth      =   14295
         TabIndex        =   144
         Top             =   360
         Width           =   14295
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
            Left            =   1800
            TabIndex        =   148
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton Command18 
            Caption         =   "Õ›Ÿ «·⁄„·Ì…"
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
            Left            =   120
            TabIndex        =   147
            Top             =   120
            Width           =   1575
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
            Left            =   6960
            TabIndex        =   146
            Top             =   120
            Width           =   3135
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
            ItemData        =   "caisse.frx":0244
            Left            =   11280
            List            =   "caisse.frx":025A
            Style           =   2  'Dropdown List
            TabIndex        =   145
            Top             =   120
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker DT15 
            Height          =   375
            Left            =   4920
            TabIndex        =   149
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   125829121
            CurrentDate     =   41154
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„»·€ «·„’—Ê›"
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
            Index           =   41
            Left            =   2640
            TabIndex        =   153
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label Label31 
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
            Index           =   40
            Left            =   4800
            TabIndex        =   152
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "»Ì«‰ «·⁄„·Ì…"
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
            Index           =   39
            Left            =   9120
            TabIndex        =   151
            Top             =   120
            Width           =   2055
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
            Left            =   12720
            TabIndex        =   150
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8895
         Left            =   120
         ScaleHeight     =   8895
         ScaleWidth      =   14295
         TabIndex        =   65
         Top             =   360
         Width           =   14295
         Begin TabDlg.SSTab SSTab2 
            Height          =   8655
            Left            =   120
            TabIndex        =   66
            Top             =   120
            Width           =   14055
            _ExtentX        =   24791
            _ExtentY        =   15266
            _Version        =   393216
            Tabs            =   2
            Tab             =   1
            TabsPerRow      =   2
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
            TabCaption(0)   =   "—ﬂ‰ «·„ «»⁄…"
            TabPicture(0)   =   "caisse.frx":02A4
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Picture7"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "—ﬂ‰ «·œ›⁄"
            TabPicture(1)   =   "caisse.frx":02C0
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "Picture5"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "Picture3"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).ControlCount=   2
            Begin VB.PictureBox Picture3 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   615
               Left            =   120
               ScaleHeight     =   615
               ScaleWidth      =   13815
               TabIndex        =   139
               Top             =   360
               Width           =   13815
               Begin VB.TextBox Text17 
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
                  Left            =   11160
                  TabIndex        =   285
                  Top             =   120
                  Width           =   1815
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
                  Left            =   7920
                  TabIndex        =   141
                  Top             =   120
                  Width           =   1455
               End
               Begin VB.CommandButton Command5 
                  Caption         =   "⁄—÷ «·Õ”«»"
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
                  Left            =   120
                  TabIndex        =   140
                  Top             =   120
                  Width           =   1455
               End
               Begin MSFlexGridLib.MSFlexGrid grd1 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   142
                  Top             =   120
                  Width           =   7695
                  _ExtentX        =   13573
                  _ExtentY        =   661
                  _Version        =   393216
                  Rows            =   1
                  Cols            =   4
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
               Begin VB.Label Label31 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Â« ›                                     √Ê «·—ﬁ„ «· ”·”·Ì"
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
                  Left            =   9360
                  TabIndex        =   143
                  Top             =   120
                  Width           =   4335
               End
            End
            Begin VB.PictureBox Picture5 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   7455
               Left            =   120
               ScaleHeight     =   7455
               ScaleWidth      =   13815
               TabIndex        =   97
               Top             =   1080
               Width           =   13815
               Begin VB.PictureBox Picture6 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   7455
                  Left            =   0
                  ScaleHeight     =   7455
                  ScaleWidth      =   13815
                  TabIndex        =   98
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   13815
                  Begin VB.CommandButton Command45 
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
                     Left            =   120
                     TabIndex        =   295
                     Top             =   600
                     Width           =   855
                  End
                  Begin VB.TextBox Text19 
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
                     Left            =   1080
                     TabIndex        =   292
                     Top             =   600
                     Width           =   2775
                  End
                  Begin VB.CommandButton Command44 
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
                     Left            =   5520
                     TabIndex        =   290
                     Top             =   600
                     Width           =   735
                  End
                  Begin VB.CheckBox Check6 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00000000&
                     Caption         =   "«· ﬁœÌ—"
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
                     Height          =   285
                     Left            =   11760
                     TabIndex        =   257
                     Top             =   6840
                     Value           =   1  'Checked
                     Width           =   255
                  End
                  Begin VB.PictureBox Picture2 
                     Height          =   6015
                     Left            =   120
                     ScaleHeight     =   5955
                     ScaleWidth      =   13515
                     TabIndex        =   109
                     Top             =   1080
                     Visible         =   0   'False
                     Width           =   13575
                     Begin VB.ComboBox Combo7 
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
                        ItemData        =   "caisse.frx":02DC
                        Left            =   1320
                        List            =   "caisse.frx":02F2
                        Style           =   2  'Dropdown List
                        TabIndex        =   287
                        Top             =   240
                        Width           =   1935
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
                        ItemData        =   "caisse.frx":033C
                        Left            =   1560
                        List            =   "caisse.frx":0352
                        Style           =   2  'Dropdown List
                        TabIndex        =   284
                        Top             =   2040
                        Width           =   1935
                     End
                     Begin VB.PictureBox Picture21 
                        Height          =   1455
                        Left            =   2040
                        ScaleHeight     =   1395
                        ScaleWidth      =   1515
                        TabIndex        =   277
                        Top             =   3240
                        Width           =   1575
                        Begin VB.CommandButton Command40 
                           Caption         =   "«· «·Ì"
                           Height          =   375
                           Left            =   0
                           TabIndex        =   280
                           Top             =   480
                           Width           =   615
                        End
                        Begin VB.CommandButton Command41 
                           Caption         =   "Õ›Ÿ"
                           Height          =   375
                           Left            =   960
                           TabIndex        =   279
                           Top             =   0
                           Width           =   495
                        End
                        Begin VB.CommandButton Command42 
                           Caption         =   "Ê’· Ê’·"
                           Height          =   375
                           Left            =   480
                           TabIndex        =   278
                           Top             =   960
                           Width           =   975
                        End
                        Begin VB.Label Label46 
                           Caption         =   "Label46"
                           Height          =   255
                           Left            =   0
                           TabIndex        =   282
                           Top             =   240
                           Width           =   735
                        End
                        Begin VB.Label Label47 
                           Caption         =   "Label47"
                           Height          =   255
                           Left            =   0
                           TabIndex        =   281
                           Top             =   0
                           Width           =   855
                        End
                     End
                     Begin VB.TextBox Text15 
                        Height          =   285
                        Left            =   2280
                        TabIndex        =   273
                        Text            =   "Text15"
                        Top             =   4800
                        Width           =   3495
                     End
                     Begin VB.CommandButton Command39 
                        Caption         =   "Command39"
                        Height          =   255
                        Left            =   2280
                        TabIndex        =   272
                        Top             =   4560
                        Width           =   1215
                     End
                     Begin VB.TextBox Text16 
                        Height          =   285
                        Left            =   2280
                        TabIndex        =   271
                        Text            =   "Text16"
                        Top             =   5040
                        Width           =   3495
                     End
                     Begin VB.CommandButton Command38 
                        Caption         =   "Envoer"
                        Height          =   375
                        Left            =   1200
                        TabIndex        =   268
                        Top             =   2520
                        Width           =   1095
                     End
                     Begin VB.TextBox Text14 
                        Height          =   375
                        Left            =   2400
                        TabIndex        =   267
                        Text            =   "195"
                        Top             =   2520
                        Width           =   855
                     End
                     Begin VB.CommandButton Command37 
                        Caption         =   "Command37"
                        Height          =   375
                        Left            =   1920
                        TabIndex        =   264
                        Top             =   1440
                        Width           =   1335
                     End
                     Begin MSComctlLib.ProgressBar ProgressBar9 
                        Height          =   255
                        Left            =   120
                        TabIndex        =   263
                        Top             =   5400
                        Width           =   3135
                        _ExtentX        =   5530
                        _ExtentY        =   450
                        _Version        =   393216
                        Appearance      =   1
                     End
                     Begin VB.Timer Timer9 
                        Enabled         =   0   'False
                        Interval        =   100
                        Left            =   840
                        Top             =   120
                     End
                     Begin VB.Timer Timer8 
                        Enabled         =   0   'False
                        Interval        =   50
                        Left            =   1200
                        Top             =   4920
                     End
                     Begin VB.ComboBox Combo12 
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
                        ItemData        =   "caisse.frx":039C
                        Left            =   1920
                        List            =   "caisse.frx":03B2
                        Style           =   2  'Dropdown List
                        TabIndex        =   232
                        Top             =   960
                        Width           =   1095
                     End
                     Begin VB.Timer Timer7 
                        Enabled         =   0   'False
                        Interval        =   50
                        Left            =   1200
                        Top             =   4440
                     End
                     Begin VB.Timer Timer6 
                        Enabled         =   0   'False
                        Interval        =   50
                        Left            =   1200
                        Top             =   3960
                     End
                     Begin VB.Timer Timer5 
                        Enabled         =   0   'False
                        Interval        =   50
                        Left            =   1200
                        Top             =   3480
                     End
                     Begin VB.Timer Timer2 
                        Enabled         =   0   'False
                        Interval        =   50
                        Left            =   120
                        Top             =   1800
                     End
                     Begin VB.Timer Timer1 
                        Enabled         =   0   'False
                        Interval        =   50
                        Left            =   120
                        Top             =   120
                     End
                     Begin VB.CommandButton Command4 
                        Caption         =   "Command4"
                        Height          =   375
                        Left            =   240
                        TabIndex        =   112
                        Top             =   600
                        Width           =   1575
                     End
                     Begin VB.Timer Timer3 
                        Enabled         =   0   'False
                        Interval        =   50
                        Left            =   120
                        Top             =   2760
                     End
                     Begin VB.Timer Timer4 
                        Enabled         =   0   'False
                        Interval        =   50
                        Left            =   1200
                        Top             =   3000
                     End
                     Begin VB.CommandButton Command1 
                        Caption         =   "Õ–› «·ﬂÊœ"
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
                        Left            =   720
                        TabIndex        =   110
                        Top             =   1920
                        Width           =   1215
                     End
                     Begin MSComCtl2.DTPicker DT6 
                        Height          =   255
                        Left            =   720
                        TabIndex        =   111
                        Top             =   1560
                        Width           =   1215
                        _ExtentX        =   2143
                        _ExtentY        =   450
                        _Version        =   393216
                        Format          =   125829121
                        CurrentDate     =   41162
                     End
                     Begin MSFlexGridLib.MSFlexGrid grd11 
                        Height          =   2775
                        Left            =   5400
                        TabIndex        =   276
                        Top             =   3240
                        Width           =   7815
                        _ExtentX        =   13785
                        _ExtentY        =   4895
                        _Version        =   393216
                        Cols            =   4
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
                     Begin MSFlexGridLib.MSFlexGrid grd80 
                        Height          =   3015
                        Left            =   3720
                        TabIndex        =   296
                        Top             =   120
                        Width           =   9735
                        _ExtentX        =   17171
                        _ExtentY        =   5318
                        _Version        =   393216
                        Cols            =   4
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
                        ForeColor       =   &H00000000&
                        Height          =   375
                        Index           =   60
                        Left            =   3000
                        TabIndex        =   291
                        Top             =   4440
                        Width           =   2055
                     End
                     Begin VB.Label Label41 
                        Caption         =   "Label41"
                        Height          =   255
                        Left            =   1080
                        TabIndex        =   253
                        Top             =   5520
                        Width           =   1815
                     End
                     Begin VB.Label Label40 
                        Caption         =   "Label40"
                        Height          =   255
                        Left            =   1680
                        TabIndex        =   252
                        Top             =   4680
                        Width           =   1815
                     End
                     Begin VB.Label Label39 
                        Caption         =   "Label39"
                        Height          =   255
                        Left            =   1680
                        TabIndex        =   251
                        Top             =   4320
                        Width           =   1815
                     End
                     Begin VB.Label Label38 
                        Caption         =   "Label38"
                        Height          =   255
                        Left            =   1680
                        TabIndex        =   250
                        Top             =   3960
                        Width           =   1815
                     End
                     Begin VB.Label Label37 
                        Caption         =   "Label37"
                        Height          =   255
                        Left            =   1680
                        TabIndex        =   249
                        Top             =   3600
                        Width           =   1815
                     End
                     Begin VB.Label Label29 
                        Caption         =   "Label29"
                        Height          =   255
                        Left            =   120
                        TabIndex        =   220
                        Top             =   5520
                        Width           =   1815
                     End
                     Begin VB.Label Label28 
                        Caption         =   "Label28"
                        Height          =   255
                        Left            =   120
                        TabIndex        =   197
                        Top             =   5160
                        Width           =   1815
                     End
                     Begin VB.Label Label27 
                        Caption         =   "Label27"
                        Height          =   255
                        Left            =   120
                        TabIndex        =   196
                        Top             =   4800
                        Width           =   1815
                     End
                     Begin VB.Label Label1 
                        Height          =   375
                        Left            =   120
                        TabIndex        =   119
                        Top             =   240
                        Width           =   1455
                     End
                     Begin VB.Label Label3 
                        Caption         =   "Label3"
                        Height          =   375
                        Left            =   240
                        TabIndex        =   118
                        Top             =   1080
                        Width           =   2175
                     End
                     Begin VB.Label Label4 
                        Caption         =   "Label4"
                        Height          =   255
                        Left            =   120
                        TabIndex        =   117
                        Top             =   2280
                        Width           =   1335
                     End
                     Begin VB.Label Label20 
                        Caption         =   "Label20"
                        Height          =   255
                        Left            =   120
                        TabIndex        =   116
                        Top             =   3240
                        Width           =   1695
                     End
                     Begin VB.Label Label21 
                        Caption         =   "Label21"
                        Height          =   255
                        Left            =   120
                        TabIndex        =   115
                        Top             =   3720
                        Width           =   1815
                     End
                     Begin VB.Label Label22 
                        Caption         =   "Label22"
                        Height          =   255
                        Left            =   120
                        TabIndex        =   114
                        Top             =   4080
                        Width           =   1695
                     End
                     Begin VB.Label Label23 
                        Caption         =   "Label23"
                        Height          =   255
                        Left            =   120
                        TabIndex        =   113
                        Top             =   4440
                        Width           =   1815
                     End
                  End
                  Begin VB.CommandButton Command7 
                     Caption         =   "„”Õ «·ﬂ·"
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
                     Left            =   6360
                     TabIndex        =   108
                     Top             =   600
                     Width           =   1095
                  End
                  Begin VB.CommandButton Command6 
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
                     Left            =   6480
                     TabIndex        =   107
                     Top             =   6840
                     Width           =   2775
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
                     Left            =   1920
                     TabIndex        =   106
                     Top             =   120
                     Width           =   1335
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
                     ItemData        =   "caisse.frx":03FC
                     Left            =   3960
                     List            =   "caisse.frx":040C
                     Style           =   2  'Dropdown List
                     TabIndex        =   105
                     Top             =   120
                     Width           =   1455
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
                     Left            =   7320
                     TabIndex        =   104
                     Top             =   120
                     Width           =   1815
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
                     Height          =   375
                     Left            =   10680
                     TabIndex        =   103
                     Top             =   120
                     Width           =   1815
                  End
                  Begin VB.CommandButton Command8 
                     Caption         =   " ÕœÌÀ"
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
                     Left            =   12960
                     TabIndex        =   102
                     Top             =   600
                     Width           =   735
                  End
                  Begin VB.CommandButton Command9 
                     Caption         =   "œ›⁄ «·ﬂ·"
                     Enabled         =   0   'False
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
                     Left            =   9960
                     TabIndex        =   101
                     Top             =   600
                     Width           =   975
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
                     ItemData        =   "caisse.frx":043F
                     Left            =   12120
                     List            =   "caisse.frx":044F
                     Style           =   2  'Dropdown List
                     TabIndex        =   100
                     Top             =   6840
                     Width           =   1455
                  End
                  Begin VB.CommandButton Command10 
                     Caption         =   "”Õ» «·Ê’·"
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
                     Left            =   9360
                     TabIndex        =   99
                     Top             =   6840
                     Width           =   1335
                  End
                  Begin MSComctlLib.ProgressBar ProgressBar2 
                     Height          =   255
                     Left            =   6480
                     TabIndex        =   120
                     Top             =   6480
                     Visible         =   0   'False
                     Width           =   5895
                     _ExtentX        =   10398
                     _ExtentY        =   450
                     _Version        =   393216
                     Appearance      =   1
                  End
                  Begin MSFlexGridLib.MSFlexGrid grd2 
                     Height          =   4695
                     Left            =   10920
                     TabIndex        =   121
                     Top             =   1080
                     Width           =   2775
                     _ExtentX        =   4895
                     _ExtentY        =   8281
                     _Version        =   393216
                     Rows            =   1
                     Cols            =   3
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
                  Begin MSFlexGridLib.MSFlexGrid grd3 
                     Height          =   4695
                     Left            =   6360
                     TabIndex        =   122
                     Top             =   1080
                     Width           =   4575
                     _ExtentX        =   8070
                     _ExtentY        =   8281
                     _Version        =   393216
                     Rows            =   1
                     Cols            =   5
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
                  Begin MSFlexGridLib.MSFlexGrid grd4 
                     Height          =   4695
                     Left            =   120
                     TabIndex        =   123
                     Top             =   1080
                     Width           =   6135
                     _ExtentX        =   10821
                     _ExtentY        =   8281
                     _Version        =   393216
                     Rows            =   1
                     Cols            =   7
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
                  Begin MSComCtl2.DTPicker DT2 
                     Height          =   375
                     Left            =   10920
                     TabIndex        =   124
                     Top             =   6000
                     Width           =   1455
                     _ExtentX        =   2566
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
                     Format          =   125501441
                     CurrentDate     =   41154
                  End
                  Begin MSFlexGridLib.MSFlexGrid grd5 
                     Height          =   1455
                     Left            =   1560
                     TabIndex        =   125
                     Top             =   5880
                     Width           =   4695
                     _ExtentX        =   8281
                     _ExtentY        =   2566
                     _Version        =   393216
                     Rows            =   1
                     Cols            =   6
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
                  Begin VB.Label Label48 
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
                     Left            =   120
                     TabIndex        =   294
                     Top             =   120
                     Width           =   1695
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·»«ﬁÌ"
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
                     Index           =   59
                     Left            =   7440
                     TabIndex        =   275
                     Top             =   6000
                     Width           =   975
                  End
                  Begin VB.Label Label45 
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
                     Left            =   6360
                     TabIndex        =   274
                     Top             =   6000
                     Width           =   1335
                  End
                  Begin VB.Label Label107 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "„⁄ «—‘Ì›"
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
                     Left            =   10560
                     TabIndex        =   258
                     Top             =   6840
                     Width           =   1095
                  End
                  Begin VB.Label Label31 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "‘ÂÊ—  Õ„· œÌÊ‰«"
                     BeginProperty Font 
                        Name            =   "Times New Roman"
                        Size            =   14.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   1095
                     Index           =   14
                     Left            =   120
                     TabIndex        =   138
                     Top             =   6240
                     Width           =   1335
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·‘ÂÊ— «·„›Ê⁄…"
                     BeginProperty Font 
                        Name            =   "Times New Roman"
                        Size            =   9
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   375
                     Index           =   13
                     Left            =   12240
                     TabIndex        =   137
                     Top             =   6480
                     Width           =   1335
                  End
                  Begin VB.Label Label2 
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
                     Left            =   8640
                     TabIndex        =   136
                     Top             =   6000
                     Width           =   1215
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·„œ›Ê⁄"
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
                     Index           =   12
                     Left            =   9840
                     TabIndex        =   135
                     Top             =   6000
                     Width           =   975
                  End
                  Begin VB.Label Label15 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   " «—ÌŒ «·œ›⁄"
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
                     Left            =   12240
                     TabIndex        =   134
                     Top             =   6000
                     Width           =   1335
                  End
                  Begin VB.Shape Shape1 
                     BorderColor     =   &H00FFFFFF&
                     BorderWidth     =   2
                     Height          =   1455
                     Left            =   6360
                     Top             =   5880
                     Width           =   7335
                  End
                  Begin VB.Label Label75 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     BeginProperty Font 
                        Name            =   "Times New Roman"
                        Size            =   9
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   375
                     Left            =   6360
                     TabIndex        =   133
                     Top             =   6480
                     Width           =   6015
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "⁄„·Ì«   „  »«·›⁄·"
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
                     Index           =   11
                     Left            =   3600
                     TabIndex        =   132
                     Top             =   600
                     Width           =   1815
                  End
                  Begin VB.Label Label31 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "⁄„·Ì«  ﬁÌœ «· ‰›Ì–"
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
                     Left            =   7080
                     TabIndex        =   131
                     Top             =   600
                     Width           =   3015
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·Ê’·"
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
                     Left            =   3000
                     TabIndex        =   130
                     Top             =   120
                     Width           =   855
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·Õ«·… «· ”ÃÌ·Ì…"
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
                     Left            =   4560
                     TabIndex        =   129
                     Top             =   120
                     Width           =   2535
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·—”Ê„ «·‘Â—Ì…"
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
                     Left            =   8520
                     TabIndex        =   128
                     Top             =   120
                     Width           =   2055
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "—”Ê„ «· ”ÃÌ·"
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
                     Left            =   11640
                     TabIndex        =   127
                     Top             =   120
                     Width           =   2055
                  End
                  Begin VB.Label Label31 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·‘ÂÊ— Ê«·—”Ê„"
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
                     Left            =   10320
                     TabIndex        =   126
                     Top             =   600
                     Width           =   3375
                  End
               End
               Begin MSFlexGridLib.MSFlexGrid grd01 
                  Height          =   4815
                  Left            =   5160
                  TabIndex        =   286
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   7935
                  _ExtentX        =   13996
                  _ExtentY        =   8493
                  _Version        =   393216
                  Rows            =   1
                  Cols            =   5
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
            End
            Begin VB.PictureBox Picture7 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   8175
               Left            =   -74880
               ScaleHeight     =   8175
               ScaleWidth      =   13815
               TabIndex        =   67
               Top             =   360
               Width           =   13815
               Begin TabDlg.SSTab SSTab3 
                  Height          =   7935
                  Left            =   120
                  TabIndex        =   68
                  Top             =   120
                  Width           =   13575
                  _ExtentX        =   23945
                  _ExtentY        =   13996
                  _Version        =   393216
                  Tabs            =   2
                  Tab             =   1
                  TabsPerRow      =   2
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
                  TabCaption(0)   =   "«·„” Õﬁ«  ⁄·Ï «· ·«„Ì–"
                  TabPicture(0)   =   "caisse.frx":0482
                  Tab(0).ControlEnabled=   0   'False
                  Tab(0).Control(0)=   "Picture9"
                  Tab(0).ControlCount=   1
                  TabCaption(1)   =   "œ›⁄ «·√Ê’«·"
                  TabPicture(1)   =   "caisse.frx":049E
                  Tab(1).ControlEnabled=   -1  'True
                  Tab(1).Control(0)=   "Picture8"
                  Tab(1).Control(0).Enabled=   0   'False
                  Tab(1).ControlCount=   1
                  Begin VB.PictureBox Picture8 
                     BackColor       =   &H00000000&
                     BorderStyle     =   0  'None
                     Height          =   7455
                     Left            =   120
                     ScaleHeight     =   7455
                     ScaleWidth      =   13335
                     TabIndex        =   78
                     Top             =   360
                     Width           =   13335
                     Begin VB.CommandButton Command43 
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
                        Height          =   285
                        Left            =   8280
                        TabIndex        =   283
                        Top             =   600
                        Width           =   1335
                     End
                     Begin VB.CommandButton Command11 
                        Caption         =   "⁄—÷ «·√Ê’«·"
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
                        Left            =   6720
                        TabIndex        =   80
                        Top             =   120
                        Width           =   1455
                     End
                     Begin MSComctlLib.ProgressBar ProgressBar3 
                        Height          =   375
                        Left            =   120
                        TabIndex        =   79
                        Top             =   120
                        Width           =   3015
                        _ExtentX        =   5318
                        _ExtentY        =   661
                        _Version        =   393216
                        Appearance      =   1
                     End
                     Begin MSComCtl2.DTPicker DT3 
                        Height          =   375
                        Left            =   11040
                        TabIndex        =   81
                        Top             =   120
                        Width           =   1335
                        _ExtentX        =   2355
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
                        Format          =   125829121
                        CurrentDate     =   41154
                     End
                     Begin MSComCtl2.DTPicker DT4 
                        Height          =   375
                        Left            =   8280
                        TabIndex        =   82
                        Top             =   120
                        Width           =   1335
                        _ExtentX        =   2355
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
                        Format          =   125829121
                        CurrentDate     =   41154
                     End
                     Begin MSFlexGridLib.MSFlexGrid grd7 
                        Height          =   6135
                        Left            =   8280
                        TabIndex        =   83
                        Top             =   960
                        Width           =   4935
                        _ExtentX        =   8705
                        _ExtentY        =   10821
                        _Version        =   393216
                        Cols            =   4
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
                     Begin MSFlexGridLib.MSFlexGrid grd8 
                        Height          =   6135
                        Left            =   4200
                        TabIndex        =   84
                        Top             =   960
                        Width           =   3975
                        _ExtentX        =   7011
                        _ExtentY        =   10821
                        _Version        =   393216
                        Cols            =   4
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
                     Begin MSFlexGridLib.MSFlexGrid grd9 
                        Height          =   6135
                        Left            =   120
                        TabIndex        =   85
                        Top             =   960
                        Width           =   3975
                        _ExtentX        =   7011
                        _ExtentY        =   10821
                        _Version        =   393216
                        Cols            =   4
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
                     Begin VB.Label Label5 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "„‰  «—ÌŒ"
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
                        Left            =   11880
                        TabIndex        =   96
                        Top             =   120
                        Width           =   1335
                     End
                     Begin VB.Label Label5 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "≈·Ï  «—ÌŒ"
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
                        Left            =   9240
                        TabIndex        =   95
                        Top             =   120
                        Width           =   1335
                     End
                     Begin VB.Label Label31 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "⁄œœ √Ê’«· Õ«·… ≈ﬂ „«·"
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
                        Index           =   15
                        Left            =   11160
                        TabIndex        =   94
                        Top             =   600
                        Width           =   2055
                     End
                     Begin VB.Label Label31 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "⁄œœ √Ê’«· Õ«·… «‰”Õ«»"
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
                        Index           =   16
                        Left            =   1920
                        TabIndex        =   93
                        Top             =   600
                        Width           =   2055
                     End
                     Begin VB.Label Label31 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "⁄œœ √Ê’«· Õ«·… ≈⁄›«¡"
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
                        Index           =   17
                        Left            =   6000
                        TabIndex        =   92
                        Top             =   600
                        Width           =   2055
                     End
                     Begin VB.Label Label31 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "„Ã„Ê⁄ ‰ﬁÊœ «·√Ê’«·"
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
                        Index           =   18
                        Left            =   4560
                        TabIndex        =   91
                        Top             =   120
                        Width           =   2055
                     End
                     Begin VB.Label Label6 
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
                        Height          =   255
                        Left            =   3240
                        TabIndex        =   90
                        Top             =   120
                        Width           =   1695
                     End
                     Begin VB.Label Label7 
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
                        Height          =   255
                        Left            =   120
                        TabIndex        =   89
                        Top             =   600
                        Width           =   1815
                     End
                     Begin VB.Label Label8 
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
                        Height          =   255
                        Left            =   4200
                        TabIndex        =   88
                        Top             =   600
                        Width           =   1935
                     End
                     Begin VB.Label Label9 
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
                        Height          =   255
                        Left            =   9720
                        TabIndex        =   87
                        Top             =   600
                        Width           =   1575
                     End
                     Begin VB.Label Label31 
                        Alignment       =   2  'Center
                        BackStyle       =   0  'Transparent
                        Caption         =   $"caisse.frx":04BA
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
                        Index           =   21
                        Left            =   0
                        TabIndex        =   86
                        Top             =   7080
                        Width           =   13335
                     End
                  End
                  Begin VB.PictureBox Picture9 
                     BackColor       =   &H00000000&
                     BorderStyle     =   0  'None
                     Height          =   7455
                     Left            =   -74880
                     ScaleHeight     =   7455
                     ScaleWidth      =   13335
                     TabIndex        =   69
                     Top             =   360
                     Width           =   13335
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
                        ItemData        =   "caisse.frx":054D
                        Left            =   8520
                        List            =   "caisse.frx":055D
                        Style           =   2  'Dropdown List
                        TabIndex        =   73
                        Top             =   120
                        Width           =   1815
                     End
                     Begin VB.CommandButton Command12 
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
                        Left            =   1680
                        TabIndex        =   72
                        Top             =   120
                        Width           =   3135
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
                        ItemData        =   "caisse.frx":0590
                        Left            =   5040
                        List            =   "caisse.frx":05A0
                        Style           =   2  'Dropdown List
                        TabIndex        =   71
                        Top             =   120
                        Width           =   1575
                     End
                     Begin VB.CommandButton Command13 
                        Caption         =   "”Õ» «··«∆Õ…"
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
                        Left            =   1920
                        TabIndex        =   70
                        Top             =   6960
                        Width           =   3135
                     End
                     Begin MSFlexGridLib.MSFlexGrid grd10 
                        Height          =   6735
                        Left            =   6720
                        TabIndex        =   74
                        Top             =   600
                        Width           =   6495
                        _ExtentX        =   11456
                        _ExtentY        =   11880
                        _Version        =   393216
                        Cols            =   4
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
                     Begin MSFlexGridLib.MSFlexGrid grd12 
                        Height          =   6255
                        Left            =   120
                        TabIndex        =   75
                        Top             =   600
                        Width           =   6495
                        _ExtentX        =   11456
                        _ExtentY        =   11033
                        _Version        =   393216
                        Cols            =   4
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
                     Begin VB.Label Label31 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "·«∆Õ…  ·«„Ì– «·ﬁ”„"
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
                        Index           =   19
                        Left            =   8880
                        TabIndex        =   77
                        Top             =   120
                        Width           =   3015
                     End
                     Begin VB.Label Label31 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "«·„” Õﬁ ⁄·ÌÂ„ ‘Â—"
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
                        Index           =   20
                        Left            =   6120
                        TabIndex        =   76
                        Top             =   120
                        Width           =   2295
                     End
                  End
               End
            End
         End
      End
      Begin VB.PictureBox Picture10 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8895
         Left            =   -74880
         ScaleHeight     =   8895
         ScaleWidth      =   14295
         TabIndex        =   14
         Top             =   360
         Width           =   14295
         Begin TabDlg.SSTab SSTab4 
            Height          =   8655
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   14055
            _ExtentX        =   24791
            _ExtentY        =   15266
            _Version        =   393216
            Tabs            =   2
            Tab             =   1
            TabsPerRow      =   2
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
            TabCaption(0)   =   "—ﬂ‰ «·„ «»⁄…"
            TabPicture(0)   =   "caisse.frx":05D3
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Picture14"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "—ﬂ‰ «·œ›⁄"
            TabPicture(1)   =   "caisse.frx":05EF
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "Picture13"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "Picture11"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).ControlCount=   2
            Begin VB.PictureBox Picture11 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   615
               Left            =   120
               ScaleHeight     =   615
               ScaleWidth      =   13815
               TabIndex        =   60
               Top             =   360
               Width           =   13815
               Begin VB.TextBox Text18 
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
                  Left            =   11160
                  TabIndex        =   289
                  Top             =   120
                  Width           =   1815
               End
               Begin VB.CommandButton Command14 
                  Caption         =   "⁄—÷ «·Õ”«»"
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
                  Left            =   120
                  TabIndex        =   62
                  Top             =   120
                  Width           =   1455
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
                  Left            =   7320
                  TabIndex        =   61
                  Top             =   120
                  Width           =   1935
               End
               Begin VB.Label Label31 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Â« ›                                     √Ê «·—ﬁ„ «· ”·”·Ì"
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
                  Left            =   9360
                  TabIndex        =   288
                  Top             =   120
                  Width           =   4335
               End
               Begin VB.Label Label10 
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
                  Left            =   3600
                  TabIndex        =   64
                  Top             =   120
                  Width           =   3615
               End
               Begin VB.Label Label19 
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
                  Left            =   1680
                  TabIndex        =   63
                  Top             =   120
                  Width           =   1815
               End
            End
            Begin VB.PictureBox Picture13 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   7455
               Left            =   120
               ScaleHeight     =   7455
               ScaleWidth      =   13815
               TabIndex        =   25
               Top             =   1080
               Width           =   13815
               Begin VB.PictureBox Picture12 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   7455
                  Left            =   0
                  ScaleHeight     =   7455
                  ScaleWidth      =   13815
                  TabIndex        =   26
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   13815
                  Begin VB.CommandButton Command24 
                     Caption         =   "”Õ»"
                     Enabled         =   0   'False
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
                     Left            =   120
                     TabIndex        =   32
                     Top             =   1440
                     Width           =   1455
                  End
                  Begin VB.CommandButton Command25 
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
                     Left            =   1680
                     TabIndex        =   31
                     Top             =   1440
                     Width           =   1695
                  End
                  Begin VB.CommandButton Command21 
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
                     Left            =   1680
                     TabIndex        =   30
                     Top             =   480
                     Width           =   3375
                  End
                  Begin VB.CommandButton Command23 
                     Caption         =   "≈·€«¡"
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
                     Left            =   120
                     TabIndex        =   29
                     Top             =   480
                     Width           =   1455
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
                     Left            =   5160
                     TabIndex        =   28
                     Top             =   480
                     Width           =   1215
                  End
                  Begin VB.CommandButton Command15 
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
                     Left            =   8880
                     TabIndex        =   27
                     Top             =   1440
                     Width           =   1815
                  End
                  Begin MSFlexGridLib.MSFlexGrid grd13 
                     Height          =   5415
                     Left            =   5040
                     TabIndex        =   33
                     Top             =   1920
                     Width           =   8775
                     _ExtentX        =   15478
                     _ExtentY        =   9551
                     _Version        =   393216
                     Cols            =   5
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
                        Size            =   9.75
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin MSComCtl2.DTPicker DT7 
                     Height          =   375
                     Left            =   7920
                     TabIndex        =   34
                     Top             =   480
                     Width           =   1335
                     _ExtentX        =   2355
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
                     Format          =   125501441
                     CurrentDate     =   41154
                  End
                  Begin MSComctlLib.ProgressBar ProgressBar5 
                     Height          =   375
                     Left            =   5160
                     TabIndex        =   35
                     Top             =   960
                     Width           =   5535
                     _ExtentX        =   9763
                     _ExtentY        =   661
                     _Version        =   393216
                     Appearance      =   1
                  End
                  Begin MSFlexGridLib.MSFlexGrid grd14 
                     Height          =   5175
                     Left            =   120
                     TabIndex        =   36
                     Top             =   1920
                     Width           =   4935
                     _ExtentX        =   8705
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
                  Begin MSComCtl2.DTPicker DT8 
                     Height          =   375
                     Left            =   1920
                     TabIndex        =   37
                     Top             =   960
                     Width           =   1455
                     _ExtentX        =   2566
                     _ExtentY        =   661
                     _Version        =   393216
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Times New Roman"
                        Size            =   11.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Format          =   125501441
                     CurrentDate     =   41154
                  End
                  Begin MSComCtl2.DTPicker DT9 
                     Height          =   375
                     Left            =   120
                     TabIndex        =   38
                     Top             =   960
                     Width           =   1455
                     _ExtentX        =   2566
                     _ExtentY        =   661
                     _Version        =   393216
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Times New Roman"
                        Size            =   11.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Format          =   125501441
                     CurrentDate     =   41154
                  End
                  Begin VB.Label Label31 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "%"
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
                     Index           =   36
                     Left            =   11880
                     TabIndex        =   59
                     Top             =   960
                     Width           =   375
                  End
                  Begin VB.Label Label18 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "30"
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
                     Left            =   11280
                     TabIndex        =   58
                     Top             =   960
                     Width           =   495
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
                     Index           =   35
                     Left            =   12120
                     TabIndex        =   57
                     Top             =   960
                     Width           =   1575
                  End
                  Begin VB.Label Label31 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "≈·Ï"
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
                     Index           =   34
                     Left            =   1440
                     TabIndex        =   56
                     Top             =   960
                     Width           =   615
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "⁄—÷ ”Ã· «·œ›⁄ „‰"
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
                     Index           =   33
                     Left            =   3000
                     TabIndex        =   55
                     Top             =   960
                     Width           =   2055
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·„»·€ «·„œ›Ê⁄"
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
                     Index           =   32
                     Left            =   5760
                     TabIndex        =   54
                     Top             =   480
                     Width           =   2055
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   " «—ÌŒ «·œ›⁄"
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
                     Index           =   31
                     Left            =   8640
                     TabIndex        =   53
                     Top             =   480
                     Width           =   2055
                  End
                  Begin VB.Label Label17 
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
                     Left            =   10800
                     TabIndex        =   52
                     Top             =   480
                     Width           =   1335
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·—’Ìœ «·‰Â«∆Ì"
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
                     Index           =   30
                     Left            =   11640
                     TabIndex        =   51
                     Top             =   480
                     Width           =   2055
                  End
                  Begin VB.Label Label16 
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
                     Left            =   120
                     TabIndex        =   50
                     Top             =   120
                     Width           =   1215
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "—’Ìœ «·œ›⁄"
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
                     Index           =   29
                     Left            =   840
                     TabIndex        =   49
                     Top             =   120
                     Width           =   1455
                  End
                  Begin VB.Label Label14 
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
                     Left            =   2400
                     TabIndex        =   48
                     Top             =   120
                     Width           =   1215
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·—’Ìœ «·«Ã„«·Ì"
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
                     Index           =   28
                     Left            =   3240
                     TabIndex        =   47
                     Top             =   120
                     Width           =   1815
                  End
                  Begin VB.Label Label31 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   " ›«’Ì· „” Õﬁ«  «·‰”»…"
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
                     Index           =   27
                     Left            =   10800
                     TabIndex        =   46
                     Top             =   1440
                     Width           =   2895
                  End
                  Begin VB.Label Label13 
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
                     Left            =   5160
                     TabIndex        =   45
                     Top             =   120
                     Width           =   1215
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "„” Õﬁ«  «·‰”»…"
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
                     Index           =   26
                     Left            =   6120
                     TabIndex        =   44
                     Top             =   120
                     Width           =   1695
                  End
                  Begin VB.Label Label12 
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
                     Left            =   7920
                     TabIndex        =   43
                     Top             =   120
                     Width           =   1335
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "„” Õﬁ«  «·‘Â—"
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
                     Left            =   8640
                     TabIndex        =   42
                     Top             =   120
                     Width           =   2055
                  End
                  Begin VB.Label Label11 
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
                     Left            =   10800
                     TabIndex        =   41
                     Top             =   120
                     Width           =   1335
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "„” Õﬁ«  «·”«⁄…"
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
                     Index           =   24
                     Left            =   11640
                     TabIndex        =   40
                     Top             =   120
                     Width           =   2055
                  End
                  Begin VB.Label Label31 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "≈‰ Õ–› √Ì „»·€ Ì⁄‰Ì «—Ã«⁄Â „‰ «·√” «– ≈·Ï «·’‰œÊﬁ „Õ«”»Ì«"
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
                     Index           =   37
                     Left            =   0
                     TabIndex        =   39
                     Top             =   7080
                     Width           =   5175
                  End
               End
            End
            Begin VB.PictureBox Picture14 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   8175
               Left            =   -74880
               ScaleHeight     =   8175
               ScaleWidth      =   13815
               TabIndex        =   16
               Top             =   360
               Width           =   13815
               Begin VB.CommandButton Command16 
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
                  Left            =   3840
                  TabIndex        =   18
                  Top             =   120
                  Width           =   3015
               End
               Begin VB.CommandButton Command17 
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
                  Left            =   1920
                  TabIndex        =   17
                  Top             =   120
                  Width           =   1815
               End
               Begin MSComCtl2.DTPicker DT10 
                  Height          =   375
                  Left            =   9600
                  TabIndex        =   19
                  Top             =   120
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
                  Format          =   125829121
                  CurrentDate     =   41154
               End
               Begin MSComCtl2.DTPicker DT11 
                  Height          =   375
                  Left            =   6960
                  TabIndex        =   20
                  Top             =   120
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
                  Format          =   125829121
                  CurrentDate     =   41154
               End
               Begin MSFlexGridLib.MSFlexGrid grd15 
                  Height          =   7095
                  Left            =   120
                  TabIndex        =   21
                  Top             =   960
                  Width           =   13575
                  _ExtentX        =   23945
                  _ExtentY        =   12515
                  _Version        =   393216
                  Cols            =   4
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
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "≈·Ï  «—ÌŒ"
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
                  Left            =   8160
                  TabIndex        =   24
                  Top             =   120
                  Width           =   1335
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‰  «—ÌŒ"
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
                  Left            =   10680
                  TabIndex        =   23
                  Top             =   120
                  Width           =   1335
               End
               Begin VB.Label Label5 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "„” Õﬁ«  Ê„œ›Ê⁄«  Ê√—’œ… √”« –… «·”«⁄… Ê«·‘Â—"
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
                  Left            =   4200
                  TabIndex        =   22
                  Top             =   600
                  Width           =   5415
               End
            End
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   -74880
         ScaleHeight     =   975
         ScaleWidth      =   14295
         TabIndex        =   1
         Top             =   360
         Width           =   14295
         Begin VB.ComboBox Combo13 
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
            ItemData        =   "caisse.frx":060B
            Left            =   3600
            List            =   "caisse.frx":0621
            Style           =   2  'Dropdown List
            TabIndex        =   230
            Top             =   120
            Visible         =   0   'False
            Width           =   1095
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
            Left            =   8400
            MaskColor       =   &H00000000&
            TabIndex        =   12
            Top             =   600
            Width           =   255
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
            Height          =   375
            Left            =   12000
            TabIndex        =   6
            Top             =   120
            Width           =   1575
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
            Height          =   375
            Left            =   8760
            TabIndex        =   5
            Top             =   120
            Width           =   2655
         End
         Begin VB.ComboBox Combo6 
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
            ItemData        =   "caisse.frx":066B
            Left            =   5880
            List            =   "caisse.frx":0681
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   120
            Width           =   1695
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Õ›Ÿ «·ﬂÊœ"
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
            Left            =   960
            TabIndex        =   3
            Top             =   120
            Width           =   2535
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
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   735
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
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
            Index           =   58
            Left            =   4440
            TabIndex        =   231
            Top             =   120
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "”Ì „ «ŸÂ«— «·ﬂÊœ ›Ì «·ﬁÊ«∆„ «·„‰”œ·…"
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
            Index           =   42
            Left            =   3600
            TabIndex        =   13
            Top             =   600
            Width           =   4695
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·ﬂÊœ"
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
            Left            =   12240
            TabIndex        =   10
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·»Ì«‰"
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
            Left            =   9960
            TabIndex        =   9
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " ’‰Ì› «·ﬂÊœ"
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
            Left            =   6720
            TabIndex        =   8
            Top             =   120
            Width           =   1935
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grd6 
         Height          =   7815
         Left            =   -74880
         TabIndex        =   11
         Top             =   1440
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   13785
         _Version        =   393216
         Rows            =   1
         Cols            =   4
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
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "caisse.frx":06CB
      Top             =   240
   End
End
Attribute VB_Name = "caisse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Public co2 As ADODB.Connection
Public cr2 As ADODB.Recordset
Public be As ADODB.Recordset
'Public ce As ADODB.Recordset
Public ru As ADODB.Recordset
Public ep As ADODB.Recordset
Public pe As ADODB.Recordset
Public pd As ADODB.Recordset
Public yp As ADODB.Recordset
Public rpp As ADODB.Recordset
Public dps As ADODB.Recordset
Public bnk As ADODB.Recordset
Public cai As ADODB.Recordset
Dim anne As String
Dim tim As Integer
Dim data As New Access.Application
Function cont2()
Set co2 = New ADODB.Connection
Set cr2 = New ADODB.Recordset
Set be = New ADODB.Recordset
'Set ce = New ADODB.Recordset
Set ru = New ADODB.Recordset
Set ep = New ADODB.Recordset
Set pe = New ADODB.Recordset
Set pd = New ADODB.Recordset
Set yp = New ADODB.Recordset
Set rpp = New ADODB.Recordset
Set dps = New ADODB.Recordset
Set bnk = New ADODB.Recordset
Set cai = New ADODB.Recordset
anne = "C" + face.SBB1.Panels(9).Text
co2.Provider = "microsoft.jet.oledb.4.0; jet oledb:database password=7346804"
co2.ConnectionString = App.Path & "\" & anne & ".mdb"
co2.Open
cr2.Open "select*from Tcarts", co2, adOpenKeyset, adLockOptimistic
be.Open "select*from Tbulletin", co2, adOpenKeyset, adLockOptimistic
'ce.Open "select*from Tcartes", co2, adOpenKeyset, adLockOptimistic
ru.Open "select*from Trecus", co2, adOpenKeyset, adLockOptimistic
ep.Open "select*from Tetudpaspaye order by num ASC", co2, adOpenKeyset, adLockOptimistic
pe.Open "select*from Tprofpaspaye order by dat ASC", co2, adOpenKeyset, adLockOptimistic
pd.Open "select*from Tprofpourcentage", co2, adOpenKeyset, adLockOptimistic
yp.Open "select*from Tpayprof order by ser ASC", co2, adOpenKeyset, adLockOptimistic
rpp.Open "select*from Trecpaypartenaires", co2, adOpenKeyset, adLockOptimistic
dps.Open "select*from Tdepenses", co2, adOpenKeyset, adLockOptimistic
bnk.Open "select*from Tbanks", co2, adOpenKeyset, adLockOptimistic
cai.Open "select*from Tcaisses", co2, adOpenKeyset, adLockOptimistic
End Function

Private Sub Check2_Click()
On Error Resume Next
If Check2.Value = 0 Then
Label31(42).Caption = "”Ì „ «ŸÂ«— «·»Ì«‰ ›Ì «·ﬁÊ«∆„ «·„‰”œ·…"
Else
Label31(42).Caption = "”Ì „ «ŸÂ«— «·ﬂÊœ ›Ì «·ﬁÊ«∆„ «·„‰”œ·…"
End If
Call cont
sr!cdc = Check2.Value
sr.Update
Call chargegrd6
End Sub

Private Sub Combo1_Change()
On Error Resume Next
Call cont
Do While Not cd.EOF
If Combo1.Text = cd!cod Then
Label21.Caption = cd!dec
Label37.Caption = cd!cas
'Text3.SetFocus
Exit Sub
End If
If Combo1.Text = cd!dec Then
Label21.Caption = cd!cod
Label37.Caption = cd!cas
Text3.SetFocus
Exit Sub
End If
cd.MoveNext
Loop
End Sub

Private Sub Combo1_Click()
On Error Resume Next
Combo1_Change
End Sub

Private Sub Combo10_Change()
On Error Resume Next
If Combo10.Text = "«·’‰œÊﬁ" Then
Label31(48).Caption = "«·‘—Ìﬂ"
Else
Label31(48).Caption = "«·’‰œÊﬁ"
End If
End Sub

Private Sub Combo10_Click()
On Error Resume Next
Combo10_Change
End Sub

Private Sub Combo11_Change()
On Error Resume Next
Call cont
Do While Not cd.EOF
If Combo11.Text = cd!cod Then
Label29.Caption = cd!dec
Label41.Caption = cd!cas
Combo12.Text = cd!der
Label31(53).Caption = "„»·€ «·" + Combo12.Text
If Combo12.Text = "—√” «·„«·" Then
Label31(53).Caption = "„»·€ " + Combo12.Text
End If
Text13.SetFocus
Exit Sub
End If
If Combo11.Text = cd!dec Then
Label29.Caption = cd!cod
Label41.Caption = cd!cas
Combo12.Text = cd!der
Label31(53).Caption = "„»·€ «·" + Combo12.Text
If Combo12.Text = "—√” «·„«·" Then
Label31(53).Caption = "„»·€ " + Combo12.Text
End If
Text13.SetFocus
Exit Sub
End If
cd.MoveNext
Loop

End Sub

Private Sub Combo11_Click()
On Error Resume Next
Combo11_Change
End Sub

Private Sub Combo2_Change()
On Error Resume Next
Text4.Text = Trim(Text4.Text)
Text5.Text = Trim(Text5.Text)
'If grd4.Rows > 1 And Combo2.Text = "≈⁄›«¡ ﬂ·Ì" Then
'MsgBox "·« Ì„ﬂ‰ ≈⁄›«¡ ﬂ·Ì", vbCritical
'Call chargec2
'Exit Sub
'End If
'If grd4.Rows = 1 And Combo2.Text = "≈⁄›«¡ Ã“∆Ì" Or grd4.Rows = 1 And Combo2.Text = "«‰”Õ«»" Then
'MsgBox "·« Ì„ﬂ‰ ≈⁄›«¡ Ã“∆Ì Ê·« «‰”Õ«»", vbCritical
'Call chargec2
'Exit Sub
'End If
grd3.Clear
grd3.Rows = 1
Label2.Caption = ""
Label75.Caption = ""
Call chargegrd2
Command7.Enabled = True
grd2.Enabled = True
If Text4.Text = "0" Then
Text4.Text = ""
End If
If Text5.Text = "0" Then
Text5.Text = ""
End If
If Combo2.Text <> "Õ«·… ≈ﬂ„«·" Then
Command7.Enabled = False
grd2.Enabled = False
Call ivaa
If Text4.Text = "" Then
Text4.Text = "0"
End If
If Text5.Text = "" Then
Text5.Text = "0"
End If
End If
Command9.Enabled = False
If grd4.Rows = 1 And Combo2.Text = "Õ«·… ≈ﬂ„«·" Then
Command9.Enabled = True
End If
End Sub

Private Sub Combo2_Click()
On Error Resume Next
Combo2_Change
End Sub

Private Sub Combo4_Change()
On Error Resume Next
grd10.Clear
grd10.Rows = 1
grd11.Clear
grd11.Rows = 1
grd12.Clear
grd12.Rows = 1

End Sub

Private Sub Combo4_Click()
On Error Resume Next
Combo4_Change
End Sub

Private Sub Combo5_Change()
On Error Resume Next
grd10.Clear
grd10.Rows = 1
grd11.Clear
grd11.Rows = 1
grd12.Clear
grd12.Rows = 1

End Sub

Private Sub Combo5_Click()
On Error Resume Next
Combo5_Change
End Sub

Private Sub Combo6_Change()
On Error Resume Next
Label31(58).Visible = False
Combo13.Visible = False
Call chargec12
If Combo6.Text = "«·»‰ﬂ" Then
Label31(58).Visible = True
Combo13.Visible = True
End If
End Sub

Private Sub Combo6_Click()
On Error Resume Next
Combo6_Change
End Sub

Private Sub Combo7_Change()
On Error Resume Next
Call cont
Do While Not cd.EOF
If Combo7.Text = cd!cod Then
Label22.Caption = cd!dec
Label38.Caption = cd!cas
'Text7.SetFocus
Exit Sub
End If
If Combo7.Text = cd!dec Then
Label22.Caption = cd!cod
Label38.Caption = cd!cas
Text7.SetFocus
Exit Sub
End If
cd.MoveNext
Loop

End Sub

Private Sub Combo7_Click()
On Error Resume Next
Combo7_Change
End Sub

Private Sub Combo8_Change()
On Error Resume Next
Call cont
Do While Not cd.EOF
If Combo8.Text = cd!cod Then
Label23.Caption = cd!dec
Label40.Caption = cd!cas
'Text9.SetFocus
Exit Sub
End If
If Combo8.Text = cd!dec Then
Label23.Caption = cd!cod
Label40.Caption = cd!cas
Text9.SetFocus
Exit Sub
End If
cd.MoveNext
Loop


End Sub

Private Sub Combo8_Click()
On Error Resume Next
Combo8_Change
End Sub

Private Sub Combo9_Change()
On Error Resume Next
Call cont
Do While Not cd.EOF
If Combo9.Text = cd!cod Then
Label27.Caption = cd!dec
Label39.Caption = cd!cas
Text11.SetFocus
Exit Sub
End If
If Combo9.Text = cd!dec Then
Label27.Caption = cd!cod
Label39.Caption = cd!cas
Text11.SetFocus
Exit Sub
End If
cd.MoveNext
Loop

End Sub

Private Sub Combo9_Click()
On Error Resume Next
Combo9_Change
End Sub

Private Sub Command1_Click()
On Error Resume Next
If Label1.Caption = "" Then
MsgBox "ﬁ„ »«·÷€ÿ ⁄·Ï «·ﬂÊœ «·„—«œ Õ–›Â", vbCritical
Exit Sub
End If
g = MsgBox("Â·  —Ìœ «·Õ–› Õﬁ« ", vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
Call cont
Do While Not cd.EOF
If Label1.Caption = cd!aut Then
cd.Delete
ProgressBar1.Value = 0
Timer1.Enabled = True
Exit Sub
End If
cd.MoveNext
Loop
End If


End Sub

Private Sub Command10_Click()
On Error Resume Next
Dim ane As String
Dim a As Double
a = Combo3.Text
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
data.DoCmd.Maximize
If Check6.Value = 0 Then
data.DoCmd.OpenReport "recu", acViewPreview, , "rec =" & a, acWindowNormal, OpenArgs
Else
data.DoCmd.OpenReport "recu2", acViewPreview, , "rec =" & a, acWindowNormal, OpenArgs
End If
'data.DoCmd.RunMacro "Opnreport", 2
data.Visible = True
Set data = Nothing

End Sub

Private Sub Command11_Click()
On Error Resume Next
Label6.Caption = "0"
Label7.Caption = "0"
Label8.Caption = "0"
Label9.Caption = "0"
grd7.Clear
grd7.Rows = 1
grd8.Clear
grd8.Rows = 1
grd9.Clear
grd9.Rows = 1
grd7.Visible = False
grd8.Visible = False
grd9.Visible = False
Call chargegrd7
grd7.Visible = True
grd8.Visible = True
grd9.Visible = True
End Sub

Private Sub Command12_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim p As Double
Dim a As Double
Dim n As Double
Dim m As Double
Dim k As Double
Dim nmm1 As String
Dim nmm2 As String
Dim nom1 As String
Dim ser1 As String
Dim tel1 As String
Dim adr1 As String
If Combo4.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·ﬁ”„", vbCritical
Exit Sub
End If
If Combo5.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·‘Â—", vbCritical
Exit Sub
End If
grd10.Visible = False
grd12.Visible = False
grd10.Clear
grd10.Rows = 1
grd10.Cols = 6
grd10.ColWidth(0) = 0
grd10.ColWidth(1) = 1100
grd10.ColWidth(2) = 3400
grd10.ColWidth(3) = 1500
grd10.ColWidth(4) = 0
grd10.ColWidth(5) = 0
grd10.row = 0
grd10.Col = 1
grd10.Text = "«·—ﬁ„"
grd10.Col = 2
grd10.Text = "«·«”„"
grd10.Col = 3
grd10.Text = "«·—ﬁ„ «· ”·”·Ì"
grd10.ColAlignment(1) = 1
grd10.ColAlignment(2) = 1
grd10.ColAlignment(3) = 1
i = 1
Call cont
grd10.Rows = et.RecordCount + 2
Do While Not et.EOF
If Combo4.Text = et!cla And Val(et!num) < 1000000 Then
grd10.row = i
grd10.Col = 0
grd10.Text = et!aut
grd10.Col = 1
grd10.Text = et!num
grd10.Col = 2
grd10.Text = et!nom
grd10.Col = 3
grd10.Text = et!ser
grd10.Col = 4
grd10.Text = et!tel
grd10.Col = 5
grd10.Text = et!adr
i = i + 1
End If
et.MoveNext
Loop
grd10.Rows = i
grd10.Col = 1
grd10.Sort = 1
'**** grd11
grd11.Clear
grd11.Rows = 1
grd11.Cols = 4
grd11.ColWidth(0) = 0
grd11.ColWidth(1) = 800
grd11.ColWidth(2) = 2500
grd11.ColWidth(3) = 1500
grd11.row = 0
grd11.Col = 1
grd11.Text = "«·—ﬁ„"
grd11.Col = 2
grd11.Text = "«·«”„"
grd11.Col = 3
grd11.Text = "«·—ﬁ„ «· ”·”·Ì"
grd11.ColAlignment(1) = 1
grd11.ColAlignment(2) = 1
grd11.ColAlignment(3) = 1
i = 1
Call cont
grd11.Rows = ce.RecordCount + 2
Do While Not ce.EOF
'DT6.Value = rc!dat
'a = DT6.Month
If Combo4.Text = ce!cla And Combo5.Text = ce!moi Then
'If ce!cas = "Õ«·… «‰”Õ«»" Or ce!cas = "Õ«·… ≈⁄›«¡" Then
grd11.row = i
grd11.Col = 0
grd11.Text = ce!aut
grd11.Col = 1
grd11.Text = ce!num
grd11.Col = 2
grd11.Text = ce!nom
grd11.Col = 3
grd11.Text = ce!ser
i = i + 1
'End If
End If
ce.MoveNext
Loop
grd11.Rows = i
grd11.Col = 1
grd11.Sort = 1
'**** grd12
grd12.Clear
grd12.Rows = 1
grd12.Cols = 6
grd12.ColWidth(0) = 0
grd12.ColWidth(1) = 1100
grd12.ColWidth(2) = 3400
grd12.ColWidth(3) = 1500
grd12.ColWidth(4) = 0
grd12.ColWidth(5) = 0
grd12.row = 0
grd12.Col = 1
grd12.Text = "«·—ﬁ„"
grd12.Col = 2
grd12.Text = "«·«”„"
grd12.Col = 3
grd12.Text = "«·—ﬁ„ «· ”·”·Ì"
grd12.ColAlignment(1) = 1
grd12.ColAlignment(2) = 1
grd12.ColAlignment(3) = 1
n = grd10.Rows
m = grd11.Rows
grd12.Rows = n + 3
p = 1
For i = 1 To n - 1
k = 0
grd10.row = i
grd10.Col = 1
nmm1 = grd10.Text
grd10.Col = 2
nom1 = grd10.Text
grd10.Col = 3
ser1 = grd10.Text
grd10.Col = 4
tel1 = grd10.Text
grd10.Col = 5
adr1 = grd10.Text
For j = 1 To m - 1
grd11.row = j
grd11.Col = 1
nmm2 = grd11.Text
If nmm2 = nmm1 Then
k = 1
j = m
End If
Next j
If k = 0 Then
grd12.row = p
grd12.Col = 0
grd12.Text = ""
grd12.Col = 1
grd12.Text = nmm1
grd12.Col = 2
grd12.Text = nom1
grd12.Col = 3
grd12.Text = ser1
grd12.Col = 4
grd12.Text = tel1
grd12.Col = 5
grd12.Text = adr1
p = p + 1
End If
Next i
grd12.Rows = p
grd12.Col = 1
grd12.Sort = 1
grd10.Visible = True
grd12.Visible = True
End Sub

Private Sub Command13_Click()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim ane As String
Command13.Enabled = False
Call cont2
Do While Not ep.EOF
ep.Delete
ep.MoveNext
Loop
n = grd12.Rows
For i = 1 To n - 1
ep.AddNew
ep!cla = Combo4.Text
ep!moi = Combo5.Text
grd12.row = i
grd12.Col = 1
ep!num = grd12.Text
grd12.Col = 2
ep!nom = grd12.Text
grd12.Col = 3
ep!ser = grd12.Text
grd12.Col = 4
ep!tel = grd12.Text
grd12.Col = 5
ep!adr = grd12.Text
ep.Update
Next i
tim = 1
ProgressBar9.Value = 0
Timer9.Enabled = True

End Sub

Private Sub Command14_Click()
On Error Resume Next
If Text7.Text = "" Then
MsgBox "«œŒ· «·—ﬁ„ «· ”·”·Ì ··√” «–", vbCritical
Text7.SetFocus
Exit Sub
End If
If Label10.Caption = "" Then
MsgBox "«œŒ· «·—ﬁ„ «· ”·”·Ì ··√” «– À„ «÷€ÿ ⁄·Ï Enter", vbCritical + arabic
Text7.SetFocus
Exit Sub
End If
grd13.Clear
grd13.Rows = 1
Call chargegrd13
grd14.Clear
grd14.Rows = 1
Call chargegrd14
Picture12.Visible = True
End Sub

Private Sub Command15_Click()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim m As Double
Dim s As Double
Dim ane As String
Command15.Enabled = False
Call cont2
Do While Not pd.EOF
pd.Delete
pd.MoveNext
Loop
s = 0
m = 0
n = grd13.Rows
For i = 1 To n - 1
pd.AddNew
pd!nom = Label10.Caption
pd!ser = Text7.Text
pd!pou = Label18.Caption
pd!tos = Label13.Caption
grd13.row = i
grd13.Col = 0
pd!moi = grd13.Text
grd13.Col = 1
pd!cla = grd13.Text
grd13.Col = 2
pd!etu = grd13.Text
grd13.Col = 3
pd!pro = grd13.Text
grd13.Col = 4
pd!res = grd13.Text
grd13.Col = 5
pd!eta = grd13.Text
grd13.Col = 6
pd!bil = grd13.Text
grd13.Col = 7
pd!hec = grd13.Text
grd13.Col = 8
pd!ppr = grd13.Text
grd13.Col = 9
pd!hep = grd13.Text
grd13.Col = 10
pd!dus = grd13.Text
pd.Update
Next i
tim = 3
ProgressBar9.Value = 0
Timer9.Enabled = True

End Sub

Private Sub Command16_Click()
On Error Resume Next
grd15.Visible = False
grd15.Clear
grd15.Rows = 1
Call chargegrd15
grd15.Visible = True
End Sub

Private Sub Command17_Click()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim d As Double
Dim p As Double
Dim r As Double
Dim sd As Double
Dim sp As Double
Dim sr As Double
Dim ane As String
Command17.Enabled = False
Call cont2
Do While Not yp.EOF
yp.Delete
yp.MoveNext
Loop
sd = 0
sp = 0
sr = 0
d = 0
p = 0
r = 0
n = grd15.Rows
For i = 1 To n - 1
grd15.row = i
grd15.Col = 2
d = grd15.Text
sd = sd + d
grd15.Col = 3
p = grd15.Text
sp = sp + p
grd15.Col = 4
r = grd15.Text
sr = sr + r
Next i
For i = 1 To n - 1
yp.AddNew
yp!dat1 = DT10.Value
yp!dat2 = DT11.Value
grd15.row = i
grd15.Col = 0
yp!ser = grd15.Text
grd15.Col = 1
yp!nom = grd15.Text
grd15.Col = 2
yp!dus = grd15.Text
grd15.Col = 3
yp!pay = grd15.Text
grd15.Col = 4
yp!res = grd15.Text
yp!tdu = sd
yp!tpy = sp
yp!trs = sr
yp.Update
Next i
tim = 4
ProgressBar9.Value = 0
Timer9.Enabled = True

End Sub

Private Sub Command18_Click()
On Error Resume Next
Dim a As Double
Dim b As Double
Dim c As Double
Dim au As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim dat4 As Date
Dim annes As String
Dim mca As Double
Dim mnv As Double
Dim j As Double
Dim i As Double
Text9.Text = Trim(Text9.Text)
Text10.Text = Trim(Text10.Text)
If Combo8.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— ﬂÊœ «·⁄„·Ì…", vbCritical
Exit Sub
End If
If Text9.Text = "" Then
MsgBox "«œŒ· »Ì«‰ «·⁄„·Ì…", vbCritical
Text9.SetFocus
Exit Sub
End If
If Text10.Text = "" Then
MsgBox "«œŒ· «·„»·€ «·„’—Ê›", vbCritical
Text10.SetFocus
Exit Sub
End If
Text15.Text = Text9.Text
Text16.Text = Mid$(Text15.Text, 1, 18)
Text16.Text = Trim(Text16.Text)
If Text16.Text = " ”œÌœ —« » «·„ÊŸ›" Then
MsgBox "€Ì— „„ﬂ‰..  ”œÌœ —Ê« » «·„ÊŸ›Ì‰ Ì „ ›Ì ﬁ”„ «·„ÊŸ›Ì‰", vbCritical
Exit Sub
End If
Call cont
'**** controle caisse ajouter
mca = sr!cca
mnv = Text10.Text
If mnv > mca Then
MsgBox "€Ì— „„ﬂ‰.. ·√‰Â ·« ÌÊÃœ ›Ì «·’‰œÊﬁ Õ«·Ì« Â–« «·„»·€", vbCritical
Exit Sub
End If
'**** end controle caisse
'**** controle Date Ajouter
annes = sr!ann      'annee scolaire
dat1 = sr!dat       'debut annee scolaire
dat2 = sr!dtf       'Fin annees scolaire
dat3 = DT15.Value    'Date Modifiable
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
'****** caisse
ca.AddNew
au = ca!aun
If Val(Combo8.Text) > 0 Then
ca!cod = Combo8.Text
ca!dec = Label23.Caption
Else
ca!dec = Combo8.Text
ca!cod = Label23.Caption
End If
ca!mem = "’—› „»·€ " + Text10.Text + " „‰ √Ã· " + Text9.Text
ca!mon = Text10.Text
ca!cas = "Œ«—Ã"
ca!heu = Time$
ca!dat = DT15.Value
ca!ger = face.SBB1.Panels(11).Text
ca!aut = au
ca!com = Label40.Caption
ca.Update
'****** journal
c = sr!ord
jr.AddNew
jr!cre = Text10.Text
jr!deb = "0"
jr!dem = "„‰ Õ‹"
jr!com = "«·„’—Ê›« "
jr!dec = "’—› „»·€ " + Text10.Text + " „‰ √Ã· " + Text9.Text
jr!ord = c
jr!dat = DT15.Value
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
jr.AddNew
jr!cre = "0"
jr!deb = Text10.Text
jr!dem = "≈·Ï Õ‹"
jr!com = "«·’‰œÊﬁ"
jr!dec = "’—› „»·€ " + Text10.Text + " „‰ √Ã· " + Text9.Text
jr!ord = c
jr!dat = DT15.Value
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
dp.AddNew
dp!aut = au
dp!dec = Text9.Text
dp!mon = Text10.Text
dp!dat = DT15.Value
dp!heu = Time$
dp!ger = face.SBB1.Panels(11).Text
If Val(Combo8.Text) > 0 Then
dp!com = Label23.Caption
Else
dp!com = Combo8.Text
End If
dp.Update
ProgressBar6.Value = 0
ProgressBar6.Visible = True
Timer6.Enabled = True

End Sub

Private Sub Command19_Click()
Dim i As Double
Dim n As Double
Dim d As Double
Dim p As Double
Dim r As Double
Dim sd As Double
Dim sp As Double
Dim sr As Double
Dim ane As String
Command19.Enabled = False
Call cont2
Do While Not dps.EOF
dps.Delete
dps.MoveNext
Loop
sd = 0
sp = 0
sr = 0
d = 0
p = 0
r = 0
n = grd19.Rows
For i = 1 To n - 1
dps.AddNew
dps!dat1 = DT16.Value
dps!dat2 = DT17.Value
dps!tos = Label43.Caption
grd19.row = i
grd19.Col = 1
dps!dec = grd19.Text
grd19.Col = 2
dps!mon = grd19.Text
grd19.Col = 3
dps!com = grd19.Text
grd19.Col = 4
dps!dat = grd19.Text
dps.Update
Next i
tim = 6
ProgressBar9.Value = 0
Timer9.Enabled = True

End Sub

Private Sub Command2_Click()
On Error Resume Next
Text1.Text = Trim(Text1.Text)
Text2.Text = Trim(Text2.Text)
If Text1.Text = "" Then
MsgBox "«œŒ· «·ﬂÊœ", vbCritical
Text1.SetFocus
Exit Sub
End If
If Val(Text1.Text) <= 0 Then
MsgBox "«·ﬂÊœ ÌÃ» √‰ ÌﬂÊ‰ —ﬁ„« ’ÕÌÕ«", vbCritical
Text1.SetFocus
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "«œŒ· «·»Ì«‰", vbCritical
Text2.SetFocus
Exit Sub
End If
If Combo6.Text = "" Then
MsgBox "«œŒ·  ’‰Ì› «·ﬂÊœ", vbCritical
Exit Sub
End If
If Combo13.Text = "" And Combo13.Visible = True Then
MsgBox "ﬁ„ »«Œ Ì«— ‰Ê⁄ «·⁄„·Ì…", vbCritical
Exit Sub
End If
Call cont
Do While Not cd.EOF
If Label1.Caption <> cd!aut And cd!cod = Val(Text1.Text) Then
MsgBox " „ «” Œœ«„ Â–« «·ﬂÊœ „‰ ﬁ»·", vbCritical
Exit Sub
End If
cd.MoveNext
Loop
If Label1.Caption <> "" Then
Call cont
Do While Not cd.EOF
If Label1.Caption = cd!aut Then
cd!cod = Val(Text1.Text)
cd!dec = Text2.Text
cd!cas = Combo6.Text
cd!der = ""
If Combo6.Text = "«·»‰ﬂ" Then
cd!der = Combo13.Text
End If
cd.Update
ProgressBar1.Value = 0
Timer1.Enabled = True
Exit Sub
End If
cd.MoveNext
Loop
End If
cd.AddNew
cd!cod = Val(Text1.Text)
cd!dec = Text2.Text
cd!cas = Combo6.Text
cd!der = ""
If Combo6.Text = "«·»‰ﬂ" Then
cd!der = Combo13.Text
End If
cd.Update
ProgressBar1.Value = 0
Timer1.Enabled = True
End Sub

Private Sub Command20_Click()
On Error Resume Next
grd19.Visible = False
grd19.Clear
grd19.Rows = 1
Call chargegrd19_2
grd19.Visible = True
Command19.Enabled = True

End Sub

Private Sub Command21_Click()
On Error Resume Next
Dim a As Double
Dim b As Double
Dim c As Double
Dim au As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim dat4 As Date
Dim annes As String
Dim mca As Double
Dim mnv As Double
Text7.Text = Trim(Text7.Text)
Text8.Text = Trim(Text8.Text)
If Combo7.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— ﬂÊœ «·⁄„·Ì…", vbCritical
Exit Sub
End If
If Text7.Text = "" Then
MsgBox "«œŒ· «·—ﬁ„ «· ”·”·Ì ··√” «–", vbCritical
Text7.SetFocus
Exit Sub
End If
If Label10.Caption = "" Then
MsgBox "«œŒ· «·—ﬁ„ «· ”·”·Ì ··√” «– À„ «÷€ÿ ⁄·Ï Enter", vbCritical + arabic
Text7.SetFocus
Exit Sub
End If
If Text8.Text = "" Then
MsgBox "«Œ· «·„»·€ «·„œ›Ê⁄", vbCritical + arabic
Text8.SetFocus
Exit Sub
End If
Call cont
'**** controle caisse ajouter
mca = sr!cca
mnv = Text8.Text
If mnv > mca Then
MsgBox "€Ì— „„ﬂ‰.. ·√‰Â ·« ÌÊÃœ ›Ì «·’‰œÊﬁ Õ«·Ì« Â–« «·„»·€", vbCritical
Exit Sub
End If
'**** end controle caisse
'**** controle Date Ajouter
annes = sr!ann      'annee scolaire
dat1 = sr!dat       'debut annee scolaire
dat2 = sr!dtf       'Fin annees scolaire
dat3 = DT7.Value    'Date Modifiable
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
'****** caisse
ca.AddNew
au = ca!aun
If Val(Combo7.Text) > 0 Then
ca!cod = Combo7.Text
ca!dec = Label22.Caption
Else
ca!dec = Combo7.Text
ca!cod = Label22.Caption
End If
ca!mem = "œ›⁄ „»·€ " + Text8.Text + " ·’«·Õ «·√” «– " + Label10.Caption + " —ﬁ„ " + Label20.Caption
ca!mon = Text8.Text
ca!cas = "Œ«—Ã"
ca!heu = Time$
ca!dat = DT7.Value
ca!ger = face.SBB1.Panels(11).Text
ca!aut = au
ca!com = Label38.Caption
ca.Update
'****** journal
c = sr!ord
jr.AddNew
jr!cre = Text8.Text
jr!deb = "0"
jr!dem = "„‰ Õ‹"
jr!com = "«·√”« –…"
jr!dec = "œ›⁄ „»·€ " + Text8.Text + " ·’«·Õ «·√” «– " + Label10.Caption + " —ﬁ„ " + Label20.Caption
jr!ord = c
jr!dat = DT7.Value
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
jr.AddNew
jr!cre = "0"
jr!deb = Text8.Text
jr!dem = "≈·Ï Õ‹"
jr!com = "«·’‰œÊﬁ"
jr!dec = "œ›⁄ „»·€ " + Text8.Text + " ·’«·Õ «·√” «– " + Label10.Caption + " —ﬁ„ " + Label20.Caption
jr!ord = c
jr!dat = DT7.Value
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
'****** cascaisse
a = sr!cca
b = Text8.Text
a = a - b
sr!cca = a
sr!ord = c + 1
sr.Update
pf.AddNew
pf!aut = au
pf!ser = Label20.Caption
pf!nom = Label10.Caption
pf!mon = Text8.Text
pf!dat = DT7.Value
pf!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
pf.Update
ProgressBar5.Value = 0
ProgressBar5.Visible = True
Timer4.Enabled = True

End Sub


Private Sub Command22_Click()
On Error Resume Next
Dim a As Double
Dim b As Double
Dim c As Double
Dim au As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim dat4 As Date
Dim annes As String
Dim mca As Double
Dim mnv As Double
Text11.Text = Trim(Text11.Text)
If Label96.Caption = "" Then
MsgBox "«÷€ÿ ⁄·Ï «”„ «·‘—Ìﬂ √Ê·«", vbCritical + arabic
Exit Sub
End If
If Combo9.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— ﬂÊœ «·⁄„·Ì…", vbCritical
Exit Sub
End If
If Text11.Text = "" Then
MsgBox "«œŒ· «·„»·€ «·„œ›Ê⁄", vbCritical
Text11.SetFocus
Exit Sub
End If
If Combo10.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— œ«›⁄ «·„»·€", vbCritical
Exit Sub
End If
Call cont
'**** controle caisse ajouter
If Combo10.Text = "«·’‰œÊﬁ" Then
mca = sr!cca
mnv = Text11.Text
If mnv > mca Then
MsgBox "€Ì— „„ﬂ‰.. ·√‰Â ·« ÌÊÃœ ›Ì «·’‰œÊﬁ Õ«·Ì« Â–« «·„»·€", vbCritical
Exit Sub
End If
End If
'**** end controle caisse
'**** controle Date Ajouter
annes = sr!ann      'annee scolaire
dat1 = sr!dat       'debut annee scolaire
dat2 = sr!dtf       'Fin annees scolaire
dat3 = DT12.Value    'Date Modifiable
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
'****** caisse
ca.AddNew
au = ca!aun
If Val(Combo9.Text) > 0 Then
ca!cod = Combo9.Text
ca!dec = Label27.Caption
Else
ca!dec = Combo9.Text
ca!cod = Label27.Caption
End If
If Combo10.Text = "«·’‰œÊﬁ" Then
ca!mem = "œ›⁄ „»·€ " + Text11.Text + " ·’«·Õ «·‘—Ìﬂ " + Label96.Caption
ca!mon = Text11.Text
ca!cas = "Œ«—Ã"
Else
ca!mem = "«” ·«„ „»·€ " + Text11.Text + " „‰ «·‘—Ìﬂ " + Label96.Caption
ca!mon = Text11.Text
ca!cas = "œ«Œ·"
End If
ca!heu = Time$
ca!dat = DT12.Value
ca!ger = face.SBB1.Panels(11).Text
ca!aut = au
ca!com = Label39.Caption
ca.Update
'****** journal
c = sr!ord
If Combo10.Text = "«·’‰œÊﬁ" Then
jr.AddNew
jr!cre = Text11.Text
jr!deb = "0"
jr!dem = "„‰ Õ‹"
jr!com = "«·‘—ﬂ«¡"
jr!dec = "œ›⁄ „»·€ " + Text11.Text + " ·’«·Õ «·‘—Ìﬂ " + Label96.Caption
jr!ord = c
jr!dat = DT12.Value
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
jr.AddNew
jr!cre = "0"
jr!deb = Text11.Text
jr!dem = "≈·Ï Õ‹"
jr!com = "«·’‰œÊﬁ"
jr!dec = "œ›⁄ „»·€ " + Text11.Text + " ·’«·Õ «·‘—Ìﬂ " + Label96.Caption
jr!ord = c
jr!dat = DT12.Value
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
Else
jr.AddNew
jr!cre = Text11.Text
jr!deb = "0"
jr!dem = "„‰ Õ‹"
jr!com = "«·’‰œÊﬁ"
jr!dec = "«” ·«„ „»·€ " + Text11.Text + " „‰ «·‘—Ìﬂ " + Label96.Caption
jr!ord = c
jr!dat = DT12.Value
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
jr.AddNew
jr!cre = "0"
jr!deb = Text11.Text
jr!dem = "≈·Ï Õ‹"
jr!com = "«·‘—ﬂ«¡"
jr!dec = "«” ·«„ „»·€ " + Text11.Text + " „‰ «·‘—Ìﬂ " + Label96.Caption
jr!ord = c
jr!dat = DT12.Value
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
End If
'****** cascaisse
a = sr!cca
b = Text11.Text
If Combo10.Text = "«·’‰œÊﬁ" Then
a = a - b
Else
a = a + b
End If
sr!cca = a
sr!ord = c + 1
sr.Update
pp.AddNew
pp!aut = au
pp!mtr = Label28.Caption
pp!nom = Label96.Caption
pp!mon = Text11.Text
pp!Mod = Combo10.Text
pp!dat = DT12.Value
pp!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
pp.Update
ProgressBar4.Value = 0
ProgressBar4.Visible = True
Timer5.Enabled = True

End Sub

Private Sub Command23_Click()
On Error Resume Next
Text8.Text = ""
Text8.SetFocus
ProgressBar5.Value = 0
ProgressBar5.Visible = True
Timer4.Enabled = False
End Sub

Private Sub Command24_Click()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim m As Double
Dim s As Double
Dim ane As String
Command24.Enabled = False
Call cont2
Do While Not pe.EOF
pe.Delete
pe.MoveNext
Loop
s = 0
m = 0
n = grd14.Rows
For i = 1 To n - 1
grd14.row = i
grd14.Col = 3
m = grd14.Text
s = s + m
Next i
For i = 1 To n - 1
pe.AddNew
pe!nom = Label10.Caption
pe!ser = Text7.Text
pe!dat1 = DT8.Value
pe!dat2 = DT9.Value
grd14.row = i
grd14.Col = 1
pe!dat = grd14.Text
grd14.Col = 2
pe!heu = grd14.Text
grd14.Col = 3
pe!mon = grd14.Text
pe!tos = s
pe.Update
Next i
tim = 2
ProgressBar9.Value = 0
Timer9.Enabled = True

End Sub

Private Sub Command25_Click()
On Error Resume Next
grd14.Visible = False
grd14.Clear
grd14.Rows = 1
Call chargegrd14_2
grd14.Visible = True
End Sub

Private Sub Command26_Click()
On Error Resume Next
Call chargec10
Text11.Text = ""
Text11.SetFocus
grd17.Visible = False
grd17.Clear
grd17.Rows = 1
grd18.Visible = False
grd18.Clear
grd18.Rows = 1
Call chargegrd17_18
grd17.Visible = True
grd18.Visible = True
ProgressBar4.Value = 0
ProgressBar4.Visible = True
Timer5.Enabled = False

End Sub

Private Sub Command27_Click()
On Error Resume Next
grd17.Visible = False
grd17.Clear
grd17.Rows = 1
grd18.Visible = False
grd18.Clear
grd18.Rows = 1
Call chargegrd17_18_2
grd17.Visible = True
grd18.Visible = True
Command28.Enabled = True
End Sub

Private Sub Command28_Click()
Dim i As Double
Dim n As Double
Dim d As Double
Dim p As Double
Dim r As Double
Dim sd As Double
Dim sp As Double
Dim sr As Double
Dim ane As String
Command28.Enabled = False
Call cont2
Do While Not rpp.EOF
rpp.Delete
rpp.MoveNext
Loop
sd = 0
sp = 0
sr = 0
d = 0
p = 0
r = 0
n = grd17.Rows
For i = 1 To n - 1
rpp.AddNew
rpp!nom = Label96.Caption
rpp!tel = Label95.Caption
rpp!dat1 = DT13.Value
rpp!dat2 = DT14.Value
rpp!mpay = Label24.Caption
rpp!mrec = Label25.Caption
rpp!dtp = ""
rpp!mop = ""
grd17.row = i
grd17.Col = 1
rpp!dtr = grd17.Text
grd17.Col = 2
rpp!mor = grd17.Text
rpp!cas = Label31(52).Caption
rpp!cre = Label26.Caption
rpp.Update
Next i
n = grd18.Rows
For i = 1 To n - 1
rpp.AddNew
rpp!nom = Label96.Caption
rpp!tel = Label95.Caption
rpp!dat1 = DT13.Value
rpp!dat2 = DT14.Value
rpp!mpay = Label24.Caption
rpp!mrec = Label25.Caption
rpp!dtr = ""
rpp!mor = ""
grd18.row = i
grd18.Col = 1
rpp!dtp = grd18.Text
grd18.Col = 2
rpp!mop = grd18.Text
rpp!cas = Label31(52).Caption
rpp!cre = Label26.Caption
rpp.Update
Next i
tim = 5
ProgressBar9.Value = 0
Timer9.Enabled = True

End Sub

Private Sub Command29_Click()
On Error Resume Next
grd20.Visible = False
grd20.Clear
grd20.Rows = 1
grd21.Visible = False
grd21.Clear
grd21.Rows = 1
Call chargegrd20_21_2
grd20.Visible = True
grd21.Visible = True
Command30.Enabled = True

End Sub

Private Sub Command3_Click()
On Error Resume Next
Text2.Text = ""
Label1.Caption = ""
Text1.Text = ""
Text1.SetFocus
ProgressBar1.Value = 0
Timer1.Enabled = False

End Sub

Private Sub Command30_Click()
Dim i As Double
Dim n As Double
Dim d As Double
Dim p As Double
Dim r As Double
Dim sd As Double
Dim sp As Double
Dim sr As Double
Dim ane As String
Command30.Enabled = False
Call cont2
Do While Not bnk.EOF
bnk.Delete
bnk.MoveNext
Loop
sd = 0
sp = 0
sr = 0
d = 0
p = 0
r = 0
n = grd21.Rows
For i = 1 To n - 1
bnk.AddNew
bnk!dat1 = DT19.Value
bnk!dat2 = DT20.Value
bnk!dep = Label33.Caption
bnk!Ret = Label32.Caption
bnk!cre = Label30.Caption
bnk!der = ""
bnk!mor = ""
bnk!dar = ""
grd21.row = i
grd21.Col = 1
bnk!ded = grd21.Text
grd21.Col = 2
bnk!Mod = grd21.Text
grd21.Col = 3
bnk!dad = grd21.Text
bnk.Update
Next i
n = grd20.Rows
For i = 1 To n - 1
bnk.AddNew
bnk!dat1 = DT19.Value
bnk!dat2 = DT20.Value
bnk!dep = Label33.Caption
bnk!Ret = Label32.Caption
bnk!cre = Label30.Caption
bnk!ded = ""
bnk!Mod = ""
bnk!dad = ""
grd20.row = i
grd20.Col = 1
bnk!der = grd20.Text
grd20.Col = 2
bnk!mor = grd20.Text
grd20.Col = 3
bnk!dar = grd20.Text
bnk.Update
Next i
tim = 7
ProgressBar9.Value = 0
Timer9.Enabled = True

End Sub

Private Sub Command31_Click()
On Error Resume Next
Dim a As Double
Dim b As Double
Dim c As Double
Dim au As Double
Dim d As Double
Dim e As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim dat4 As Date
Dim annes As String
Dim mca As Double
Dim mnv As Double
Text13.Text = Trim(Text13.Text)
Text12.Text = Trim(Text12.Text)
If Combo11.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— ﬂÊœ «·⁄„·Ì…", vbCritical
Exit Sub
End If
If Text13.Text = "" Then
MsgBox "«œŒ· »Ì«‰ «·⁄„·Ì…", vbCritical
Text13.SetFocus
Exit Sub
End If
If Combo12.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— ‰Ê⁄ «·⁄„·Ì…", vbCritical
Exit Sub
End If
If Text12.Text = "" Then
MsgBox "«œŒ· „»·€ «·" + Combo12.Text, vbCritical
Text12.SetFocus
Exit Sub
End If
d = Label30.Caption
e = Text12.Text
If e > d And Combo12.Text = "”Õ»" Then
MsgBox "€Ì— „„ﬂ‰.. ·«ÌÊÃœ ›Ì —’Ìœ «·»‰ﬂ ”ÊÏ " + Label30.Caption, vbCritical
Text12.SetFocus
Exit Sub
End If
Call cont
'**** controle caisse ajouter
If Combo12.Text = "«Ìœ«⁄" Then
mca = sr!cca
mnv = Text12.Text
If mnv > mca Then
MsgBox "€Ì— „„ﬂ‰.. ·√‰Â ·« ÌÊÃœ ›Ì «·’‰œÊﬁ Õ«·Ì« Â–« «·„»·€", vbCritical
Exit Sub
End If
End If
'**** end controle caisse
'**** controle Date Ajouter
annes = sr!ann      'annee scolaire
dat1 = sr!dat       'debut annee scolaire
dat2 = sr!dtf       'Fin annees scolaire
dat3 = DT18.Value    'Date Modifiable
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
'**** if capital
If Combo12.Text = "—√” «·„«·" Then
Call cont
'****** caisse
ca.AddNew
au = ca!aun
If Val(Combo11.Text) > 0 Then
ca!cod = Combo11.Text
ca!dec = Label29.Caption
Else
ca!dec = Combo11.Text
ca!cod = Label29.Caption
End If
ca!mon = "0"
ca!mem = "œ›⁄ „»·€ " + Text12.Text + " ≈·Ï «·»‰ﬂ ﬂ—√” „«· ,  Õ  «·»Ì«‰ " + Text13.Text
ca!cas = "Œ«—Ã"
ca!heu = Time$
ca!dat = DT18.Value
ca!ger = face.SBB1.Panels(11).Text
ca!aut = au
ca!com = Label41.Caption
ca.Update
'****** journal
c = sr!ord
jr.AddNew
jr!cre = Text12.Text
jr!deb = "0"
jr!dem = "„‰ Õ‹"
jr!com = "«·»‰ﬂ"
jr!dec = "œ›⁄ „»·€ " + Text12.Text + " ≈·Ï «·»‰ﬂ ﬂ—√” „«· ,  Õ  «·»Ì«‰ " + Text13.Text
jr!ord = c
jr!dat = DT18.Value
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
jr.AddNew
jr!cre = "0"
jr!deb = Text12.Text
jr!dem = "≈·Ï Õ‹"
jr!com = "—√” «·„«·"
jr!dec = "œ›⁄ „»·€ " + Text12.Text + " ≈·Ï «·»‰ﬂ ﬂ—√” „«· ,  Õ  «·»Ì«‰ " + Text13.Text
jr!ord = c
jr!dat = DT18.Value
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
sr!ord = c + 1
sr.Update
bn.AddNew
bn!aut = au
bn!dec = Text13.Text
bn!Mod = Combo12.Text
bn!mon = Text12.Text
bn!dat = DT18.Value
bn!heu = Time$
bn!ger = face.SBB1.Panels(11).Text
bn.Update
ProgressBar7.Value = 0
ProgressBar7.Visible = True
Timer7.Enabled = True
Exit Sub
End If
'**** if not capital
Call cont
'****** caisse
ca.AddNew
au = ca!aun
If Val(Combo11.Text) > 0 Then
ca!cod = Combo11.Text
ca!dec = Label29.Caption
Else
ca!dec = Combo11.Text
ca!cod = Label29.Caption
End If
ca!mon = Text12.Text
If Combo12.Text = "”Õ»" Then
ca!mem = "”Õ» „»·€ " + Text12.Text + " „‰ «·»‰ﬂ ,  Õ  «·»Ì«‰ " + Text13.Text
ca!cas = "œ«Œ·"
Else
ca!mem = "«Ìœ«⁄ „»·€ " + Text12.Text + " ›Ì «·»‰ﬂ ,  Õ  «·»Ì«‰ " + Text13.Text
ca!cas = "Œ«—Ã"
End If
ca!heu = Time$
ca!dat = DT18.Value
ca!ger = face.SBB1.Panels(11).Text
ca!aut = au
ca!com = Label41.Caption
ca.Update
'****** journal
c = sr!ord
If Combo12.Text = "”Õ»" Then
jr.AddNew
jr!cre = Text12.Text
jr!deb = "0"
jr!dem = "„‰ Õ‹"
jr!com = "«·’‰œÊﬁ"
jr!dec = "”Õ» „»·€ " + Text12.Text + " „‰ «·»‰ﬂ ,  Õ  «·»Ì«‰ " + Text13.Text
jr!ord = c
jr!dat = DT18.Value
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
jr.AddNew
jr!cre = "0"
jr!deb = Text12.Text
jr!dem = "≈·Ï Õ‹"
jr!com = "«·»‰ﬂ"
jr!dec = "”Õ» „»·€ " + Text12.Text + " „‰ «·»‰ﬂ ,  Õ  «·»Ì«‰ " + Text13.Text
jr!ord = c
jr!dat = DT18.Value
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
Else
jr.AddNew
jr!cre = Text12.Text
jr!deb = "0"
jr!dem = "„‰ Õ‹"
jr!com = "«·»‰ﬂ"
jr!dec = "«Ìœ«⁄ „»·€ " + Text12.Text + " ›Ì «·»‰ﬂ ,  Õ  «·»Ì«‰ " + Text13.Text
jr!ord = c
jr!dat = DT18.Value
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
jr.AddNew
jr!cre = "0"
jr!deb = Text12.Text
jr!dem = "≈·Ï Õ‹"
jr!com = "«·’‰œÊﬁ"
jr!dec = "«Ìœ«⁄ „»·€ " + Text12.Text + " ›Ì «·»‰ﬂ ,  Õ  «·»Ì«‰ " + Text13.Text
jr!ord = c
jr!dat = DT18.Value
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
End If
'****** cascaisse
a = sr!cca
b = Text12.Text
If Combo12.Text = "”Õ»" Then
a = a + b
Else
a = a - b
End If
sr!cca = a
sr!ord = c + 1
sr.Update
bn.AddNew
bn!aut = au
bn!dec = Text13.Text
bn!Mod = Combo12.Text
bn!mon = Text12.Text
bn!dat = DT18.Value
bn!heu = Time$
bn!ger = face.SBB1.Panels(11).Text
bn.Update
ProgressBar7.Value = 0
ProgressBar7.Visible = True
Timer7.Enabled = True

End Sub

Private Sub Command32_Click()
Dim i As Double
Dim n As Double
Dim d As Double
Dim p As Double
Dim r As Double
Dim sd As Double
Dim sp As Double
Dim sr As Double
Dim ane As String
Command32.Enabled = False
Call cont2
Do While Not cai.EOF
cai.Delete
cai.MoveNext
Loop
sd = 0
sp = 0
sr = 0
d = 0
p = 0
r = 0
n = grd22.Rows
For i = 1 To n - 1
cai.AddNew
cai!crp = Label42.Caption
cai!cra = Label36.Caption
cai!moe = Label34.Caption
cai!mos = Label35.Caption
cai!dif = Label44.Caption
cai!dat1 = DT21.Value
cai!dat2 = DT22.Value
grd22.row = i
grd22.Col = 1
cai!dec = grd22.Text
grd22.Col = 2
cai!mon = grd22.Text
grd22.Col = 6
cai!dat = grd22.Text
grd22.Col = 7
cai!heu = grd22.Text
grd22.Col = 8
cai!ger = grd22.Text
cai.Update
Next i
tim = 8
ProgressBar9.Value = 0
Timer9.Enabled = True

End Sub

Private Sub Command33_Click()
On Error Resume Next
grd22.Visible = False
grd22.Clear
grd22.Rows = 1
grd23.Visible = False
grd23.Clear
grd23.Rows = 1
Call chargegrd22_23
grd22.Visible = True
grd23.Visible = True
Call monauj
Label5(26).Caption = "«·Õ’Ì·…"
End Sub

Private Sub Command34_Click()
On Error Resume Next
DT21.Value = Date
DT22.Value = Date
grd22.Visible = False
grd22.Clear
grd22.Rows = 1
grd23.Visible = False
grd23.Clear
grd23.Rows = 1
Call chargegrd22_23
grd22.Visible = True
grd23.Visible = True
Call monauj
Label5(26).Caption = "Õ’Ì·… «·ÌÊ„"

End Sub

Private Sub Command35_Click()
On Error Resume Next
Dim a As Integer
DT22.Value = Date
a = DT22.Day
a = a - 1
DT21.Value = Date - a
grd22.Visible = False
grd22.Clear
grd22.Rows = 1
grd23.Visible = False
grd23.Clear
grd23.Rows = 1
Call chargegrd22_23
grd22.Visible = True
grd23.Visible = True
Call monauj
Label5(26).Caption = "Õ’Ì·… «·‘Â—"

End Sub

Private Sub Command36_Click()
On Error Resume Next
Call cont
DT21.Value = sr!dat
DT22.Value = sr!dtf
grd22.Visible = False
grd22.Clear
grd22.Rows = 1
grd23.Visible = False
grd23.Clear
grd23.Rows = 1
Call chargegrd22_23
grd22.Visible = True
grd23.Visible = True
Call monauj
Label5(26).Caption = "Õ’Ì·… «·”‰…"

End Sub

Private Sub Command37_Click()
On Error Resume Next
Dim i As Double
Dim a As Double
grd3.Clear
grd3.Rows = 13
i = 1
Call cont
Do While Not ev.EOF
If Val(Text3.Text) = Val(ev!ser) Then
'*** oct
a = 0
If ev!Oct <> "" Then
If Val(ev!Oct) > 0 Then
grd3.row = i
grd3.Col = 0
grd3.Text = "«ﬂ Ê»—"
grd3.Col = 1
grd3.Text = Text5.Text
grd3.Col = 2
grd3.Text = Val(ev!Oct)
a = Val(Text5.Text) - Val(ev!Oct)
grd3.Col = 3
grd3.Text = a
grd3.Col = 4
grd3.Text = 10
i = i + 1
End If
End If
'*** fra
If ev!fra <> "" Then
a = 0
If Val(ev!fra) > 0 Then
grd3.row = i
grd3.Col = 0
grd3.Text = "«· ”ÃÌ·"
grd3.Col = 1
grd3.Text = Text4.Text
grd3.Col = 2
grd3.Text = Val(ev!fra)
a = Val(Text4.Text) - Val(ev!fra)
grd3.Col = 3
grd3.Text = a
grd3.Col = 4
grd3.Text = DT2.Month
i = i + 1
End If
End If
'*** nov
a = 0
If ev!nov <> "" Then
If Val(ev!nov) > 0 Then
grd3.row = i
grd3.Col = 0
grd3.Text = "‰Ê›„»—"
grd3.Col = 1
grd3.Text = Text5.Text
grd3.Col = 2
grd3.Text = Val(ev!nov)
a = Val(Text5.Text) - Val(ev!nov)
grd3.Col = 3
grd3.Text = a
grd3.Col = 4
grd3.Text = 11
i = i + 1
End If
End If
'*** dec
a = 0
If ev!dec <> "" Then
If Val(ev!dec) > 0 Then
grd3.row = i
grd3.Col = 0
grd3.Text = "œÌ”„»—"
grd3.Col = 1
grd3.Text = Text5.Text
grd3.Col = 2
grd3.Text = Val(ev!dec)
a = Val(Text5.Text) - Val(ev!dec)
grd3.Col = 3
grd3.Text = a
grd3.Col = 4
grd3.Text = 12
i = i + 1
End If
End If
'*** jan
a = 0
If ev!jan <> "" Then
If Val(ev!jan) > 0 Then
grd3.row = i
grd3.Col = 0
grd3.Text = "Ì‰«Ì—"
grd3.Col = 1
grd3.Text = Text5.Text
grd3.Col = 2
grd3.Text = Val(ev!jan)
a = Val(Text5.Text) - Val(ev!jan)
grd3.Col = 3
grd3.Text = a
grd3.Col = 4
grd3.Text = 1
i = i + 1
End If
End If
'*** fev
a = 0
If ev!fev <> "" Then
If Val(ev!fev) > 0 Then
grd3.row = i
grd3.Col = 0
grd3.Text = "›»—«Ì—"
grd3.Col = 1
grd3.Text = Text5.Text
grd3.Col = 2
grd3.Text = Val(ev!fev)
a = Val(Text5.Text) - Val(ev!fev)
grd3.Col = 3
grd3.Text = a
grd3.Col = 4
grd3.Text = 2
i = i + 1
End If
End If
'*** mar
a = 0
If ev!mar <> "" Then
If Val(ev!mar) > 0 Then
grd3.row = i
grd3.Col = 0
grd3.Text = "„«—”"
grd3.Col = 1
grd3.Text = Text5.Text
grd3.Col = 2
grd3.Text = Val(ev!mar)
a = Val(Text5.Text) - Val(ev!mar)
grd3.Col = 3
grd3.Text = a
grd3.Col = 4
grd3.Text = 3
i = i + 1
End If
End If
'*** avr
a = 0
If ev!avr <> "" Then
If Val(ev!avr) > 0 Then
grd3.row = i
grd3.Col = 0
grd3.Text = "«»—Ì·"
grd3.Col = 1
grd3.Text = Text5.Text
grd3.Col = 2
grd3.Text = Val(ev!avr)
a = Val(Text5.Text) - Val(ev!avr)
grd3.Col = 3
grd3.Text = a
grd3.Col = 4
grd3.Text = 4
i = i + 1
End If
End If
'*** mai
a = 0
If ev!mai <> "" Then
If Val(ev!mai) > 0 Then
grd3.row = i
grd3.Col = 0
grd3.Text = "„«ÌÊ"
grd3.Col = 1
grd3.Text = Text5.Text
grd3.Col = 2
grd3.Text = Val(ev!mai)
a = Val(Text5.Text) - Val(ev!mai)
grd3.Col = 3
grd3.Text = a
grd3.Col = 4
grd3.Text = 5
i = i + 1
End If
End If
'*** jun
a = 0
If ev!jun <> "" Then
If Val(ev!jun) > 0 Then
grd3.row = i
grd3.Col = 0
grd3.Text = "ÌÊ‰ÌÊ"
grd3.Col = 1
grd3.Text = Text5.Text
grd3.Col = 2
grd3.Text = Val(ev!jun)
a = Val(Text5.Text) - Val(ev!jun)
grd3.Col = 3
grd3.Text = a
grd3.Col = 4
grd3.Text = 6
i = i + 1
End If
End If
End If
ev.MoveNext
Loop
grd3.Rows = i
Call totalmontant
End Sub

Private Sub Command38_Click()
On Error Resume Next
If Text14.Text = Text3.Text Then
MsgBox "OK", vbInformation
Exit Sub
End If
Command4_Click
Command5_Click
Combo2.Text = "Õ«·… ≈ﬂ„«·"
Command37_Click
Command6_Click
Text3.Text = Val(Text3.Text) + 1
Command38_Click
End Sub


Private Sub Command4_Click()
On Error Resume Next
grd1.Clear
grd1.Rows = 1
grd1.Cols = 4
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1200
grd1.ColWidth(2) = 1200
grd1.ColWidth(3) = 3800
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
Call cont
Do While Not et.EOF
If Val(et!num) < 1000000 Then
If Val(et!act) = 1 Then
If Text3.Text = et!ser Or Val(et!ser) = Text3.Text Then
Label3.Caption = et!cla
Label48.Caption = et!tel
grd1.row = 0
grd1.Col = 1
grd1.Text = et!cla
grd1.Col = 2
grd1.Text = et!num
grd1.Col = 3
grd1.Text = et!nom
Label4.Caption = et!ser
Exit Sub
End If
End If
End If
et.MoveNext
Loop
MsgBox "«·—ﬁ„ «· ”·”·Ì «·„œŒ· €Ì— „Œ“‰ .. Ì—ÃÏ «· √ﬂœ „‰Â", vbExclamation
Text3.Text = ""
Text3.SetFocus
End Sub

Private Sub Command40_Click()
Dim i As Double
Dim n As Double
Dim r As Double
Dim p As Double
Dim m As Double
Text3.Text = Val(Text3.Text) + 1
Command4_Click
Command5_Click
p = 0
r = 0
m = 0
n = grd4.Rows
Label75.Caption = ""
For i = 1 To n - 1
grd4.row = i
grd4.Col = 1
If i = 1 Then
Label75.Caption = grd4.Text
Else
Label75.Caption = Label75.Caption + " + " + grd4.Text
End If
grd4.Col = 3
m = grd4.Text
p = p + m
grd4.Col = 4
m = grd4.Text
r = r + m
Next i
Label2.Caption = p
Label45.Caption = r
End Sub

Private Sub Command41_Click()
Command41.Enabled = False
'****** recu pour imprimer
Call cont2
ru.AddNew
ru!rec = Combo3.Text
ru!ser = Label4.Caption
grd1.row = 0
grd1.Col = 1
ru!cla = grd1.Text
grd1.Col = 2
ru!num = grd1.Text
grd1.Col = 3
ru!nom = grd1.Text
ru!mon = Label2.Caption
ru!mois = Label75.Caption
ru!dat = Label46.Caption
ru!cas = Label47.Caption
ru!ann = face.SBB1.Panels(9).Text
ru!eco = face.SBB1.Panels(13).Text
ru!res = Label45.Caption
ru.Update
MsgBox "OK"
End Sub

Private Sub Command42_Click()
Dim i As Double
Dim n As Double
Dim r As Double
Dim p As Double
Dim m As Double
Call chargegrd4_rec
p = 0
r = 0
m = 0
n = grd4.Rows
Label75.Caption = ""
For i = 1 To n - 1
grd4.row = i
grd4.Col = 1
If i = 1 Then
Label75.Caption = grd4.Text
Else
Label75.Caption = Label75.Caption + " + " + grd4.Text
End If
grd4.Col = 3
m = grd4.Text
p = p + m
grd4.Col = 4
m = grd4.Text
r = r + m
Next i
Label2.Caption = p
Label45.Caption = r
Command41.Enabled = True

End Sub

Private Sub Command43_Click()
On Error Resume Next
Dim r As Double
Dim i As Double
Dim rcu As String
Dim mnt As String
Dim dps As String
Dim dte As String
Dim Security As SECURITY_ATTRIBUTES
Dim x$
x$ = ""
x$ = Dir$(App.Path & "\Recus")
If x$ = "" Then
'Create a directory dossier POINTAGES
Ret& = CreateDirectory(App.Path & "\Recus", Security)
End If
x$ = Dir$(App.Path & "\Recus\Recus.txt")
If x$ <> "" Then
Kill App.Path & "\Recus\Recus.txt"
End If
r = grd7.Rows
Open (App.Path & "\Recus\Recus.txt") For Append As #1
Print #1, "ReÁu |  |  Montant |  |     Code       |  |   Date"
For i = 1 To r - 1
grd7.row = i
grd7.Col = 1
rcu = grd7.Text
grd7.Col = 2
mnt = grd7.Text
grd7.Col = 3
dps = grd7.Text
grd7.Col = 4
dte = grd7.Text
Print #1, rcu + " |  |     " + mnt + "    |  |     " + dps + "  |  |     " + dte
Next i
Close #1
Shell "notepad.exe" & " " & App.Path & "\Recus\Recus.txt", vbNormalFocus
End Sub

Private Sub Command44_Click()
On Error GoTo u
Dim i As Double
Dim j As Double
Dim n As Double
Dim k As Double
Dim d As Double
Dim sd As Double
Dim nom As String
Dim cla As String
Dim du As Double
Dim py As Double
Dim rs As Double

If grd4.Rows < 2 Then
MsgBox "·«  ÊÃœ »Ì«‰« ", vbCritical
Exit Sub
End If
grd1.row = 0
grd1.Col = 1
cla = grd1.Text
grd1.Col = 3
nom = grd1.Text
FileCopy App.Path & "\HC010.xls", App.Path & "\Historique de compte.xls"
Command44.Enabled = False
n = grd4.Rows
Set kb = CreateObject("Excel.application")
kb.Workbooks.Open (App.Path & "\Historique de compte.xls")
kb.Visible = True
du = 0
py = 0
rs = 0
For i = 0 To n - 1
For j = 1 To 6
grd4.row = i
grd4.Col = j
k = 7 - j
kb.Workbooks("Historique de compte").Sheets(1).Cells(i + 7, k).Value = grd4.Text
If k = 5 Then
du = du + Val(grd4.Text)
End If
If k = 4 Then
py = py + Val(grd4.Text)
End If
If k = 3 Then
rs = rs + Val(grd4.Text)
End If
Next j
Next i
k = i + 7
kb.Workbooks("Historique de compte").Sheets(1).Cells(k, 6).Value = "«·„Ã„Ê⁄"
kb.Workbooks("Historique de compte").Sheets(1).Cells(k, 5).Value = du
kb.Workbooks("Historique de compte").Sheets(1).Cells(k, 4).Value = py
kb.Workbooks("Historique de compte").Sheets(1).Cells(k, 3).Value = rs

kb.Workbooks("Historique de compte").Sheets(1).Range("D3").Value = face.SBB1.Panels(13).Text
kb.Workbooks("Historique de compte").Sheets(1).Range("A3").Value = face.SBB1.Panels(9).Text
kb.Workbooks("Historique de compte").Sheets(1).Range("D5").Value = nom
kb.Workbooks("Historique de compte").Sheets(1).Range("A5").Value = cla

'kb.Workbooks("Historique de compte").Sheets(1).Cells(k + 2, 2).Value = "«·≈œ«—…"

'kb.Workbooks("fiche de presences").Sheets(1).Range("B5").Value = DT11.Value
'Workbooks("Etudiants").Close savechanges:=False
'Worksheets(1).Activate
Command44.Enabled = True
Exit Sub
u:
MsgBox "ÌÃ» «€·«ﬁ ’›Õ… «ﬂ”· «‰ ﬂ«‰  „› ÊÕ…", vbExclamation
Command44.Enabled = True

End Sub

Private Sub Command45_Click()
On Error GoTo u
Dim tel As String
tel = Label48.Caption
grd80.Visible = False
Call chargegrd80
grd80.Visible = True
If grd80.Rows < 2 Then
MsgBox "·«  ÊÃœ »Ì«‰« ", vbCritical
Exit Sub
End If
FileCopy App.Path & "\DU010.xls", App.Path & "\Dus de compte.xls"
Command45.Enabled = False
n = grd80.Rows
Set kb = CreateObject("Excel.application")
kb.Workbooks.Open (App.Path & "\Dus de compte.xls")
kb.Visible = True
For i = 1 To n - 1
For j = 1 To 11
grd80.row = i
grd80.Col = j
k = 12 - j
kb.Workbooks("Dus de compte").Sheets(1).Cells(i + 10, k).Value = grd80.Text
Next j
Next i
kb.Workbooks("Dus de compte").Sheets(1).Range("H3").Value = face.SBB1.Panels(13).Text
kb.Workbooks("Dus de compte").Sheets(1).Range("B3").Value = face.SBB1.Panels(9).Text
kb.Workbooks("Dus de compte").Sheets(1).Range("H5").Value = tel
Command45.Enabled = True
Exit Sub
u:
MsgBox "ÌÃ» «€·«ﬁ ’›Õ… «ﬂ”· «‰ ﬂ«‰  „› ÊÕ…", vbExclamation
Command45.Enabled = True

End Sub

Private Sub Command5_Click()
On Error Resume Next
If Text3.Text = "" Then
MsgBox "«œŒ· «·—ﬁ„ «· ”·”·Ì ·· ·„Ì–", vbCritical
Text3.SetFocus
Exit Sub
End If
If Label3.Caption = "" Then
MsgBox "«œŒ· «·—ﬁ„ «· ”·”·Ì ·· ·„Ì–À„ «÷€ÿ ⁄·Ï Enter", vbCritical + arabic
Text3.SetFocus
Exit Sub
End If
grd1.Visible = False
Call chargec2
Call chargec3
grd4.Clear
grd4.Rows = 1
grd3.Clear
grd3.Rows = 1
grd2.Visible = False
grd3.Visible = False
grd4.Visible = False
grd5.Visible = False
Label2.Caption = ""
Label75.Caption = ""
Call chargegrd4
Call chargegrd2
Call chargegrd5
grd2.Visible = True
grd3.Visible = True
grd4.Visible = True
grd5.Visible = True
Picture6.Visible = True
grd1.Visible = True
Text19.Text = ""
Call cont2
Do While Not ru.EOF
If Text3.Text = ru!ser Or Val(ru!ser) = Text3.Text Then
Text19.Text = ru!obs
'Exit Sub
End If
ru.MoveNext
Loop

End Sub

Private Sub Command6_Click()
On Error Resume Next
Dim n As Double
Dim i As Double
Dim a As Double
Dim b As Double
Dim c As Double
Dim du1 As Double
Dim du2 As Double
Dim py As Double
Dim rs As Double
Dim au As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim dat4 As Date
Dim clas As String
Dim numero As String
Dim annes As String
Text3.Text = Trim(Text3.Text)
If Combo1.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— ﬂÊœ «·⁄„·Ì…", vbCritical
Exit Sub
End If
If Text3.Text = "" Then
MsgBox "«œŒ· «·—ﬁ„ «· ”·”·Ì ·· ·„Ì–", vbCritical
Text3.SetFocus
Exit Sub
End If
If Combo2.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·Õ«·… «· ””ÃÌ·Ì… ·· ·„Ì–", vbCritical
Text3.SetFocus
Exit Sub
End If
If grd3.Rows < 2 Then
MsgBox "·« ÊÃœ ⁄„·Ì«  ﬁÌœ «· ‰›Ì–", vbCritical
Exit Sub
End If
Text6.Text = Val(Text6.Text)
n = grd3.Rows
Call cont
Do While Not rc.EOF
If Text6.Text = rc!rec Then
MsgBox "—ﬁ„ «·Ê’· «·„œŒ·  „ ÕÃ“Â ”«»ﬁ«", vbCritical
Exit Sub
End If
rc.MoveNext
Loop
'**** controle Date Ajouter
annes = sr!ann      'annee scolaire
dat1 = sr!dat       'debut annee scolaire
dat2 = sr!dtf       'Fin annees scolaire
dat3 = DT2.Value    'Date Modifiable
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
Command6.Enabled = False
Command7.Enabled = False
Command9.Enabled = False
grd1.row = 0
grd1.Col = 1
clas = grd1.Text
grd1.Col = 2
numero = grd1.Text
'****** caisse
ca.AddNew
au = ca!aun
If Val(Combo1.Text) > 0 Then
ca!cod = Combo1.Text
ca!dec = Label21.Caption
Else
ca!dec = Combo1.Text
ca!cod = Label21.Caption
End If
ca!mem = "œ›⁄ „»·€ " + Label2.Caption + " „‰ ÿ—› " + Label4.Caption + "_" + clas + "_" + numero
ca!mon = Label2.Caption
ca!cas = "œ«Œ·"
ca!heu = Time$
ca!dat = DT2.Value
ca!ger = face.SBB1.Panels(11).Text
ca!aut = au
ca!com = Label37.Caption
ca.Update
'****** journal
c = sr!ord
jr.AddNew
jr!cre = Label2.Caption
jr!deb = "0"
jr!dem = "„‰ Õ‹"
jr!com = "«·’‰œÊﬁ"
jr!dec = "œ›⁄ „»·€ " + Label2.Caption + " „‰ ÿ—› " + Label4.Caption + "_" + clas + "_" + numero
jr!ord = c
jr!dat = DT2.Value
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
jr.AddNew
jr!cre = "0"
jr!deb = Label2.Caption
jr!dem = "≈·Ï Õ‹"
jr!com = "«· ·«„Ì–"
jr!dec = "œ›⁄ „»·€ " + Label2.Caption + " „‰ ÿ—› " + Label4.Caption + "_" + clas + "_" + numero
jr!ord = c
jr!dat = DT2.Value
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
'****** cascaisse
a = sr!cca
b = Label2.Caption
a = a + b
sr!cca = a
sr!ord = c + 1
sr!rec = Val(Text6.Text) + 1
sr.Update
'****** recu
rc.AddNew
rc!aut = au
rc!rec = Text6.Text
rc!ser = Label4.Caption
grd1.row = 0
grd1.Col = 1
rc!cla = grd1.Text
grd1.Col = 2
rc!num = grd1.Text
grd1.Col = 3
rc!nom = grd1.Text
rc!mon = Label2.Caption
rc!mois = Label75.Caption
rc!dat = DT2.Value
rc!cas = Combo2.Text
rc.Update
'******* compte etudiant
For i = 1 To n - 1
ce.AddNew
ce!ser = Label4.Caption
grd1.row = 0
grd1.Col = 1
ce!cla = grd1.Text
grd1.Col = 2
ce!num = grd1.Text
grd1.Col = 3
ce!nom = grd1.Text
ce!fra = Text4.Text
ce!man = Text5.Text
ce!cas = Combo2.Text
ce!rec = Text6.Text
grd3.row = i
grd3.Col = 0
ce!moi = grd3.Text
If grd3.Text = "«· ”ÃÌ·" Then
du1 = Text4.Text
Else
du1 = Text5.Text
End If
grd3.Col = 1
du2 = grd3.Text
grd3.Col = 2
py = grd3.Text
grd3.Col = 3
rs = grd3.Text
If du1 = du2 Then
ce!mon = du2
ce!pay = py
ce!res = rs
Else
ce!mon = 0
ce!pay = py
ce!res = -py
End If
grd3.Col = 4
ce!mois = grd3.Text
ce!dat = DT2.Value
ce.Update
Next i
'****** recu pour imprimer
Call cont2
ru.AddNew
ru!rec = Text6.Text
ru!ser = Label4.Caption
grd1.row = 0
grd1.Col = 1
ru!cla = grd1.Text
grd1.Col = 2
ru!num = grd1.Text
grd1.Col = 3
ru!nom = grd1.Text
ru!mon = Label2.Caption
ru!mois = Label75.Caption
ru!dat = DT2.Value
ru!cas = Combo2.Text
ru!ann = face.SBB1.Panels(9).Text
ru!eco = face.SBB1.Panels(13).Text
ru!res = Label45.Caption
ru!tel = face.SBB1.Panels(11).Text
ru!obs = Text19.Text
ru.Update
Text6.Text = Val(Text6.Text) + 1
Command6.Enabled = True
Command7.Enabled = True
Command9.Enabled = True
ProgressBar2.Value = 0
ProgressBar2.Visible = True
Timer2.Enabled = True

End Sub

Private Sub Command7_Click()
On Error Resume Next
grd3.Clear
grd3.Rows = 1
grd3.ColWidth(0) = 1100
grd3.ColWidth(1) = 1100
grd3.ColWidth(2) = 1100
grd3.ColWidth(3) = 1100
grd3.ColAlignment(0) = 1
grd3.ColAlignment(1) = 1
grd3.ColAlignment(2) = 1
grd3.ColAlignment(3) = 1
grd3.row = 0
grd3.Col = 0
grd3.Text = "«·‘Â—"
grd3.Col = 1
grd3.Text = "«·„” Õﬁ"
grd3.Col = 2
grd3.Text = "«·„œ›Ê⁄"
grd3.Col = 3
grd3.Text = "«·»«ﬁÌ"
Label2.Caption = ""
Label45.Caption = ""
Label75.Caption = ""

End Sub

Private Sub Command8_Click()
On Error Resume Next
If Text4.Text = "" Then
MsgBox "ÌÃ» «œŒ«· —”Ê„ «· ”ÃÌ·", vbCritical
Text4.SetFocus
Exit Sub
End If
If Text5.Text = "" Then
MsgBox "ÌÃ» «œŒ«· «·—”Ê„ «·‘Â—Ì…", vbCritical
Text5.SetFocus
Exit Sub
End If
If Combo2.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·Õ«·… «· ””ÃÌ·Ì… ·· ·„Ì–", vbCritical
Text3.SetFocus
Exit Sub
End If
grd2.Visible = False
grd3.Visible = False
grd4.Visible = False
grd5.Visible = False
Call chargegrd2
grd2.Visible = True
grd3.Visible = True
grd4.Visible = True
grd5.Visible = True

End Sub

Private Sub Command9_Click()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim m As Double
Dim tx1 As String
Dim a As Double
n = grd2.Rows
grd3.Rows = n + 1
For i = 1 To n - 1
grd2.row = i
grd2.Col = 0
tx1 = grd2.Text
grd2.Col = 1
a = grd2.Text
grd2.Col = 2
m = grd2.Text
grd3.row = i
grd3.Col = 0
grd3.Text = tx1
grd3.Col = 1
grd3.Text = a
grd3.Col = 2
grd3.Text = a
grd3.Col = 3
grd3.Text = "0"
grd3.Col = 4
grd3.Text = m
Next i
grd3.Rows = i
Call totalmontant
End Sub



Private Sub DT10_Change()
On Error Resume Next
grd15.Clear
grd15.Rows = 1

End Sub

Private Sub DT10_Click()
On Error Resume Next
DT10_Change
End Sub

Private Sub DT11_Change()
On Error Resume Next
grd15.Clear
grd15.Rows = 1

End Sub

Private Sub DT11_Click()
On Error Resume Next
DT11_Change
End Sub

Private Sub DT3_Change()
On Error Resume Next
Label6.Caption = "0"
Label7.Caption = "0"
Label8.Caption = "0"
Label9.Caption = "0"
grd7.Clear
grd7.Rows = 1
grd8.Clear
grd8.Rows = 1
grd9.Clear
grd9.Rows = 1
End Sub

Private Sub DT3_Click()
On Error Resume Next
DT3_Change
End Sub


Private Sub DT4_Change()
On Error Resume Next
Label6.Caption = "0"
Label7.Caption = "0"
Label8.Caption = "0"
Label9.Caption = "0"
grd7.Clear
grd7.Rows = 1
grd8.Clear
grd8.Rows = 1
grd9.Clear
grd9.Rows = 1

End Sub

Private Sub DT4_Click()
On Error Resume Next
DT4_Change
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Top = 100
Skin1.LoadSkin App.Path & "\green.skn"
Skin1.ApplySkin Me.hWnd
Call chargec1
Call chargec2
Call chargec4
Call chargec5
Call chargec10
Call chargec12
Call chargec14
Call chargegrd6
Call cont
Text6.Text = sr!rec
Label18.Caption = sr!pou
'face.SBB1.Panels(16).Width = 100
face.SBB1.Panels(16).Text = sr!cca
DT2.Value = Date
DT3.Value = Date
DT4.Value = Date
DT7.Value = Date
DT8.Value = Date - 30
DT9.Value = Date
DT10.Value = Date - 30
DT11.Value = Date
Call chargegrd16
DT12.Value = Date
DT13.Value = Date - 30
DT14.Value = Date
DT15.Value = Date
DT16.Value = Date - 30
DT17.Value = Date
DT18.Value = Date
DT19.Value = Date - 30
DT20.Value = Date
DT21.Value = Date
DT22.Value = Date
grd19.Clear
grd19.Rows = 1
Call chargegrd19
grd20.Clear
grd20.Rows = 1
grd21.Clear
grd21.Rows = 1
Call chargegrd20_21
If face.SBB1.Panels(7).Text = "0" Then
DT3.Enabled = False
DT4.Enabled = False
End If
End Sub
Public Sub chargec1()
On Error Resume Next
Combo6.Clear
Combo6.AddItem "Õ”«» «· ·«„Ì–"
Combo6.AddItem "Õ”«» «·√”« –…"
Combo6.AddItem "Õ”«» «·‘—ﬂ«¡"
Combo6.AddItem "Õ”«» «·⁄„«·"
Combo6.AddItem "«·„’—Ê›« "
Combo6.AddItem "«·»‰ﬂ"
End Sub
Public Sub chargec2()
On Error Resume Next
Combo2.Clear
Combo2.AddItem "Õ«·… ≈ﬂ„«·"
Combo2.AddItem "Õ«·… ≈⁄›«¡"
Combo2.AddItem "Õ«·… «‰”Õ«»"
End Sub
Public Sub chargec3()
On Error Resume Next
Dim t As Double
Dim r As Double
Dim i As Double
Combo3.Clear
i = 0
t = 0
r = 0
Call cont
Do While Not rc.EOF
If Text3.Text = rc!ser Or Val(rc!ser) = Text3.Text Then
i = 1
Combo3.AddItem rc!rec
r = rc!rec
If r > t Then
t = r
End If
End If
rc.MoveNext
Loop
If i = 1 Then
Combo3.Text = t
End If
End Sub
Public Sub chargec4()
On Error Resume Next
Call cont
Combo4.Clear
  Do While Not cl.EOF
  If cl!act = "1" Then
    Combo4.AddItem cl!cla
    End If
    cl.MoveNext
  Loop
End Sub
Public Sub chargec5()
On Error Resume Next
Combo5.Clear
Combo5.AddItem "«ﬂ Ê»—"
Combo5.AddItem "‰Ê›„»—"
Combo5.AddItem "œÌ”„»—"
Combo5.AddItem "Ì‰«Ì—"
Combo5.AddItem "›»—«Ì—"
Combo5.AddItem "„«—”"
Combo5.AddItem "«»—Ì·"
Combo5.AddItem "„«ÌÊ"
Combo5.AddItem "ÌÊ‰ÌÊ"
Combo5.AddItem "ÌÊ·ÌÊ"
Combo5.AddItem "√€”ÿ”"
Combo5.AddItem "”» „»—"
End Sub
Public Sub chargec10()
On Error Resume Next
Combo10.Clear
Combo10.AddItem "«·’‰œÊﬁ"
Combo10.AddItem "«·‘—Ìﬂ"
End Sub
Public Sub chargec12()
On Error Resume Next
Combo12.Clear
Combo13.Clear
Combo12.AddItem "”Õ»"
Combo12.AddItem "«Ìœ«⁄"
Combo12.AddItem "—√” «·„«·"
Combo13.AddItem "”Õ»"
Combo13.AddItem "«Ìœ«⁄"
Combo13.AddItem "—√” «·„«·"
End Sub
Public Sub chargec14()
On Error Resume Next
Combo14.Clear
Combo14.AddItem "«·ﬂ·"
Combo14.AddItem "«·œ«Œ·"
Combo14.AddItem "«·Œ«—Ã"
Combo14.Text = "«·ﬂ·"
End Sub

Private Sub chargegrd6()
On Error Resume Next
Dim i As Double
Dim k As Integer
grd6.Clear
grd6.Cols = 4
grd6.Rows = 1
grd6.ColWidth(0) = 0
grd6.ColWidth(1) = 3500
grd6.ColWidth(2) = 7300
grd6.ColWidth(3) = 3000
grd6.ColAlignment(1) = 1
grd6.ColAlignment(2) = 1
grd6.ColAlignment(3) = 1
grd6.row = 0
grd6.Col = 0
grd6.Text = ""
grd6.Col = 1
grd6.Text = "«·ﬂÊœ"
grd6.Col = 2
grd6.Text = "«·»Ì«‰"
grd6.Col = 3
grd6.Text = "«· ’‰Ì›"
i = 1
Call cont
k = sr!cdc
Check2.Value = k
If cd.RecordCount > 0 Then
cd.MoveFirst
End If
Combo1.Clear
Combo7.Clear
Combo8.Clear
Combo9.Clear
Combo11.Clear
grd6.Rows = cd.RecordCount + 3
Do While Not cd.EOF
If Val(cd!cod) > 0 Then
grd6.row = i
grd6.Col = 0
grd6.Text = cd!aut
grd6.Col = 1
grd6.Text = cd!cod
grd6.Col = 2
grd6.Text = cd!dec
grd6.Col = 3
grd6.Text = cd!cas
End If
i = i + 1
'**** if codes
If k = 1 Then
If cd!cas = "Õ”«» «· ·«„Ì–" Then
Combo1.AddItem cd!cod
Combo1.Text = cd!cod
'Text3.SetFocus
End If
If cd!cas = "Õ”«» «·√”« –…" Then
Combo7.AddItem cd!cod
Combo7.Text = cd!cod
End If
If cd!cas = "«·„’—Ê›« " Then
Combo8.AddItem cd!cod
Combo8.Text = cd!cod
End If
If cd!cas = "Õ”«» «·‘—ﬂ«¡" Then
Combo9.AddItem cd!cod
'Combo8.Text = cd!cod
End If
If cd!cas = "«·»‰ﬂ" Then
Combo11.AddItem cd!cod
'Combo11.Text = cd!cod
End If
End If
'**** if dec
If k = 0 Then
If cd!cas = "Õ”«» «· ·«„Ì–" Then
Combo1.AddItem cd!dec
'Combo1.Text = cd!dec
End If
If cd!cas = "Õ”«» «·√”« –…" Then
Combo7.AddItem cd!dec
'Combo7.Text = cd!dec
End If
If cd!cas = "«·„’—Ê›« " Then
Combo8.AddItem cd!dec
'Combo8.Text = cd!dec
End If
If cd!cas = "Õ”«» «·‘—ﬂ«¡" Then
Combo9.AddItem cd!dec
'Combo8.Text = cd!dec
End If
If cd!cas = "«·»‰ﬂ" Then
Combo11.AddItem cd!dec
'Combo11.Text = cd!dec
End If
End If
'*****
cd.MoveNext
Loop
grd6.Rows = i
grd6.Col = 1
grd6.Sort = 1
End Sub

Private Sub grd01_Click()
Dim r As Double
Dim c As Double
r = grd01.row
c = grd01.Col
If r > 0 Then
grd01.row = r
grd01.Col = 2
Text3.Text = grd01.Text
grd01.Visible = False
Command4_Click
End If
End Sub

Private Sub grd13_Click()
On Error Resume Next
Dim s1 As String
Dim s2 As String
Dim s3 As String
Dim i As Double
Dim j As Double
Dim k As Double
Dim n As Double
Dim a As Double
Dim b As Double
i = grd13.row
j = grd13.Col
n = grd13.Rows
b = 0
If i > 0 Then
If j = 0 Then
MsgBox grd13.Rows
End If
If j = 0 Then
grd13.row = i
grd13.Col = 0
s1 = grd13.Text
For k = 1 To n - 1
grd13.row = k
grd13.Col = 0
s2 = grd13.Text
If s1 = s2 Then
grd13.row = k
grd13.Col = 10
a = grd13.Text
b = b + a
End If
Next k
s3 = b
MsgBox "„Ã„Ê⁄ „” Õﬁ«  «·‘Â—: " + s1 + " ÂÌ : " + s3, vbInformation + arabic
End If
End If
End Sub

Private Sub grd14_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim a As Double
Dim b As Double
Dim c As Double
Dim aut1 As Double
Dim mon1 As String
Dim mon2 As Double
Dim ser1 As String
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim dat4 As Date
Dim annes As String
i = grd14.Col
j = grd14.row
If j > 0 Then
If i = 4 Then
grd14.row = j
grd14.Col = 0
aut1 = grd14.Text
grd14.Col = 3
mon1 = grd14.Text
ser1 = Label20.Caption
g = MsgBox("Â·  —Ìœ Õﬁ« Õ–› «·„»·€ " + mon1, vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
mon2 = mon1
Call cont
'**** controle Date supprimer
annes = sr!ann      'annee scolaire
dat1 = sr!dat       'debut annee scolaire
dat2 = sr!dtf       'Fin annees scolaire
dat3 = Date         'Date Modifiable
dat4 = Date         'Date Machine
If dat3 < dat1 Then
MsgBox "€Ì— „„ﬂ‰..  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“) ”«»ﬁ · «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì… " + annes, vbCritical + arabic
Exit Sub
End If
If dat3 > dat2 Then
MsgBox "€Ì— „„ﬂ‰..  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“) „ √Œ— ⁄‰  «—ÌŒ ‰Â«Ì… «·”‰… «·œ—«”Ì… " + annes, vbCritical + arabic
Exit Sub
End If
'**** end controle Date supprimer
'****** caisse
Call cont
Do While Not ca.EOF
If aut1 = ca!aut Then
'k = 1
ca.Delete
If ca.RecordCount > 0 Then
ca.MoveLast
End If
End If
ca.MoveNext
Loop
'****** journal
c = sr!ord
jr.AddNew
jr!cre = mon2
jr!deb = "0"
jr!dem = "„‰ Õ‹"
jr!com = "«·’‰œÊﬁ"
jr!dec = "«—Ã«⁄ „»·€ " + mon1 + " „‰ ÿ—› «·√” «– ’«Õ» «·—ﬁ„ «· ”·”·Ì " + Label20.Caption
jr!ord = c
jr!dat = Date
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
jr.AddNew
jr!cre = "0"
jr!deb = mon2
jr!dem = "≈·Ï Õ‹"
jr!com = "«·√”« –…"
jr!dec = "«—Ã«⁄ „»·€ " + mon1 + " „‰ ÿ—› «·√” «– ’«Õ» «·—ﬁ„ «· ”·”·Ì " + Label20.Caption
jr!ord = c
jr!dat = Date
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
'****** cascaisse
a = sr!cca
b = mon2
a = a + b
sr!cca = a
sr!ord = c + 1
sr.Update
ProgressBar5.Value = 0
ProgressBar5.Visible = True
Timer4.Enabled = True
End If
End If
End If
End Sub

Private Sub grd16_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
i = grd16.row
j = grd16.Col
If i > 0 Then
grd16.row = i
grd16.Col = 0
Label28.Caption = grd16.Text
grd16.Col = 1
Label96.Caption = grd16.Text
grd16.Col = 2
Label95.Caption = grd16.Text
grd17.Visible = False
grd17.Clear
grd17.Rows = 1
grd18.Visible = False
grd18.Clear
grd18.Rows = 1
Call chargegrd17_18
grd17.Visible = True
grd18.Visible = True
End If
End Sub

Private Sub grd17_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim a As Double
Dim b As Double
Dim c As Double
Dim aut1 As Double
Dim mon1 As String
Dim mon2 As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim dat4 As Date
Dim annes As String
i = grd17.Col
j = grd17.row
If j > 0 Then
If i = 3 Then
grd17.row = j
grd17.Col = 0
aut1 = grd17.Text
grd17.Col = 2
mon1 = grd17.Text
g = MsgBox("Â·  —Ìœ Õﬁ« Õ–› «·„»·€ " + mon1, vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
mon2 = mon1
Call cont
'**** controle Date supprimer
annes = sr!ann      'annee scolaire
dat1 = sr!dat       'debut annee scolaire
dat2 = sr!dtf       'Fin annees scolaire
dat3 = Date         'Date Modifiable
dat4 = Date         'Date Machine
If dat3 < dat1 Then
MsgBox "€Ì— „„ﬂ‰..  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“) ”«»ﬁ · «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì… " + annes, vbCritical + arabic
Exit Sub
End If
If dat3 > dat2 Then
MsgBox "€Ì— „„ﬂ‰..  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“) „ √Œ— ⁄‰  «—ÌŒ ‰Â«Ì… «·”‰… «·œ—«”Ì… " + annes, vbCritical + arabic
Exit Sub
End If
'**** end controle Date supprimer
'****** caisse
Call cont
Do While Not ca.EOF
If aut1 = ca!aut Then
'k = 1
ca.Delete
If ca.RecordCount > 0 Then
ca.MoveLast
End If
End If
ca.MoveNext
Loop
'****** journal
c = sr!ord
jr.AddNew
jr!cre = mon2
jr!deb = "0"
jr!dem = "„‰ Õ‹"
jr!com = "«·’‰œÊﬁ"
jr!dec = "—œ „»·€ " + mon1 + " „‰ «·‘—Ìﬂ " + Label96.Caption
jr!ord = c
jr!dat = Date
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
jr.AddNew
jr!cre = "0"
jr!deb = mon2
jr!dem = "≈·Ï Õ‹"
jr!com = "«·‘—ﬂ«¡"
jr!dec = "—œ „»·€ " + mon1 + " „‰ «·‘—Ìﬂ " + Label96.Caption
jr!ord = c
jr!dat = Date
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
'****** cascaisse
a = sr!cca
b = mon2
a = a + b
sr!cca = a
sr!ord = c + 1
sr.Update
ProgressBar4.Value = 0
ProgressBar4.Visible = True
Timer5.Enabled = True
End If
End If
End If

End Sub

Private Sub grd18_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim a As Double
Dim b As Double
Dim c As Double
Dim aut1 As Double
Dim mon1 As String
Dim mon2 As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim dat4 As Date
Dim annes As String
Dim mca As Double
Dim mnv As Double
i = grd18.Col
j = grd18.row
If j > 0 Then
If i = 3 Then
grd18.row = j
grd18.Col = 0
aut1 = grd18.Text
grd18.Col = 2
mon1 = grd18.Text
g = MsgBox("Â·  —Ìœ Õﬁ« Õ–› «·„»·€ " + mon1, vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
mon2 = mon1
Call cont
'**** controle caisse supprimee
mca = sr!cca
mnv = mon2
If mnv > mca Then
MsgBox "€Ì— „„ﬂ‰.. ·√‰ Õ–› √Ì œ›⁄ „‰ «·‘—Ìﬂ ··’‰œÊﬁ Ì⁄‰Ì «—Ã«⁄ «·„»·€ ··‘—Ìﬂ.. Ê·«ÌÊÃœ ›Ì «·’‰œÊﬁ Õ«·Ì« –«ﬂ «·„»·€", vbCritical
Exit Sub
End If
'**** end controle caisse
'**** controle Date supprimer
annes = sr!ann      'annee scolaire
dat1 = sr!dat       'debut annee scolaire
dat2 = sr!dtf       'Fin annees scolaire
dat3 = Date         'Date Modifiable
dat4 = Date         'Date Machine
If dat3 < dat1 Then
MsgBox "€Ì— „„ﬂ‰..  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“) ”«»ﬁ · «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì… " + annes, vbCritical + arabic
Exit Sub
End If
If dat3 > dat2 Then
MsgBox "€Ì— „„ﬂ‰..  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“) „ √Œ— ⁄‰  «—ÌŒ ‰Â«Ì… «·”‰… «·œ—«”Ì… " + annes, vbCritical + arabic
Exit Sub
End If
'**** end controle Date supprimer
'****** caisse
Call cont
Do While Not ca.EOF
If aut1 = ca!aut Then
'k = 1
ca.Delete
If ca.RecordCount > 0 Then
ca.MoveLast
End If
End If
ca.MoveNext
Loop
'****** journal
c = sr!ord
jr.AddNew
jr!cre = mon2
jr!deb = "0"
jr!dem = "„‰ Õ‹"
jr!com = "«·‘—ﬂ«¡"
jr!dec = "«—Ã«⁄ „»·€ " + mon1 + " ··‘—Ìﬂ " + Label96.Caption
jr!ord = c
jr!dat = Date
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
jr.AddNew
jr!cre = "0"
jr!deb = mon2
jr!dem = "≈·Ï Õ‹"
jr!com = "«·’‰œÊﬁ"
jr!dec = "«—Ã«⁄ „»·€ " + mon1 + " ··‘—Ìﬂ " + Label96.Caption
jr!ord = c
jr!dat = Date
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
'****** cascaisse
a = sr!cca
b = mon2
a = a - b
sr!cca = a
sr!ord = c + 1
sr.Update
ProgressBar4.Value = 0
ProgressBar4.Visible = True
Timer5.Enabled = True
End If
End If
End If

End Sub

Private Sub grd19_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim a As Double
Dim b As Double
Dim c As Double
Dim aut1 As Double
Dim mon1 As String
Dim mon2 As Double
Dim tx1 As String
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim dat4 As Date
Dim annes As String
i = grd19.Col
j = grd19.row
If j > 0 Then
If i = 6 Then
grd19.row = j
grd19.Col = 0
aut1 = grd19.Text
grd19.Col = 1
tx1 = grd19.Text
grd19.Col = 2
mon1 = grd19.Text
g = MsgBox("Â·  —Ìœ Õﬁ« Õ–› «·„»·€ " + mon1, vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
mon2 = mon1
Call cont
'**** controle Date supprimer
annes = sr!ann      'annee scolaire
dat1 = sr!dat       'debut annee scolaire
dat2 = sr!dtf       'Fin annees scolaire
dat3 = Date         'Date Modifiable
dat4 = Date         'Date Machine
If dat3 < dat1 Then
MsgBox "€Ì— „„ﬂ‰..  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“) ”«»ﬁ · «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì… " + annes, vbCritical + arabic
Exit Sub
End If
If dat3 > dat2 Then
MsgBox "€Ì— „„ﬂ‰..  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“) „ √Œ— ⁄‰  «—ÌŒ ‰Â«Ì… «·”‰… «·œ—«”Ì… " + annes, vbCritical + arabic
Exit Sub
End If
'**** end controle Date supprimer
'****** caisse
Call cont
Do While Not ca.EOF
If aut1 = ca!aut Then
'k = 1
ca.Delete
If ca.RecordCount > 0 Then
ca.MoveLast
End If
End If
ca.MoveNext
Loop
Text15.Text = tx1
Text16.Text = Mid$(Text15.Text, 1, 18)
Text16.Text = Trim(Text16.Text)
If Text16.Text = " ”œÌœ —« » «·„ÊŸ›" Then
'****** journal
c = sr!ord
jr.AddNew
jr!cre = mon2
jr!deb = "0"
jr!dem = "„‰ Õ‹"
jr!com = "«·⁄„«·"
jr!dec = "Õ–› „»·€ " + mon1 + " ﬂ«‰ ﬁœ «” Œœ„ „‰ √Ã· " + tx1
jr!ord = c
jr!dat = Date
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
jr.AddNew
jr!cre = "0"
jr!deb = mon2
jr!dem = "≈·Ï Õ‹"
jr!com = "«·„’—Ê›« "
jr!dec = "Õ–› „»·€ " + mon1 + " ﬂ«‰ ﬁœ «” Œœ„ „‰ √Ã· " + tx1
jr!ord = c
jr!dat = Date
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
'****** cascaisse
sr!ord = c + 1
sr.Update
ProgressBar6.Value = 0
ProgressBar6.Visible = True
Timer6.Enabled = True
Exit Sub
End If
'****** journal
c = sr!ord
jr.AddNew
jr!cre = mon2
jr!deb = "0"
jr!dem = "„‰ Õ‹"
jr!com = "«·’‰œÊﬁ"
jr!dec = "Õ–› „»·€ " + mon1 + " ﬂ«‰ ﬁœ «” Œœ„ „‰ √Ã· " + tx1
jr!ord = c
jr!dat = Date
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
jr.AddNew
jr!cre = "0"
jr!deb = mon2
jr!dem = "≈·Ï Õ‹"
jr!com = "«·„’—Ê›« "
jr!dec = "Õ–› „»·€ " + mon1 + " ﬂ«‰ ﬁœ «” Œœ„ „‰ √Ã· " + tx1
jr!ord = c
jr!dat = Date
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
'****** cascaisse
a = sr!cca
b = mon2
a = a + b
sr!cca = a
sr!ord = c + 1
sr.Update
ProgressBar6.Value = 0
ProgressBar6.Visible = True
Timer6.Enabled = True
End If
End If
End If

End Sub

Private Sub grd2_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim r As Double
Dim l As Double
Dim tx1 As String
Dim tx2 As String
Dim tx3 As String
Dim tx4 As String
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim n As Double
Dim m As Double
Dim du As Double
Dim py As Double
Dim dus As Double
Dim pys As Double
Text4.Text = Trim(Text4.Text)
Text5.Text = Trim(Text5.Text)
r = grd2.row
l = grd2.Col
If r > 0 Then
If Text4.Text = "" Then
MsgBox "ÌÃ» «œŒ«· —”Ê„ «· ”ÃÌ·", vbCritical
Text4.SetFocus
Exit Sub
End If
If Text5.Text = "" Then
MsgBox "ÌÃ» «œŒ«· «·—”Ê„ «·‘Â—Ì…", vbCritical
Text5.SetFocus
Exit Sub
End If
If Combo2.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·Õ«·… «· ””ÃÌ·Ì… ·· ·„Ì–", vbCritical
Text3.SetFocus
Exit Sub
End If
'grd2.row = 0
'grd2.col = 1
'tx1 = grd2.Text
'If tx1 <> "0" And Val(tx1) = 0 Then
'grd2.Clear
'grd2.Rows = 1
'Call chargegrd2
'Exit Sub
'End If
grd2.row = r
grd2.Col = 0
tx1 = grd2.Text
grd2.Col = 1
a = grd2.Text
grd2.Col = 2
m = grd2.Text
tx4 = a
n = grd4.Rows
j = 0
dus = 0
pys = 0
For i = 1 To n - 1
grd4.row = i
grd4.Col = 1
tx2 = grd4.Text
If tx1 = tx2 Then
grd4.Col = 2
du = grd4.Text
dus = dus + du
grd4.Col = 3
py = grd4.Text
pys = pys + py
j = 1
End If
Next i
g = InputBox("«œŒ· «·„»·€ «·„œ›Ê⁄", tx1 + "  " + tx4, tx4)
If g = Cancel Then
Exit Sub
End If
If Val(g) < 0 Then
Exit Sub
End If
c = g
If j = 1 Then
If c > (dus - pys) Then
MsgBox "Â–« «·‘Â—  „ œ›⁄Â ”«»ﬁ« Ê·«Ì„ﬂ‰ œ›⁄ √ﬂÀ— „„« ÂÊ »«ﬁ ﬂœÌ‰ ›ÌÂ", vbCritical
Exit Sub
End If
Else
If c > a Then
MsgBox "·« Ì„ﬂ‰ œ›⁄ √ﬂÀ— „‰ «·„»·€ «·„ÿ·Ê»", vbCritical
Exit Sub
End If
End If
If j = 1 Then
d = b - c
Else
d = a - c
End If
If c = 0 Then
d = 0
End If
n = grd3.Rows
For i = 1 To n - 1
grd3.row = i
grd3.Col = 0
tx3 = grd3.Text
If tx3 = tx1 Then
grd3.row = i
grd3.Col = 0
grd3.Text = tx1
grd3.Col = 1
If j = 1 Then
grd3.Text = b
Else
grd3.Text = a
End If
grd3.Col = 2
grd3.Text = c
grd3.Col = 3
grd3.Text = d
grd3.Col = 4
grd3.Text = m
Call totalmontant
Exit Sub
End If
Next i
grd3.Rows = n + 1
grd3.row = n
grd3.Col = 0
grd3.Text = tx1
grd3.Col = 1
If j = 1 Then
grd3.Text = b
Else
grd3.Text = a
End If
grd3.Col = 2
grd3.Text = c
grd3.Col = 3
grd3.Text = d
grd3.Col = 4
grd3.Text = m
Call totalmontant
End If
End Sub

Private Sub grd20_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim a As Double
Dim b As Double
Dim c As Double
Dim aut1 As Double
Dim mon1 As String
Dim mon2 As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim dat4 As Date
Dim annes As String
Dim act1 As String
Dim mca As Double
Dim mnv As Double
i = grd20.Col
j = grd20.row
If j > 0 Then
If i = 4 Then
grd20.row = j
grd20.Col = 0
aut1 = grd20.Text
grd20.Col = 2
mon1 = grd20.Text
grd20.Col = 5
act1 = grd20.Text
If act1 = "1" Then
MsgBox "·« Ì„ﬂ‰ Õ–› Â–« «·»‰œ", vbCritical
Exit Sub
End If
g = MsgBox("Â·  —Ìœ Õﬁ« Õ–› «·„»·€ " + mon1, vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
mon2 = mon1
Call cont
'**** controle caisse supprimee
mca = sr!cca
mnv = mon2
If mnv > mca Then
MsgBox "€Ì— „„ﬂ‰.. ·√‰ Õ–› √Ì ”Õ» Ì⁄‰Ì «—Ã«⁄Â ≈·Ï «·»‰ﬂ.. Ê·«ÌÊÃœ ›Ì «·’‰œÊﬁ Õ«·Ì« –«ﬂ «·„»·€", vbCritical
Exit Sub
End If
'**** end controle caisse
'**** controle Date supprimer
annes = sr!ann      'annee scolaire
dat1 = sr!dat       'debut annee scolaire
dat2 = sr!dtf       'Fin annees scolaire
dat3 = Date         'Date Modifiable
dat4 = Date         'Date Machine
If dat3 < dat1 Then
MsgBox "€Ì— „„ﬂ‰..  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“) ”«»ﬁ · «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì… " + annes, vbCritical + arabic
Exit Sub
End If
If dat3 > dat2 Then
MsgBox "€Ì— „„ﬂ‰..  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“) „ √Œ— ⁄‰  «—ÌŒ ‰Â«Ì… «·”‰… «·œ—«”Ì… " + annes, vbCritical + arabic
Exit Sub
End If
'**** end controle Date supprimer
'****** caisse
Call cont
Do While Not ca.EOF
If aut1 = ca!aut Then
'k = 1
ca.Delete
If ca.RecordCount > 0 Then
ca.MoveLast
End If
End If
ca.MoveNext
Loop
'****** journal
c = sr!ord
jr.AddNew
jr!cre = mon2
jr!deb = "0"
jr!dem = "„‰ Õ‹"
jr!com = "«·»‰ﬂ"
jr!dec = "Õ–› „»·€ " + mon1 + " ﬂ«‰ ﬁœ  „ ”Õ»Â „‰ «·»‰ﬂ "
jr!ord = c
jr!dat = Date
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
jr.AddNew
jr!cre = "0"
jr!deb = mon2
jr!dem = "≈·Ï Õ‹"
jr!com = "«·’‰œÊﬁ"
jr!dec = "Õ–› „»·€ " + mon1 + " ﬂ«‰ ﬁœ  „ ”Õ»Â „‰ «·»‰ﬂ "
jr!ord = c
jr!dat = Date
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
'****** cascaisse
a = sr!cca
b = mon2
a = a - b
sr!cca = a
sr!ord = c + 1
sr.Update
ProgressBar7.Value = 0
ProgressBar7.Visible = True
Timer7.Enabled = True
End If
End If
End If

End Sub

Private Sub grd21_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim a As Double
Dim b As Double
Dim c As Double
Dim aut1 As Double
Dim mon1 As String
Dim mon2 As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim dat4 As Date
Dim annes As String
Dim act1 As String
i = grd21.Col
j = grd21.row
If j > 0 Then
If i = 4 Then
grd21.row = j
grd21.Col = 0
aut1 = grd21.Text
grd21.Col = 2
mon1 = grd21.Text
grd21.Col = 5
act1 = grd21.Text
If act1 = "1" Then
MsgBox "·« Ì„ﬂ‰ Õ–› Â–« «·»‰œ", vbCritical
Exit Sub
End If
g = MsgBox("Â·  —Ìœ Õﬁ« Õ–› «·„»·€ " + mon1, vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
mon2 = mon1
Call cont
'**** controle Date supprimer
annes = sr!ann      'annee scolaire
dat1 = sr!dat       'debut annee scolaire
dat2 = sr!dtf       'Fin annees scolaire
dat3 = Date         'Date Modifiable
dat4 = Date         'Date Machine
If dat3 < dat1 Then
MsgBox "€Ì— „„ﬂ‰..  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“) ”«»ﬁ · «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì… " + annes, vbCritical + arabic
Exit Sub
End If
If dat3 > dat2 Then
MsgBox "€Ì— „„ﬂ‰..  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“) „ √Œ— ⁄‰  «—ÌŒ ‰Â«Ì… «·”‰… «·œ—«”Ì… " + annes, vbCritical + arabic
Exit Sub
End If
'**** end controle Date supprimer
'****** caisse
Call cont
Do While Not ca.EOF
If aut1 = ca!aut Then
'k = 1
ca.Delete
If ca.RecordCount > 0 Then
ca.MoveLast
End If
End If
ca.MoveNext
Loop
'****** journal
c = sr!ord
jr.AddNew
jr!cre = mon2
jr!deb = "0"
jr!dem = "„‰ Õ‹"
jr!com = "«·’‰œÊﬁ"
jr!dec = "Õ–› „»·€ " + mon1 + " ﬂ«‰ ﬁœ  „ «Ìœ«⁄Â ›Ì «·»‰ﬂ "
jr!ord = c
jr!dat = Date
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
jr.AddNew
jr!cre = "0"
jr!deb = mon2
jr!dem = "≈·Ï Õ‹"
jr!com = "«·»‰ﬂ"
jr!dec = "Õ–› „»·€ " + mon1 + " ﬂ«‰ ﬁœ  „ «Ìœ«⁄Â ›Ì «·»‰ﬂ "
jr!ord = c
jr!dat = Date
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
'****** cascaisse
a = sr!cca
b = mon2
a = a + b
sr!cca = a
sr!ord = c + 1
sr.Update
ProgressBar7.Value = 0
ProgressBar7.Visible = True
Timer7.Enabled = True
End If
End If
End If

End Sub


Private Sub grd22_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim a As Double
Dim b As Double
Dim c As Double
Dim aut1 As Double
Dim mon1 As String
Dim com1 As String
Dim com2 As String
Dim dec1 As String
Dim mon2 As Double
Dim dkh As String
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim dat4 As Date
Dim annes As String
Dim act1 As String
Dim mca As Double
Dim mnv As Double
i = grd22.Col
j = grd22.row
If j > 0 Then
If i = 9 Then
grd22.row = j
grd22.Col = 0
aut1 = grd22.Text
grd22.Col = 1
dec1 = grd22.Text
grd22.Col = 2
mon1 = grd22.Text
grd22.Col = 5
dkh = grd22.Text
grd22.Col = 10
com1 = grd22.Text
grd22.Col = 11
act1 = grd22.Text
If act1 = "1" Then
MsgBox "·« Ì„ﬂ‰ Õ–› Â–« «·»‰œ", vbCritical
Exit Sub
End If
If com1 = "Õ”«» «· ·«„Ì–" Then
com1 = "«· ·«„Ì–"
com2 = "«·’‰œÊﬁ"
End If
If com1 = "Õ”«» «·√”« –…" Then
com1 = "«·’‰œÊﬁ"
com2 = "«·√”« –…"
End If
If com1 = "Õ”«» «·‘—ﬂ«¡" Then
If dkh = "œ«Œ·" Then
com1 = "«·‘—ﬂ«¡"
com2 = "«·’‰œÊﬁ"
Else
com1 = "«·’‰œÊﬁ"
com2 = "«·‘—ﬂ«¡"
End If
End If
If com1 = "Õ”«» «·⁄„«·" Then
com1 = "«·’‰œÊﬁ"
com2 = "«·⁄„«·"
End If
g = MsgBox("Â·  —Ìœ Õﬁ« Õ–› «·„»·€ " + mon1, vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
mon2 = mon1
Call cont
'**** controle caisse supprimee
If dkh = "œ«Œ·" Then
mca = sr!cca
mnv = mon2
If mnv > mca Then
MsgBox "€Ì— „„ﬂ‰.. ·√‰ Õ–› √Ì „»·€ œ«Œ· ··’‰œÊﬁ Ì⁄‰Ì «Œ—«ÃÂ „‰ «·’‰œÊﬁ.. Ê·«ÌÊÃœ ›Ì «·’‰œÊﬁ Õ«·Ì« –«ﬂ «·„»·€", vbCritical
Exit Sub
End If
End If
'**** end controle caisse
'**** controle Date supprimer
annes = sr!ann      'annee scolaire
dat1 = sr!dat       'debut annee scolaire
dat2 = sr!dtf       'Fin annees scolaire
dat3 = Date         'Date Modifiable
dat4 = Date         'Date Machine
If dat3 < dat1 Then
MsgBox "€Ì— „„ﬂ‰..  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“) ”«»ﬁ · «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì… " + annes, vbCritical + arabic
Exit Sub
End If
If dat3 > dat2 Then
MsgBox "€Ì— „„ﬂ‰..  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“) „ √Œ— ⁄‰  «—ÌŒ ‰Â«Ì… «·”‰… «·œ—«”Ì… " + annes, vbCritical + arabic
Exit Sub
End If
'**** end controle Date supprimer
'****** caisse
Call cont
Do While Not ca.EOF
If aut1 = ca!aut Then
'k = 1
ca.Delete
If ca.RecordCount > 0 Then
ca.MoveLast
End If
End If
ca.MoveNext
Loop
'****** journal
c = sr!ord
jr.AddNew
jr!cre = mon2
jr!deb = "0"
jr!dem = "„‰ Õ‹"
jr!com = com1
jr!dec = "Õ–› „»·€ " + mon1 + " ﬂ«‰  Õ  «·»Ì«‰ '" + dec1 + "'"
jr!ord = c
jr!dat = Date
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
jr.AddNew
jr!cre = "0"
jr!deb = mon2
jr!dem = "≈·Ï Õ‹"
jr!com = com2
jr!dec = "Õ–› „»·€ " + mon1 + " ﬂ«‰  Õ  «·»Ì«‰ '" + dec1 + "'"
jr!ord = c
jr!dat = Date
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
'****** cascaisse
a = sr!cca
b = mon2
If dkh = "Œ«—Ã" Then
a = a + b
Else
a = a - b
End If
sr!cca = a
sr!ord = c + 1
sr.Update
ProgressBar8.Value = 0
ProgressBar8.Visible = True
Timer8.Enabled = True
End If
End If
End If
grd22.ToolTipText = grd22.Text

End Sub

Private Sub grd3_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim n As Double
Dim tx As String
Dim du As Double
Dim py As Double
Dim rs As Double
If grd3.Rows > 1 Then
g = MsgBox("Â·  —Ìœ Õ–› Â–« «·”ÿ—", vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
j = grd3.row
n = grd3.Rows
For i = j To n - 2
grd3.row = i + 1
grd3.Col = 0
tx = grd3.Text
grd3.Col = 1
du = grd3.Text
grd3.Col = 2
py = grd3.Text
grd3.Col = 3
rs = grd3.Text
grd3.row = i
grd3.Col = 0
grd3.Text = tx
grd3.Col = 1
grd3.Text = du
grd3.Col = 2
grd3.Text = py
grd3.Col = 3
grd3.Text = rs
Next i
grd3.Rows = i
Call totalmontant
End If
End If
End Sub

Private Sub grd6_Click()
On Error Resume Next
Dim i As Double
i = grd6.row
grd6.row = i
grd6.Col = 0
Label1.Caption = grd6.Text
grd6.Col = 1
Text1.Text = grd6.Text
grd6.Col = 2
Text2.Text = grd6.Text
grd6.Col = 3
Combo6.Text = grd6.Text
End Sub


Private Sub grd7_Click()
On Error Resume Next
Dim ane As String
Dim i As Double
Dim j As Double
Dim k As Integer
Dim tx As String
Dim seri As String
Dim o As Double
Dim m As Double
Dim c As Double
Dim x As Double
Dim recu As String
Dim au As String
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim dat4 As Date
Dim annes As String
Dim mca As Double
Dim mnv As Double
i = grd7.row
j = grd7.Col
If grd7.Rows < 2 Then
MsgBox "·« ÌÊÃœ √Ì Ê’· ", vbCritical
Exit Sub
End If
k = 0
If j = 1 Then
Dim a As Double
grd7.row = i
grd7.Col = 1
a = grd7.Text
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
data.DoCmd.Maximize
data.DoCmd.OpenReport "recu", acViewPreview, , "rec =" & a, acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
data.Visible = True
Set data = Nothing
End If
If j = 5 Then
Call cont
'**** controle caisse supprimee
mca = sr!cca
grd7.row = i
grd7.Col = 2
m = grd7.Text
mnv = m
If mnv > mca Then
MsgBox "€Ì— „„ﬂ‰.. ·√‰ Õ–› √Ì Ê’· Ì⁄‰Ì «—Ã«⁄ „»·€Â ≈·Ï ’«Õ»Â.. Ê·«ÌÊÃœ ›Ì «·’‰œÊﬁ Õ«·Ì« –«ﬂ «·„»·€", vbCritical
Exit Sub
End If
'**** end controle caisse
'**** controle Date supprimer
annes = sr!ann      'annee scolaire
dat1 = sr!dat       'debut annee scolaire
dat2 = sr!dtf       'Fin annees scolaire
dat3 = Date         'Date Modifiable
dat4 = Date         'Date Machine
If dat3 < dat1 Then
MsgBox "€Ì— „„ﬂ‰..  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“) ”«»ﬁ · «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì… " + annes, vbCritical + arabic
Exit Sub
End If
If dat3 > dat2 Then
MsgBox "€Ì— „„ﬂ‰..  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“) „ √Œ— ⁄‰  «—ÌŒ ‰Â«Ì… «·”‰… «·œ—«”Ì… " + annes, vbCritical + arabic
Exit Sub
End If
'**** end controle Date supprimer
grd7.row = i
grd7.Col = 0
tx = grd7.Text
grd7.Col = 1
recu = grd7.Text
grd7.Col = 2
m = grd7.Text
grd7.Col = 3
seri = grd7.Text
grd7.Col = 6
au = grd7.Text
g = MsgBox("Â·  —Ìœ «·Õ–› Õﬁ« ", vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
g = MsgBox("ÌÃ» «·«‰ »«Â ≈·Ï √‰ Õ–› √Ì Ê’· Ì⁄‰Ì «—Ã«⁄ ﬂ«„· „»«·€Â ≈·Ï œ«›⁄Â œÊ‰ «” À‰«¡ ··‘ÂÊ—, ›Â·  —Ìœ «·«” „—«—ø ", vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
Call cont
Do While Not ca.EOF
If au = ca!aut Then
k = 1
ca.Delete
If ca.RecordCount > 0 Then
ca.MoveLast
End If
End If
ca.MoveNext
Loop
If k = 1 Then
recu = m
Call cont
'****** journal
o = sr!ord
jr.AddNew
jr!cre = m
jr!deb = "0"
jr!dem = "„‰ Õ‹"
jr!com = "«· ·«„Ì–"
jr!dec = "Õ–› „»·€ " + recu + " «·„œ›Ê⁄ „‰ ÿ—› «· ·„Ì– ’«Õ» «·—ﬁ„ «· ”·”·Ì " + seri
jr!ord = o
jr!dat = Date
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
jr.AddNew
jr!cre = "0"
jr!deb = m
jr!dem = "≈·Ï Õ‹"
jr!com = "«·’‰œÊﬁ"
jr!dec = "Õ–› „»·€ " + recu + " «·„œ›Ê⁄ „‰ ÿ—› «· ·„Ì– ’«Õ» «·—ﬁ„ «· ”·”·Ì " + seri
jr!ord = o
jr!dat = Date
jr!heu = Time$
jr!ger = face.SBB1.Panels(11).Text
jr.Update
'****** cascaisse
c = sr!cca
x = m
c = c - x
sr!cca = c
sr!ord = o + 1
sr.Update
'****** recu pour imprimer
Call cont2
Do While Not ru.EOF
If recu = ru!rec Then
ru.Delete
If ru.RecordCount > 0 Then
ru.MoveLast
End If
End If
ru.MoveNext
Loop
ProgressBar3.Value = 0
Timer3.Enabled = True
Exit Sub
End If
End If
End If
End If
End Sub

Private Sub grd8_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim k As Integer
Dim tx As String
Dim seri As String
Dim o As Double
Dim m As Double
Dim c As Double
Dim x As Double
Dim recu As String
Dim au As String
i = grd8.row
j = grd8.Col
If grd8.Rows < 2 Then
MsgBox "·« ÌÊÃœ √Ì Ê’· ", vbCritical
Exit Sub
End If
k = 0
If j = 5 Then
grd8.row = i
grd8.Col = 0
tx = grd8.Text
grd8.Col = 1
recu = grd8.Text
grd8.Col = 2
m = grd8.Text
grd8.Col = 3
seri = grd8.Text
grd8.Col = 6
au = grd8.Text
g = MsgBox("Â·  —Ìœ «·Õ–› Õﬁ« ", vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
Call cont
Do While Not ca.EOF
If au = ca!aut Then
k = 1
ca.Delete
If ca.RecordCount > 0 Then
ca.MoveLast
End If
End If
ca.MoveNext
Loop
If k = 1 Then
Call cont
'****** journal
o = sr!ord
jr.AddNew
jr!cre = m
jr!deb = "0"
jr!dem = "„‰ Õ‹"
jr!com = "«· ·«„Ì–"
jr!dec = "Õ–› Ê’· —ﬁ„ " + recu + " «·„œ›Ê⁄ „‰ ÿ—› «· ·„Ì– ’«Õ» «·—ﬁ„ «· ”·”·Ì " + seri
jr!ord = o
jr!dat = Date
jr!heu = Time$
jr!ger = ""
jr.Update
jr.AddNew
jr!cre = "0"
jr!deb = m
jr!dem = "≈·Ï Õ‹"
jr!com = "«·’‰œÊﬁ"
jr!dec = "Õ–› Ê’· —ﬁ„ " + recu + " «·„œ›Ê⁄ „‰ ÿ—› «· ·„Ì– ’«Õ» «·—ﬁ„ «· ”·”·Ì " + seri
jr!ord = o
jr!dat = Date
jr!heu = Time$
jr!ger = ""
jr.Update
'****** cascaisse
c = sr!cca
x = m
c = c - x
sr!cca = c
sr!ord = o + 1
sr.Update
'****** recu pour imprimer
Call cont2
Do While Not ru.EOF
If recu = ru!rec Then
ru.Delete
ProgressBar3.Value = 0
Timer3.Enabled = True
Exit Sub
End If
ru.MoveNext
Loop
End If
End If
End If

End Sub

Private Sub grd9_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim k As Integer
Dim tx As String
Dim seri As String
Dim o As Double
Dim m As Double
Dim c As Double
Dim x As Double
Dim recu As String
Dim au As String
i = grd9.row
j = grd9.Col
If grd9.Rows < 2 Then
MsgBox "·« ÌÊÃœ √Ì Ê’· ", vbCritical
Exit Sub
End If
k = 0
If j = 5 Then
grd9.row = i
grd9.Col = 0
tx = grd9.Text
grd9.Col = 1
recu = grd9.Text
grd9.Col = 2
m = grd9.Text
grd9.Col = 3
seri = grd9.Text
grd9.Col = 6
au = grd9.Text
g = MsgBox("Â·  —Ìœ «·Õ–› Õﬁ« ", vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
Call cont
Do While Not ca.EOF
If au = ca!aut Then
k = 1
ca.Delete
If ca.RecordCount > 0 Then
ca.MoveLast
End If
End If
ca.MoveNext
Loop
If k = 1 Then
Call cont
'****** journal
o = sr!ord
jr.AddNew
jr!cre = m
jr!deb = "0"
jr!dem = "„‰ Õ‹"
jr!com = "«· ·«„Ì–"
jr!dec = "Õ–› Ê’· —ﬁ„ " + recu + " «·„œ›Ê⁄ „‰ ÿ—› «· ·„Ì– ’«Õ» «·—ﬁ„ «· ”·”·Ì " + seri
jr!ord = o
jr!dat = Date
jr!heu = Time$
jr!ger = ""
jr.Update
jr.AddNew
jr!cre = "0"
jr!deb = m
jr!dem = "≈·Ï Õ‹"
jr!com = "«·’‰œÊﬁ"
jr!dec = "Õ–› Ê’· —ﬁ„ " + recu + " «·„œ›Ê⁄ „‰ ÿ—› «· ·„Ì– ’«Õ» «·—ﬁ„ «· ”·”·Ì " + seri
jr!ord = o
jr!dat = Date
jr!heu = Time$
jr!ger = ""
jr.Update
'****** cascaisse
c = sr!cca
x = m
c = c - x
sr!cca = c
sr!ord = o + 1
sr.Update
'****** recu pour imprimer
Call cont2
Do While Not ru.EOF
If recu = ru!rec Then
ru.Delete
ProgressBar3.Value = 0
Timer3.Enabled = True
Exit Sub
End If
ru.MoveNext
Loop
End If
End If
End If

End Sub



Private Sub SSTab1_Click(PreviousTab As Integer)
If face.SBB1.Panels(7).Text = "0" Then
SSTab1.Tab = 6
End If
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
On Error Resume Next
Label6.Caption = "0"
Label7.Caption = "0"
Label8.Caption = "0"
Label9.Caption = "0"
grd7.Clear
grd7.Rows = 1
grd8.Clear
grd8.Rows = 1
grd9.Clear
grd9.Rows = 1
grd10.Clear
grd10.Rows = 1
grd11.Clear
grd11.Rows = 1
grd12.Clear
grd12.Rows = 1

End Sub
Private Sub SSTab3_Click(PreviousTab As Integer)
On Error Resume Next
Label6.Caption = "0"
Label7.Caption = "0"
Label8.Caption = "0"
Label9.Caption = "0"
grd7.Clear
grd7.Rows = 1
grd8.Clear
grd8.Rows = 1
grd9.Clear
grd9.Rows = 1
grd10.Clear
grd10.Rows = 1
grd11.Clear
grd11.Rows = 1
grd12.Clear
grd12.Rows = 1

End Sub

Private Sub SSTab4_Click(PreviousTab As Integer)
On Error Resume Next
grd15.Clear
grd15.Rows = 1


End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii <> 8 Then
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End If

End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim i As Integer
Dim j As Integer
Dim n As Double
Dim vg As String
Text10.Text = Trim(Text10.Text)
n = Len(Text10.Text)
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
vg = Mid$(Text10.Text, i, 1)
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

Private Sub Text11_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim i As Integer
Dim j As Integer
Dim n As Double
Dim vg As String
Text11.Text = Trim(Text11.Text)
n = Len(Text11.Text)
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
vg = Mid$(Text11.Text, i, 1)
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

Private Sub Text12_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim i As Integer
Dim j As Integer
Dim n As Double
Dim vg As String
Text12.Text = Trim(Text12.Text)
n = Len(Text12.Text)
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
vg = Mid$(Text12.Text, i, 1)
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

Private Sub Text17_Change()
'On Error Resume Next
grd4.Clear
grd4.Rows = 1
grd3.Clear
grd3.Rows = 1
Label2.Caption = ""
Label3.Caption = ""
Label75.Caption = ""
Picture6.Visible = False
grd1.Clear
grd01.Visible = False
Text3.Text = ""
End Sub

Private Sub Text17_Click()
On Error Resume Next
Text17_Change

End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii <> 8 Then
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End If

End Sub

Private Sub Text17_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error Resume Next
If Text17.Text <> "" Then
If KeyCode = 13 Then
Call chargegrd01
End If
End If

End Sub
Private Sub chargegrd01()
'On Error Resume Next
Dim i As Double
Dim j As Double
Dim s1 As Double
Dim s2 As Double
Dim tx As String
grd01.Clear
grd01.Rows = 1
grd01.Cols = 5
grd01.ColWidth(0) = 700
grd01.ColWidth(1) = 3000
grd01.ColWidth(2) = 1500
grd01.ColWidth(3) = 1000
grd01.ColWidth(4) = 1000
grd01.ColAlignment(1) = 1
grd01.ColAlignment(2) = 1
grd01.ColAlignment(0) = 1
grd01.ColAlignment(3) = 1
grd01.ColAlignment(4) = 1
grd01.row = 0
grd01.Col = 0
grd01.Text = "«·—ﬁ„"
grd01.Col = 1
grd01.Text = "«·«”„"
grd01.Col = 2
grd01.Text = "«·—ﬁ„ «· ”·”·Ì"
grd01.Col = 3
grd01.Text = "«·ﬁ”„"
i = 1
Call cont
grd01.Rows = et.RecordCount + 30
Do While Not et.EOF
If Text17.Text = et!tel Then
If Val(et!num) < 1000000 Then
grd01.row = i
grd01.Col = 0
grd01.Text = et!num
grd01.Col = 1
grd01.Text = et!nom
grd01.Col = 2
grd01.Text = et!ser
grd01.Col = 3
grd01.Text = et!cla
grd01.Col = 4
grd01.Text = ""
i = i + 1
End If
End If
et.MoveNext
Loop
grd01.Rows = i
If (i > 1) Then
Call cont
Do While Not ce.EOF
s1 = ce!ser
For j = 1 To i - 1
grd01.row = j
grd01.Col = 2
s2 = grd01.Text
If s1 = s2 Then
grd01.row = j
grd01.Col = 4
grd01.Text = ce!cas
End If
Next j
ce.MoveNext
Loop
grd01.Visible = True
Else
MsgBox "·« ÌÊÃœ  ·„Ì– „”Ã·  Õ  Â–« «·—ﬁ„", vbExclamation
Text17.Text = ""
End If
End Sub
Private Sub chargegrd80()
'On Error Resume Next
Dim i As Double
Dim j As Double
Dim s1 As Double
Dim s2 As Double
Dim tx As String
Dim cas As String
Dim cla As String
Dim rs10 As Double
Dim rs11 As Double
Dim rs12 As Double
Dim rs1 As Double
Dim rs2 As Double
Dim rs3 As Double
Dim rs4 As Double
Dim rs5 As Double
Dim rs6 As Double
grd80.Clear
grd80.Rows = 1
grd80.Cols = 12
grd80.ColWidth(0) = 1500
grd80.ColWidth(1) = 3000
grd80.ColWidth(3) = 1200
grd80.ColWidth(4) = 1200
grd80.ColWidth(5) = 1200
grd80.ColWidth(6) = 1200
grd80.ColWidth(7) = 1200
grd80.ColWidth(8) = 1200
grd80.ColWidth(9) = 1200
grd80.ColWidth(10) = 1200
grd80.ColWidth(11) = 1200
grd80.ColAlignment(1) = 1
grd80.ColAlignment(2) = 1
grd80.ColAlignment(0) = 1
grd80.ColAlignment(3) = 1
grd80.ColAlignment(4) = 1
grd80.ColAlignment(5) = 1
grd80.ColAlignment(6) = 1
grd80.ColAlignment(7) = 1
grd80.ColAlignment(8) = 1
grd80.ColAlignment(9) = 1
grd80.ColAlignment(10) = 1
grd80.ColAlignment(11) = 1
grd80.row = 0
grd80.Col = 0
grd80.Text = "«·—ﬁ„"
grd80.Col = 1
grd80.Text = "«·«”„"
grd80.Col = 2
grd80.Text = "10"
grd80.Col = 3
grd80.Text = "11"
grd80.Col = 4
grd80.Text = "12"
grd80.Col = 5
grd80.Text = "1"
grd80.Col = 6
grd80.Text = "2"
grd80.Col = 7
grd80.Text = "3"
grd80.Col = 8
grd80.Text = "4"
grd80.Col = 9
grd80.Text = "5"
grd80.Col = 10
grd80.Text = "6"
grd80.Col = 11
grd80.Text = "«·—”Ê„"
i = 1
Call cont
grd80.Rows = et.RecordCount + 30
Do While Not et.EOF
If Label48.Caption = et!tel Then
If Val(et!num) < 1000000 Then
cla = et!cla
grd80.row = i
grd80.Col = 0
grd80.Text = et!ser
grd80.Col = 1
grd80.Text = et!nom
grd80.Col = 2
grd80.Text = "0"
grd80.Col = 3
grd80.Text = "0"
grd80.Col = 4
grd80.Text = "0"
grd80.Col = 5
grd80.Text = "0"
grd80.Col = 6
grd80.Text = "0"
grd80.Col = 7
grd80.Text = "0"
grd80.Col = 8
grd80.Text = "0"
grd80.Col = 9
grd80.Text = "0"
grd80.Col = 10
grd80.Text = "0"
grd80.Col = 11
grd80.Text = ""
i = i + 1
End If
End If
et.MoveNext
Loop
grd80.Rows = i
If i > 1 Then
j = 0
Call cont
Do While Not ce.EOF
For j = 1 To grd80.Rows - 1
grd80.row = j
grd80.Col = 0
tx = grd80.Text
If Val(tx) = Val(ce!ser) Then
grd80.row = j
grd80.Col = 2
grd80.Text = ce!man
grd80.Col = 3
grd80.Text = ce!man
grd80.Col = 4
grd80.Text = ce!man
grd80.Col = 5
grd80.Text = ce!man
grd80.Col = 6
grd80.Text = ce!man
grd80.Col = 7
grd80.Text = ce!man
grd80.Col = 8
grd80.Text = ce!man
grd80.Col = 9
grd80.Text = ce!man
grd80.Col = 10
grd80.Text = ce!man
grd80.Col = 11
grd80.Text = ""
'ce.MoveLast
End If
Next j
ce.MoveNext
Loop
j = 0
Call cont
Do While Not ce.EOF
For j = 1 To grd80.Rows - 1
grd80.row = j
grd80.Col = 0
tx = grd80.Text
If Val(tx) = Val(ce!ser) Then
'Õ«·… ≈⁄›«¡
If ce!cas = "Õ«·… ≈⁄›«¡" Then
grd80.Col = 2
grd80.Text = "0"
grd80.Col = 3
grd80.Text = "0"
grd80.Col = 4
grd80.Text = "0"
grd80.Col = 5
grd80.Text = "0"
grd80.Col = 6
grd80.Text = "0"
grd80.Col = 7
grd80.Text = "0"
grd80.Col = 8
grd80.Text = "0"
grd80.Col = 9
grd80.Text = "0"
grd80.Col = 10
grd80.Text = "0"
grd80.Col = 11
grd80.Text = "≈⁄›«¡"
End If
'Õ«·… ≈ﬂ„«·
If ce!cas = "Õ«·… ≈ﬂ„«·" Then
'10
If Val(ce!mois) = 10 Then
grd80.Col = 2
rs10 = grd80.Text
rs10 = rs10 - ce!pay
grd80.Text = rs10
End If
'11
If Val(ce!mois) = 11 Then
grd80.row = j
grd80.Col = 3
rs11 = grd80.Text
rs11 = rs11 - ce!pay
grd80.Text = rs11
End If
'12
If Val(ce!mois) = 12 Then
grd80.row = j
grd80.Col = 4
rs12 = grd80.Text
rs12 = rs12 - ce!pay
grd80.Text = rs12
End If
'1
If Val(ce!mois) = 1 Then
grd80.row = j
grd80.Col = 5
rs1 = grd80.Text
rs1 = rs1 - ce!pay
grd80.Text = rs1
End If
'2
If Val(ce!mois) = 2 Then
grd80.row = j
grd80.Col = 6
rs2 = grd80.Text
rs2 = rs2 - ce!pay
grd80.Text = rs2
End If
'3
If Val(ce!mois) = 3 Then
grd80.row = j
grd80.Col = 7
rs3 = grd80.Text
rs3 = rs3 - ce!pay
grd80.Text = rs3
End If
'4
If Val(ce!mois) = 4 Then
grd80.row = j
grd80.Col = 8
rs4 = grd80.Text
rs4 = rs4 - ce!pay
grd80.Text = rs4
End If
'5
If Val(ce!mois) = 5 Then
grd80.row = j
grd80.Col = 9
rs5 = grd80.Text
rs5 = rs5 - ce!pay
grd80.Text = rs5
End If
'6
If Val(ce!mois) = 6 Then
grd80.row = j
grd80.Col = 10
rs6 = grd80.Text
rs6 = rs6 - ce!pay
grd80.Text = rs6
End If
grd80.Col = 11
grd80.Text = "≈ﬂ„«·"
End If
End If
Next j
ce.MoveNext
Loop
End If

End Sub

Private Sub Text18_Change()
On Error Resume Next
Label10.Caption = ""
Label19.Caption = ""
Label20.Caption = ""
Text8.Text = ""
Text7.Text = ""
Picture12.Visible = False

End Sub

Private Sub Text18_Click()
Text18_Change
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii <> 8 Then
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End If

End Sub

Private Sub Text18_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If Text18.Text <> "" Then
If KeyCode = 13 Then
Call cont
Do While Not pr.EOF
If Text18.Text = pr!tel Then
If pr!act = "1" Then
Text7.Text = pr!ser
Label10.Caption = pr!nom
Label19.Caption = pr!tel
Label20.Caption = pr!ser
Exit Sub
End If
End If
pr.MoveNext
Loop
MsgBox "—ﬁ„ «·Â« › «·„œŒ· €Ì— „Œ“‰ .. Ì—ÃÏ «· √ﬂœ „‰Â", vbExclamation
Text18.Text = ""
Text18.SetFocus
End If
End If
End Sub

Private Sub Text3_Change()
On Error Resume Next
grd4.Clear
grd4.Rows = 1
grd3.Clear
grd3.Rows = 1
Label2.Caption = ""
Label3.Caption = ""
Label75.Caption = ""
Picture6.Visible = False
grd1.Clear
grd01.Visible = False
End Sub

Private Sub Text3_Click()
On Error Resume Next
Text3_Change
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii <> 8 Then
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End If

End Sub

Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If Text3.Text <> "" Then
If KeyCode = 13 Then
Command4_Click
End If
End If

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim i As Integer
Dim j As Integer
Dim n As Double
Dim vg As String
Text4.Text = Trim(Text4.Text)
n = Len(Text4.Text)
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
vg = Mid$(Text4.Text, i, 1)
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

Private Sub Text5_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim i As Integer
Dim j As Integer
Dim n As Double
Dim vg As String
Text5.Text = Trim(Text5.Text)
n = Len(Text5.Text)
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
vg = Mid$(Text5.Text, i, 1)
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

Private Sub Text6_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii <> 8 Then
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End If

End Sub

Private Sub Text7_Change()
On Error Resume Next
Label10.Caption = ""
Label19.Caption = ""
Label20.Caption = ""
Text8.Text = ""
Picture12.Visible = False
End Sub

Private Sub Text7_Click()
On Error Resume Next
Text7_Change
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii <> 8 Then
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End If

End Sub

Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If Text7.Text <> "" Then
If KeyCode = 13 Then
Call cont
Do While Not pr.EOF
If Text7.Text = pr!ser Or Val(pr!ser) = Text7.Text Then
If pr!act = "1" Then
Label10.Caption = pr!nom
Label19.Caption = pr!tel
Label20.Caption = pr!ser
Exit Sub
End If
End If
pr.MoveNext
Loop
MsgBox "«·—ﬁ„ «· ”·”·Ì «·„œŒ· €Ì— „Œ“‰ .. Ì—ÃÏ «· √ﬂœ „‰Â", vbExclamation
Text7.Text = ""
Text7.SetFocus
End If
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim i As Integer
Dim j As Integer
Dim n As Double
Dim vg As String
Text8.Text = Trim(Text8.Text)
n = Len(Text8.Text)
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
vg = Mid$(Text8.Text, i, 1)
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

Private Sub Timer1_Timer()
On Error Resume Next
ProgressBar1.Value = ProgressBar1.Value + 8
If ProgressBar1.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
Text2.Text = ""
Label1.Caption = ""
Text1.Text = ""
Text1.SetFocus
ProgressBar1.Value = 0
Timer1.Enabled = False
Label31(58).Visible = False
Combo13.Visible = False
Call chargec1
grd6.Visible = False
Call chargegrd6
grd6.Visible = True
End If

End Sub
Private Sub chargegrd2()
On Error Resume Next
Dim i As Double
grd2.Clear
grd2.Rows = 14
grd2.ColWidth(2) = 0
If Combo2.Text = "Õ«·… ≈ﬂ„«·" Then
grd2.row = 0
grd2.Col = 0
grd2.Text = "«·‘Â—"
grd2.Col = 1
grd2.Text = "«·—”Ê„"
grd2.row = 1
grd2.Col = 0
grd2.Text = "«· ”ÃÌ·"
grd2.Col = 1
grd2.Text = Text4.Text
grd2.Col = 2
grd2.Text = DT2.Month
grd2.row = 2
grd2.Col = 0
grd2.Text = "«ﬂ Ê»—"
grd2.Col = 1
grd2.Text = Text5.Text
grd2.Col = 2
grd2.Text = "10"
grd2.row = 3
grd2.Col = 0
grd2.Text = "‰Ê›„»—"
grd2.Col = 1
grd2.Text = Text5.Text
grd2.Col = 2
grd2.Text = "11"
grd2.row = 4
grd2.Col = 0
grd2.Text = "œÌ”„»—"
grd2.Col = 1
grd2.Text = Text5.Text
grd2.Col = 2
grd2.Text = "12"
grd2.row = 5
grd2.Col = 0
grd2.Text = "Ì‰«Ì—"
grd2.Col = 1
grd2.Text = Text5.Text
grd2.Col = 2
grd2.Text = "1"
grd2.row = 6
grd2.Col = 0
grd2.Text = "›»—«Ì—"
grd2.Col = 1
grd2.Text = Text5.Text
grd2.Col = 2
grd2.Text = "2"
grd2.row = 7
grd2.Col = 0
grd2.Text = "„«—”"
grd2.Col = 1
grd2.Text = Text5.Text
grd2.Col = 2
grd2.Text = "3"
grd2.row = 8
grd2.Col = 0
grd2.Text = "«»—Ì·"
grd2.Col = 1
grd2.Text = Text5.Text
grd2.Col = 2
grd2.Text = "4"
grd2.row = 9
grd2.Col = 0
grd2.Text = "„«ÌÊ"
grd2.Col = 1
grd2.Text = Text5.Text
grd2.Col = 2
grd2.Text = "5"
grd2.row = 10
grd2.Col = 0
grd2.Text = "ÌÊ‰ÌÊ"
grd2.Col = 1
grd2.Text = Text5.Text
grd2.Col = 2
grd2.Text = "6"
grd2.row = 11
grd2.Col = 0
grd2.Text = "ÌÊ·ÌÊ"
grd2.Col = 1
grd2.Text = Text5.Text
grd2.Col = 2
grd2.Text = "7"
grd2.row = 12
grd2.Col = 0
grd2.Text = "√€”ÿ”"
grd2.Col = 1
grd2.Text = Text5.Text
grd2.Col = 2
grd2.Text = "8"
grd2.row = 13
grd2.Col = 0
grd2.Text = "”» „»—"
grd2.Col = 1
grd2.Text = Text5.Text
grd2.Col = 2
grd2.Text = "9"
Else
grd2.row = 0
grd2.Col = 0
grd2.Text = "«·‘Â—"
grd2.Col = 1
grd2.Text = "«·—”Ê„"
grd2.row = 1
grd2.Col = 0
grd2.Text = "«· ”ÃÌ·"
grd2.Col = 1
grd2.Text = "0"
grd2.Col = 2
grd2.Text = "10"
grd2.row = 2
grd2.Col = 0
grd2.Text = "«ﬂ Ê»—"
grd2.Col = 1
grd2.Text = "0"
grd2.Col = 2
grd2.Text = "10"
grd2.row = 3
grd2.Col = 0
grd2.Text = "‰Ê›„»—"
grd2.Col = 1
grd2.Text = "0"
grd2.Col = 2
grd2.Text = "11"
grd2.row = 4
grd2.Col = 0
grd2.Text = "œÌ”„»—"
grd2.Col = 1
grd2.Text = "0"
grd2.Col = 2
grd2.Text = "12"
grd2.row = 5
grd2.Col = 0
grd2.Text = "Ì‰«Ì—"
grd2.Col = 1
grd2.Text = "0"
grd2.Col = 2
grd2.Text = "1"
grd2.row = 6
grd2.Col = 0
grd2.Text = "›»—«Ì—"
grd2.Col = 1
grd2.Text = "0"
grd2.Col = 2
grd2.Text = "2"
grd2.row = 7
grd2.Col = 0
grd2.Text = "„«—”"
grd2.Col = 1
grd2.Text = "0"
grd2.Col = 2
grd2.Text = "3"
grd2.row = 8
grd2.Col = 0
grd2.Text = "«»—Ì·"
grd2.Col = 1
grd2.Text = "0"
grd2.Col = 2
grd2.Text = "4"
grd2.row = 9
grd2.Col = 0
grd2.Text = "„«ÌÊ"
grd2.Col = 1
grd2.Text = "0"
grd2.Col = 2
grd2.Text = "5"
grd2.row = 10
grd2.Col = 0
grd2.Text = "ÌÊ‰ÌÊ"
grd2.Col = 1
grd2.Text = "0"
grd2.Col = 2
grd2.Text = "6"
grd2.row = 11
grd2.Col = 0
grd2.Text = "ÌÊ·ÌÊ"
grd2.Col = 1
grd2.Text = "0"
grd2.Col = 2
grd2.Text = "7"
grd2.row = 12
grd2.Col = 0
grd2.Text = "√€”ÿ”"
grd2.Col = 1
grd2.Text = "0"
grd2.Col = 2
grd2.Text = "8"
grd2.row = 13
grd2.Col = 0
grd2.Text = "”» „»—"
grd2.Col = 1
grd2.Text = "0"
grd2.Col = 2
grd2.Text = "9"
End If
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd3.ColWidth(0) = 1100
grd3.ColWidth(1) = 1100
grd3.ColWidth(2) = 1100
grd3.ColWidth(3) = 1100
grd3.ColWidth(4) = 0
grd3.ColAlignment(0) = 1
grd3.ColAlignment(1) = 1
grd3.ColAlignment(2) = 1
grd3.ColAlignment(3) = 1
grd4.ColWidth(0) = 0
grd4.ColWidth(1) = 1100
grd4.ColWidth(2) = 1100
grd4.ColWidth(3) = 1100
grd4.ColWidth(4) = 1100
grd4.ColWidth(6) = 0
grd4.ColAlignment(0) = 1
grd4.ColAlignment(1) = 1
grd4.ColAlignment(2) = 1
grd4.ColAlignment(3) = 1
grd4.ColAlignment(4) = 1
grd4.ColAlignment(5) = 1
grd3.row = 0
grd3.Col = 0
grd3.Text = "«·‘Â—"
grd3.Col = 1
grd3.Text = "«·„” Õﬁ"
grd3.Col = 2
grd3.Text = "«·„œ›Ê⁄"
grd3.Col = 3
grd3.Text = "«·»«ﬁÌ"
grd4.row = 0
grd4.Col = 1
grd4.Text = "«·‘Â—"
grd4.Col = 2
grd4.Text = "«·„” Õﬁ"
grd4.Col = 3
grd4.Text = "«·„œ›Ê⁄"
grd4.Col = 4
grd4.Text = "«·»«ﬁÌ"
grd4.Col = 5
grd4.Text = "—ﬁ„ «·Ê’·"
grd4.Col = 6
grd4.Text = "«· «—ÌŒ"
End Sub
Private Sub chargegrd4()
On Error Resume Next
Dim i As Double
Dim tx As String
Dim k As Double
Dim tx2 As String
i = 1
k = 0
Call cont
grd4.Rows = ce.RecordCount + 3
Do While Not ce.EOF
grd1.row = 0
grd1.Col = 1
tx = grd1.Text
If tx = ce!cla Then
Text4.Text = ce!fra
Text5.Text = ce!man
End If
If Text3.Text = ce!ser Or Val(ce!ser) = Text3.Text Then
k = 1
tx2 = ce!cas
Label46.Caption = ce!dat
Label47.Caption = ce!cas
grd4.row = i
grd4.Col = 0
grd4.Text = ce!aut
grd4.Col = 1
grd4.Text = ce!moi
grd4.Col = 2
grd4.Text = ce!mon
grd4.Col = 3
grd4.Text = ce!pay
grd4.Col = 4
grd4.Text = ce!res
grd4.Col = 5
grd4.Text = ce!rec
grd4.Col = 6
grd4.Text = ce!dat
i = i + 1
End If
ce.MoveNext
Loop
grd4.Rows = i
grd4.Col = 0
grd4.Sort = 1
If k = 1 Then
Combo2.Text = tx2
End If
End Sub
Private Sub chargegrd4_rec()
On Error Resume Next
Dim i As Double
Dim tx As String
Dim k As Double
Dim tx2 As String
i = 1
k = 0
Call cont
grd4.Rows = ce.RecordCount + 3
Do While Not ce.EOF
grd1.row = 0
grd1.Col = 1
tx = grd1.Text
If tx = ce!cla Then
Text4.Text = ce!fra
Text5.Text = ce!man
End If
If Text3.Text = ce!ser Or Val(ce!ser) = Text3.Text Then
If ce!rec = Combo3.Text Then
k = 1
tx2 = ce!cas
Label46.Caption = ce!dat
Label47.Caption = ce!cas
grd4.row = i
grd4.Col = 0
grd4.Text = ce!aut
grd4.Col = 1
grd4.Text = ce!moi
grd4.Col = 2
grd4.Text = ce!mon
grd4.Col = 3
grd4.Text = ce!pay
grd4.Col = 4
grd4.Text = ce!res
grd4.Col = 5
grd4.Text = ce!rec
i = i + 1
End If
End If
ce.MoveNext
Loop
grd4.Rows = i
grd4.Col = 0
grd4.Sort = 1
If k = 1 Then
Combo2.Text = tx2
End If
End Sub
Private Sub totalmontant()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim a As Double
Dim s As Double
Dim r As Double
Dim tx1 As String
Dim tx2 As String
n = grd3.Rows
s = 0
tx2 = ""
For i = 1 To n - 1
grd3.row = i
grd3.Col = 0
tx1 = grd3.Text
grd3.Col = 2
a = grd3.Text
s = s + a
grd3.Col = 3
a = grd3.Text
r = r + a
If i = 1 Then
tx2 = tx1
Else
tx2 = tx2 + "+" + tx1
End If
Next i
Label75.Caption = tx2
Label2.Caption = s
Label45.Caption = r
End Sub
Private Sub chargegrd5()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim j As Double
Dim m As Double
Dim tx1 As String
Dim tx2 As String
Dim a As Double
Dim b As Double
Dim c As Double
Dim sa As Double
Dim sb As Double
Dim sc As Double
Dim cre As Double
Dim k As Double
n = grd2.Rows
m = grd4.Rows
grd5.Clear
grd5.Rows = n + 2
grd5.Cols = 2
grd5.ColWidth(0) = 2200
grd5.ColWidth(1) = 2200
grd5.row = 0
grd5.Col = 0
grd5.Text = "«·‘Â—"
grd5.Col = 1
grd5.Text = "«·œÌ‰"
grd5.ColAlignment(0) = 1
grd5.ColAlignment(1) = 1
k = 1
For i = 1 To n - 1
sa = 0
sb = 0
sc = 0
grd2.row = i
grd2.Col = 0
tx1 = grd2.Text
For j = 1 To m - 1
a = 0
b = 0
c = 0
grd4.row = j
grd4.Col = 1
tx2 = grd4.Text
If tx1 = tx2 Then
grd4.row = j
grd4.Col = 2
a = grd4.Text
grd4.Col = 3
b = grd4.Text
grd4.Col = 4
c = grd4.Text
sa = sa + a
sb = sb + b
sc = c
End If
Next j
cre = sa - sb
If cre <> 0 Then
grd5.row = k
grd5.Col = 0
grd5.Text = tx1
grd5.Col = 1
grd5.Text = cre
k = k + 1
End If
Next i
grd5.Rows = k
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
ProgressBar2.Value = ProgressBar2.Value + 8
If ProgressBar2.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
grd1.Visible = False
grd4.Clear
grd4.Rows = 1
grd3.Clear
grd3.Rows = 1
grd2.Visible = False
grd3.Visible = False
grd4.Visible = False
grd5.Visible = False
Label2.Caption = ""
Label45.Caption = ""
Label75.Caption = ""
Call chargec3
face.SBB1.Panels(16).Text = sr!cca
Call chargegrd4
'Call chargegrd2
Call chargegrd5
grd2.Visible = True
grd3.Visible = True
grd4.Visible = True
grd5.Visible = True
Picture6.Visible = True
ProgressBar2.Value = 0
ProgressBar2.Visible = False
Timer2.Enabled = False
grd1.Visible = True
End If


End Sub
Private Sub ivaa()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim n As Double
Dim m As Double
Dim mm As Double
Dim p As Double
Dim k As Double
Dim tx1 As String
Dim tx2 As String
n = grd2.Rows
m = grd4.Rows
For i = 1 To n - 1
grd2.row = i
grd2.Col = 0
tx1 = grd2.Text
grd2.Col = 2
mm = grd2.Text
k = 0
For j = 1 To m - 1
grd4.row = j
grd4.Col = 1
tx2 = grd4.Text
If tx1 = tx2 Then
k = 1
j = m
End If
Next j
If k = 0 Then
p = grd3.Rows
grd3.Rows = p + 1
grd3.row = p
grd3.Col = 0
grd3.Text = tx1
grd3.Col = 1
grd3.Text = "0"
grd3.Col = 2
grd3.Text = "0"
grd3.Col = 3
grd3.Text = "0"
grd3.Col = 4
grd3.Text = mm
Call totalmontant
End If
Next i
End Sub
Private Sub chargegrd7()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim k As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim a As Double
Dim s As Double
grd7.Rows = 1
grd7.Cols = 7
grd7.ColWidth(0) = 0
grd7.ColWidth(1) = 800
grd7.ColWidth(2) = 900
grd7.ColWidth(3) = 1000
grd7.ColWidth(4) = 1200
grd7.ColWidth(5) = 600
grd7.ColWidth(6) = 0
grd7.ColAlignment(0) = 1
grd7.ColAlignment(1) = 1
grd7.ColAlignment(2) = 1
grd7.ColAlignment(3) = 1
grd7.ColAlignment(4) = 1
grd7.ColAlignment(5) = 1
grd7.row = 0
grd7.Col = 1
grd7.Text = "«·Ê’·"
grd7.Col = 2
grd7.Text = "«·„»·€"
grd7.Col = 3
grd7.Text = "«·œ«›⁄"
grd7.Col = 4
grd7.Text = "«· «—ÌŒ"
grd7.Col = 5
grd7.Text = ""
'***** grd8
grd8.Rows = 1
grd8.Cols = 7
grd8.ColWidth(0) = 0
grd8.ColWidth(1) = 800
grd8.ColWidth(2) = 0
grd8.ColWidth(3) = 1000
grd8.ColWidth(4) = 1200
grd8.ColWidth(5) = 600
grd8.ColWidth(6) = 0
grd8.ColAlignment(0) = 1
grd8.ColAlignment(1) = 1
grd8.ColAlignment(2) = 1
grd8.ColAlignment(3) = 1
grd8.ColAlignment(4) = 1
grd8.ColAlignment(5) = 1
grd8.row = 0
grd8.Col = 1
grd8.Text = "«·Ê’·"
grd8.Col = 2
grd8.Text = "«·„»·€"
grd8.Col = 3
grd8.Text = "«·œ«›⁄"
grd8.Col = 4
grd8.Text = "«· «—ÌŒ"
grd8.Col = 5
grd8.Text = ""
'***** grd9
grd9.Rows = 1
grd9.Cols = 7
grd9.ColWidth(0) = 0
grd9.ColWidth(1) = 800
grd9.ColWidth(2) = 0
grd9.ColWidth(3) = 1000
grd9.ColWidth(4) = 1200
grd9.ColWidth(5) = 600
grd9.ColWidth(6) = 0
grd9.ColAlignment(0) = 1
grd9.ColAlignment(1) = 1
grd9.ColAlignment(2) = 1
grd9.ColAlignment(3) = 1
grd9.ColAlignment(4) = 1
grd9.ColAlignment(5) = 1
grd9.row = 0
grd9.Col = 1
grd9.Text = "«·Ê’·"
grd9.Col = 2
grd9.Text = "«·„»·€"
grd9.Col = 3
grd9.Text = "«·œ«›⁄"
grd9.Col = 4
grd9.Text = "«· «—ÌŒ"
grd9.Col = 5
grd9.Text = ""
i = 1
j = 1
k = 1
s = 0
dat1 = DT3.Value
dat2 = DT4.Value
Call cont
grd7.Rows = rc.RecordCount + 3
grd8.Rows = rc.RecordCount + 3
grd9.Rows = rc.RecordCount + 3
Do While Not rc.EOF
dat3 = rc!dat
If dat3 >= dat1 And dat3 <= dat2 Then
a = rc!mon
s = s + a
'**** grd7
If rc!cas = "Õ«·… ≈ﬂ„«·" Then
grd7.row = i
grd7.Col = 0
grd7.Text = rc!aut
grd7.Col = 1
grd7.Text = rc!rec
grd7.Col = 2
grd7.Text = rc!mon
grd7.Col = 3
grd7.Text = rc!ser
grd7.Col = 4
grd7.Text = rc!dat
grd7.Col = 5
grd7.Text = "Õ–›"
grd7.Col = 6
grd7.Text = rc!aut
i = i + 1
End If
'**** grd8
If rc!cas = "Õ«·… ≈⁄›«¡" Then
grd8.row = j
grd8.Col = 0
grd8.Text = rc!aut
grd8.Col = 1
grd8.Text = rc!rec
grd8.Col = 2
grd8.Text = rc!mon
grd8.Col = 3
grd8.Text = rc!ser
grd8.Col = 4
grd8.Text = rc!dat
grd8.Col = 5
grd8.Text = "Õ–›"
grd8.Col = 6
grd8.Text = rc!aut
j = j + 1
End If
'**** grd9
If rc!cas = "Õ«·… «‰”Õ«»" Then
grd9.row = k
grd9.Col = 0
grd9.Text = rc!aut
grd9.Col = 1
grd9.Text = rc!rec
grd9.Col = 2
grd9.Text = rc!mon
grd9.Col = 3
grd9.Text = rc!ser
grd9.Col = 4
grd9.Text = rc!dat
grd9.Col = 5
grd9.Text = "Õ–›"
grd9.Col = 6
grd9.Text = rc!aut
k = k + 1
End If
End If
rc.MoveNext
Loop
grd7.Rows = i
grd8.Rows = j
grd9.Rows = k
Label6.Caption = s
Label7.Caption = k - 1
Label8.Caption = j - 1
Label9.Caption = i - 1
grd7.Col = 1
grd7.Sort = 1
grd8.Col = 1
grd8.Sort = 1
grd9.Col = 1
grd9.Sort = 1
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
ProgressBar3.Value = ProgressBar3.Value + 8
If ProgressBar3.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
Command11_Click
Call cont
face.SBB1.Panels(16).Text = sr!cca
ProgressBar3.Value = 0
ProgressBar3.Visible = True
Timer3.Enabled = False
End If


End Sub
Private Sub chargegrd13()
On Error Resume Next
Dim i As Double
Dim cla1 As String
Dim cla2 As String
Dim mois1 As String
Dim mois2 As String
Dim n As Double
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim e As Double
Dim f As Double
Dim h As Double
Dim r As Double
Dim p As Double
Dim s As Double
Dim w As Double
Dim mh As Double
Dim sh As Double
Dim mm As Double
Dim sm As Double
Dim mp As Double
Dim sp As Double
grd13.Rows = 1
grd13.Cols = 11
grd13.ColWidth(0) = 500
grd13.ColWidth(1) = 600
grd13.ColWidth(2) = 900
grd13.ColWidth(3) = 900
grd13.ColWidth(4) = 900
grd13.ColWidth(5) = 900
grd13.ColWidth(6) = 900
grd13.ColWidth(7) = 500
grd13.ColWidth(8) = 900
grd13.ColWidth(9) = 500
grd13.ColWidth(10) = 900
grd13.ColAlignment(0) = 1
grd13.ColAlignment(1) = 1
grd13.ColAlignment(2) = 1
grd13.ColAlignment(3) = 1
grd13.ColAlignment(4) = 1
grd13.ColAlignment(5) = 1
grd13.ColAlignment(6) = 1
grd13.ColAlignment(7) = 1
grd13.ColAlignment(8) = 1
grd13.ColAlignment(9) = 1
grd13.ColAlignment(10) = 1
grd13.row = 0
grd13.Col = 0
grd13.Text = "«·‘Â—"
grd13.Col = 1
grd13.Text = "«·ﬁ”„"
grd13.Col = 2
grd13.Text = "«· ·«„Ì–"
grd13.Col = 3
grd13.Text = "«·√”« –…"
grd13.Col = 4
grd13.Text = "«·»«ﬁÌ"
grd13.Col = 5
grd13.Text = "«·„ƒ””…"
grd13.Col = 6
grd13.Text = "«·Õ’Ì·…"
grd13.Col = 7
grd13.Text = "”.ﬁ"
grd13.Col = 8
grd13.Text = "‰’Ì».”"
grd13.Col = 9
grd13.Text = "”.√"
grd13.Col = 10
grd13.Text = "«·„” Õﬁ"
i = 0
cla1 = ""
mois1 = ""
sh = 0
sm = 0
sp = 0
'**** charge mois, classes et nbr
Call cont
grd13.Rows = ps.RecordCount + 5
Do While Not ps.EOF
If ps!ser = Text7.Text Or Val(ps!ser) = Text7.Text Then
'**** heur
If ps!cas = "h" Then
mh = ps!tot
sh = sh + mh
mh = ps!prm
sh = sh + mh
mh = ps!rtr
sh = sh + mh
End If
'**** mois
If ps!cas = "m" Then
mm = ps!tot
sm = sm + mm
mm = ps!prm
sm = sm + mm
mm = ps!rtr
sm = sm + mm
End If
'**** grd13
If ps!cas = "p" Then
cla2 = ps!cla
mois2 = ps!mois
If cla1 <> cla2 Or mois1 <> mois2 Then
i = i + 1
End If
grd13.row = i
grd13.Col = 0
grd13.Text = ps!mois
grd13.Col = 1
grd13.Text = ps!cla
grd13.Col = 2
grd13.Text = "0"
grd13.Col = 3
grd13.Text = "0"
grd13.Col = 7
grd13.Text = "0"
grd13.Col = 9
grd13.Text = ps!nbr
cla1 = ps!cla
mois1 = ps!mois
End If
End If
ps.MoveNext
Loop
grd13.Rows = i + 1
'**** professeurs pourcentage
n = grd13.Rows
If n > 1 Then
'**** charge montants etudiants
c = 0
Call cont
Do While Not ce.EOF
For i = 1 To n - 1
grd13.row = i
grd13.Col = 0
mois1 = grd13.Text
grd13.Col = 1
cla1 = grd13.Text
grd13.Col = 2
a = grd13.Text
If ce!mois = mois1 And ce!cla = cla1 And ce!moi <> "«· ”ÃÌ·" Then
b = ce!pay
c = a + b
grd13.row = i
grd13.Col = 2
grd13.Text = c
End If
Next i
ce.MoveNext
Loop
'**** charge montants professeurs,nbr des professeurs
c = 0
r = 0
f = 0
Call cont
Do While Not ps.EOF
For i = 1 To n - 1
grd13.row = i
grd13.Col = 0
mois1 = grd13.Text
grd13.Col = 1
cla1 = grd13.Text
grd13.Col = 3
a = grd13.Text
grd13.Col = 7
d = grd13.Text
If ps!mois = mois1 And ps!cla = cla1 And ps!cas <> "p" Then
e = ps!tot
f = a + e
e = ps!prm
f = f + e
e = ps!rtr
f = f + e
grd13.row = i
grd13.Col = 3
grd13.Text = f
End If
If ps!mois = mois1 And ps!cla = cla1 And ps!cas = "p" Then
h = ps!nbr
r = d + h
grd13.row = i
grd13.Col = 7
grd13.Text = r
End If
Next i
ps.MoveNext
Loop
'***** calcule
a = 0
b = 0
c = 0
d = 0
e = 0
f = 0
h = 0
r = 0
p = 0
a = 0
b = 0
s = 0
For i = 1 To n - 1
grd13.row = i
grd13.Col = 2
a = grd13.Text
grd13.Col = 3
b = grd13.Text
c = a - b
d = Label18.Caption
e = (c * d) / 100
MyNumber = Round(e, 0)
e = MyNumber
f = c - e
grd13.row = i
grd13.Col = 7
h = grd13.Text
If h > 0 Then
r = f / h
w = f / h
MyNumber = Round(r, 0)
r = MyNumber
End If
grd13.row = i
grd13.Col = 9
p = grd13.Text
s = p * w
MyNumber = Round(s, 0)
s = MyNumber
grd13.row = i
grd13.Col = 4
grd13.Text = c
grd13.Col = 5
grd13.Text = e
grd13.Col = 6
grd13.Text = f
grd13.Col = 8
grd13.Text = r
grd13.Col = 10
grd13.Text = s
mp = s
sp = sp + mp
Next i
'grd13.Rows = 70
End If
Label11.Caption = sh
Label12.Caption = sm
Label13.Caption = sp
Label14.Caption = sh + sm + sp
End Sub

Private Sub Timer4_Timer()
On Error Resume Next
ProgressBar5.Value = ProgressBar5.Value + 8
If ProgressBar5.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
Text8.Text = ""
grd14.Visible = False
grd14.Clear
grd14.Rows = 1
Call chargegrd14
face.SBB1.Panels(16).Text = sr!cca
grd14.Visible = True
ProgressBar5.Value = 0
ProgressBar5.Visible = True
Timer4.Enabled = False
End If
End Sub
Private Sub chargegrd14()
On Error Resume Next
Dim i As Double
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
grd14.Rows = 1
grd14.Cols = 5
grd14.ColWidth(0) = 0
grd14.ColWidth(1) = 1200
grd14.ColWidth(2) = 1000
grd14.ColWidth(3) = 1100
grd14.ColWidth(4) = 1300
grd14.ColAlignment(1) = 1
grd14.ColAlignment(2) = 1
grd14.ColAlignment(3) = 1
grd14.ColAlignment(4) = 1
grd14.row = 0
grd14.Col = 0
grd14.Text = ""
grd14.Col = 1
grd14.Text = "«· «—ÌŒ"
grd14.Col = 2
grd14.Text = "«·”«⁄…"
grd14.Col = 3
grd14.Text = "«·„»·€"
grd14.Col = 4
grd14.Text = ""
i = 1
b = 0
Call cont
grd14.Rows = pf.RecordCount + 3
Do While Not pf.EOF
If pf!ser = Text7.Text Or Val(Text7.Text) = pf!ser Then
a = pf!mon
b = b + a
grd14.row = i
grd14.Col = 0
grd14.Text = pf!aut
grd14.Col = 1
grd14.Text = pf!dat
grd14.Col = 2
grd14.Text = pf!heu
grd14.Col = 3
grd14.Text = pf!mon
grd14.Col = 4
grd14.Text = "Õ–›"
i = i + 1
End If
pf.MoveNext
Loop
grd14.Rows = i
a = Label14.Caption
c = a - b
Label16.Caption = b
Label17.Caption = c
Command24.Enabled = False
End Sub
Private Sub chargegrd14_2()
On Error Resume Next
Dim i As Double
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
grd14.Rows = 1
grd14.Cols = 5
grd14.ColWidth(0) = 0
grd14.ColWidth(1) = 1200
grd14.ColWidth(2) = 1000
grd14.ColWidth(3) = 1100
grd14.ColWidth(4) = 1300
grd14.ColAlignment(1) = 1
grd14.ColAlignment(2) = 1
grd14.ColAlignment(3) = 1
grd14.ColAlignment(4) = 1
grd14.row = 0
grd14.Col = 0
grd14.Text = ""
grd14.Col = 1
grd14.Text = "«· «—ÌŒ"
grd14.Col = 2
grd14.Text = "«·”«⁄…"
grd14.Col = 3
grd14.Text = "«·„»·€"
grd14.Col = 4
grd14.Text = ""
i = 1
b = 0
dat1 = DT8.Value
dat2 = DT9.Value
Call cont
grd14.Rows = pf.RecordCount + 3
Do While Not pf.EOF
dat3 = pf!dat
If dat3 >= dat1 And dat3 <= dat2 Then
If pf!ser = Text7.Text Or Val(Text7.Text) = pf!ser Then
a = pf!mon
b = b + a
grd14.row = i
grd14.Col = 0
grd14.Text = pf!aut
grd14.Col = 1
grd14.Text = pf!dat
grd14.Col = 2
grd14.Text = pf!heu
grd14.Col = 3
grd14.Text = pf!mon
grd14.Col = 4
grd14.Text = "Õ–›"
i = i + 1
End If
End If
pf.MoveNext
Loop
grd14.Rows = i
Command24.Enabled = True
'a = Label14.Caption
'c = a - b
'Label16.Caption = b
'Label17.Caption = c
End Sub

Private Sub chargegrd15()
On Error Resume Next
Dim i As Double
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim n As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim ser1 As String
Dim k As Double
grd15.Rows = 1
grd15.Cols = 5
grd15.ColWidth(0) = 1500
grd15.ColWidth(1) = 5000
grd15.ColWidth(2) = 2200
grd15.ColWidth(3) = 2200
grd15.ColWidth(4) = 2200
grd15.ColAlignment(0) = 1
grd15.ColAlignment(1) = 1
grd15.ColAlignment(2) = 1
grd15.ColAlignment(3) = 1
grd15.ColAlignment(4) = 1
grd15.row = 0
grd15.Col = 0
grd15.Text = "«·—ﬁ„ «· ”·”·Ì"
grd15.Col = 1
grd15.Text = "«·«”„"
grd15.Col = 2
grd15.Text = "«·„” Õﬁ"
grd15.Col = 3
grd15.Text = "«·„œ›Ê⁄"
grd15.Col = 4
grd15.Text = "«·—’Ìœ"
i = 1
Call cont
grd15.Rows = pr.RecordCount + 3
Do While Not pr.EOF
grd15.row = i
grd15.Col = 0
grd15.Text = pr!ser
grd15.Col = 1
grd15.Text = pr!nom
grd15.Col = 2
grd15.Text = "0"
grd15.Col = 3
grd15.Text = "0"
grd15.Col = 4
grd15.Text = "0"
i = i + 1
pr.MoveNext
Loop
grd15.Rows = i
n = grd15.Rows
'***** dus
i = 1
b = 0
c = 0
dat1 = DT10.Value
dat2 = DT11.Value
Call cont
Do While Not ps.EOF
dat3 = ps!dat
For i = 1 To n - 1
b = 0
c = 0
If dat3 >= dat1 And dat3 <= dat2 Then
grd15.row = i
grd15.Col = 0
ser1 = grd15.Text
If ps!ser = ser1 Then
grd15.row = i
grd15.Col = 2
c = grd15.Text
a = ps!tot
b = b + a
a = ps!prm
b = b + a
a = ps!rtr
b = b + a
b = b + c
grd15.row = i
grd15.Col = 2
grd15.Text = b
End If
End If
Next i
ps.MoveNext
Loop
'***** dus
i = 1
b = 0
c = 0
dat1 = DT10.Value
dat2 = DT11.Value
Call cont
Do While Not pf.EOF
dat3 = pf!dat
For i = 1 To n - 1
b = 0
c = 0
If dat3 >= dat1 And dat3 <= dat2 Then
grd15.row = i
grd15.Col = 0
ser1 = grd15.Text
If pf!ser = ser1 Then
grd15.row = i
grd15.Col = 3
c = grd15.Text
a = pf!mon
b = b + a + c
grd15.row = i
grd15.Col = 3
grd15.Text = b
End If
End If
Next i
pf.MoveNext
Loop
k = 1
For i = 1 To n - 1
grd15.row = i
grd15.Col = 2
a = grd15.Text
grd15.Col = 3
b = grd15.Text
c = a - b
grd15.row = i
grd15.Col = 4
grd15.Text = c
If a = 0 Then
grd15.row = i
grd15.Col = 2
grd15.Text = "-----"
grd15.Col = 3
grd15.Text = "-----"
grd15.Col = 4
grd15.Text = "-----"
Else
k = k + 1
End If
Next i
grd15.Col = 2
grd15.Sort = 2
grd15.Rows = k
End Sub
Private Sub chargegrd16()
On Error Resume Next
Dim i As Double
grd16.Rows = 1
grd16.Cols = 3
grd16.ColWidth(0) = 0
grd16.ColWidth(1) = 3000
grd16.ColWidth(2) = 1500
grd16.ColAlignment(0) = 1
grd16.ColAlignment(1) = 1
grd16.ColAlignment(2) = 1
grd16.row = 0
grd16.Col = 0
grd16.Text = ""
grd16.Col = 1
grd16.Text = "«·«”„"
grd16.Col = 2
grd16.Text = "—ﬁ„ «·Â« ›"
i = 1
Call cont
grd16.Rows = pa.RecordCount + 3
Do While Not pa.EOF
grd16.row = i
grd16.Col = 0
grd16.Text = pa!mtr
grd16.Col = 1
grd16.Text = pa!nom
grd16.Col = 2
grd16.Text = pa!tel
i = i + 1
pa.MoveNext
Loop
grd16.Rows = i
End Sub
Private Sub chargegrd17_18()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim p As Double
Dim sp As Double
Dim r As Double
Dim sr As Double
Dim cre As Double
grd17.Rows = 1
grd17.Cols = 4
grd17.ColWidth(0) = 0
grd17.ColWidth(1) = 1500
grd17.ColWidth(2) = 1500
grd17.ColWidth(3) = 1000
grd17.ColAlignment(0) = 1
grd17.ColAlignment(1) = 1
grd17.ColAlignment(2) = 1
grd17.row = 0
grd17.Col = 0
grd17.Text = ""
grd17.Col = 1
grd17.Text = "«· «—ÌŒ"
grd17.Col = 2
grd17.Text = "«·„»·€"
grd18.Rows = 1
grd18.Cols = 4
grd18.ColWidth(0) = 0
grd18.ColWidth(1) = 1500
grd18.ColWidth(2) = 1500
grd18.ColWidth(3) = 1000
grd18.ColAlignment(0) = 1
grd18.ColAlignment(1) = 1
grd18.ColAlignment(2) = 1
grd18.row = 0
grd18.Col = 0
grd18.Text = ""
grd18.Col = 1
grd18.Text = "«· «—ÌŒ"
grd18.Col = 2
grd18.Text = "«·„»·€"
i = 1
j = 1
sp = 0
sr = 0
Call cont
grd17.Rows = pp.RecordCount + 3
grd18.Rows = pp.RecordCount + 3
Do While Not pp.EOF
If Label28.Caption = pp!mtr Then
If pp!Mod = "«·’‰œÊﬁ" Then
r = pp!mon
sr = sr + r
grd17.row = i
grd17.Col = 0
grd17.Text = pp!aut
grd17.Col = 1
grd17.Text = pp!dat
grd17.Col = 2
grd17.Text = pp!mon
grd17.Col = 3
grd17.Text = "Õ–›"
i = i + 1
End If
If pp!Mod = "«·‘—Ìﬂ" Then
p = pp!mon
sp = sp + p
grd18.row = j
grd18.Col = 0
grd18.Text = pp!aut
grd18.Col = 1
grd18.Text = pp!dat
grd18.Col = 2
grd18.Text = pp!mon
grd18.Col = 3
grd18.Text = "Õ–›"
j = j + 1
End If
End If
pp.MoveNext
Loop
grd17.Rows = i
grd18.Rows = j
cre = sp - sr
Label24.Caption = sp
Label25.Caption = sr
If cre < 0 Then
Label31(52).Caption = "„œÌ‰ »‹ "
Label26.Caption = cre * -1
Else
Label31(52).Caption = "œ«∆‰ »‹ "
Label26.Caption = cre
End If
If cre = 0 Then
Label31(52).Caption = "·« „œÌ‰ Ê·« œ«∆‰"
Label26.Caption = ""
End If
End Sub

Private Sub chargegrd17_18_2()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim p As Double
Dim sp As Double
Dim r As Double
Dim sr As Double
Dim cre As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
grd17.Rows = 1
grd17.Cols = 4
grd17.ColWidth(0) = 0
grd17.ColWidth(1) = 1500
grd17.ColWidth(2) = 1500
grd17.ColWidth(3) = 1000
grd17.ColAlignment(0) = 1
grd17.ColAlignment(1) = 1
grd17.ColAlignment(2) = 1
grd17.row = 0
grd17.Col = 0
grd17.Text = ""
grd17.Col = 1
grd17.Text = "«· «—ÌŒ"
grd17.Col = 2
grd17.Text = "«·„»·€"
grd18.Rows = 1
grd18.Cols = 4
grd18.ColWidth(0) = 0
grd18.ColWidth(1) = 1500
grd18.ColWidth(2) = 1500
grd18.ColWidth(3) = 1000
grd18.ColAlignment(0) = 1
grd18.ColAlignment(1) = 1
grd18.ColAlignment(2) = 1
grd18.row = 0
grd18.Col = 0
grd18.Text = ""
grd18.Col = 1
grd18.Text = "«· «—ÌŒ"
grd18.Col = 2
grd18.Text = "«·„»·€"
i = 1
j = 1
sp = 0
sr = 0
dat1 = DT13.Value
dat2 = DT14.Value
Call cont
grd17.Rows = pp.RecordCount + 3
grd18.Rows = pp.RecordCount + 3
Do While Not pp.EOF
dat3 = pp!dat
If dat3 >= dat1 And dat3 <= dat2 Then
If Label28.Caption = pp!mtr Then
If pp!Mod = "«·’‰œÊﬁ" Then
r = pp!mon
sr = sr + r
grd17.row = i
grd17.Col = 0
grd17.Text = pp!aut
grd17.Col = 1
grd17.Text = pp!dat
grd17.Col = 2
grd17.Text = pp!mon
i = i + 1
End If
If pp!Mod = "«·‘—Ìﬂ" Then
p = pp!mon
sp = sp + p
grd18.row = j
grd18.Col = 0
grd18.Text = pp!aut
grd18.Col = 1
grd18.Text = pp!dat
grd18.Col = 2
grd18.Text = pp!mon
j = j + 1
End If
End If
End If
pp.MoveNext
Loop
grd17.Rows = i
grd18.Rows = j
cre = sp - sr
Label24.Caption = sp
Label25.Caption = sr
If cre < 0 Then
Label31(52).Caption = "„œÌ‰ »‹ "
Label26.Caption = cre * -1
Else
Label31(52).Caption = "œ«∆‰ »‹ "
Label26.Caption = cre
End If
If cre = 0 Then
Label31(52).Caption = "·« „œÌ‰ Ê·« œ«∆‰"
Label26.Caption = ""
End If
End Sub

Private Sub Timer5_Timer()
On Error Resume Next
ProgressBar4.Value = ProgressBar4.Value + 8
If ProgressBar4.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
Command28.Enabled = False
Call chargec10
face.SBB1.Panels(16).Text = sr!cca
Text11.Text = ""
Text11.SetFocus
grd17.Visible = False
grd17.Clear
grd17.Rows = 1
grd18.Visible = False
grd18.Clear
grd18.Rows = 1
Call chargegrd17_18
grd17.Visible = True
grd18.Visible = True
ProgressBar4.Value = 0
ProgressBar4.Visible = True
Timer5.Enabled = False
End If

End Sub

Private Sub Timer6_Timer()
On Error Resume Next
ProgressBar6.Value = ProgressBar6.Value + 8
If ProgressBar6.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
Command19.Enabled = False
Text9.Text = ""
Text10.Text = ""
Text9.SetFocus
grd19.Visible = False
grd19.Clear
grd19.Rows = 1
Call chargegrd19
face.SBB1.Panels(16).Text = sr!cca
grd19.Visible = True
ProgressBar6.Value = 0
ProgressBar6.Visible = True
Timer6.Enabled = False
End If

End Sub
Private Sub chargegrd19()
On Error Resume Next
Dim i As Double
Dim a As Double
Dim s As Double
grd19.Rows = 1
grd19.Cols = 7
grd19.ColWidth(0) = 0
grd19.ColWidth(1) = 4700
grd19.ColWidth(2) = 2000
grd19.ColWidth(3) = 3000
grd19.ColWidth(4) = 2000
grd19.ColWidth(5) = 1200
grd19.ColWidth(6) = 800
grd19.ColAlignment(0) = 1
grd19.ColAlignment(1) = 1
grd19.ColAlignment(2) = 1
grd19.ColAlignment(3) = 1
grd19.ColAlignment(4) = 1
grd19.ColAlignment(5) = 1
grd19.ColAlignment(6) = 1
grd19.row = 0
grd19.Col = 0
grd19.Text = ""
grd19.Col = 1
grd19.Text = "»Ì«‰ «·⁄„·Ì…"
grd19.Col = 2
grd19.Text = "«·„»·€ «·„’—Ê›"
grd19.Col = 3
grd19.Text = "Õ”«» «·⁄„·Ì…"
grd19.Col = 4
grd19.Text = "«· «—ÌŒ"
grd19.Col = 5
grd19.Text = "«·”«⁄…"
i = 1
s = 0
a = 0
Call cont
grd19.Rows = dp.RecordCount + 3
Do While Not dp.EOF
a = dp!mon
s = s + a
grd19.row = i
grd19.Col = 0
grd19.Text = dp!aut
grd19.Col = 1
grd19.Text = dp!dec
grd19.Col = 2
grd19.Text = dp!mon
grd19.Col = 3
grd19.Text = dp!com
grd19.Col = 4
grd19.Text = dp!dat
grd19.Col = 5
grd19.Text = dp!heu
grd19.Col = 6
grd19.Text = "Õ–›"
i = i + 1
dp.MoveNext
Loop
grd19.Rows = i
Label43.Caption = s
End Sub

Private Sub chargegrd19_2()
On Error Resume Next
Dim i As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim a As Double
Dim s As Double
grd19.Rows = 1
grd19.Cols = 7
grd19.ColWidth(0) = 0
grd19.ColWidth(1) = 4700
grd19.ColWidth(2) = 2000
grd19.ColWidth(3) = 3000
grd19.ColWidth(4) = 2000
grd19.ColWidth(5) = 1200
grd19.ColWidth(6) = 800
grd19.ColAlignment(0) = 1
grd19.ColAlignment(1) = 1
grd19.ColAlignment(2) = 1
grd19.ColAlignment(3) = 1
grd19.ColAlignment(4) = 1
grd19.ColAlignment(5) = 1
grd19.ColAlignment(6) = 1
grd19.row = 0
grd19.Col = 0
grd19.Text = ""
grd19.Col = 1
grd19.Text = "»Ì«‰ «·⁄„·Ì…"
grd19.Col = 2
grd19.Text = "«·„»·€ «·„’—Ê›"
grd19.Col = 3
grd19.Text = "Õ”«» «·⁄„·Ì…"
grd19.Col = 4
grd19.Text = "«· «—ÌŒ"
grd19.Col = 5
grd19.Text = "«·”«⁄…"
i = 1
dat1 = DT16.Value
dat2 = DT17.Value
a = 0
s = 0
Call cont
grd19.Rows = dp.RecordCount + 3
Do While Not dp.EOF
dat3 = dp!dat
If dat3 >= dat1 And dat3 <= dat2 Then
a = dp!mon
s = s + a
grd19.row = i
grd19.Col = 0
grd19.Text = dp!aut
grd19.Col = 1
grd19.Text = dp!dec
grd19.Col = 2
grd19.Text = dp!mon
grd19.Col = 3
grd19.Text = dp!com
grd19.Col = 4
grd19.Text = dp!dat
grd19.Col = 5
grd19.Text = dp!heu
grd19.Col = 6
grd19.Text = "Õ–›"
i = i + 1
End If
dp.MoveNext
Loop
grd19.Rows = i
Label43.Caption = s
End Sub

Private Sub Timer7_Timer()
On Error Resume Next
ProgressBar7.Value = ProgressBar7.Value + 8
If ProgressBar7.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
Command30.Enabled = False
Text13.Text = ""
Text12.Text = ""
Text13.SetFocus
grd20.Visible = False
grd20.Clear
grd20.Rows = 1
grd21.Visible = False
grd21.Clear
grd21.Rows = 1
Call chargegrd20_21
face.SBB1.Panels(16).Text = sr!cca
grd20.Visible = True
grd21.Visible = True
ProgressBar7.Value = 0
ProgressBar7.Visible = True
Timer7.Enabled = False
End If

End Sub
Private Sub chargegrd20_21()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim p As Double
Dim sp As Double
Dim r As Double
Dim sr As Double
Dim cre As Double
grd20.Rows = 1
grd20.Cols = 6
grd20.ColWidth(0) = 0
grd20.ColWidth(1) = 3000
grd20.ColWidth(2) = 1500
grd20.ColWidth(3) = 1300
grd20.ColWidth(4) = 700
grd20.ColWidth(5) = 0
grd20.ColAlignment(0) = 1
grd20.ColAlignment(1) = 1
grd20.ColAlignment(2) = 1
grd20.ColAlignment(3) = 1
grd20.ColAlignment(4) = 1
grd20.row = 0
grd20.Col = 0
grd20.Text = ""
grd20.Col = 1
grd20.Text = "»Ì«‰ «·⁄„·Ì…"
grd20.Col = 2
grd20.Text = "«·„»·€"
grd20.Col = 3
grd20.Text = "«· «—ÌŒ"
grd21.Rows = 1
grd21.Cols = 6
grd21.ColWidth(0) = 0
grd21.ColWidth(1) = 3000
grd21.ColWidth(2) = 1500
grd21.ColWidth(3) = 1300
grd21.ColWidth(4) = 700
grd21.ColWidth(5) = 0
grd21.ColAlignment(0) = 1
grd21.ColAlignment(1) = 1
grd21.ColAlignment(2) = 1
grd21.ColAlignment(3) = 1
grd21.ColAlignment(4) = 1
grd21.row = 0
grd21.Col = 0
grd21.Text = ""
grd21.Col = 1
grd21.Text = "»Ì«‰ «·⁄„·Ì…"
grd21.Col = 2
grd21.Text = "«·„»·€"
grd21.Col = 3
grd21.Text = "«· «—ÌŒ"
i = 1
j = 1
sp = 0
sr = 0
Call cont
grd20.Rows = bn.RecordCount + 3
grd21.Rows = bn.RecordCount + 3
Do While Not bn.EOF
If bn!Mod = "”Õ»" Then
r = bn!mon
sr = sr + r
grd20.row = i
grd20.Col = 0
grd20.Text = bn!aut
grd20.Col = 1
grd20.Text = bn!dec
grd20.Col = 2
grd20.Text = bn!mon
grd20.Col = 3
grd20.Text = bn!dat
grd20.Col = 4
grd20.Text = "Õ–›"
grd20.Col = 5
grd20.Text = bn!act
i = i + 1
End If
If bn!Mod = "«Ìœ«⁄" Or bn!Mod = "—√” «·„«·" Then
p = bn!mon
sp = sp + p
grd21.row = j
grd21.Col = 0
grd21.Text = bn!aut
grd21.Col = 1
grd21.Text = bn!dec
grd21.Col = 2
grd21.Text = bn!mon
grd21.Col = 3
grd21.Text = bn!dat
grd21.Col = 4
grd21.Text = "Õ–›"
grd21.Col = 5
grd21.Text = bn!act
j = j + 1
End If
bn.MoveNext
Loop
grd20.Rows = i
grd21.Rows = j
cre = sp - sr
Label33.Caption = sp
Label32.Caption = sr
Label30.Caption = cre
End Sub
Private Sub chargegrd20_21_2()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim p As Double
Dim sp As Double
Dim r As Double
Dim sr As Double
Dim cre As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
grd20.Rows = 1
grd20.Cols = 6
grd20.ColWidth(0) = 0
grd20.ColWidth(1) = 3000
grd20.ColWidth(2) = 1500
grd20.ColWidth(3) = 1300
grd20.ColWidth(4) = 700
grd20.ColWidth(5) = 0
grd20.ColAlignment(0) = 1
grd20.ColAlignment(1) = 1
grd20.ColAlignment(2) = 1
grd20.ColAlignment(3) = 1
grd20.ColAlignment(4) = 1
grd20.row = 0
grd20.Col = 0
grd20.Text = ""
grd20.Col = 1
grd20.Text = "»Ì«‰ «·⁄„·Ì…"
grd20.Col = 2
grd20.Text = "«·„»·€"
grd20.Col = 3
grd20.Text = "«· «—ÌŒ"
grd21.Rows = 1
grd21.Cols = 6
grd21.ColWidth(0) = 0
grd21.ColWidth(1) = 3000
grd21.ColWidth(2) = 1500
grd21.ColWidth(3) = 1300
grd21.ColWidth(4) = 700
grd21.ColWidth(5) = 0
grd21.ColAlignment(0) = 1
grd21.ColAlignment(1) = 1
grd21.ColAlignment(2) = 1
grd21.ColAlignment(3) = 1
grd21.ColAlignment(4) = 1
grd21.row = 0
grd21.Col = 0
grd21.Text = ""
grd21.Col = 1
grd21.Text = "»Ì«‰ «·⁄„·Ì…"
grd21.Col = 2
grd21.Text = "«·„»·€"
grd21.Col = 3
grd21.Text = "«· «—ÌŒ"
i = 1
j = 1
sp = 0
sr = 0
dat1 = DT19.Value
dat2 = DT20.Value
Call cont
grd20.Rows = bn.RecordCount + 3
grd21.Rows = bn.RecordCount + 3
Do While Not bn.EOF
dat3 = bn!dat
If dat3 >= dat1 And dat3 <= dat2 Then
If bn!Mod = "”Õ»" Then
r = bn!mon
sr = sr + r
grd20.row = i
grd20.Col = 0
grd20.Text = bn!aut
grd20.Col = 1
grd20.Text = bn!dec
grd20.Col = 2
grd20.Text = bn!mon
grd20.Col = 3
grd20.Text = bn!dat
grd20.Col = 4
grd20.Text = "Õ–›"
grd20.Col = 5
grd20.Text = bn!act
i = i + 1
End If
If bn!Mod = "«Ìœ«⁄" Or bn!Mod = "—√” «·„«·" Then
p = bn!mon
sp = sp + p
grd21.row = j
grd21.Col = 0
grd21.Text = bn!aut
grd21.Col = 1
grd21.Text = bn!dec
grd21.Col = 2
grd21.Text = bn!mon
grd21.Col = 3
grd21.Text = bn!dat
grd21.Col = 4
grd21.Text = "Õ–›"
grd21.Col = 5
grd21.Text = bn!act
j = j + 1
End If
End If
bn.MoveNext
Loop
grd20.Rows = i
grd21.Rows = j
cre = sp - sr
Label33.Caption = sp
Label32.Caption = sr
Label30.Caption = cre
End Sub
Private Sub chargegrd22_23()
On Error Resume Next
Dim txtCasD As String
Dim txtCasK As String
Dim i As Double
Dim j As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim e1 As Double
Dim se1 As Double
Dim f1 As Double
Dim sf1 As Double
Dim p1 As Double
Dim sp1 As Double
Dim d1 As Double
Dim sd1 As Double
Dim b1 As Double
Dim sb1 As Double
Dim w1 As Double
Dim sw1 As Double
Dim s1 As Double
Dim e2 As Double
Dim se2 As Double
Dim f2 As Double
Dim sf2 As Double
Dim p2 As Double
Dim sp2 As Double
Dim d2 As Double
Dim sd2 As Double
Dim b2 As Double
Dim sb2 As Double
Dim w2 As Double
Dim sw2 As Double
Dim s2 As Double
Dim s As Double
Dim en As Double
Dim sen As Double
Dim sr As Double
Dim ssr As Double
Dim o1 As Double
Dim so1 As Double
Dim o2 As Double
Dim so2 As Double
grd22.Rows = 1
grd22.Cols = 12
grd22.ColWidth(0) = 0
grd22.ColWidth(1) = 4400
grd22.ColWidth(2) = 1200
grd22.ColWidth(3) = 1000
grd22.ColWidth(4) = 2000
grd22.ColWidth(5) = 700
grd22.ColWidth(6) = 1300
grd22.ColWidth(7) = 1000
grd22.ColWidth(8) = 1200
grd22.ColWidth(9) = 700
grd22.ColWidth(10) = 0
grd22.ColWidth(11) = 0
grd22.ColAlignment(0) = 1
grd22.ColAlignment(1) = 1
grd22.ColAlignment(2) = 1
grd22.ColAlignment(3) = 1
grd22.ColAlignment(4) = 1
grd22.ColAlignment(5) = 1
grd22.ColAlignment(6) = 1
grd22.ColAlignment(7) = 1
grd22.ColAlignment(8) = 1
grd22.ColAlignment(9) = 1
grd22.row = 0
grd22.Col = 0
grd22.Text = ""
grd22.Col = 1
grd22.Text = "»Ì«‰ «·⁄„·Ì…"
grd22.Col = 2
grd22.Text = "«·„»·€"
grd22.Col = 3
grd22.Text = "«·ﬂÊœ"
grd22.Col = 4
grd22.Text = "«·„”„Ï"
grd22.Col = 5
grd22.Text = "«·Õ«·…"
grd22.Col = 6
grd22.Text = "«· «—ÌŒ"
grd22.Col = 7
grd22.Text = "«·”«⁄…"
grd22.Col = 8
grd22.Text = "«·„‰›–"
grd22.Col = 9
grd22.Text = ""
grd23.Rows = 9
grd23.Cols = 4
grd23.ColWidth(0) = 3400
grd23.ColWidth(1) = 2800
grd23.ColWidth(2) = 2800
grd23.ColWidth(3) = 2800
grd23.ColWidth(4) = 2800
grd23.ColAlignment(0) = 1
grd23.ColAlignment(1) = 1
grd23.ColAlignment(2) = 1
grd23.ColAlignment(3) = 1
grd23.row = 0
grd23.Col = 0
grd23.Text = "«·Õ”«»"
grd23.Col = 1
grd23.Text = "„Ã„Ê⁄ «·œ«Œ·"
grd23.Col = 2
grd23.Text = "„Ã„Ê⁄ «·Œ«—Ã"
grd23.Col = 3
grd23.Text = "«·—’Ìœ"
grd23.Col = 0
grd23.row = 1
grd23.Text = "Õ”«» «· ·«„Ì–"
grd23.row = 2
grd23.Text = "Õ”«» «·√”« –…"
grd23.row = 3
grd23.Text = "Õ”«» «·‘—ﬂ«¡"
grd23.row = 4
grd23.Text = "Õ”«» «·⁄„«·"
grd23.row = 5
grd23.Text = "Õ”«» «·„’—Ê›« "
grd23.row = 6
grd23.Text = "Õ”«» «·»‰ﬂ"
grd23.row = 7
grd23.Text = "«· Ê“Ì⁄"
grd23.row = 8
grd23.Text = "«·„Ã„Ê⁄"
i = 1
j = 1
dat1 = DT21.Value
dat2 = DT22.Value
txtCasD = "œ«Œ·"
txtCasK = "Œ«—Ã"
If Combo14.Text = "«·œ«Œ·" Then
txtCasK = "œ«Œ·"
End If
If Combo14.Text = "«·Œ«—Ã" Then
txtCasD = "Œ«—Ã"
End If
se1 = 0
se2 = 0
sf1 = 0
sf2 = 0
sp1 = 0
sp2 = 0
sd1 = 0
sd2 = 0
sb1 = 0
sb2 = 0
sw1 = 0
sw2 = 0
sen = 0
ssr = 0
Call cont
grd22.Rows = ca.RecordCount + 5
Do While Not ca.EOF
dat3 = ca!dat
If dat3 < dat1 Then
If ca!cas = "œ«Œ·" Then
en = ca!mon
sen = sen + en
Else
sr = ca!mon
ssr = ssr + sr
End If
End If
If txtCasK = ca!cas Or txtCasD = ca!cas Then
If dat3 >= dat1 And dat3 <= dat2 Then
grd22.row = i
grd22.Col = 0
grd22.Text = ca!aut
grd22.Col = 1
grd22.Text = ca!mem
grd22.Col = 2
grd22.Text = ca!mon
grd22.Col = 3
grd22.Text = ca!cod
grd22.Col = 4
grd22.Text = ca!dec
grd22.Col = 5
grd22.Text = ca!cas
grd22.Col = 6
grd22.Text = ca!dat
grd22.Col = 7
grd22.Text = ca!heu
grd22.Col = 8
grd22.Text = ca!ger
grd22.Col = 9
grd22.Text = "Õ–›"
grd22.Col = 10
grd22.Text = ca!com
grd22.Col = 11
grd22.Text = ca!act
i = i + 1
If ca!com = "Õ”«» «· ·«„Ì–" Then
If ca!cas = "œ«Œ·" Then
e1 = ca!mon
se1 = se1 + e1
Else
e2 = ca!mon
se2 = se2 + e2
End If
End If
If ca!com = "Õ”«» «·√”« –…" Then
If ca!cas = "œ«Œ·" Then
f1 = ca!mon
sf1 = sf1 + f1
Else
f2 = ca!mon
sf2 = sf2 + f2
End If
End If
If ca!com = "Õ”«» «·‘—ﬂ«¡" Then
If ca!cas = "œ«Œ·" Then
p1 = ca!mon
sp1 = sp1 + p1
Else
p2 = ca!mon
sp2 = sp2 + p2
End If
End If
If ca!com = "Õ”«» «·⁄„«·" Then
If ca!cas = "œ«Œ·" Then
o1 = ca!mon
so1 = so1 + o1
Else
o2 = ca!mon
so2 = so2 + o2
End If
End If
If ca!com = "«·„’—Ê›« " Then
If ca!cas = "œ«Œ·" Then
d1 = ca!mon
sd1 = sd1 + d1
Else
d2 = ca!mon
sd2 = sd2 + d2
End If
End If
If ca!com = "«·»‰ﬂ" Then
If ca!cas = "œ«Œ·" Then
b1 = ca!mon
sb1 = sb1 + b1
Else
b2 = ca!mon
sb2 = sb2 + b2
End If
End If
If ca!com = "«· Ê“Ì⁄" Then
If ca!cas = "œ«Œ·" Then
w1 = ca!mon
sw1 = sw1 + w1
Else
w2 = ca!mon
sw2 = sw2 + w2
End If
End If
End If
End If
ca.MoveNext
Loop
grd22.Rows = i
s1 = se1 + sf1 + sp1 + sd1 + sb1 + sw1 + so1
s2 = se2 + sf2 + sp2 + sd2 + sb2 + sw2 + so2
s = s1 - s2
Label34.Caption = s1
Label35.Caption = s2
Label42.Caption = sen - ssr
Label36.Caption = s + (sen - ssr)
grd23.row = 1
grd23.Col = 1
grd23.Text = se1
grd23.Col = 2
grd23.Text = se2
grd23.Col = 3
grd23.Text = se1 - se2
grd23.row = 2
grd23.Col = 1
grd23.Text = sf1
grd23.Col = 2
grd23.Text = sf2
grd23.Col = 3
grd23.Text = sf1 - sf2
grd23.row = 3
grd23.Col = 1
grd23.Text = sp1
grd23.Col = 2
grd23.Text = sp2
grd23.Col = 3
grd23.Text = sp1 - sp2
grd23.row = 4
grd23.Col = 1
grd23.Text = so1
grd23.Col = 2
grd23.Text = so2
grd23.Col = 3
grd23.Text = so1 - so2
grd23.row = 5
grd23.Col = 1
grd23.Text = sd1
grd23.Col = 2
grd23.Text = sd2
grd23.Col = 3
grd23.Text = sd1 - sd2
grd23.row = 6
grd23.Col = 1
grd23.Text = sb1
grd23.Col = 2
grd23.Text = sb2
grd23.Col = 3
grd23.Text = sb1 - sb2
grd23.row = 7
grd23.Col = 1
grd23.Text = sw1
grd23.Col = 2
grd23.Text = sw2
grd23.Col = 3
grd23.Text = sw1 - sw2
grd23.row = 8
grd23.Col = 1
grd23.Text = s1
grd23.Col = 2
grd23.Text = s2
grd23.Col = 3
grd23.Text = s
End Sub

Private Sub Timer8_Timer()
On Error Resume Next
ProgressBar8.Value = ProgressBar8.Value + 8
If ProgressBar8.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
grd22.Visible = False
grd22.Clear
grd22.Rows = 1
grd23.Visible = False
grd23.Clear
grd23.Rows = 1
Call chargegrd22_23
face.SBB1.Panels(16).Text = sr!cca
grd22.Visible = True
grd23.Visible = True
ProgressBar8.Value = 0
ProgressBar8.Visible = True
Timer8.Enabled = False
End If

End Sub

Private Sub Timer9_Timer()
On Error Resume Next
Dim ane As String
ProgressBar9.Value = ProgressBar9.Value + 8
If ProgressBar9.Value > 90 Then
Timer9.Enabled = False
ProgressBar9.Value = 0
Call cont2
'******* etu non paye
If tim = 1 Then
'End If
'FileCopy "C:\photos\4.jpg", "C:\photos\2.jpg"
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
'data.DoCmd.ApplyFilter cla
data.DoCmd.Maximize
'data.DoCmd.OpenReport "Bulletin", acViewPreview, , "num =" & a, acWindowNormal, OpenArgs
data.DoCmd.OpenReport "Tetudpaspaye", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
'data.DCount
    'data.CloseCurrentDatabase
data.Visible = True
Set data = Nothing
Command13.Enabled = True
End If
'******* prof paye
If tim = 2 Then
'End If
'FileCopy "C:\photos\4.jpg", "C:\photos\2.jpg"
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
'data.DoCmd.ApplyFilter cla
data.DoCmd.Maximize
'data.DoCmd.OpenReport "Bulletin", acViewPreview, , "num =" & a, acWindowNormal, OpenArgs
data.DoCmd.OpenReport "Tprofpaspaye", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
'data.DCount
    'data.CloseCurrentDatabase
data.Visible = True
Set data = Nothing
Command13.Enabled = True
End If
'******* porcentage prof
If tim = 3 Then
'End If
'FileCopy "C:\photos\4.jpg", "C:\photos\2.jpg"
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
'data.DoCmd.ApplyFilter cla
data.DoCmd.Maximize
'data.DoCmd.OpenReport "Bulletin", acViewPreview, , "num =" & a, acWindowNormal, OpenArgs
data.DoCmd.OpenReport "Tprofpourcentage", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
'data.DCount
    'data.CloseCurrentDatabase
data.Visible = True
Set data = Nothing
Command15.Enabled = True
End If
'******* porcentage prof
If tim = 4 Then
'End If
'FileCopy "C:\photos\4.jpg", "C:\photos\2.jpg"
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
'data.DoCmd.ApplyFilter cla
data.DoCmd.Maximize
'data.DoCmd.OpenReport "Bulletin", acViewPreview, , "num =" & a, acWindowNormal, OpenArgs
data.DoCmd.OpenReport "Tpayprof", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
'data.DCount
    'data.CloseCurrentDatabase
data.Visible = True
Set data = Nothing
Command17.Enabled = True
End If
If tim = 5 Then
'End If
'FileCopy "C:\photos\4.jpg", "C:\photos\2.jpg"
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
'data.DoCmd.ApplyFilter cla
data.DoCmd.Maximize
'data.DoCmd.OpenReport "Bulletin", acViewPreview, , "num =" & a, acWindowNormal, OpenArgs
data.DoCmd.OpenReport "Trecpaypartenaires", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
'data.DCount
    'data.CloseCurrentDatabase
data.Visible = True
Set data = Nothing
Command28.Enabled = True
End If
If tim = 6 Then
'End If
'FileCopy "C:\photos\4.jpg", "C:\photos\2.jpg"
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
'data.DoCmd.ApplyFilter cla
data.DoCmd.Maximize
'data.DoCmd.OpenReport "Bulletin", acViewPreview, , "num =" & a, acWindowNormal, OpenArgs
data.DoCmd.OpenReport "Tdepenses", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
'data.DCount
    'data.CloseCurrentDatabase
data.Visible = True
Set data = Nothing
Command19.Enabled = True
End If
If tim = 7 Then
'End If
'FileCopy "C:\photos\4.jpg", "C:\photos\2.jpg"
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
'data.DoCmd.ApplyFilter cla
data.DoCmd.Maximize
'data.DoCmd.OpenReport "Bulletin", acViewPreview, , "num =" & a, acWindowNormal, OpenArgs
data.DoCmd.OpenReport "Tbanks", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
'data.DCount
    'data.CloseCurrentDatabase
data.Visible = True
Set data = Nothing
Command30.Enabled = True
End If
If tim = 8 Then
'End If
'FileCopy "C:\photos\4.jpg", "C:\photos\2.jpg"
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
'data.DoCmd.ApplyFilter cla
data.DoCmd.Maximize
'data.DoCmd.OpenReport "Bulletin", acViewPreview, , "num =" & a, acWindowNormal, OpenArgs
data.DoCmd.OpenReport "caisses", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
'data.DCount
    'data.CloseCurrentDatabase
data.Visible = True
Set data = Nothing
Command32.Enabled = True
End If




End If
End Sub
Private Sub monauj()
Dim a As Double
Dim b As Double
Dim c As Double
a = Label34.Caption
b = Label35.Caption
c = (a - b)
Label44.Caption = c
End Sub
