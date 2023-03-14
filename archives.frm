VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form archives 
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
      TabCaption(0)   =   "«·‰”Œ «·«Õ Ì«ÿÌ"
      TabPicture(0)   =   "archives.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "«” —Ã«⁄ »Ì«‰«  „‰ «·«—‘Ì›"
      TabPicture(1)   =   "archives.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "„”Õ ”‰… „‰ «·«—‘Ì›"
      TabPicture(2)   =   "archives.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "⁄—÷ ”‰… „‰ «·«—‘Ì›"
      TabPicture(3)   =   "archives.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture1"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "«” —Ã«⁄ «·„Õ–Ê›« "
      TabPicture(4)   =   "archives.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Picture9"
      Tab(4).ControlCount=   1
      Begin VB.PictureBox Picture9 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8775
         Left            =   -74880
         ScaleHeight     =   8775
         ScaleWidth      =   14295
         TabIndex        =   86
         Top             =   360
         Width           =   14295
         Begin TabDlg.SSTab SSTab3 
            Height          =   8535
            Left            =   120
            TabIndex        =   87
            Top             =   120
            Width           =   14055
            _ExtentX        =   24791
            _ExtentY        =   15055
            _Version        =   393216
            Tab             =   2
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
            TabCaption(0)   =   "«·”‰Ê«  «·œ—«”Ì…"
            TabPicture(0)   =   "archives.frx":008C
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Picture12"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "«·√”« –…"
            TabPicture(1)   =   "archives.frx":00A8
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Picture11"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "«· ·«„Ì–"
            TabPicture(2)   =   "archives.frx":00C4
            Tab(2).ControlEnabled=   -1  'True
            Tab(2).Control(0)=   "Picture10"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).ControlCount=   1
            Begin VB.PictureBox Picture12 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   8055
               Left            =   -74880
               ScaleHeight     =   8055
               ScaleWidth      =   13815
               TabIndex        =   121
               Top             =   360
               Width           =   13815
               Begin VB.ComboBox Combo4 
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "Times New Roman"
                     Size            =   15.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   480
                  ItemData        =   "archives.frx":00E0
                  Left            =   600
                  List            =   "archives.frx":00F6
                  Style           =   2  'Dropdown List
                  TabIndex        =   129
                  Top             =   3360
                  Width           =   2055
               End
               Begin VB.CommandButton Command16 
                  Caption         =   "‰“⁄ «·Õ–› ⁄‰ Â–Â «·”‰…"
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
                  Left            =   2880
                  TabIndex        =   123
                  Top             =   4560
                  Width           =   3975
               End
               Begin VB.CommandButton Command15 
                  Caption         =   "⁄—÷ «·”‰Ê«  «·œ—«”Ì… «·„Õ–Ê›…"
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
                  Left            =   9000
                  TabIndex        =   122
                  Top             =   120
                  Width           =   3855
               End
               Begin MSFlexGridLib.MSFlexGrid grd11 
                  Height          =   7335
                  Left            =   7080
                  TabIndex        =   124
                  Top             =   600
                  Width           =   6615
                  _ExtentX        =   11668
                  _ExtentY        =   12938
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
               Begin MSComctlLib.ProgressBar ProgressBar3 
                  Height          =   375
                  Left            =   2880
                  TabIndex        =   125
                  Top             =   4080
                  Width           =   3975
                  _ExtentX        =   7011
                  _ExtentY        =   661
                  _Version        =   393216
                  Appearance      =   1
               End
               Begin VB.Label Label91 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·«”„ «·ÃœÌœ ··”‰… «·œ—«”Ì… «·„Õ–Ê›…"
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
                  Height          =   735
                  Index           =   5
                  Left            =   2760
                  TabIndex        =   130
                  Top             =   3480
                  Width           =   4095
               End
               Begin VB.Label Label24 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  BeginProperty Font 
                     Name            =   "Times New Roman"
                     Size            =   24
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   495
                  Left            =   480
                  TabIndex        =   127
                  Top             =   2760
                  Width           =   2295
               End
               Begin VB.Label Label91 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·”‰… «·œ—«”Ì… «·„Õ–Ê›…"
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
                  Height          =   735
                  Index           =   4
                  Left            =   3720
                  TabIndex        =   126
                  Top             =   2880
                  Width           =   3255
               End
               Begin VB.Shape Shape3 
                  BorderColor     =   &H00FFFFFF&
                  BorderWidth     =   3
                  Height          =   2415
                  Left            =   120
                  Shape           =   4  'Rounded Rectangle
                  Top             =   2640
                  Width           =   6855
               End
            End
            Begin VB.PictureBox Picture11 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   8055
               Left            =   -74880
               ScaleHeight     =   8055
               ScaleWidth      =   13815
               TabIndex        =   101
               Top             =   360
               Width           =   13815
               Begin VB.CommandButton Command14 
                  Caption         =   "⁄—÷ «·√”« –… «·„Õ–Ê›Ì‰"
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
                  Left            =   10080
                  TabIndex        =   103
                  Top             =   120
                  Width           =   2535
               End
               Begin VB.CommandButton Command13 
                  Caption         =   "‰“⁄ «·Õ–› ⁄‰ Â–« «·√” «–"
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
                  Left            =   2280
                  TabIndex        =   102
                  Top             =   4560
                  Width           =   4575
               End
               Begin MSFlexGridLib.MSFlexGrid grd10 
                  Height          =   7335
                  Left            =   7080
                  TabIndex        =   104
                  Top             =   600
                  Width           =   6615
                  _ExtentX        =   11668
                  _ExtentY        =   12938
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
               Begin MSComctlLib.ProgressBar ProgressBar2 
                  Height          =   375
                  Left            =   2280
                  TabIndex        =   105
                  Top             =   4080
                  Width           =   4575
                  _ExtentX        =   8070
                  _ExtentY        =   661
                  _Version        =   393216
                  Appearance      =   1
               End
               Begin VB.Image Image1 
                  Appearance      =   0  'Flat
                  Height          =   1695
                  Left            =   240
                  Stretch         =   -1  'True
                  Top             =   3240
                  Width           =   1695
               End
               Begin VB.Shape Shape2 
                  BorderColor     =   &H00FFFFFF&
                  BorderWidth     =   3
                  Height          =   1935
                  Left            =   120
                  Shape           =   4  'Rounded Rectangle
                  Top             =   3120
                  Width           =   6855
               End
               Begin VB.Label Label21 
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
                  Left            =   2040
                  TabIndex        =   111
                  Top             =   3600
                  Width           =   4215
               End
               Begin VB.Label Label20 
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
                  TabIndex        =   110
                  Top             =   3240
                  Width           =   1575
               End
               Begin VB.Label Label19 
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
                  Left            =   4920
                  TabIndex        =   109
                  Top             =   3600
                  Width           =   1935
               End
               Begin VB.Label Label91 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Â« ›"
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
                  Left            =   5760
                  TabIndex        =   108
                  Top             =   3240
                  Width           =   1095
               End
               Begin VB.Label Label91 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«· ”·”·Ì"
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
                  Left            =   3480
                  TabIndex        =   107
                  Top             =   3240
                  Width           =   1095
               End
               Begin VB.Label Label18 
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
                  Left            =   2040
                  TabIndex        =   106
                  Top             =   3240
                  Width           =   1575
               End
            End
            Begin VB.PictureBox Picture10 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   8055
               Left            =   120
               ScaleHeight     =   8055
               ScaleWidth      =   13815
               TabIndex        =   88
               Top             =   360
               Width           =   13815
               Begin VB.CommandButton Command28 
                  Caption         =   "‰“⁄ «·Õ–› ⁄‰ Â–« «· ·„Ì–"
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
                  Left            =   2280
                  TabIndex        =   91
                  Top             =   4560
                  Width           =   4575
               End
               Begin VB.CommandButton Command12 
                  Caption         =   "⁄—÷ «· ·«„Ì– «·„Õ–Ê›Ì‰"
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
                  Left            =   10080
                  TabIndex        =   89
                  Top             =   120
                  Width           =   2535
               End
               Begin MSFlexGridLib.MSFlexGrid grd9 
                  Height          =   7335
                  Left            =   7080
                  TabIndex        =   90
                  Top             =   600
                  Width           =   6615
                  _ExtentX        =   11668
                  _ExtentY        =   12938
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
               Begin MSComctlLib.ProgressBar ProgressBar1 
                  Height          =   375
                  Left            =   2280
                  TabIndex        =   100
                  Top             =   4080
                  Width           =   4575
                  _ExtentX        =   8070
                  _ExtentY        =   661
                  _Version        =   393216
                  Appearance      =   1
               End
               Begin VB.Label Label15 
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
                  Left            =   2040
                  TabIndex        =   97
                  Top             =   3240
                  Width           =   1575
               End
               Begin VB.Label Label91 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«· ”·”·Ì"
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
                  Left            =   3480
                  TabIndex        =   96
                  Top             =   3240
                  Width           =   1095
               End
               Begin VB.Label Label91 
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
                  Index           =   0
                  Left            =   5760
                  TabIndex        =   95
                  Top             =   3240
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
                  Left            =   4920
                  TabIndex        =   94
                  Top             =   3600
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
                  Left            =   4680
                  TabIndex        =   93
                  Top             =   3240
                  Width           =   1575
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
                  Left            =   2040
                  TabIndex        =   92
                  Top             =   3600
                  Width           =   4215
               End
               Begin VB.Shape Shape8 
                  BorderColor     =   &H00FFFFFF&
                  BorderWidth     =   3
                  Height          =   1935
                  Left            =   120
                  Shape           =   4  'Rounded Rectangle
                  Top             =   3120
                  Width           =   6855
               End
               Begin VB.Image Image3 
                  Appearance      =   0  'Flat
                  Height          =   1695
                  Left            =   240
                  Stretch         =   -1  'True
                  Top             =   3240
                  Width           =   1695
               End
            End
         End
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8775
         Left            =   120
         ScaleHeight     =   8775
         ScaleWidth      =   14295
         TabIndex        =   57
         Top             =   360
         Width           =   14295
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            Caption         =   "«·— »…"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   2880
            MaskColor       =   &H00000000&
            TabIndex        =   133
            Top             =   8040
            Width           =   255
         End
         Begin VB.CheckBox Check2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            Caption         =   "«·— »…"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   2880
            MaskColor       =   &H00000000&
            TabIndex        =   112
            Top             =   2280
            Width           =   255
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   600
            Top             =   3600
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton Command10 
            Caption         =   "« „«„ «·⁄„·Ì…"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   27.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   480
            TabIndex        =   78
            Top             =   6480
            Width           =   2655
         End
         Begin VB.DriveListBox Drive1 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   4080
            TabIndex        =   73
            Top             =   6840
            Width           =   3135
         End
         Begin VB.CommandButton Command9 
            Caption         =   "« „«„ «·⁄„·Ì…"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   27.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   480
            TabIndex        =   61
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label Label68 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÊ«⁄œ «·»Ì«‰«  «·„Õ–Ê›…"
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
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   132
            Top             =   8040
            Width           =   2535
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   4
            X1              =   3240
            X2              =   3480
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "-3"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   29
            Left            =   13440
            TabIndex        =   117
            Top             =   2640
            Width           =   495
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "txt"
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
            Height          =   495
            Index           =   0
            Left            =   8880
            TabIndex        =   116
            Top             =   2760
            Width           =   735
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "mdb"
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
            Height          =   495
            Index           =   0
            Left            =   9600
            TabIndex        =   115
            Top             =   2760
            Width           =   855
         End
         Begin VB.Label Label68 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   $"archives.frx":0140
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
            Left            =   120
            TabIndex        =   114
            Top             =   2760
            Width           =   13455
         End
         Begin VB.Label Label68 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÊ«⁄œ «·»Ì«‰«  «·„Õ–Ê›…"
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
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   113
            Top             =   2280
            Width           =   2535
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   4
            X1              =   3360
            X2              =   3360
            Y1              =   360
            Y2              =   2640
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "-1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   28
            Left            =   13440
            TabIndex        =   81
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "-2"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   27
            Left            =   13440
            TabIndex        =   80
            Top             =   2160
            Width           =   495
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "≈–« ﬂ‰  ﬁœ ﬁ„  ”«»ﬁ« »Õ›Ÿ ‰”Œ… «Õ Ì«ÿÌ… Ì ÊÃ» Õ–›  ·ﬂ «·‰”Œ… „‰ «·„ﬂ«‰ «·–Ì  —Ìœ √‰  Õ›Ÿ ›ÌÂ Â–Â «·‰”Œ… "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   855
            Index           =   26
            Left            =   3360
            TabIndex        =   79
            Top             =   1320
            Width           =   10215
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁ„ »«·÷€ÿ ⁄·Ï “— « „«„ «·⁄„·Ì… «·„ﬁ«»· ·« „«„ ⁄„·Ì… ≈⁄«œ… «·‰”Œ… «·«Õ Ì«ÿÌ… ·ﬁÊ«⁄œ »Ì«‰«  «·»—‰«„Ã"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   3240
            TabIndex        =   77
            Top             =   7440
            Width           =   9255
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "-4"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   25
            Left            =   12480
            TabIndex        =   76
            Top             =   7440
            Width           =   495
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "-3"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   24
            Left            =   12480
            TabIndex        =   75
            Top             =   6840
            Width           =   495
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÌÃ»  ÕœÌœ „ﬂ«‰ «·‰”Œ… «·«Õ Ì«ÿÌ… „‰ Â‰«"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   6240
            TabIndex        =   74
            Top             =   6840
            Width           =   6255
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "F:\AGEP6_0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   3240
            TabIndex        =   72
            Top             =   6315
            Width           =   3615
         End
         Begin VB.Label Label30 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "CD"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   1320
            TabIndex        =   71
            Top             =   5880
            Width           =   735
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "USB"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Index           =   23
            Left            =   2160
            TabIndex        =   70
            Top             =   5880
            Width           =   855
         End
         Begin VB.Label Label32 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "D:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   4080
            TabIndex        =   69
            Top             =   5880
            Width           =   375
         End
         Begin VB.Label Label33 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "C:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   4800
            TabIndex        =   68
            Top             =   5880
            Width           =   375
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "-2"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   22
            Left            =   12480
            TabIndex        =   67
            Top             =   5880
            Width           =   495
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   $"archives.frx":01F0
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   855
            Left            =   240
            TabIndex        =   66
            Top             =   5880
            Width           =   12255
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "-1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   21
            Left            =   12480
            TabIndex        =   65
            Top             =   5280
            Width           =   495
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "AGEP6_0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   5040
            TabIndex        =   64
            Top             =   5280
            Width           =   2295
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÌÃ» «‰  ﬂÊ‰ «·‰”Œ… «·«Õ Ì«ÿÌ…  Õ  «·«”„"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   3840
            TabIndex        =   63
            Top             =   5280
            Width           =   8655
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "·· „ﬂ‰ „‰ «⁄«œ… «·‰”Œ… «·«Õ Ì«ÿÌ… »‘ﬂ· ’ÕÌÕ ÌÃ» „—«⁄«  «·¬ Ì:‹ "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   360
            TabIndex        =   62
            Top             =   4680
            Width           =   13335
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   4
            Height          =   5175
            Index           =   1
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   3360
            Width           =   13935
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   4
            Height          =   2895
            Index           =   0
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   360
            Width           =   13935
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁ„ »«·÷€ÿ ⁄·Ï “— « „«„ «·⁄„·Ì… «·„ﬁ«»· ·Õ›Ÿ ‰”Œ… «Õ Ì«ÿÌ… ·ﬁÊ«⁄œ »Ì«‰«  «·»—‰«„Ã"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   20
            Left            =   1200
            TabIndex        =   60
            Top             =   2160
            Width           =   12255
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "≈⁄«œ… ‰”Œ… «Õ Ì«ÿÌ… ·ﬁÊ«⁄œ »Ì«‰«  «·»—‰«„Ã"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   36
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1095
            Index           =   19
            Left            =   360
            TabIndex        =   59
            Top             =   3480
            Width           =   13335
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ›Ÿ ‰”Œ… «Õ Ì«ÿÌ… ·ﬁÊ«⁄œ »Ì«‰«  «·»—‰«„Ã"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   36
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1095
            Index           =   18
            Left            =   3120
            TabIndex        =   58
            Top             =   480
            Width           =   10695
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8775
         Left            =   -74880
         ScaleHeight     =   8775
         ScaleWidth      =   14295
         TabIndex        =   13
         Top             =   360
         Width           =   14295
         Begin TabDlg.SSTab SSTab2 
            Height          =   8535
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Width           =   14055
            _ExtentX        =   24791
            _ExtentY        =   15055
            _Version        =   393216
            Tab             =   2
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
            TabCaption(0)   =   "≈⁄«œ… √”« –…"
            TabPicture(0)   =   "archives.frx":0286
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Picture7"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "≈⁄«œ…  ·«„Ì–"
            TabPicture(1)   =   "archives.frx":02A2
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Picture5"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "≈⁄«œ… √ﬁ”«„"
            TabPicture(2)   =   "archives.frx":02BE
            Tab(2).ControlEnabled=   -1  'True
            Tab(2).Control(0)=   "Picture4"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).ControlCount=   1
            Begin VB.PictureBox Picture7 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   8055
               Left            =   -74880
               ScaleHeight     =   8055
               ScaleWidth      =   13815
               TabIndex        =   46
               Top             =   360
               Width           =   13815
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
                  ItemData        =   "archives.frx":02DA
                  Left            =   9360
                  List            =   "archives.frx":02EA
                  Style           =   2  'Dropdown List
                  TabIndex        =   49
                  Top             =   120
                  Width           =   2055
               End
               Begin VB.CommandButton Command7 
                  Caption         =   "⁄—÷ √”« –… «·”‰ Ì‰ «·œ—«”Ì Ì‰"
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
                  TabIndex        =   48
                  Top             =   120
                  Width           =   4215
               End
               Begin VB.CommandButton Command6 
                  Caption         =   " √ﬂÌœ ⁄„·Ì… «·«” —Ã«⁄"
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
                  Left            =   4440
                  TabIndex        =   47
                  Top             =   7560
                  Width           =   4935
               End
               Begin MSFlexGridLib.MSFlexGrid grd7 
                  Height          =   6255
                  Left            =   7680
                  TabIndex        =   50
                  Top             =   1200
                  Width           =   6015
                  _ExtentX        =   10610
                  _ExtentY        =   11033
                  _Version        =   393216
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
                  Height          =   6255
                  Left            =   240
                  TabIndex        =   51
                  Top             =   1200
                  Width           =   7335
                  _ExtentX        =   12938
                  _ExtentY        =   11033
                  _Version        =   393216
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
                  Caption         =   "«·”‰… «·œ—«”Ì… «·„⁄«œ „‰Â«"
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
                  Left            =   11160
                  TabIndex        =   56
                  Top             =   120
                  Width           =   2535
               End
               Begin VB.Label Label31 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·”‰… «·œ—«”Ì… «·„⁄«œ ≈·ÌÂ«"
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
                  Left            =   6600
                  TabIndex        =   55
                  Top             =   120
                  Width           =   2535
               End
               Begin VB.Label Label7 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "2010-2011"
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
                  TabIndex        =   54
                  Top             =   120
                  Width           =   2055
               End
               Begin VB.Label Label31 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "√”« –… «·”‰… «·œ—«”Ì… «·„⁄«œ „‰Â«"
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
                  Left            =   10440
                  TabIndex        =   53
                  Top             =   720
                  Width           =   3255
               End
               Begin VB.Label Label31 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "√”« –… «·”‰… «·œ—«”Ì… «·„⁄«œ ≈·ÌÂ«"
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
                  Left            =   4320
                  TabIndex        =   52
                  Top             =   720
                  Width           =   3255
               End
            End
            Begin VB.PictureBox Picture5 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   8055
               Left            =   -74880
               ScaleHeight     =   8055
               ScaleWidth      =   13815
               TabIndex        =   26
               Top             =   360
               Width           =   13815
               Begin VB.PictureBox Picture6 
                  Height          =   3975
                  Left            =   3960
                  ScaleHeight     =   3915
                  ScaleWidth      =   6795
                  TabIndex        =   43
                  Top             =   2760
                  Visible         =   0   'False
                  Width           =   6855
                  Begin VB.Timer Timer3 
                     Enabled         =   0   'False
                     Interval        =   50
                     Left            =   1800
                     Top             =   120
                  End
                  Begin VB.Timer Timer2 
                     Enabled         =   0   'False
                     Interval        =   50
                     Left            =   1320
                     Top             =   120
                  End
                  Begin VB.Timer Timer1 
                     Enabled         =   0   'False
                     Interval        =   50
                     Left            =   840
                     Top             =   120
                  End
                  Begin VB.CommandButton Command11 
                     Caption         =   "Command11"
                     Height          =   735
                     Left            =   360
                     TabIndex        =   83
                     Top             =   2520
                     Width           =   1935
                  End
                  Begin VB.TextBox Text1 
                     Height          =   285
                     Left            =   0
                     TabIndex        =   82
                     Text            =   "Text1"
                     Top             =   960
                     Width           =   5175
                  End
                  Begin VB.Label Label25 
                     Caption         =   "Label25"
                     Height          =   375
                     Left            =   2640
                     TabIndex        =   131
                     Top             =   1800
                     Width           =   2175
                  End
                  Begin VB.Label Label36 
                     Caption         =   "Label36"
                     Height          =   255
                     Left            =   1680
                     TabIndex        =   128
                     Top             =   1560
                     Width           =   1695
                  End
                  Begin VB.Label Label17 
                     Caption         =   "Label17"
                     Height          =   255
                     Left            =   1440
                     TabIndex        =   99
                     Top             =   3480
                     Width           =   1575
                  End
                  Begin VB.Label Label16 
                     Caption         =   "Label16"
                     Height          =   375
                     Left            =   240
                     TabIndex        =   98
                     Top             =   3480
                     Width           =   1935
                  End
                  Begin VB.Label Label11 
                     Caption         =   "Label11"
                     Height          =   255
                     Left            =   240
                     TabIndex        =   85
                     Top             =   1680
                     Width           =   1575
                  End
                  Begin VB.Label Label12 
                     Caption         =   "2011-2012"
                     Height          =   375
                     Left            =   240
                     TabIndex        =   84
                     Top             =   2040
                     Width           =   1575
                  End
                  Begin VB.Label Label5 
                     Caption         =   "Label5"
                     Height          =   255
                     Left            =   0
                     TabIndex        =   45
                     Top             =   240
                     Width           =   975
                  End
                  Begin VB.Label Label6 
                     Caption         =   "Label6"
                     Height          =   255
                     Left            =   0
                     TabIndex        =   44
                     Top             =   600
                     Width           =   1935
                  End
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
                  ItemData        =   "archives.frx":031D
                  Left            =   9360
                  List            =   "archives.frx":032D
                  Style           =   2  'Dropdown List
                  TabIndex        =   29
                  Top             =   120
                  Width           =   2055
               End
               Begin VB.CommandButton Command4 
                  Caption         =   "⁄—÷ √ﬁ”«„ «·”‰ Ì‰ «·œ—«”Ì Ì‰"
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
                  TabIndex        =   28
                  Top             =   120
                  Width           =   4215
               End
               Begin VB.CommandButton Command3 
                  Caption         =   " √ﬂÌœ ⁄„·Ì… «·«” —Ã«⁄"
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
                  TabIndex        =   27
                  Top             =   840
                  Width           =   2535
               End
               Begin MSFlexGridLib.MSFlexGrid grd3 
                  Height          =   6615
                  Left            =   11040
                  TabIndex        =   30
                  Top             =   1320
                  Width           =   2655
                  _ExtentX        =   4683
                  _ExtentY        =   11668
                  _Version        =   393216
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
               Begin MSFlexGridLib.MSFlexGrid grd5 
                  Height          =   6255
                  Left            =   6960
                  TabIndex        =   36
                  Top             =   1320
                  Width           =   3975
                  _ExtentX        =   7011
                  _ExtentY        =   11033
                  _Version        =   393216
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
                  Height          =   6615
                  Left            =   120
                  TabIndex        =   39
                  Top             =   1320
                  Width           =   2655
                  _ExtentX        =   4683
                  _ExtentY        =   11668
                  _Version        =   393216
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
               Begin MSFlexGridLib.MSFlexGrid grd6 
                  Height          =   6255
                  Left            =   2880
                  TabIndex        =   40
                  Top             =   1320
                  Width           =   3975
                  _ExtentX        =   7011
                  _ExtentY        =   11033
                  _Version        =   393216
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
               Begin VB.Label Label4 
                  Alignment       =   2  'Center
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
                  Left            =   2880
                  TabIndex        =   42
                  Top             =   840
                  Width           =   1575
               End
               Begin VB.Label Label31 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   " ·«„Ì– «·ﬁ”„"
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
                  Left            =   3720
                  TabIndex        =   41
                  Top             =   840
                  Width           =   1815
               End
               Begin VB.Label Label31 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   " ·«„Ì– «·ﬁ”„"
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
                  TabIndex        =   38
                  Top             =   840
                  Width           =   1935
               End
               Begin VB.Label Label3 
                  Alignment       =   2  'Center
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
                  Left            =   8280
                  TabIndex        =   37
                  Top             =   840
                  Width           =   1575
               End
               Begin VB.Label Label31 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·”‰… «·œ—«”Ì… «·„⁄«œ „‰Â«"
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
                  Left            =   11160
                  TabIndex        =   35
                  Top             =   120
                  Width           =   2535
               End
               Begin VB.Label Label31 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·”‰… «·œ—«”Ì… «·„⁄«œ ≈·ÌÂ«"
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
                  Left            =   6600
                  TabIndex        =   34
                  Top             =   120
                  Width           =   2535
               End
               Begin VB.Label Label2 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "2010-2011"
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
                  TabIndex        =   33
                  Top             =   120
                  Width           =   2055
               End
               Begin VB.Label Label31 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·”‰… «·œ—«”Ì… «·„⁄«œ „‰Â«"
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
                  Left            =   11040
                  TabIndex        =   32
                  Top             =   840
                  Width           =   2655
               End
               Begin VB.Label Label31 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·”‰… «·œ—«”Ì… «·„⁄«œ ≈·ÌÂ«"
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
                  Left            =   120
                  TabIndex        =   31
                  Top             =   840
                  Width           =   2655
               End
            End
            Begin VB.PictureBox Picture4 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   8055
               Left            =   120
               ScaleHeight     =   8055
               ScaleWidth      =   13815
               TabIndex        =   15
               Top             =   360
               Width           =   13815
               Begin VB.CommandButton Command2 
                  Caption         =   " √ﬂÌœ ⁄„·Ì… «·«” —Ã«⁄"
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
                  Left            =   4440
                  TabIndex        =   25
                  Top             =   7560
                  Width           =   4935
               End
               Begin VB.CommandButton Command8 
                  Caption         =   "⁄—÷ √ﬁ”«„ «·”‰ Ì‰ «·œ—«”Ì Ì‰"
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
                  TabIndex        =   20
                  Top             =   120
                  Width           =   4215
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
                  ItemData        =   "archives.frx":0360
                  Left            =   9360
                  List            =   "archives.frx":0370
                  Style           =   2  'Dropdown List
                  TabIndex        =   17
                  Top             =   120
                  Width           =   2055
               End
               Begin MSFlexGridLib.MSFlexGrid grd1 
                  Height          =   6255
                  Left            =   7680
                  TabIndex        =   22
                  Top             =   1200
                  Width           =   3255
                  _ExtentX        =   5741
                  _ExtentY        =   11033
                  _Version        =   393216
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
               Begin MSFlexGridLib.MSFlexGrid grd2 
                  Height          =   6255
                  Left            =   240
                  TabIndex        =   24
                  Top             =   1200
                  Width           =   5895
                  _ExtentX        =   10398
                  _ExtentY        =   11033
                  _Version        =   393216
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
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "√ﬁ”«„ «·”‰… «·œ—«”Ì… «·„⁄«œ ≈·ÌÂ«"
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
                  Left            =   2880
                  TabIndex        =   23
                  Top             =   720
                  Width           =   3255
               End
               Begin VB.Label Label31 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "√ﬁ”«„ «·”‰… «·œ—«”Ì… «·„⁄«œ „‰Â«"
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
                  Left            =   7680
                  TabIndex        =   21
                  Top             =   720
                  Width           =   3255
               End
               Begin VB.Label Label13 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "2010-2011"
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
                  TabIndex        =   19
                  Top             =   120
                  Width           =   2055
               End
               Begin VB.Label Label31 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·”‰… «·œ—«”Ì… «·„⁄«œ ≈·ÌÂ«"
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
                  Left            =   6600
                  TabIndex        =   18
                  Top             =   120
                  Width           =   2535
               End
               Begin VB.Label Label31 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·”‰… «·œ—«”Ì… «·„⁄«œ „‰Â«"
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
                  Left            =   11160
                  TabIndex        =   16
                  Top             =   120
                  Width           =   2535
               End
            End
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8775
         Left            =   -74880
         ScaleHeight     =   8775
         ScaleWidth      =   14295
         TabIndex        =   7
         Top             =   360
         Width           =   14295
         Begin VB.CommandButton Command17 
            Caption         =   "„”Õ „Õ ÊÌ«  «·”‰… «·Ã«—Ì…"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   24
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   2280
            TabIndex        =   140
            Top             =   8040
            Width           =   6255
         End
         Begin VB.ListBox List2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   36
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8280
            Left            =   10920
            TabIndex        =   9
            Top             =   480
            Width           =   3255
         End
         Begin VB.CommandButton Command1 
            Caption         =   " √ﬂÌœ „”Õ Â–Â «·”‰…"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   36
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   2400
            TabIndex        =   8
            Top             =   3720
            Width           =   6255
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
            Index           =   3
            Left            =   4200
            TabIndex        =   138
            Top             =   5520
            Width           =   2655
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
            Index           =   1
            Left            =   4200
            TabIndex        =   137
            Top             =   6960
            Width           =   2655
         End
         Begin VB.Label Label331 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   $"archives.frx":03A3
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
            Height          =   735
            Index           =   0
            Left            =   240
            TabIndex        =   136
            Top             =   7320
            Width           =   10095
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "mdb"
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
            Height          =   495
            Index           =   1
            Left            =   6720
            TabIndex        =   120
            Top             =   6000
            Width           =   975
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "txt"
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
            Height          =   495
            Index           =   1
            Left            =   6000
            TabIndex        =   119
            Top             =   6000
            Width           =   855
         End
         Begin VB.Label Label68 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   $"archives.frx":043D
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
            Index           =   2
            Left            =   240
            TabIndex        =   118
            Top             =   6000
            Width           =   10575
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "”‰Ê«  «·«—‘Ì›"
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
            Left            =   10920
            TabIndex        =   12
            Top             =   120
            Width           =   3255
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·”‰… «·œ—«”Ì… «· Ì ”Ì „ „”ÕÂ« „‰ «·«—‘Ì›"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   27.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   855
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   10695
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   72
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1575
            Left            =   240
            TabIndex        =   10
            Top             =   1680
            Width           =   10575
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
         Begin VB.CommandButton Command5 
            Caption         =   " √ﬂÌœ ⁄—÷ Â–Â «·”‰…"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   36
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   2400
            TabIndex        =   6
            Top             =   3720
            Width           =   6255
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   36
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8280
            Left            =   10920
            TabIndex        =   2
            Top             =   480
            Width           =   3255
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   27.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   855
            Index           =   30
            Left            =   120
            TabIndex        =   139
            Top             =   5520
            Width           =   10695
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
            Left            =   4200
            TabIndex        =   135
            Top             =   6840
            Width           =   2655
         End
         Begin VB.Label Label331 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   $"archives.frx":05CE
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
            Left            =   3240
            TabIndex        =   134
            Top             =   7320
            Width           =   4455
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   72
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1575
            Left            =   240
            TabIndex        =   5
            Top             =   1680
            Width           =   10575
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·”‰… «·œ—«”Ì… «· Ì ”Ì „ ⁄—÷Â« „‰ «·«—‘Ì›"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   27.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   855
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   10695
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "”‰Ê«  «·«—‘Ì›"
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
            Left            =   10920
            TabIndex        =   3
            Top             =   120
            Width           =   3255
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "archives.frx":0664
      Top             =   120
   End
End
Attribute VB_Name = "archives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const FO_MOVE As Long = &H1
Private Const FO_COPY As Long = &H2
Private Const FO_DELETE As Long = &H3
Private Const FO_RENAME As Long = &H4
Private Const FOF_MULTIDESTFILES As Long = &H1
Private Const FOF_CONFIRMMOUSE As Long = &H2
Private Const FOF_SILENT As Long = &H4
Private Const FOF_RENAMEONCOLLISION As Long = &H8
Private Const FOF_NOCONFIRMATION As Long = &H10
Private Const FOF_WANTMAPPINGHANDLE As Long = &H20
Private Const FOF_CREATEPROGRESSDLG As Long = &H0
Private Const FOF_ALLOWUNDO As Long = &H40
Private Const FOF_FILESONLY As Long = &H80
Private Const FOF_SIMPLEPROGRESS As Long = &H100
Private Const FOF_NOCONFIRMMKDIR As Long = &H200

Private Type SHFILEOPSTRUCT
     hWnd As Long
     wFunc As Long
     pFrom As String
     pTo As String
     fFlags As Long
     fAnyOperationsAborted As Long
     hNameMappings As Long
     lpszProgressTitle As String
End Type

Private Declare Function SHFileOperation Lib "Shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Dim Text As String
Public co2 As ADODB.Connection
Public nn As ADODB.Recordset
Public ns As ADODB.Recordset
Dim PicFilev As String
Dim strStream As ADODB.Stream
Dim fName As String
Function cont2()
Set co2 = New ADODB.Connection
Set nn = New ADODB.Recordset
Set ns = New ADODB.Recordset
co2.Provider = "microsoft.jet.oledb.4.0; jet oledb:database password=7346804"
co2.ConnectionString = App.Path & "\ANNEES.mdb"
co2.Open
nn.Open "select*from Tannees", co2, adOpenKeyset, adLockOptimistic
ns.Open "select*from Tannesssuprimees", co2, adOpenKeyset, adLockOptimistic
End Function
 Private Function LoadPictureFromDB()
On Error Resume Next
Call cont
    Do Until et.EOF
    Set strStream = New ADODB.Stream
    strStream.Type = adTypeBinary
    strStream.Open
  If et!ser = Label6.Caption Then
    strStream.Write et.Fields(4).Value
   strStream.SaveToFile App.Path & "\aboubekrine.bmp", adSaveCreateOverWrite
  '     PicFilev = App.Path & "\aboubekrine.bmp"
 'Image2.Picture = LoadPicture(App.Path & "\aboubekrine.bmp")
'fName = App.Path & "\aboubekrine.bmp"
 'Label11.Caption = "01"

    LoadPictureFromDB = True
  End If
    et.MoveNext
    Loop
 ' If Dir$(App.Path & "\aboubekrine.bmp", vbNormal) <> "" Then Kill (App.Path & "\aboubekrine.bmp")

End Function
Public Function SavePictureToDB(sFileName As String)
On Error Resume Next
    Call cont
    Do Until et.EOF
    If et!ser = Label6.Caption Then
    Set strStream = New ADODB.Stream
        strStream.Type = adTypeBinary
    strStream.Open
    strStream.LoadFromFile sFileName
    et.Fields(4).Value = strStream.Read
    et.Update
        Exit Function
    End If
  et.MoveNext
    Loop
   End Function
Private Function LoadPictureFromDB4()
On Error Resume Next
Image3.Picture = LoadPicture("")
Call cont
    Do Until et.EOF
    Set strStream = New ADODB.Stream
    strStream.Type = adTypeBinary
    strStream.Open
  If et!ser = Label15.Caption Then
    strStream.Write et.Fields(4).Value
    strStream.SaveToFile App.Path & "\aboubekrine.bmp", adSaveCreateOverWrite
    Image3.Picture = LoadPicture(App.Path & "\aboubekrine.bmp")
    '    FileCopy App.Path & "\aboubekrine.bmp", "C:\photos\1.jpg"
    Kill (App.Path & "\aboubekrine.bmp")
    LoadPictureFromDB4 = True
    End If
    et.MoveNext
    Loop
  If Dir$(App.Path & "\aboubekrine.bmp", vbNormal) <> "" Then Kill (App.Path & "\aboubekrine.bmp")

End Function
Function DelTree(ByVal strDir As String) As Long
On Error Resume Next
Dim x As Long
Dim intAttr As Integer
Dim strAllDirs As String
Dim strFile As String
DelTree = -1
On Error Resume Next
strDir = Trim$(strDir)
If Len(strDir) = 0 Then Exit Function
If Right$(strDir, 1) = "\" Then strDir = Left$(strDir, Len(strDir) - 1)
If InStr(strDir, "\") = 0 Then Exit Function
intAttr = GetAttr(strDir)
If (intAttr And vbDirectory) = 0 Then Exit Function
strFile = Dir$(strDir & "\*.*", vbSystem Or vbDirectory Or vbHidden)
Do While Len(strFile)
If strFile <> "." And strFile <> ".." Then
  intAttr = GetAttr(strDir & "\" & strFile)
  If (intAttr And vbDirectory) Then
   strAllDirs = strAllDirs & strFile & Chr$(0)
  Else
   If intAttr <> vbNormal Then
    SetAttr strDir & "\" & strFile, vbNormal
    If Err Then DelTree = Err: Exit Function
   End If
   Kill strDir & "\" & strFile
   If Err Then DelTree = Err: Exit Function
  End If
End If
strFile = Dir$
Loop
Do While Len(strAllDirs)
x = InStr(strAllDirs, Chr$(0))
strFile = Left$(strAllDirs, x - 1)
strAllDirs = Mid$(strAllDirs, x + 1)
x = DelTree(strDir & "\" & strFile)
If x Then DelTree = x: Exit Function
Loop
RmDir strDir
If Err Then
DelTree = Err
Else
DelTree = 0
End If
End Function

Private Sub Combo1_Change()
On Error Resume Next
grd1.Clear
grd1.Rows = 1
grd2.Clear
grd2.Rows = 1
End Sub

Private Sub Combo1_Click()
On Error Resume Next
Combo1_Change
End Sub

Private Sub Combo2_Change()
On Error Resume Next
grd3.Clear
grd3.Rows = 1
grd4.Clear
grd4.Rows = 1
grd5.Clear
grd5.Rows = 1
grd6.Clear
grd6.Rows = 1

End Sub

Private Sub Combo2_Click()
On Error Resume Next
Combo2_Change
End Sub

Private Sub Combo3_Change()
On Error Resume Next
grd7.Clear
grd7.Rows = 1
grd8.Clear
grd8.Rows = 1

End Sub

Private Sub Combo3_Click()
On Error Resume Next
Combo3_Change
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim numar As String
If Label1.Caption = "" Then
MsgBox "ﬁ„ »«·÷€ÿ ⁄·Ï «·”‰… «·„—«œ Õ–›Â«", vbCritical
Exit Sub
End If
g = MsgBox("Â·  —Ìœ «·Õ–› Õﬁ« ", vbInformation + vbYesNo, "AGEP")
If g = vbYes Then
If Label1.Caption = face.SBB1.Panels(9).Text Then
MsgBox "Â–Â «·”‰… ·«Ì„ﬂ‰ Õ–›Â« ·√‰Â« ﬁÌœ «· ’›Õ", vbInformation + arabic
Exit Sub
End If
Call cont2
Do While Not nn.EOF
If nn!act = "1" And Label1.Caption = nn!ann Then
MsgBox "Â–Â «·”‰… ·«Ì„ﬂ‰ Õ–›Â« ·√‰Â« ﬁÌœ «·«” ⁄„«·", vbInformation + arabic
Label1.Caption = ""
Exit Sub
End If
If Label1.Caption = nn!ann Then
nn!sup = "1"
nn.Update
End If
nn.MoveNext
Loop
ns.AddNew
numar = ns!num
ns!ann = Label1.Caption
ns!dat = Date
numar = Label1.Caption + "_" + numar
ns!nom = numar
ns!heu = Time$
ns.Update
Label11.Caption = face.SBB1.Panels(9).Text
face.SBB1.Panels(9).Text = Label1.Caption
start.Label1.Caption = Label1.Caption
Call cont
co.Close
face.SBB1.Panels(9).Text = Label11.Caption
start.Label1.Caption = Label11.Caption
FileCopy App.Path & "\R" & Label1.Caption & ".mdb", App.Path & "\" & numar & ".txt"
Kill App.Path & "\R" & Label1.Caption & ".mdb"
MsgBox " „  «“«·… «·”‰… «·œ—«”Ì… " + Label1.Caption + " „‰ «·«—‘Ì› ", vbInformation + arabic
Call chargelists
Label1.Caption = ""
End If

End Sub

Private Sub Command10_Click()
On Error GoTo p
Text1.Text = Drive1.Drive & face.Caption
Text = Text1.Text
Call bases
'Call Coder
Exit Sub
p:
MsgBox "Êﬁ⁄ Œÿ√ «À‰«¡ «⁄«œ… «·‰”Œ «·«Õ Ì«ÿÌ —»„« ÌﬂÊ‰ «·„”«— Œÿ√, «·—Ã«¡ «⁄«œ… «·„Õ«Ê·…", vbExclamation

End Sub

Private Sub Command11_Click()
On Error Resume Next
FileCopy App.Path & "\" & Label12.Caption & ".txt", App.Path & "\" & Label12.Caption & ".mdb"

End Sub

Private Sub Command12_Click()
On Error Resume Next
If face.SBB1.Panels(10).Text = "«—‘Ì›" Then
MsgBox "·« Ì„ﬂ‰ .. ·√‰ «·”‰… «·œ—«”Ì… «· Ì   ’›ÕÂ« «·¬‰ ÂÌ „‰ «·«—‘Ì›", vbCritical
Exit Sub
End If
grd9.Visible = False
Label95.Caption = ""
Label96.Caption = ""
Label15.Caption = ""
Label16.Caption = ""
Label17.Caption = ""
Image3.Picture = LoadPicture("")
Call chargegrd9
grd9.Visible = True
End Sub

Private Sub Command13_Click()
On Error Resume Next
If Label18.Caption = "" Then
MsgBox "ﬁ„ »«·÷€ÿ ⁄·Ï «”„ «·√” «– «·„—«œ ‰“⁄ «·Õ–› ⁄‰Â", vbCritical
Exit Sub
End If
g = MsgBox("Â·  —Ìœ Õﬁ« ‰“⁄ «·Õ–› ⁄‰ Â–« «·√” «– ø", vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
Call cont
Do While Not pr.EOF
If Label18.Caption = pr!ser Then
pr!act = "1"
pr.Update
ProgressBar2.Value = 0
ProgressBar2.Visible = True
Timer2.Enabled = True
Exit Sub
End If
pr.MoveNext
Loop
End If

End Sub

Private Sub Command14_Click()
On Error Resume Next
If face.SBB1.Panels(10).Text = "«—‘Ì›" Then
MsgBox "·« Ì„ﬂ‰ .. ·√‰ «·”‰… «·œ—«”Ì… «· Ì   ’›ÕÂ« «·¬‰ ÂÌ „‰ «·«—‘Ì›", vbCritical
Exit Sub
End If
grd10.Visible = False
Label18.Caption = ""
Label20.Caption = ""
Label21.Caption = ""
Image1.Picture = LoadPicture("")
Call chargegrd10
grd10.Visible = True

End Sub

Private Sub Command15_Click()
On Error Resume Next
If face.SBB1.Panels(10).Text = "«—‘Ì›" Then
MsgBox "·« Ì„ﬂ‰ .. ·√‰ «·”‰… «·œ—«”Ì… «· Ì   ’›ÕÂ« «·¬‰ ÂÌ „‰ «·«—‘Ì›", vbCritical
Exit Sub
End If
grd11.Visible = False
Label24.Caption = ""
Call chargegrd11
grd11.Visible = True

End Sub

Private Sub Command16_Click()
On Error Resume Next
Dim k As Integer
Dim x$
If Label24.Caption = "" Then
MsgBox "ﬁ„ »«·÷€ÿ ⁄·Ï «·”‰… «·œ—«”Ì…«·„—«œ ‰“⁄ «·Õ–› ⁄‰Â«", vbCritical
Exit Sub
End If
If Combo4.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·«”„ «·ÃœÌœ ··”‰… «·œ—«”Ì… «·„—«œ ‰“⁄ «·Õ–› ⁄‰Â«", vbCritical
Exit Sub
End If
g = MsgBox("Â·  —Ìœ Õﬁ« ‰“⁄ «·Õ–› ⁄‰ Â–Â «·”‰… «·œ—«”Ì… ø", vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
Call cont2
Do While Not nn.EOF
If Combo4.Text = nn!ann And nn!sup = "0" Then
MsgBox "€Ì— „„ﬂ‰ .. «·«”„ «·ÃœÌœ «·„œŒ· „ÕÃÊ“ ·”‰… √Œ—Ï Ã«—Ì «·⁄„· »Â«", vbCritical
Exit Sub
End If
nn.MoveNext
Loop
k = 0
x$ = Dir$(App.Path & "\" & Label25.Caption & ".txt")
If x$ <> "" Then
FileCopy App.Path & "\" & Label25.Caption & ".txt", App.Path & "\R" & Combo4.Text & ".mdb"
Kill App.Path & "\" & Label25.Caption & ".txt"
Else
MsgBox "€Ì— „„ﬂ‰ ... «·”‰… «·„—«œ ‰“⁄ «·Õ–› ⁄‰Â« ·« ÌÊÃœ „·›Â« ›Ì «·«—‘Ì›", vbCritical
Exit Sub
End If
Call cont2
Do While Not nn.EOF
If Combo4.Text = nn!ann And nn!sup = "1" Then
k = 1
nn!sup = "0"
nn.Update
nn.MoveLast
End If
nn.MoveNext
Loop
End If
If k = 0 Then
nn.AddNew
nn!ann = Combo4.Text
nn!act = "0"
nn!sup = "0"
nn.Update
End If
Call cont2
Do While Not ns.EOF
If Label25.Caption = ns!nom Then
ns.Delete
ProgressBar3.Value = 0
ProgressBar3.Visible = True
Timer3.Enabled = True
Exit Sub
End If
ns.MoveNext
Loop
End Sub

Private Sub Command17_Click()
Dim xd As String
Dim xm As String
Dim xy As String
Dim xh As String
Dim xt As String
Dim xs As String
Dim xdy As String
g = MsgBox("Â·  —Ìœ Õ–› „Õ ÊÌ«  Â–Â «·”‰… Õﬁ« ", vbInformation + vbYesNo, "AGEP")
If g = vbYes Then
xd = Day(Date)
xm = Month(Date)
xy = Year(Date)
xh = Hour(Time$)
xt = Minute(Time$)
xs = Second(Time$)
xdy = xd + "-" + xm + "-" + xy + "-" + xh + "-" + xt + "-" + xs
Call cont
co.Close
FileCopy App.Path & "\" & face.SBB1.Panels(9).Text & ".mdb", App.Path & "\" & xdy & ".mdb"
FileCopy App.Path & "\C" & face.SBB1.Panels(9).Text & ".mdb", App.Path & "\C" & xdy & ".mdb"
FileCopy App.Path & "\AAI.mdb", App.Path & "\" & face.SBB1.Panels(9).Text & ".mdb"
FileCopy App.Path & "\AAC.mdb", App.Path & "\C" & face.SBB1.Panels(9).Text & ".mdb"
MsgBox " „ „”Õ „Õ ÊÌ«  Â–Â «·”‰… Ê „  Œ“Ì‰Â«  Õ  «·«”„ " + xd + "-" + xm + "-" + xy + "-" + xh + "-" + xt + "-" + xs
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim a As Double
Dim k As Double
Dim cla1 As String
Dim tx1 As String
If Combo1.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·”‰… «·œ—«”Ì… «·„⁄«œ „‰Â«", vbCritical
Exit Sub
End If
n = grd2.Rows
If n < 2 Then
MsgBox "€Ì— „„ﬂ‰ ... ·« ÊÃœ √ﬁ”«„", vbCritical
Exit Sub
End If
k = 0
g = MsgBox("Â·  —Ìœ Õﬁ« «” —Ã«⁄ Â–Â «·√ﬁ”«„", vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
Call cont
For i = 1 To n - 1
grd2.row = i
grd2.Col = 0
cla1 = grd2.Text
grd2.Col = 1
tx1 = grd2.Text
If tx1 = "ﬁÌœ «·≈÷«›…" Then
k = k + 1
cl.AddNew
a = cl!aut
cl!cla = cla1
cl!num = a
cl!act = "1"
cl.Update
grd2.row = i
grd2.Col = 1
grd2.Text = "„ÊÃÊœ ”·›«"
End If
Next i
If k > 0 Then
cla1 = k
MsgBox " „ «” —Ã«⁄ " + cla1 + " √ﬁ”«„", vbInformation
Else
MsgBox "·„ Ì „ «” —Ã«⁄ √Ì ﬁ”„ ·√‰ Ã„Ì⁄ «·√ﬁ”«„ „ÊÃÊœ… ”·›«", vbInformation
End If
End If
End Sub

Private Sub Command28_Click()
On Error Resume Next
If Label15.Caption = "" Then
MsgBox "ﬁ„ »«·÷€ÿ ⁄·Ï «”„ «· ·„Ì– «·„—«œ ‰“⁄ «·Õ–› ⁄‰Â", vbCritical
Exit Sub
End If
g = MsgBox("Â·  —Ìœ Õﬁ« ‰“⁄ «·Õ–› ⁄‰ Â–« «· ·„Ì–ø", vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
Call numsetu2
Call cont
Do While Not et.EOF
If Label17.Caption = et!aut Then
et!num = Label16.Caption
et.Update
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
et.MoveNext
Loop
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim a As Double
Dim k As Double
Dim nom1 As String
Dim ser1 As String
Dim sex1 As String
Dim tel1 As String
Dim adr1 As String
Dim tx1 As String
Dim cla1 As String
If Combo2.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·”‰… «·œ—«”Ì… «·„⁄«œ „‰Â«", vbCritical
Exit Sub
End If
n = grd6.Rows
If n < 2 Then
MsgBox "€Ì— „„ﬂ‰ ... ·«ÌÊÃœ  ·«„Ì–", vbCritical
Exit Sub
End If
k = 0
g = MsgBox("Â·  —Ìœ Õﬁ« «” —Ã«⁄ Âƒ·«¡ «· ·«„Ì–", vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
For i = 1 To n - 1
grd6.row = i
grd6.Col = 1
tx1 = grd6.Text
grd6.Col = 2
nom1 = grd6.Text
grd6.Col = 3
ser1 = grd6.Text
If tx1 = "*" Then
Call cont
Do While Not et.EOF
If ser1 = et!ser Then
cla1 = et!cla
MsgBox "·« Ì„ﬂ‰ ≈÷«›… «· ·„Ì– ’«Õ» «·—ﬁ„ «· ”·”·Ì " + ser1 + " ≈·Ï Â–« «·ﬁ”„ Ê«·”»» √‰Â „ÊÃÊœ ›Ì «·ﬁ”„ " + cla1 + " ·–« ÌÃ» Õ–›Â „‰ Â–Â «·ﬁ«∆„… ·· „ﬂ‰ „‰ „Ê«’·… ⁄„·Ì… «” —Ã«⁄ «· ·«„Ì– „‰ «·«—‘Ì›", vbCritical + arabic
If k > 0 Then
nom1 = k
MsgBox " „ «” —Ã«⁄ " + nom1 + "  ·„Ì– ", vbInformation
End If
Exit Sub
End If
et.MoveNext
Loop
k = k + 1
Call cont
co.Close
start.Label1.Caption = Combo2.Text
face.SBB1.Panels(9).Text = Combo2.Text
Call cont
Do While Not et.EOF
If et!ser = ser1 Then
Label6.Caption = et!ser
sex1 = et!sex
tel1 = et!tel
adr1 = et!adr
Call LoadPictureFromDB
et.MoveLast
End If
et.MoveNext
Loop
Call cont
co.Close
start.Label1.Caption = Label2.Caption
face.SBB1.Panels(9).Text = Label2.Caption
Call cont
et.AddNew
et!cla = Label4.Caption
et!num = Label5.Caption
et!dat = Date
et!nom = nom1
et!sex = sex1
et!pho = "01"
et!tel = tel1
et!adr = adr1
et!ser = ser1
et!act = "1"
et.Update
fName = App.Path & "\aboubekrine.bmp"
Call cont
co.Close
start.Label1.Caption = Label2.Caption
face.SBB1.Panels(9).Text = Label2.Caption
Call cont
Call SavePictureToDB(fName)
grd6.row = i
grd6.Col = 1
grd6.Text = Label5.Caption
Kill App.Path & "\aboubekrine.bmp"
Label5.Caption = Val(Label5.Caption) + 1
End If
Next i
If k > 0 Then
nom1 = k
MsgBox " „ «” —Ã«⁄ " + nom1 + "  ·„Ì– ", vbInformation
Else
MsgBox "·„ Ì „ «” —Ã«⁄ √Ì  ·„Ì– ·√‰ Ã„Ì⁄ «· ·«„Ì– „ÊÃÊœÊ‰ ”·›«", vbInformation
End If
End If

End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim x$
If face.SBB1.Panels(10).Text = "«—‘Ì›" Then
MsgBox "·« Ì„ﬂ‰ .. ·√‰ «·”‰… «·œ—«”Ì… «· Ì   ’›ÕÂ« «·¬‰ ÂÌ „‰ «·«—‘Ì›", vbCritical
Exit Sub
End If
If Combo2.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·”‰… «·œ—«”Ì… «·„⁄«œ „‰Â«", vbCritical
Exit Sub
End If
Command4.Enabled = False
grd3.Clear
grd3.Rows = 1
grd3.Cols = 1
grd3.ColWidth(0) = 2300
grd3.ColAlignment(0) = 0
grd3.row = 0
grd3.Col = 0
grd3.Text = "«·ﬁ”„"
grd4.Clear
grd4.Rows = 1
grd4.Cols = 1
grd4.ColWidth(0) = 2300
grd4.ColAlignment(0) = 0
grd4.row = 0
grd4.Col = 0
grd4.Text = "«·ﬁ”„"
i = 1
j = 1
'**** grd1
x$ = Dir$(App.Path & "\R" & Combo2.Text & ".mdb")
If x$ <> "" Then
FileCopy x$, App.Path & "\" & Combo2.Text & ".mdb"
Else
MsgBox "ﬁ«⁄œ… »Ì«‰«  «·”‰… «·œ—«”Ì… «·„œŒ·… €Ì— „ÊÃÊœ…", vbExclamation
Command4.Enabled = True
Exit Sub
End If
Call cont
co.Close
start.Label1.Caption = Combo2.Text
face.SBB1.Panels(9).Text = Combo2.Text
Call cont
grd3.Rows = cl.RecordCount + 3
Do While Not cl.EOF
grd3.row = i
grd3.Col = 0
grd3.Text = cl!cla
i = i + 1
cl.MoveNext
Loop
grd3.Rows = i
grd3.Col = 0
grd3.Sort = 1
'**** grd2
Call cont
co.Close
start.Label1.Caption = Label2.Caption
face.SBB1.Panels(9).Text = Label2.Caption
Call cont
grd4.Rows = cl.RecordCount + 3
Do While Not cl.EOF
grd4.row = j
grd4.Col = 0
grd4.Text = cl!cla
j = j + 1
cl.MoveNext
Loop
grd4.Rows = j
grd4.Col = 0
grd4.Sort = 1
Command4.Enabled = True
End Sub

Private Sub Command5_Click()
On Error Resume Next
Dim x$
Dim k As Integer
Command5.Enabled = False
If Label14.Caption = "" Then
MsgBox "ﬁ„ »«·÷€ÿ ⁄·Ï «·”‰… «·„—«œ ≈⁄«œ Â«", vbCritical
Command5.Enabled = True
Label31(30).Caption = ""
Exit Sub
End If
If face.SBB1.Panels(9).Text = Label14.Caption Then
MsgBox "√‰  «·¬‰  ⁄„· ⁄·Ï Â–Â «·”‰… «·œ—«”Ì… " + Label14.Caption, vbCritical
Command5.Enabled = True
Label31(30).Caption = ""
Exit Sub
End If
k = 0
Call cont2
Do While Not nn.EOF
If nn!ann = Label14.Caption Then
If nn!act = "1" Then
k = 1
Else
k = 2
End If
nn.MoveLast
End If
nn.MoveNext
Loop
If k = 2 Then
x$ = Dir$(App.Path & "\R" & Label14.Caption & ".mdb")
Else
x$ = Dir$(App.Path & "\" & Label14.Caption & ".mdb")
End If
If x$ <> "" Then
If k = 2 Then
FileCopy x$, App.Path & "\" & Label14.Caption & ".mdb"
face.SBB1.Panels(10).Text = "«—‘Ì›"
Else
face.SBB1.Panels(10).Text = "«·”‰… «·œ—«”Ì…"
End If
Else
MsgBox "ﬁ«⁄œ… »Ì«‰«  «·”‰… «·œ—«”Ì… «·„œŒ·… €Ì— „ÊÃÊœ…", vbExclamation
Exit Sub
End If
start.Label1.Caption = Label14.Caption
face.SBB1.Panels(9).Text = Label14.Caption
MsgBox "  „  ≈⁄«œ… «·”‰… «·œ—«”Ì… " + Label14.Caption + " Ì„ﬂ‰ﬂ «·¬‰ «·⁄„· ⁄·ÌÂ« ", vbInformation + arabic
Label14.Caption = ""
Command5.Enabled = True
Label31(30).Caption = ""

End Sub

Private Sub Command6_Click()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim a As Double
Dim k As Double
Dim ser1 As String
Dim nom1 As String
Dim tel1 As String
Dim mat1 As String
Dim adr1 As String
Dim tx1 As String
If Combo3.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·”‰… «·œ—«”Ì… «·„⁄«œ „‰Â«", vbCritical
Exit Sub
End If
n = grd8.Rows
If n < 2 Then
MsgBox "€Ì— „„ﬂ‰ ... ·«ÌÊÃœ √”« –…", vbCritical
Exit Sub
End If
k = 0
g = MsgBox("Â·  —Ìœ Õﬁ« «” —Ã«⁄ Âƒ·«¡ «·√”« –…", vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
Call cont
For i = 1 To n - 1
grd8.row = i
grd8.Col = 0
ser1 = grd8.Text
grd8.Col = 1
nom1 = grd8.Text
grd8.Col = 2
tel1 = grd8.Text
grd8.Col = 3
mat1 = grd8.Text
grd8.Col = 4
adr1 = grd8.Text
grd8.Col = 5
tx1 = grd8.Text
If tx1 = "ﬁÌœ «·≈÷«›…" Then
k = k + 1
pr.AddNew
pr!dat = Date
pr!tel = tel1
pr!nom = nom1
pr!mat = mat1
pr!adr = adr1
pr!ser = ser1
pr!act = "1"
pr.Update
grd8.row = i
grd8.Col = 5
grd8.Text = "„ÊÃÊœ ”·›«"
End If
Next i
If k > 0 Then
ser1 = k
MsgBox " „ «” —Ã«⁄ " + ser1 + " √” «–", vbInformation
Else
MsgBox "·„ Ì „ «” —Ã«⁄ √Ì √” «– ·√‰ Ã„Ì⁄ «·√”« –… „ÊÃÊœÊ‰ ”·›«", vbInformation
End If
End If

End Sub

Private Sub Command7_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim x$
If face.SBB1.Panels(10).Text = "«—‘Ì›" Then
MsgBox "·« Ì„ﬂ‰ .. ·√‰ «·”‰… «·œ—«”Ì… «· Ì   ’›ÕÂ« «·¬‰ ÂÌ „‰ «·«—‘Ì›", vbCritical
Exit Sub
End If
If Combo3.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·”‰… «·œ—«”Ì… «·„⁄«œ „‰Â«", vbCritical
Exit Sub
End If
Command7.Enabled = False
grd7.Clear
grd7.Rows = 1
grd7.Cols = 5
grd7.ColWidth(0) = 1200
grd7.ColWidth(1) = 2700
grd7.ColWidth(2) = 1400
grd7.ColWidth(3) = 0
grd7.ColWidth(4) = 0
grd7.ColAlignment(0) = 1
grd7.ColAlignment(1) = 1
grd7.ColAlignment(2) = 1
grd7.row = 0
grd7.Col = 0
grd7.Text = "«· ”·”·Ì"
grd7.Col = 1
grd7.Text = "«·«”„"
grd7.Col = 2
grd7.Text = "«·Â« ›"
grd8.Clear
grd8.Rows = 1
grd8.Cols = 6
grd8.ColWidth(0) = 1200
grd8.ColWidth(1) = 2700
grd8.ColWidth(2) = 1400
grd8.ColWidth(3) = 0
grd8.ColWidth(4) = 0
grd8.ColWidth(5) = 1500
grd8.ColAlignment(0) = 1
grd8.ColAlignment(1) = 1
grd8.ColAlignment(2) = 1
grd8.ColAlignment(5) = 1
grd8.row = 0
grd8.Col = 0
grd8.Text = "«· ”·”·Ì"
grd8.Col = 1
grd8.Text = "«·«”„"
grd8.Col = 2
grd8.Text = "«·Â« ›"
grd8.Col = 5
grd8.Text = "«·Õ«·…"
i = 1
j = 1
'**** grd1
x$ = Dir$(App.Path & "\R" & Combo3.Text & ".mdb")
If x$ <> "" Then
FileCopy x$, App.Path & "\" & Combo3.Text & ".mdb"
Else
MsgBox "ﬁ«⁄œ… »Ì«‰«  «·”‰… «·œ—«”Ì… «·„œŒ·… €Ì— „ÊÃÊœ…", vbExclamation
Command7.Enabled = True
Exit Sub
End If
Call cont
co.Close
start.Label1.Caption = Combo3.Text
face.SBB1.Panels(9).Text = Combo3.Text
Call cont
grd7.Rows = pr.RecordCount + 3
Do While Not pr.EOF
grd7.row = i
grd7.Col = 0
grd7.Text = pr!ser
grd7.Col = 1
grd7.Text = pr!nom
grd7.Col = 2
grd7.Text = pr!tel
grd7.Col = 3
grd7.Text = pr!mat
grd7.Col = 4
grd7.Text = pr!adr
i = i + 1
pr.MoveNext
Loop
grd7.Rows = i
grd7.Col = 0
grd7.Sort = 1
'**** grd8
Call cont
co.Close
start.Label1.Caption = Label7.Caption
face.SBB1.Panels(9).Text = Label7.Caption
Call cont
grd8.Rows = pr.RecordCount + 3
Do While Not pr.EOF
grd8.row = j
grd8.Col = 0
grd8.Text = pr!ser
grd8.Col = 1
grd8.Text = pr!nom
grd8.Col = 2
grd8.Text = pr!tel
grd8.Col = 3
grd8.Text = pr!mat
grd8.Col = 4
grd8.Text = pr!adr
grd8.Col = 5
grd8.Text = "„ÊÃÊœ ”·›«"
j = j + 1
pr.MoveNext
Loop
grd8.Rows = j
grd8.Col = 0
grd8.Sort = 1
Command7.Enabled = True
End Sub

Private Sub Command8_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim x$
If face.SBB1.Panels(10).Text = "«—‘Ì›" Then
MsgBox "·« Ì„ﬂ‰ .. ·√‰ «·”‰… «·œ—«”Ì… «· Ì   ’›ÕÂ« «·¬‰ ÂÌ „‰ «·«—‘Ì›", vbCritical
Exit Sub
End If
If Combo1.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·”‰… «·œ—«”Ì… «·„⁄«œ „‰Â«", vbCritical
Exit Sub
End If
Command8.Enabled = False
grd1.Clear
grd1.Rows = 1
grd1.Cols = 1
grd1.ColWidth(0) = 2500
grd1.ColAlignment(0) = 0
grd1.row = 0
grd1.Col = 0
grd1.Text = "«·ﬁ”„"
grd2.Clear
grd2.Rows = 1
grd2.Cols = 2
grd2.ColWidth(0) = 2500
grd2.ColWidth(1) = 2500
grd2.ColAlignment(0) = 0
grd2.ColAlignment(1) = 0
grd2.row = 0
grd2.Col = 0
grd2.Text = "«·ﬁ”„"
grd2.Col = 1
grd2.Text = "«·Õ«·…"
i = 1
j = 1
'**** grd1
x$ = Dir$(App.Path & "\R" & Combo1.Text & ".mdb")
If x$ <> "" Then
FileCopy x$, App.Path & "\" & Combo1.Text & ".mdb"
Else
MsgBox "ﬁ«⁄œ… »Ì«‰«  «·”‰… «·œ—«”Ì… «·„œŒ·… €Ì— „ÊÃÊœ…", vbExclamation
Command8.Enabled = True
Exit Sub
End If
Call cont
co.Close
start.Label1.Caption = Combo1.Text
face.SBB1.Panels(9).Text = Combo1.Text
Call cont
grd1.Rows = cl.RecordCount + 3
Do While Not cl.EOF
grd1.row = i
grd1.Col = 0
grd1.Text = cl!cla
i = i + 1
cl.MoveNext
Loop
grd1.Rows = i
grd1.Col = 0
grd1.Sort = 1
'**** grd2
Call cont
co.Close
start.Label1.Caption = Label13.Caption
face.SBB1.Panels(9).Text = Label13.Caption
Call cont
grd2.Rows = cl.RecordCount + 3
Do While Not cl.EOF
grd2.row = j
grd2.Col = 0
grd2.Text = cl!cla
grd2.Col = 1
grd2.Text = "„ÊÃÊœ ”·›«"
j = j + 1
cl.MoveNext
Loop
grd2.Rows = j
grd2.Col = 0
grd2.Sort = 1
Command8.Enabled = True
End Sub

Private Sub Command9_Click()
On Error GoTo p
Dim result As Long, fileop As SHFILEOPSTRUCT
CommonDialog1.CancelError = True
CommonDialog1.DialogTitle = "«Œ — «·„ﬂ«‰ «·„ÿ·Ê» ··Õ›Ÿ"
CommonDialog1.Filter = "|*.*|"
CommonDialog1.FileName = face.Caption
CommonDialog1.ShowSave
Text = CommonDialog1.FileName
'Replace the 'C:\MyDir' below with the name of the directory you want to delete.
'x = DelTree(CommonDialog1.FileName)
'Select Case x
'Case 0: MsgBox "Deleted"
'Case -1: MsgBox "Invalid Directory"
'Case Else: MsgBox "An Error was occured"
'End Select
With fileop
        .hWnd = Me.hWnd
        .wFunc = FO_COPY
        .pFrom = App.Path & "\*.mdb" & vbNullChar & vbNullChar
        .pTo = CommonDialog1.FileName & vbNullChar & vbNullChar
    'If Check2.Value = 1 Then
     '   .pFrom = App.Path & "\*.txt" & vbNullChar & vbNullChar
     '   .pTo = CommonDialog1.FileName & vbNullChar & vbNullChar
   ' End If
        .fFlags = FOF_SIMPLEPROGRESS Or FOF_FILESONLY
End With
result = SHFileOperation(fileop)
If result <> 0 Then
      MsgBox "·ﬁœ ﬁ„  »«·€«¡ ⁄·Ì… «·‰”Œ", vbExclamation
Else
        If fileop.fAnyOperationsAborted <> 0 Then
                    ' MsgBox "Operation Failed"
         End If
End If

If Check2.Value = 1 Then
    
With fileop
        .hWnd = Me.hWnd
        .wFunc = FO_COPY
        '.pFrom = App.Path & "\*.mdb" & vbNullChar & vbNullChar
        '.pTo = CommonDialog1.FileName & vbNullChar & vbNullChar
    'If Check2.Value = 1 Then
        .pFrom = App.Path & "\*.txt" & vbNullChar & vbNullChar
        .pTo = CommonDialog1.FileName & vbNullChar & vbNullChar
    'End If
        .fFlags = FOF_SIMPLEPROGRESS Or FOF_FILESONLY
End With
result = SHFileOperation(fileop)
If result <> 0 Then
      MsgBox "·ﬁœ ﬁ„  »«·€«¡ ⁄·Ì… «·‰”Œ", vbExclamation
Else
        If fileop.fAnyOperationsAborted <> 0 Then
                    ' MsgBox "Operation Failed"
         End If
End If

End If
MsgBox "·ﬁœ  „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
'Call Coder
 Exit Sub
p:
MsgBox "Êﬁ⁄ Œÿ√ «À‰«¡ Õ›Ÿ «·‰”Œ «·«Õ Ì«ÿÌ , «·—Ã«¡ «⁄«œ… «·„Õ«Ê·…", vbExclamation

End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Top = 100
Skin1.LoadSkin App.Path & "\green.skn"
Skin1.ApplySkin Me.hWnd
Label26.Caption = face.Caption
Label8.Caption = "F:\" + face.Caption
Call chargelists
Call chargec4
End Sub
Private Sub chargelists()
On Error Resume Next
List1.Clear
List2.Clear
Combo1.Clear
Combo2.Clear
Combo3.Clear
Call cont2
Do While Not nn.EOF
If nn!sup = "0" Then
If nn!act = "0" Then
List1.AddItem nn!ann
List2.AddItem nn!ann
Combo1.AddItem nn!ann
Combo2.AddItem nn!ann
Combo3.AddItem nn!ann
Else
List1.AddItem nn!ann
Label13.Caption = nn!ann
Label2.Caption = nn!ann
Label7.Caption = nn!ann
End If
End If
nn.MoveNext
Loop
End Sub

Private Sub grd1_Click()
On Error Resume Next
Dim i As Double
Dim cla1 As String
Dim cla2 As String
i = grd1.row
grd1.row = i
grd1.Col = 0
cla1 = grd1.Text
For i = 1 To grd2.Rows - 1
grd2.row = i
grd2.Col = 0
cla2 = grd2.Text
If cla1 = cla2 Then
Exit Sub
End If
Next i
i = grd2.Rows
grd2.Rows = grd2.Rows + 1
grd2.row = i
grd2.Col = 0
grd2.Text = cla1
grd2.Col = 1
grd2.Text = "ﬁÌœ «·≈÷«›…"

End Sub

Private Sub grd10_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim tx As String
i = grd10.row
j = grd10.Col
If grd10.Rows > 1 Then
grd10.row = i
grd10.Col = 0
tx = grd10.Text
grd10.Col = 1
Label20.Caption = grd10.Text
grd10.Col = 2
Label21.Caption = grd10.Text
grd10.Col = 3
Label18.Caption = grd10.Text
End If

End Sub

Private Sub grd11_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim tx As String
i = grd11.row
j = grd11.Col
If grd11.Rows > 1 Then
grd11.row = i
grd11.Col = 0
tx = grd11.Text
grd11.Col = 1
Label24.Caption = grd11.Text
grd11.Col = 2
Label25.Caption = grd11.Text
Label36.Caption = tx
End If

End Sub

Private Sub grd2_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim n As Double
Dim tx1 As String
Dim tx2 As String
If grd2.Rows > 1 Then
j = grd2.row
grd2.row = j
grd2.Col = 1
tx1 = grd2.Text
If tx1 = "„ÊÃÊœ ”·›«" Then
Exit Sub
End If
g = MsgBox("Â·  —Ìœ Õ–› Â–« «·”ÿ—", vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
j = grd2.row
n = grd2.Rows
For i = j To n - 2
grd2.row = i + 1
grd2.Col = 0
tx1 = grd2.Text
grd2.Col = 1
tx2 = grd2.Text
grd2.row = i
grd2.Col = 0
grd2.Text = tx1
grd2.Col = 1
grd2.Text = tx2
Next i
grd2.Rows = i
End If
End If
End Sub

Private Sub grd3_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
If Combo2.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·”‰… «·œ—«”Ì… «·„⁄«œ „‰Â«", vbCritical
Exit Sub
End If
If grd3.Rows > 1 Then
Label3.Caption = grd3.Text
grd5.Clear
grd5.Rows = 1
grd5.Cols = 4
grd5.ColWidth(0) = 0
grd5.ColWidth(1) = 600
grd5.ColWidth(2) = 2000
grd5.ColWidth(3) = 1000
grd5.ColAlignment(1) = 0
grd5.ColAlignment(2) = 0
grd5.ColAlignment(3) = 0
grd5.row = 0
grd5.Col = 1
grd5.Text = "«·—ﬁ„"
grd5.Col = 2
grd5.Text = "«·«”„"
grd5.Col = 3
grd5.Text = "«· ”·”·Ì"
i = 1
'**** grd5
Call cont
co.Close
start.Label1.Caption = Combo2.Text
face.SBB1.Panels(9).Text = Combo2.Text
Call cont
grd5.Rows = et.RecordCount + 3
Do While Not et.EOF
If Label3.Caption = et!cla Then
If Val(et!num) < 1000000 Then
grd5.row = i
grd5.Col = 0
grd5.Text = et!aut
grd5.Col = 1
grd5.Text = et!num
grd5.Col = 2
grd5.Text = et!nom
grd5.Col = 3
grd5.Text = et!ser
i = i + 1
End If
End If
et.MoveNext
Loop
grd5.Rows = i
grd5.Col = 1
grd5.Sort = 1
Call cont
co.Close
start.Label1.Caption = Label2.Caption
face.SBB1.Panels(9).Text = Label2.Caption
End If
End Sub

Private Sub grd4_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
If Combo2.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·”‰… «·œ—«”Ì… «·„⁄«œ „‰Â«", vbCritical
Exit Sub
End If
If grd4.Rows > 1 Then
Label4.Caption = grd4.Text
grd6.Clear
grd6.Rows = 1
grd6.Cols = 4
grd6.ColWidth(0) = 0
grd6.ColWidth(1) = 600
grd6.ColWidth(2) = 2000
grd6.ColWidth(3) = 1000
grd6.ColAlignment(1) = 0
grd6.ColAlignment(2) = 0
grd6.ColAlignment(3) = 0
grd6.row = 0
grd6.Col = 1
grd6.Text = "«·—ﬁ„"
grd6.Col = 2
grd6.Text = "«·«”„"
grd6.Col = 3
grd6.Text = "«· ”·”·Ì"
i = 1
'**** grd6
Call cont
co.Close
start.Label1.Caption = Label2.Caption
face.SBB1.Panels(9).Text = Label2.Caption
Call cont
grd6.Rows = et.RecordCount + 3
Do While Not et.EOF
If Label4.Caption = et!cla Then
If Val(et!num) < 1000000 Then
grd6.row = i
grd6.Col = 0
grd6.Text = et!aut
grd6.Col = 1
grd6.Text = et!num
grd6.Col = 2
grd6.Text = et!nom
grd6.Col = 3
grd6.Text = et!ser
i = i + 1
End If
End If
et.MoveNext
Loop
grd6.Rows = i
grd6.Col = 1
grd6.Sort = 1
Call numsetu
End If

End Sub

Private Sub grd5_Click()
On Error Resume Next
Dim i As Double
Dim ser1 As String
Dim nom1 As String
Dim ser2 As String
i = grd5.row
If i > 0 Then
If Label4.Caption = "" Then
MsgBox "ÌÃ» «·÷€ÿ ⁄·Ï «·ﬁ”„ «·„⁄«œ ≈·ÌÂ", vbCritical
Exit Sub
End If
grd5.row = i
grd5.Col = 2
nom1 = grd5.Text
grd5.Col = 3
ser1 = grd5.Text
For i = 1 To grd6.Rows - 1
grd6.row = i
grd6.Col = 3
ser2 = grd6.Text
If ser1 = ser2 Then
Exit Sub
End If
Next i
i = grd6.Rows
grd6.Rows = grd6.Rows + 1
grd6.row = i
grd6.Col = 1
grd6.Text = "*"
grd6.Col = 2
grd6.Text = nom1
grd6.Col = 3
grd6.Text = ser1
End If
End Sub

Private Sub grd6_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim n As Double
Dim tx0 As String
Dim tx1 As String
Dim tx2 As String
Dim tx3 As String
If grd6.Rows > 1 Then
j = grd6.row
grd6.row = j
grd6.Col = 1
tx1 = grd6.Text
If tx1 <> "*" Then
Exit Sub
End If
g = MsgBox("Â·  —Ìœ Õ–› Â–« «·”ÿ—", vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
j = grd6.row
n = grd6.Rows
For i = j To n - 2
grd6.row = i + 1
grd6.Col = 0
tx0 = grd6.Text
grd6.Col = 1
tx1 = grd6.Text
grd6.Col = 2
tx2 = grd6.Text
grd6.Col = 3
tx3 = grd6.Text
grd6.row = i
grd6.Col = 0
grd6.Text = tx0
grd6.Col = 1
grd6.Text = tx1
grd6.Col = 2
grd6.Text = tx2
grd6.Col = 3
grd6.Text = tx3
Next i
grd6.Rows = i
End If
End If

End Sub

Private Sub grd7_Click()
On Error Resume Next
Dim i As Double
Dim ser1 As String
Dim nom1 As String
Dim tel1 As String
Dim mat1 As String
Dim adr1 As String
Dim ser2 As String
i = grd7.row
grd7.row = i
grd7.Col = 0
ser1 = grd7.Text
grd7.Col = 1
nom1 = grd7.Text
grd7.Col = 2
tel1 = grd7.Text
grd7.Col = 3
mat1 = grd7.Text
grd7.Col = 4
adr1 = grd7.Text
For i = 1 To grd8.Rows - 1
grd8.row = i
grd8.Col = 0
ser2 = grd8.Text
If ser1 = ser2 Then
Exit Sub
End If
Next i
i = grd8.Rows
grd8.Rows = grd8.Rows + 1
grd8.row = i
grd8.Col = 0
grd8.Text = ser1
grd8.Col = 1
grd8.Text = nom1
grd8.Col = 2
grd8.Text = tel1
grd8.Col = 3
grd8.Text = mat1
grd8.Col = 4
grd8.Text = adr1
grd8.Col = 5
grd8.Text = "ﬁÌœ «·≈÷«›…"

End Sub

Private Sub grd9_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim tx As String
i = grd9.row
j = grd9.Col
If grd9.Rows > 1 Then
grd9.row = i
grd9.Col = 0
tx = grd9.Text
Call cont
Do While Not et.EOF
If tx = et!aut Then
Label17.Caption = et!aut
Label95.Caption = et!cla
Label96.Caption = et!nom
Label15.Caption = et!ser
et.MoveLast
End If
et.MoveNext
Loop
Call LoadPictureFromDB4
End If
End Sub

Private Sub List1_Click()
On Error Resume Next
Label14.Caption = List1.Text
End Sub
Private Sub numsetu()
On Error Resume Next
Dim a As Double
Dim b As Double
a = 0
Call cont
Do While Not et.EOF
If Val(et!num) < 1000000 Then
If et!cla = Label4.Caption Then
b = et!num
If b > a Then
a = b
End If
End If
End If
et.MoveNext
Loop
a = a + 1
Label5.Caption = a
End Sub
Private Sub numsetu2()
On Error Resume Next
Dim a As Double
Dim b As Double
a = 0
Call cont
Do While Not et.EOF
If Val(et!num) < 1000000 Then
If et!cla = Label95.Caption Then
b = et!num
If b > a Then
a = b
End If
End If
End If
et.MoveNext
Loop
a = a + 1
Label16.Caption = a
End Sub
Private Sub bases()
On Error Resume Next
Dim result As Long, fileop As SHFILEOPSTRUCT
With fileop
        .hWnd = Me.hWnd
        .wFunc = FO_COPY
        .pFrom = Text1.Text & "\*.mdb" & vbNullChar & vbNullChar
        .pTo = App.Path & vbNullChar & vbNullChar
        .fFlags = FOF_SIMPLEPROGRESS Or FOF_FILESONLY
End With
result = SHFileOperation(fileop)
If Check1.Value = 1 Then
With fileop
        .hWnd = Me.hWnd
        .wFunc = FO_COPY
        .pFrom = Text1.Text & "\*.txt" & vbNullChar & vbNullChar
        .pTo = App.Path & vbNullChar & vbNullChar
        .fFlags = FOF_SIMPLEPROGRESS Or FOF_FILESONLY
End With
result = SHFileOperation(fileop)
End If
If result <> 0 Then
     MsgBox "·ﬁœ ﬁ„  »«·€«¡ ⁄„·Ì… «·‰”Œ", vbExclamation
     Exit Sub
Else
        If fileop.fAnyOperationsAborted <> 0 Then
                    ' MsgBox "Operation Failed"
         End If
End If

MsgBox "·ﬁœ  „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
End Sub

Private Sub List2_Click()
On Error Resume Next
Label1.Caption = List2.Text

End Sub
Private Sub chargegrd9()
On Error Resume Next
Dim i As Double
Dim j As Double
grd9.Clear
grd9.Rows = 1
grd9.Cols = 4
grd9.ColWidth(0) = 0
grd9.ColWidth(1) = 1200
grd9.ColWidth(2) = 3500
grd9.ColWidth(3) = 1500
grd9.ColAlignment(1) = 1
grd9.ColAlignment(2) = 1
grd9.ColAlignment(3) = 1
grd9.row = 0
grd9.Col = 1
grd9.Text = "«·ﬁ”„"
grd9.Col = 2
grd9.Text = "«·«”„"
grd9.Col = 3
grd9.Text = "«·—ﬁ„ «· ”·”·Ì"
i = 1
Call cont
grd9.Rows = et.RecordCount + 30
Do While Not et.EOF
If Val(et!num) >= 1000000 Then
grd9.row = i
grd9.Col = 0
grd9.Text = et!aut
grd9.Col = 1
grd9.Text = et!cla
grd9.Col = 2
grd9.Text = et!nom
grd9.Col = 3
grd9.Text = et!ser
i = i + 1
End If
et.MoveNext
Loop
grd9.Rows = i
grd9.Col = 1
grd9.Sort = 1
End Sub
Private Sub chargegrd10()
On Error Resume Next
Dim i As Double
Dim j As Double
grd10.Clear
grd10.Rows = 1
grd10.Cols = 4
grd10.ColWidth(0) = 0
grd10.ColWidth(1) = 1200
grd10.ColWidth(2) = 3500
grd10.ColWidth(3) = 1500
grd10.ColAlignment(1) = 1
grd10.ColAlignment(2) = 1
grd10.ColAlignment(3) = 1
grd10.row = 0
grd10.Col = 1
grd10.Text = "«·Â« ›"
grd10.Col = 2
grd10.Text = "«·«”„"
grd10.Col = 3
grd10.Text = "«·—ﬁ„ «· ”·”·Ì"
i = 1
Call cont
grd10.Rows = pr.RecordCount + 30
Do While Not pr.EOF
If pr!act = "0" Then
grd10.row = i
grd10.Col = 0
grd10.Text = pr!aut
grd10.Col = 1
grd10.Text = pr!tel
grd10.Col = 2
grd10.Text = pr!nom
grd10.Col = 3
grd10.Text = pr!ser
i = i + 1
End If
pr.MoveNext
Loop
grd10.Rows = i
grd10.Col = 1
grd10.Sort = 1
End Sub
Private Sub chargegrd11()
On Error Resume Next
Dim i As Double
Dim j As Double
grd11.Clear
grd11.Rows = 1
grd11.Cols = 5
grd11.ColWidth(0) = 0
grd11.ColWidth(1) = 1500
grd11.ColWidth(2) = 1800
grd11.ColWidth(3) = 1500
grd11.ColWidth(4) = 1500
grd11.ColAlignment(1) = 1
grd11.ColAlignment(2) = 1
grd11.ColAlignment(3) = 1
grd11.ColAlignment(4) = 1
grd11.row = 0
grd11.Col = 1
grd11.Text = "«·”‰… «·„Õ–Ê›…"
grd11.Col = 2
grd11.Text = "«·”‰… Ê—ﬁ„ Õ–›Â«"
grd11.Col = 3
grd11.Text = " «—ÌŒ «·Õ–›"
grd11.Col = 4
grd11.Text = "”«⁄… «·Õ–›"
i = 1
Call cont2
grd11.Rows = ns.RecordCount + 3
Do While Not ns.EOF
grd11.row = i
grd11.Col = 0
grd11.Text = ns!num
grd11.Col = 1
grd11.Text = ns!ann
grd11.Col = 2
grd11.Text = ns!nom
grd11.Col = 3
grd11.Text = ns!dat
grd11.Col = 4
grd11.Text = ns!heu
i = i + 1
ns.MoveNext
Loop
grd11.Rows = i
grd11.Col = 1
grd11.Sort = 1
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
ProgressBar1.Value = ProgressBar1.Value + 8
If ProgressBar1.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
Command12_Click
ProgressBar1.Value = 0
Timer1.Enabled = False
End If

End Sub

Private Sub Timer2_Timer()
On Error Resume Next
ProgressBar2.Value = ProgressBar2.Value + 8
If ProgressBar2.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
Command14_Click
ProgressBar2.Value = 0
Timer2.Enabled = False
End If

End Sub
Public Sub chargec4()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim k As Double
Combo4.Clear
Call cont
Do While Not an.EOF
Combo4.AddItem an!ann
an.MoveNext
Loop

End Sub

Private Sub Timer3_Timer()
On Error Resume Next
ProgressBar3.Value = ProgressBar3.Value + 8
If ProgressBar3.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
Call chargelists
Call chargec4
Command15_Click
ProgressBar3.Value = 0
Timer3.Enabled = False
End If

End Sub
