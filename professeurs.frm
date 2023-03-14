VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{8E515444-86DF-11D3-A630-444553540001}#1.0#0"; "barcodex.ocx"
Begin VB.Form professeurs 
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   9255
      Left            =   120
      ScaleHeight     =   9255
      ScaleWidth      =   14535
      TabIndex        =   1
      Top             =   120
      Width           =   14535
      Begin VB.PictureBox Picture13 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8295
         Left            =   840
         ScaleHeight     =   8295
         ScaleWidth      =   12975
         TabIndex        =   134
         Top             =   720
         Visible         =   0   'False
         Width           =   12975
         Begin VB.CommandButton Command27 
            Caption         =   "”Õ» «·√”„«¡"
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
            TabIndex        =   140
            Top             =   7800
            Width           =   3735
         End
         Begin MSComctlLib.ProgressBar ProgressBar6 
            Height          =   375
            Left            =   4440
            TabIndex        =   139
            Top             =   7800
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Timer Timer8 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   360
            Top             =   360
         End
         Begin MSFlexGridLib.MSFlexGrid grd6 
            Height          =   7455
            Left            =   120
            TabIndex        =   135
            Top             =   240
            Width           =   12735
            _ExtentX        =   22463
            _ExtentY        =   13150
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
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8655
         Left            =   120
         ScaleHeight     =   8655
         ScaleWidth      =   14415
         TabIndex        =   23
         Top             =   600
         Width           =   14415
         Begin TabDlg.SSTab SSTab2 
            Height          =   8535
            Left            =   120
            TabIndex        =   24
            Top             =   120
            Visible         =   0   'False
            Width           =   14175
            _ExtentX        =   25003
            _ExtentY        =   15055
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
            TabCaption(0)   =   "”Ã· Õ÷Ê— «·√” «–"
            TabPicture(0)   =   "professeurs.frx":0000
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Picture10"
            Tab(0).Control(1)=   "Picture12"
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "»Ì«‰«  «·√” «–"
            TabPicture(1)   =   "professeurs.frx":001C
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "Picture7"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).ControlCount=   1
            Begin VB.PictureBox Picture12 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   615
               Left            =   -74880
               ScaleHeight     =   615
               ScaleWidth      =   13935
               TabIndex        =   77
               Top             =   360
               Width           =   13935
               Begin VB.CommandButton Command13 
                  Caption         =   "⁄—÷ ”Ã· «·Õ÷Ê—"
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
                  TabIndex        =   78
                  Top             =   120
                  Width           =   3375
               End
               Begin VB.Label Label37 
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
                  Left            =   9720
                  TabIndex        =   82
                  Top             =   120
                  Width           =   2055
               End
               Begin VB.Label Label34 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·—ﬁ„ «· ”·”·Ì"
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
                  Left            =   11400
                  TabIndex        =   81
                  Top             =   120
                  Width           =   1695
               End
               Begin VB.Label Label31 
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
                  Left            =   7560
                  TabIndex        =   80
                  Top             =   120
                  Width           =   1935
               End
               Begin VB.Label Label38 
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
                  Left            =   4200
                  TabIndex        =   79
                  Top             =   120
                  Width           =   4695
               End
               Begin VB.Shape Shape5 
                  BorderColor     =   &H00FFFFFF&
                  BorderWidth     =   2
                  Height          =   375
                  Left            =   120
                  Shape           =   4  'Rounded Rectangle
                  Top             =   120
                  Width           =   13695
               End
            End
            Begin VB.PictureBox Picture10 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   7335
               Left            =   -74880
               ScaleHeight     =   7335
               ScaleWidth      =   13935
               TabIndex        =   48
               Top             =   1080
               Width           =   13935
               Begin TabDlg.SSTab SSTab3 
                  Height          =   7095
                  Left            =   120
                  TabIndex        =   49
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   13695
                  _ExtentX        =   24156
                  _ExtentY        =   12515
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
                  TabCaption(0)   =   "Ì ﬁ«÷Ï »«·‰”»…"
                  TabPicture(0)   =   "professeurs.frx":0038
                  Tab(0).ControlEnabled=   0   'False
                  Tab(0).Control(0)=   "Picture5"
                  Tab(0).ControlCount=   1
                  TabCaption(1)   =   "Ì ﬁ«÷Ï »«·‘Â—"
                  TabPicture(1)   =   "professeurs.frx":0054
                  Tab(1).ControlEnabled=   0   'False
                  Tab(1).Control(0)=   "Picture3"
                  Tab(1).ControlCount=   1
                  TabCaption(2)   =   "Ì ﬁ«÷Ï »«·”«⁄…"
                  TabPicture(2)   =   "professeurs.frx":0070
                  Tab(2).ControlEnabled=   -1  'True
                  Tab(2).Control(0)=   "Picture11"
                  Tab(2).Control(0).Enabled=   0   'False
                  Tab(2).ControlCount=   1
                  Begin VB.PictureBox Picture5 
                     BackColor       =   &H00000000&
                     BorderStyle     =   0  'None
                     Height          =   6615
                     Left            =   -74880
                     ScaleHeight     =   6615
                     ScaleWidth      =   13455
                     TabIndex        =   110
                     Top             =   360
                     Width           =   13455
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
                        Left            =   1560
                        MaskColor       =   &H00000000&
                        TabIndex        =   145
                        Top             =   1320
                        Width           =   255
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
                        ItemData        =   "professeurs.frx":008C
                        Left            =   6240
                        List            =   "professeurs.frx":00B4
                        Style           =   2  'Dropdown List
                        TabIndex        =   137
                        Top             =   240
                        Width           =   2895
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
                        Left            =   3840
                        Style           =   2  'Dropdown List
                        TabIndex        =   132
                        Top             =   240
                        Width           =   1455
                     End
                     Begin MSComctlLib.ProgressBar ProgressBar5 
                        Height          =   375
                        Left            =   3840
                        TabIndex        =   131
                        Top             =   720
                        Width           =   3255
                        _ExtentX        =   5741
                        _ExtentY        =   661
                        _Version        =   393216
                        Appearance      =   1
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
                        Left            =   4680
                        TabIndex        =   119
                        Top             =   1320
                        Width           =   1095
                     End
                     Begin VB.CommandButton Command24 
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
                        Left            =   3840
                        TabIndex        =   118
                        Top             =   1320
                        Width           =   735
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
                        Left            =   6840
                        MaskColor       =   &H00000000&
                        TabIndex        =   117
                        Top             =   800
                        Width           =   255
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
                        Left            =   240
                        TabIndex        =   116
                        Top             =   720
                        Width           =   735
                     End
                     Begin VB.CommandButton Command22 
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
                        Left            =   1080
                        TabIndex        =   115
                        Top             =   720
                        Width           =   735
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
                        Left            =   1920
                        TabIndex        =   114
                        Top             =   720
                        Width           =   1815
                     End
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
                        Left            =   7200
                        TabIndex        =   113
                        Text            =   "0"
                        Top             =   720
                        Width           =   1935
                     End
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
                        Left            =   10200
                        TabIndex        =   112
                        Text            =   "0"
                        Top             =   720
                        Width           =   1575
                     End
                     Begin VB.TextBox Text16 
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
                        Left            =   240
                        TabIndex        =   111
                        Top             =   240
                        Width           =   1695
                     End
                     Begin MSComCtl2.DTPicker DT9 
                        Height          =   375
                        Left            =   10200
                        TabIndex        =   120
                        Top             =   240
                        Width           =   1575
                        _ExtentX        =   2778
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
                        Format          =   108396545
                        CurrentDate     =   41154
                     End
                     Begin MSFlexGridLib.MSFlexGrid grd3 
                        Height          =   4575
                        Left            =   120
                        TabIndex        =   121
                        Top             =   1920
                        Width           =   13215
                        _ExtentX        =   23310
                        _ExtentY        =   8070
                        _Version        =   393216
                        FixedCols       =   0
                        BackColor       =   0
                        ForeColor       =   16777215
                        BackColorFixed  =   0
                        ForeColorFixed  =   16777215
                        ForeColorSel    =   8388608
                        BackColorBkg    =   0
                        Enabled         =   -1  'True
                        RightToLeft     =   -1  'True
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
                     Begin MSComCtl2.DTPicker DT10 
                        Height          =   375
                        Left            =   7920
                        TabIndex        =   122
                        Top             =   1320
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
                        Format          =   108396545
                        CurrentDate     =   41154
                     End
                     Begin MSComCtl2.DTPicker DT11 
                        Height          =   375
                        Left            =   5880
                        TabIndex        =   123
                        Top             =   1320
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
                        Format          =   108396545
                        CurrentDate     =   41154
                     End
                     Begin VB.Label Label46 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   " ⁄œÌ· «·»Ì«‰« "
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
                        Left            =   360
                        TabIndex        =   146
                        Top             =   1320
                        Width           =   1095
                     End
                     Begin VB.Label Label45 
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
                        Left            =   8760
                        TabIndex        =   138
                        Top             =   240
                        Width           =   1335
                     End
                     Begin VB.Label Label30 
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
                        Left            =   4800
                        TabIndex        =   133
                        Top             =   240
                        Width           =   1335
                     End
                     Begin VB.Shape Shape3 
                        BorderColor     =   &H00FFFFFF&
                        BorderWidth     =   2
                        Height          =   615
                        Index           =   5
                        Left            =   120
                        Shape           =   4  'Rounded Rectangle
                        Top             =   1200
                        Width           =   13215
                     End
                     Begin VB.Label Label44 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "⁄—÷ ”Ã· «·Õ÷Ê— „‰  «—ÌŒ"
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
                        Left            =   9000
                        TabIndex        =   130
                        Top             =   1320
                        Width           =   3015
                     End
                     Begin VB.Label Label43 
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
                        Left            =   7200
                        TabIndex        =   129
                        Top             =   1320
                        Width           =   855
                     End
                     Begin VB.Label Label42 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "Ã„Ì⁄ «·„»«·€ «·„„«À·… ·‹ "
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
                        Left            =   3840
                        TabIndex        =   128
                        Top             =   795
                        Width           =   3015
                     End
                     Begin VB.Shape Shape3 
                        BorderColor     =   &H00FFFFFF&
                        BorderWidth     =   2
                        Height          =   1095
                        Index           =   4
                        Left            =   120
                        Shape           =   4  'Rounded Rectangle
                        Top             =   120
                        Width           =   13215
                     End
                     Begin VB.Label Label41 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "„ √Œ—« "
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
                        TabIndex        =   127
                        Top             =   720
                        Width           =   1335
                     End
                     Begin VB.Label Label40 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "⁄·«Ê« "
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
                        Left            =   11760
                        TabIndex        =   126
                        Top             =   720
                        Width           =   1335
                     End
                     Begin VB.Label Label39 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "⁄œœ ”«⁄«  «·· œ—Ì”"
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
                        Left            =   840
                        TabIndex        =   125
                        Top             =   240
                        Width           =   2895
                     End
                     Begin VB.Label Label35 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   " «—ÌŒ «· ﬁÌÌœ"
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
                        Left            =   11760
                        TabIndex        =   124
                        Top             =   240
                        Width           =   1335
                     End
                  End
                  Begin VB.PictureBox Picture3 
                     BackColor       =   &H00000000&
                     BorderStyle     =   0  'None
                     Height          =   6615
                     Left            =   -74880
                     ScaleHeight     =   6615
                     ScaleWidth      =   13455
                     TabIndex        =   84
                     Top             =   360
                     Width           =   13455
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
                        Left            =   1560
                        MaskColor       =   &H00000000&
                        TabIndex        =   143
                        Top             =   1320
                        Width           =   255
                     End
                     Begin MSComctlLib.ProgressBar ProgressBar4 
                        Height          =   375
                        Left            =   3840
                        TabIndex        =   109
                        Top             =   720
                        Width           =   3255
                        _ExtentX        =   5741
                        _ExtentY        =   661
                        _Version        =   393216
                        Appearance      =   1
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
                        ItemData        =   "professeurs.frx":010F
                        Left            =   6240
                        List            =   "professeurs.frx":0137
                        Style           =   2  'Dropdown List
                        TabIndex        =   95
                        Top             =   240
                        Width           =   2895
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
                        Left            =   3840
                        Style           =   2  'Dropdown List
                        TabIndex        =   94
                        Top             =   240
                        Width           =   1455
                     End
                     Begin VB.TextBox Text15 
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
                        Left            =   240
                        TabIndex        =   93
                        Top             =   240
                        Width           =   2295
                     End
                     Begin VB.TextBox Text14 
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
                        Left            =   10200
                        TabIndex        =   92
                        Text            =   "0"
                        Top             =   720
                        Width           =   1575
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
                        Left            =   7200
                        TabIndex        =   91
                        Text            =   "0"
                        Top             =   720
                        Width           =   1935
                     End
                     Begin VB.CommandButton Command20 
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
                        Left            =   1920
                        TabIndex        =   90
                        Top             =   720
                        Width           =   1815
                     End
                     Begin VB.CommandButton Command19 
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
                        Left            =   1080
                        TabIndex        =   89
                        Top             =   720
                        Width           =   735
                     End
                     Begin VB.CommandButton Command18 
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
                        Left            =   240
                        TabIndex        =   88
                        Top             =   720
                        Width           =   735
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
                        Left            =   6840
                        MaskColor       =   &H00000000&
                        TabIndex        =   87
                        Top             =   800
                        Width           =   255
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
                        Left            =   3840
                        TabIndex        =   86
                        Top             =   1320
                        Width           =   735
                     End
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
                        Left            =   4680
                        TabIndex        =   85
                        Top             =   1320
                        Width           =   1095
                     End
                     Begin MSComCtl2.DTPicker DT6 
                        Height          =   375
                        Left            =   10200
                        TabIndex        =   96
                        Top             =   240
                        Width           =   1575
                        _ExtentX        =   2778
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
                        Format          =   108396545
                        CurrentDate     =   41154
                     End
                     Begin MSFlexGridLib.MSFlexGrid grd2 
                        Height          =   4575
                        Left            =   120
                        TabIndex        =   97
                        Top             =   1920
                        Width           =   13215
                        _ExtentX        =   23310
                        _ExtentY        =   8070
                        _Version        =   393216
                        FixedCols       =   0
                        BackColor       =   0
                        ForeColor       =   16777215
                        BackColorFixed  =   0
                        ForeColorFixed  =   16777215
                        ForeColorSel    =   8388608
                        BackColorBkg    =   0
                        Enabled         =   -1  'True
                        RightToLeft     =   -1  'True
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
                     Begin MSComCtl2.DTPicker DT7 
                        Height          =   375
                        Left            =   7920
                        TabIndex        =   98
                        Top             =   1320
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
                        Format          =   108396545
                        CurrentDate     =   41154
                     End
                     Begin MSComCtl2.DTPicker DT8 
                        Height          =   375
                        Left            =   5880
                        TabIndex        =   99
                        Top             =   1320
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
                        Format          =   108396545
                        CurrentDate     =   41154
                     End
                     Begin VB.Label Label46 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   " ⁄œÌ· «·»Ì«‰« "
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
                        Left            =   240
                        TabIndex        =   144
                        Top             =   1320
                        Width           =   1215
                     End
                     Begin VB.Label Label28 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "«·„»·€ ··‘Â—"
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
                        TabIndex        =   108
                        Top             =   240
                        Width           =   1455
                     End
                     Begin VB.Label Label33 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   " «—ÌŒ «· ﬁÌÌœ"
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
                        Left            =   11760
                        TabIndex        =   107
                        Top             =   240
                        Width           =   1335
                     End
                     Begin VB.Label Label32 
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
                        Left            =   8760
                        TabIndex        =   106
                        Top             =   240
                        Width           =   1335
                     End
                     Begin VB.Label Label29 
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
                        Left            =   4800
                        TabIndex        =   105
                        Top             =   240
                        Width           =   1335
                     End
                     Begin VB.Label Label27 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "⁄·«Ê« "
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
                        Left            =   11760
                        TabIndex        =   104
                        Top             =   720
                        Width           =   1335
                     End
                     Begin VB.Label Label26 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "„ √Œ—« "
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
                        TabIndex        =   103
                        Top             =   720
                        Width           =   1335
                     End
                     Begin VB.Shape Shape3 
                        BorderColor     =   &H00FFFFFF&
                        BorderWidth     =   2
                        Height          =   1095
                        Index           =   3
                        Left            =   120
                        Shape           =   4  'Rounded Rectangle
                        Top             =   120
                        Width           =   13215
                     End
                     Begin VB.Label Label25 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "Ã„Ì⁄ «·„»«·€ «·„„«À·… ·‹ "
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
                        Left            =   3840
                        TabIndex        =   102
                        Top             =   795
                        Width           =   3015
                     End
                     Begin VB.Label Label24 
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
                        Left            =   7200
                        TabIndex        =   101
                        Top             =   1320
                        Width           =   855
                     End
                     Begin VB.Label Label23 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "⁄—÷ ”Ã· «·Õ÷Ê— „‰  «—ÌŒ"
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
                        Left            =   9000
                        TabIndex        =   100
                        Top             =   1320
                        Width           =   3015
                     End
                     Begin VB.Shape Shape3 
                        BorderColor     =   &H00FFFFFF&
                        BorderWidth     =   2
                        Height          =   615
                        Index           =   2
                        Left            =   120
                        Shape           =   4  'Rounded Rectangle
                        Top             =   1200
                        Width           =   13215
                     End
                  End
                  Begin VB.PictureBox Picture11 
                     BackColor       =   &H00000000&
                     BorderStyle     =   0  'None
                     Height          =   6615
                     Left            =   120
                     ScaleHeight     =   6615
                     ScaleWidth      =   13455
                     TabIndex        =   50
                     Top             =   360
                     Width           =   13455
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
                        Left            =   1560
                        MaskColor       =   &H00000000&
                        TabIndex        =   141
                        Top             =   1320
                        Width           =   255
                     End
                     Begin MSComctlLib.ProgressBar ProgressBar3 
                        Height          =   375
                        Left            =   3840
                        TabIndex        =   83
                        Top             =   720
                        Width           =   3255
                        _ExtentX        =   5741
                        _ExtentY        =   661
                        _Version        =   393216
                        Appearance      =   1
                     End
                     Begin VB.CommandButton Command14 
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
                        Left            =   4680
                        TabIndex        =   62
                        Top             =   1320
                        Width           =   1095
                     End
                     Begin VB.CommandButton Command12 
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
                        Left            =   3840
                        TabIndex        =   61
                        Top             =   1320
                        Width           =   735
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
                        Left            =   6840
                        MaskColor       =   &H00000000&
                        TabIndex        =   60
                        Top             =   800
                        Width           =   255
                     End
                     Begin VB.CommandButton Command11 
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
                        Left            =   240
                        TabIndex        =   59
                        Top             =   720
                        Width           =   735
                     End
                     Begin VB.CommandButton Command10 
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
                        Left            =   1080
                        TabIndex        =   58
                        Top             =   720
                        Width           =   735
                     End
                     Begin VB.CommandButton Command8 
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
                        Left            =   1920
                        TabIndex        =   57
                        Top             =   720
                        Width           =   1815
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
                        Left            =   7200
                        TabIndex        =   56
                        Text            =   "0"
                        Top             =   720
                        Width           =   1935
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
                        Left            =   10200
                        TabIndex        =   55
                        Text            =   "0"
                        Top             =   720
                        Width           =   1575
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
                        Left            =   240
                        TabIndex        =   54
                        Top             =   240
                        Width           =   2295
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
                        Left            =   3840
                        Style           =   2  'Dropdown List
                        TabIndex        =   53
                        Top             =   240
                        Width           =   1455
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
                        Left            =   6240
                        Style           =   2  'Dropdown List
                        TabIndex        =   52
                        Top             =   240
                        Width           =   855
                     End
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
                        Left            =   8280
                        Style           =   2  'Dropdown List
                        TabIndex        =   51
                        Top             =   240
                        Width           =   855
                     End
                     Begin MSComCtl2.DTPicker DT3 
                        Height          =   375
                        Left            =   10200
                        TabIndex        =   63
                        Top             =   240
                        Width           =   1575
                        _ExtentX        =   2778
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
                        Format          =   108396545
                        CurrentDate     =   41154
                     End
                     Begin MSFlexGridLib.MSFlexGrid grd1 
                        Height          =   4575
                        Left            =   120
                        TabIndex        =   64
                        Top             =   1920
                        Width           =   13215
                        _ExtentX        =   23310
                        _ExtentY        =   8070
                        _Version        =   393216
                        FixedCols       =   0
                        BackColor       =   0
                        ForeColor       =   16777215
                        BackColorFixed  =   0
                        ForeColorFixed  =   16777215
                        ForeColorSel    =   8388608
                        BackColorBkg    =   0
                        Enabled         =   -1  'True
                        RightToLeft     =   -1  'True
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
                     Begin MSComCtl2.DTPicker DT4 
                        Height          =   375
                        Left            =   7920
                        TabIndex        =   65
                        Top             =   1320
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
                        Format          =   108396545
                        CurrentDate     =   41154
                     End
                     Begin MSComCtl2.DTPicker DT5 
                        Height          =   375
                        Left            =   5880
                        TabIndex        =   66
                        Top             =   1320
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
                        Format          =   108396545
                        CurrentDate     =   41154
                     End
                     Begin VB.Label Label46 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   " ⁄œÌ· «·»Ì«‰« "
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
                        Left            =   240
                        TabIndex        =   142
                        Top             =   1320
                        Width           =   1215
                     End
                     Begin VB.Shape Shape3 
                        BorderColor     =   &H00FFFFFF&
                        BorderWidth     =   2
                        Height          =   615
                        Index           =   1
                        Left            =   120
                        Shape           =   4  'Rounded Rectangle
                        Top             =   1200
                        Width           =   13215
                     End
                     Begin VB.Label Label19 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "⁄—÷ ”Ã· «·Õ÷Ê— „‰  «—ÌŒ"
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
                        Left            =   9000
                        TabIndex        =   76
                        Top             =   1320
                        Width           =   3015
                     End
                     Begin VB.Label Label18 
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
                        Left            =   7200
                        TabIndex        =   75
                        Top             =   1320
                        Width           =   855
                     End
                     Begin VB.Label Label68 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "Ã„Ì⁄ «·„»«·€ «·„„«À·… ·‹ "
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
                        Left            =   3840
                        TabIndex        =   74
                        Top             =   795
                        Width           =   3015
                     End
                     Begin VB.Shape Shape3 
                        BorderColor     =   &H00FFFFFF&
                        BorderWidth     =   2
                        Height          =   1095
                        Index           =   0
                        Left            =   120
                        Shape           =   4  'Rounded Rectangle
                        Top             =   120
                        Width           =   13215
                     End
                     Begin VB.Label Label17 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "„ √Œ—« "
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
                        TabIndex        =   73
                        Top             =   720
                        Width           =   1335
                     End
                     Begin VB.Label Label16 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "⁄·«Ê« "
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
                        Left            =   11760
                        TabIndex        =   72
                        Top             =   720
                        Width           =   1335
                     End
                     Begin VB.Label Label86 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "«·„»·€ ··”«⁄…"
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
                        TabIndex        =   71
                        Top             =   240
                        Width           =   1335
                     End
                     Begin VB.Label Label85 
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
                        Left            =   4800
                        TabIndex        =   70
                        Top             =   240
                        Width           =   1335
                     End
                     Begin VB.Label Label84 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "≈·Ï «·”«⁄…"
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
                        TabIndex        =   69
                        Top             =   240
                        Width           =   1335
                     End
                     Begin VB.Label Label83 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "„‰ «·”«⁄…"
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
                        TabIndex        =   68
                        Top             =   240
                        Width           =   1335
                     End
                     Begin VB.Label Label15 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   " «—ÌŒ «· œ—Ì”"
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
                        Left            =   11760
                        TabIndex        =   67
                        Top             =   240
                        Width           =   1335
                     End
                  End
               End
            End
            Begin VB.PictureBox Picture7 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   8055
               Left            =   120
               ScaleHeight     =   8055
               ScaleWidth      =   13935
               TabIndex        =   25
               Top             =   360
               Width           =   13935
               Begin VB.PictureBox Picture9 
                  Height          =   5535
                  Left            =   480
                  ScaleHeight     =   5475
                  ScaleWidth      =   2595
                  TabIndex        =   42
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   2655
                  Begin VB.Timer Timer6 
                     Enabled         =   0   'False
                     Interval        =   50
                     Left            =   1320
                     Top             =   1440
                  End
                  Begin VB.Timer Timer5 
                     Enabled         =   0   'False
                     Interval        =   50
                     Left            =   720
                     Top             =   1440
                  End
                  Begin VB.Timer Timer2 
                     Enabled         =   0   'False
                     Interval        =   50
                     Left            =   720
                     Top             =   120
                  End
                  Begin VB.Timer Timer4 
                     Enabled         =   0   'False
                     Interval        =   50
                     Left            =   120
                     Top             =   1440
                  End
                  Begin VB.Timer Timer3 
                     Enabled         =   0   'False
                     Interval        =   50
                     Left            =   1200
                     Top             =   120
                  End
                  Begin VB.Timer Timer1 
                     Enabled         =   0   'False
                     Interval        =   50
                     Left            =   120
                     Top             =   120
                  End
                  Begin VB.CommandButton Command9 
                     Caption         =   "Command9"
                     Height          =   375
                     Left            =   480
                     TabIndex        =   43
                     Top             =   600
                     Width           =   1695
                  End
                  Begin VB.Label Label36 
                     Caption         =   "Label36"
                     Height          =   495
                     Left            =   240
                     TabIndex        =   136
                     Top             =   2880
                     Width           =   2055
                  End
                  Begin VB.Label Label22 
                     Caption         =   "Label22"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   47
                     Top             =   2400
                     Width           =   1335
                  End
                  Begin VB.Label Label21 
                     Caption         =   "Label21"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   46
                     Top             =   2040
                     Width           =   1575
                  End
                  Begin VB.Label Label20 
                     Height          =   375
                     Left            =   120
                     TabIndex        =   45
                     Top             =   1080
                     Width           =   1455
                  End
                  Begin VB.Label Label14 
                     Height          =   255
                     Left            =   600
                     TabIndex        =   44
                     Top             =   240
                     Width           =   1695
                  End
               End
               Begin VB.PictureBox Picture8 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   4695
                  Left            =   3240
                  ScaleHeight     =   4695
                  ScaleWidth      =   7335
                  TabIndex        =   26
                  Top             =   1440
                  Visible         =   0   'False
                  Width           =   7335
                  Begin VB.CommandButton Command7 
                     Caption         =   "Õ–› «·√” «–"
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
                     Left            =   1560
                     TabIndex        =   33
                     Top             =   3240
                     Width           =   1575
                  End
                  Begin VB.CommandButton Command3 
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
                     Left            =   3240
                     TabIndex        =   32
                     Top             =   3240
                     Width           =   3855
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
                     Left            =   240
                     TabIndex        =   31
                     Top             =   720
                     Width           =   5535
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
                     Left            =   240
                     TabIndex        =   30
                     Top             =   240
                     Width           =   2535
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
                     Height          =   825
                     Left            =   3240
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   29
                     Top             =   1800
                     Width           =   2535
                  End
                  Begin VB.CommandButton Command2 
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
                     Left            =   240
                     TabIndex        =   28
                     Top             =   3240
                     Width           =   1215
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
                     Left            =   3000
                     TabIndex        =   27
                     Top             =   1200
                     Width           =   2775
                  End
                  Begin MSComCtl2.DTPicker DT2 
                     Height          =   375
                     Left            =   3960
                     TabIndex        =   34
                     Top             =   240
                     Width           =   1815
                     _ExtentX        =   3201
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
                     Format          =   108396545
                     CurrentDate     =   41154
                  End
                  Begin MSComctlLib.ProgressBar ProgressBar1 
                     Height          =   375
                     Left            =   240
                     TabIndex        =   35
                     Top             =   2760
                     Width           =   5535
                     _ExtentX        =   9763
                     _ExtentY        =   661
                     _Version        =   393216
                     Appearance      =   1
                  End
                  Begin BARCODEXLib.BarcodeX BarcodeX2 
                     BeginProperty DataFormat 
                        Type            =   0
                        Format          =   "0"
                        HaveTrueFalseNull=   0
                        FirstDayOfWeek  =   0
                        FirstWeekOfYear =   0
                        LCID            =   1036
                        SubFormatType   =   0
                     EndProperty
                     Height          =   855
                     Left            =   240
                     Top             =   1800
                     Width           =   2775
                     _Version        =   65536
                     _ExtentX        =   4895
                     _ExtentY        =   1508
                     _StockProps     =   13
                     ForeColor       =   0
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Times New Roman"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Caption         =   "00000000"
                     hasText         =   -1
                     BarcodeType     =   6
                  End
                  Begin VB.Label Label13 
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
                     Left            =   5160
                     TabIndex        =   41
                     Top             =   720
                     Width           =   1935
                  End
                  Begin VB.Label Label12 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   " «—ÌŒ «·«‰ ”«»"
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
                     Left            =   5760
                     TabIndex        =   40
                     Top             =   240
                     Width           =   1335
                  End
                  Begin VB.Label Label11 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "„«œ… «· œ—Ì”"
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
                     Left            =   2640
                     TabIndex        =   39
                     Top             =   240
                     Width           =   1215
                  End
                  Begin VB.Label Label10 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·⁄‰Ê«‰"
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
                     Left            =   5760
                     TabIndex        =   38
                     Top             =   1800
                     Width           =   1335
                  End
                  Begin VB.Label Label9 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·—ﬁ„ «· ”·”·Ì ··√” «–"
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
                     Height          =   615
                     Left            =   240
                     TabIndex        =   37
                     Top             =   1200
                     Width           =   2895
                  End
                  Begin VB.Shape Shape2 
                     BorderColor     =   &H8000000E&
                     BorderWidth     =   2
                     Height          =   3615
                     Left            =   120
                     Top             =   120
                     Width           =   7095
                  End
                  Begin VB.Label Label8 
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
                     Left            =   5880
                     TabIndex        =   36
                     Top             =   1200
                     Width           =   1215
                  End
               End
            End
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   4695
         Left            =   3480
         ScaleHeight     =   4695
         ScaleWidth      =   7335
         TabIndex        =   6
         Top             =   2520
         Visible         =   0   'False
         Width           =   7335
         Begin VB.CommandButton Command31 
            Caption         =   "ÃœÌœ"
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
            TabIndex        =   14
            Top             =   3240
            Width           =   2055
         End
         Begin VB.CommandButton Command6 
            Caption         =   "≈€·«ﬁ"
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
            TabIndex        =   13
            Top             =   3240
            Width           =   735
         End
         Begin VB.CommandButton Command5 
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
            Left            =   4080
            TabIndex        =   12
            Top             =   3240
            Width           =   3015
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
            Left            =   240
            TabIndex        =   11
            Top             =   720
            Width           =   5535
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
            Left            =   240
            TabIndex        =   10
            Top             =   240
            Width           =   2535
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
            Height          =   825
            Left            =   3240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   1800
            Width           =   2535
         End
         Begin VB.CommandButton Command4 
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
            Left            =   1080
            TabIndex        =   8
            Top             =   3240
            Width           =   735
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
            Left            =   3240
            TabIndex        =   7
            Top             =   1200
            Width           =   2535
         End
         Begin MSComCtl2.DTPicker DT1 
            Height          =   375
            Left            =   3960
            TabIndex        =   15
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
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
            Format          =   108396545
            CurrentDate     =   41154
         End
         Begin MSComctlLib.ProgressBar ProgressBar2 
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Top             =   2760
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
         End
         Begin BARCODEXLib.BarcodeX BarcodeX1 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1036
               SubFormatType   =   0
            EndProperty
            Height          =   855
            Left            =   240
            Top             =   1800
            Width           =   2775
            _Version        =   65536
            _ExtentX        =   4895
            _ExtentY        =   1508
            _StockProps     =   13
            ForeColor       =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "00000000"
            hasText         =   -1
            BarcodeType     =   6
         End
         Begin VB.Label Label7 
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
            Left            =   5160
            TabIndex        =   22
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·«‰ ”«»"
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
            Left            =   5760
            TabIndex        =   21
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„«œ… «· œ—Ì”"
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
            Left            =   2640
            TabIndex        =   20
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·⁄‰Ê«‰"
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
            Left            =   5760
            TabIndex        =   19
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·—ﬁ„ «· ”·”·Ì ··√” «–"
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
            Height          =   615
            Left            =   240
            TabIndex        =   18
            Top             =   1200
            Width           =   2895
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H8000000E&
            BorderWidth     =   2
            Height          =   3615
            Left            =   120
            Top             =   120
            Width           =   7095
         End
         Begin VB.Label Label6 
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
            Left            =   5880
            TabIndex        =   17
            Top             =   1200
            Width           =   1215
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   14055
         TabIndex        =   2
         Top             =   0
         Width           =   14055
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
            Left            =   7800
            TabIndex        =   0
            Top             =   120
            Width           =   4575
         End
         Begin VB.CommandButton Command1 
            Caption         =   "≈÷«›… √” «– ÃœÌœ"
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
            TabIndex        =   4
            Top             =   120
            Width           =   2895
         End
         Begin VB.CommandButton Command15 
            Caption         =   "⁄—÷ √”„«¡ Ã„Ì⁄ «·√”« –…"
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
            Width           =   2895
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·√” «–"
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
            Left            =   11760
            TabIndex        =   5
            Top             =   120
            Width           =   1935
         End
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   360
         OleObjectBlob   =   "professeurs.frx":0192
         Top             =   1080
      End
   End
End
Attribute VB_Name = "professeurs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim st1 As Integer
Dim tim As Integer
Public co2 As ADODB.Connection
Public np As ADODB.Recordset
Public ph As ADODB.Recordset
Public pm As ADODB.Recordset
Public pn As ADODB.Recordset
Dim anes As String
Dim data As New Access.Application
Function cont2()
Set co2 = New ADODB.Connection
Set np = New ADODB.Recordset
Set ph = New ADODB.Recordset
Set pm = New ADODB.Recordset
Set pn = New ADODB.Recordset
co2.Provider = "microsoft.jet.oledb.4.0; jet oledb:database password=7346804"
anes = "C" + face.SBB1.Panels(9).Text
co2.ConnectionString = App.Path & "\" & anes & ".mdb"
co2.Open
np.Open "select*from Tprofesseurs", co2, adOpenKeyset, adLockOptimistic
ph.Open "select*from Tpresencesh order by dat ASC", co2, adOpenKeyset, adLockOptimistic
pm.Open "select*from Tpresencesm order by dat ASC", co2, adOpenKeyset, adLockOptimistic
pn.Open "select*from Tpresencesp order by dat ASC", co2, adOpenKeyset, adLockOptimistic
End Function

Private Sub Combo2_Change()
On Error Resume Next
Text16.SetFocus
End Sub

Private Sub Combo2_Click()
On Error Resume Next
Combo2_Change
End Sub

Private Sub Combo3_Change()
On Error Resume Next
If Combo3.Text = "«ﬂ Ê»—" Then
Label36.Caption = "10"
End If
If Combo3.Text = "‰Ê›„»—" Then
Label36.Caption = "11"
End If
If Combo3.Text = "œÌ”„»—" Then
Label36.Caption = "12"
End If
If Combo3.Text = "Ì‰«Ì—" Then
Label36.Caption = "1"
End If
If Combo3.Text = "›»—«Ì—" Then
Label36.Caption = "2"
End If
If Combo3.Text = "„«—”" Then
Label36.Caption = "3"
End If
If Combo3.Text = "«»—Ì·" Then
Label36.Caption = "4"
End If
If Combo3.Text = "„«ÌÊ" Then
Label36.Caption = "5"
End If
If Combo3.Text = "ÌÊ‰ÌÊ" Then
Label36.Caption = "6"
End If
If Combo3.Text = "ÌÊ·ÌÊ" Then
Label36.Caption = "7"
End If
If Combo3.Text = "√€”ÿ”" Then
Label36.Caption = "8"
End If
If Combo3.Text = "”» „»—" Then
Label36.Caption = "9"
End If
End Sub

Private Sub Combo3_Click()
On Error Resume Next
Combo3_Change
End Sub

Private Sub Combo4_Change()
On Error Resume Next
If Combo4.Text = "«ﬂ Ê»—" Then
Label36.Caption = "10"
End If
If Combo4.Text = "‰Ê›„»—" Then
Label36.Caption = "11"
End If
If Combo4.Text = "œÌ”„»—" Then
Label36.Caption = "12"
End If
If Combo4.Text = "Ì‰«Ì—" Then
Label36.Caption = "1"
End If
If Combo4.Text = "›»—«Ì—" Then
Label36.Caption = "2"
End If
If Combo4.Text = "„«—”" Then
Label36.Caption = "3"
End If
If Combo4.Text = "«»—Ì·" Then
Label36.Caption = "4"
End If
If Combo4.Text = "„«ÌÊ" Then
Label36.Caption = "5"
End If
If Combo4.Text = "ÌÊ‰ÌÊ" Then
Label36.Caption = "6"
End If
If Combo4.Text = "ÌÊ·ÌÊ" Then
Label36.Caption = "7"
End If
If Combo4.Text = "√€”ÿ”" Then
Label36.Caption = "8"
End If
If Combo4.Text = "”» „»—" Then
Label36.Caption = "9"
End If

End Sub

Private Sub Combo4_Click()
On Error Resume Next
Combo4_Change
End Sub

Private Sub Combo9_Change()
On Error Resume Next
Text10.SetFocus
End Sub

Private Sub Combo9_Click()
On Error Resume Next
Combo9_Change
End Sub

Private Sub Command1_Click()
On Error Resume Next
Command31_Click
Picture6.Visible = False
Picture13.Visible = False
End Sub

Private Sub Command10_Click()
On Error Resume Next
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
If Label20.Caption = "" Then
MsgBox "ﬁ„ »«·÷€ÿ ⁄·Ï «·»Ì«‰«  «·„—«œ Õ–›Â«", vbCritical
Exit Sub
End If
g = MsgBox("Â·  —Ìœ «·Õ–› Õﬁ« ", vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
Call cont
'**** controle Date
dat1 = Date
dat2 = sr!dat
dat3 = Date
If dat1 < dat2 Then
'MsgBox "€Ì— „„ﬂ‰.. «· «—ÌŒ «·„œŒ· ”«»ﬁ · «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì…", vbCritical+arabic
MsgBox "€Ì— „„ﬂ‰..  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“) ”«»ﬁ · «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì…", vbCritical + arabic
Exit Sub
End If
If dat3 < dat1 Then
MsgBox "€Ì— „„ﬂ‰.. «· «—ÌŒ «·„œŒ· „ ﬁœ„ ⁄‰  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“)", vbCritical + arabic
Exit Sub
End If
'**** end controle Date
'***** 1
If Check2.Value = 0 Then
Call cont
Do While Not ps.EOF
If Label20.Caption = ps!aut Then
ps.Delete
ProgressBar3.Visible = True
ProgressBar3.Value = 0
Timer4.Enabled = True
Exit Sub
End If
ps.MoveNext
Loop
End If
'***** n
If Check2.Value = 1 Then
Call cont
Do While Not ps.EOF
If Label37.Caption = ps!ser And ps!cas = "h" Then
If ps(Label21.Caption) = Label22.Caption Then
If Label21.Caption = "dat" Then
ps.Delete
End If
If Label21.Caption = "mon" Then
ps.Delete
End If
If Label21.Caption = "cla" Then
ps.Delete
End If
End If
End If
If Not ps.EOF Then
ps.MoveNext
End If
Loop
ProgressBar3.Visible = True
ProgressBar3.Value = 0
Timer4.Enabled = True
End If

End If

End Sub

Private Sub Command11_Click()
On Error Resume Next
ProgressBar3.Visible = True
ProgressBar3.Value = 0
Timer4.Enabled = False
Label20.Caption = ""
Check2.Value = 0
Check4.Value = 0
'grd1.Enabled = False
'Command26.Enabled = True
'Call chargec1
'Call chargec3
grd1.Visible = False
Call chargegrd1
SSTab3.Tab = 2
grd1.Visible = True

End Sub

Private Sub Command12_Click()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim m As Double
Dim s As Double
Dim ane As String
Command12.Enabled = False
Call cont2
Do While Not ph.EOF
ph.Delete
ph.MoveNext
Loop
s = 0
m = 0
n = grd1.Rows
For i = 1 To n - 1
grd1.row = i
grd1.Col = 7
m = grd1.Text
s = s + m
grd1.Col = 8
m = grd1.Text
s = s + m
grd1.Col = 9
m = grd1.Text
s = s + m
Next i
For i = 1 To n - 1
ph.AddNew
ph!nom = Label38.Caption
ph!ser = Label37.Caption
ph!dat1 = DT4.Value
ph!dat2 = DT5.Value
grd1.row = i
grd1.Col = 1
ph!dat = grd1.Text
grd1.Col = 2
ph!cla = grd1.Text
grd1.Col = 3
ph!de = grd1.Text
grd1.Col = 4
ph!a = grd1.Text
grd1.Col = 5
ph!nbr = grd1.Text
grd1.Col = 6
ph!mon = grd1.Text
grd1.Col = 7
ph!tot = grd1.Text
grd1.Col = 8
ph!prm = grd1.Text
grd1.Col = 9
ph!rtr = grd1.Text
ph!tos = s
ph.Update
Next i
tim = 2
Timer8.Enabled = True

End Sub

Private Sub Command13_Click()
On Error Resume Next
If Label37.Caption = "" Then
MsgBox "ÌÃ» «œŒ«· «·—ﬁ„ «· ”·”·Ì ··√” «– À„ «·÷€ÿ ⁄·Ï Enter", vbCritical + arabic
Exit Sub
End If
Call chargec1
Call chargec3
grd1.Visible = False
grd2.Visible = False
grd3.Visible = False
Call chargegrd1
grd1.Visible = True
grd2.Visible = True
grd3.Visible = True
SSTab3.Tab = st1
SSTab3.Visible = True
End Sub

Private Sub Command14_Click()
On Error Resume Next
Dim dat1 As Date
Dim dat2 As Date
dat1 = DT4.Value
dat2 = DT5.Value
If dat2 < dat1 Then
MsgBox " «—ÌŒ «·»œ«Ì… ÌÃ» √‰ ÌﬂÊ‰ ﬁ»·  «—ÌŒ «·‰Â«Ì…", vbCritical
Exit Sub
End If
grd1.Visible = False
Call chargegrd2
grd1.Visible = True
Command12.Enabled = True
End Sub

Private Sub Command15_Click()
On Error Resume Next
Text1.Text = ""
Picture6.Visible = False
Call chargegrd6
Picture13.Visible = True
End Sub

Private Sub Command16_Click()
On Error Resume Next
Dim dat1 As Date
Dim dat2 As Date
dat1 = DT7.Value
dat2 = DT8.Value
If dat2 < dat1 Then
MsgBox " «—ÌŒ «·»œ«Ì… ÌÃ» √‰ ÌﬂÊ‰ ﬁ»·  «—ÌŒ «·‰Â«Ì…", vbCritical
Exit Sub
End If
grd2.Visible = False
Call chargegrd3
grd2.Visible = True
Command17.Enabled = True

End Sub

Private Sub Command17_Click()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim m As Double
Dim s As Double
Dim ane As String
Command17.Enabled = False
Call cont2
Do While Not pm.EOF
pm.Delete
pm.MoveNext
Loop
s = 0
m = 0
n = grd2.Rows
For i = 1 To n - 1
grd2.row = i
grd2.Col = 4
m = grd2.Text
s = s + m
grd2.Col = 5
m = grd2.Text
s = s + m
grd2.Col = 6
m = grd2.Text
s = s + m
Next i
For i = 1 To n - 1
pm.AddNew
pm!nom = Label38.Caption
pm!ser = Label37.Caption
pm!dat1 = DT7.Value
pm!dat2 = DT8.Value
grd2.row = i
grd2.Col = 1
pm!dat = grd2.Text
grd2.Col = 2
pm!cla = grd2.Text
grd2.Col = 3
pm!moi = grd2.Text
grd2.Col = 4
pm!mon = grd2.Text
pm!tot = grd2.Text
grd2.Col = 5
pm!prm = grd2.Text
grd2.Col = 6
pm!rtr = grd2.Text
pm!tos = s
pm.Update
Next i
tim = 3
Timer8.Enabled = True


End Sub

Private Sub Command18_Click()
On Error Resume Next
ProgressBar4.Visible = True
ProgressBar4.Value = 0
Timer5.Enabled = False
Label20.Caption = ""
Check1.Value = 0
Check5.Value = 0
'grd2.Enabled = False
'Command28.Enabled = True
'Call chargec1
'Call chargec3
grd2.Visible = False
Call chargegrd1
SSTab3.Tab = 1
grd2.Visible = True

End Sub

Private Sub Command19_Click()
On Error Resume Next
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
If Label20.Caption = "" Then
MsgBox "ﬁ„ »«·÷€ÿ ⁄·Ï «·»Ì«‰«  «·„—«œ Õ–›Â«", vbCritical
Exit Sub
End If
g = MsgBox("Â·  —Ìœ «·Õ–› Õﬁ« ", vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
Call cont
'**** controle Date
dat1 = Date
dat2 = sr!dat
dat3 = Date
If dat1 < dat2 Then
'MsgBox "€Ì— „„ﬂ‰.. «· «—ÌŒ «·„œŒ· ”«»ﬁ · «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì…", vbCritical+arabic
MsgBox "€Ì— „„ﬂ‰..  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“) ”«»ﬁ · «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì…", vbCritical + arabic
Exit Sub
End If
If dat3 < dat1 Then
MsgBox "€Ì— „„ﬂ‰.. «· «—ÌŒ «·„œŒ· „ ﬁœ„ ⁄‰  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“)", vbCritical + arabic
Exit Sub
End If
'***** 1
If Check1.Value = 0 Then
Call cont
Do While Not ps.EOF
If Label20.Caption = ps!aut Then
ps.Delete
ProgressBar4.Visible = True
ProgressBar4.Value = 0
Timer5.Enabled = True
Exit Sub
End If
ps.MoveNext
Loop
End If
'***** n
If Check1.Value = 1 Then
Call cont
Do While Not ps.EOF
If Label37.Caption = ps!ser And ps!cas = "m" Then
If ps(Label21.Caption) = Label22.Caption Then
If Label21.Caption = "dat" Then
ps.Delete
End If
If Label21.Caption = "mon" Then
ps.Delete
End If
If Label21.Caption = "cla" Then
ps.Delete
End If
End If
End If
If Not ps.EOF Then
ps.MoveNext
End If
Loop
ProgressBar4.Visible = True
ProgressBar4.Value = 0
Timer5.Enabled = True
End If

End If

End Sub

Private Sub Command20_Click()
On Error Resume Next
Dim i As Double
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim e As Double
Dim dat1 As Date
Dim dat2 As Date
Text15.Text = Trim(Text15.Text)
Text14.Text = Trim(Text14.Text)
Text13.Text = Trim(Text13.Text)
If Label37.Caption = "" Then
MsgBox "ÌÃ» «œŒ«· «·—ﬁ„ «· ”·”·Ì ··√” «– À„ «·÷€ÿ ⁄·Ï Enter", vbCritical + arabic
Exit Sub
End If
If Combo3.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·‘Â—", vbCritical
Exit Sub
End If
If Combo1.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·ﬁ”„", vbCritical
Exit Sub
End If
If Text15.Text = "" Then
MsgBox "«œŒ· «·„»·€", vbCritical
Text15.SetFocus
Exit Sub
End If
If Text14.Text = "" Then
MsgBox "«œŒ· «·⁄·«Ê« ", vbCritical
Text14.SetFocus
Exit Sub
End If
If Text13.Text = "" Then
MsgBox "«œŒ· «·„ √Œ—« ", vbCritical
Text13.SetFocus
Exit Sub
End If
Call cont
'**** controle Date
dat1 = DT6.Value
dat2 = sr!dat
dat3 = Date
If dat1 < dat2 Then
MsgBox "€Ì— „„ﬂ‰.. «· «—ÌŒ «·„œŒ· ”«»ﬁ · «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì…", vbCritical + arabic
'MsgBox "€Ì— „„ﬂ‰..  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“) ”«»ﬁ · «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì…", vbCritical + arabic
Exit Sub
End If
If dat3 < dat1 Then
MsgBox "€Ì— „„ﬂ‰.. «· «—ÌŒ «·„œŒ· „ ﬁœ„ ⁄‰  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“)", vbCritical + arabic
Exit Sub
End If
'**** end controle Date
c = 1
d = Text15.Text
e = d * c
If Check1.Value = 0 Then
Call cont
Do While Not ps.EOF
dat1 = DT6.Value
dat2 = ps!dat
If ps!cas = "m" And Label20.Caption <> ps!aut And ps!ser = Label37.Caption And ps!moi = Combo3.Text And ps!cla = Combo1.Text Then
MsgBox " „ ÕÃ“ ‰’Ì» Â–« «·ﬁ”„ „‰ Â–« «·‘Â—", vbCritical
Exit Sub
End If
ps.MoveNext
Loop
End If
If Label20.Caption <> "" Then
If Check1.Value = 0 Then
Call cont
Do While Not ps.EOF
If Label20.Caption = ps!aut Then
ps!ser = Label37.Caption
ps!nom = Label38.Caption
ps!cas = "m"
ps!he1 = ""
ps!he2 = ""
ps!nbr = c
ps!cla = Combo1.Text
ps!mon = Text15.Text
ps!tot = e
ps!dat = DT6.Value
ps!moi = Combo3.Text
ps!prm = Text14.Text
ps!rtr = Text13.Text
ps!mois = Label36.Caption
ps.Update
ProgressBar4.Visible = True
ProgressBar4.Value = 0
Timer5.Enabled = True
Exit Sub
End If
ps.MoveNext
Loop
End If
If Check1.Value = 1 Then
Call cont
Do While Not ps.EOF
If Label37.Caption = ps!ser And ps!cas = "m" Then
If ps(Label21.Caption) = Label22.Caption Then
If Label21.Caption = "dat" Then
ps!dat = DT6.Value
ps.Update
End If
If Label21.Caption = "mon" Then
ps!mon = Text15.Text
ps.Update
End If
If Label21.Caption = "cla" Then
ps!cla = Combo1.Text
ps.Update
End If
End If
End If
ps.MoveNext
Loop
ProgressBar4.Visible = True
ProgressBar4.Value = 0
Timer5.Enabled = True
Exit Sub
End If
Exit Sub
End If
ps.AddNew
ps!ser = Label37.Caption
ps!nom = Label38.Caption
ps!cas = "m"
ps!he1 = ""
ps!he2 = ""
ps!nbr = c
ps!cla = Combo1.Text
ps!mon = Text15.Text
ps!tot = e
ps!dat = DT6.Value
ps!moi = Combo3.Text
ps!prm = Text14.Text
ps!rtr = Text13.Text
ps!mois = Label36.Caption
ps.Update
ProgressBar4.Visible = True
ProgressBar4.Value = 0
Timer5.Enabled = True

End Sub

Private Sub Command21_Click()
On Error Resume Next
Dim i As Double
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim e As Double
Dim dat1 As Date
Dim dat2 As Date
Dim m As Double
Text16.Text = Trim(Text16.Text)
Text17.Text = Trim(Text17.Text)
Text18.Text = Trim(Text18.Text)
If Label37.Caption = "" Then
MsgBox "ÌÃ» «œŒ«· «·—ﬁ„ «· ”·”·Ì ··√” «– À„ «·÷€ÿ ⁄·Ï Enter", vbCritical + arabic
Exit Sub
End If
If Combo4.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·‘Â—", vbCritical
Exit Sub
End If
If Combo2.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·ﬁ”„", vbCritical
Exit Sub
End If
If Text16.Text = "" Then
MsgBox "«œŒ· ⁄œœ ”«⁄«  «· œ—Ì”", vbCritical
Text16.SetFocus
Exit Sub
End If
If Text17.Text = "" Then
MsgBox "«œŒ· «·⁄·«Ê« ", vbCritical
Text17.SetFocus
Exit Sub
End If
If Text18.Text = "" Then
MsgBox "«œŒ· «·„ √Œ—« ", vbCritical
Text18.SetFocus
Exit Sub
End If
Call cont
'**** controle Date
dat1 = DT9.Value
dat2 = sr!dat
dat3 = Date
If dat1 < dat2 Then
'MsgBox "€Ì— „„ﬂ‰.. «· «—ÌŒ «·„œŒ· ”«»ﬁ · «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì…", vbCritical+arabic
MsgBox "€Ì— „„ﬂ‰..  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“) ”«»ﬁ · «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì…", vbCritical + arabic
Exit Sub
End If
If dat3 < dat1 Then
MsgBox "€Ì— „„ﬂ‰.. «· «—ÌŒ «·„œŒ· „ ﬁœ„ ⁄‰  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“)", vbCritical + arabic
Exit Sub
End If
'**** end controle Date
c = 1
d = Text16.Text
e = d * c
If Check3.Value = 0 Then
Call cont
Do While Not ps.EOF
dat1 = DT9.Value
dat2 = ps!dat
If ps!cas = "p" And Label20.Caption <> ps!aut And ps!ser = Label37.Caption And ps!moi = Combo4.Text And ps!cla = Combo2.Text Then
MsgBox " „ ÕÃ“ ⁄œœ ”«⁄«  «· œ—Ì” ·Â–« «·ﬁ”„ ›Ì Â–« «·‘Â—", vbCritical
Exit Sub
End If
ps.MoveNext
Loop
End If
m = DT9.Month
If Label20.Caption <> "" Then
If Check3.Value = 0 Then
Call cont
Do While Not ps.EOF
If Label20.Caption = ps!aut Then
ps!ser = Label37.Caption
ps!nom = Label38.Caption
ps!cas = "p"
ps!he1 = ""
ps!he2 = ""
ps!nbr = e
ps!cla = Combo2.Text
ps!mon = ""
ps!tot = "0"
ps!dat = DT9.Value
ps!moi = Combo4.Text
ps!prm = Text17.Text
ps!rtr = Text18.Text
ps!mois = Label36.Caption
ps.Update
ProgressBar5.Visible = True
ProgressBar5.Value = 0
Timer6.Enabled = True
Exit Sub
End If
ps.MoveNext
Loop
End If
If Check3.Value = 1 Then
Call cont
Do While Not ps.EOF
If Label37.Caption = ps!ser And ps!cas = "p" Then
If ps(Label21.Caption) = Label22.Caption Then
If Label21.Caption = "dat" Then
ps!dat = DT9.Value
ps.Update
End If
If Label21.Caption = "nbr" Then
ps!nbr = Text16.Text
ps.Update
End If
If Label21.Caption = "cla" Then
ps!cla = Combo2.Text
ps.Update
End If
End If
End If
ps.MoveNext
Loop
ProgressBar5.Visible = True
ProgressBar5.Value = 0
Timer6.Enabled = True
Exit Sub
End If
Exit Sub
End If
ps.AddNew
ps!ser = Label37.Caption
ps!nom = Label38.Caption
ps!cas = "p"
ps!he1 = ""
ps!he2 = ""
ps!nbr = e
ps!cla = Combo2.Text
ps!mon = ""
ps!tot = "0"
ps!dat = DT9.Value
ps!moi = Combo4.Text
ps!prm = Text17.Text
ps!rtr = Text18.Text
ps!mois = Label36.Caption
ps.Update
ProgressBar5.Visible = True
ProgressBar5.Value = 0
Timer6.Enabled = True
End Sub

Private Sub Command22_Click()
On Error Resume Next
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
If Label20.Caption = "" Then
MsgBox "ﬁ„ »«·÷€ÿ ⁄·Ï «·»Ì«‰«  «·„—«œ Õ–›Â«", vbCritical
Exit Sub
End If
g = MsgBox("Â·  —Ìœ «·Õ–› Õﬁ« ", vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
Call cont
'**** controle Date
dat1 = Date
dat2 = sr!dat
dat3 = Date
If dat1 < dat2 Then
'MsgBox "€Ì— „„ﬂ‰.. «· «—ÌŒ «·„œŒ· ”«»ﬁ · «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì…", vbCritical+arabic
MsgBox "€Ì— „„ﬂ‰..  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“) ”«»ﬁ · «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì…", vbCritical + arabic
Exit Sub
End If
If dat3 < dat1 Then
MsgBox "€Ì— „„ﬂ‰.. «· «—ÌŒ «·„œŒ· „ ﬁœ„ ⁄‰  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“)", vbCritical + arabic
Exit Sub
End If
'***** 1
If Check3.Value = 0 Then
Call cont
Do While Not ps.EOF
If Label20.Caption = ps!aut Then
ps.Delete
ProgressBar5.Visible = True
ProgressBar5.Value = 0
Timer6.Enabled = True
Exit Sub
End If
ps.MoveNext
Loop
End If
'***** n
If Check3.Value = 1 Then
Call cont
Do While Not ps.EOF
If Label37.Caption = ps!ser And ps!cas = "p" Then
If ps(Label21.Caption) = Label22.Caption Then
If Label21.Caption = "dat" Then
ps.Delete
End If
If Label21.Caption = "nbr" Then
ps.Delete
End If
If Label21.Caption = "cla" Then
ps.Delete
End If
End If
End If
If Not ps.EOF Then
ps.MoveNext
End If
Loop
ProgressBar5.Visible = True
ProgressBar5.Value = 0
Timer6.Enabled = True
End If

End If

End Sub

Private Sub Command23_Click()
On Error Resume Next
ProgressBar5.Visible = True
ProgressBar5.Value = 0
Timer6.Enabled = False
Label20.Caption = ""
Check3.Value = 0
Check6.Value = 0
'grd3.Enabled = False
'Command29.Enabled = True
'Call chargec1
'Call chargec3
grd3.Visible = False
Call chargegrd1
SSTab3.Tab = 0
grd3.Visible = True

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
Do While Not pn.EOF
pn.Delete
pn.MoveNext
Loop
s = 0
m = 0
n = grd3.Rows
For i = 1 To n - 1
grd3.row = i
grd3.Col = 5
m = grd3.Text
s = s + m
grd3.Col = 6
m = grd3.Text
s = s + m
Next i
For i = 1 To n - 1
pn.AddNew
pn!nom = Label38.Caption
pn!ser = Label37.Caption
pn!dat1 = DT10.Value
pn!dat2 = DT11.Value
grd3.row = i
grd3.Col = 1
pn!dat = grd3.Text
grd3.Col = 2
pn!cla = grd3.Text
grd3.Col = 3
pn!moi = grd3.Text
grd3.Col = 4
pn!nbr = grd3.Text
grd3.Col = 5
pn!prm = grd3.Text
grd3.Col = 6
pn!rtr = grd3.Text
pn!tos = s
pn.Update
Next i
tim = 4
Timer8.Enabled = True

End Sub

Private Sub Command25_Click()
On Error Resume Next
Dim dat1 As Date
Dim dat2 As Date
dat1 = DT10.Value
dat2 = DT11.Value
If dat2 < dat1 Then
MsgBox " «—ÌŒ «·»œ«Ì… ÌÃ» √‰ ÌﬂÊ‰ ﬁ»·  «—ÌŒ «·‰Â«Ì…", vbCritical
Exit Sub
End If
grd3.Visible = False
Call chargegrd4
grd3.Visible = True
Command24.Enabled = True

End Sub

Private Sub Command27_Click()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim ane As String
Command27.Visible = False
Call cont2
Do While Not np.EOF
np.Delete
np.MoveNext
Loop
n = grd6.Rows
For i = 1 To n - 1
np.AddNew
grd6.row = i
grd6.Col = 1
np!ser = grd6.Text
grd6.Col = 2
np!nom = grd6.Text
grd6.Col = 3
np!tel = grd6.Text
grd6.Col = 4
np!mat = grd6.Text
grd6.Col = 5
np!dat = grd6.Text
grd6.Col = 6
np!adr = grd6.Text
np.Update
Next i
tim = 1
Timer8.Enabled = True

End Sub



Private Sub Command3_Click()
On Error Resume Next
Text6.Text = Trim(Text6.Text)
Text7.Text = Trim(Text7.Text)
Text8.Text = Trim(Text8.Text)
Text9.Text = Trim(Text9.Text)
If Text6.Text = "" Then
MsgBox "«œŒ· —ﬁ„ «·Â« ›", vbCritical
Text6.SetFocus
Exit Sub
End If
If Text9.Text = "" Then
MsgBox "«œŒ· «”„ «·√” «–", vbCritical
Text9.SetFocus
Exit Sub
End If
Label38.Caption = Text9.Text
Call cont
Call cont
Do While Not pr.EOF
If Label37.Caption = pr!ser Then
pr!dat = DT2.Value
pr!tel = Text6.Text
pr!nom = Text9.Text
pr!mat = Text8.Text
pr!adr = Text7.Text
pr!ser = BarcodeX2.Caption
pr.Update
Timer1.Enabled = True
Exit Sub
End If
pr.MoveNext
Loop
End Sub

Private Sub Command31_Click()
On Error Resume Next
Text5.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text1.Text = ""
Call serial
Text2.Enabled = True
Text2.SetFocus
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
DT1.Enabled = True
Command5.Enabled = True
Command4.Enabled = True

End Sub

Private Sub Command4_Click()
On Error Resume Next
Command31_Click
End Sub

Private Sub Command5_Click()
On Error Resume Next
Text2.Text = Trim(Text2.Text)
Text3.Text = Trim(Text3.Text)
Text4.Text = Trim(Text4.Text)
Text5.Text = Trim(Text5.Text)
If Text3.Text = "" Then
MsgBox "«œŒ· —ﬁ„ «·Â« ›", vbCritical
Text3.SetFocus
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "«œŒ· «”„ «·√” «–", vbCritical
Text2.SetFocus
Exit Sub
End If
Text3.Enabled = False
Text2.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Command5.Enabled = False
Command4.Enabled = False
DT1.Enabled = False
Call cont
pr.AddNew
pr!dat = DT1.Value
pr!tel = Text3.Text
pr!nom = Text2.Text
pr!mat = Text4.Text
pr!adr = Text5.Text
pr!ser = BarcodeX1.Caption
pr!act = "1"
pr.Update
Call cont
a = sr!mtr
sr!mtr = a + 1
sr.Update
Timer2.Enabled = True
End Sub

Private Sub Command6_Click()
On Error Resume Next
Picture4.Visible = False
End Sub

Private Sub Command7_Click()
On Error Resume Next
If Label14.Caption = "" Then
MsgBox "·« ÌÊÃœ ·Â– «·√” «– √Ì —ﬁ„  ”·”·Ì Ì„ﬂ‰ «·Õ–› ⁄·Ï √”«”Â", vbCritical
Exit Sub
End If
g = MsgBox("Â·  —Ìœ «·Õ–› Õﬁ« ", vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
Call cont
Do While Not pr.EOF
If Label14.Caption = pr!aut Then
pr!act = "0"
pr.Update
Timer3.Enabled = True
Exit Sub
End If
pr.MoveNext
Loop
End If

End Sub

Private Sub Command8_Click()
On Error Resume Next
Dim i As Double
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim e As Double
Dim dat1 As Date
Dim dat2 As Date
Dim m As Double
Text10.Text = Trim(Text10.Text)
Text11.Text = Trim(Text11.Text)
Text12.Text = Trim(Text12.Text)
If Label37.Caption = "" Then
MsgBox "ÌÃ» «œŒ«· «·—ﬁ„ «· ”·”·Ì ··√” «– À„ «·÷€ÿ ⁄·Ï Enter", vbCritical + arabic
Exit Sub
End If
If Combo7.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— ”«⁄… «·»œ«Ì…", vbCritical
Exit Sub
End If
If Combo8.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— ”«⁄… «·‰Â«Ì…", vbCritical
Exit Sub
End If
If Combo9.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·ﬁ”„", vbCritical
Exit Sub
End If
If Text10.Text = "" Then
MsgBox "«œŒ· «·„»·€", vbCritical
Text10.SetFocus
Exit Sub
End If
If Text11.Text = "" Then
MsgBox "«œŒ· «·⁄·«Ê« ", vbCritical
Text11.SetFocus
Exit Sub
End If
If Text12.Text = "" Then
MsgBox "«œŒ· «·„ √Œ—« ", vbCritical
Text12.SetFocus
Exit Sub
End If
Call cont
'**** controle Date
dat1 = DT3.Value
dat2 = sr!dat
dat3 = Date
If dat1 < dat2 Then
MsgBox "€Ì— „„ﬂ‰.. «· «—ÌŒ «·„œŒ· ”«»ﬁ · «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì…", vbCritical + arabic
'MsgBox "€Ì— „„ﬂ‰..  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“) ”«»ﬁ · «—ÌŒ »œ¡ «·”‰… «·œ—«”Ì…", vbCritical + arabic
Exit Sub
End If
If dat3 < dat1 Then
MsgBox "€Ì— „„ﬂ‰.. «· «—ÌŒ «·„œŒ· „ ﬁœ„ ⁄‰  «—ÌŒ «·ÌÊ„ ( «—ÌŒ «·ÃÂ«“)", vbCritical + arabic
Exit Sub
End If
'**** end controle Date
a = Combo7.Text
b = Combo8.Text
c = 0
If a = 0 Then
a = 24
End If
If b = 0 Then
b = 24
End If
If a > b Then
MsgBox "”«⁄… «·»œ«Ì… ÌÃ» √‰  ﬂÊ‰ ﬁ»· ”«⁄… «·‰Â«Ì…", vbCritical
Exit Sub
End If
c = b - a
d = Text10.Text
e = d * c
If Check2.Value = 0 Then
Call cont
Do While Not ps.EOF
dat1 = DT3.Value
dat2 = ps!dat
If ps!cas = "p" And Label20.Caption <> ps!aut And ps!ser = Label37.Caption And ps!he1 = Combo7.Text And ps!he2 = Combo8.Text And dat1 = dat2 Then
MsgBox "Â–Â «·”«⁄… „ÕÃÊ“… ·ﬁ”„ ¬Œ—", vbCritical
Exit Sub
End If
If Label20.Caption <> ps!aut And ps!cla = Combo9.Text And ps!he1 = Combo7.Text And ps!he2 = Combo8.Text And dat1 = dat2 Then
MsgBox "Â–Â «·”«⁄… „ÕÃÊ“… „‰ ÿ—› √” «– ¬Œ—", vbCritical
Exit Sub
End If
ps.MoveNext
Loop
End If
m = DT3.Month
If Label20.Caption <> "" Then
If Check2.Value = 0 Then
Call cont
Do While Not ps.EOF
If Label20.Caption = ps!aut Then
ps!ser = Label37.Caption
ps!nom = Label38.Caption
ps!cas = "h"
ps!he1 = Combo7.Text
ps!he2 = Combo8.Text
ps!nbr = c
ps!cla = Combo9.Text
ps!mon = Text10.Text
ps!tot = e
ps!dat = DT3.Value
ps!moi = ""
ps!prm = Text11.Text
ps!rtr = Text12.Text
ps!mois = m
ps.Update
ProgressBar3.Visible = True
ProgressBar3.Value = 0
Timer4.Enabled = True
Exit Sub
End If
ps.MoveNext
Loop
End If
If Check2.Value = 1 Then
Call cont
Do While Not ps.EOF
If Label37.Caption = ps!ser And ps!cas = "h" Then
If ps(Label21.Caption) = Label22.Caption Then
If Label21.Caption = "dat" Then
ps!dat = DT3.Value
ps!mois = m
ps.Update
End If
If Label21.Caption = "mon" Then
ps!mon = Text10.Text
ps.Update
End If
If Label21.Caption = "cla" Then
ps!cla = Combo9.Text
ps.Update
End If
End If
End If
ps.MoveNext
Loop
ProgressBar3.Visible = True
ProgressBar3.Value = 0
Timer4.Enabled = True
Exit Sub
End If
Exit Sub
End If
ps.AddNew
ps!ser = Label37.Caption
ps!nom = Label38.Caption
ps!cas = "h"
ps!he1 = Combo7.Text
ps!he2 = Combo8.Text
ps!nbr = c
ps!cla = Combo9.Text
ps!mon = Text10.Text
ps!tot = e
ps!dat = DT3.Value
ps!moi = ""
ps!prm = Text11.Text
ps!rtr = Text12.Text
ps!mois = m
ps.Update
ProgressBar3.Visible = True
ProgressBar3.Value = 0
Timer4.Enabled = True

End Sub

Private Sub Command9_Click()
On Error Resume Next
Dim t As String
Dim a As Integer
Dim cl1 As String
Dim cl2 As String
a = 0
SSTab2.Tab = 1
Call cont
Do While Not pr.EOF
If pr!act = "1" Then
If Text1.Text = pr!ser Or Val(Text1.Text) = Val(pr!ser) Then
a = 1
Label14.Caption = pr!aut
Label38.Caption = pr!nom
Text9.Text = pr!nom
Text6.Text = pr!tel
DT2.Value = pr!dat
Text7.Text = pr!adr
Text8.Text = pr!mat
BarcodeX2.Caption = pr!ser
Label37.Caption = pr!ser
pr.MoveLast
End If
End If
pr.MoveNext
Loop
If a = 1 Then
Picture4.Visible = False
Picture8.Visible = True
Picture6.Visible = True
SSTab2.Visible = True
Else
MsgBox "«·—ﬁ„ «· ”·”·Ì «·„œŒ· €Ì— „Œ“‰ .. Ì—ÃÏ «· √ﬂœ „‰Â", vbExclamation
Text1.Text = ""
Text1.SetFocus
End If

End Sub

Private Sub Form_Load()
'On Error Resume Next
Me.Top = 100
Skin1.LoadSkin App.Path & "\green.skn"
Skin1.ApplySkin Me.hWnd
DT1.Value = Date
DT2.Value = Date
DT3.Value = Date
DT4.Value = Date - 30
DT5.Value = Date
DT6.Value = Date
DT7.Value = Date - 30
DT8.Value = Date
DT9.Value = Date
DT10.Value = Date - 30
DT11.Value = Date
st1 = 2
End Sub
Private Sub serial()
On Error Resume Next
Dim a As Double
Dim b As String
Call cont
a = Val(sr!mtr)
If a > 999999 Then
b = a
ElseIf a > 99999 And a <= 999999 Then
b = a
b = "0" + b
ElseIf a > 9999 And a <= 99999 Then
b = a
b = "00" + b
ElseIf a > 999 And a <= 9999 Then
b = a
b = "000" + b
ElseIf a > 99 And a <= 999 Then
b = a
b = "0000" + b
ElseIf a > 9 And a <= 99 Then
b = a
b = "00000" + b
ElseIf a > 0 And a <= 9 Then
b = a
b = "000000" + b
End If
If a >= 99999999 Then
MsgBox "«·—ﬁ„ «· ”·”·Ì ·«Ì„ﬂ‰ √‰ Ì’· ≈·Ï „«∆… „·ÌÊ‰", vbCritical
Exit Sub
End If
BarcodeX1.Caption = b
Picture4.Visible = True

End Sub

Private Sub grd1_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim tx1 As String
Dim tx2 As String
Dim tx3 As String
i = grd1.row
j = grd1.Col
If Check4.Value = 1 Then
If i > 0 Then
grd1.row = i
grd1.Col = 0
Label20.Caption = grd1.Text
grd1.Col = 1
DT3.Value = grd1.Text
grd1.Col = 2
Combo9.Text = grd1.Text
grd1.Col = 3
Combo7.Text = grd1.Text
grd1.Col = 4
Combo8.Text = grd1.Text
grd1.Col = 6
Text10.Text = grd1.Text
grd1.Col = 8
Text11.Text = grd1.Text
grd1.Col = 9
Text12.Text = grd1.Text
ProgressBar3.Visible = True
Check2.Value = 0
grd1.row = i
grd1.Col = j
tx1 = grd1.Text
If j = 1 Then
tx2 = "Ã„Ì⁄ «· Ê—ÌŒ «·„„«À·… ·‹ " + tx1
ProgressBar3.Visible = False
Label21.Caption = "dat"
Label22.Caption = tx1
End If
If j = 2 Then
tx2 = "Ã„Ì⁄ «·√ﬁ”«„ «·„„«À·… ·‹ " + tx1
ProgressBar3.Visible = False
Label21.Caption = "cla"
Label22.Caption = tx1
End If
If j = 6 Then
tx2 = "Ã„Ì⁄ «·„»«·€ «·„„«À·… ·‹ " + tx1
ProgressBar3.Visible = False
Label21.Caption = "mon"
Label22.Caption = tx1
End If
Label68.Caption = tx2
End If
End If
End Sub

Private Sub grd2_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim tx1 As String
Dim tx2 As String
Dim tx3 As String
i = grd2.row
j = grd2.Col
If Check5.Value = 1 Then
If i > 0 Then
grd2.row = i
grd2.Col = 0
Label20.Caption = grd2.Text
grd2.Col = 1
DT6.Value = grd2.Text
grd2.Col = 2
Combo1.Text = grd2.Text
grd2.Col = 3
Combo3.Text = grd2.Text
grd2.Col = 4
Text15.Text = grd2.Text
grd2.Col = 5
Text14.Text = grd2.Text
grd2.Col = 6
Text13.Text = grd2.Text
ProgressBar4.Visible = True
Check1.Value = 0
grd2.row = i
grd2.Col = j
tx1 = grd2.Text
If j = 1 Then
tx2 = "Ã„Ì⁄ «· Ê—ÌŒ «·„„«À·… ·‹ " + tx1
ProgressBar4.Visible = False
Label21.Caption = "dat"
Label22.Caption = tx1
End If
If j = 2 Then
tx2 = "Ã„Ì⁄ «·√ﬁ”«„ «·„„«À·… ·‹ " + tx1
ProgressBar4.Visible = False
Label21.Caption = "cla"
Label22.Caption = tx1
End If
If j = 4 Then
tx2 = "Ã„Ì⁄ «·„»«·€ «·„„«À·… ·‹ " + tx1
ProgressBar4.Visible = False
Label21.Caption = "mon"
Label22.Caption = tx1
End If
Label25.Caption = tx2
End If
End If
End Sub

Private Sub grd3_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim tx1 As String
Dim tx2 As String
Dim tx3 As String
i = grd3.row
j = grd3.Col
If Check6.Value = 1 Then
If i > 0 Then
grd3.row = i
grd3.Col = 0
Label20.Caption = grd3.Text
grd3.Col = 1
DT9.Value = grd3.Text
grd3.Col = 2
Combo2.Text = grd3.Text
grd3.Col = 3
Combo4.Text = grd3.Text
grd3.Col = 4
Text16.Text = grd3.Text
grd3.Col = 5
Text17.Text = grd3.Text
grd3.Col = 6
Text18.Text = grd3.Text
ProgressBar5.Visible = True
Check3.Value = 0
grd3.row = i
grd3.Col = j
tx1 = grd3.Text
If j = 1 Then
tx2 = "Ã„Ì⁄ «· Ê—ÌŒ «·„„«À·… ·‹ " + tx1
ProgressBar5.Visible = False
Label21.Caption = "dat"
Label22.Caption = tx1
End If
If j = 2 Then
tx2 = "Ã„Ì⁄ «·√ﬁ”«„ «·„„«À·… ·‹ " + tx1
ProgressBar5.Visible = False
Label21.Caption = "cla"
Label22.Caption = tx1
End If
If j = 4 Then
tx2 = "Ã„Ì⁄ «·”«⁄«  «·„„«À·… ·‹ " + tx1
ProgressBar5.Visible = False
Label21.Caption = "nbr"
Label22.Caption = tx1
End If
Label42.Caption = tx2
End If
End If
End Sub

Private Sub grd6_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
i = grd6.Col
j = grd6.row
If i = 7 Then
grd6.row = j
grd6.Col = 1
Text1.Text = grd6.Text
Command9_Click
End If
End Sub

Private Sub Text1_Change()
On Error Resume Next
Picture4.Visible = False
Picture6.Visible = False
Picture13.Visible = False
Text10.Text = ""
Text15.Text = ""
Text16.Text = ""
SSTab2.Visible = False
SSTab3.Visible = False
End Sub

Private Sub Text1_Click()
On Error Resume Next
Text1_Change
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If Text1.Text <> "" Then
If KeyCode = 13 Then
Command9_Click
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

Private Sub Text13_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim i As Integer
Dim j As Integer
Dim n As Double
Dim vg As String
Text13.Text = Trim(Text13.Text)
n = Len(Text13.Text)
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
vg = Mid$(Text13.Text, i, 1)
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

Private Sub Text14_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim i As Integer
Dim j As Integer
Dim n As Double
Dim vg As String
Text14.Text = Trim(Text14.Text)
n = Len(Text14.Text)
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
vg = Mid$(Text14.Text, i, 1)
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

Private Sub Text15_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim i As Integer
Dim j As Integer
Dim n As Double
Dim vg As String
Text15.Text = Trim(Text15.Text)
n = Len(Text15.Text)
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
vg = Mid$(Text15.Text, i, 1)
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

Private Sub Text16_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim i As Integer
Dim j As Integer
Dim n As Double
Dim vg As String
Text16.Text = Trim(Text16.Text)
n = Len(Text16.Text)
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
vg = Mid$(Text16.Text, i, 1)
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

Private Sub Text17_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim i As Integer
Dim j As Integer
Dim n As Double
Dim vg As String
Text17.Text = Trim(Text17.Text)
n = Len(Text17.Text)
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
vg = Mid$(Text17.Text, i, 1)
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

Private Sub Text18_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim i As Integer
Dim j As Integer
Dim n As Double
Dim vg As String
Text18.Text = Trim(Text18.Text)
n = Len(Text18.Text)
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
vg = Mid$(Text18.Text, i, 1)
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

Private Sub Text6_KeyPress(KeyAscii As Integer)
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
ProgressBar1.Value = 0
Timer1.Enabled = False
End If

End Sub

Private Sub Timer2_Timer()
On Error Resume Next
ProgressBar2.Value = ProgressBar2.Value + 8
If ProgressBar2.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
ProgressBar2.Value = 0
Timer2.Enabled = False
End If


End Sub

Private Sub Timer3_Timer()
On Error Resume Next
ProgressBar1.Value = ProgressBar1.Value + 8
If ProgressBar1.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
ProgressBar1.Value = 0
Timer3.Enabled = False
Text1.Text = ""
Text1.SetFocus
End If

End Sub
Public Sub chargec3()
On Error Resume Next
Combo7.Clear
Combo8.Clear
Combo7.AddItem "08"
Combo8.AddItem "08"
Combo7.AddItem "09"
Combo8.AddItem "09"
Combo7.AddItem "10"
Combo8.AddItem "10"
Combo7.AddItem "11"
Combo8.AddItem "11"
Combo7.AddItem "12"
Combo8.AddItem "12"
Combo7.AddItem "13"
Combo8.AddItem "13"
Combo7.AddItem "14"
Combo8.AddItem "14"
Combo7.AddItem "15"
Combo8.AddItem "15"
Combo7.AddItem "16"
Combo8.AddItem "16"
Combo7.AddItem "17"
Combo8.AddItem "17"
Combo7.AddItem "18"
Combo8.AddItem "18"
Combo7.AddItem "19"
Combo8.AddItem "19"
Combo7.AddItem "20"
Combo8.AddItem "20"
Combo7.AddItem "21"
Combo8.AddItem "21"
Combo7.AddItem "22"
Combo8.AddItem "22"
Combo7.AddItem "23"
Combo8.AddItem "23"
Combo7.AddItem "00"
Combo8.AddItem "00"
End Sub
Public Sub chargec1()
On Error Resume Next
Call cont
Combo9.Clear
Combo1.Clear
Combo2.Clear
  Do While Not cl.EOF
    Combo9.AddItem cl!cla
    Combo1.AddItem cl!cla
    Combo2.AddItem cl!cla
cl.MoveNext
  Loop
End Sub
Private Sub chargegrd1()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim k As Double
Dim a As Double
Dim b As Double
Dim c As Double
grd1.Clear
grd1.Rows = 1
grd1.Cols = 10
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1500
grd1.ColWidth(2) = 1500
grd1.ColWidth(3) = 1000
grd1.ColWidth(4) = 1000
grd1.ColWidth(5) = 500
grd1.ColWidth(6) = 1750
grd1.ColWidth(7) = 1750
grd1.ColWidth(8) = 1750
grd1.ColWidth(9) = 1750
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 1
grd1.ColAlignment(6) = 1
grd1.ColAlignment(7) = 1
grd1.ColAlignment(8) = 1
grd1.ColAlignment(9) = 1
grd1.row = 0
grd1.Col = 1
grd1.Text = "«· «—ÌŒ"
grd1.Col = 2
grd1.Text = "«·ﬁ”„"
grd1.Col = 3
grd1.Text = "„‰"
grd1.Col = 4
grd1.Text = "≈·Ï"
grd1.Col = 5
grd1.Text = "«·⁄œœ"
grd1.Col = 6
grd1.Text = "«·„»·€"
grd1.Col = 7
grd1.Text = "«·«Ã„«·Ì"
grd1.Col = 8
grd1.Text = "«·⁄·«Ê« "
grd1.Col = 9
grd1.Text = "«·„ √Œ—« "
'****** grd2
grd2.Clear
grd2.Rows = 1
grd2.Cols = 7
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 1500
grd2.ColWidth(2) = 1500
grd2.ColWidth(3) = 2000
grd2.ColWidth(4) = 2000
grd2.ColWidth(5) = 2000
grd2.ColWidth(6) = 3800
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.ColAlignment(4) = 1
grd2.ColAlignment(5) = 1
grd2.ColAlignment(6) = 1
grd2.row = 0
grd2.Col = 1
grd2.Text = "«· «—ÌŒ"
grd2.Col = 2
grd2.Text = "«·ﬁ”„"
grd2.Col = 3
grd2.Text = "«·‘Â—"
grd2.Col = 4
grd2.Text = "«·„»·€"
grd2.Col = 5
grd2.Text = "«·⁄·«Ê« "
grd2.Col = 6
grd2.Text = "«·„ √Œ—« "
'****** grd3
grd3.Clear
grd3.Rows = 1
grd3.Cols = 7
grd3.ColWidth(0) = 0
grd3.ColWidth(1) = 1500
grd3.ColWidth(2) = 2200
grd3.ColWidth(3) = 2200
grd3.ColWidth(4) = 2500
grd3.ColWidth(5) = 2200
grd3.ColWidth(6) = 2200
grd3.ColAlignment(1) = 1
grd3.ColAlignment(2) = 1
grd3.ColAlignment(3) = 1
grd3.ColAlignment(4) = 1
grd3.ColAlignment(5) = 1
grd3.ColAlignment(6) = 1
grd3.row = 0
grd3.Col = 1
grd3.Text = "«· «—ÌŒ"
grd3.Col = 2
grd3.Text = "«·ﬁ”„"
grd3.Col = 3
grd3.Text = "«·‘Â—"
grd3.Col = 4
grd3.Text = "⁄œœ ”«⁄«  «· œ—Ì”"
grd3.Col = 5
grd3.Text = "«·⁄·«Ê« "
grd3.Col = 6
grd3.Text = "«·„ √Œ—« "
i = 1
j = 1
k = 1
Command12.Enabled = False
Command17.Enabled = False
Command24.Enabled = False
Call cont
grd1.Rows = ps.RecordCount + 3
grd2.Rows = ps.RecordCount + 3
grd3.Rows = ps.RecordCount + 3
Do While Not ps.EOF
If ps!ser = Label37.Caption Then
'**** H
If ps!cas = "h" Then
grd1.row = i
grd1.Col = 0
grd1.Text = ps!aut
grd1.Col = 1
grd1.Text = ps!dat
grd1.Col = 2
grd1.Text = ps!cla
grd1.Col = 3
grd1.Text = ps!he1
grd1.Col = 4
grd1.Text = ps!he2
grd1.Col = 5
grd1.Text = ps!nbr
grd1.Col = 6
grd1.Text = ps!mon
a = ps!nbr
b = ps!mon
c = a * b
grd1.Col = 7
grd1.Text = c
grd1.Col = 8
grd1.Text = ps!prm
grd1.Col = 9
grd1.Text = ps!rtr
i = i + 1
Text10.Text = ps!mon
st1 = 2
End If
'**** M
If ps!cas = "m" Then
grd2.row = j
grd2.Col = 0
grd2.Text = ps!aut
grd2.Col = 1
grd2.Text = ps!dat
grd2.Col = 2
grd2.Text = ps!cla
grd2.Col = 3
grd2.Text = ps!moi
grd2.Col = 4
grd2.Text = ps!mon
grd2.Col = 5
grd2.Text = ps!prm
grd2.Col = 6
grd2.Text = ps!rtr
j = j + 1
Text15.Text = ps!mon
st1 = 1
End If
'**** P
If ps!cas = "p" Then
grd3.row = k
grd3.Col = 0
grd3.Text = ps!aut
grd3.Col = 1
grd3.Text = ps!dat
grd3.Col = 2
grd3.Text = ps!cla
grd3.Col = 3
grd3.Text = ps!moi
grd3.Col = 4
grd3.Text = ps!nbr
grd3.Col = 5
grd3.Text = ps!prm
grd3.Col = 6
grd3.Text = ps!rtr
k = k + 1
Text16.Text = ""
st1 = 0
End If
End If
ps.MoveNext
Loop
grd1.Rows = i
grd1.Col = 0
grd1.Sort = 2
grd2.Rows = j
grd2.Col = 0
grd2.Sort = 2
grd3.Rows = k
grd3.Col = 0
grd3.Sort = 2

End Sub

Private Sub Timer4_Timer()
On Error Resume Next
ProgressBar3.Value = ProgressBar3.Value + 8
If ProgressBar3.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
ProgressBar3.Value = 0
Timer4.Enabled = False
Label20.Caption = ""
Check2.Value = 0
Check4.Value = 0
'grd1.Enabled = False
'Command26.Enabled = True
'Call chargec1
'Call chargec3
grd1.Visible = False
Call chargegrd1
'SSTab3.Tab = ts1
grd1.Visible = True
End If

End Sub

Private Sub Timer5_Timer()
On Error Resume Next
ProgressBar4.Value = ProgressBar4.Value + 8
If ProgressBar4.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
ProgressBar4.Value = 0
Timer5.Enabled = False
Label20.Caption = ""
Check1.Value = 0
Check5.Value = 0
'grd2.Enabled = False
'Command28.Enabled = True
'Call chargec1
'Call chargec3
grd2.Visible = False
Call chargegrd1
'SSTab3.Tab = ts1
grd2.Visible = True
End If

End Sub

Private Sub Timer6_Timer()
On Error Resume Next
ProgressBar5.Value = ProgressBar5.Value + 8
If ProgressBar5.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
ProgressBar5.Value = 0
Timer6.Enabled = False
Label20.Caption = ""
Check3.Value = 0
Check6.Value = 0
'grd3.Enabled = False
'Command29.Enabled = True
'Call chargec1
'Call chargec3
grd3.Visible = False
Call chargegrd1
'SSTab3.Tab = ts1
grd3.Visible = True
End If

End Sub
Private Sub chargegrd2()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim k As Double
Dim a As Double
Dim b As Double
Dim c As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
grd1.Clear
grd1.Rows = 1
grd1.Cols = 10
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1500
grd1.ColWidth(2) = 1500
grd1.ColWidth(3) = 1000
grd1.ColWidth(4) = 1000
grd1.ColWidth(5) = 500
grd1.ColWidth(6) = 1750
grd1.ColWidth(7) = 1750
grd1.ColWidth(8) = 1750
grd1.ColWidth(9) = 1750
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 1
grd1.ColAlignment(6) = 1
grd1.ColAlignment(7) = 1
grd1.ColAlignment(8) = 1
grd1.ColAlignment(9) = 1
grd1.row = 0
grd1.Col = 1
grd1.Text = "«· «—ÌŒ"
grd1.Col = 2
grd1.Text = "«·ﬁ”„"
grd1.Col = 3
grd1.Text = "„‰"
grd1.Col = 4
grd1.Text = "≈·Ï"
grd1.Col = 5
grd1.Text = "«·⁄œœ"
grd1.Col = 6
grd1.Text = "«·„»·€"
grd1.Col = 7
grd1.Text = "«·«Ã„«·Ì"
grd1.Col = 8
grd1.Text = "«·⁄·«Ê« "
grd1.Col = 9
grd1.Text = "«·„ √Œ—« "
i = 1
dat1 = DT4.Value
dat2 = DT5.Value
Call cont
grd1.Rows = ps.RecordCount + 3
Do While Not ps.EOF
dat3 = ps!dat
If ps!ser = Label37.Caption Then
'**** H
If ps!cas = "h" Then
If dat3 >= dat1 And dat3 <= dat2 Then
grd1.row = i
grd1.Col = 0
grd1.Text = ps!aut
grd1.Col = 1
grd1.Text = ps!dat
grd1.Col = 2
grd1.Text = ps!cla
grd1.Col = 3
grd1.Text = ps!he1
grd1.Col = 4
grd1.Text = ps!he2
grd1.Col = 5
grd1.Text = ps!nbr
grd1.Col = 6
grd1.Text = ps!mon
a = ps!nbr
b = ps!mon
c = a * b
grd1.Col = 7
grd1.Text = c
grd1.Col = 8
grd1.Text = ps!prm
grd1.Col = 9
grd1.Text = ps!rtr
i = i + 1
Text10.Text = ps!mon
st1 = 2
End If
End If
End If
ps.MoveNext
Loop
grd1.Rows = i
grd1.Col = 0
grd1.Sort = 2

End Sub
Private Sub chargegrd3()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim k As Double
Dim a As Double
Dim b As Double
Dim c As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
'****** grd2
grd2.Clear
grd2.Rows = 1
grd2.Cols = 7
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 1500
grd2.ColWidth(2) = 1500
grd2.ColWidth(3) = 2000
grd2.ColWidth(4) = 2000
grd2.ColWidth(5) = 2000
grd2.ColWidth(6) = 3800
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.ColAlignment(4) = 1
grd2.ColAlignment(5) = 1
grd2.ColAlignment(6) = 1
grd2.row = 0
grd2.Col = 1
grd2.Text = "«· «—ÌŒ"
grd2.Col = 2
grd2.Text = "«·ﬁ”„"
grd2.Col = 3
grd2.Text = "«·‘Â—"
grd2.Col = 4
grd2.Text = "«·„»·€"
grd2.Col = 5
grd2.Text = "«·⁄·«Ê« "
grd2.Col = 6
grd2.Text = "«·„ √Œ—« "
j = 1
dat1 = DT7.Value
dat2 = DT8.Value
Call cont
grd2.Rows = ps.RecordCount + 3
Do While Not ps.EOF
dat3 = ps!dat
If ps!ser = Label37.Caption Then
'**** M
If ps!cas = "m" Then
If dat3 >= dat1 And dat3 <= dat2 Then
grd2.row = j
grd2.Col = 0
grd2.Text = ps!aut
grd2.Col = 1
grd2.Text = ps!dat
grd2.Col = 2
grd2.Text = ps!cla
grd2.Col = 3
grd2.Text = ps!moi
grd2.Col = 4
grd2.Text = ps!mon
grd2.Col = 5
grd2.Text = ps!prm
grd2.Col = 6
grd2.Text = ps!rtr
j = j + 1
Text15.Text = ps!mon
st1 = 1
End If
End If
End If
ps.MoveNext
Loop
grd2.Rows = j
grd2.Col = 0
grd2.Sort = 2

End Sub
Private Sub chargegrd4()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim k As Double
Dim a As Double
Dim b As Double
Dim c As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
'****** grd3
grd3.Clear
grd3.Rows = 1
grd3.Cols = 7
grd3.ColWidth(0) = 0
grd3.ColWidth(1) = 1500
grd3.ColWidth(2) = 2200
grd3.ColWidth(3) = 2200
grd3.ColWidth(4) = 2500
grd3.ColWidth(5) = 2200
grd3.ColWidth(6) = 2200
grd3.ColAlignment(1) = 1
grd3.ColAlignment(2) = 1
grd3.ColAlignment(3) = 1
grd3.ColAlignment(4) = 1
grd3.ColAlignment(5) = 1
grd3.ColAlignment(6) = 1
grd3.row = 0
grd3.Col = 1
grd3.Text = "«· «—ÌŒ"
grd3.Col = 2
grd3.Text = "«·ﬁ”„"
grd3.Col = 3
grd3.Text = "«·‘Â—"
grd3.Col = 4
grd3.Text = "⁄œœ ”«⁄«  «· œ—Ì”"
grd3.Col = 5
grd3.Text = "«·⁄·«Ê« "
grd3.Col = 6
grd3.Text = "«·„ √Œ—« "
k = 1
dat1 = DT7.Value
dat2 = DT8.Value
Call cont
grd3.Rows = ps.RecordCount + 3
Do While Not ps.EOF
dat3 = ps!dat
If ps!ser = Label37.Caption Then
'**** P
If ps!cas = "p" Then
If dat3 >= dat1 And dat3 <= dat2 Then
grd3.row = k
grd3.Col = 0
grd3.Text = ps!aut
grd3.Col = 1
grd3.Text = ps!dat
grd3.Col = 2
grd3.Text = ps!cla
grd3.Col = 3
grd3.Text = ps!moi
grd3.Col = 4
grd3.Text = ps!nbr
grd3.Col = 5
grd3.Text = ps!prm
grd3.Col = 6
grd3.Text = ps!rtr
k = k + 1
Text16.Text = ""
st1 = 0
End If
End If
End If
ps.MoveNext
Loop
grd3.Rows = k
grd3.Col = 0
grd3.Sort = 2

End Sub
Private Sub chargegrd6()
On Error Resume Next
Dim i As Double
grd6.Clear
grd6.Cols = 8
grd6.Rows = 1
grd6.ColWidth(0) = 0
grd6.ColWidth(1) = 1400
grd6.ColWidth(2) = 3000
grd6.ColWidth(3) = 1500
grd6.ColWidth(4) = 1500
grd6.ColWidth(5) = 1500
grd6.ColWidth(6) = 2000
grd6.ColWidth(7) = 1000
grd6.ColAlignment(1) = 1
grd6.ColAlignment(2) = 1
grd6.ColAlignment(3) = 1
grd6.ColAlignment(4) = 1
grd6.ColAlignment(5) = 1
grd6.ColAlignment(6) = 1
grd6.ColAlignment(7) = 1
grd6.row = 0
grd6.Col = 1
grd6.Text = "«·—ﬁ„ «· ”·”·Ì"
grd6.Col = 2
grd6.Text = "«·«”„"
grd6.Col = 3
grd6.Text = "«·Â« ›"
grd6.Col = 4
grd6.Text = "«·„«œ…"
grd6.Col = 5
grd6.Text = " «—ÌŒ «·«‰ ”«»"
grd6.Col = 6
grd6.Text = "«·⁄‰Ê«‰"
grd6.Col = 7
grd6.Text = ""
i = 1
Call cont
grd6.Rows = pr.RecordCount + 3
Do While Not pr.EOF
If pr!act = "1" Then
grd6.row = i
grd6.Col = 0
grd6.Text = pr!aut
grd6.Col = 1
grd6.Text = pr!ser
grd6.Col = 2
grd6.Text = pr!nom
grd6.Col = 3
grd6.Text = pr!tel
grd6.Col = 4
grd6.Text = pr!mat
grd6.Col = 5
grd6.Text = pr!dat
grd6.Col = 6
grd6.Text = pr!adr
grd6.Col = 7
grd6.Text = "⁄—÷"
i = i + 1
End If
pr.MoveNext
Loop
grd6.Rows = i
grd6.Col = 1
grd6.Sort = 1
End Sub

Private Sub Timer8_Timer()
On Error Resume Next
'On Error Resume Next
Dim ane As String
ProgressBar6.Value = ProgressBar6.Value + 8
If ProgressBar6.Value > 90 Then
Timer8.Enabled = False
ProgressBar6.Value = 0
Call cont2
'******** noms prof
If tim = 1 Then
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
'data.DoCmd.ApplyFilter cla
data.DoCmd.Maximize
'data.DoCmd.OpenReport "Bulletin", acViewPreview, , "num =" & a, acWindowNormal, OpenArgs
data.DoCmd.OpenReport "Tprofesseurs", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
'data.DCount
    'data.CloseCurrentDatabase
data.Visible = True
'Set data = Nothing
Command27.Visible = True
End If
'******** fich heurs pres prof
If tim = 2 Then
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
'data.DoCmd.ApplyFilter cla
data.DoCmd.Maximize
'data.DoCmd.OpenReport "Bulletin", acViewPreview, , "num =" & a, acWindowNormal, OpenArgs
data.DoCmd.OpenReport "Tpresencesh", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
'data.DCount
    'data.CloseCurrentDatabase
data.Visible = True
'Set data = Nothing
Command12.Enabled = True
End If
'******** fich mois pres prof
If tim = 3 Then
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
'data.DoCmd.ApplyFilter cla
data.DoCmd.Maximize
'data.DoCmd.OpenReport "Bulletin", acViewPreview, , "num =" & a, acWindowNormal, OpenArgs
data.DoCmd.OpenReport "Tpresencesm", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
'data.DCount
    'data.CloseCurrentDatabase
data.Visible = True
'Set data = Nothing
Command17.Enabled = True
End If
'******** fich pourcentage pres prof
If tim = 4 Then
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
'data.DoCmd.ApplyFilter cla
data.DoCmd.Maximize
'data.DoCmd.OpenReport "Bulletin", acViewPreview, , "num =" & a, acWindowNormal, OpenArgs
data.DoCmd.OpenReport "Tpresencesp", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
'data.DCount
    'data.CloseCurrentDatabase
data.Visible = True
'Set data = Nothing
Command24.Enabled = True
End If


End If

End Sub
