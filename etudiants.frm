VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ActiveSkin.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{8E515444-86DF-11D3-A630-444553540001}#1.0#0"; "barcodex.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.ocx"
Begin VB.Form etudiants 
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
      TabIndex        =   1
      Top             =   120
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   16325
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
      TabCaption(0)   =   "»ÕÀ ⁄«„"
      TabPicture(0)   =   "etudiants.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Skin1"
      Tab(0).Control(1)=   "Picture22"
      Tab(0).Control(2)=   "Picture21"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "ﬁ”„ „⁄Ì‰"
      TabPicture(1)   =   "etudiants.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture16"
      Tab(1).Control(1)=   "Picture15"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   " ·„Ì– „⁄Ì‰"
      TabPicture(2)   =   "etudiants.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Picture1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Picture2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Picture6"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.PictureBox Picture21 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8055
         Left            =   -74880
         ScaleHeight     =   8055
         ScaleWidth      =   14295
         TabIndex        =   208
         Top             =   1080
         Width           =   14295
         Begin MSFlexGridLib.MSFlexGrid grd10 
            Height          =   6735
            Left            =   6600
            TabIndex        =   209
            Top             =   600
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   11880
            _Version        =   393216
            Rows            =   1
            Cols            =   4
            FixedCols       =   0
            BackColor       =   0
            ForeColor       =   16777215
            BackColorFixed  =   0
            ForeColorFixed  =   16777215
            BackColorBkg    =   0
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
         Begin BARCODEXLib.BarcodeX BarcodeX3 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1036
               SubFormatType   =   0
            EndProperty
            Height          =   615
            Left            =   2040
            Top             =   240
            Width           =   2415
            _Version        =   65536
            _ExtentX        =   4260
            _ExtentY        =   1085
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
         Begin VB.Shape Shape9 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   3
            Height          =   1575
            Left            =   240
            Shape           =   4  'Rounded Rectangle
            Top             =   120
            Width           =   6255
         End
         Begin VB.Label Label113 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰ «∆Ã «·»ÕÀ"
            BeginProperty Font 
               Name            =   "Times New Roman"
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
            TabIndex        =   220
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label114 
            Alignment       =   2  'Center
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
            Left            =   9720
            TabIndex        =   219
            Top             =   120
            Width           =   3255
         End
         Begin VB.Image Image4 
            Appearance      =   0  'Flat
            Height          =   1335
            Left            =   240
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label119 
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
            Left            =   1920
            TabIndex        =   218
            Top             =   1320
            Width           =   3255
         End
         Begin VB.Label Label89 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Â«› «·ÊﬂÌ·"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   5280
            TabIndex        =   217
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label118 
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
            Left            =   2040
            TabIndex        =   216
            Top             =   960
            Width           =   3735
         End
         Begin VB.Label Label94 
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
            Left            =   5280
            TabIndex        =   215
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label117 
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
            Left            =   4560
            TabIndex        =   214
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label93 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·—ﬁ„"
            BeginProperty Font 
               Name            =   "Times New Roman"
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
            TabIndex        =   213
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label116 
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
            Left            =   4560
            TabIndex        =   212
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label1021 
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
            Left            =   5040
            TabIndex        =   211
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.PictureBox Picture22 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   -74880
         ScaleHeight     =   615
         ScaleWidth      =   14295
         TabIndex        =   203
         Top             =   360
         Width           =   14295
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
            Index           =   3
            Left            =   3120
            TabIndex        =   223
            Top             =   120
            Width           =   1695
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
            Index           =   2
            Left            =   6000
            TabIndex        =   222
            Top             =   120
            Width           =   1455
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
            Index           =   1
            Left            =   8880
            TabIndex        =   221
            Top             =   120
            Width           =   1095
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
            Index           =   0
            Left            =   11280
            TabIndex        =   205
            Top             =   120
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker DT4 
            Height          =   375
            Left            =   240
            TabIndex        =   206
            Top             =   120
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
            Format          =   124518401
            CurrentDate     =   41154
         End
         Begin VB.Label Label111 
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
            Index           =   4
            Left            =   12240
            TabIndex        =   252
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label111 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·—ﬁ„ ›Ì «·ﬁ”„"
            BeginProperty Font 
               Name            =   "Times New Roman"
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
            Left            =   9240
            TabIndex        =   250
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label111 
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
            Index           =   2
            Left            =   6840
            TabIndex        =   249
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label111 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «· ”ÃÌ·"
            BeginProperty Font 
               Name            =   "Times New Roman"
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
            Left            =   1080
            TabIndex        =   248
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label111 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Â« › «·ÊﬂÌ·"
            BeginProperty Font 
               Name            =   "Times New Roman"
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
            Left            =   3960
            TabIndex        =   204
            Top             =   120
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture16 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8055
         Left            =   -74880
         ScaleHeight     =   8055
         ScaleWidth      =   14295
         TabIndex        =   170
         Top             =   1080
         Width           =   14295
         Begin TabDlg.SSTab SSTab3 
            Height          =   7815
            Left            =   120
            TabIndex        =   171
            Top             =   120
            Width           =   14055
            _ExtentX        =   24791
            _ExtentY        =   13785
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
            TabCaption(0)   =   "‰ﬁ«ÿ  ·«„Ì– «·ﬁ”„"
            TabPicture(0)   =   "etudiants.frx":0054
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Picture20"
            Tab(0).Control(1)=   "Picture19"
            Tab(0).Control(2)=   "grd8"
            Tab(0).ControlCount=   3
            TabCaption(1)   =   "»Ì«‰«   ·«„Ì– «·ﬁ”„"
            TabPicture(1)   =   "etudiants.frx":0070
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "Picture17"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "Picture18"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).ControlCount=   2
            Begin VB.PictureBox Picture20 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   615
               Left            =   -74880
               ScaleHeight     =   615
               ScaleWidth      =   13815
               TabIndex        =   195
               Top             =   7080
               Width           =   13815
               Begin VB.CommandButton Command41 
                  Caption         =   "«·‰ «∆Ã Õ”» «·„⁄œ·« "
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
                  TabIndex        =   251
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1935
               End
               Begin VB.CommandButton Command34 
                  Caption         =   "ﬂ‘› «·œ—Ã«  ·Ã„Ì⁄  ·«„Ì– «·ﬁ”„"
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
                  TabIndex        =   198
                  Top             =   120
                  Width           =   3855
               End
               Begin VB.CheckBox Check9 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
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
                  Height          =   285
                  Left            =   11760
                  TabIndex        =   197
                  Top             =   120
                  Value           =   1  'Checked
                  Width           =   255
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
                  Left            =   12960
                  TabIndex        =   196
                  Top             =   120
                  Value           =   1  'Checked
                  Width           =   255
               End
               Begin VB.Label Label107 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·— »…"
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
                  Left            =   13200
                  TabIndex        =   200
                  Top             =   135
                  Width           =   495
               End
               Begin VB.Label Label105 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«· «—ÌŒ"
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
                  Left            =   12000
                  TabIndex        =   199
                  Top             =   135
                  Width           =   615
               End
            End
            Begin VB.PictureBox Picture19 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   615
               Left            =   -74880
               ScaleHeight     =   615
               ScaleWidth      =   13815
               TabIndex        =   188
               Top             =   360
               Width           =   13815
               Begin VB.CommandButton Command2 
                  Caption         =   "”Õ» ‰ «∆Ã «·„«œ…"
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
                  Left            =   2160
                  TabIndex        =   226
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
                  Left            =   3840
                  TabIndex        =   225
                  Top             =   120
                  Width           =   2655
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
                  Left            =   10800
                  Style           =   2  'Dropdown List
                  TabIndex        =   191
                  Top             =   120
                  Width           =   1695
               End
               Begin VB.CommandButton Command33 
                  Caption         =   "⁄—÷ «·‰ «∆Ã"
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
                  TabIndex        =   189
                  Top             =   120
                  Width           =   1695
               End
               Begin MSFlexGridLib.MSFlexGrid grd30 
                  Height          =   285
                  Left            =   7560
                  TabIndex        =   234
                  Top             =   120
                  Width           =   1365
                  _ExtentX        =   2408
                  _ExtentY        =   503
                  _Version        =   393216
                  Rows            =   1
                  Cols            =   1
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColor       =   0
                  ForeColor       =   16777215
                  BackColorFixed  =   0
                  ForeColorFixed  =   16777215
                  ForeColorSel    =   8388608
                  BackColorBkg    =   0
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
               Begin VB.Label Label90 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«”„ «·√” «–"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
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
                  Left            =   5520
                  TabIndex        =   224
                  Top             =   120
                  Width           =   1935
               End
               Begin VB.Shape Shape10 
                  BorderColor     =   &H00FFFFFF&
                  BorderWidth     =   2
                  Height          =   375
                  Left            =   120
                  Shape           =   4  'Rounded Rectangle
                  Top             =   120
                  Width           =   13575
               End
               Begin VB.Label Label103 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„«œ…"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
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
                  TabIndex        =   190
                  Top             =   120
                  Width           =   1215
               End
            End
            Begin VB.PictureBox Picture18 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   7335
               Left            =   120
               ScaleHeight     =   7335
               ScaleWidth      =   13815
               TabIndex        =   185
               Top             =   360
               Width           =   13815
            End
            Begin VB.PictureBox Picture17 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   7335
               Left            =   120
               ScaleHeight     =   7335
               ScaleWidth      =   13815
               TabIndex        =   172
               Top             =   360
               Width           =   13815
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
                  Left            =   6480
                  MaskColor       =   &H00000000&
                  TabIndex        =   186
                  Top             =   1080
                  Width           =   255
               End
               Begin VB.CommandButton Command30 
                  Caption         =   "„”Õ «·ﬁ«∆„…"
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
                  Left            =   360
                  TabIndex        =   182
                  Top             =   6840
                  Width           =   2175
               End
               Begin VB.CommandButton Command29 
                  Caption         =   "”Õ» »ÿ«ﬁ«  «·œŒÊ·"
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
                  Left            =   2760
                  TabIndex        =   181
                  Top             =   6840
                  Width           =   3735
               End
               Begin VB.CommandButton Command28 
                  Caption         =   "≈÷«›… Â–« «· ·„Ì– ≈·Ï ﬁ«∆„… »ÿ«ﬁ«  «·œŒÊ·"
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
                  TabIndex        =   179
                  Top             =   1560
                  Width           =   4575
               End
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
                  Left            =   9000
                  TabIndex        =   176
                  Top             =   6840
                  Visible         =   0   'False
                  Width           =   3735
               End
               Begin MSFlexGridLib.MSFlexGrid grd6 
                  Height          =   6615
                  Left            =   7080
                  TabIndex        =   175
                  Top             =   120
                  Width           =   6615
                  _ExtentX        =   11668
                  _ExtentY        =   11668
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
               Begin MSFlexGridLib.MSFlexGrid grd7 
                  Height          =   3975
                  Left            =   120
                  TabIndex        =   180
                  Top             =   2280
                  Width           =   6735
                  _ExtentX        =   11880
                  _ExtentY        =   7011
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
               Begin VB.Label Label115 
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
                  Left            =   2280
                  TabIndex        =   210
                  Top             =   240
                  Width           =   2175
               End
               Begin VB.Label Label101 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "≈÷«›…  ·ﬁ«∆Ì…"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
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
                  TabIndex        =   187
                  Top             =   1080
                  Width           =   1575
               End
               Begin VB.Image Image3 
                  Appearance      =   0  'Flat
                  Height          =   1695
                  Left            =   240
                  Stretch         =   -1  'True
                  Top             =   240
                  Width           =   1695
               End
               Begin VB.Shape Shape8 
                  BorderColor     =   &H00FFFFFF&
                  BorderWidth     =   3
                  Height          =   1935
                  Left            =   120
                  Shape           =   4  'Rounded Rectangle
                  Top             =   120
                  Width           =   6855
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
                  Left            =   2280
                  TabIndex        =   178
                  Top             =   600
                  Width           =   3975
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
                  TabIndex        =   177
                  Top             =   240
                  Width           =   1575
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
                  TabIndex        =   174
                  Top             =   600
                  Width           =   1935
               End
               Begin VB.Label Label91 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·—ﬁ„"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
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
                  TabIndex        =   173
                  Top             =   240
                  Width           =   1095
               End
            End
            Begin MSFlexGridLib.MSFlexGrid grd8 
               Height          =   5895
               Left            =   -74880
               TabIndex        =   192
               Top             =   1080
               Width           =   13815
               _ExtentX        =   24368
               _ExtentY        =   10398
               _Version        =   393216
               FixedRows       =   0
               FixedCols       =   0
               BackColor       =   0
               ForeColor       =   16777215
               BackColorFixed  =   0
               ForeColorFixed  =   16777215
               ForeColorSel    =   8388608
               BackColorBkg    =   0
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
         End
      End
      Begin VB.PictureBox Picture15 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   -74880
         ScaleHeight     =   615
         ScaleWidth      =   14295
         TabIndex        =   167
         Top             =   360
         Width           =   14295
         Begin VB.CommandButton Command40 
            Caption         =   "”Õ» ‰ «∆Ã «·»«ﬂ·Ê—Ì«"
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
            TabIndex        =   258
            Top             =   120
            Visible         =   0   'False
            Width           =   1935
         End
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
            Left            =   7200
            Style           =   2  'Dropdown List
            TabIndex        =   168
            Top             =   120
            Width           =   1695
         End
         Begin MSComctlLib.ProgressBar ProgressBar5 
            Height          =   375
            Left            =   2280
            TabIndex        =   239
            Top             =   120
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label Label90 
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
            Left            =   7560
            TabIndex        =   169
            Top             =   120
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8055
         Left            =   120
         ScaleHeight     =   8055
         ScaleWidth      =   14295
         TabIndex        =   21
         Top             =   1080
         Width           =   14295
         Begin TabDlg.SSTab SSTab2 
            Height          =   7815
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Visible         =   0   'False
            Width           =   14055
            _ExtentX        =   24791
            _ExtentY        =   13785
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
            TabCaption(0)   =   "€Ì«»«  «· ·„Ì–"
            TabPicture(0)   =   "etudiants.frx":008C
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Picture13"
            Tab(0).Control(1)=   "Picture14"
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "‰ﬁ«ÿ «· ·„Ì–"
            TabPicture(1)   =   "etudiants.frx":00A8
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Line1"
            Tab(1).Control(1)=   "Line2"
            Tab(1).Control(2)=   "grd1"
            Tab(1).Control(3)=   "Picture8"
            Tab(1).Control(4)=   "Picture9"
            Tab(1).Control(5)=   "Picture10"
            Tab(1).Control(6)=   "Picture11"
            Tab(1).ControlCount=   7
            TabCaption(2)   =   "»Ì«‰«  «· ·„Ì–"
            TabPicture(2)   =   "etudiants.frx":00C4
            Tab(2).ControlEnabled=   -1  'True
            Tab(2).Control(0)=   "Picture3"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).ControlCount=   1
            Begin VB.PictureBox Picture14 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   6615
               Left            =   -74880
               ScaleHeight     =   6615
               ScaleWidth      =   13815
               TabIndex        =   147
               Top             =   1080
               Width           =   13815
               Begin VB.PictureBox Picture12 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   6615
                  Left            =   0
                  ScaleHeight     =   6615
                  ScaleWidth      =   13815
                  TabIndex        =   148
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   13815
                  Begin VB.CommandButton Command45 
                     Caption         =   "”Õ» ”Ã· «·€Ì«»"
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
                     TabIndex        =   275
                     Top             =   720
                     Width           =   3735
                  End
                  Begin VB.CommandButton Command25 
                     Caption         =   "Õ–› «·»Ì«‰« "
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
                     TabIndex        =   164
                     Top             =   720
                     Width           =   1335
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
                     Left            =   9000
                     Style           =   2  'Dropdown List
                     TabIndex        =   154
                     Top             =   120
                     Width           =   855
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
                     Left            =   6960
                     Style           =   2  'Dropdown List
                     TabIndex        =   153
                     Top             =   120
                     Width           =   855
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
                     Left            =   4680
                     Style           =   2  'Dropdown List
                     TabIndex        =   152
                     Top             =   120
                     Width           =   1335
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
                     Left            =   120
                     TabIndex        =   151
                     Top             =   120
                     Width           =   3615
                  End
                  Begin VB.CommandButton Command22 
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
                     Left            =   2520
                     TabIndex        =   150
                     Top             =   720
                     Width           =   2055
                  End
                  Begin VB.CommandButton Command23 
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
                     TabIndex        =   149
                     Top             =   720
                     Width           =   855
                  End
                  Begin MSComCtl2.DTPicker DT3 
                     Height          =   375
                     Left            =   11040
                     TabIndex        =   155
                     Top             =   120
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
                     Format          =   124518401
                     CurrentDate     =   41154
                  End
                  Begin MSComctlLib.ProgressBar ProgressBar4 
                     Height          =   375
                     Left            =   4680
                     TabIndex        =   156
                     Top             =   720
                     Width           =   5175
                     _ExtentX        =   9128
                     _ExtentY        =   661
                     _Version        =   393216
                     Appearance      =   1
                  End
                  Begin MSFlexGridLib.MSFlexGrid grd2 
                     Height          =   5175
                     Left            =   120
                     TabIndex        =   157
                     Top             =   1320
                     Width           =   13575
                     _ExtentX        =   23945
                     _ExtentY        =   9128
                     _Version        =   393216
                     FixedRows       =   0
                     FixedCols       =   0
                     BackColor       =   0
                     ForeColor       =   16777215
                     BackColorFixed  =   0
                     ForeColorFixed  =   16777215
                     ForeColorSel    =   8388608
                     BackColorBkg    =   0
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
                  Begin VB.Label Label82 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   " «—ÌŒ «· €Ì»"
                     BeginProperty Font 
                        Name            =   "Times New Roman"
                        Size            =   12
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   375
                     Left            =   12360
                     TabIndex        =   162
                     Top             =   120
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
                     Left            =   9480
                     TabIndex        =   161
                     Top             =   120
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
                     Left            =   7560
                     TabIndex        =   160
                     Top             =   120
                     Width           =   1335
                  End
                  Begin VB.Label Label85 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·„«œ…"
                     BeginProperty Font 
                        Name            =   "Times New Roman"
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
                     TabIndex        =   159
                     Top             =   120
                     Width           =   1335
                  End
                  Begin VB.Label Label86 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "„·«ÕŸ…"
                     BeginProperty Font 
                        Name            =   "Times New Roman"
                        Size            =   12
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   375
                     Left            =   3120
                     TabIndex        =   158
                     Top             =   120
                     Width           =   1335
                  End
               End
            End
            Begin VB.PictureBox Picture13 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   615
               Left            =   -74880
               ScaleHeight     =   615
               ScaleWidth      =   13815
               TabIndex        =   137
               Top             =   360
               Width           =   13815
               Begin VB.CommandButton Command24 
                  Caption         =   "⁄—÷ ”Ã· «·€Ì«»"
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
                  TabIndex        =   146
                  Top             =   120
                  Width           =   2895
               End
               Begin VB.Label Label81 
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
                  Left            =   11640
                  TabIndex        =   145
                  Top             =   120
                  Width           =   1455
               End
               Begin VB.Label Label80 
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
                  TabIndex        =   144
                  Top             =   120
                  Width           =   1335
               End
               Begin VB.Label Label79 
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
                  Left            =   12600
                  TabIndex        =   143
                  Top             =   120
                  Width           =   975
               End
               Begin VB.Label Label78 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·—ﬁ„"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
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
                  TabIndex        =   142
                  Top             =   120
                  Width           =   855
               End
               Begin VB.Label Label77 
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
                  TabIndex        =   141
                  Top             =   120
                  Width           =   1935
               End
               Begin VB.Label Label76 
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
                  Left            =   5040
                  TabIndex        =   140
                  Top             =   120
                  Width           =   3855
               End
               Begin VB.Shape Shape7 
                  BorderColor     =   &H00FFFFFF&
                  BorderWidth     =   2
                  Height          =   375
                  Left            =   120
                  Shape           =   4  'Rounded Rectangle
                  Top             =   120
                  Width           =   13575
               End
               Begin VB.Label Label75 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·€Ì«»"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
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
                  TabIndex        =   139
                  Top             =   120
                  Width           =   1215
               End
               Begin VB.Label Label74 
                  Alignment       =   2  'Center
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
                  Left            =   3120
                  TabIndex        =   138
                  Top             =   120
                  Width           =   1095
               End
            End
            Begin VB.PictureBox Picture11 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   6615
               Left            =   -74880
               ScaleHeight     =   6615
               ScaleWidth      =   5415
               TabIndex        =   71
               Top             =   1080
               Visible         =   0   'False
               Width           =   5415
               Begin VB.TextBox mens 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   0
                  Left            =   2640
                  TabIndex        =   102
                  Top             =   1440
                  Width           =   1815
               End
               Begin VB.TextBox coff 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   0
                  Left            =   2640
                  TabIndex        =   101
                  Top             =   360
                  Width           =   495
               End
               Begin VB.TextBox coff 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   1
                  Left            =   480
                  TabIndex        =   100
                  Top             =   360
                  Width           =   495
               End
               Begin VB.TextBox coff 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   2
                  Left            =   2640
                  TabIndex        =   99
                  Top             =   720
                  Width           =   495
               End
               Begin VB.TextBox coff 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   3
                  Left            =   480
                  TabIndex        =   98
                  Top             =   720
                  Width           =   495
               End
               Begin VB.TextBox coff 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   4
                  Left            =   2640
                  TabIndex        =   97
                  Top             =   1080
                  Width           =   1815
               End
               Begin VB.TextBox coff 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   5
                  Left            =   480
                  TabIndex        =   96
                  Top             =   1080
                  Width           =   1815
               End
               Begin VB.TextBox coff 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   6
                  Left            =   2640
                  TabIndex        =   95
                  Top             =   1800
                  Width           =   1815
               End
               Begin VB.TextBox coff 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   7
                  Left            =   480
                  TabIndex        =   94
                  Top             =   1800
                  Width           =   1815
               End
               Begin VB.TextBox coff 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   8
                  Left            =   2640
                  TabIndex        =   93
                  Top             =   2520
                  Width           =   1815
               End
               Begin VB.TextBox coff 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   9
                  Left            =   480
                  TabIndex        =   92
                  Top             =   2520
                  Width           =   1815
               End
               Begin VB.TextBox coff 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   10
                  Left            =   2640
                  TabIndex        =   91
                  Top             =   3240
                  Width           =   1815
               End
               Begin VB.TextBox coff 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   11
                  Left            =   480
                  TabIndex        =   90
                  Top             =   3240
                  Width           =   1815
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
                  Left            =   2640
                  TabIndex        =   89
                  Top             =   5520
                  Width           =   2415
               End
               Begin VB.TextBox coff 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   12
                  Left            =   2640
                  TabIndex        =   88
                  Top             =   3960
                  Width           =   1815
               End
               Begin VB.TextBox coff 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   13
                  Left            =   480
                  TabIndex        =   87
                  Top             =   3960
                  Width           =   1815
               End
               Begin VB.TextBox coff 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   14
                  Left            =   2640
                  TabIndex        =   86
                  Top             =   4680
                  Width           =   1815
               End
               Begin VB.TextBox coff 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   15
                  Left            =   480
                  TabIndex        =   85
                  Top             =   4680
                  Width           =   1815
               End
               Begin VB.TextBox mens 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   1
                  Left            =   2640
                  TabIndex        =   84
                  Top             =   2160
                  Width           =   1815
               End
               Begin VB.TextBox mens 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   2
                  Left            =   2640
                  TabIndex        =   83
                  Top             =   2880
                  Width           =   1815
               End
               Begin VB.TextBox mens 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   3
                  Left            =   2640
                  TabIndex        =   82
                  Top             =   3600
                  Width           =   1815
               End
               Begin VB.TextBox mens 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   4
                  Left            =   2640
                  TabIndex        =   81
                  Top             =   4320
                  Width           =   1815
               End
               Begin VB.TextBox mens 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   5
                  Left            =   2640
                  TabIndex        =   80
                  Top             =   5040
                  Width           =   1815
               End
               Begin VB.TextBox mens 
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   6
                  Left            =   480
                  TabIndex        =   79
                  Top             =   1440
                  Width           =   1815
               End
               Begin VB.TextBox mens 
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   7
                  Left            =   480
                  TabIndex        =   78
                  Top             =   2160
                  Width           =   1815
               End
               Begin VB.TextBox mens 
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   8
                  Left            =   480
                  TabIndex        =   77
                  Top             =   2880
                  Width           =   1815
               End
               Begin VB.TextBox mens 
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   9
                  Left            =   480
                  TabIndex        =   76
                  Top             =   3600
                  Width           =   1815
               End
               Begin VB.TextBox mens 
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   10
                  Left            =   480
                  TabIndex        =   75
                  Top             =   4320
                  Width           =   1815
               End
               Begin VB.TextBox mens 
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   11
                  Left            =   480
                  TabIndex        =   74
                  Top             =   5040
                  Width           =   1815
               End
               Begin VB.CommandButton Command19 
                  Caption         =   "ÃœÊ·  ·ﬁ«∆Ì"
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
                  Left            =   1320
                  TabIndex        =   73
                  Top             =   5520
                  Width           =   1215
               End
               Begin VB.CommandButton Command18 
                  Caption         =   "≈Œ›«¡"
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
                  TabIndex        =   72
                  Top             =   5520
                  Width           =   1095
               End
               Begin MSFlexGridLib.MSFlexGrid grd22 
                  Height          =   5895
                  Left            =   5640
                  TabIndex        =   240
                  Top             =   120
                  Width           =   4335
                  _ExtentX        =   7646
                  _ExtentY        =   10398
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
               Begin MSFlexGridLib.MSFlexGrid grd23 
                  Height          =   615
                  Left            =   120
                  TabIndex        =   241
                  Top             =   6000
                  Width           =   4935
                  _ExtentX        =   8705
                  _ExtentY        =   1085
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
               Begin VB.Label Label55 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«· ﬁœÌ—"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Index           =   5
                  Left            =   2640
                  TabIndex        =   247
                  Top             =   5040
                  Width           =   2295
               End
               Begin VB.Label Label55 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«· ﬁœÌ—"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Index           =   4
                  Left            =   2640
                  TabIndex        =   246
                  Top             =   4320
                  Width           =   2295
               End
               Begin VB.Label Label55 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«· ﬁœÌ—"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Index           =   3
                  Left            =   2640
                  TabIndex        =   245
                  Top             =   3600
                  Width           =   2295
               End
               Begin VB.Label Label55 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«· ﬁœÌ—"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Index           =   2
                  Left            =   2640
                  TabIndex        =   244
                  Top             =   2880
                  Width           =   2295
               End
               Begin VB.Label Label55 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«· ﬁœÌ—"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Index           =   0
                  Left            =   2640
                  TabIndex        =   243
                  Top             =   2160
                  Width           =   2295
               End
               Begin VB.Label Label55 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«· ﬁœÌ—"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Index           =   1
                  Left            =   2640
                  TabIndex        =   242
                  Top             =   1440
                  Width           =   2295
               End
               Begin VB.Label Label66 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "÷«—» „⁄œ· «·≈Œ »«—« "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   3120
                  TabIndex        =   119
                  Top             =   360
                  Width           =   1815
               End
               Begin VB.Label Label65 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "÷«—» „⁄œ· «„ Õ«‰ 1"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   600
                  TabIndex        =   118
                  Top             =   360
                  Width           =   1935
               End
               Begin VB.Label Label62 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "÷«—» „⁄œ· «„ Õ«‰ 2"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   2880
                  TabIndex        =   117
                  Top             =   720
                  Width           =   2055
               End
               Begin VB.Label Label61 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "÷«—» „⁄œ· «„ Õ«‰ 3"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   600
                  TabIndex        =   116
                  Top             =   720
                  Width           =   1935
               End
               Begin VB.Label Label60 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‰"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   4440
                  TabIndex        =   115
                  Top             =   1080
                  Width           =   495
               End
               Begin VB.Label Label59 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "≈·Ï"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   2040
                  TabIndex        =   114
                  Top             =   1080
                  Width           =   855
               End
               Begin VB.Label Label57 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‰"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   4440
                  TabIndex        =   113
                  Top             =   1800
                  Width           =   495
               End
               Begin VB.Label Label56 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "≈·Ï"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   2040
                  TabIndex        =   112
                  Top             =   1800
                  Width           =   855
               End
               Begin VB.Label Label54 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‰"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   4440
                  TabIndex        =   111
                  Top             =   3240
                  Width           =   495
               End
               Begin VB.Label Label53 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "≈·Ï"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   2040
                  TabIndex        =   110
                  Top             =   2520
                  Width           =   855
               End
               Begin VB.Label Label51 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‰"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   4440
                  TabIndex        =   109
                  Top             =   2520
                  Width           =   495
               End
               Begin VB.Label Label50 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "≈·Ï"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   2040
                  TabIndex        =   108
                  Top             =   3240
                  Width           =   855
               End
               Begin VB.Label Label48 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‰"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   4440
                  TabIndex        =   107
                  Top             =   4680
                  Width           =   495
               End
               Begin VB.Label Label47 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "≈·Ï"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   2040
                  TabIndex        =   106
                  Top             =   3960
                  Width           =   855
               End
               Begin VB.Label Label43 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‰"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   4440
                  TabIndex        =   105
                  Top             =   3960
                  Width           =   495
               End
               Begin VB.Label Label44 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "≈·Ï"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   2040
                  TabIndex        =   104
                  Top             =   4680
                  Width           =   855
               End
               Begin VB.Label Label64 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "ÃœÊ· «·÷Ê«—» Ê«· ﬁœÌ—« "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   480
                  TabIndex        =   103
                  Top             =   0
                  Width           =   4455
               End
            End
            Begin VB.PictureBox Picture10 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   615
               Left            =   -74880
               ScaleHeight     =   615
               ScaleWidth      =   13815
               TabIndex        =   65
               Top             =   5280
               Width           =   13815
               Begin VB.CommandButton Command42 
                  Caption         =   "‰ «∆Ã «· ·„Ì–"
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
                  TabIndex        =   254
                  Top             =   120
                  Width           =   2175
               End
               Begin VB.CheckBox Check3 
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
                  Left            =   4560
                  TabIndex        =   129
                  Top             =   140
                  Value           =   1  'Checked
                  Width           =   255
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
                  Left            =   5400
                  MaskColor       =   &H00000000&
                  TabIndex        =   128
                  Top             =   140
                  Value           =   1  'Checked
                  Width           =   255
               End
               Begin VB.CheckBox Check5 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
                  Caption         =   "«·€Ì«»"
                  BeginProperty Font 
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
                  Left            =   2760
                  TabIndex        =   127
                  Top             =   120
                  Value           =   1  'Checked
                  Width           =   255
               End
               Begin VB.CheckBox Check4 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
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
                  Height          =   285
                  Left            =   3600
                  TabIndex        =   126
                  Top             =   140
                  Value           =   1  'Checked
                  Width           =   255
               End
               Begin VB.CommandButton Command21 
                  Caption         =   "«·— »…"
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
                  Left            =   7800
                  TabIndex        =   121
                  Top             =   120
                  Width           =   735
               End
               Begin VB.Label Label71 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·€Ì«»"
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
                  Left            =   3000
                  TabIndex        =   134
                  Top             =   135
                  Width           =   495
               End
               Begin VB.Label Label70 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«· «—ÌŒ"
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
                  TabIndex        =   133
                  Top             =   135
                  Width           =   615
               End
               Begin VB.Label Label69 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«· ﬁœÌ—"
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
                  Left            =   4800
                  TabIndex        =   132
                  Top             =   135
                  Width           =   495
               End
               Begin VB.Label Label68 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·— »…"
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
                  Left            =   5760
                  TabIndex        =   131
                  Top             =   140
                  Width           =   495
               End
               Begin VB.Shape Shape6 
                  BorderColor     =   &H00FFFFFF&
                  Height          =   375
                  Left            =   2640
                  Top             =   120
                  Width           =   3855
               End
               Begin VB.Label Label29 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Label29"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
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
                  TabIndex        =   120
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   2535
               End
               Begin VB.Label Label42 
                  Alignment       =   2  'Center
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
                  Left            =   6840
                  TabIndex        =   70
                  Top             =   120
                  Width           =   855
               End
               Begin VB.Label Label40 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "10.25"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
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
                  TabIndex        =   69
                  Top             =   120
                  Width           =   2535
               End
               Begin VB.Label Label39 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
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
                  Height          =   375
                  Left            =   10680
                  TabIndex        =   68
                  Top             =   120
                  Width           =   1215
               End
               Begin VB.Label Label33 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "10.25"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
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
                  TabIndex        =   67
                  Top             =   120
                  Width           =   735
               End
               Begin VB.Label Label32 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„⁄œ· «·⁄«„"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
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
                  TabIndex        =   66
                  Top             =   120
                  Width           =   1935
               End
            End
            Begin VB.PictureBox Picture9 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   615
               Left            =   -74880
               ScaleHeight     =   615
               ScaleWidth      =   13815
               TabIndex        =   57
               Top             =   360
               Width           =   13815
               Begin VB.TextBox Text17 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00000000&
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
                  ForeColor       =   &H00FFFFFF&
                  Height          =   375
                  Left            =   5520
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   273
                  Top             =   120
                  Width           =   1455
               End
               Begin VB.CommandButton Command17 
                  Caption         =   "ÃœÊ· «· ﬁœÌ—« "
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
                  TabIndex        =   130
                  Top             =   120
                  Width           =   1455
               End
               Begin VB.CommandButton Command13 
                  Caption         =   "⁄—÷ «·‰ «∆Ã"
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
                  TabIndex        =   64
                  Top             =   120
                  Width           =   1335
               End
               Begin VB.Label Label15 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "RIM"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
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
                  Left            =   4920
                  TabIndex        =   272
                  Top             =   150
                  Width           =   615
               End
               Begin VB.Label Label73 
                  Alignment       =   2  'Center
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
                  Left            =   3120
                  TabIndex        =   136
                  Top             =   120
                  Width           =   1095
               End
               Begin VB.Label Label72 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·€Ì«»"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
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
                  TabIndex        =   135
                  Top             =   120
                  Width           =   1215
               End
               Begin VB.Shape Shape5 
                  BorderColor     =   &H00FFFFFF&
                  BorderWidth     =   2
                  Height          =   375
                  Left            =   120
                  Shape           =   4  'Rounded Rectangle
                  Top             =   120
                  Width           =   13575
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
                  Left            =   6960
                  TabIndex        =   63
                  Top             =   120
                  Width           =   2775
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
                  Left            =   9600
                  TabIndex        =   62
                  Top             =   120
                  Width           =   735
               End
               Begin VB.Label Label34 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·—ﬁ„"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   375
                  Left            =   11040
                  TabIndex        =   61
                  Top             =   120
                  Width           =   495
               End
               Begin VB.Label Label35 
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
                  Left            =   12600
                  TabIndex        =   60
                  Top             =   120
                  Width           =   975
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
                  Left            =   10320
                  TabIndex        =   59
                  Top             =   120
                  Width           =   735
               End
               Begin VB.Label Label36 
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
                  Left            =   11640
                  TabIndex        =   58
                  Top             =   120
                  Width           =   1455
               End
            End
            Begin VB.PictureBox Picture8 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   1695
               Left            =   -74880
               ScaleHeight     =   1695
               ScaleWidth      =   13815
               TabIndex        =   44
               Top             =   6000
               Width           =   13815
               Begin VB.CommandButton Command16 
                  Caption         =   "ﬂ‘› «·œ—Ã«  «·⁄«„"
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
                  TabIndex        =   253
                  Top             =   120
                  Width           =   2175
               End
               Begin VB.Label Label27 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "3"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
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
                  TabIndex        =   53
                  Top             =   0
                  Width           =   735
               End
               Begin VB.Label Label26 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "2"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
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
                  TabIndex        =   52
                  Top             =   0
                  Width           =   735
               End
               Begin VB.Label Label25 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "1"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   375
                  Left            =   10320
                  TabIndex        =   51
                  Top             =   0
                  Width           =   735
               End
               Begin VB.Label Label24 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "3"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   375
                  Left            =   11880
                  TabIndex        =   50
                  Top             =   0
                  Width           =   735
               End
               Begin VB.Label Label63 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   $"etudiants.frx":00E0
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
                  Index           =   5
                  Left            =   0
                  TabIndex        =   49
                  Top             =   1200
                  Width           =   12975
               End
               Begin VB.Label Label63 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "÷«—» «·«Œ »«—«       ÷«—» «„ Õ«‰ 1     ÷«—» «„ Õ«‰ 2      ÷«—» «„ Õ«‰ 3"
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
                  Left            =   7200
                  TabIndex        =   48
                  Top             =   120
                  Width           =   6495
               End
               Begin VB.Label Label63 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„⁄œ· «·«Œ »«—«   = „Ã„Ê⁄ «·«Œ »«—«  / ⁄œœ «·«Œ »«—«  („«  „ Õ”»«‰Â)‹ "
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
                  Left            =   5280
                  TabIndex        =   47
                  Top             =   480
                  Width           =   7695
               End
               Begin VB.Label Label63 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   $"etudiants.frx":01DC
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
                  Left            =   -600
                  TabIndex        =   46
                  Top             =   720
                  Width           =   13575
               End
               Begin VB.Label Label63 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„Ã„Ê⁄ = „⁄œ· «·„«œ… * ÷«—» «·„«œ…       «·„⁄œ· «·⁄«„ = „Ã„Ê⁄ „Ã«„Ì⁄ ﬂ· «·„Ê«œ / „Ã„Ê⁄ ÷Ê«—» ﬂ· «·„Ê«œ "
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
                  Left            =   4560
                  TabIndex        =   45
                  Top             =   960
                  Width           =   8415
               End
            End
            Begin VB.PictureBox Picture3 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   7335
               Left            =   120
               ScaleHeight     =   7335
               ScaleWidth      =   13815
               TabIndex        =   24
               Top             =   360
               Width           =   13815
               Begin VB.PictureBox Picture5 
                  Height          =   6255
                  Left            =   600
                  ScaleHeight     =   6195
                  ScaleWidth      =   12435
                  TabIndex        =   38
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   12495
                  Begin VB.CommandButton Command39 
                     Caption         =   "ﬂ‘› «·œ—Ã«  ··›’· «·À«‰Ì"
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
                     Left            =   6840
                     TabIndex        =   257
                     Top             =   5160
                     Width           =   2655
                  End
                  Begin VB.CommandButton Command37 
                     Caption         =   "ﬂ‘› «·œ—Ã«  ··›’· «·À«‰Ì"
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
                     TabIndex        =   256
                     Top             =   5400
                     Width           =   2415
                  End
                  Begin VB.CommandButton Command35 
                     Caption         =   "ﬂ‘› «·œ—Ã«  ··›’· «·√Ê·"
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
                     TabIndex        =   255
                     Top             =   5400
                     Width           =   2415
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
                     Left            =   1080
                     Style           =   2  'Dropdown List
                     TabIndex        =   235
                     Top             =   4200
                     Width           =   1575
                  End
                  Begin VB.CommandButton Command36 
                     Caption         =   "Command36"
                     Height          =   495
                     Left            =   240
                     TabIndex        =   207
                     Top             =   1800
                     Width           =   1335
                  End
                  Begin VB.CommandButton Command32 
                     Caption         =   " Ê“Ì⁄ «·— »"
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
                     TabIndex        =   202
                     Top             =   3600
                     Width           =   1935
                  End
                  Begin MSFlexGridLib.MSFlexGrid grd9 
                     Height          =   4575
                     Left            =   3840
                     TabIndex        =   194
                     Top             =   600
                     Width           =   2895
                     _ExtentX        =   5106
                     _ExtentY        =   8070
                     _Version        =   393216
                  End
                  Begin VB.Timer Timer8 
                     Enabled         =   0   'False
                     Interval        =   100
                     Left            =   2040
                     Top             =   3120
                  End
                  Begin VB.Timer Timer7 
                     Enabled         =   0   'False
                     Interval        =   50
                     Left            =   2040
                     Top             =   2280
                  End
                  Begin VB.Timer Timer6 
                     Enabled         =   0   'False
                     Interval        =   50
                     Left            =   2040
                     Top             =   1800
                  End
                  Begin MSFlexGridLib.MSFlexGrid grd5 
                     Height          =   4815
                     Left            =   7080
                     TabIndex        =   123
                     Top             =   240
                     Width           =   5295
                     _ExtentX        =   9340
                     _ExtentY        =   8493
                     _Version        =   393216
                     Cols            =   5
                  End
                  Begin VB.CommandButton Command8 
                     Caption         =   "Command8"
                     Height          =   375
                     Left            =   840
                     TabIndex        =   56
                     Top             =   840
                     Width           =   1695
                  End
                  Begin VB.Timer Timer4 
                     Enabled         =   0   'False
                     Interval        =   50
                     Left            =   1920
                     Top             =   120
                  End
                  Begin VB.Timer Timer3 
                     Enabled         =   0   'False
                     Interval        =   50
                     Left            =   1320
                     Top             =   120
                  End
                  Begin VB.Timer Timer1 
                     Enabled         =   0   'False
                     Interval        =   50
                     Left            =   720
                     Top             =   120
                  End
                  Begin VB.Timer Timer2 
                     Enabled         =   0   'False
                     Interval        =   50
                     Left            =   120
                     Top             =   120
                  End
                  Begin MSComctlLib.ProgressBar ProgressBar3 
                     Height          =   375
                     Left            =   0
                     TabIndex        =   122
                     Top             =   2760
                     Width           =   3735
                     _ExtentX        =   6588
                     _ExtentY        =   661
                     _Version        =   393216
                     Appearance      =   1
                  End
                  Begin VB.Label Label4 
                     Alignment       =   2  'Center
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
                     ForeColor       =   &H000000FF&
                     Height          =   375
                     Left            =   2400
                     TabIndex        =   271
                     Top             =   720
                     Visible         =   0   'False
                     Width           =   1095
                  End
                  Begin VB.Label Label5 
                     Alignment       =   2  'Center
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
                     ForeColor       =   &H000000FF&
                     Height          =   375
                     Left            =   2400
                     TabIndex        =   270
                     Top             =   120
                     Visible         =   0   'False
                     Width           =   1095
                  End
                  Begin VB.Label Label7 
                     Alignment       =   2  'Center
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
                     ForeColor       =   &H000000FF&
                     Height          =   375
                     Left            =   2400
                     TabIndex        =   269
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   1095
                  End
                  Begin VB.Label Label106 
                     Caption         =   "Label106"
                     Height          =   255
                     Left            =   240
                     TabIndex        =   201
                     Top             =   4560
                     Width           =   1575
                  End
                  Begin VB.Label Label104 
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
                     TabIndex        =   193
                     Top             =   4560
                     Width           =   1215
                  End
                  Begin VB.Label Label100 
                     Caption         =   "Label100"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   183
                     Top             =   4200
                     Width           =   1575
                  End
                  Begin VB.Label Label88 
                     Height          =   255
                     Left            =   120
                     TabIndex        =   166
                     Top             =   4080
                     Width           =   2295
                  End
                  Begin VB.Label Label87 
                     Height          =   255
                     Left            =   120
                     TabIndex        =   163
                     Top             =   3720
                     Width           =   1455
                  End
                  Begin VB.Label Label67 
                     Caption         =   "Label67"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   125
                     Top             =   3480
                     Width           =   2295
                  End
                  Begin VB.Label Label41 
                     Caption         =   "Label41"
                     Height          =   375
                     Left            =   120
                     TabIndex        =   124
                     Top             =   3240
                     Width           =   1095
                  End
                  Begin VB.Label Label21 
                     Caption         =   "Label21"
                     Height          =   255
                     Left            =   0
                     TabIndex        =   40
                     Top             =   1200
                     Width           =   1815
                  End
                  Begin VB.Label Label11 
                     Caption         =   "Label11"
                     Height          =   375
                     Left            =   0
                     TabIndex        =   39
                     Top             =   960
                     Width           =   1215
                  End
               End
               Begin VB.PictureBox Picture7 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   4695
                  Left            =   3000
                  ScaleHeight     =   4695
                  ScaleWidth      =   7335
                  TabIndex        =   25
                  Top             =   960
                  Width           =   7335
                  Begin VB.TextBox Text21 
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
                     Left            =   4080
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   282
                     Top             =   2160
                     Width           =   1935
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
                     Left            =   240
                     TabIndex        =   276
                     Top             =   1200
                     Width           =   2775
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
                     Left            =   4440
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   267
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
                     Left            =   240
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   260
                     Top             =   240
                     Width           =   3015
                  End
                  Begin VB.CommandButton Command26 
                     Caption         =   "Õ–› «· ·„Ì–"
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
                     TabIndex        =   165
                     Top             =   3960
                     Width           =   1695
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
                     Left            =   4080
                     Style           =   2  'Dropdown List
                     TabIndex        =   41
                     Top             =   1680
                     Width           =   1935
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
                     Left            =   2160
                     Style           =   2  'Dropdown List
                     TabIndex        =   33
                     Top             =   720
                     Width           =   1575
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
                     Left            =   240
                     TabIndex        =   32
                     Top             =   720
                     Width           =   1215
                  End
                  Begin VB.CommandButton Command9 
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
                     TabIndex        =   31
                     Top             =   3960
                     Width           =   3015
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
                     Left            =   3240
                     TabIndex        =   30
                     Top             =   1200
                     Width           =   2775
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
                     TabIndex        =   29
                     Top             =   1680
                     Width           =   2535
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
                     Left            =   4440
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   28
                     Top             =   240
                     Width           =   1575
                  End
                  Begin VB.CommandButton Command11 
                     Caption         =   "«—›«ﬁ"
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
                     Left            =   6120
                     TabIndex        =   27
                     Top             =   2640
                     Width           =   975
                  End
                  Begin VB.CommandButton Command12 
                     Caption         =   "„”Õ"
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
                     Left            =   6120
                     TabIndex        =   26
                     Top             =   3240
                     Width           =   975
                  End
                  Begin MSComCtl2.DTPicker DT2 
                     Height          =   375
                     Left            =   240
                     TabIndex        =   34
                     Top             =   2160
                     Width           =   2535
                     _ExtentX        =   4471
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
                     Format          =   124518401
                     CurrentDate     =   41154
                  End
                  Begin MSComctlLib.ProgressBar ProgressBar1 
                     Height          =   255
                     Left            =   240
                     TabIndex        =   35
                     Top             =   3600
                     Width           =   3735
                     _ExtentX        =   6588
                     _ExtentY        =   450
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
                     Height          =   615
                     Left            =   240
                     Top             =   3000
                     Width           =   3735
                     _Version        =   65536
                     _ExtentX        =   6588
                     _ExtentY        =   1085
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
                  Begin VB.Label Label6 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "ﬂÊœ «· ·„Ì–"
                     BeginProperty Font 
                        Name            =   "Times New Roman"
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
                     Left            =   1200
                     TabIndex        =   283
                     Top             =   2640
                     Width           =   1935
                  End
                  Begin VB.Label Label6 
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
                     Index           =   1
                     Left            =   5160
                     TabIndex        =   274
                     Top             =   1200
                     Width           =   1935
                  End
                  Begin VB.Label Label15 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "Ã‰” «· ·„Ì–                                    Â« › «·ÊﬂÌ·"
                     BeginProperty Font 
                        Name            =   "Times New Roman"
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
                     Left            =   240
                     TabIndex        =   264
                     Top             =   1680
                     Width           =   6855
                  End
                  Begin VB.Label Label15 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "    «·—ﬁ„ «·Êÿ‰Ì                              „Õ· «·„Ì·«œ"
                     BeginProperty Font 
                        Name            =   "Times New Roman"
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
                     Left            =   120
                     TabIndex        =   259
                     Top             =   240
                     Width           =   6975
                  End
                  Begin VB.Label Label15 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "” … «·„Ì·«œ                               «·ﬁ”„                                «·—ﬁ„"
                     BeginProperty Font 
                        Name            =   "Times New Roman"
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
                     Left            =   240
                     TabIndex        =   237
                     Top             =   720
                     Width           =   6855
                  End
                  Begin VB.Label Label15 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   " «—ÌŒ «· ”ÃÌ·                                          RIM"
                     BeginProperty Font 
                        Name            =   "Times New Roman"
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
                     Left            =   2640
                     TabIndex        =   37
                     Top             =   2160
                     Width           =   4335
                  End
                  Begin VB.Image Image2 
                     Appearance      =   0  'Flat
                     Height          =   1095
                     Left            =   4080
                     Stretch         =   -1  'True
                     Top             =   2760
                     Width           =   1935
                  End
                  Begin VB.Shape Shape3 
                     BorderColor     =   &H8000000E&
                     BorderWidth     =   2
                     Height          =   4335
                     Left            =   120
                     Top             =   120
                     Width           =   7095
                  End
                  Begin VB.Shape Shape4 
                     BorderColor     =   &H8000000E&
                     BorderWidth     =   2
                     Height          =   1095
                     Left            =   4080
                     Top             =   2760
                     Width           =   1935
                  End
                  Begin VB.Label Label20 
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
                     Height          =   1095
                     Left            =   4080
                     TabIndex        =   36
                     Top             =   2760
                     Width           =   1935
                  End
               End
               Begin MSComDlg.CommonDialog CommonDialog1 
                  Left            =   2280
                  Top             =   2760
                  _ExtentX        =   847
                  _ExtentY        =   847
                  _Version        =   393216
               End
               Begin VB.Label Label30 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   $"etudiants.frx":0270
                  BeginProperty Font 
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
                  Left            =   2760
                  TabIndex        =   236
                  Top             =   5640
                  Width           =   7695
               End
            End
            Begin MSFlexGridLib.MSFlexGrid grd1 
               Height          =   4095
               Left            =   -74880
               TabIndex        =   43
               Top             =   1080
               Width           =   13815
               _ExtentX        =   24368
               _ExtentY        =   7223
               _Version        =   393216
               FixedRows       =   0
               FixedCols       =   0
               BackColor       =   0
               ForeColor       =   16777215
               BackColorFixed  =   0
               ForeColorFixed  =   16777215
               ForeColorSel    =   8388608
               BackColorBkg    =   0
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
            Begin VB.Line Line2 
               BorderColor     =   &H00FFFFFF&
               X1              =   -68640
               X2              =   -68640
               Y1              =   5520
               Y2              =   6120
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00FFFFFF&
               BorderWidth     =   4
               X1              =   -68760
               X2              =   -68760
               Y1              =   5520
               Y2              =   6120
            End
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8055
         Left            =   120
         ScaleHeight     =   8055
         ScaleWidth      =   14295
         TabIndex        =   5
         Top             =   1080
         Width           =   14295
         Begin VB.CommandButton Command43 
            Caption         =   "”Õ» „⁄ —ﬁ„ «·Â« ›"
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
            Left            =   7200
            TabIndex        =   268
            Top             =   7560
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.CommandButton Command44 
            Caption         =   "”Õ» „⁄ «·—ﬁ„ «·Êÿ‰Ì"
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
            Left            =   9480
            TabIndex        =   265
            Top             =   7560
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.CommandButton Command10 
            Caption         =   "”Õ» „⁄ «·—ﬁ„ «· ”·”·Ì"
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
            Left            =   11880
            TabIndex        =   238
            Top             =   7560
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   5055
            Left            =   3480
            ScaleHeight     =   5055
            ScaleWidth      =   7335
            TabIndex        =   6
            Top             =   1320
            Visible         =   0   'False
            Width           =   7335
            Begin VB.TextBox Text20 
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
               Left            =   4080
               ScrollBars      =   2  'Vertical
               TabIndex        =   279
               Top             =   2160
               Width           =   1935
            End
            Begin VB.TextBox Text19 
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
               TabIndex        =   277
               Top             =   1200
               Width           =   2895
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
               Left            =   4440
               ScrollBars      =   2  'Vertical
               TabIndex        =   266
               Top             =   720
               Width           =   1575
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
               ScrollBars      =   2  'Vertical
               TabIndex        =   263
               Top             =   240
               Width           =   3015
            End
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
               TabIndex        =   184
               Top             =   3960
               Width           =   735
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
               Left            =   4080
               Style           =   2  'Dropdown List
               TabIndex        =   42
               Top             =   1680
               Width           =   1935
            End
            Begin VB.CommandButton Command14 
               Caption         =   "»ÿ«ﬁ… «·œŒÊ·"
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
               Left            =   2760
               TabIndex        =   23
               Top             =   3960
               Width           =   1215
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
               Left            =   2040
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   720
               Width           =   1695
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
               Left            =   240
               TabIndex        =   15
               Top             =   720
               Width           =   1095
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
               TabIndex        =   14
               Top             =   3960
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
               TabIndex        =   13
               Top             =   3960
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
               Left            =   3240
               TabIndex        =   12
               Top             =   1200
               Width           =   2775
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
               TabIndex        =   11
               Top             =   1680
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
               Height          =   375
               Left            =   4440
               ScrollBars      =   2  'Vertical
               TabIndex        =   10
               Top             =   240
               Width           =   1575
            End
            Begin VB.CommandButton Command3 
               Caption         =   "«—›«ﬁ"
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
               Left            =   6120
               TabIndex        =   9
               Top             =   2640
               Width           =   975
            End
            Begin VB.CommandButton Command7 
               Caption         =   "„”Õ"
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
               Left            =   6120
               TabIndex        =   8
               Top             =   3240
               Width           =   975
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
               TabIndex        =   7
               Top             =   3960
               Width           =   735
            End
            Begin MSComCtl2.DTPicker DT1 
               Height          =   375
               Left            =   240
               TabIndex        =   17
               Top             =   2160
               Width           =   2535
               _ExtentX        =   4471
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
               Format          =   124518401
               CurrentDate     =   41154
            End
            Begin MSComctlLib.ProgressBar ProgressBar2 
               Height          =   255
               Left            =   240
               TabIndex        =   18
               Top             =   3600
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   450
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
               Height          =   495
               Left            =   240
               Top             =   3000
               Width           =   3735
               _Version        =   65536
               _ExtentX        =   6588
               _ExtentY        =   873
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
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "ﬂÊœ «· ·„Ì–"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   1200
               TabIndex        =   281
               Top             =   2640
               Width           =   1575
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·Ã‰”                                               —ﬁ„ «·Â« ›"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   2880
               TabIndex        =   280
               Top             =   1680
               Width           =   4215
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ «· ”ÃÌ·                                        RIM     "
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               TabIndex        =   278
               Top             =   2160
               Width           =   4215
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "    «·—ﬁ„ «·Êÿ‰Ì                              „Õ· «·„Ì·«œ"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   120
               TabIndex        =   262
               Top             =   240
               Width           =   6975
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "”‰… «·„Ì·«œ                               «·ﬁ”„                                «·—ﬁ„"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   240
               TabIndex        =   261
               Top             =   720
               Width           =   6855
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00FFFFFF&
               BorderWidth     =   2
               X1              =   120
               X2              =   7200
               Y1              =   4440
               Y2              =   4440
            End
            Begin VB.Label Label90 
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
               Index           =   7
               Left            =   240
               TabIndex        =   232
               Top             =   4560
               Width           =   1095
            End
            Begin VB.Label Label90 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·≈‰«À"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   1200
               TabIndex        =   231
               Top             =   4560
               Width           =   855
            End
            Begin VB.Label Label90 
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
               Index           =   5
               Left            =   2280
               TabIndex        =   230
               Top             =   4560
               Width           =   1095
            End
            Begin VB.Label Label90 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·–ﬂÊ—"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   3240
               TabIndex        =   229
               Top             =   4560
               Width           =   855
            End
            Begin VB.Label Label90 
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
               Index           =   3
               Left            =   4320
               TabIndex        =   228
               Top             =   4560
               Width           =   1335
            End
            Begin VB.Label Label90 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "⁄œœ  ·«„Ì– «·ﬁ”„"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   5160
               TabIndex        =   227
               Top             =   4560
               Width           =   1935
            End
            Begin VB.Label Label6 
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
               Index           =   0
               Left            =   5160
               TabIndex        =   20
               Top             =   1200
               Width           =   1935
            End
            Begin VB.Image Image1 
               Appearance      =   0  'Flat
               Height          =   1215
               Left            =   4080
               Stretch         =   -1  'True
               Top             =   2640
               Width           =   1935
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H8000000E&
               BorderWidth     =   2
               Height          =   4815
               Left            =   120
               Top             =   120
               Width           =   7095
            End
            Begin VB.Shape Shape2 
               BorderColor     =   &H8000000E&
               BorderWidth     =   2
               Height          =   1215
               Left            =   4080
               Top             =   2640
               Width           =   1935
            End
            Begin VB.Label Label10 
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
               Height          =   1095
               Left            =   4080
               TabIndex        =   19
               Top             =   2640
               Width           =   1935
            End
         End
         Begin MSFlexGridLib.MSFlexGrid grd25 
            Height          =   7335
            Left            =   7200
            TabIndex        =   233
            Top             =   120
            Width           =   6975
            _ExtentX        =   12303
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
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   14295
         TabIndex        =   2
         Top             =   360
         Width           =   14295
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
            Left            =   7200
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   120
            Width           =   1695
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
            Left            =   9720
            TabIndex        =   0
            Top             =   120
            Width           =   2535
         End
         Begin VB.CommandButton Command1 
            Caption         =   "≈÷«›…  ·„Ì– ÃœÌœ"
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
            TabIndex        =   3
            Top             =   120
            Width           =   4455
         End
         Begin VB.Label Label28 
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
            Left            =   7560
            TabIndex        =   54
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·—ﬁ„ «· ”·”·Ì ·· ·„Ì–"
            BeginProperty Font 
               Name            =   "Times New Roman"
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
            TabIndex        =   4
            Top             =   120
            Width           =   1935
         End
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   -74400
         OleObjectBlob   =   "etudiants.frx":035D
         Top             =   1440
      End
   End
End
Attribute VB_Name = "etudiants"
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
Dim PicFilev As String
Dim strStream As ADODB.Stream
Dim fName As String
Public co2 As ADODB.Connection
Public cr2 As ADODB.Recordset
Public be As ADODB.Recordset
Public ce As ADODB.Recordset
Public oe As ADODB.Recordset
Public nm As ADODB.Recordset
Public nn As ADODB.Recordset
Public ru As ADODB.Recordset
Dim anes As String
Dim tim As Double
Dim data As New Access.Application
Function cont2()
Set co2 = New ADODB.Connection
Set cr2 = New ADODB.Recordset
Set be = New ADODB.Recordset
Set ce = New ADODB.Recordset
Set oe = New ADODB.Recordset
Set nm = New ADODB.Recordset
Set nn = New ADODB.Recordset
Set ru = New ADODB.Recordset
co2.Provider = "microsoft.jet.oledb.4.0; jet oledb:database password=7346804"
anes = "C" + face.SBB1.Panels(9).Text
co2.ConnectionString = App.Path & "\" & anes & ".mdb"
co2.Open
cr2.Open "select*from Tcarts", co2, adOpenKeyset, adLockOptimistic
be.Open "select*from Tbulletin", co2, adOpenKeyset, adLockOptimistic
ce.Open "select*from Tcartes", co2, adOpenKeyset, adLockOptimistic
oe.Open "select*from Tetudiants", co2, adOpenKeyset, adLockOptimistic
nm.Open "select*from Tnotesmat order by num ASC", co2, adOpenKeyset, adLockOptimistic
nn.Open "select*from Tnni order by aut ASC", co2, adOpenKeyset, adLockOptimistic
ru.Open "select*from Trecus", co2, adOpenKeyset, adLockOptimistic
End Function
 Public Function SavePictureToDB(sFileName As String)
On Error Resume Next
    Call cont
    Do Until et.EOF
    If et!cla = Combo1.Text And et!num = Text3.Text Then
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
 Public Function SavePictureToDB2(sFileName As String)
On Error Resume Next
    Call cont
    Do Until et.EOF
    If et!cla = Combo2.Text And et!num = Text6.Text Then
    Set strStream = New ADODB.Stream
    If Label11.Caption = "" Then
    et.Fields(4).Value = "01"
    et.Update
    Else
    strStream.Type = adTypeBinary
    strStream.Open
    strStream.LoadFromFile sFileName
    et.Fields(4).Value = strStream.Read
    et.Update
    End If
     Exit Function
    End If
  et.MoveNext
    Loop
   End Function
 Private Function LoadPictureFromDB()
On Error Resume Next
Image2.Picture = LoadPicture("")
Call cont
    Do Until et.EOF
    Set strStream = New ADODB.Stream
    strStream.Type = adTypeBinary
    strStream.Open
  If et!ser = Text1.Text Or Val(et!ser) = Val(Text1.Text) Then
    strStream.Write et.Fields(4).Value
   strStream.SaveToFile App.Path & "\aboubekrine.bmp", adSaveCreateOverWrite
       PicFilev = App.Path & "\aboubekrine.bmp"
 Image2.Picture = LoadPicture(App.Path & "\aboubekrine.bmp")
fName = App.Path & "\aboubekrine.bmp"
 Label11.Caption = "01"

    LoadPictureFromDB = True
  End If
    et.MoveNext
    Loop
 ' If Dir$(App.Path & "\aboubekrine.bmp", vbNormal) <> "" Then Kill (App.Path & "\aboubekrine.bmp")

End Function
Private Function LoadPictureFromDB2()
On Error Resume Next
'Image4.Picture = LoadPicture("")
Call cont
    Do Until et.EOF
    Set strStream = New ADODB.Stream
    strStream.Type = adTypeBinary
    strStream.Open
  If et!cla = Combo2.Text And et!num = Text6.Text Then
    strStream.Write et.Fields(4).Value
    strStream.SaveToFile App.Path & "\aboubekrine.bmp", adSaveCreateOverWrite
    'Image4.Picture = LoadPicture(App.Path & "\aboubekrine.bmp")
        FileCopy App.Path & "\aboubekrine.bmp", "C:\photos\1.jpg"
    Kill (App.Path & "\aboubekrine.bmp")
    LoadPictureFromDB6 = True
    End If
    et.MoveNext
    Loop
  If Dir$(App.Path & "\aboubekrine.bmp", vbNormal) <> "" Then Kill (App.Path & "\aboubekrine.bmp")

End Function
Private Function LoadPictureFromDB3()
On Error Resume Next
'Image4.Picture = LoadPicture("")
Call cont
    Do Until et.EOF
    Set strStream = New ADODB.Stream
    strStream.Type = adTypeBinary
    strStream.Open
  If et!cla = Combo1.Text And et!num = Text3.Text Then
    strStream.Write et.Fields(4).Value
    strStream.SaveToFile App.Path & "\aboubekrine.bmp", adSaveCreateOverWrite
    'Image4.Picture = LoadPicture(App.Path & "\aboubekrine.bmp")
        FileCopy App.Path & "\aboubekrine.bmp", "C:\photos\1.jpg"
    Kill (App.Path & "\aboubekrine.bmp")
    LoadPictureFromDB3 = True
    End If
    et.MoveNext
    Loop
  If Dir$(App.Path & "\aboubekrine.bmp", vbNormal) <> "" Then Kill (App.Path & "\aboubekrine.bmp")

End Function
Private Function LoadPictureFromDB4()
On Error Resume Next
'Image4.Picture = LoadPicture("")
Call cont
    Do Until et.EOF
    Set strStream = New ADODB.Stream
    strStream.Type = adTypeBinary
    strStream.Open
  If et!cla = Combo11.Text And et!num = Label95.Caption Then
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
Private Function LoadPictureFromDB5()
On Error Resume Next
'Image4.Picture = LoadPicture("")
Call cont
    Do Until et.EOF
    Set strStream = New ADODB.Stream
    strStream.Type = adTypeBinary
    strStream.Open
  If et!cla = Combo11.Text And et!num = Label95.Caption Then
    strStream.Write et.Fields(4).Value
    strStream.SaveToFile App.Path & "\aboubekrine.bmp", adSaveCreateOverWrite
    'Image4.Picture = LoadPicture(App.Path & "\aboubekrine.bmp")
        FileCopy App.Path & "\aboubekrine.bmp", "C:\photos\" & Label100.Caption & ".jpg"
    Kill (App.Path & "\aboubekrine.bmp")
    LoadPictureFromDB6 = True
    End If
    et.MoveNext
    Loop
  If Dir$(App.Path & "\aboubekrine.bmp", vbNormal) <> "" Then Kill (App.Path & "\aboubekrine.bmp")

End Function
Private Function LoadPictureFromDB6()
On Error Resume Next
Image4.Picture = LoadPicture("")
Call cont
    Do Until et.EOF
    Set strStream = New ADODB.Stream
    strStream.Type = adTypeBinary
    strStream.Open
  If et!cla = Label116.Caption And et!num = Label117.Caption Then
    strStream.Write et.Fields(4).Value
    strStream.SaveToFile App.Path & "\aboubekrine.bmp", adSaveCreateOverWrite
    Image4.Picture = LoadPicture(App.Path & "\aboubekrine.bmp")
    '    FileCopy App.Path & "\aboubekrine.bmp", "C:\photos\1.jpg"
    Kill (App.Path & "\aboubekrine.bmp")
    LoadPictureFromDB6 = True
    End If
    et.MoveNext
    Loop
  If Dir$(App.Path & "\aboubekrine.bmp", vbNormal) <> "" Then Kill (App.Path & "\aboubekrine.bmp")

End Function


Private Sub Combo1_Change()
On Error Resume Next
Call numsetu
Call numclasfornni1
Text2.SetFocus
End Sub

Private Sub Combo1_Click()
On Error Resume Next
Combo1_Change
End Sub

Private Sub Combo10_Change()
On Error Resume Next
grd8.Clear
grd8.Rows = 1
Call cont
Do While Not mt.EOF
If mt!cla = Combo11.Text And mt!mat = Combo10.Text Then
Label41.Caption = mt!nbr
Label67.Caption = mt!nme
Label104.Caption = mt!cof
Exit Sub
End If
mt.MoveNext
Loop
End Sub

Private Sub Combo10_Click()
On Error Resume Next
Combo10_Change
End Sub

Private Sub Combo11_Change()
On Error Resume Next
Dim i As Double
SSTab3.Visible = False
grd6.Visible = False
grd8.Clear
grd8.Rows = 1
Check1.Value = 0
Image3.Picture = LoadPicture("")
Call cont
Do While Not cl.EOF
If Combo11.Text = cl!cla Then
Label106.Caption = cl!num
cl.MoveLast
End If
cl.MoveNext
Loop
Call chargegrd6
grd6.Visible = True
Label95.Caption = ""
Label96.Caption = ""
Label115.Caption = ""
grd7.Clear
grd7.Rows = 1
grd7.Cols = 4
grd7.ColWidth(0) = 0
grd7.ColWidth(1) = 1200
grd7.ColWidth(2) = 3500
grd7.ColWidth(3) = 1500
grd7.row = 0
grd7.Col = 1
grd7.Text = "«·—ﬁ„"
grd7.Col = 2
grd7.Text = "«·«”„"
grd7.Col = 3
grd7.Text = "«·—ﬁ„ «· ”·”·Ì"
grd7.ColAlignment(1) = 1
grd7.ColAlignment(2) = 1
grd7.ColAlignment(3) = 1
SSTab3.Visible = True
Combo10.Clear
grd9.Clear
grd9.Rows = 1
i = 1
Call cont
grd9.Rows = mt.RecordCount + 2
Do While Not mt.EOF
If Combo11.Text = mt!cla Then
Combo10.AddItem mt!mat
grd9.row = i
grd9.Col = 1
grd9.Text = mt!mat
i = i + 1
End If
mt.MoveNext
Loop
grd9.Rows = i
Picture18.Visible = False
End Sub

Private Sub Combo11_Click()
On Error Resume Next
Combo11_Change
End Sub

Private Sub Combo2_Change()
On Error Resume Next
Call numsetu2
Call numclasfornni2
If SSTab2.Visible = True Then
Text7.SetFocus
End If
End Sub

Private Sub Combo2_Click()
On Error Resume Next
Combo2_Change
End Sub

Private Sub Combo4_Change()
On Error Resume Next
Text4.SetFocus
End Sub

Private Sub Combo4_Click()
On Error Resume Next
Combo4_Change
End Sub

Private Sub Combo5_Change()
On Error Resume Next
Picture4.Visible = False
Combo6.Clear
Command10.Visible = True
Command43.Visible = True
Command44.Visible = True
Call cont
Do While Not et.EOF
If Combo5.Text = et!cla And Val(et!num) < 1000000 Then
Combo6.AddItem et!num
End If
et.MoveNext
Loop
Picture6.Visible = False
Text1.Text = ""
grd25.Visible = False
Call chargegrd25
grd25.Visible = True
Call numclasfornni3
End Sub

Private Sub Combo5_Click()
On Error Resume Next
Combo5_Change
End Sub

Private Sub Combo6_Change()
On Error Resume Next
Text1.Text = ""
Call cont
Do While Not et.EOF
If et!cla = Combo5.Text And Combo6.Text = et!num And Val(et!num) < 1000000 Then
Text1.Text = et!ser
Command8_Click
Exit Sub
End If
et.MoveNext
Loop
End Sub

Private Sub Combo6_Click()
On Error Resume Next
Combo6_Change
End Sub





Private Sub Command1_Click()
On Error Resume Next
Command10.Visible = False
Command43.Visible = False
Command44.Visible = False
Label11.Caption = ""
PicFilev = ""
Text5.Text = ""
Text15.Text = ""
Text2.Text = ""
Text20.Text = ""
Text19.Text = ""
Text3.Text = ""
Text4.Text = ""
Text1.Text = ""
Label11.Caption = ""
grd25.Visible = False
Call chargec1
Call chargec2
PicFilev = ""
Label11.Caption = ""
Image1.Picture = LoadPicture(PicFilev) 'Afficher l'image
Call serial
Picture6.Visible = False
Command14.Enabled = False
Combo1.Enabled = True
Combo4.Enabled = True
Text2.Enabled = True
Text20.Enabled = True
Text19.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
DT1.Enabled = True
Command3.Enabled = True
Command7.Enabled = True
Command5.Enabled = True
Command4.Enabled = True
End Sub

Private Sub Command10_Click()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim ane As String
Command10.Enabled = False
grd25.Visible = False
Call chargegrd25
grd25.Visible = True
Call cont2
Do While Not oe.EOF
oe.Delete
oe.MoveNext
Loop
n = grd25.Rows
For i = 1 To n - 1
oe.AddNew
oe!cla = Combo5.Text
grd25.row = i
grd25.Col = 1
oe!num = grd25.Text
grd25.Col = 2
oe!nom = grd25.Text
grd25.Col = 3
oe!ser = grd25.Text
grd25.Col = 4
oe!tel = grd25.Text
grd25.Col = 5
oe!adr = grd25.Text
oe.Update
Next i
Call cont2
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
'data.DoCmd.ApplyFilter cla
data.DoCmd.Maximize
'data.DoCmd.OpenReport "Bulletin", acViewPreview, , "num =" & a, acWindowNormal, OpenArgs
data.DoCmd.OpenReport "Tetudiants", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
'data.DCount
    'data.CloseCurrentDatabase
data.Visible = True
Set data = Nothing
Command10.Enabled = True
End Sub

Private Sub Command11_Click()
On Error GoTo p
Dim clnm As String
Text6.Text = Trim(Text6.Text)
Text7.Text = Trim(Text7.Text)
Text18.Text = Trim(Text18.Text)
Text21.Text = Trim(Text21.Text)
Text8.Text = Trim(Text8.Text)
Text9.Text = Trim(Text9.Text)
If Combo2.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·ﬁ”„", vbCritical
Exit Sub
End If
If Text6.Text = "" Then
MsgBox "«œŒ· —ﬁ„ «· ·„Ì–", vbCritical
Text6.SetFocus
Exit Sub
End If
If Text7.Text = "" Then
MsgBox "«œŒ· «”„ «· ·„Ì–", vbCritical
Text7.SetFocus
Exit Sub
End If
If Combo3.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— Ã‰” «· ·„Ì–", vbCritical
Exit Sub
End If
If Text8.Text = "" Then
MsgBox "«œŒ· —ﬁ„ ÊﬂÌ· «· ·„Ì–", vbCritical
Text8.SetFocus
Exit Sub
End If
Label11.Caption = ""
PicFilev = ""
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "Picture JPG |*.jpg|Picture Gif |*.Gif|Picture Bmp |*.Bmp|Picture Icon |*.ICO|All Picture |*.*"
    CommonDialog1.DialogTitle = "Picture"
    CommonDialog1.ShowOpen
    PicFilev = CommonDialog1.FileName 'lien d'image
    Image2.Picture = LoadPicture(PicFilev) 'Afficher l'image
 fName = CommonDialog1.FileName
 Label11.Caption = CommonDialog1.FileName
 'Image1.Width = 1900
 'Image1.Height = 1500
 'Command4.Enabled = True
 'Command2.Enabled = True
Exit Sub
p:
MsgBox "Êﬁ⁄ Œÿ√ ›Ì  Õ„Ì· «·ÊÀÌﬁ… , «·—Ã«¡ «⁄«œ… «·„Õ«Ê·…", vbExclamation
PicFilev = ""
Label11.Caption = ""
   Image2.Picture = LoadPicture(PicFilev) 'Afficher l'image
 fName = ""
End Sub

Private Sub Command12_Click()
On Error Resume Next
PicFilev = ""
Label11.Caption = ""
Image2.Picture = LoadPicture(PicFilev) 'Afficher l'image
fName = ""

End Sub


Private Sub Command13_Click()
On Error Resume Next
grd1.Clear
grd1.Rows = 1
grd1.Visible = False
Call cont
Do While Not cl.EOF
If Label36.Caption = cl!cla Then
Label106.Caption = cl!num
cl.MoveLast
End If
cl.MoveNext
Loop
Call coffes
Call chargegrd1
Call calculmoyenne
grd1.Visible = True
Label42.Caption = ""
End Sub

Private Sub Command14_Click()
On Error Resume Next
Dim Security As SECURITY_ATTRIBUTES
Dim x$
If PicFilev = "" Then
MsgBox "ÌÃ» «—›«ﬁ ’Ê—…", vbCritical
Exit Sub
End If
x$ = Dir$("C:\photos\1.jpg")
If x$ = "" Then
    'Create a directory
    Ret& = CreateDirectory("C:\photos", Security)
FileCopy App.Path & "\nophoto.jpg", "C:\photos\1.jpg"
    'If CreateDirectory returns 0, the function has failed
    'If Ret& = 0 Then MsgBox "Error : Couldn't create directory !", vbCritical + vbOKOnly
End If
Call cont2
Do While Not cr2.EOF
cr2.Delete
cr2.MoveNext
Loop
cr2.AddNew
cr2!cla = Combo1.Text
cr2!num = Text3.Text
cr2!dat = DT1.Value
cr2!nom = Text2.Text
cr2!tel = Text4.Text
cr2!adr = Text5.Text
cr2!ser = BarcodeX1.Caption
cr2!eta = face.SBB1.Panels(13).Text
cr2!ann = face.SBB1.Panels(9).Text
cr2.Update
Timer4.Enabled = True

End Sub

Private Sub Command15_Click()
On Error Resume Next
Dim Security As SECURITY_ATTRIBUTES
Dim x$
If PicFilev = "" Then
MsgBox "ÌÃ» «—›«ﬁ ’Ê—…", vbCritical
Exit Sub
End If
x$ = Dir$("C:\photos\1.jpg")
If x$ = "" Then
    'Create a directory
    Ret& = CreateDirectory("C:\photos", Security)
FileCopy App.Path & "\nophoto.jpg", "C:\photos\1.jpg"
    'If CreateDirectory returns 0, the function has failed
    'If Ret& = 0 Then MsgBox "Error : Couldn't create directory !", vbCritical + vbOKOnly
End If
Call cont2
Do While Not cr2.EOF
cr2.Delete
cr2.MoveNext
Loop
cr2.AddNew
cr2!cla = Combo2.Text
cr2!num = Text6.Text
cr2!dat = DT2.Value
cr2!nom = Text7.Text
cr2!tel = Text8.Text
cr2!adr = Text9.Text
cr2!ser = BarcodeX2.Caption
cr2!eta = face.SBB1.Panels(13).Text
cr2!ann = face.SBB1.Panels(9).Text
cr2.Update
Timer3.Enabled = True
End Sub


Private Sub Command16_Click()
On Error Resume Next
Dim ane As String
Dim a As Double
Dim f As Double
Dim tx As String
Dim i As Double
Dim n As Double
Dim momat As Double
Dim comat As Double
Dim tomat As Double
Dim momats As Double
Dim comats As Double
Dim tomats As Double
Dim moy As Double
Dim b As Double
Dim moyenne As Double
Dim menar As String
Dim menfr As String
Dim cs As Double
Dim ts As Double
If Label33.Caption = "" Then
MsgBox "·« Ì„ﬂ‰ ⁄—÷ ﬂ‘› «·œ—Ã«  Õ Ï Ì „  ÕœÌœ «·„⁄œ· «·⁄«„", vbCritical
Exit Sub
End If
Call cont2
Do While Not be.EOF
If Label36.Caption = be!cla Then
momats = 0
comats = 0
tomats = 0
momat = 0
comat = 0
tomat = 0
If Val(be!mm1) > 0 Or be!mm1 = "0" Then
momat = be!mm1
momats = momats + momat
comat = be!cof1
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm2) > 0 Or be!mm2 = "0" Then
momat = be!mm2
momats = momats + momat
comat = be!cof2
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm3) > 0 Or be!mm3 = "0" Then
momat = be!mm3
momats = momats + momat
comat = be!cof3
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm4) > 0 Or be!mm4 = "0" Then
momat = be!mm4
momats = momats + momat
comat = be!cof4
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm5) > 0 Or be!mm5 = "0" Then
momat = be!mm5
momats = momats + momat
comat = be!cof5
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm6) > 0 Or be!mm6 = "0" Then
momat = be!mm6
momats = momats + momat
comat = be!cof6
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm7) > 0 Or be!mm7 = "0" Then
momat = be!mm7
momats = momats + momat
comat = be!cof7
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm8) > 0 Or be!mm8 = "0" Then
momat = be!mm8
momats = momats + momat
comat = be!cof8
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm9) > 0 Or be!mm9 = "0" Then
momat = be!mm9
momats = momats + momat
comat = be!cof9
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm10) > 0 Or be!mm10 = "0" Then
momat = be!mm10
momats = momats + momat
comat = be!cof10
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm11) > 0 Or be!mm11 = "0" Then
momat = be!mm11
momats = momats + momat
comat = be!cof11
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm12) > 0 Or be!mm12 = "0" Then
momat = be!mm12
momats = momats + momat
comat = be!cof12
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm13) > 0 Or be!mm13 = "0" Then
momat = be!mm13
momats = momats + momat
comat = be!cof13
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm14) > 0 Or be!mm14 = "0" Then
momat = be!mm14
momats = momats + momat
comat = be!cof14
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm15) > 0 Or be!mm15 = "0" Then
momat = be!mm15
momats = momats + momat
comat = be!cof15
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm16) > 0 Or be!mm16 = "0" Then
momat = be!mm16
momats = momats + momat
comat = be!cof16
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm17) > 0 Or be!mm17 = "0" Then
momat = be!mm17
momats = momats + momat
comat = be!cof17
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm18) > 0 Or be!mm18 = "0" Then
momat = be!mm18
momats = momats + momat
comat = be!cof18
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm19) > 0 Or be!mm19 = "0" Then
momat = be!mm19
momats = momats + momat
comat = be!cof19
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm20) > 0 Or be!mm20 = "0" Then
momat = be!mm20
momats = momats + momat
comat = be!cof20
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
moy = 0
If comats > 0 Then
moy = tomats / comats
MyNumber = Round(moy, 2)
moy = MyNumber
End If
Label33.Caption = moy
a = 0
f = 6
'***** Mentions
Call coffes
Label40.Caption = ""
Label29.Caption = ""
For i = 5 To 15
If moy = 0 And comats = 0 Then
Label40.Caption = ""
Label29.Caption = ""
Label33.Caption = ""
i = 15
'Exit Sub
End If
If moy = 0 And comats > 0 Then
Label40.Caption = mens(5).Text
Label29.Caption = mens(11).Text
i = 15
'Exit Sub
End If
cs = coff(i - 1).Text
ts = coff(i).Text
If moy > ts And moy <= cs Then
Label40.Caption = mens(a).Text
Label29.Caption = mens(f).Text
i = 15
End If
If i = 5 Or i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 15 Then
a = a + 1
f = f + 1
End If
Next i
If Check4.Value = 1 Then
be!dat = Date
Else
be!dat = ""
End If
If Check5.Value = 1 Then
be!Abs = Label73.Caption
Else
be!Abs = ""
End If
If Check2.Value = 1 Then
be!ran = Label42.Caption
Else
be!ran = ""
End If
If Check3.Value = 1 Then
be!mena = Label40.Caption
be!menf = Label29.Caption
Else
be!mena = ""
be!menf = ""
End If
If Label37.Caption = be!mtr Then
be!nom = Label38.Caption
be!ser = BarcodeX2.Caption
moyenne = Label33.Caption
menar = Label40.Caption
menfr = Label29.Caption
b = be!num
End If
be!numm = Label106.Caption
be!eco = face.SBB1.Panels(13).Text
be!ann = face.SBB1.Panels(9).Text
be!moy = Label33.Caption
be.Update
'be.MoveLast
End If
be.MoveNext
Loop
Label33.Caption = moyenne
Label40.Caption = menar
Label29.Caption = menfr
If Check2.Value = 1 Then
Command21_Click
End If
Call cont2
i = Label41.Caption
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
data.DoCmd.Maximize
If i <= 7 Then
data.DoCmd.OpenReport "notes7", acViewPreview, , "num =" & b, acWindowNormal, OpenArgs
ElseIf i > 7 And i <= 12 Then
data.DoCmd.OpenReport "notes7", acViewPreview, , "num =" & b, acWindowNormal, OpenArgs
Else
data.DoCmd.OpenReport "Bulletin20", acViewPreview, , "num =" & b, acWindowNormal, OpenArgs
End If
'data.DoCmd.RunMacro "Opnreport", 2
data.Visible = True
'Set data = Nothing

End Sub

Private Sub Command17_Click()
On Error Resume Next
Call coffes
Picture11.Visible = True
End Sub

Private Sub Command18_Click()
On Error Resume Next
Picture11.Visible = False
End Sub

Private Sub Command19_Click()
On Error Resume Next
Call coffes1
Command20_Click
End Sub


Private Sub Command2_Click()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim ane As String
Command2.Enabled = False
Call cont2
Do While Not nm.EOF
nm.Delete
nm.MoveNext
Loop
n = grd8.Rows
For i = 1 To n - 1
nm.AddNew
nm!cla = Combo11.Text
nm!mat = Combo10.Text
nm!pro = Text12.Text
grd8.row = i
grd8.Col = 19
nm!num = Val(grd8.Text)
grd8.Col = 1
nm!nom = grd8.Text
grd8.Col = 2
nm!cof = grd8.Text
grd8.Col = 3
nm!dv1 = grd8.Text
grd8.Col = 4
nm!dv2 = grd8.Text
grd8.Col = 5
nm!dv3 = grd8.Text
grd8.Col = 6
nm!dv4 = grd8.Text
grd8.Col = 7
nm!dv5 = grd8.Text
grd8.Col = 8
nm!dv6 = grd8.Text
grd8.Col = 9
nm!dv7 = grd8.Text
grd8.Col = 10
nm!dv8 = grd8.Text
grd8.Col = 11
nm!dv9 = grd8.Text
grd8.Col = 12
nm!dv10 = grd8.Text
grd8.Col = 13
nm!mdv = grd8.Text
grd8.Col = 14
nm!ex1 = grd8.Text
grd8.Col = 15
nm!ex2 = grd8.Text
grd8.Col = 16
nm!ex3 = grd8.Text
grd8.Col = 17
nm!mmt = grd8.Text
grd8.Col = 18
nm!tot = grd8.Text
nm.Update
Next i
tim = 2
Timer8.Enabled = True
End Sub

Private Sub Command20_Click()
On Error Resume Next
Dim i As Integer
Dim j As Integer
Dim a As Double
Dim b As Double
Dim c As String
Dim d As String
For i = 0 To 11
If mens(i).Text = "" Or coff(i).Text = "" Then
MsgBox "·«Ì‰»€Ì ÊÃÊœ Õﬁ· ›«—€", vbCritical
Exit Sub
End If
Next i
For i = 11 To 15
If coff(i).Text = "" Then
MsgBox "·«Ì‰»€Ì ÊÃÊœ Õﬁ· ›«—€", vbCritical
Exit Sub
End If
Next i
j = 0
For i = 4 To 15
j = j + 1
d = j
If coff(i).Text <> "" Then
If i = 5 Or i = 7 Or i = 9 Or i = 11 Or i = 13 Then
a = coff(i).Text
b = coff(i - 1).Text
c = b
coff(i + 1).Text = coff(i).Text
If a >= b Then
MsgBox " «·⁄œœ «·„œŒ· ÌÃ» √‰ ÌﬂÊ‰ √’€— „‰ " + c + " ›Ì «·Œ«‰… —ﬁ„ " + d, vbCritical
coff(i).Text = ""
Exit Sub
End If
End If
End If
Next i
Label24.Caption = coff(0).Text
Label25.Caption = coff(1).Text
Label26.Caption = coff(2).Text
Label27.Caption = coff(3).Text
Call cont
cf!cof0 = coff(0).Text
cf!cof1 = coff(1).Text
cf!cof2 = coff(2).Text
cf!cof3 = coff(3).Text
cf!cof4 = coff(4).Text
cf!cof5 = coff(5).Text
cf!cof6 = coff(6).Text
cf!cof7 = coff(7).Text
cf!cof8 = coff(8).Text
cf!cof9 = coff(9).Text
cf!cof10 = coff(10).Text
cf!cof11 = coff(11).Text
cf!cof12 = coff(12).Text
cf!cof13 = coff(13).Text
cf!cof14 = coff(14).Text
cf!cof15 = coff(15).Text
cf!tex9 = mens(0).Text
cf!tex12 = mens(1).Text
cf!tex15 = mens(2).Text
cf!tex18 = mens(3).Text
cf!tex19 = mens(4).Text
cf!tex20 = mens(5).Text
cf!tex21 = mens(6).Text
cf!tex22 = mens(7).Text
cf!tex23 = mens(8).Text
cf!tex24 = mens(9).Text
cf!tex25 = mens(10).Text
cf!tex26 = mens(11).Text
cf.Update
Command18.Enabled = False
Command19.Enabled = False
Command20.Enabled = False
Call changementdecoffcients
Command18.Enabled = True
Command19.Enabled = True
Command20.Enabled = True
End Sub

Private Sub Command21_Click()
On Error Resume Next
Dim i As Double
Dim tx1 As String
Dim tx2 As String
If Label33.Caption = "" Then
MsgBox "·« Ì„ﬂ‰  ÕœÌœ «·— »… Õ Ï Ì „  ÕœÌœ «·„⁄œ· «·⁄«„", vbCritical
Exit Sub
End If
Command21.Enabled = False
Label42.Visible = False
i = 1
Call cont2
grd5.Rows = be.RecordCount + 2
Do While Not be.EOF
If be!cla = Label36.Caption Then
grd5.row = i
grd5.Col = 0
grd5.Text = be!mtr
grd5.Col = 1
grd5.Text = be!nom
grd5.Col = 2
grd5.Text = be!moy
i = i + 1
End If
be.MoveNext
Loop
grd5.Rows = i
n = grd5.Rows
grd5.Col = 2
grd5.Sort = 2
For i = 1 To n - 1
grd5.row = i
grd5.Col = 0
tx = grd5.Text
If tx = Label37.Caption Then
Label42.Caption = i
End If
grd5.Col = 3
grd5.Text = i
Next i
n = grd5.Rows
Call cont2
Do While Not be.EOF
If Label36.Caption = be!cla Then
For i = 1 To n - 1
grd5.row = i
grd5.Col = 0
tx1 = grd5.Text
grd5.Col = 3
tx2 = grd5.Text
If be!mtr = tx1 Then
be!ran = tx2
be.Update
i = n
End If
Next i
End If
be.MoveNext
Loop
Label42.Visible = True
Command21.Enabled = True
'Timer5.Enabled = True
End Sub

Private Sub Command22_Click()
On Error Resume Next
Dim a As Double
Dim b As Double
Dim c As Double
Dim dat1 As Date
Dim dat2 As Date
If Combo7.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— ”«⁄… «·»œ«Ì…", vbCritical
Exit Sub
End If
If Combo8.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— ”«⁄… «·‰Â«Ì…", vbCritical
Exit Sub
End If
If Combo9.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·„«œ…", vbCritical
Exit Sub
End If
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
Call cont
Do While Not ab.EOF
dat1 = DT3.Value
dat2 = ab!dat
If Label87.Caption <> ab!aut And Label36.Caption = ab!cla And Label37.Caption = ab!num And dat1 = dat2 And Combo7.Text = ab!hr1 And ab!hr2 = Combo8.Text Then
MsgBox "Â–Â «·„⁄·Ê„… ”»ﬁ Ê√‰  „ Õ›ŸÂ«", vbCritical
Exit Sub
End If
ab.MoveNext
Loop
If Label87.Caption <> "" Then
Call cont
Do While Not ab.EOF
If Label87.Caption = ab!aut Then
ab!dat = DT3.Value
ab!hr1 = Combo7.Text
ab!hr2 = Combo8.Text
ab!nbr = c
ab!mat = Combo9.Text
ab!rem = Text10.Text
ab.Update
Timer6.Enabled = True
Exit Sub
End If
ab.MoveNext
Loop
End If
ab.AddNew
ab!cla = Label36.Caption
ab!num = Label37.Caption
ab!nom = Label76.Caption
ab!dat = DT3.Value
ab!hr1 = Combo7.Text
ab!hr2 = Combo8.Text
ab!nbr = c
ab!mat = Combo9.Text
ab!rem = Text10.Text
ab.Update
Timer6.Enabled = True

End Sub

Private Sub Command23_Click()
On Error Resume Next
Label87.Caption = ""
ProgressBar4.Value = 0
Timer6.Enabled = False
Command24_Click
End Sub

Private Sub Command24_Click()
On Error Resume Next
Call chargec3
Combo9.Clear
Call cont
Do While Not mt.EOF
If Label36.Caption = mt!cla Then
Combo9.AddItem mt!mat
End If
mt.MoveNext
Loop
grd2.Visible = False
grd2.Clear
grd2.Rows = 1
grd2.Cols = 7
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 2000
grd2.ColWidth(2) = 800
grd2.ColWidth(3) = 800
grd2.ColWidth(4) = 1400
grd2.ColWidth(5) = 3000
grd2.ColWidth(6) = 5200
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.ColAlignment(4) = 1
grd2.ColAlignment(5) = 1
grd2.ColAlignment(6) = 1
Call chargegrd2
grd2.Visible = True
Picture12.Visible = True
End Sub

Private Sub Command25_Click()
On Error Resume Next
If Label87.Caption = "" Then
MsgBox "ﬁ„ »«·÷€ÿ ⁄·Ï «·»Ì«‰«  «·„—«œ Õ–›Â«", vbCritical
Exit Sub
End If
g = MsgBox("Â·  —Ìœ «·Õ–› Õﬁ« ", vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
Call cont
Do While Not ab.EOF
If Label87.Caption = ab!aut Then
ab.Delete
Timer6.Enabled = True
Exit Sub
End If
ab.MoveNext
Loop
End If

End Sub

Private Sub Command26_Click()
On Error Resume Next
Dim a As Double
If Label88.Caption = "" Then
MsgBox "·« ÌÊÃœ ·Â– «· ·„Ì– √Ì —ﬁ„  ”·”·Ì Ì„ﬂ‰ «·Õ–› ⁄·Ï √”«”Â", vbCritical
Exit Sub
End If
g = MsgBox("Â·  —Ìœ «·Õ–› Õﬁ« ", vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
Call cont2
Do While Not be.EOF
If BarcodeX2.Caption = be!ser Then
be.Delete
If be.RecordCount > 0 Then
be.MoveLast
End If
End If
be.MoveNext
Loop
Call cont2
Do While Not nn.EOF
If BarcodeX2.Caption = nn!ser Then
nn.Delete
If nn.RecordCount > 0 Then
nn.MoveLast
End If
End If
nn.MoveNext
Loop
Call cont
a = sr!nes
sr!nes = a + 1
sr.Update
Call cont
Do While Not et.EOF
If Label88.Caption = et!ser Then
et!num = a
'et!act = "0"
et.Update
Timer7.Enabled = True
Exit Sub
End If
et.MoveNext
Loop
End If

End Sub

Private Sub Command27_Click()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim ane As String
Command27.Enabled = False
Call cont2
Do While Not oe.EOF
oe.Delete
oe.MoveNext
Loop
n = grd6.Rows
For i = 1 To n - 1
oe.AddNew
oe!cla = Combo11.Text
grd6.row = i
grd6.Col = 1
oe!num = grd6.Text
grd6.Col = 2
oe!nom = grd6.Text
grd6.Col = 3
oe!ser = grd6.Text
oe.Update
Next i
Call cont2
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
'data.DoCmd.ApplyFilter cla
data.DoCmd.Maximize
'data.DoCmd.OpenReport "Bulletin", acViewPreview, , "num =" & a, acWindowNormal, OpenArgs
data.DoCmd.OpenReport "Tetudiants", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
'data.DCount
    'data.CloseCurrentDatabase
data.Visible = True
Set data = Nothing
Command27.Enabled = True
End Sub

Private Sub Command28_Click()
On Error Resume Next
Dim n As Double
Dim i As Double
Dim cla As String
Dim num As String
Dim Security As SECURITY_ATTRIBUTES
Dim x$
x$ = Dir$("C:\photos\11.jpg")
If x$ = "" Then
    'Create a directory
    Ret& = CreateDirectory("C:\photos", Security)
FileCopy App.Path & "\TETE.jpg", "C:\photos\11.jpg"
    'If CreateDirectory returns 0, the function has failed
    'If Ret& = 0 Then MsgBox "Error : Couldn't create directory !", vbCritical + vbOKOnly
End If
If Label95.Caption = "" Then
MsgBox " «÷€ÿ ⁄·Ï «·≈”„ ·≈÷«› Â ", vbCritical
Exit Sub
End If
n = grd7.Rows
If n >= 11 Then
MsgBox " «·Õœ «·√⁄·Ï ··≈÷«›… 10  ·«„Ì–", vbCritical
Exit Sub
End If
For i = 1 To n - 1
grd7.row = i
grd7.Col = 0
cla = grd7.Text
grd7.Col = 1
num = grd7.Text
If cla = Combo11.Text And num = Label95.Caption Then
MsgBox " ”»ﬁ Ê√‰ √÷Ì› Â–« «· ·„Ì– ", vbCritical
Exit Sub
End If
Next i
grd7.Rows = n + 1
grd7.row = n
grd7.Col = 0
grd7.Text = Combo11.Text
grd7.Col = 1
grd7.Text = Label95.Caption
grd7.Col = 2
grd7.Text = Label96.Caption
grd7.Col = 3
grd7.Text = Label115.Caption
Label100.Caption = n
Call LoadPictureFromDB5

End Sub

Private Sub Command29_Click()
On Error Resume Next
Dim a As Double
a = 0
If grd7.Rows < 11 Then
MsgBox " ·«Ì„ﬂ‰ ⁄—÷ »ÿ«ﬁ«  «·œŒÊ· ≈·« »⁄œ ≈ﬂ„«·Â« »⁄‘—…  ·«„Ì– ", vbCritical
Exit Sub
End If
'If a = 10 Then
Call cont2
'ce.AddNew
ce!eco = face.SBB1.Panels(13).Text
ce!ann = face.SBB1.Panels(9).Text
grd7.row = 1
grd7.Col = 0
ce!cla1 = grd7.Text
grd7.Col = 1
ce!mtr1 = grd7.Text
grd7.Col = 2
ce!nom1 = grd7.Text
grd7.Col = 3
ce!ser1 = grd7.Text
grd7.row = 2
grd7.Col = 0
ce!cla2 = grd7.Text
grd7.Col = 1
ce!mtr2 = grd7.Text
grd7.Col = 2
ce!nom2 = grd7.Text
grd7.Col = 3
ce!ser2 = grd7.Text
grd7.row = 3
grd7.Col = 0
ce!cla3 = grd7.Text
grd7.Col = 1
ce!mtr3 = grd7.Text
grd7.Col = 2
ce!nom3 = grd7.Text
grd7.Col = 3
ce!ser3 = grd7.Text
grd7.row = 4
grd7.Col = 0
ce!cla4 = grd7.Text
grd7.Col = 1
ce!mtr4 = grd7.Text
grd7.Col = 2
ce!nom4 = grd7.Text
grd7.Col = 3
ce!ser4 = grd7.Text
grd7.row = 5
grd7.Col = 0
ce!cla5 = grd7.Text
grd7.Col = 1
ce!mtr5 = grd7.Text
grd7.Col = 2
ce!nom5 = grd7.Text
grd7.Col = 3
ce!ser5 = grd7.Text
grd7.row = 6
grd7.Col = 0
ce!cla6 = grd7.Text
grd7.Col = 1
ce!mtr6 = grd7.Text
grd7.Col = 2
ce!nom6 = grd7.Text
grd7.Col = 3
ce!ser6 = grd7.Text
grd7.row = 7
grd7.Col = 0
ce!cla7 = grd7.Text
grd7.Col = 1
ce!mtr7 = grd7.Text
grd7.Col = 2
ce!nom7 = grd7.Text
grd7.Col = 3
ce!ser7 = grd7.Text
grd7.row = 8
grd7.Col = 0
ce!cla8 = grd7.Text
grd7.Col = 1
ce!mtr8 = grd7.Text
grd7.Col = 2
ce!nom8 = grd7.Text
grd7.Col = 3
ce!ser8 = grd7.Text
grd7.row = 9
grd7.Col = 0
ce!cla9 = grd7.Text
grd7.Col = 1
ce!mtr9 = grd7.Text
grd7.Col = 2
ce!nom9 = grd7.Text
grd7.Col = 3
ce!ser9 = grd7.Text
grd7.row = 10
grd7.Col = 0
ce!cla10 = grd7.Text
grd7.Col = 1
ce!mtr10 = grd7.Text
grd7.Col = 2
ce!nom10 = grd7.Text
grd7.Col = 3
ce!ser10 = grd7.Text
ce.Update
tim = 1
Timer8.Enabled = True
End Sub

Private Sub Command3_Click()
On Error GoTo p
Dim clnm As String
Text2.Text = Trim(Text2.Text)
Text20.Text = Trim(Text20.Text)
Text19.Text = Trim(Text19.Text)
Text3.Text = Trim(Text3.Text)
Text4.Text = Trim(Text4.Text)
Text5.Text = Trim(Text5.Text)
If Combo1.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·ﬁ”„", vbCritical
Exit Sub
End If
If Text3.Text = "" Then
MsgBox "«œŒ· —ﬁ„ «· ·„Ì–", vbCritical
Text3.SetFocus
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "«œŒ· «”„ «· ·„Ì–", vbCritical
Text2.SetFocus
Exit Sub
End If
If Combo4.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— Ã‰” «· ·„Ì–", vbCritical
Exit Sub
End If
If Text4.Text = "" Then
MsgBox "«œŒ· —ﬁ„ ÊﬂÌ· «· ·„Ì–", vbCritical
Text4.SetFocus
Exit Sub
End If
Label11.Caption = ""
PicFilev = ""
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "Picture JPG |*.jpg|Picture Gif |*.Gif|Picture Bmp |*.Bmp|Picture Icon |*.ICO|All Picture |*.*"
    CommonDialog1.DialogTitle = "Picture"
    CommonDialog1.ShowOpen
    PicFilev = CommonDialog1.FileName 'lien d'image
    Image1.Picture = LoadPicture(PicFilev) 'Afficher l'image
 fName = CommonDialog1.FileName
 Label11.Caption = "01"
 'Image1.Width = 1900
 'Image1.Height = 1500
 'Command4.Enabled = True
 'Command2.Enabled = True
Exit Sub
p:
MsgBox "Êﬁ⁄ Œÿ√ ›Ì  Õ„Ì· «·ÊÀÌﬁ… , «·—Ã«¡ «⁄«œ… «·„Õ«Ê·…", vbExclamation
PicFilev = ""
Label11.Caption = ""
Image1.Picture = LoadPicture(PicFilev) 'Afficher l'image
End Sub

Private Sub Command30_Click()
On Error Resume Next
Image3.Picture = LoadPicture("")
grd7.Clear
grd7.Rows = 1
grd7.Cols = 4
grd7.ColWidth(0) = 0
grd7.ColWidth(1) = 1200
grd7.ColWidth(2) = 3500
grd7.ColWidth(3) = 1500
grd7.row = 0
grd7.Col = 1
grd7.Text = "«·—ﬁ„"
grd7.Col = 2
grd7.Text = "«·«”„"
grd7.Col = 3
grd7.Text = "«·—ﬁ„ «· ”·”·Ì"
grd7.ColAlignment(1) = 1
grd7.ColAlignment(2) = 1
grd7.ColAlignment(3) = 1

End Sub

Private Sub Command31_Click()
On Error Resume Next
Label11.Caption = ""
PicFilev = ""
Text5.Text = ""
Text2.Text = ""
Text20.Text = ""
Text19.Text = ""
Text3.Text = ""
Text4.Text = ""
Text1.Text = ""
Label11.Caption = ""
'Call chargec1
Call chargec2
PicFilev = ""
Label11.Caption = ""
Image1.Picture = LoadPicture(PicFilev) 'Afficher l'image
Text2.Enabled = True
Text20.Enabled = True
Text19.Enabled = True
Call numsetu
Text2.SetFocus
Call serial
Picture6.Visible = False
Command14.Enabled = False
Combo1.Enabled = True
Combo4.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
DT1.Enabled = True
Command3.Enabled = True
Command7.Enabled = True
Command5.Enabled = True
Command4.Enabled = True

End Sub

Private Sub Command32_Click()
On Error Resume Next
Dim i As Double
Dim tx1 As String
Dim tx2 As String
If Combo11.Text = "" Then
MsgBox "·« Ì„ﬂ‰  Ê“Ì⁄ «·— » Õ Ï Ì „  ÕœÌœ «·ﬁ”„", vbCritical
Exit Sub
End If
Command32.Enabled = False
i = 1
grd5.Clear
grd5.Rows = 1
Call cont2
grd5.Rows = be.RecordCount + 2
Do While Not be.EOF
If be!cla = Combo11.Text Then
grd5.row = i
grd5.Col = 0
grd5.Text = be!mtr
grd5.Col = 1
grd5.Text = be!nom
grd5.Col = 2
grd5.Text = be!moy
i = i + 1
End If
be.MoveNext
Loop
grd5.Rows = i
n = grd5.Rows
grd5.Col = 2
grd5.Sort = 2
For i = 1 To n - 1
grd5.row = i
grd5.Col = 0
tx = grd5.Text
grd5.Col = 3
grd5.Text = i
Next i
n = grd5.Rows
Call cont2
Do While Not be.EOF
If Combo11.Text = be!cla Then
For i = 1 To n - 1
grd5.row = i
grd5.Col = 0
tx1 = grd5.Text
grd5.Col = 3
tx2 = grd5.Text
If be!mtr = tx1 Then
be!ran = tx2
be.Update
i = n
End If
Next i
End If
be.MoveNext
Loop
Command32.Enabled = True
'Timer5.Enabled = True

End Sub

Private Sub Command33_Click()
On Error Resume Next
If Combo11.Text = "" Then
MsgBox "·« Ì„ﬂ‰ ⁄—÷ «·‰ «∆Ã Õ Ï Ì „  ÕœÌœ «·ﬁ”„", vbCritical
Exit Sub
End If
If Combo10.Text = "" Then
MsgBox "ÌÃ»  ÕœÌœ √Ì „«œ… ·Â–« «·ﬁ”„", vbCritical
Exit Sub
End If
Command33.Enabled = False
Command34.Enabled = False
Command2.Enabled = False
grd8.Visible = False
Call coffes
Call chargegrd8
grd8.Visible = True
Command33.Enabled = True
Command34.Enabled = True
Command2.Enabled = True
End Sub

Private Sub Command34_Click()
On Error Resume Next
Dim ane As String
Dim a As Double
Dim f As Double
Dim tx As String
Dim i As Double
Dim n As Double
Dim momat As Double
Dim comat As Double
Dim tomat As Double
Dim momats As Double
Dim comats As Double
Dim tomats As Double
Dim moy As Double
Dim cs As Double
Dim ts As Double
If Combo11.Text = "" Then
MsgBox "·« Ì„ﬂ‰ ⁄—÷ ﬂ‘› «·œ—Ã«  Õ Ï Ì „  ÕœÌœ «·ﬁ”„", vbCritical
Exit Sub
End If
If Combo10.Text = "" Then
MsgBox "ÌÃ»  ÕœÌœ √Ì „«œ… ·Â–« «·ﬁ”„", vbCritical
Exit Sub
End If
Command34.Enabled = False
'****** controle moyenne
Call coffes
Call cont2
Do While Not be.EOF
momats = 0
comats = 0
tomats = 0
momat = 0
comat = 0
tomat = 0
If Combo11.Text = be!cla Then
If Val(be!mm1) > 0 Or be!mm1 = "0" Then
momat = be!mm1
momats = momats + momat
comat = be!cof1
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm2) > 0 Or be!mm2 = "0" Then
momat = be!mm2
momats = momats + momat
comat = be!cof2
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm3) > 0 Or be!mm3 = "0" Then
momat = be!mm3
momats = momats + momat
comat = be!cof3
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm4) > 0 Or be!mm4 = "0" Then
momat = be!mm4
momats = momats + momat
comat = be!cof4
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm5) > 0 Or be!mm5 = "0" Then
momat = be!mm5
momats = momats + momat
comat = be!cof5
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm6) > 0 Or be!mm6 = "0" Then
momat = be!mm6
momats = momats + momat
comat = be!cof6
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm7) > 0 Or be!mm7 = "0" Then
momat = be!mm7
momats = momats + momat
comat = be!cof7
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm8) > 0 Or be!mm8 = "0" Then
momat = be!mm8
momats = momats + momat
comat = be!cof8
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm9) > 0 Or be!mm9 = "0" Then
momat = be!mm9
momats = momats + momat
comat = be!cof9
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm10) > 0 Or be!mm10 = "0" Then
momat = be!mm10
momats = momats + momat
comat = be!cof10
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm11) > 0 Or be!mm11 = "0" Then
momat = be!mm11
momats = momats + momat
comat = be!cof11
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm12) > 0 Or be!mm12 = "0" Then
momat = be!mm12
momats = momats + momat
comat = be!cof12
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm13) > 0 Or be!mm13 = "0" Then
momat = be!mm13
momats = momats + momat
comat = be!cof13
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm14) > 0 Or be!mm14 = "0" Then
momat = be!mm14
momats = momats + momat
comat = be!cof14
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm15) > 0 Or be!mm15 = "0" Then
momat = be!mm15
momats = momats + momat
comat = be!cof15
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm16) > 0 Or be!mm16 = "0" Then
momat = be!mm16
momats = momats + momat
comat = be!cof16
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm17) > 0 Or be!mm17 = "0" Then
momat = be!mm17
momats = momats + momat
comat = be!cof17
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm18) > 0 Or be!mm18 = "0" Then
momat = be!mm18
momats = momats + momat
comat = be!cof18
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm19) > 0 Or be!mm19 = "0" Then
momat = be!mm19
momats = momats + momat
comat = be!cof19
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!mm20) > 0 Or be!mm20 = "0" Then
momat = be!mm20
momats = momats + momat
comat = be!cof20
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
moy = 0
If comats > 0 Then
moy = tomats / comats
MyNumber = Round(moy, 2)
moy = MyNumber
End If
Label33.Caption = moy
a = 0
f = 6
'Mention
Call coffes
Label40.Caption = ""
Label29.Caption = ""
For i = 5 To 15
If moy = 0 And comats = 0 Then
Label40.Caption = ""
Label29.Caption = ""
Label33.Caption = ""
i = 15
'Exit Sub
End If
If moy = 0 And comats > 0 Then
Label40.Caption = mens(5).Text
Label29.Caption = mens(11).Text
i = 15
'Exit Sub
End If
cs = coff(i - 1).Text
ts = coff(i).Text
If moy > ts And moy <= cs Then
Label40.Caption = mens(a).Text
Label29.Caption = mens(f).Text
i = 15
End If
If i = 5 Or i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 15 Then
a = a + 1
f = f + 1
End If
Next i

If Check9.Value = 1 Then
be!dat = Date
Else
be!dat = ""
End If
If Check6.Value = 0 Then
be!ran = ""
End If
be!Abs = ""
be!mena = Label40.Caption
be!menf = Label29.Caption
be!moy = Label33.Caption
be!numm = Label106.Caption
be!eco = face.SBB1.Panels(13).Text
be!ann = face.SBB1.Panels(9).Text
be.Update
a = Label106.Caption
End If
be.MoveNext
Loop
If Check6.Value = 1 Then
Command32_Click
End If
Call cont2
i = Label41.Caption
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
data.DoCmd.Maximize
If i <= 7 Then
data.DoCmd.OpenReport "notes7", acViewPreview, , "numm =" & a, acWindowNormal, OpenArgs
ElseIf i > 7 And i <= 12 Then
data.DoCmd.OpenReport "notes7", acViewPreview, , "numm =" & a, acWindowNormal, OpenArgs
Else
data.DoCmd.OpenReport "Bulletin20", acViewPreview, , "numm =" & a, acWindowNormal, OpenArgs
End If
'data.DoCmd.RunMacro "Opnreport", 2
data.Visible = True
'Set data = Nothing
Command34.Enabled = True
Exit Sub

End Sub

Private Sub Command35_Click()
On Error Resume Next
Dim ane As String
Dim a As Double
Dim f As Double
Dim tx As String
Dim i As Double
Dim n As Double
Dim momat As Double
Dim comat As Double
Dim tomat As Double
Dim momats As Double
Dim comats As Double
Dim tomats As Double
Dim moy As Double
Dim b As Double
Dim moyenne As Double
Dim menar As String
Dim menfr As String
Dim cs As Double
Dim ts As Double
If Label33.Caption = "" Then
MsgBox "·« Ì„ﬂ‰ ⁄—÷ ﬂ‘› «·œ—Ã«  Õ Ï Ì „  ÕœÌœ «·„⁄œ· «·⁄«„", vbCritical
Exit Sub
End If
Call cont2
Do While Not be.EOF
If Label36.Caption = be!cla Then
momats = 0
comats = 0
tomats = 0
momat = 0
comat = 0
tomat = 0
If Val(be!e11) > 0 Or be!e11 = "0" Then
momat = be!e11
momats = momats + momat
comat = be!cof1
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e12) > 0 Or be!e12 = "0" Then
momat = be!e12
momats = momats + momat
comat = be!cof2
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e13) > 0 Or be!e13 = "0" Then
momat = be!e13
momats = momats + momat
comat = be!cof3
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e14) > 0 Or be!e14 = "0" Then
momat = be!e14
momats = momats + momat
comat = be!cof4
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e15) > 0 Or be!e15 = "0" Then
momat = be!e15
momats = momats + momat
comat = be!cof5
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e16) > 0 Or be!e16 = "0" Then
momat = be!e16
momats = momats + momat
comat = be!cof6
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e17) > 0 Or be!e17 = "0" Then
momat = be!e17
momats = momats + momat
comat = be!cof7
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e18) > 0 Or be!e18 = "0" Then
momat = be!e18
momats = momats + momat
comat = be!cof8
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e19) > 0 Or be!e19 = "0" Then
momat = be!e19
momats = momats + momat
comat = be!cof9
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e110) > 0 Or be!e110 = "0" Then
momat = be!e110
momats = momats + momat
comat = be!cof10
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e111) > 0 Or be!e111 = "0" Then
momat = be!e111
momats = momats + momat
comat = be!cof11
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e112) > 0 Or be!e112 = "0" Then
momat = be!e112
momats = momats + momat
comat = be!cof12
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e113) > 0 Or be!e113 = "0" Then
momat = be!e113
momats = momats + momat
comat = be!cof13
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e114) > 0 Or be!e114 = "0" Then
momat = be!e114
momats = momats + momat
comat = be!cof14
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e115) > 0 Or be!e115 = "0" Then
momat = be!e115
momats = momats + momat
comat = be!cof15
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e116) > 0 Or be!e116 = "0" Then
momat = be!e116
momats = momats + momat
comat = be!cof16
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e117) > 0 Or be!e117 = "0" Then
momat = be!e117
momats = momats + momat
comat = be!cof17
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e118) > 0 Or be!e118 = "0" Then
momat = be!e118
momats = momats + momat
comat = be!cof18
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e119) > 0 Or be!e119 = "0" Then
momat = be!e119
momats = momats + momat
comat = be!cof19
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e120) > 0 Or be!e120 = "0" Then
momat = be!e120
momats = momats + momat
comat = be!cof20
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
moy = 0
If comats > 0 Then
moy = tomats / comats
MyNumber = Round(moy, 2)
moy = MyNumber
End If
Label33.Caption = moy
a = 0
f = 6
'***** Mentions
Call coffes
Label40.Caption = ""
Label29.Caption = ""
For i = 5 To 15
If moy = 0 And comats = 0 Then
Label40.Caption = ""
Label29.Caption = ""
Label33.Caption = ""
i = 15
'Exit Sub
End If
If moy = 0 And comats > 0 Then
Label40.Caption = mens(5).Text
Label29.Caption = mens(11).Text
i = 15
'Exit Sub
End If
cs = coff(i - 1).Text
ts = coff(i).Text
If moy > ts And moy <= cs Then
Label40.Caption = mens(a).Text
Label29.Caption = mens(f).Text
i = 15
End If
If i = 5 Or i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 15 Then
a = a + 1
f = f + 1
End If
Next i
If Check4.Value = 1 Then
be!dat = Date
Else
be!dat = ""
End If
If Check5.Value = 1 Then
be!Abs = Label73.Caption
Else
be!Abs = ""
End If
If Check2.Value = 1 Then
be!ran = Label42.Caption
Else
be!ran = ""
End If
If Check3.Value = 1 Then
be!mena = Label40.Caption
be!menf = Label29.Caption
Else
be!mena = ""
be!menf = ""
End If
If Label37.Caption = be!mtr Then
be!nom = Label38.Caption
be!ser = BarcodeX2.Caption
moyenne = Label33.Caption
menar = Label40.Caption
menfr = Label29.Caption
b = be!num
End If
be!numm = Label106.Caption
be!eco = face.SBB1.Panels(13).Text
be!ann = face.SBB1.Panels(9).Text
be!moy = Label33.Caption
be.Update
'be.MoveLast
End If
be.MoveNext
Loop
Label33.Caption = moyenne
Label40.Caption = menar
Label29.Caption = menfr
If Check2.Value = 1 Then
Command21_Click
End If
Call cont2
i = Label41.Caption
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
data.DoCmd.Maximize
If i <= 7 Then
data.DoCmd.OpenReport "Bulletin107", acViewPreview, , "num =" & b, acWindowNormal, OpenArgs
ElseIf i > 7 And i <= 13 Then
data.DoCmd.OpenReport "Bulletin113", acViewPreview, , "num =" & b, acWindowNormal, OpenArgs
Else
data.DoCmd.OpenReport "Bulletin120", acViewPreview, , "num =" & b, acWindowNormal, OpenArgs
End If
'data.DoCmd.RunMacro "Opnreport", 2
data.Visible = True
'Set data = Nothing

End Sub

Private Sub Command36_Click()
On Error Resume Next
Call cont
Combo11.Clear
  Do While Not cl.EOF
  If cl!act = "1" Then
    Combo11.AddItem cl!cla
    End If
cl.MoveNext
  Loop
SSTab3.Visible = False
grd6.Visible = False
grd8.Clear
grd8.Rows = 1
Label95.Caption = ""
Label96.Caption = ""
grd7.Clear
grd7.Rows = 1
grd7.Cols = 4
grd7.ColWidth(0) = 0
grd7.ColWidth(1) = 1200
grd7.ColWidth(2) = 3500
grd7.ColWidth(3) = 1500
grd7.row = 0
grd7.Col = 1
grd7.Text = "«·—ﬁ„"
grd7.Col = 2
grd7.Text = "«·«”„"
grd7.Col = 3
grd7.Text = "«·—ﬁ„ «· ”·”·Ì"
grd7.ColAlignment(1) = 1
grd7.ColAlignment(2) = 1
grd7.ColAlignment(3) = 1

End Sub

Private Sub Command37_Click()
On Error Resume Next
Dim ane As String
Dim a As Double
Dim f As Double
Dim tx As String
Dim i As Double
Dim n As Double
Dim momat As Double
Dim comat As Double
Dim tomat As Double
Dim momats As Double
Dim comats As Double
Dim tomats As Double
Dim moy As Double
Dim cs As Double
Dim ts As Double
Dim b As Double
Dim moyenne As Double
Dim menar As String
Dim menfr As String
If Label33.Caption = "" Then
MsgBox "·« Ì„ﬂ‰ ⁄—÷ ﬂ‘› «·œ—Ã«  Õ Ï Ì „  ÕœÌœ «·„⁄œ· «·⁄«„", vbCritical
Exit Sub
End If
Call cont2
Do While Not be.EOF
If Label36.Caption = be!cla Then
momats = 0
comats = 0
tomats = 0
momat = 0
comat = 0
tomat = 0
If Val(be!e21) > 0 Or be!e21 = "0" Then
momat = be!e21
momats = momats + momat
comat = be!cof1
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e22) > 0 Or be!e22 = "0" Then
momat = be!e22
momats = momats + momat
comat = be!cof2
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e23) > 0 Or be!e23 = "0" Then
momat = be!e23
momats = momats + momat
comat = be!cof3
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e24) > 0 Or be!e24 = "0" Then
momat = be!e24
momats = momats + momat
comat = be!cof4
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e25) > 0 Or be!e25 = "0" Then
momat = be!e25
momats = momats + momat
comat = be!cof5
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e26) > 0 Or be!e26 = "0" Then
momat = be!e26
momats = momats + momat
comat = be!cof6
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e27) > 0 Or be!e27 = "0" Then
momat = be!e27
momats = momats + momat
comat = be!cof7
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e28) > 0 Or be!e28 = "0" Then
momat = be!e28
momats = momats + momat
comat = be!cof8
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e29) > 0 Or be!e29 = "0" Then
momat = be!e29
momats = momats + momat
comat = be!cof9
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e210) > 0 Or be!e210 = "0" Then
momat = be!e210
momats = momats + momat
comat = be!cof10
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e211) > 0 Or be!e211 = "0" Then
momat = be!e211
momats = momats + momat
comat = be!cof11
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e212) > 0 Or be!e212 = "0" Then
momat = be!e212
momats = momats + momat
comat = be!cof12
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e213) > 0 Or be!e213 = "0" Then
momat = be!e213
momats = momats + momat
comat = be!cof13
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e214) > 0 Or be!e214 = "0" Then
momat = be!e214
momats = momats + momat
comat = be!cof14
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e215) > 0 Or be!e215 = "0" Then
momat = be!e215
momats = momats + momat
comat = be!cof15
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e216) > 0 Or be!e216 = "0" Then
momat = be!e216
momats = momats + momat
comat = be!cof16
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e217) > 0 Or be!e217 = "0" Then
momat = be!e217
momats = momats + momat
comat = be!cof17
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e218) > 0 Or be!e218 = "0" Then
momat = be!e218
momats = momats + momat
comat = be!cof18
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e219) > 0 Or be!e219 = "0" Then
momat = be!e219
momats = momats + momat
comat = be!cof19
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e220) > 0 Or be!e220 = "0" Then
momat = be!e220
momats = momats + momat
comat = be!cof20
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
moy = 0
If comats > 0 Then
moy = tomats / comats
MyNumber = Round(moy, 2)
moy = MyNumber
End If
Label33.Caption = moy
a = 0
f = 6
'***** Mentions
Call coffes
Label40.Caption = ""
Label29.Caption = ""
For i = 5 To 15
If moy = 0 And comats = 0 Then
Label40.Caption = ""
Label29.Caption = ""
Label33.Caption = ""
i = 15
'Exit Sub
End If
If moy = 0 And comats > 0 Then
Label40.Caption = mens(5).Text
Label29.Caption = mens(11).Text
i = 15
'Exit Sub
End If
cs = coff(i - 1).Text
ts = coff(i).Text
If moy > ts And moy <= cs Then
Label40.Caption = mens(a).Text
Label29.Caption = mens(f).Text
i = 15
End If
If i = 5 Or i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 15 Then
a = a + 1
f = f + 1
End If
Next i
If Check4.Value = 1 Then
be!dat = Date
Else
be!dat = ""
End If
If Check5.Value = 1 Then
be!Abs = Label73.Caption
Else
be!Abs = ""
End If
If Check2.Value = 1 Then
be!ran = Label42.Caption
Else
be!ran = ""
End If
If Check3.Value = 1 Then
be!mena = Label40.Caption
be!menf = Label29.Caption
Else
be!mena = ""
be!menf = ""
End If
If Label37.Caption = be!mtr Then
be!nom = Label38.Caption
be!ser = BarcodeX2.Caption
moyenne = Label33.Caption
menar = Label40.Caption
menfr = Label29.Caption
b = be!num
End If
be!numm = Label106.Caption
be!eco = face.SBB1.Panels(13).Text
be!ann = face.SBB1.Panels(9).Text
be!moy = Label33.Caption
be.Update
'be.MoveLast
End If
be.MoveNext
Loop
Label33.Caption = moyenne
Label40.Caption = menar
Label29.Caption = menfr
If Check2.Value = 1 Then
Command21_Click
End If
Call cont2
i = Label41.Caption
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
data.DoCmd.Maximize
If i <= 7 Then
data.DoCmd.OpenReport "Bulletin207", acViewPreview, , "num =" & b, acWindowNormal, OpenArgs
ElseIf i > 7 And i <= 13 Then
data.DoCmd.OpenReport "Bulletin213", acViewPreview, , "num =" & b, acWindowNormal, OpenArgs
Else
data.DoCmd.OpenReport "Bulletin220", acViewPreview, , "num =" & b, acWindowNormal, OpenArgs
End If
'data.DoCmd.RunMacro "Opnreport", 2
data.Visible = True
'Set data = Nothing

End Sub

Private Sub Command38_Click()
On Error Resume Next
Dim ane As String
Dim a As Double
Dim f As Double
Dim tx As String
Dim i As Double
Dim n As Double
Dim momat As Double
Dim comat As Double
Dim tomat As Double
Dim momats As Double
Dim comats As Double
Dim tomats As Double
Dim moy As Double
Dim cs As Double
Dim ts As Double
If Combo11.Text = "" Then
MsgBox "·« Ì„ﬂ‰ ⁄—÷ ﬂ‘› «·œ—Ã«  Õ Ï Ì „  ÕœÌœ «·ﬁ”„", vbCritical
Exit Sub
End If
If Combo10.Text = "" Then
MsgBox "ÌÃ»  ÕœÌœ √Ì „«œ… ·Â–« «·ﬁ”„", vbCritical
Exit Sub
End If
Command38.Enabled = False
'****** controle moyenne
Call coffes
Call cont2
Do While Not be.EOF
momats = 0
comats = 0
tomats = 0
momat = 0
comat = 0
tomat = 0
If Combo11.Text = be!cla Then
If Val(be!e11) > 0 Or be!e11 = "0" Then
momat = be!e11
momats = momats + momat
comat = be!cof1
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e12) > 0 Or be!e12 = "0" Then
momat = be!e12
momats = momats + momat
comat = be!cof2
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e13) > 0 Or be!e13 = "0" Then
momat = be!e13
momats = momats + momat
comat = be!cof3
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e14) > 0 Or be!e14 = "0" Then
momat = be!e14
momats = momats + momat
comat = be!cof4
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e15) > 0 Or be!e15 = "0" Then
momat = be!e15
momats = momats + momat
comat = be!cof5
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e16) > 0 Or be!e16 = "0" Then
momat = be!e16
momats = momats + momat
comat = be!cof6
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e17) > 0 Or be!e17 = "0" Then
momat = be!e17
momats = momats + momat
comat = be!cof7
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e18) > 0 Or be!e18 = "0" Then
momat = be!e18
momats = momats + momat
comat = be!cof8
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e19) > 0 Or be!e19 = "0" Then
momat = be!e19
momats = momats + momat
comat = be!cof9
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e110) > 0 Or be!e110 = "0" Then
momat = be!e110
momats = momats + momat
comat = be!cof10
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e111) > 0 Or be!e111 = "0" Then
momat = be!e111
momats = momats + momat
comat = be!cof11
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e112) > 0 Or be!e112 = "0" Then
momat = be!e112
momats = momats + momat
comat = be!cof12
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e113) > 0 Or be!e113 = "0" Then
momat = be!e113
momats = momats + momat
comat = be!cof13
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e114) > 0 Or be!e114 = "0" Then
momat = be!e114
momats = momats + momat
comat = be!cof14
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e115) > 0 Or be!e115 = "0" Then
momat = be!e115
momats = momats + momat
comat = be!cof15
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e116) > 0 Or be!e116 = "0" Then
momat = be!e116
momats = momats + momat
comat = be!cof16
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e117) > 0 Or be!e117 = "0" Then
momat = be!e117
momats = momats + momat
comat = be!cof17
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e118) > 0 Or be!e118 = "0" Then
momat = be!e118
momats = momats + momat
comat = be!cof18
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e119) > 0 Or be!e119 = "0" Then
momat = be!e119
momats = momats + momat
comat = be!cof19
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e120) > 0 Or be!e120 = "0" Then
momat = be!e120
momats = momats + momat
comat = be!cof20
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
moy = 0
If comats > 0 Then
moy = tomats / comats
MyNumber = Round(moy, 2)
moy = MyNumber
End If
Label33.Caption = moy
a = 0
f = 6
'mention
Call coffes
Label40.Caption = ""
Label29.Caption = ""
For i = 5 To 15
If moy = 0 And comats = 0 Then
Label40.Caption = ""
Label29.Caption = ""
Label33.Caption = ""
i = 15
'Exit Sub
End If
If moy = 0 And comats > 0 Then
Label40.Caption = mens(5).Text
Label29.Caption = mens(11).Text
i = 15
'Exit Sub
End If
cs = coff(i - 1).Text
ts = coff(i).Text
If moy > ts And moy <= cs Then
Label40.Caption = mens(a).Text
Label29.Caption = mens(f).Text
i = 15
End If
If i = 5 Or i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 15 Then
a = a + 1
f = f + 1
End If
Next i

If Check9.Value = 1 Then
be!dat = Date
Else
be!dat = ""
End If
If Check6.Value = 0 Then
be!ran = ""
End If
be!Abs = ""
be!mena = Label40.Caption
be!menf = Label29.Caption
be!moy = Label33.Caption
be!numm = Label106.Caption
be!eco = face.SBB1.Panels(13).Text
be!ann = face.SBB1.Panels(9).Text
be.Update
a = Label106.Caption
End If
be.MoveNext
Loop
If Check6.Value = 1 Then
Command32_Click
End If
Call cont2
i = Label41.Caption
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
data.DoCmd.Maximize
If i <= 7 Then
data.DoCmd.OpenReport "Bulletin107", acViewPreview, , "numm =" & a, acWindowNormal, OpenArgs
ElseIf i > 7 And i <= 13 Then
data.DoCmd.OpenReport "Bulletin113", acViewPreview, , "numm =" & a, acWindowNormal, OpenArgs
Else
data.DoCmd.OpenReport "Bulletin120", acViewPreview, , "numm =" & a, acWindowNormal, OpenArgs
End If
'data.DoCmd.RunMacro "Opnreport", 2
data.Visible = True
'Set data = Nothing
Command38.Enabled = True
Exit Sub

End Sub

Private Sub Command39_Click()
On Error Resume Next
Dim ane As String
Dim a As Double
Dim f As Double
Dim tx As String
Dim i As Double
Dim n As Double
Dim momat As Double
Dim comat As Double
Dim tomat As Double
Dim momats As Double
Dim comats As Double
Dim tomats As Double
Dim moy As Double
Dim cs As Double
Dim ts As Double
If Combo11.Text = "" Then
MsgBox "·« Ì„ﬂ‰ ⁄—÷ ﬂ‘› «·œ—Ã«  Õ Ï Ì „  ÕœÌœ «·ﬁ”„", vbCritical
Exit Sub
End If
If Combo10.Text = "" Then
MsgBox "ÌÃ»  ÕœÌœ √Ì „«œ… ·Â–« «·ﬁ”„", vbCritical
Exit Sub
End If
'Command39.Enabled = False
'****** controle moyenne
Call coffes
Call cont2
Do While Not be.EOF
momats = 0
comats = 0
tomats = 0
momat = 0
comat = 0
tomat = 0
If Combo11.Text = be!cla Then
If Val(be!e21) > 0 Or be!e21 = "0" Then
momat = be!e21
momats = momats + momat
comat = be!cof1
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e22) > 0 Or be!e22 = "0" Then
momat = be!e22
momats = momats + momat
comat = be!cof2
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e23) > 0 Or be!e23 = "0" Then
momat = be!e23
momats = momats + momat
comat = be!cof3
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e24) > 0 Or be!e24 = "0" Then
momat = be!e24
momats = momats + momat
comat = be!cof4
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e25) > 0 Or be!e25 = "0" Then
momat = be!e25
momats = momats + momat
comat = be!cof5
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e26) > 0 Or be!e26 = "0" Then
momat = be!e26
momats = momats + momat
comat = be!cof6
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e27) > 0 Or be!e27 = "0" Then
momat = be!e27
momats = momats + momat
comat = be!cof7
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e28) > 0 Or be!e28 = "0" Then
momat = be!e28
momats = momats + momat
comat = be!cof8
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e29) > 0 Or be!e29 = "0" Then
momat = be!e29
momats = momats + momat
comat = be!cof9
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e210) > 0 Or be!e210 = "0" Then
momat = be!e210
momats = momats + momat
comat = be!cof10
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e211) > 0 Or be!e211 = "0" Then
momat = be!e211
momats = momats + momat
comat = be!cof11
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e212) > 0 Or be!e212 = "0" Then
momat = be!e212
momats = momats + momat
comat = be!cof12
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e213) > 0 Or be!e213 = "0" Then
momat = be!e213
momats = momats + momat
comat = be!cof13
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e214) > 0 Or be!e214 = "0" Then
momat = be!e214
momats = momats + momat
comat = be!cof14
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e215) > 0 Or be!e215 = "0" Then
momat = be!e215
momats = momats + momat
comat = be!cof15
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e216) > 0 Or be!e216 = "0" Then
momat = be!e216
momats = momats + momat
comat = be!cof16
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e217) > 0 Or be!e217 = "0" Then
momat = be!e217
momats = momats + momat
comat = be!cof17
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e218) > 0 Or be!e218 = "0" Then
momat = be!e218
momats = momats + momat
comat = be!cof18
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e219) > 0 Or be!e219 = "0" Then
momat = be!e219
momats = momats + momat
comat = be!cof19
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
If Val(be!e220) > 0 Or be!e220 = "0" Then
momat = be!e220
momats = momats + momat
comat = be!cof20
comats = comats + comat
tomat = momat * comat
tomats = tomats + tomat
End If
moy = 0
If comats > 0 Then
moy = tomats / comats
MyNumber = Round(moy, 2)
moy = MyNumber
End If
Label33.Caption = moy
a = 0
f = 6
'Mention
Call coffes
Label40.Caption = ""
Label29.Caption = ""
For i = 5 To 15
If moy = 0 And comats = 0 Then
Label40.Caption = ""
Label29.Caption = ""
Label33.Caption = ""
i = 15
'Exit Sub
End If
If moy = 0 And comats > 0 Then
Label40.Caption = mens(5).Text
Label29.Caption = mens(11).Text
i = 15
'Exit Sub
End If
cs = coff(i - 1).Text
ts = coff(i).Text
If moy > ts And moy <= cs Then
Label40.Caption = mens(a).Text
Label29.Caption = mens(f).Text
i = 15
End If
If i = 5 Or i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 15 Then
a = a + 1
f = f + 1
End If
Next i

If Check9.Value = 1 Then
be!dat = Date
Else
be!dat = ""
End If
If Check6.Value = 0 Then
be!ran = ""
End If
be!Abs = ""
be!mena = Label40.Caption
be!menf = Label29.Caption
be!moy = Label33.Caption
be!numm = Label106.Caption
be!eco = face.SBB1.Panels(13).Text
be!ann = face.SBB1.Panels(9).Text
be.Update
a = Label106.Caption
End If
be.MoveNext
Loop
If Check6.Value = 1 Then
Command32_Click
End If
Call cont2
i = Label41.Caption
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
data.DoCmd.Maximize
If i <= 7 Then
data.DoCmd.OpenReport "Bulletin207", acViewPreview, , "numm =" & a, acWindowNormal, OpenArgs
ElseIf i > 7 And i <= 13 Then
data.DoCmd.OpenReport "Bulletin213", acViewPreview, , "numm =" & a, acWindowNormal, OpenArgs
Else
data.DoCmd.OpenReport "Bulletin220", acViewPreview, , "numm =" & a, acWindowNormal, OpenArgs
End If
'data.DoCmd.RunMacro "Opnreport", 2
data.Visible = True
'Set data = Nothing
Command39.Enabled = True

End Sub

Private Sub Command4_Click()
On Error Resume Next
Text5.Text = ""
Text2.Text = ""
Text20.Text = ""
Text19.Text = ""
Text3.Text = ""
Text4.Text = ""
Label11.Caption = ""
Text2.SetFocus
PicFilev = ""
Label11.Caption = ""
Image1.Picture = LoadPicture(PicFilev) 'Afficher l'image
If Combo1.Text <> "" Then
Call numsetu
End If
Call serial
ProgressBar2.Value = 0
Timer2.Enabled = False

End Sub

Private Sub Command40_Click()
On Error GoTo u
Dim a As Double
Dim b As Double
Dim c As Double
Dim r As Double
Dim m As String
Dim t As String
Dim i As Double
Dim j As Double
Dim n As String
Dim k As Double
Dim h As Double
Dim p As Double
If Combo11.Text = "" Then
MsgBox "ÌÃ»  ÕœÌœ «·ﬁ”„ √Ê·«", vbCritical
Exit Sub
End If
Command40.Enabled = False
grd10.Visible = False
FileCopy App.Path & "\Export_Notes.xls", App.Path & "\Notes_Etudiants.xls"
Set kb = CreateObject("Excel.application")
kb.Workbooks.Open (App.Path & "\Notes_Etudiants.xls")
For k = 1 To 1
grd10.Clear
grd10.Rows = 1
'matiers
i = 2
grd10.row = 0
grd10.Col = 0
grd10.Text = "«·ﬂÊœ"
grd10.Col = 1
grd10.Text = "«·«”„ «·ﬂ«„·"
Call cont
grd10.Cols = mt.RecordCount + 2
Do While Not mt.EOF
If Combo11.Text = mt!cla Then
grd10.row = 0
grd10.Col = i
grd10.Text = mt!mat
i = i + 1
End If
mt.MoveNext
Loop
grd10.Cols = i
'noms
i = 1
et.MoveFirst
grd10.Rows = et.RecordCount + 2
Do While Not et.EOF
If Combo11.Text = et!cla Then
grd10.row = i
grd10.Col = 0
grd10.Text = et!ser
grd10.Col = 1
grd10.Text = et!nom
i = i + 1
End If
et.MoveNext
Loop
grd10.Rows = i
grd10.Col = 0
grd10.Sort = 1
'notes
i = 1
c = grd10.Cols
r = grd10.Rows
nt.MoveFirst
Do While Not nt.EOF
If Combo11.Text = nt!cla Then
a = nt!ser
m = nt!mat
If k = 1 Then
n = nt!de1
ElseIf k = 2 Then
n = nt!de2
ElseIf k = 3 Then
n = nt!de3
ElseIf k = 4 Then
n = nt!de4
ElseIf k = 5 Then
n = nt!de5
ElseIf k = 6 Then
n = nt!de6
ElseIf k = 7 Then
n = nt!de7
ElseIf k = 8 Then
n = nt!de8
ElseIf k = 9 Then
n = nt!de9
ElseIf k = 10 Then
n = nt!de10
ElseIf k = 11 Then
n = nt!ex1
ElseIf k = 12 Then
n = nt!ex2
ElseIf k = 13 Then
n = nt!ex3
End If
For i = 1 To r - 1
grd10.row = i
grd10.Col = 0
b = grd10.Text
If a = b Then
For j = 2 To c - 1
grd10.Col = j
grd10.row = 0
t = grd10.Text
If t = m Then
grd10.Col = j
grd10.row = i
grd10.Text = n
i = r
j = c
End If
Next j
End If
Next i
End If
nt.MoveNext
Loop
h = 100 * k / 2
ProgressBar5.Value = h
'Excel
For i = 0 To r - 1
For j = 0 To c - 1
grd10.row = i
grd10.Col = j
p = c - j
kb.Workbooks("Notes_Etudiants").Sheets(k).Cells(i + 3, j + 1).Value = grd10.Text
kb.Workbooks("Notes_Etudiants").Sheets(k).Range("B1").Value = Combo11.Text
Next j
Next i
Next k
kb.Visible = True
grd10.Visible = True
grd10.Clear
grd10.Rows = 1
grd10.Cols = 4
ProgressBar5.Value = 0
Command40.Enabled = True
Exit Sub
u:
MsgBox "ÌÃ» «€·«ﬁ ’›Õ… «ﬂ”· «‰ ﬂ«‰  „› ÊÕ…", vbExclamation
Command40.Enabled = True
End Sub

Private Sub Command41_Click()
On Error Resume Next
Dim i As Double
Dim a As String
Dim b As String
If Combo11.Text = "" Then
MsgBox "ÌÃ»  ÕœÌœ «·ﬁ”„ √Ê·«", vbCritical
Exit Sub
End If
Call cont2
Do While Not be.EOF
b = be!ran
a = Val(be!ran)
If a < 10 Then
b = "00" + a
ElseIf a < 100 And a > 9 Then
b = "0" + a
End If
be!ran = b
be.Update
be.MoveNext
Loop
Call cont2
i = Label106.Caption
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
data.DoCmd.Maximize
data.DoCmd.OpenReport "Tmoyennes", acViewPreview, , "numm =" & i, acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
data.Visible = True

End Sub

Private Sub Command42_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim c As Double
Dim r As Double
c = grd1.Cols
r = grd1.Rows
FileCopy App.Path & "\Notes_etu0.xls", App.Path & "\Notes_Etudiant.xls"
Set kb = CreateObject("Excel.application")
kb.Workbooks.Open (App.Path & "\Notes_Etudiant.xls")
For i = 0 To r - 1
For j = 1 To c - 1
grd1.row = i
grd1.Col = j
kb.Workbooks("Notes_Etudiant").Sheets(1).Cells(i + 3, j).Value = grd1.Text
Next j
Next i
kb.Workbooks("Notes_Etudiant").Sheets(1).Range("Q1").Value = Text1.Text
kb.Workbooks("Notes_Etudiant").Sheets(1).Range("O1").Value = Label36.Caption
kb.Workbooks("Notes_Etudiant").Sheets(1).Range("M1").Value = Label37.Caption
kb.Workbooks("Notes_Etudiant").Sheets(1).Range("I1").Value = Label38.Caption
kb.Workbooks("Notes_Etudiant").Sheets(1).Range("G1").Value = Label73.Caption
kb.Workbooks("Notes_Etudiant").Sheets(1).Range("E1").Value = Label42.Caption
kb.Workbooks("Notes_Etudiant").Sheets(1).Range("C1").Value = Label33.Caption
kb.Workbooks("Notes_Etudiant").Sheets(1).Range("A1").Value = Label40.Caption
kb.Visible = True
End Sub





Private Sub Command43_Click()
On Error Resume Next
Dim r As Double
Dim i As Double
Dim j  As Double
Command43.Enabled = False
grd25.Visible = False
Call chargegrd25_2
grd25.Visible = True
r = grd25.Rows
FileCopy App.Path & "\Export_tel.xls", App.Path & "\TEL.xls"
Set kb = CreateObject("Excel.application")
kb.Workbooks.Open (App.Path & "\TEL.xls")
kb.Workbooks("TEL").Sheets(1).Cells(3, 5).Value = Combo5.Text
For i = 1 To r - 1
For j = 0 To 5
grd25.row = i
grd25.Col = (5 - j)
kb.Workbooks("TEL").Sheets(1).Cells(i + 4, j + 1).Value = grd25.Text
Next j
Next i
kb.Visible = True
Command43.Enabled = True
End Sub

Private Sub Command44_Click()
On Error Resume Next
Dim r As Double
Dim i As Double
Dim j  As Double
Command44.Enabled = False
grd25.Visible = False
Call chargegrd25_2
grd25.Visible = True
r = grd25.Rows
FileCopy App.Path & "\Export_nni.xls", App.Path & "\NNI.xls"
Set kb = CreateObject("Excel.application")
kb.Workbooks.Open (App.Path & "\NNI.xls")
kb.Workbooks("NNI").Sheets(1).Cells(3, 5).Value = Combo5.Text
For i = 1 To r - 1
For j = 0 To 5
grd25.row = i
grd25.Col = (5 - j)
kb.Workbooks("NNI").Sheets(1).Cells(i + 4, j + 1).Value = grd25.Text
Next j
Next i
kb.Visible = True

''Dim a As Double
''Command44.Enabled = False
''a = Label7.Caption
''Call cont2
''ane = "C" + face.SBB1.Panels(9).Text
''data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
''data.DoCmd.Maximize
''data.DoCmd.OpenReport "Tetudiantsnni", acViewPreview, , "ncla =" & a, acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
''data.Visible = True
'Set data = Nothing
Command44.Enabled = True
End Sub

Private Sub Command45_Click()
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

If grd2.Rows < 2 Then
MsgBox "·«  ÊÃœ »Ì«‰« ", vbCritical
Exit Sub
End If
cla = Label81.Caption
nom = Label76.Caption
FileCopy App.Path & "\Abs010.xls", App.Path & "\Absences.xls"
Command45.Enabled = False
n = grd2.Rows
Set kb = CreateObject("Excel.application")
kb.Workbooks.Open (App.Path & "\Absences.xls")
kb.Visible = True
For i = 0 To n - 1
For j = 1 To 6
grd2.row = i
grd2.Col = j
k = 7 - j
kb.Workbooks("Absences").Sheets(1).Cells(i + 7, k).Value = grd2.Text
Next j
Next i

kb.Workbooks("Absences").Sheets(1).Range("D3").Value = face.SBB1.Panels(13).Text
kb.Workbooks("Absences").Sheets(1).Range("A3").Value = face.SBB1.Panels(9).Text
kb.Workbooks("Absences").Sheets(1).Range("D5").Value = nom
kb.Workbooks("Absences").Sheets(1).Range("A5").Value = cla

'kb.Workbooks("Historique de compte").Sheets(1).Cells(k + 2, 2).Value = "«·≈œ«—…"

'kb.Workbooks("fiche de presences").Sheets(1).Range("B5").Value = DT11.Value
'Workbooks("Etudiants").Close savechanges:=False
'Worksheets(1).Activate
Command45.Enabled = True
Exit Sub
u:
MsgBox "ÌÃ» «€·«ﬁ ’›Õ… «ﬂ”· «‰ ﬂ«‰  „› ÊÕ…", vbExclamation
Command45.Enabled = True

End Sub

Private Sub Command5_Click()
On Error Resume Next
Dim a As Double
'Dim nni As Double
'Dim md As Double
Text2.Text = Trim(Text2.Text)
Text20.Text = Trim(Text20.Text)
Text3.Text = Trim(Text3.Text)
Text4.Text = Trim(Text4.Text)
Text5.Text = Trim(Text5.Text)
Text19.Text = Trim(Text19.Text)
If Combo1.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·ﬁ”„", vbCritical
Exit Sub
End If
If Text3.Text = "" Then
MsgBox "«œŒ· —ﬁ„ «· ·„Ì–", vbCritical
Text3.SetFocus
Exit Sub
End If
If Val(Text3.Text) = 0 Then
MsgBox "—ﬁ„ «· ·„Ì– «·„œŒ· €Ì— ”·Ì„", vbCritical
Exit Sub
End If
If Val(Text3.Text) > 999999 Then
MsgBox "ÌÃ» √‰ ·«Ì Ã«Ê“ —ﬁ„ «· ·„Ì– ”  ‘›—« ", vbCritical
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "«œŒ· «”„ «· ·„Ì–", vbCritical
Text2.SetFocus
Exit Sub
End If
If Combo4.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— Ã‰” «· ·„Ì–", vbCritical
Exit Sub
End If
If Text4.Text = "" Then
MsgBox "«œŒ· —ﬁ„ ÊﬂÌ· «· ·„Ì–", vbCritical
Text4.SetFocus
Exit Sub
End If
'Control NNI
'If Text5.Text <> "" Then
'nni = Text5.Text
'md = Round((nni - 1) / 97, 2)
'If md <> 1 Then
'MsgBox "«·—ﬁ„ «·Êÿ‰Ì «·„œŒ· €Ì— ’ÕÌÕ", vbCritical
'MsgBox nni, vbCritical
'MsgBox md, vbCritical
'Text5.SetFocus
'Exit Sub
'End If
'End If

Call cont
Do While Not et.EOF
If et!cla = Combo1.Text And et!num = Text3.Text Then
MsgBox "€Ì— „„ﬂ‰ .. ÌÊÃœ  ·„Ì– ¬Œ— ÌÕ„· ‰›” «·—ﬁ„ ›Ì ‰›” «·ﬁ”„", vbCritical
Exit Sub
End If
et.MoveNext
Loop
If Label11.Caption = "" Then
g = MsgBox("·„ Ì „ «—›«ﬁ ’Ê—… ·· ·„Ì– , Â·  —Ìœ «·«” „—«—ø", vbInformation + vbYesNo, "Pressing")
If g = vbYes Then
Label11.Caption = "01"
If Combo4.Text = "√‰ÀÏ" Then
PicFilev = App.Path & "\nophotof.jpg"
Else
PicFilev = App.Path & "\nophotom.jpg"
End If
Image1.Picture = LoadPicture(PicFilev)
fName = PicFilev
Else
Exit Sub
End If
End If
Combo1.Enabled = False
Combo4.Enabled = False
Text3.Enabled = False
Text2.Enabled = False
Text20.Enabled = False
Text19.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Command3.Enabled = False
Command7.Enabled = False
Command5.Enabled = False
Command4.Enabled = False
DT1.Enabled = False
et.AddNew
et!cla = Combo1.Text
et!num = Text3.Text
et!dat = DT1.Value
et!nom = Text2.Text
et!nof = Text19.Text
et!rim = Text20.Text
et!sex = Combo4.Text
et!pho = "01"
et!tel = Text4.Text
et!adr = Text5.Text
et!ser = BarcodeX1.Caption
et!act = "1"
et.Update
Call cont
a = sr!num
sr!num = a + 1
sr.Update
Call cont2
nn.AddNew
nn!ser = BarcodeX1.Caption
nn!nni = Text5.Text
nn!liu = Text15.Text
nn!dat = Text14.Text
nn!cla = Combo1.Text
nn!num = Text3.Text
nn!dti = DT1.Value
nn!nom = Text2.Text
nn!nof = Text19.Text
nn!rim = Text20.Text
nn!sex = Combo4.Text
nn!tel = Text4.Text
nn!tel = Text4.Text
nn!ncla = Label4.Caption
nn.Update
Timer2.Enabled = True
Command14.Enabled = True
End Sub

Private Sub Command6_Click()
On Error Resume Next
Picture4.Visible = False
End Sub

Private Sub Command7_Click()
On Error Resume Next
PicFilev = ""
Label11.Caption = ""
Image1.Picture = LoadPicture(PicFilev) 'Afficher l'image
fName = ""
End Sub

Private Sub Command8_Click()
On Error Resume Next
Dim t As String
Dim a As Integer
Dim cl1 As String
Dim cl2 As String
a = 0
SSTab2.Tab = 2
Call cont
Do While Not et.EOF
If Val(et!num) < 1000000 Then
If Val(et!act) = 1 Then
If Text1.Text = et!ser Or Val(Text1.Text) = Val(et!ser) Then
a = 1
Label21.Caption = et!aut
Text7.Text = et!nom
Text18.Text = et!nof
Text21.Text = et!rim
Text17.Text = et!rim
Label36.Caption = et!cla
Label37.Caption = et!num
Label38.Caption = et!nom
Label81.Caption = et!cla
Label80.Caption = et!num
Label76.Caption = et!nom
Combo3.Text = et!sex
Text8.Text = et!tel
DT2.Value = et!dat
Text9.Text = et!adr
t = et!num
BarcodeX2.Caption = et!ser
Label88.Caption = et!ser
cl1 = et!cla
cl2 = et!cla
et.MoveLast
End If
End If
End If
et.MoveNext
Loop
Call cont2
Do While Not nn.EOF
If Text9.Text = nn!nni Then
Text13.Text = nn!liu
Text16.Text = nn!dat
nn.MoveLast
End If
nn.MoveNext
Loop
If a = 1 Then
Call LoadPictureFromDB
Combo2.Text = cl1
Text6.Text = t
Picture4.Visible = False
Picture6.Visible = True
SSTab2.Visible = True
Else
MsgBox "«·—ﬁ„ «· ”·”·Ì «·„œŒ· €Ì— „Œ“‰ .. Ì—ÃÏ «· √ﬂœ „‰Â", vbExclamation
Text1.Text = ""
Text1.SetFocus
End If

End Sub

Private Sub Command9_Click()
On Error Resume Next
Dim a As Double
Dim x As Double
'Dim nni As Double
'Dim md As Double
Text6.Text = Trim(Text6.Text)
Text7.Text = Trim(Text7.Text)
Text21.Text = Trim(Text21.Text)
Text18.Text = Trim(Text18.Text)
Text8.Text = Trim(Text8.Text)
Text9.Text = Trim(Text9.Text)
If Combo2.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·ﬁ”„", vbCritical
Exit Sub
End If
If Text6.Text = "" Then
MsgBox "«œŒ· —ﬁ„ «· ·„Ì–", vbCritical
Text6.SetFocus
Exit Sub
End If
If Val(Text6.Text) = 0 Then
MsgBox "—ﬁ„ «· ·„Ì– «·„œŒ· €Ì— ”·Ì„", vbCritical
Exit Sub
End If
If Val(Text6.Text) > 999999 Then
MsgBox "ÌÃ» √‰ ·«Ì Ã«Ê“ —ﬁ„ «· ·„Ì– ”  ‘›—« ", vbCritical
Exit Sub
End If
If Text7.Text = "" Then
MsgBox "«œŒ· «”„ «· ·„Ì–", vbCritical
Text7.SetFocus
Exit Sub
End If
If Combo3.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— Ã‰” «· ·„Ì–", vbCritical
Exit Sub
End If
'Control NNI
'If Text9.Text <> "" Then
'nni = Text9.Text
'md = nni Mod 97
'If md <> 1 Then
'MsgBox "«·—ﬁ„ «·Êÿ‰Ì «·„œŒ· €Ì— ’ÕÌÕ", vbCritical
''Text9.SetFocus
'Exit Sub
'End If
'End If

'If Text8.Text = "" Then
'MsgBox "«œŒ· —ﬁ„ ÊﬂÌ· «· ·„Ì–", vbCritical
'Text8.SetFocus
'Exit Sub
'End If

Call cont
Do While Not et.EOF
If Label21.Caption <> et!aut And et!cla = Combo2.Text And et!num = Text6.Text Then
MsgBox "€Ì— „„ﬂ‰ .. ÌÊÃœ  ·„Ì– ¬Œ— ÌÕ„· ‰›” «·—ﬁ„ ›Ì ‰›” «·ﬁ”„", vbCritical
Exit Sub
End If
et.MoveNext
Loop
If Label11.Caption = "" Then
g = MsgBox("·„ Ì „ «—›«ﬁ ’Ê—… ·· ·„Ì– , Â·  —Ìœ «·«” „—«—ø", vbInformation + vbYesNo, "Pressing")
If g = vbYes Then
Label11.Caption = "01"
If Combo3.Text = "√‰ÀÏ" Then
PicFilev = App.Path & "\nophotof.jpg"
Else
PicFilev = App.Path & "\nophotom.jpg"
End If
fName = PicFilev
Image2.Picture = LoadPicture(PicFilev) 'Afficher l'image
Else
Exit Sub
End If
End If
'bulltin
Call cont2
Do While Not be.EOF
If BarcodeX2.Caption = be!ser Then
be!cla = Combo2.Text
be!mtr = Text6.Text
be!nom = Text7.Text
be!obs19 = Text21.Text
be!ann = face.SBB1.Panels(9).Text
be!eco = face.SBB1.Panels(13).Text
be.Update
End If
be.MoveNext
Loop
'Recu
Call cont2
Do While Not ru.EOF
If BarcodeX2.Caption = ru!ser Then
ru!cla = Combo2.Text
ru!num = Text6.Text
ru!nom = Text7.Text
ru!ann = face.SBB1.Panels(9).Text
ru!eco = face.SBB1.Panels(13).Text
ru.Update
End If
ru.MoveNext
Loop
'nni
x = 0
Call cont2
Do While Not nn.EOF
If BarcodeX2.Caption = nn!ser Then
nn!nni = Text9.Text
nn!liu = Text13.Text
nn!dat = Text16.Text
nn!cla = Combo2.Text
nn!num = Text6.Text
nn!dti = DT2.Value
nn!nom = Text7.Text
nn!nof = Text18.Text
nn!rim = Text21.Text
nn!sex = Combo3.Text
nn!tel = Text8.Text
nn!ncla = Label5.Caption
nn.Update
nn.MoveLast
x = 1
End If
nn.MoveNext
Loop
If x = 0 Then
nn.AddNew
nn!ser = BarcodeX2.Caption
nn!nni = Text9.Text
nn!liu = Text13.Text
nn!dat = Text16.Text
nn!cla = Combo2.Text
nn!num = Text6.Text
nn!dti = DT2.Value
nn!nom = Text7.Text
nn!nof = Text18.Text
nn!rim = Text21.Text
nn!sex = Combo3.Text
nn!tel = Text8.Text
nn!ncla = Label5.Caption
nn.Update
End If
Call cont
Do While Not et.EOF
If BarcodeX2.Caption = et!ser Then
et!cla = Combo2.Text
et!num = Text6.Text
et!dat = DT2.Value
et!nom = Text7.Text
et!nof = Text18.Text
et!rim = Text21.Text
et!sex = Combo3.Text
If Label11.Caption = "" Then
et!pho = "01"
End If
et!tel = Text8.Text
et!adr = Text9.Text
et.Update
Label36.Caption = Combo2.Text
Label37.Caption = Text6.Text
Label38.Caption = Text7.Text
Label81.Caption = Combo2.Text
Label80.Caption = Text6.Text
Label76.Caption = Text7.Text
Text17.Text = Text21.Text
grd1.Clear
grd1.Rows = 1
Timer1.Enabled = True
Exit Sub
End If
et.MoveNext
Loop

End Sub


Private Sub DT4_Change()
On Error Resume Next
grd10.Clear
grd10.Rows = 1
PicFilev = ""
Image4.Picture = LoadPicture(PicFilev)
For j = 0 To 3
Text11(j).Text = ""
Next j
Call recheche
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
tim = 0
DT1.Value = Date
DT2.Value = Date
DT3.Value = Date
End Sub
Public Sub chargec1()
On Error Resume Next
Call cont
Combo1.Clear
Combo2.Clear
Combo5.Clear
Combo11.Clear
  Do While Not cl.EOF
  If cl!act = "1" Then
    Combo1.AddItem cl!cla
    Combo2.AddItem cl!cla
    Combo5.AddItem cl!cla
    Combo11.AddItem cl!cla
End If
cl.MoveNext
  Loop
End Sub
Public Sub chargec2()
On Error Resume Next
Combo3.Clear
Combo4.Clear
    Combo3.AddItem "–ﬂ—"
    Combo3.AddItem "√‰ÀÏ"
Combo4.AddItem "–ﬂ—"
Combo4.AddItem "√‰ÀÏ"
Combo4.Text = "–ﬂ—"
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
Private Sub numsetu()
On Error Resume Next
Dim a As Double
Dim b As Double
Dim s As Double
Dim f As Double
Dim m As Double
a = 0
s = 0
m = 0
f = 0
Call cont
Do While Not et.EOF
If Val(et!num) < 1000000 Then
If et!cla = Combo1.Text Then
b = et!num
If b > a Then
a = b
End If
s = s + 1
If et!sex = "–ﬂ—" Then
m = m + 1
Else
f = f + 1
End If
End If
End If
et.MoveNext
Loop
a = a + 1
Text3.Text = a
Label90(3).Caption = s
Label90(5).Caption = m
Label90(7).Caption = f
End Sub
Private Sub numsetu2()
On Error Resume Next
Dim a As Double
Dim b As Double
a = 0
Call cont
Do While Not et.EOF
If Val(et!num) < 1000000 Then
If et!cla = Combo2.Text Then
b = et!num
If b > a Then
a = b
End If
End If
End If
et.MoveNext
Loop
a = a + 1
Text6.Text = a
End Sub
Private Sub numclasfornni1()
On Error Resume Next
Call cont
Do While Not cl.EOF
If cl!cla = Combo1.Text Then
Label4.Caption = cl!num
Exit Sub
End If
cl.MoveNext
Loop
End Sub
Private Sub numclasfornni2()
On Error Resume Next
Call cont
Do While Not cl.EOF
If cl!cla = Combo2.Text Then
Label5.Caption = cl!num
Exit Sub
End If
cl.MoveNext
Loop
End Sub
Private Sub numclasfornni3()
On Error Resume Next
Call cont
Do While Not cl.EOF
If cl!cla = Combo5.Text Then
Label7.Caption = cl!num
Exit Sub
End If
cl.MoveNext
Loop
End Sub
Private Sub serial()
On Error Resume Next
Dim a As Double
Dim b As String
Call cont
a = Val(sr!num)
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
SSTab2.Visible = False
Picture4.Visible = True

End Sub

Private Sub grd1_Click()
On Error Resume Next
Dim r As Double
Dim c As Double
Dim tx1 As String
Dim tx2 As String
Dim tx3 As String
Dim a As Double
Dim b As Double
Dim f As Double
Dim j As Double
Dim k As Double
Dim l As Double
Dim controle As Double
If Text1.Text = "" Then
MsgBox "ÌÃ» «œŒ«· «·—ﬁ„ «· ”·”·Ì ·· ·„Ì–", vbCritical
Text1.SetFocus
Exit Sub
End If
If Label36.Caption = "" Then
MsgBox "ÌÃ» «œŒ«· «·—ﬁ„ «· ”·”·Ì ·· ·„Ì– À„ «·÷€ÿ ⁄·Ï Enter", vbCritical + arabic
Exit Sub
End If
If grd1.Rows < 2 Then
grd1.row = 0
grd1.Col = 1
tx3 = grd1.Text
If tx3 = "«·„«œ…" Then
MsgBox "·«  ÊÃœ ·Â–« «·ﬁ”„ √Ì „«œ…, ÌÃ» ≈÷«›… „Ê«œ ·Â–« «·ﬁ”„ „‰ —ﬂ‰ «·√ﬁ”«„", vbCritical
Exit Sub
Else
MsgBox "ÌÃ» «·÷€ÿ ⁄·Ï “— ⁄—÷ «·‰ «∆Ã", vbCritical
Exit Sub
End If
Exit Sub
End If
c = grd1.Col
r = grd1.row
If c > 2 And r > 0 And c <> 13 And c <> 17 And c <> 18 Then
grd1.Col = c
grd1.row = 0
tx1 = grd1.Text
grd1.row = r
grd1.Col = 1
tx2 = grd1.Text
g = InputBox("«œŒ· «·‰ ÌÃ…", tx1 + "  " + tx2)
grd1.Visible = False
If g = Cancel Then
Exit Sub
grd1.Visible = True
End If
'controle
If Val(g) = 0 Then
Else
b = g
If b < 0 Then
grd1.Visible = True
MsgBox "·« Ì„ﬂ‰ ··‰ ÌÃ… √‰  ﬂÊ‰  Õ  «·’›—", vbCritical
Exit Sub
End If
If b > 20 Then
grd1.Visible = True
MsgBox "·« Ì„ﬂ‰ ··‰ ÌÃ… √‰  ›Êﬁ 20", vbCritical
Exit Sub
End If
End If
Label42.Caption = ""
grd1.Enabled = False
grd1.Col = c
grd1.row = r
grd1.Text = g
Call calculmoyenne
j = grd1.Rows
controle = 0
Call cont2
Do While Not be.EOF
If be!cla = Label36.Caption And be!mtr = Label37.Caption Then
controle = 1
be!obs18 = Text9.Text
be!obs19 = Text17.Text
be!cla = Label36.Caption
be!mtr = Label37.Caption
be!nom = Label38.Caption
be!mena = Label40.Caption
be!menf = Label29.Caption
be!numm = Label106.Caption
be!eco = face.SBB1.Panels(13).Text
be!ann = face.SBB1.Panels(9).Text
be!moy = Label33.Caption
be!dat = Date
be!Abs = Label73.Caption
be!ran = Label42.Caption
be!ser = BarcodeX2.Caption
If j > 1 Then
grd1.row = 1
grd1.Col = 1
be!mat1 = grd1.Text
grd1.Col = 2
be!cof1 = grd1.Text
grd1.Col = 13
be!md1 = grd1.Text
grd1.Col = 14
be!e11 = grd1.Text
grd1.Col = 15
be!e21 = grd1.Text
grd1.Col = 16
be!e31 = grd1.Text
grd1.Col = 17
be!mm1 = grd1.Text
'be!obs1 = ne!obs
be.Update
End If
If j > 2 Then
grd1.row = 2
grd1.Col = 1
be!mat2 = grd1.Text
grd1.Col = 2
be!cof2 = grd1.Text
grd1.Col = 13
be!md2 = grd1.Text
grd1.Col = 14
be!e12 = grd1.Text
grd1.Col = 15
be!e22 = grd1.Text
grd1.Col = 16
be!e32 = grd1.Text
grd1.Col = 17
be!mm2 = grd1.Text
'be!obs2 = ne!obs
be.Update
End If
If j > 3 Then
grd1.row = 3
grd1.Col = 1
be!mat3 = grd1.Text
grd1.Col = 2
be!cof3 = grd1.Text
grd1.Col = 13
be!md3 = grd1.Text
grd1.Col = 14
be!e13 = grd1.Text
grd1.Col = 15
be!e23 = grd1.Text
grd1.Col = 16
be!e33 = grd1.Text
grd1.Col = 17
be!mm3 = grd1.Text
'be!obs3 = ne!obs
be.Update
End If
If j > 4 Then
grd1.row = 4
grd1.Col = 1
be!mat4 = grd1.Text
grd1.Col = 2
be!cof4 = grd1.Text
grd1.Col = 13
be!md4 = grd1.Text
grd1.Col = 14
be!e14 = grd1.Text
grd1.Col = 15
be!e24 = grd1.Text
grd1.Col = 16
be!e34 = grd1.Text
grd1.Col = 17
be!mm4 = grd1.Text
'be!obs4 = ne!obs
be.Update
End If
If j > 5 Then
grd1.row = 5
grd1.Col = 1
be!mat5 = grd1.Text
grd1.Col = 2
be!cof5 = grd1.Text
grd1.Col = 13
be!md5 = grd1.Text
grd1.Col = 14
be!e15 = grd1.Text
grd1.Col = 15
be!e25 = grd1.Text
grd1.Col = 16
be!e35 = grd1.Text
grd1.Col = 17
be!mm5 = grd1.Text
'grd1.Col = 1
'be!obs5 = ne!obs
be.Update
End If
If j > 6 Then
grd1.row = 6
grd1.Col = 1
be!mat6 = grd1.Text
grd1.Col = 2
be!cof6 = grd1.Text
grd1.Col = 13
be!md6 = grd1.Text
grd1.Col = 14
be!e16 = grd1.Text
grd1.Col = 15
be!e26 = grd1.Text
grd1.Col = 16
be!e36 = grd1.Text
grd1.Col = 17
be!mm6 = grd1.Text
'grd1.Col = 1
'be!obs6 = ne!obs
be.Update
End If
If j > 7 Then
grd1.row = 7
grd1.Col = 1
be!mat7 = grd1.Text
grd1.Col = 2
be!cof7 = grd1.Text
grd1.Col = 13
be!md7 = grd1.Text
grd1.Col = 14
be!e17 = grd1.Text
grd1.Col = 15
be!e27 = grd1.Text
grd1.Col = 16
be!e37 = grd1.Text
grd1.Col = 17
be!mm7 = grd1.Text
'be!obs7 = ne!obs
be.Update
End If
If j > 8 Then
grd1.row = 8
grd1.Col = 1
be!mat8 = grd1.Text
grd1.Col = 2
be!cof8 = grd1.Text
grd1.Col = 13
be!md8 = grd1.Text
grd1.Col = 14
be!e18 = grd1.Text
grd1.Col = 15
be!e28 = grd1.Text
grd1.Col = 16
be!e38 = grd1.Text
grd1.Col = 17
be!mm8 = grd1.Text
'be!obs8 = ne!obs
be.Update
End If
If j > 9 Then
grd1.row = 9
grd1.Col = 1
be!mat9 = grd1.Text
grd1.Col = 2
be!cof9 = grd1.Text
grd1.Col = 13
be!md9 = grd1.Text
grd1.Col = 14
be!e19 = grd1.Text
grd1.Col = 15
be!e29 = grd1.Text
grd1.Col = 16
be!e39 = grd1.Text
grd1.Col = 17
be!mm9 = grd1.Text
'be!obs9 = ne!obs
be.Update
End If
If j > 10 Then
grd1.row = 10
grd1.Col = 1
be!mat10 = grd1.Text
grd1.Col = 2
be!cof10 = grd1.Text
grd1.Col = 13
be!md10 = grd1.Text
grd1.Col = 14
be!e110 = grd1.Text
grd1.Col = 15
be!e210 = grd1.Text
grd1.Col = 16
be!e310 = grd1.Text
grd1.Col = 17
be!mm10 = grd1.Text
'be!obs10 = ne!obs
be.Update
End If
If j > 11 Then
grd1.row = 11
grd1.Col = 1
be!mat11 = grd1.Text
grd1.Col = 2
be!cof11 = grd1.Text
grd1.Col = 13
be!md11 = grd1.Text
grd1.Col = 14
be!e111 = grd1.Text
grd1.Col = 15
be!e211 = grd1.Text
grd1.Col = 16
be!e311 = grd1.Text
grd1.Col = 17
be!mm11 = grd1.Text
'be!obs11 = ne!obs
be.Update
End If
If j > 12 Then
grd1.row = 12
grd1.Col = 1
be!mat12 = grd1.Text
grd1.Col = 2
be!cof12 = grd1.Text
grd1.Col = 13
be!md12 = grd1.Text
grd1.Col = 14
be!e112 = grd1.Text
grd1.Col = 15
be!e212 = grd1.Text
grd1.Col = 16
be!e312 = grd1.Text
grd1.Col = 17
be!mm12 = grd1.Text
'be!obs12 = ne!obs
be.Update
End If
If j > 13 Then
grd1.row = 13
grd1.Col = 1
be!mat13 = grd1.Text
grd1.Col = 2
be!cof13 = grd1.Text
grd1.Col = 13
be!md13 = grd1.Text
grd1.Col = 14
be!e113 = grd1.Text
grd1.Col = 15
be!e213 = grd1.Text
grd1.Col = 16
be!e313 = grd1.Text
grd1.Col = 17
be!mm13 = grd1.Text
'be!obs13 = ne!obs
be.Update
End If
If j > 14 Then
grd1.row = 14
grd1.Col = 1
be!mat14 = grd1.Text
grd1.Col = 2
be!cof14 = grd1.Text
grd1.Col = 13
be!md14 = grd1.Text
grd1.Col = 14
be!e114 = grd1.Text
grd1.Col = 15
be!e214 = grd1.Text
grd1.Col = 16
be!e314 = grd1.Text
grd1.Col = 17
be!mm14 = grd1.Text
'be!obs14 = ne!obs
be.Update
End If
If j > 15 Then
grd1.row = 15
grd1.Col = 1
be!mat15 = grd1.Text
grd1.Col = 2
be!cof15 = grd1.Text
grd1.Col = 13
be!md15 = grd1.Text
grd1.Col = 14
be!e115 = grd1.Text
grd1.Col = 15
be!e215 = grd1.Text
grd1.Col = 16
be!e315 = grd1.Text
grd1.Col = 17
be!mm15 = grd1.Text
'be!obs15 = ne!obs
be.Update
End If
If j > 16 Then
grd1.row = 16
grd1.Col = 1
be!mat16 = grd1.Text
grd1.Col = 2
be!cof16 = grd1.Text
grd1.Col = 13
be!md16 = grd1.Text
grd1.Col = 14
be!e116 = grd1.Text
grd1.Col = 15
be!e216 = grd1.Text
grd1.Col = 16
be!e316 = grd1.Text
grd1.Col = 17
be!mm16 = grd1.Text
'be!obs16 = ne!obs
be.Update
End If
If j > 17 Then
grd1.row = 17
grd1.Col = 1
be!mat17 = grd1.Text
grd1.Col = 2
be!cof17 = grd1.Text
grd1.Col = 13
be!md17 = grd1.Text
grd1.Col = 14
be!e117 = grd1.Text
grd1.Col = 15
be!e217 = grd1.Text
grd1.Col = 16
be!e317 = grd1.Text
grd1.Col = 17
be!mm17 = grd1.Text
'be!obs17 = ne!obs
be.Update
End If
If j > 18 Then
grd1.row = 18
grd1.Col = 1
be!mat18 = grd1.Text
grd1.Col = 2
be!cof18 = grd1.Text
grd1.Col = 13
be!md18 = grd1.Text
grd1.Col = 14
be!e118 = grd1.Text
grd1.Col = 15
be!e218 = grd1.Text
grd1.Col = 16
be!e318 = grd1.Text
grd1.Col = 17
be!mm18 = grd1.Text
'be!obs18 = ne!obs
be.Update
End If
If j > 19 Then
grd1.row = 19
grd1.Col = 1
be!mat19 = grd1.Text
grd1.Col = 2
be!cof19 = grd1.Text
grd1.Col = 13
be!md19 = grd1.Text
grd1.Col = 14
be!e119 = grd1.Text
grd1.Col = 15
be!e219 = grd1.Text
grd1.Col = 16
be!e319 = grd1.Text
grd1.Col = 17
be!mm19 = grd1.Text
be.Update
End If
If j > 20 Then
grd1.row = 20
grd1.Col = 1
be!mat20 = grd1.Text
grd1.Col = 2
be!cof20 = grd1.Text
grd1.Col = 13
be!md20 = grd1.Text
grd1.Col = 14
be!e120 = grd1.Text
grd1.Col = 15
be!e220 = grd1.Text
grd1.Col = 16
be!e320 = grd1.Text
grd1.Col = 17
be!mm20 = grd1.Text
'be!obs20 = ne!obs
be.Update
End If
be.MoveLast
End If
be.MoveNext
Loop
k = 0
Call cont
Do While Not nt.EOF
If nt!cla = Label36.Caption And nt!num = Label37.Caption And nt!mat = tx2 Then
k = 1
grd1.row = r
grd1.Col = 1
nt!mat = grd1.Text
'grd1.Col = 2
'nt!cof = grd1.Text
grd1.Col = 3
nt!de1 = grd1.Text
grd1.Col = 4
nt!de2 = grd1.Text
grd1.Col = 5
nt!de3 = grd1.Text
grd1.Col = 6
nt!de4 = grd1.Text
grd1.Col = 7
nt!de5 = grd1.Text
grd1.Col = 8
nt!de6 = grd1.Text
grd1.Col = 9
nt!de7 = grd1.Text
grd1.Col = 10
nt!de8 = grd1.Text
grd1.Col = 11
nt!de9 = grd1.Text
grd1.Col = 12
nt!de10 = grd1.Text
grd1.Col = 13
nt!md = grd1.Text
grd1.Col = 14
nt!ex1 = grd1.Text
grd1.Col = 15
nt!ex2 = grd1.Text
grd1.Col = 16
nt!ex3 = grd1.Text
grd1.Col = 17
nt!mm = grd1.Text
grd1.Col = 18
nt!tot = grd1.Text
nt.Update
grd1.Enabled = True
'Exit Sub
nt.MoveLast
End If
nt.MoveNext
Loop
If k = 0 Then
nt.AddNew
nt!cla = Label36.Caption
nt!num = Label37.Caption
nt!nom = Label38.Caption
nt!ser = BarcodeX2.Caption
nt!nbr = Label41.Caption
nt!nme = Label67.Caption
grd1.row = r
grd1.Col = 1
nt!mat = grd1.Text
grd1.Col = 2
nt!cof = grd1.Text
grd1.Col = 3
nt!de1 = grd1.Text
grd1.Col = 4
nt!de2 = grd1.Text
grd1.Col = 5
nt!de3 = grd1.Text
grd1.Col = 6
nt!de4 = grd1.Text
grd1.Col = 7
nt!de5 = grd1.Text
grd1.Col = 8
nt!de6 = grd1.Text
grd1.Col = 9
nt!de7 = grd1.Text
grd1.Col = 10
nt!de8 = grd1.Text
grd1.Col = 11
nt!de9 = grd1.Text
grd1.Col = 12
nt!de10 = grd1.Text
grd1.Col = 13
nt!md = grd1.Text
grd1.Col = 14
nt!ex1 = grd1.Text
grd1.Col = 15
nt!ex2 = grd1.Text
grd1.Col = 16
nt!ex3 = grd1.Text
grd1.Col = 17
nt!mm = grd1.Text
grd1.Col = 18
nt!tot = grd1.Text
nt.Update
Call calculmoyenne
grd1.Enabled = True
End If
'******** bultin
If controle = 0 Then
Call cont2
be.AddNew
be!cla = Label36.Caption
be!mtr = Label37.Caption
be!nom = Label38.Caption
be!mena = ""
be!menf = ""
be!numm = Label106.Caption
be!eco = face.SBB1.Panels(13).Text
be!ann = face.SBB1.Panels(9).Text
be!moy = ""
be!dat = ""
be!Abs = ""
be!ran = ""
be!ser = BarcodeX2.Caption
be!mat1 = ""
be!cof1 = ""
be!md1 = ""
be!e11 = ""
be!e21 = ""
be!e31 = ""
be!mm1 = ""
be!obs1 = ""
be!mat2 = ""
be!cof2 = ""
be!md2 = ""
be!e12 = ""
be!e22 = ""
be!e32 = ""
be!mm2 = ""
be!obs2 = ""
be!mat3 = ""
be!cof3 = ""
be!md3 = ""
be!e13 = ""
be!e23 = ""
be!e33 = ""
be!mm3 = ""
be!obs3 = ""
be!mat4 = ""
be!cof4 = ""
be!md4 = ""
be!e14 = ""
be!e24 = ""
be!e34 = ""
be!mm4 = ""
be!obs4 = ""
be!mat5 = ""
be!cof5 = ""
be!md5 = ""
be!e15 = ""
be!e25 = ""
be!e35 = ""
be!mm5 = ""
be!obs5 = ""
be!mat6 = ""
be!cof6 = ""
be!md6 = ""
be!e16 = ""
be!e26 = ""
be!e36 = ""
be!mm6 = ""
be!obs6 = ""
be!mat7 = ""
be!cof7 = ""
be!md7 = ""
be!e17 = ""
be!e27 = ""
be!e37 = ""
be!mm7 = ""
be!obs7 = ""
be!mat8 = ""
be!cof8 = ""
be!md8 = ""
be!e18 = ""
be!e28 = ""
be!e38 = ""
be!mm8 = ""
be!obs8 = ""
be!mat9 = ""
be!cof9 = ""
be!md9 = ""
be!e19 = ""
be!e29 = ""
be!e39 = ""
be!mm9 = ""
be!obs9 = ""
be!mat10 = ""
be!cof10 = ""
be!md10 = ""
be!e110 = ""
be!e210 = ""
be!e310 = ""
be!mm10 = ""
be!obs10 = ""
be!mat11 = ""
be!cof11 = ""
be!md11 = ""
be!e111 = ""
be!e211 = ""
be!e311 = ""
be!mm11 = ""
be!obs11 = ""
be!mat12 = ""
be!cof12 = ""
be!md12 = ""
be!e112 = ""
be!e212 = ""
be!e312 = ""
be!mm12 = ""
be!obs12 = ""
be!mat13 = ""
be!cof13 = ""
be!md13 = ""
be!e113 = ""
be!e213 = ""
be!e313 = ""
be!mm13 = ""
be!obs13 = ""
be!mat14 = ""
be!cof14 = ""
be!md14 = ""
be!e114 = ""
be!e214 = ""
be!e314 = ""
be!mm14 = ""
be!obs14 = ""
be!mat15 = ""
be!cof15 = ""
be!md15 = ""
be!e115 = ""
be!e215 = ""
be!e315 = ""
be!mm15 = ""
be!obs15 = ""
be!mat16 = ""
be!cof16 = ""
be!md16 = ""
be!e116 = ""
be!e216 = ""
be!e316 = ""
be!mm16 = ""
be!obs16 = ""
be!mat17 = ""
be!cof17 = ""
be!md17 = ""
be!e117 = ""
be!e217 = ""
be!e317 = ""
be!mm17 = ""
be!obs17 = ""
be!mat18 = ""
be!cof18 = ""
be!md18 = ""
be!e118 = ""
be!e218 = ""
be!e318 = ""
be!mm18 = ""
be!mat19 = ""
be!cof19 = ""
be!md19 = ""
be!e119 = ""
be!e219 = ""
be!e319 = ""
be!mm19 = ""
be!obs18 = Text9.Text
be!obs19 = Text17.Text
be!mat20 = ""
be!cof20 = ""
be!md20 = ""
be!e120 = ""
be!e220 = ""
be!e320 = ""
be!mm20 = ""
be!obs20 = ""
be.Update
be!cla = Label36.Caption
be!mtr = Label37.Caption
be!nom = Label38.Caption
be!mena = Label40.Caption
be!menf = Label29.Caption
be!numm = Label106.Caption
be!eco = face.SBB1.Panels(13).Text
be!ann = face.SBB1.Panels(9).Text
be!moy = Label33.Caption
be!dat = Date
be!Abs = Label73.Caption
be!ran = Label42.Caption
be!ser = BarcodeX2.Caption
If j > 1 Then
grd1.row = 1
grd1.Col = 1
be!mat1 = grd1.Text
grd1.Col = 2
be!cof1 = grd1.Text
grd1.Col = 13
be!md1 = grd1.Text
grd1.Col = 14
be!e11 = grd1.Text
grd1.Col = 15
be!e21 = grd1.Text
grd1.Col = 16
be!e31 = grd1.Text
grd1.Col = 17
be!mm1 = grd1.Text
'be!obs1 = ne!obs
be.Update
End If
If j > 2 Then
grd1.row = 2
grd1.Col = 1
be!mat2 = grd1.Text
grd1.Col = 2
be!cof2 = grd1.Text
grd1.Col = 13
be!md2 = grd1.Text
grd1.Col = 14
be!e12 = grd1.Text
grd1.Col = 15
be!e22 = grd1.Text
grd1.Col = 16
be!e32 = grd1.Text
grd1.Col = 17
be!mm2 = grd1.Text
'be!obs2 = ne!obs
be.Update
End If
If j > 3 Then
grd1.row = 3
grd1.Col = 1
be!mat3 = grd1.Text
grd1.Col = 2
be!cof3 = grd1.Text
grd1.Col = 13
be!md3 = grd1.Text
grd1.Col = 14
be!e13 = grd1.Text
grd1.Col = 15
be!e23 = grd1.Text
grd1.Col = 16
be!e33 = grd1.Text
grd1.Col = 17
be!mm3 = grd1.Text
'be!obs3 = ne!obs
be.Update
End If
If j > 4 Then
grd1.row = 4
grd1.Col = 1
be!mat4 = grd1.Text
grd1.Col = 2
be!cof4 = grd1.Text
grd1.Col = 13
be!md4 = grd1.Text
grd1.Col = 14
be!e14 = grd1.Text
grd1.Col = 15
be!e24 = grd1.Text
grd1.Col = 16
be!e34 = grd1.Text
grd1.Col = 17
be!mm4 = grd1.Text
'be!obs4 = ne!obs
be.Update
End If
If j > 5 Then
grd1.row = 5
grd1.Col = 1
be!mat5 = grd1.Text
grd1.Col = 2
be!cof5 = grd1.Text
grd1.Col = 13
be!md5 = grd1.Text
grd1.Col = 14
be!e15 = grd1.Text
grd1.Col = 15
be!e25 = grd1.Text
grd1.Col = 16
be!e35 = grd1.Text
grd1.Col = 17
be!mm5 = grd1.Text
'grd1.Col = 1
'be!obs5 = ne!obs
be.Update
End If
If j > 6 Then
grd1.row = 6
grd1.Col = 1
be!mat6 = grd1.Text
grd1.Col = 2
be!cof6 = grd1.Text
grd1.Col = 13
be!md6 = grd1.Text
grd1.Col = 14
be!e16 = grd1.Text
grd1.Col = 15
be!e26 = grd1.Text
grd1.Col = 16
be!e36 = grd1.Text
grd1.Col = 17
be!mm6 = grd1.Text
'grd1.Col = 1
'be!obs6 = ne!obs
be.Update
End If
If j > 7 Then
grd1.row = 7
grd1.Col = 1
be!mat7 = grd1.Text
grd1.Col = 2
be!cof7 = grd1.Text
grd1.Col = 13
be!md7 = grd1.Text
grd1.Col = 14
be!e17 = grd1.Text
grd1.Col = 15
be!e27 = grd1.Text
grd1.Col = 16
be!e37 = grd1.Text
grd1.Col = 17
be!mm7 = grd1.Text
'be!obs7 = ne!obs
be.Update
End If
If j > 8 Then
grd1.row = 8
grd1.Col = 1
be!mat8 = grd1.Text
grd1.Col = 2
be!cof8 = grd1.Text
grd1.Col = 13
be!md8 = grd1.Text
grd1.Col = 14
be!e18 = grd1.Text
grd1.Col = 15
be!e28 = grd1.Text
grd1.Col = 16
be!e38 = grd1.Text
grd1.Col = 17
be!mm8 = grd1.Text
'be!obs8 = ne!obs
be.Update
End If
If j > 9 Then
grd1.row = 9
grd1.Col = 1
be!mat9 = grd1.Text
grd1.Col = 2
be!cof9 = grd1.Text
grd1.Col = 13
be!md9 = grd1.Text
grd1.Col = 14
be!e19 = grd1.Text
grd1.Col = 15
be!e29 = grd1.Text
grd1.Col = 16
be!e39 = grd1.Text
grd1.Col = 17
be!mm9 = grd1.Text
'be!obs9 = ne!obs
be.Update
End If
If j > 10 Then
grd1.row = 10
grd1.Col = 1
be!mat10 = grd1.Text
grd1.Col = 2
be!cof10 = grd1.Text
grd1.Col = 13
be!md10 = grd1.Text
grd1.Col = 14
be!e110 = grd1.Text
grd1.Col = 15
be!e210 = grd1.Text
grd1.Col = 16
be!e310 = grd1.Text
grd1.Col = 17
be!mm10 = grd1.Text
'be!obs10 = ne!obs
be.Update
End If
If j > 11 Then
grd1.row = 11
grd1.Col = 1
be!mat11 = grd1.Text
grd1.Col = 2
be!cof11 = grd1.Text
grd1.Col = 13
be!md11 = grd1.Text
grd1.Col = 14
be!e111 = grd1.Text
grd1.Col = 15
be!e211 = grd1.Text
grd1.Col = 16
be!e311 = grd1.Text
grd1.Col = 17
be!mm11 = grd1.Text
'be!obs11 = ne!obs
be.Update
End If
If j > 12 Then
grd1.row = 12
grd1.Col = 1
be!mat12 = grd1.Text
grd1.Col = 2
be!cof12 = grd1.Text
grd1.Col = 13
be!md12 = grd1.Text
grd1.Col = 14
be!e112 = grd1.Text
grd1.Col = 15
be!e212 = grd1.Text
grd1.Col = 16
be!e312 = grd1.Text
grd1.Col = 17
be!mm12 = grd1.Text
'be!obs12 = ne!obs
be.Update
End If
If j > 13 Then
grd1.row = 13
grd1.Col = 1
be!mat13 = grd1.Text
grd1.Col = 2
be!cof13 = grd1.Text
grd1.Col = 13
be!md13 = grd1.Text
grd1.Col = 14
be!e113 = grd1.Text
grd1.Col = 15
be!e213 = grd1.Text
grd1.Col = 16
be!e313 = grd1.Text
grd1.Col = 17
be!mm13 = grd1.Text
'be!obs13 = ne!obs
be.Update
End If
If j > 14 Then
grd1.row = 14
grd1.Col = 1
be!mat14 = grd1.Text
grd1.Col = 2
be!cof14 = grd1.Text
grd1.Col = 13
be!md14 = grd1.Text
grd1.Col = 14
be!e114 = grd1.Text
grd1.Col = 15
be!e214 = grd1.Text
grd1.Col = 16
be!e314 = grd1.Text
grd1.Col = 17
be!mm14 = grd1.Text
'be!obs14 = ne!obs
be.Update
End If
If j > 15 Then
grd1.row = 15
grd1.Col = 1
be!mat15 = grd1.Text
grd1.Col = 2
be!cof15 = grd1.Text
grd1.Col = 13
be!md15 = grd1.Text
grd1.Col = 14
be!e115 = grd1.Text
grd1.Col = 15
be!e215 = grd1.Text
grd1.Col = 16
be!e315 = grd1.Text
grd1.Col = 17
be!mm15 = grd1.Text
'be!obs15 = ne!obs
be.Update
End If
If j > 16 Then
grd1.row = 16
grd1.Col = 1
be!mat16 = grd1.Text
grd1.Col = 2
be!cof16 = grd1.Text
grd1.Col = 13
be!md16 = grd1.Text
grd1.Col = 14
be!e116 = grd1.Text
grd1.Col = 15
be!e216 = grd1.Text
grd1.Col = 16
be!e316 = grd1.Text
grd1.Col = 17
be!mm16 = grd1.Text
'be!obs16 = ne!obs
be.Update
End If
If j > 17 Then
grd1.row = 17
grd1.Col = 1
be!mat17 = grd1.Text
grd1.Col = 2
be!cof17 = grd1.Text
grd1.Col = 13
be!md17 = grd1.Text
grd1.Col = 14
be!e117 = grd1.Text
grd1.Col = 15
be!e217 = grd1.Text
grd1.Col = 16
be!e317 = grd1.Text
grd1.Col = 17
be!mm17 = grd1.Text
'be!obs17 = ne!obs
be.Update
End If
If j > 18 Then
grd1.row = 18
grd1.Col = 1
be!mat18 = grd1.Text
grd1.Col = 2
be!cof18 = grd1.Text
grd1.Col = 13
be!md18 = grd1.Text
grd1.Col = 14
be!e118 = grd1.Text
grd1.Col = 15
be!e218 = grd1.Text
grd1.Col = 16
be!e318 = grd1.Text
grd1.Col = 17
be!mm18 = grd1.Text
'be!obs18 = ne!obs
be.Update
End If
If j > 19 Then
grd1.row = 19
grd1.Col = 1
be!mat19 = grd1.Text
grd1.Col = 2
be!cof19 = grd1.Text
grd1.Col = 13
be!md19 = grd1.Text
grd1.Col = 14
be!e119 = grd1.Text
grd1.Col = 15
be!e219 = grd1.Text
grd1.Col = 16
be!e319 = grd1.Text
grd1.Col = 17
be!mm19 = grd1.Text
'be!obs19 = Text17.Text
be.Update
End If
If j > 20 Then
grd1.row = 20
grd1.Col = 1
be!mat20 = grd1.Text
grd1.Col = 2
be!cof20 = grd1.Text
grd1.Col = 13
be!md20 = grd1.Text
grd1.Col = 14
be!e120 = grd1.Text
grd1.Col = 15
be!e220 = grd1.Text
grd1.Col = 16
be!e320 = grd1.Text
grd1.Col = 17
be!mm20 = grd1.Text
'be!obs20 = ne!obs
be.Update
End If
End If
End If
grd1.Visible = True
End Sub


Private Sub grd10_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim tx1 As String
Dim tx2 As String
i = grd10.row
j = grd10.Col
grd10.row = i
grd10.Col = 0
tx1 = grd10.Text
grd10.Col = 1
tx2 = grd10.Text
Call cont
Do While Not et.EOF
If tx1 = et!num And tx2 = et!cla Then
Label116.Caption = et!cla
Label117.Caption = et!num
Label118.Caption = et!nom
BarcodeX3.Caption = et!ser
Label119.Caption = et!tel
et.MoveLast
End If
et.MoveNext
Loop
Call LoadPictureFromDB6
End Sub

Private Sub grd2_Click()
On Error Resume Next
Dim i As Double
i = grd2.row
grd2.row = i
grd2.Col = 0
Label87.Caption = grd2.Text
grd2.Col = 1
DT3.Value = grd2.Text
grd2.Col = 2
Combo7.Text = grd2.Text
grd2.Col = 3
Combo8.Text = grd2.Text
grd2.Col = 5
Combo9.Text = grd2.Text
grd2.Col = 6
Text10.Text = grd2.Text
End Sub

Private Sub grd25_Click()
On Error Resume Next
Dim i As Double
i = grd25.row
If i > 0 Then
grd25.row = i
grd25.Col = 3
Text1.Text = grd25.Text
Command8_Click
End If
End Sub

Private Sub grd6_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim tx As String
Image3.Picture = LoadPicture("")
i = grd6.row
j = grd6.Col
grd6.row = i
grd6.Col = 0
tx = grd6.Text
Call cont
Do While Not et.EOF
If tx = et!aut Then
Label95.Caption = et!num
Label96.Caption = et!nom
Label115.Caption = et!ser
et.MoveLast
End If
et.MoveNext
Loop
Call LoadPictureFromDB4
If Check1.Value = 1 Then
Command28_Click
End If
End Sub

Private Sub grd8_Click()
On Error Resume Next
Dim r As Double
Dim c As Double
Dim tx1 As String
Dim tx2 As String
Dim tx3 As String
Dim tx4 As String
Dim a As Double
Dim b As Double
Dim f As Double
Dim j As Double
Dim controle As Double
If Combo11.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·ﬁ”„", vbCritical
Exit Sub
End If
If Combo10.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·„«œ…", vbCritical
Exit Sub
End If
If grd8.Rows < 2 Then
grd8.row = 0
grd8.Col = 1
tx3 = grd8.Text
If tx3 = "«·«”„" Then
MsgBox "·« ÌÊÃœ ›Ì Â–« «·ﬁ”„ √Ì  ·„Ì–, ÌÃ» ≈÷«›…  ·«„Ì– ·Â–« «·ﬁ”„ √Ê·«", vbCritical
Exit Sub
Else
MsgBox "ÌÃ» «·÷€ÿ ⁄·Ï “— ⁄—÷ «·‰ «∆Ã", vbCritical
Exit Sub
End If
Exit Sub
End If
c = grd8.Col
r = grd8.row
If c = 0 Then
grd8.Col = 0
grd8.Sort = 1
Exit Sub
End If
If c = 1 Then
grd8.Col = 1
grd8.Sort = 1
Exit Sub
End If
If c > 2 And r > 0 And c <> 13 And c <> 17 And c <> 18 Then
grd8.Col = c
grd8.row = 0
tx1 = grd8.Text
grd8.row = r
grd8.Col = 1
tx2 = grd8.Text
g = InputBox("«œŒ· «·‰ ÌÃ…", tx1 + "  " + tx2)
If g = Cancel Then
Exit Sub
End If
'controle
grd8.Visible = False
If Val(g) = 0 Then
Else
b = g
If b < 0 Then
MsgBox "·« Ì„ﬂ‰ ··‰ ÌÃ… √‰  ﬂÊ‰  Õ  «·’›—", vbCritical
Exit Sub
End If
If b > 20 Then
MsgBox "·« Ì„ﬂ‰ ··‰ ÌÃ… √‰  ›Êﬁ 20", vbCritical
Exit Sub
End If
End If
grd8.row = r
grd8.Col = 0
tx1 = grd8.Text
grd8.Col = 1
tx2 = grd8.Text
'Label102.Caption = ""
grd8.Enabled = False
grd8.Col = c
grd8.row = r
grd8.Text = g
Call calculmoyenne2
j = grd9.Rows
controle = 0
Call cont2
Do While Not be.EOF
If be!cla = Combo11.Text And be!mtr = tx1 Then
controle = 1
be!cla = Combo11.Text
be!mtr = tx1
be!nom = tx2
be!mena = Label40.Caption
be!menf = Label29.Caption
be!numm = Label106.Caption
be!eco = face.SBB1.Panels(13).Text
be!ann = face.SBB1.Panels(9).Text
be!moy = Label33.Caption
be!dat = Date
'be!Abs = ""
'be!ran = ""
grd8.row = r
grd8.Col = 19
be!ser = grd8.Text
If j > 1 Then
grd9.row = 1
grd9.Col = 1
tx4 = grd9.Text
If tx4 = Combo10.Text Then
be!mat1 = tx4
grd8.row = r
grd8.Col = 2
be!cof1 = grd8.Text
grd8.Col = 13
be!md1 = grd8.Text
grd8.Col = 14
be!e11 = grd8.Text
grd8.Col = 15
be!e21 = grd8.Text
grd8.Col = 16
be!e31 = grd8.Text
grd8.Col = 17
be!mm1 = grd8.Text
'be!obs1 = ne!obs
be.Update
End If
be.Update
End If
If j > 2 Then
grd9.row = 2
grd9.Col = 1
tx4 = grd9.Text
If tx4 = Combo10.Text Then
be!mat2 = tx4
grd8.row = r
grd8.Col = 2
be!cof2 = grd8.Text
grd8.Col = 13
be!md2 = grd8.Text
grd8.Col = 14
be!e12 = grd8.Text
grd8.Col = 15
be!e22 = grd8.Text
grd8.Col = 16
be!e32 = grd8.Text
grd8.Col = 17
be!mm2 = grd8.Text
'be!obs2 = ne!obs
be.Update
End If
be.Update
End If
If j > 3 Then
grd9.row = 3
grd9.Col = 1
tx4 = grd9.Text
If tx4 = Combo10.Text Then
be!mat3 = tx4
grd8.row = r
grd8.Col = 2
be!cof3 = grd8.Text
grd8.Col = 13
be!md3 = grd8.Text
grd8.Col = 14
be!e13 = grd8.Text
grd8.Col = 15
be!e23 = grd8.Text
grd8.Col = 16
be!e33 = grd8.Text
grd8.Col = 17
be!mm3 = grd8.Text
'be!obs3 = ne!obs
be.Update
End If
be.Update
End If
If j > 4 Then
grd9.row = 4
grd9.Col = 1
tx4 = grd9.Text
If tx4 = Combo10.Text Then
be!mat4 = tx4
grd8.row = r
grd8.Col = 2
be!cof4 = grd8.Text
grd8.Col = 13
be!md4 = grd8.Text
grd8.Col = 14
be!e14 = grd8.Text
grd8.Col = 15
be!e24 = grd8.Text
grd8.Col = 16
be!e34 = grd8.Text
grd8.Col = 17
be!mm4 = grd8.Text
'beobs4 = ne!obs
be.Update
End If
be.Update
End If
If j > 5 Then
grd9.row = 5
grd9.Col = 1
tx4 = grd9.Text
If tx4 = Combo10.Text Then
be!mat5 = tx4
grd8.row = r
grd8.Col = 2
be!cof5 = grd8.Text
grd8.Col = 13
be!md5 = grd8.Text
grd8.Col = 14
be!e15 = grd8.Text
grd8.Col = 15
be!e25 = grd8.Text
grd8.Col = 16
be!e35 = grd8.Text
grd8.Col = 17
be!mm5 = grd8.Text
'grd1.Col = 1
'be!obs5 = ne!obs
be.Update
End If
be.Update
End If
If j > 6 Then
grd9.row = 6
grd9.Col = 1
tx4 = grd9.Text
If tx4 = Combo10.Text Then
be!mat6 = tx4
grd8.row = r
grd8.Col = 2
be!cof6 = grd8.Text
grd8.Col = 13
be!md6 = grd8.Text
grd8.Col = 14
be!e16 = grd8.Text
grd8.Col = 15
be!e26 = grd8.Text
grd8.Col = 16
be!e36 = grd8.Text
grd8.Col = 17
be!mm6 = grd8.Text
'grd1.Col = 1
'be!obs6 = ne!obs
be.Update
End If
be.Update
End If
If j > 7 Then
grd9.row = 7
grd9.Col = 1
tx4 = grd9.Text
If tx4 = Combo10.Text Then
be!mat7 = tx4
grd8.row = r
grd8.Col = 2
be!cof7 = grd8.Text
grd8.Col = 13
be!md7 = grd8.Text
grd8.Col = 14
be!e17 = grd8.Text
grd8.Col = 15
be!e27 = grd8.Text
grd8.Col = 16
be!e37 = grd8.Text
grd8.Col = 17
be!mm7 = grd8.Text
'be!obs7 = ne!obs
be.Update
End If
be.Update
End If
If j > 8 Then
grd9.row = 8
grd9.Col = 1
tx4 = grd9.Text
If tx4 = Combo10.Text Then
be!mat8 = tx4
grd8.row = r
grd8.Col = 2
be!cof8 = grd8.Text
grd8.Col = 13
be!md8 = grd8.Text
grd8.Col = 14
be!e18 = grd8.Text
grd8.Col = 15
be!e28 = grd8.Text
grd8.Col = 16
be!e38 = grd8.Text
grd8.Col = 17
be!mm8 = grd8.Text
'be!obs8 = ne!obs
be.Update
End If
be.Update
End If
If j > 9 Then
grd9.row = 9
grd9.Col = 1
tx4 = grd9.Text
If tx4 = Combo10.Text Then
be!mat9 = tx4
grd8.row = r
grd8.Col = 2
be!cof9 = grd8.Text
grd8.Col = 13
be!md9 = grd8.Text
grd8.Col = 14
be!e19 = grd8.Text
grd8.Col = 15
be!e29 = grd8.Text
grd8.Col = 16
be!e39 = grd8.Text
grd8.Col = 17
be!mm9 = grd8.Text
'be!obs9 = ne!obs
be.Update
End If
be.Update
End If
If j > 10 Then
grd9.row = 10
grd9.Col = 1
tx4 = grd9.Text
If tx4 = Combo10.Text Then
be!mat10 = tx4
grd8.row = r
grd8.Col = 2
be!cof10 = grd8.Text
grd8.Col = 13
be!md10 = grd8.Text
grd8.Col = 14
be!e110 = grd8.Text
grd8.Col = 15
be!e210 = grd8.Text
grd8.Col = 16
be!e310 = grd8.Text
grd8.Col = 17
be!mm10 = grd8.Text
'be!obs10 = ne!obs
be.Update
End If
be.Update
End If
If j > 11 Then
grd9.row = 11
grd9.Col = 1
tx4 = grd9.Text
If tx4 = Combo10.Text Then
be!mat11 = tx4
grd8.row = r
grd8.Col = 2
be!cof11 = grd8.Text
grd8.Col = 13
be!md11 = grd8.Text
grd8.Col = 14
be!e111 = grd8.Text
grd8.Col = 15
be!e211 = grd8.Text
grd8.Col = 16
be!e311 = grd8.Text
grd8.Col = 17
be!mm11 = grd8.Text
'be!obs11 = ne!obs
be.Update
End If
be.Update
End If
If j > 12 Then
grd9.row = 12
grd9.Col = 1
tx4 = grd9.Text
If tx4 = Combo10.Text Then
be!mat12 = tx4
grd8.row = r
grd8.Col = 2
be!cof12 = grd8.Text
grd8.Col = 13
be!md12 = grd8.Text
grd8.Col = 14
be!e112 = grd8.Text
grd8.Col = 15
be!e212 = grd8.Text
grd8.Col = 16
be!e312 = grd8.Text
grd8.Col = 17
be!mm12 = grd8.Text
'be!obs12 = ne!obs
be.Update
End If
be.Update
End If
If j > 13 Then
grd9.row = 13
grd9.Col = 1
tx4 = grd9.Text
If tx4 = Combo10.Text Then
be!mat13 = tx4
grd8.row = r
grd8.Col = 2
be!cof13 = grd8.Text
grd8.Col = 13
be!md13 = grd8.Text
grd8.Col = 14
be!e113 = grd8.Text
grd8.Col = 15
be!e213 = grd8.Text
grd8.Col = 16
be!e313 = grd8.Text
grd8.Col = 17
be!mm13 = grd8.Text
'be!obs13 = ne!obs
be.Update
End If
be.Update
End If
If j > 14 Then
grd9.row = 14
grd9.Col = 1
tx4 = grd9.Text
If tx4 = Combo10.Text Then
be!mat14 = tx4
grd8.row = r
grd8.Col = 2
be!cof14 = grd8.Text
grd8.Col = 13
be!md14 = grd8.Text
grd8.Col = 14
be!e114 = grd8.Text
grd8.Col = 15
be!e214 = grd8.Text
grd8.Col = 16
be!e314 = grd8.Text
grd8.Col = 17
be!mm14 = grd8.Text
'be!obs14 = ne!obs
be.Update
End If
be.Update
End If
If j > 15 Then
grd9.row = 15
grd9.Col = 1
tx4 = grd9.Text
If tx4 = Combo10.Text Then
be!mat15 = tx4
grd8.row = r
grd8.Col = 2
be!cof15 = grd8.Text
grd8.Col = 13
be!md15 = grd8.Text
grd8.Col = 14
be!e115 = grd8.Text
grd8.Col = 15
be!e215 = grd8.Text
grd8.Col = 16
be!e315 = grd8.Text
grd8.Col = 17
be!mm15 = grd8.Text
'be!obs15 = ne!obs
be.Update
End If
be.Update
End If
If j > 16 Then
grd9.row = 16
grd9.Col = 1
tx4 = grd9.Text
If tx4 = Combo10.Text Then
be!mat16 = tx4
grd8.row = r
grd8.Col = 2
be!cof16 = grd8.Text
grd8.Col = 13
be!md16 = grd8.Text
grd8.Col = 14
be!e116 = grd8.Text
grd8.Col = 15
be!e216 = grd8.Text
grd8.Col = 16
be!e316 = grd8.Text
grd8.Col = 17
be!mm16 = grd8.Text
'be!obs16 = ne!obs
be.Update
End If
be.Update
End If
If j > 17 Then
grd9.row = 17
grd9.Col = 1
tx4 = grd9.Text
If tx4 = Combo10.Text Then
be!mat17 = tx4
grd8.row = r
grd8.Col = 2
be!cof17 = grd8.Text
grd8.Col = 13
be!md17 = grd8.Text
grd8.Col = 14
be!e117 = grd8.Text
grd8.Col = 15
be!e217 = grd8.Text
grd8.Col = 16
be!e317 = grd8.Text
grd8.Col = 17
be!mm17 = grd8.Text
'be!obs17 = ne!obs
be.Update
End If
be.Update
End If
If j > 18 Then
grd9.row = 18
grd9.Col = 1
tx4 = grd9.Text
If tx4 = Combo10.Text Then
be!mat18 = tx4
grd8.row = r
grd8.Col = 2
be!cof18 = grd8.Text
grd8.Col = 13
be!md18 = grd8.Text
grd8.Col = 14
be!e118 = grd8.Text
grd8.Col = 15
be!e218 = grd8.Text
grd8.Col = 16
be!e318 = grd8.Text
grd8.Col = 17
be!mm18 = grd8.Text
'be!obs18 = ne!obs
be.Update
End If
be.Update
End If
If j > 19 Then
grd9.row = 19
grd9.Col = 1
tx4 = grd9.Text
If tx4 = Combo10.Text Then
be!mat19 = tx4
grd8.row = r
grd8.Col = 2
be!cof19 = grd8.Text
grd8.Col = 13
be!md19 = grd8.Text
grd8.Col = 14
be!e119 = grd8.Text
grd8.Col = 15
be!e219 = grd8.Text
grd8.Col = 16
be!e319 = grd8.Text
grd8.Col = 17
be!mm19 = grd8.Text
'be!obs19 = Text17.Text
be.Update
End If
be.Update
End If
If j > 20 Then
grd9.row = 20
grd9.Col = 1
tx4 = grd9.Text
If tx4 = Combo10.Text Then
be!mat20 = tx4
grd8.row = r
grd8.Col = 2
be!cof20 = grd8.Text
grd8.Col = 13
be!md20 = grd8.Text
grd8.Col = 14
be!e120 = grd8.Text
grd8.Col = 15
be!e220 = grd8.Text
grd8.Col = 16
be!e320 = grd8.Text
grd8.Col = 17
be!mm20 = grd8.Text
'be!obs20 = ne!obs
be.Update
End If
be.Update
End If
be.MoveLast
End If
be.MoveNext
Loop
k = 0
Call cont
Do While Not nt.EOF
If nt!cla = Combo11.Text And nt!num = tx1 And nt!mat = Combo10.Text Then
k = 1
grd8.row = r
grd8.Col = 19
nt!ser = grd8.Text
grd8.row = r
'grd8.Col = 1
nt!mat = Combo10.Text
'grd8.Col = 2
'nt!cof = Label104.Caption
grd8.Col = 3
nt!de1 = grd8.Text
grd8.Col = 4
nt!de2 = grd8.Text
grd8.Col = 5
nt!de3 = grd8.Text
grd8.Col = 6
nt!de4 = grd8.Text
grd8.Col = 7
nt!de5 = grd8.Text
grd8.Col = 8
nt!de6 = grd8.Text
grd8.Col = 9
nt!de7 = grd8.Text
grd8.Col = 10
nt!de8 = grd8.Text
grd8.Col = 11
nt!de9 = grd8.Text
grd8.Col = 12
nt!de10 = grd8.Text
grd8.Col = 13
nt!md = grd8.Text
grd8.Col = 14
nt!ex1 = grd8.Text
grd8.Col = 15
nt!ex2 = grd8.Text
grd8.Col = 16
nt!ex3 = grd8.Text
grd8.Col = 17
nt!mm = grd8.Text
grd8.Col = 18
nt!tot = grd8.Text
nt.Update
grd8.Enabled = True
grd8.Visible = True
'Exit Sub
nt.MoveLast
End If
nt.MoveNext
Loop
If k = 0 Then
nt.AddNew
nt!cla = Combo11.Text
nt!num = tx1
nt!nom = tx2
grd8.row = r
grd8.Col = 19
nt!ser = grd8.Text
nt!nbr = Label41.Caption
nt!nme = Label67.Caption
grd8.row = r
'grd8.Col = 1
nt!mat = Combo10.Text
'grd8.Col = 2
nt!cof = Label104.Caption
grd8.Col = 3
nt!de1 = grd8.Text
grd8.Col = 4
nt!de2 = grd8.Text
grd8.Col = 5
nt!de3 = grd8.Text
grd8.Col = 6
nt!de4 = grd8.Text
grd8.Col = 7
nt!de5 = grd8.Text
grd8.Col = 8
nt!de6 = grd8.Text
grd8.Col = 9
nt!de7 = grd8.Text
grd8.Col = 10
nt!de8 = grd8.Text
grd8.Col = 11
nt!de9 = grd8.Text
grd8.Col = 12
nt!de10 = grd8.Text
grd8.Col = 13
nt!md = grd8.Text
grd8.Col = 14
nt!ex1 = grd8.Text
grd8.Col = 15
nt!ex2 = grd8.Text
grd8.Col = 16
nt!ex3 = grd8.Text
grd8.Col = 17
nt!mm = grd8.Text
grd8.Col = 18
nt!tot = grd8.Text
nt.Update
Call calculmoyenne2
grd8.Enabled = True
grd8.Visible = True
End If
'******** bultin
If controle = 0 Then
Call cont2
be.AddNew
be!cla = Combo11.Text
be!mtr = tx1
be!nom = tx2
be!mena = ""
be!menf = ""
be!numm = Label106.Caption
be!eco = face.SBB1.Panels(13).Text
be!ann = face.SBB1.Panels(9).Text
be!moy = ""
be!dat = ""
be!Abs = ""
be!ran = ""
be!ser = ""
be!mat1 = ""
be!cof1 = ""
be!md1 = ""
be!e11 = ""
be!e21 = ""
be!e31 = ""
be!mm1 = ""
be!obs1 = ""
be!mat2 = ""
be!cof2 = ""
be!md2 = ""
be!e12 = ""
be!e22 = ""
be!e32 = ""
be!mm2 = ""
be!obs2 = ""
be!mat3 = ""
be!cof3 = ""
be!md3 = ""
be!e13 = ""
be!e23 = ""
be!e33 = ""
be!mm3 = ""
be!obs3 = ""
be!mat4 = ""
be!cof4 = ""
be!md4 = ""
be!e14 = ""
be!e24 = ""
be!e34 = ""
be!mm4 = ""
be!obs4 = ""
be!mat5 = ""
be!cof5 = ""
be!md5 = ""
be!e15 = ""
be!e25 = ""
be!e35 = ""
be!mm5 = ""
be!obs5 = ""
be!mat6 = ""
be!cof6 = ""
be!md6 = ""
be!e16 = ""
be!e26 = ""
be!e36 = ""
be!mm6 = ""
be!obs6 = ""
be!mat7 = ""
be!cof7 = ""
be!md7 = ""
be!e17 = ""
be!e27 = ""
be!e37 = ""
be!mm7 = ""
be!obs7 = ""
be!mat8 = ""
be!cof8 = ""
be!md8 = ""
be!e18 = ""
be!e28 = ""
be!e38 = ""
be!mm8 = ""
be!obs8 = ""
be!mat9 = ""
be!cof9 = ""
be!md9 = ""
be!e19 = ""
be!e29 = ""
be!e39 = ""
be!mm9 = ""
be!obs9 = ""
be!mat10 = ""
be!cof10 = ""
be!md10 = ""
be!e110 = ""
be!e210 = ""
be!e310 = ""
be!mm10 = ""
be!obs10 = ""
be!mat11 = ""
be!cof11 = ""
be!md11 = ""
be!e111 = ""
be!e211 = ""
be!e311 = ""
be!mm11 = ""
be!obs11 = ""
be!mat12 = ""
be!cof12 = ""
be!md12 = ""
be!e112 = ""
be!e212 = ""
be!e312 = ""
be!mm12 = ""
be!obs12 = ""
be!mat13 = ""
be!cof13 = ""
be!md13 = ""
be!e113 = ""
be!e213 = ""
be!e313 = ""
be!mm13 = ""
be!obs13 = ""
be!mat14 = ""
be!cof14 = ""
be!md14 = ""
be!e114 = ""
be!e214 = ""
be!e314 = ""
be!mm14 = ""
be!obs14 = ""
be!mat15 = ""
be!cof15 = ""
be!md15 = ""
be!e115 = ""
be!e215 = ""
be!e315 = ""
be!mm15 = ""
be!obs15 = ""
be!mat16 = ""
be!cof16 = ""
be!md16 = ""
be!e116 = ""
be!e216 = ""
be!e316 = ""
be!mm16 = ""
be!obs16 = ""
be!mat17 = ""
be!cof17 = ""
be!md17 = ""
be!e117 = ""
be!e217 = ""
be!e317 = ""
be!mm17 = ""
be!obs17 = ""
be!mat18 = ""
be!cof18 = ""
be!md18 = ""
be!e118 = ""
be!e218 = ""
be!e318 = ""
be!mm18 = ""
'be!obs18 = ""
be!mat19 = ""
be!cof19 = ""
be!md19 = ""
be!e119 = ""
be!e219 = ""
be!e319 = ""
be!mm19 = ""
'be!obs19 = ""
be!mat20 = ""
be!cof20 = ""
be!md20 = ""
be!e120 = ""
be!e220 = ""
be!e320 = ""
be!mm20 = ""
be!obs20 = ""
be.Update
be!cla = Combo11.Text
be!mtr = tx1
be!nom = tx2
be!mena = Label40.Caption
be!menf = Label29.Caption
be!numm = Label106.Caption
be!eco = face.SBB1.Panels(13).Text
be!ann = face.SBB1.Panels(9).Text
be!moy = Label33.Caption
be!dat = Date
'be!Abs = ""
'be!ran = ""
grd8.row = r
grd8.Col = 19
be!ser = grd8.Text
If j > 1 Then
grd9.row = 1
grd9.Col = 1
tx4 = grd9.Text
be!mat1 = tx4
If tx4 = Combo10.Text Then
be!mat1 = tx4
grd8.row = r
grd8.Col = 2
be!cof1 = grd8.Text
grd8.Col = 13
be!md1 = grd8.Text
grd8.Col = 14
be!e11 = grd8.Text
grd8.Col = 15
be!e21 = grd8.Text
grd8.Col = 16
be!e31 = grd8.Text
grd8.Col = 17
be!mm1 = grd8.Text
'be!obs1 = ne!obs
be.Update
End If
be.Update
End If
If j > 2 Then
grd9.row = 2
grd9.Col = 1
tx4 = grd9.Text
be!mat2 = tx4
If tx4 = Combo10.Text Then
be!mat2 = tx4
grd8.row = r
grd8.Col = 2
be!cof2 = grd8.Text
grd8.Col = 13
be!md2 = grd8.Text
grd8.Col = 14
be!e12 = grd8.Text
grd8.Col = 15
be!e22 = grd8.Text
grd8.Col = 16
be!e32 = grd8.Text
grd8.Col = 17
be!mm2 = grd8.Text
'be!obs2 = ne!obs
be.Update
End If
be.Update
End If
If j > 3 Then
grd9.row = 3
grd9.Col = 1
tx4 = grd9.Text
be!mat3 = tx4
If tx4 = Combo10.Text Then
be!mat3 = tx4
grd8.row = r
grd8.Col = 2
be!cof3 = grd8.Text
grd8.Col = 13
be!md3 = grd8.Text
grd8.Col = 14
be!e13 = grd8.Text
grd8.Col = 15
be!e23 = grd8.Text
grd8.Col = 16
be!e33 = grd8.Text
grd8.Col = 17
be!mm3 = grd8.Text
'be!obs3 = ne!obs
be.Update
End If
be.Update
End If
If j > 4 Then
grd9.row = 4
grd9.Col = 1
tx4 = grd9.Text
be!mat4 = tx4
If tx4 = Combo10.Text Then
be!mat4 = tx4
grd8.row = r
grd8.Col = 2
be!cof4 = grd8.Text
grd8.Col = 13
be!md4 = grd8.Text
grd8.Col = 14
be!e14 = grd8.Text
grd8.Col = 15
be!e24 = grd8.Text
grd8.Col = 16
be!e34 = grd8.Text
grd8.Col = 17
be!mm4 = grd8.Text
'beobs4 = ne!obs
be.Update
End If
be.Update
End If
If j > 5 Then
grd9.row = 5
grd9.Col = 1
tx4 = grd9.Text
be!mat5 = tx4
If tx4 = Combo10.Text Then
be!mat5 = tx4
grd8.row = r
grd8.Col = 2
be!cof5 = grd8.Text
grd8.Col = 13
be!md5 = grd8.Text
grd8.Col = 14
be!e15 = grd8.Text
grd8.Col = 15
be!e25 = grd8.Text
grd8.Col = 16
be!e35 = grd8.Text
grd8.Col = 17
be!mm5 = grd8.Text
'grd1.Col = 1
'be!obs5 = ne!obs
be.Update
End If
be.Update
End If
If j > 6 Then
grd9.row = 6
grd9.Col = 1
tx4 = grd9.Text
be!mat6 = tx4
If tx4 = Combo10.Text Then
be!mat6 = tx4
grd8.row = r
grd8.Col = 2
be!cof6 = grd8.Text
grd8.Col = 13
be!md6 = grd8.Text
grd8.Col = 14
be!e16 = grd8.Text
grd8.Col = 15
be!e26 = grd8.Text
grd8.Col = 16
be!e36 = grd8.Text
grd8.Col = 17
be!mm6 = grd8.Text
'grd1.Col = 1
'be!obs6 = ne!obs
be.Update
End If
be.Update
End If
If j > 7 Then
grd9.row = 7
grd9.Col = 1
tx4 = grd9.Text
be!mat7 = tx4
If tx4 = Combo10.Text Then
be!mat7 = tx4
grd8.row = r
grd8.Col = 2
be!cof7 = grd8.Text
grd8.Col = 13
be!md7 = grd8.Text
grd8.Col = 14
be!e17 = grd8.Text
grd8.Col = 15
be!e27 = grd8.Text
grd8.Col = 16
be!e37 = grd8.Text
grd8.Col = 17
be!mm7 = grd8.Text
'be!obs7 = ne!obs
be.Update
End If
be.Update
End If
If j > 8 Then
grd9.row = 8
grd9.Col = 1
tx4 = grd9.Text
be!mat8 = tx4
If tx4 = Combo10.Text Then
be!mat8 = tx4
grd8.row = r
grd8.Col = 2
be!cof8 = grd8.Text
grd8.Col = 13
be!md8 = grd8.Text
grd8.Col = 14
be!e18 = grd8.Text
grd8.Col = 15
be!e28 = grd8.Text
grd8.Col = 16
be!e38 = grd8.Text
grd8.Col = 17
be!mm8 = grd8.Text
'be!obs8 = ne!obs
be.Update
End If
be.Update
End If
If j > 9 Then
grd9.row = 9
grd9.Col = 1
tx4 = grd9.Text
be!mat9 = tx4
If tx4 = Combo10.Text Then
be!mat9 = tx4
grd8.row = r
grd8.Col = 2
be!cof9 = grd8.Text
grd8.Col = 13
be!md9 = grd8.Text
grd8.Col = 14
be!e19 = grd8.Text
grd8.Col = 15
be!e29 = grd8.Text
grd8.Col = 16
be!e39 = grd8.Text
grd8.Col = 17
be!mm9 = grd8.Text
'be!obs9 = ne!obs
be.Update
End If
be.Update
End If
If j > 10 Then
grd9.row = 10
grd9.Col = 1
tx4 = grd9.Text
be!mat10 = tx4
If tx4 = Combo10.Text Then
be!mat10 = tx4
grd8.row = r
grd8.Col = 2
be!cof10 = grd8.Text
grd8.Col = 13
be!md10 = grd8.Text
grd8.Col = 14
be!e110 = grd8.Text
grd8.Col = 15
be!e210 = grd8.Text
grd8.Col = 16
be!e310 = grd8.Text
grd8.Col = 17
be!mm10 = grd8.Text
'be!obs10 = ne!obs
be.Update
End If
be.Update
End If
If j > 11 Then
grd9.row = 11
grd9.Col = 1
tx4 = grd9.Text
be!mat11 = tx4
If tx4 = Combo10.Text Then
be!mat11 = tx4
grd8.row = r
grd8.Col = 2
be!cof11 = grd8.Text
grd8.Col = 13
be!md11 = grd8.Text
grd8.Col = 14
be!e111 = grd8.Text
grd8.Col = 15
be!e211 = grd8.Text
grd8.Col = 16
be!e311 = grd8.Text
grd8.Col = 17
be!mm11 = grd8.Text
'be!obs11 = ne!obs
be.Update
End If
be.Update
End If
If j > 12 Then
grd9.row = 12
grd9.Col = 1
tx4 = grd9.Text
be!mat12 = tx4
If tx4 = Combo10.Text Then
be!mat12 = tx4
grd8.row = r
grd8.Col = 2
be!cof12 = grd8.Text
grd8.Col = 13
be!md12 = grd8.Text
grd8.Col = 14
be!e112 = grd8.Text
grd8.Col = 15
be!e212 = grd8.Text
grd8.Col = 16
be!e312 = grd8.Text
grd8.Col = 17
be!mm12 = grd8.Text
'be!obs12 = ne!obs
be.Update
End If
be.Update
End If
If j > 13 Then
grd9.row = 13
grd9.Col = 1
tx4 = grd9.Text
be!mat13 = tx4
If tx4 = Combo10.Text Then
be!mat13 = tx4
grd8.row = r
grd8.Col = 2
be!cof13 = grd8.Text
grd8.Col = 13
be!md13 = grd8.Text
grd8.Col = 14
be!e113 = grd8.Text
grd8.Col = 15
be!e213 = grd8.Text
grd8.Col = 16
be!e313 = grd8.Text
grd8.Col = 17
be!mm13 = grd8.Text
'be!obs13 = ne!obs
be.Update
End If
be.Update
End If
If j > 14 Then
grd9.row = 14
grd9.Col = 1
tx4 = grd9.Text
be!mat14 = tx4
If tx4 = Combo10.Text Then
be!mat14 = tx4
grd8.row = r
grd8.Col = 2
be!cof14 = grd8.Text
grd8.Col = 13
be!md14 = grd8.Text
grd8.Col = 14
be!e114 = grd8.Text
grd8.Col = 15
be!e214 = grd8.Text
grd8.Col = 16
be!e314 = grd8.Text
grd8.Col = 17
be!mm14 = grd8.Text
'be!obs14 = ne!obs
be.Update
End If
be.Update
End If
If j > 15 Then
grd9.row = 15
grd9.Col = 1
tx4 = grd9.Text
be!mat15 = tx4
If tx4 = Combo10.Text Then
be!mat15 = tx4
grd8.row = r
grd8.Col = 2
be!cof15 = grd8.Text
grd8.Col = 13
be!md15 = grd8.Text
grd8.Col = 14
be!e115 = grd8.Text
grd8.Col = 15
be!e215 = grd8.Text
grd8.Col = 16
be!e315 = grd8.Text
grd8.Col = 17
be!mm15 = grd8.Text
'be!obs15 = ne!obs
be.Update
End If
be.Update
End If
If j > 16 Then
grd9.row = 16
grd9.Col = 1
tx4 = grd9.Text
be!mat16 = tx4
If tx4 = Combo10.Text Then
be!mat16 = tx4
grd8.row = r
grd8.Col = 2
be!cof16 = grd8.Text
grd8.Col = 13
be!md16 = grd8.Text
grd8.Col = 14
be!e116 = grd8.Text
grd8.Col = 15
be!e216 = grd8.Text
grd8.Col = 16
be!e316 = grd8.Text
grd8.Col = 17
be!mm16 = grd8.Text
'be!obs16 = ne!obs
be.Update
End If
be.Update
End If
If j > 17 Then
grd9.row = 17
grd9.Col = 1
tx4 = grd9.Text
be!mat17 = tx4
If tx4 = Combo10.Text Then
be!mat17 = tx4
grd8.row = r
grd8.Col = 2
be!cof17 = grd8.Text
grd8.Col = 13
be!md17 = grd8.Text
grd8.Col = 14
be!e117 = grd8.Text
grd8.Col = 15
be!e217 = grd8.Text
grd8.Col = 16
be!e317 = grd8.Text
grd8.Col = 17
be!mm17 = grd8.Text
'be!obs17 = ne!obs
be.Update
End If
be.Update
End If
If j > 18 Then
grd9.row = 18
grd9.Col = 1
tx4 = grd9.Text
be!mat18 = tx4
If tx4 = Combo10.Text Then
be!mat18 = tx4
grd8.row = r
grd8.Col = 2
be!cof18 = grd8.Text
grd8.Col = 13
be!md18 = grd8.Text
grd8.Col = 14
be!e118 = grd8.Text
grd8.Col = 15
be!e218 = grd8.Text
grd8.Col = 16
be!e318 = grd8.Text
grd8.Col = 17
be!mm18 = grd8.Text
'be!obs18 = ne!obs
be.Update
End If
be.Update
End If
If j > 19 Then
grd9.row = 19
grd9.Col = 1
tx4 = grd9.Text
be!mat19 = tx4
If tx4 = Combo10.Text Then
be!mat19 = tx4
grd8.row = r
grd8.Col = 2
be!cof19 = grd8.Text
grd8.Col = 13
be!md19 = grd8.Text
grd8.Col = 14
be!e119 = grd8.Text
grd8.Col = 15
be!e219 = grd8.Text
grd8.Col = 16
be!e319 = grd8.Text
grd8.Col = 17
be!mm19 = grd8.Text
'be!obs19 = Text17.Text
be.Update
End If
be.Update
End If
If j > 20 Then
grd9.row = 20
grd9.Col = 1
tx4 = grd9.Text
be!mat20 = tx4
If tx4 = Combo10.Text Then
be!mat20 = tx4
grd8.row = r
grd8.Col = 2
be!cof20 = grd8.Text
grd8.Col = 13
be!md20 = grd8.Text
grd8.Col = 14
be!e120 = grd8.Text
grd8.Col = 15
be!e220 = grd8.Text
grd8.Col = 16
be!e320 = grd8.Text
grd8.Col = 17
be!mm20 = grd8.Text
'be!obs20 = ne!obs
be.Update
End If
be.Update
End If
End If
End If

End Sub

Private Sub Label29_Click()
On Error Resume Next
Label29.Visible = False
Label40.Visible = True
End Sub

Private Sub Label40_Click()
On Error Resume Next
Label40.Visible = False
Label29.Visible = True

End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
Dim i As Integer
i = SSTab1.Tab
If i = 2 Then
Combo5_Change
Text1.Text = ""
Text1.SetFocus
End If
If i = 1 Then
Command36_Click
End If
End Sub
Private Sub Text1_Change()
On Error Resume Next
Picture4.Visible = False
Picture6.Visible = False
SSTab2.Visible = False
grd1.Clear
grd1.Rows = 1
Label33.Caption = ""
Label29.Caption = ""
Label40.Caption = ""
Label42.Caption = ""
Text17.Text = ""
grd25.Visible = False
Command10.Visible = False
Command43.Visible = False
Command44.Visible = False
End Sub

Private Sub Text1_Click()
On Error Resume Next
Text1_Change
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If Text1.Text <> "" Then
If KeyCode = 13 Then
Command8_Click
End If
End If

End Sub


Private Sub Text11_Change(Index As Integer)
On Error Resume Next
Dim i As Integer
Dim j As Integer
i = Index
grd10.Clear
grd10.Rows = 1
PicFilev = ""
Label114.Caption = ""
Label116.Caption = ""
Label117.Caption = ""
Label118.Caption = ""
Label119.Caption = ""
BarcodeX3.Caption = "000000"
Image4.Picture = LoadPicture(PicFilev)
If Text11(i).Text <> "" Then
For j = 0 To 3
If i <> j Then
Text11(j).Text = ""
End If
Next j
Call recheche
End If

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii <> 8 Then
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End If

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii <> 8 Then
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End If

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
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

Private Sub Text8_KeyPress(KeyAscii As Integer)
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
Call SavePictureToDB2(fName)
Text7.SetFocus
'If Dir$(App.Path & "\aboubekrine.bmp", vbNormal) <> "" Then Kill (App.Path & "\aboubekrine.bmp")
ProgressBar1.Value = 0
Timer1.Enabled = False
End If

End Sub

Private Sub Timer2_Timer()
On Error Resume Next
ProgressBar2.Value = ProgressBar2.Value + 8
If ProgressBar2.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
Call SavePictureToDB(fName)
ProgressBar2.Value = 0
Timer2.Enabled = False
End If

End Sub

Private Sub Timer3_Timer()
On Error Resume Next
Dim ane As String
ProgressBar1.Value = ProgressBar1.Value + 8
If ProgressBar1.Value > 90 Then
Call LoadPictureFromDB2
ProgressBar1.Value = 0
Timer3.Enabled = False
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
'data.DoCmd.ApplyFilter cla
data.DoCmd.Maximize
'data.DoCmd.OpenReport "Bulletin", acViewPreview, , "num =" & a, acWindowNormal, OpenArgs
data.DoCmd.OpenReport "cartes", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
'data.DCount
'data.CloseCurrentDatabase
data.Visible = True
Set data = Nothing

End If

End Sub

Private Sub Timer4_Timer()
On Error Resume Next
Dim ane As String
ProgressBar2.Value = ProgressBar2.Value + 8
If ProgressBar2.Value > 90 Then
Call LoadPictureFromDB3
ProgressBar2.Value = 0
Timer4.Enabled = False
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
'data.DoCmd.ApplyFilter cla
data.DoCmd.Maximize
'data.DoCmd.OpenReport "Bulletin", acViewPreview, , "num =" & a, acWindowNormal, OpenArgs
data.DoCmd.OpenReport "cartes", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
'data.DCount
    'data.CloseCurrentDatabase
data.Visible = True
Set data = Nothing

End If

End Sub
Private Sub chargegrd1()
On Error Resume Next
Dim i As Double
Dim tx1 As String
grd1.Cols = 19
grd1.Rows = 1
grd1.row = 0
grd1.Col = 1
grd1.Text = "«·„«œ…"
grd1.Col = 2
grd1.Text = "÷«—»"
grd1.Col = 3
grd1.Text = "«Œ‹ 1"
grd1.Col = 4
grd1.Text = "«Œ‹ 2"
grd1.Col = 5
grd1.Text = "«Œ‹ 3"
grd1.Col = 6
grd1.Text = "«Œ‹ 4"
grd1.Col = 7
grd1.Text = "«Œ‹ 5"
grd1.Col = 8
grd1.Text = "«Œ‹ 6"
grd1.Col = 9
grd1.Text = "«Œ‹ 7"
grd1.Col = 10
grd1.Text = "«Œ‹ 8"
grd1.Col = 11
grd1.Text = "«Œ‹ 9"
grd1.Col = 12
grd1.Text = "«Œ‹ 10"
grd1.Col = 13
grd1.Text = "„⁄œ· «·«Œ‹"
grd1.Col = 14
grd1.Text = "«„‹ 1"
grd1.Col = 15
grd1.Text = "«„‹ 2"
grd1.Col = 16
grd1.Text = "«„‹ 3"
grd1.Col = 17
grd1.Text = "„⁄œ· «·„«œ…"
grd1.Col = 18
grd1.Text = "«·„Ã„Ê⁄"
For i = 0 To 18
grd1.ColWidth(i) = 600
grd1.ColAlignment(i) = 1
If i = 0 Then
grd1.ColWidth(i) = 0
End If
If i = 1 Then
grd1.ColWidth(i) = 2300
End If
If i = 13 Then
grd1.ColWidth(i) = 1000
End If
If i = 17 Then
grd1.ColWidth(i) = 1000
End If
If i = 18 Then
grd1.ColWidth(i) = 900
End If
Next i
i = 1
Call cont
grd1.Rows = mt.RecordCount + 3
Do While Not mt.EOF
If mt!cla = Label36.Caption Then
Label41.Caption = mt!nbr
Label67.Caption = mt!nme
grd1.row = i
grd1.Col = 1
grd1.Text = mt!mat
grd1.Col = 2
grd1.Text = mt!cof
i = i + 1
End If
mt.MoveNext
Loop
grd1.Rows = i
n = grd1.Rows
Call cont
Do While Not nt.EOF
For i = 1 To n - 1
grd1.row = i
grd1.Col = 1
tx1 = grd1.Text
If nt!cla = Label36.Caption And nt!num = Label37.Caption And nt!mat = tx1 Then
grd1.row = i
grd1.Col = 3
grd1.Text = nt!de1
grd1.Col = 4
grd1.Text = nt!de2
grd1.Col = 5
grd1.Text = nt!de3
grd1.Col = 6
grd1.Text = nt!de4
grd1.Col = 7
grd1.Text = nt!de5
grd1.Col = 8
grd1.Text = nt!de6
grd1.Col = 9
grd1.Text = nt!de7
grd1.Col = 10
grd1.Text = nt!de8
grd1.Col = 11
grd1.Text = nt!de9
grd1.Col = 12
grd1.Text = nt!de10
grd1.Col = 13
grd1.Text = nt!md
grd1.Col = 14
grd1.Text = nt!ex1
grd1.Col = 15
grd1.Text = nt!ex2
grd1.Col = 16
grd1.Text = nt!ex3
grd1.Col = 17
grd1.Text = nt!mm
grd1.Col = 18
grd1.Text = nt!tot
i = n
End If
Next i
nt.MoveNext
Loop
'Call cont2
'Do While Not be.EOF
'If be!cla = Label36.Caption And be!mtr = Label37.Caption Then
'Text17.Text = be!obs19
'Exit Sub
'End If
'be.MoveNext
'Loop

End Sub
Private Sub chargegrd8()
On Error Resume Next
Dim i As Double
Dim k As Double
Dim m As Double
Dim l As Double
Dim tx1 As String
grd8.Cols = 20
grd8.Rows = 1
grd8.row = 0
grd8.Col = 0
grd8.Text = ".«·—ﬁ„"
grd8.Col = 1
grd8.Text = ".«·«”„"
grd8.Col = 2
grd8.Text = "÷«—»"
grd8.Col = 3
grd8.Text = "«Œ‹ 1"
grd8.Col = 4
grd8.Text = "«Œ‹ 2"
grd8.Col = 5
grd8.Text = "«Œ‹ 3"
grd8.Col = 6
grd8.Text = "«Œ‹ 4"
grd8.Col = 7
grd8.Text = "«Œ‹ 5"
grd8.Col = 8
grd8.Text = "«Œ‹ 6"
grd8.Col = 9
grd8.Text = "«Œ‹ 7"
grd8.Col = 10
grd8.Text = "«Œ‹ 8"
grd8.Col = 11
grd8.Text = "«Œ‹ 9"
grd8.Col = 12
grd8.Text = "«Œ‹ 10"
grd8.Col = 13
grd8.Text = "„⁄œ· «·«Œ‹"
grd8.Col = 14
grd8.Text = "«„‹ 1"
grd8.Col = 15
grd8.Text = "«„‹ 2"
grd8.Col = 16
grd8.Text = "«„‹ 3"
grd8.Col = 17
grd8.Text = "„⁄œ· «·„«œ…"
grd8.Col = 18
grd8.Text = "«·„Ã„Ê⁄"
For i = 0 To 19
grd8.ColWidth(i) = 600
grd8.ColAlignment(i) = 1
If i = 0 Then
grd8.ColWidth(i) = 700
End If
If i = 1 Then
grd8.ColWidth(i) = 1900
End If
If i = 13 Then
grd8.ColWidth(i) = 800
End If
If i = 17 Then
grd8.ColWidth(i) = 800
End If
If i = 18 Then
grd8.ColWidth(i) = 900
End If
If i = 19 Then
grd8.ColWidth(i) = 0
End If
Next i
i = 1
k = 0
Call cont
grd8.Rows = et.RecordCount + 3
Do While Not et.EOF
If et!num < 1000000 Then
If et!cla = Combo11.Text Then
grd8.row = i
grd8.Col = 0
grd8.Text = et!num
grd8.Col = 1
grd8.Text = et!nom
grd8.Col = 2
grd8.Text = Label104.Caption
grd8.Col = 19
grd8.Text = et!ser
i = i + 1
'***** grd30
k = k + 1
grd30.row = 0
grd30.Col = 0
grd30.Text = k
'***** end grd30
End If
End If
et.MoveNext
Loop
m = 0
grd8.Rows = i
grd8.Col = 0
grd8.Sort = 1
n = grd8.Rows
Call cont
nt.Filter = "[cla]" & "Like '*" & Combo11 & "*'"
nt.Filter = "[mat]" & "Like '*" & Combo10 & "*'"
Do While Not nt.EOF
'k = 0
For i = 1 To n - 1
grd8.row = i
grd8.Col = 0
tx1 = grd8.Text
If nt!cla = Combo11.Text And nt!num = tx1 And nt!mat = Combo10.Text Then
'MsgBox ""
grd8.row = i
grd8.Col = 3
grd8.Text = nt!de1
grd8.Col = 4
grd8.Text = nt!de2
grd8.Col = 5
grd8.Text = nt!de3
grd8.Col = 6
grd8.Text = nt!de4
grd8.Col = 7
grd8.Text = nt!de5
grd8.Col = 8
grd8.Text = nt!de6
grd8.Col = 9
grd8.Text = nt!de7
grd8.Col = 10
grd8.Text = nt!de8
grd8.Col = 11
grd8.Text = nt!de9
grd8.Col = 12
grd8.Text = nt!de10
grd8.Col = 13
grd8.Text = nt!md
grd8.Col = 14
grd8.Text = nt!ex1
grd8.Col = 15
grd8.Text = nt!ex2
grd8.Col = 16
grd8.Text = nt!ex3
grd8.Col = 17
grd8.Text = nt!mm
grd8.Col = 18
grd8.Text = nt!tot
i = n
End If
'***** grd30
m = m + 1
l = m / 1000
l = l * 2
l = l / k
l = l * 100
MyNumber = Round(l, 2)
l = MyNumber
grd30.row = 0
grd30.Col = 0
grd30.Text = l
'***** end grd30
Next i
nt.MoveNext
Loop

End Sub

Private Sub calculmoyenne()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim c As Double
Dim cs As Double
Dim d1 As Double
Dim d2 As Double
Dim d3 As Double
Dim d4 As Double
Dim d5 As Double
Dim d6 As Double
Dim d7 As Double
Dim d8 As Double
Dim d9 As Double
Dim d10 As Double
Dim ds As Double
Dim m1 As Double
Dim e1 As Double
Dim e2 As Double
Dim e3 As Double
Dim m2 As Double
Dim t As Double
Dim ts As Double
Dim j As Double
Dim cd As Double
Dim ce1 As Double
Dim ce2 As Double
Dim ce3 As Double
Dim tc As Double
Dim moy As Double
ts = 0
cs = 0
moy = 0
n = grd1.Rows
For i = 1 To n - 1
t = 0
c = 0
j = 0
ds = 0
cd = 0
ce1 = 0
ce2 = 0
ce3 = 0
e1 = 0
e2 = 0
e3 = 0
d1 = 0
d2 = 0
d3 = 0
d4 = 0
d5 = 0
d6 = 0
d7 = 0
d8 = 0
d9 = 0
d10 = 0
m1 = 0
m2 = 0
grd1.row = i
grd1.Col = 3
If Val(grd1.Text) > 0 Or Val(grd1.Text) = 0 And grd1.Text = "0" Then
d1 = grd1.Text
j = j + 1
ds = ds + d1
End If
grd1.Col = 4
If Val(grd1.Text) > 0 Or Val(grd1.Text) = 0 And grd1.Text = "0" Then
d2 = grd1.Text
j = j + 1
ds = ds + d2
End If
grd1.Col = 5
If Val(grd1.Text) > 0 Or Val(grd1.Text) = 0 And grd1.Text = "0" Then
d3 = grd1.Text
j = j + 1
ds = ds + d3
End If
grd1.Col = 6
If Val(grd1.Text) > 0 Or Val(grd1.Text) = 0 And grd1.Text = "0" Then
d4 = grd1.Text
j = j + 1
ds = ds + d4
End If
grd1.Col = 7
If Val(grd1.Text) > 0 Or Val(grd1.Text) = 0 And grd1.Text = "0" Then
d5 = grd1.Text
j = j + 1
ds = ds + d5
End If
grd1.Col = 8
If Val(grd1.Text) > 0 Or Val(grd1.Text) = 0 And grd1.Text = "0" Then
d6 = grd1.Text
j = j + 1
ds = ds + d6
End If
grd1.Col = 9
If Val(grd1.Text) > 0 Or Val(grd1.Text) = 0 And grd1.Text = "0" Then
d7 = grd1.Text
j = j + 1
ds = ds + d7
End If
grd1.Col = 10
If Val(grd1.Text) > 0 Or Val(grd1.Text) = 0 And grd1.Text = "0" Then
d8 = grd1.Text
j = j + 1
ds = ds + d8
End If
grd1.Col = 11
If Val(grd1.Text) > 0 Or Val(grd1.Text) = 0 And grd1.Text = "0" Then
d9 = grd1.Text
j = j + 1
ds = ds + d9
End If
grd1.Col = 12
If Val(grd1.Text) > 0 Or Val(grd1.Text) = 0 And grd1.Text = "0" Then
d10 = grd1.Text
j = j + 1
ds = ds + d10
End If
grd1.row = i
grd1.Col = 13
grd1.Text = ""
If j > 0 Then
ds = ds / j
MyNumber = Round(ds, 2)
m1 = MyNumber
grd1.row = i
grd1.Col = 13
grd1.Text = m1
cd = Label24.Caption
End If
grd1.row = i
grd1.Col = 14
If Val(grd1.Text) > 0 Or Val(grd1.Text) = 0 And grd1.Text = "0" Then
e1 = grd1.Text
ce1 = Label25.Caption
End If
grd1.Col = 15
If Val(grd1.Text) > 0 Or Val(grd1.Text) = 0 And grd1.Text = "0" Then
e2 = grd1.Text
ce2 = Label26.Caption
End If
grd1.Col = 16
If Val(grd1.Text) > 0 Or Val(grd1.Text) = 0 And grd1.Text = "0" Then
e3 = grd1.Text
ce3 = Label27.Caption
End If
tc = (cd + ce1 + ce2 + ce3)
grd1.row = i
grd1.Col = 17
grd1.Text = ""
grd1.Col = 18
grd1.Text = ""
If tc > 0 Then
m2 = ((m1 * cd) + (e1 * ce1) + (e2 * ce2) + (e3 * ce3)) / tc
MyNumber = Round(m2, 2)
m2 = MyNumber
grd1.row = i
grd1.Col = 17
grd1.Text = m2
grd1.Col = 2
c = grd1.Text
cs = cs + c
t = (m2 * c)
grd1.row = i
grd1.Col = 18
grd1.Text = t
ts = ts + t
End If
Next i
If cs > 0 Then
moy = ts / cs
MyNumber = Round(moy, 2)
moy = MyNumber
End If
Label33.Caption = moy
Call coffes
a = 0
f = 6
For i = 5 To 15
If moy = 0 And cs = 0 Then
Label40.Caption = ""
Label29.Caption = ""
Label33.Caption = ""
i = 15
Exit Sub
End If
If moy = 0 And cs > 0 Then
Label40.Caption = mens(5).Text
Label29.Caption = mens(11).Text
i = 15
Exit Sub
End If
cs = coff(i - 1).Text
ts = coff(i).Text
If moy > ts And moy <= cs Then
Label40.Caption = mens(a).Text
Label29.Caption = mens(f).Text
i = 15
End If
If i = 5 Or i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 15 Then
a = a + 1
f = f + 1
End If
Next i
End Sub
Private Sub calculmoyenne2()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim c As Double
Dim cs As Double
Dim d1 As Double
Dim d2 As Double
Dim d3 As Double
Dim d4 As Double
Dim d5 As Double
Dim d6 As Double
Dim d7 As Double
Dim d8 As Double
Dim d9 As Double
Dim d10 As Double
Dim ds As Double
Dim m1 As Double
Dim e1 As Double
Dim e2 As Double
Dim e3 As Double
Dim m2 As Double
Dim t As Double
Dim ts As Double
Dim j As Double
Dim cd As Double
Dim ce1 As Double
Dim ce2 As Double
Dim ce3 As Double
Dim tc As Double
Dim moy As Double
Dim r As Double
ts = 0
cs = 0
moy = 0
n = grd8.Rows
i = grd8.row
'For i = 1 To n - 1
t = 0
c = 0
j = 0
ds = 0
cd = 0
ce1 = 0
ce2 = 0
ce3 = 0
e1 = 0
e2 = 0
e3 = 0
d1 = 0
d2 = 0
d3 = 0
d4 = 0
d5 = 0
d6 = 0
d7 = 0
d8 = 0
d9 = 0
d10 = 0
m1 = 0
m2 = 0
grd8.row = i
grd8.Col = 3
If Val(grd8.Text) > 0 Or Val(grd8.Text) = 0 And grd8.Text = "0" Then
d1 = grd8.Text
j = j + 1
ds = ds + d1
End If
grd8.Col = 4
If Val(grd8.Text) > 0 Or Val(grd8.Text) = 0 And grd8.Text = "0" Then
d2 = grd8.Text
j = j + 1
ds = ds + d2
End If
grd8.Col = 5
If Val(grd8.Text) > 0 Or Val(grd8.Text) = 0 And grd8.Text = "0" Then
d3 = grd8.Text
j = j + 1
ds = ds + d3
End If
grd8.Col = 6
If Val(grd8.Text) > 0 Or Val(grd8.Text) = 0 And grd8.Text = "0" Then
d4 = grd8.Text
j = j + 1
ds = ds + d4
End If
grd8.Col = 7
If Val(grd8.Text) > 0 Or Val(grd8.Text) = 0 And grd8.Text = "0" Then
d5 = grd8.Text
j = j + 1
ds = ds + d5
End If
grd8.Col = 8
If Val(grd8.Text) > 0 Or Val(grd8.Text) = 0 And grd8.Text = "0" Then
d6 = grd8.Text
j = j + 1
ds = ds + d6
End If
grd8.Col = 9
If Val(grd8.Text) > 0 Or Val(grd8.Text) = 0 And grd8.Text = "0" Then
d7 = grd8.Text
j = j + 1
ds = ds + d7
End If
grd8.Col = 10
If Val(grd8.Text) > 0 Or Val(grd8.Text) = 0 And grd8.Text = "0" Then
d8 = grd8.Text
j = j + 1
ds = ds + d8
End If
grd8.Col = 11
If Val(grd8.Text) > 0 Or Val(grd8.Text) = 0 And grd8.Text = "0" Then
d9 = grd8.Text
j = j + 1
ds = ds + d9
End If
grd8.Col = 12
If Val(grd8.Text) > 0 Or Val(grd8.Text) = 0 And grd8.Text = "0" Then
d10 = grd8.Text
j = j + 1
ds = ds + d10
End If
grd8.row = i
grd8.Col = 13
grd8.Text = ""
If j > 0 Then
ds = ds / j
MyNumber = Round(ds, 2)
m1 = MyNumber
grd8.row = i
grd8.Col = 13
grd8.Text = m1
cd = Label24.Caption
End If
grd8.row = i
grd8.Col = 14
If Val(grd8.Text) > 0 Or Val(grd8.Text) = 0 And grd8.Text = "0" Then
e1 = grd8.Text
ce1 = Label25.Caption
End If
grd8.Col = 15
If Val(grd8.Text) > 0 Or Val(grd8.Text) = 0 And grd8.Text = "0" Then
e2 = grd8.Text
ce2 = Label26.Caption
End If
grd8.Col = 16
If Val(grd8.Text) > 0 Or Val(grd8.Text) = 0 And grd8.Text = "0" Then
e3 = grd8.Text
ce3 = Label27.Caption
End If
tc = (cd + ce1 + ce2 + ce3)
grd8.row = i
grd8.Col = 17
grd8.Text = ""
grd8.Col = 18
grd8.Text = ""
If tc > 0 Then
m2 = ((m1 * cd) + (e1 * ce1) + (e2 * ce2) + (e3 * ce3)) / tc
MyNumber = Round(m2, 2)
m2 = MyNumber
grd8.row = i
grd8.Col = 17
grd8.Text = m2
grd8.Col = 2
c = grd8.Text
cs = cs + c
t = (m2 * c)
grd8.row = i
grd8.Col = 18
grd8.Text = t
ts = ts + t
End If
'Next i
End Sub
Public Sub coffes1()
On Error Resume Next
Call cont
coff(0).Text = cf1!cof0
coff(1).Text = cf1!cof1
coff(2).Text = cf1!cof2
coff(3).Text = cf1!cof3
coff(4).Text = cf1!cof4
coff(5).Text = cf1!cof5
coff(6).Text = cf1!cof6
coff(7).Text = cf1!cof7
coff(8).Text = cf1!cof8
coff(9).Text = cf1!cof9
coff(10).Text = cf1!cof10
coff(11).Text = cf1!cof11
coff(12).Text = cf1!cof12
coff(13).Text = cf1!cof13
coff(14).Text = cf1!cof14
coff(15).Text = cf1!cof15
mens(0).Text = cf1!tex9
mens(1).Text = cf1!tex12
mens(2).Text = cf1!tex15
mens(3).Text = cf1!tex18
mens(4).Text = cf1!tex19
mens(5).Text = cf1!tex20
mens(6).Text = cf1!tex21
mens(7).Text = cf1!tex22
mens(8).Text = cf1!tex23
mens(9).Text = cf1!tex24
mens(10).Text = cf1!tex25
mens(11).Text = cf1!tex26
End Sub
Public Sub coffes()
On Error Resume Next
Call cont
coff(0).Text = cf!cof0
coff(1).Text = cf!cof1
coff(2).Text = cf!cof2
coff(3).Text = cf!cof3
Label24.Caption = cf!cof0
Label25.Caption = cf!cof1
Label26.Caption = cf!cof2
Label27.Caption = cf!cof3
coff(4).Text = cf!cof4
coff(5).Text = cf!cof5
coff(6).Text = cf!cof6
coff(7).Text = cf!cof7
coff(8).Text = cf!cof8
coff(9).Text = cf!cof9
coff(10).Text = cf!cof10
coff(11).Text = cf!cof11
coff(12).Text = cf!cof12
coff(13).Text = cf!cof13
coff(14).Text = cf!cof14
coff(15).Text = cf!cof15
mens(0).Text = cf!tex9
mens(1).Text = cf!tex12
mens(2).Text = cf!tex15
mens(3).Text = cf!tex18
mens(4).Text = cf!tex19
mens(5).Text = cf!tex20
mens(6).Text = cf!tex21
mens(7).Text = cf!tex22
mens(8).Text = cf!tex23
mens(9).Text = cf!tex24
mens(10).Text = cf!tex25
mens(11).Text = cf!tex26
End Sub

Private Sub chargegrd2()
On Error Resume Next
Dim i As Double
Dim a As Double
Dim b As Double
grd2.row = 0
grd2.Col = 0
grd2.Text = ""
grd2.Col = 1
grd2.Text = "«· «—ÌŒ"
grd2.Col = 2
grd2.Text = "„‰"
grd2.Col = 3
grd2.Text = "≈·Ï"
grd2.Col = 4
grd2.Text = "⁄œœ «·”«⁄« "
grd2.Col = 5
grd2.Text = "«·„«œ…"
grd2.Col = 6
grd2.Text = "«·„·«ÕŸ…"
i = 1
b = 0
Call cont
grd2.Rows = ab.RecordCount + 2
Do While Not ab.EOF
If ab!cla = Label36.Caption And ab!num = Label37.Caption Then
a = ab!nbr
b = b + a
grd2.row = i
grd2.Col = 0
grd2.Text = ab!aut
grd2.Col = 1
grd2.Text = ab!dat
grd2.Col = 2
grd2.Text = ab!hr1
grd2.Col = 3
grd2.Text = ab!hr2
grd2.Col = 4
grd2.Text = ab!nbr
grd2.Col = 5
grd2.Text = ab!mat
grd2.Col = 6
grd2.Text = ab!rem
i = i + 1
End If
ab.MoveNext
Loop
grd2.Rows = i
Label73.Caption = b
Label74.Caption = b

End Sub

Private Sub Timer6_Timer()
On Error Resume Next
ProgressBar4.Value = ProgressBar4.Value + 8
If ProgressBar4.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
Label87.Caption = ""
grd2.Clear
grd2.Rows = 1
grd2.Visible = False
Call chargegrd2
grd2.Visible = True
ProgressBar4.Value = 0
Timer6.Enabled = False
End If

End Sub

Private Sub Timer7_Timer()
On Error Resume Next
ProgressBar1.Value = ProgressBar1.Value + 8
If ProgressBar1.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
ProgressBar1.Value = 0
Timer7.Enabled = False
Text1.Text = ""
End If

End Sub
Private Sub chargegrd6()
On Error Resume Next
Dim i As Double
Dim j As Double
grd6.Clear
grd6.Rows = 1
grd6.Cols = 4
grd6.ColWidth(0) = 0
grd6.ColWidth(1) = 1200
grd6.ColWidth(2) = 3500
grd6.ColWidth(3) = 1500
grd6.ColAlignment(1) = 1
grd6.ColAlignment(2) = 1
grd6.ColAlignment(3) = 1
grd6.row = 0
grd6.Col = 1
grd6.Text = "«·—ﬁ„"
grd6.Col = 2
grd6.Text = "«·«”„"
grd6.Col = 3
grd6.Text = "«·—ﬁ„ «· ”·”·Ì"
i = 1
Call cont
grd6.Rows = et.RecordCount + 30
Do While Not et.EOF
If Combo11.Text = et!cla Then
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
End Sub
Private Sub chargegrd25()
On Error Resume Next
Dim i As Double
Dim j As Double
grd25.Clear
grd25.Rows = 1
grd25.Cols = 6
grd25.ColWidth(0) = 0
grd25.ColWidth(1) = 1500
grd25.ColWidth(2) = 3500
grd25.ColWidth(3) = 1500
grd25.ColWidth(4) = 0
grd25.ColWidth(5) = 0
grd25.ColAlignment(1) = 1
grd25.ColAlignment(2) = 1
grd25.ColAlignment(3) = 1
grd25.row = 0
grd25.Col = 1
grd25.Text = "«·—ﬁ„"
grd25.Col = 2
grd25.Text = "«·«”„"
grd25.Col = 3
grd25.Text = "«·—ﬁ„ «· ”·”·Ì"
i = 1
Call cont
grd25.Rows = et.RecordCount + 30
Do While Not et.EOF
If Combo5.Text = et!cla Then
If Val(et!num) < 1000000 Then
grd25.row = i
grd25.Col = 0
grd25.Text = et!aut
grd25.Col = 1
grd25.Text = et!num
grd25.Col = 2
grd25.Text = et!nom
grd25.Col = 3
grd25.Text = et!ser
grd25.Col = 4
grd25.Text = et!tel
grd25.Col = 5
grd25.Text = et!adr
i = i + 1
End If
End If
et.MoveNext
Loop
grd25.Rows = i
grd25.Col = 1
grd25.Sort = 1
End Sub
Private Sub chargegrd25_2()
On Error Resume Next
Dim i As Double
Dim j As Double
grd25.Clear
grd25.Rows = 1
grd25.Cols = 6
grd25.ColWidth(0) = 0
grd25.ColWidth(1) = 3000
grd25.ColWidth(2) = 1200
grd25.ColWidth(3) = 1200
grd25.ColWidth(4) = 1100
grd25.ColWidth(5) = 0
grd25.ColAlignment(1) = 1
grd25.ColAlignment(2) = 1
grd25.ColAlignment(3) = 1
grd25.row = 0
grd25.Col = 1
grd25.Text = "«·«”„"
grd25.Col = 2
grd25.Text = "«·—ﬁ„ «·Êÿ‰Ì"
grd25.Col = 3
grd25.Text = " «—ÌŒ «·„Ì·«œ"
grd25.Col = 4
grd25.Text = "„Õ· «·„Ì·«œ"
i = 1
Call cont2
grd25.Rows = nn.RecordCount + 30
Do While Not nn.EOF
If Combo5.Text = nn!cla Then
grd25.row = i
grd25.Col = 0
grd25.Text = nn!num
grd25.Col = 1
grd25.Text = nn!nom
grd25.Col = 2
grd25.Text = nn!nni
grd25.Col = 3
grd25.Text = nn!dat
grd25.Col = 4
grd25.Text = nn!liu
grd25.Col = 5
grd25.Text = nn!tel
i = i + 1
End If
nn.MoveNext
Loop
grd25.Rows = i
grd25.Col = 1
grd25.Sort = 0


End Sub
Private Sub Timer8_Timer()
On Error Resume Next
Dim ane As String
ProgressBar5.Value = ProgressBar5.Value + 8
If ProgressBar5.Value > 90 Then
Timer8.Enabled = False
ProgressBar5.Value = 0
Call cont2
If tim = 1 Then
'End If
'FileCopy "C:\photos\4.jpg", "C:\photos\2.jpg"
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
'data.DoCmd.ApplyFilter cla
data.DoCmd.Maximize
'data.DoCmd.OpenReport "Bulletin", acViewPreview, , "num =" & a, acWindowNormal, OpenArgs
data.DoCmd.OpenReport "cartess", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
'data.DCount
    'data.CloseCurrentDatabase
data.Visible = True
Set data = Nothing
End If
If tim = 2 Then
ane = "C" + face.SBB1.Panels(9).Text
data.OpenCurrentDatabase App.Path & "\" & ane & ".mdb", False, "7346804"
'data.DoCmd.ApplyFilter cla
data.DoCmd.Maximize
'data.DoCmd.OpenReport "Bulletin", acViewPreview, , "num =" & a, acWindowNormal, OpenArgs
data.DoCmd.OpenReport "Tnotesmat", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
'data.DCount
    'data.CloseCurrentDatabase
data.Visible = True
'Set data = Nothing
Command2.Enabled = True
End If
End If

End Sub
Public Sub recheche()
On Error Resume Next
Dim i As Double
Label114.Caption = ""
grd10.Clear
grd10.Rows = 1
grd10.Cols = 4
grd10.ColWidth(0) = 1000
grd10.ColWidth(1) = 1000
grd10.ColWidth(2) = 3500
grd10.ColWidth(3) = 1400
grd10.ColAlignment(1) = 1
grd10.ColAlignment(2) = 1
grd10.ColAlignment(3) = 1
grd10.row = 0
grd10.Col = 0
grd10.Text = "«·—ﬁ„"
grd10.Col = 1
grd10.Text = "«·ﬁ”„"
grd10.Col = 2
grd10.Text = "«·«”„"
grd10.Col = 3
grd10.Text = "«·—ﬁ„ «· ”·”·Ì"
 ' **  **  **  ** chargemant des donnÈes **  **  **  **
Call cont
 i = 1
'***** recherch par nom
If Text11(0).Text <> "" Then
n = et.RecordCount
grd10.Rows = n + 2
et.Filter = "[nom]" & "Like '*" & Text11(0) & "*'" 'entre
Do While Not et.EOF
If Val(et!num) < 1000000 Then
If Val(et!act) = 1 Then
grd10.row = i
grd10.Col = 0
grd10.Text = et!num
grd10.Col = 1
grd10.Text = et!cla
grd10.Col = 2
grd10.Text = et!nom
grd10.Col = 3
grd10.Text = et!ser
i = i + 1
End If
End If
et.MoveNext
Loop
End If
'***** recherch par num
If Text11(1).Text <> "" Then
n = et.RecordCount
grd10.Rows = n + 2
et.Filter = "[num]" & "Like '*" & Text11(1) & "*'" 'entre
Do While Not et.EOF
If Val(et!num) < 1000000 Then
If Val(et!act) = 1 Then
grd10.row = i
grd10.Col = 0
grd10.Text = et!num
grd10.Col = 1
grd10.Text = et!cla
grd10.Col = 2
grd10.Text = et!nom
grd10.Col = 3
grd10.Text = et!ser
i = i + 1
End If
End If
et.MoveNext
Loop
End If
'***** recherch par ser
If Text11(2).Text <> "" Then
n = et.RecordCount
grd10.Rows = n + 2
et.Filter = "[ser]" & "Like '*" & Text11(2) & "*'" 'entre
Do While Not et.EOF
If Val(et!num) < 1000000 Then
If Val(et!act) = 1 Then
grd10.row = i
grd10.Col = 0
grd10.Text = et!num
grd10.Col = 1
grd10.Text = et!cla
grd10.Col = 2
grd10.Text = et!nom
grd10.Col = 3
grd10.Text = et!ser
i = i + 1
End If
End If
et.MoveNext
Loop
End If
'***** recherch par tel
If Text11(3).Text <> "" Then
n = et.RecordCount
grd10.Rows = n + 2
et.Filter = "[tel]" & "Like '*" & Text11(3) & "*'" 'entre
Do While Not et.EOF
If Val(et!num) < 1000000 Then
If Val(et!act) = 1 Then
grd10.row = i
grd10.Col = 0
grd10.Text = et!num
grd10.Col = 1
grd10.Text = et!cla
grd10.Col = 2
grd10.Text = et!nom
grd10.Col = 3
grd10.Text = et!ser
i = i + 1
End If
End If
et.MoveNext
Loop
End If
'***** recherch par dat
If Text11(0).Text = "" And Text11(1).Text = "" And Text11(2).Text = "" And Text11(3).Text = "" Then
n = et.RecordCount
grd10.Rows = n + 2
et.Filter = "[dat]" & "Like '*" & DT4.Value & "*'" 'entre
Do While Not et.EOF
If Val(et!num) < 1000000 Then
If Val(et!act) = 1 Then
grd10.row = i
grd10.Col = 0
grd10.Text = et!num
grd10.Col = 1
grd10.Text = et!cla
grd10.Col = 2
grd10.Text = et!nom
grd10.Col = 3
grd10.Text = et!ser
i = i + 1
End If
End If
et.MoveNext
Loop
End If
'**** end
grd10.Rows = i
grd10.Col = 0
grd10.Sort = 1
grd10.ColAlignment(2) = 1
Label114.Caption = i - 1
'pro2.MoveFirst
'****************************************
End Sub
Private Sub changementdecoffcients()
On Error Resume Next
Dim cd1 As Double
Dim ce1 As Double
Dim ce2 As Double
Dim ce3 As Double
Dim md1 As Double
Dim e1 As Double
Dim e2 As Double
Dim e3 As Double
Dim mm1 As Double
Dim tos As Double
Dim cm1 As Double
Dim sc As Double
Dim sb As Double
Dim b As Double
Dim n As Double
Dim u As Double
Dim y As Double
Dim c As Double
Dim tx1 As String
Dim tx2 As String
Dim tx3 As String
Dim tx4 As String
Dim tx5 As String
md1 = 0
e1 = 0
e2 = 0
e3 = 0
grd22.Clear
grd22.Rows = 1
grd22.Cols = 4
grd22.Visible = False
grd23.Clear
grd23.Rows = 1
grd23.Cols = 1
grd23.ColWidth(0) = 2500
grd23.ColAlignment(0) = 1
grd23.row = 0
grd23.Col = 0
grd23.Text = "«·⁄„·Ì… ﬁœ  √Œ– Êﬁ « ·–« Ì—ÃÏ «·«‰ Ÿ«—"
Call cont
Do While Not nt.EOF
sc = 0
If Val(nt!md) = 0 And nt!md <> "0" Then
cd1 = 0
Else
cd1 = coff(0).Text
md1 = nt!md
End If
If Val(nt!ex1) = 0 And nt!ex1 <> "0" Then
ce1 = 0
Else
ce1 = coff(1).Text
e1 = nt!ex1
End If
If Val(nt!ex2) = 0 And nt!ex2 <> "0" Then
ce2 = 0
Else
ce2 = coff(2).Text
e2 = nt!ex2
End If
If Val(nt!ex3) = 0 And nt!ex3 <> "0" Then
ce3 = 0
Else
ce3 = coff(3).Text
e3 = nt!ex3
End If
cm1 = nt!cof
sc = (cd1 + ce1 + ce2 + ce3)
If sc > 0 Then
mm1 = (((md1 * cd1) + (e1 * ce1) + (e2 * ce2) + (e3 * ce3)) / sc)
tos = (mm1 * cm1)
MyNumber = Round(mm1, 2)
nt!mm = MyNumber
MyNumber = Round(tos, 2)
nt!tot = MyNumber
nt.Update
n = grd22.Rows
grd22.Rows = n + 1
grd22.row = n
grd22.Col = 0
grd22.Text = nt!ser
grd22.Col = 1
grd22.Text = nt!mat
MyNumber = Round(mm1, 2)
grd22.Col = 2
grd22.Text = MyNumber
grd22.Col = 3
grd22.Text = nt!cof
End If
nt.MoveNext
Loop
grd23.Clear
grd23.Rows = 1
grd23.Cols = 4
grd23.ColWidth(0) = 1300
grd23.ColWidth(1) = 700
grd23.ColWidth(2) = 1300
grd23.ColWidth(3) = 1000
grd23.ColAlignment(0) = 1
grd23.ColAlignment(1) = 3
grd23.ColAlignment(2) = 2
grd23.ColAlignment(3) = 1
sc = 0
u = 0
grd22.Col = 0
grd22.Sort = 1
Call cont2
y = be.RecordCount
tx5 = y
tx5 = tx5 + "  ·„Ì–"
Do While Not be.EOF
u = u + 1
'If u = 10 Then
'be.MoveLast
'Exit Sub
'End If
grd23.row = 0
grd23.Col = 0
grd23.Text = "  „  „⁄«·Ã… "
grd23.Col = 1
grd23.Text = u
grd23.Col = 2
grd23.Text = "  ·„Ì–« „‰ √’· "
grd23.Col = 3
grd23.Text = tx5
sb = 0
sc = 0
tx1 = be!ser
n = grd22.Rows
For i = 1 To n - 1
grd22.row = i
grd22.Col = 0
tx2 = grd22.Text
grd22.Col = 1
tx3 = grd22.Text
grd22.Col = 2
tx4 = grd22.Text
If tx1 = tx2 Then
b = tx4
sb = sb + b
'**** 1
If tx3 = be!mat1 Then
c = be!cof1
sc = sc + c
be!mm1 = tx4
be.Update
ElseIf tx3 = be!mat2 Then
c = be!cof2
sc = sc + c
be!mm2 = tx4
be.Update
ElseIf tx3 = be!mat3 Then
c = be!cof3
sc = sc + c
be!mm3 = tx4
be.Update
ElseIf tx3 = be!mat4 Then
c = be!cof4
sc = sc + c
be!mm4 = tx4
be.Update
ElseIf tx3 = be!mat5 Then
c = be!cof5
sc = sc + c
be!mm5 = tx4
be.Update
ElseIf tx3 = be!mat6 Then
c = be!cof6
sc = sc + c
be!mm6 = tx4
be.Update
ElseIf tx3 = be!mat7 Then
c = be!cof7
sc = sc + c
be!mm7 = tx4
be.Update
ElseIf tx3 = be!mat8 Then
c = be!cof8
sc = sc + c
be!mm8 = tx4
be.Update
ElseIf tx3 = be!mat9 Then
c = be!cof9
sc = sc + c
be!mm9 = tx4
be.Update
ElseIf tx3 = be!mat10 Then
c = be!cof10
sc = sc + c
be!mm10 = tx4
be.Update
ElseIf tx3 = be!mat11 Then
c = be!cof11
sc = sc + c
be!mm11 = tx4
be.Update
ElseIf tx3 = be!mat12 Then
c = be!cof12
sc = sc + c
be!mm12 = tx4
be.Update
ElseIf tx3 = be!mat13 Then
c = be!cof13
sc = sc + c
be!mm13 = tx4
be.Update
ElseIf tx3 = be!mat14 Then
c = be!cof14
sc = sc + c
be!mm14 = tx4
be.Update
ElseIf tx3 = be!mat15 Then
c = be!cof15
sc = sc + c
be!mm15 = tx4
be.Update
ElseIf tx3 = be!mat16 Then
c = be!cof16
sc = sc + c
be!mm16 = tx4
be.Update
ElseIf tx3 = be!mat17 Then
c = be!cof17
sc = sc + c
be!mm17 = tx4
be.Update
ElseIf tx3 = be!mat18 Then
c = be!cof18
sc = sc + c
be!mm18 = tx4
be.Update
ElseIf tx3 = be!mat19 Then
c = be!cof19
sc = sc + c
be!mm19 = tx4
be.Update
ElseIf tx3 = be!mat20 Then
c = be!cof20
sc = sc + c
be!mm20 = tx4
be.Update
End If
End If
Next i
If sc > 0 Then
sb = sb / sc
MyNumber = Round(sb, 2)
be!moy = MyNumber
be.Update
End If
be.MoveNext
Loop
grd22.Visible = True
grd23.Clear
grd23.Rows = 1
grd23.Cols = 1
grd23.ColWidth(0) = 4700
grd23.ColAlignment(0) = 1
grd23.row = 0
grd23.Col = 0
grd23.Text = "  „  «·⁄„·Ì… »‰Ã«Õ ⁄·Ï ﬂ«›…  ·«„Ì– «·„ƒ””… " + tx5 + "«"
End Sub
