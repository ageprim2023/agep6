VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form comptabilite 
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
      Height          =   8535
      Left            =   120
      ScaleHeight     =   8535
      ScaleWidth      =   14415
      TabIndex        =   7
      Top             =   840
      Width           =   14415
      Begin TabDlg.SSTab SSTab1 
         Height          =   8295
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   14631
         _Version        =   393216
         Tabs            =   7
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
         TabCaption(0)   =   "«Ì—«œ«  «·√ﬁ”«„"
         TabPicture(0)   =   "comptabilite.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Picture7"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "«·„œ›Ê⁄«  «·”‰ÊÌ…"
         TabPicture(1)   =   "comptabilite.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Picture6"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "—ﬂ‰ «· Õﬁﬁ"
         TabPicture(2)   =   "comptabilite.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Picture13"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "«·„—ﬂ“ «·„«·Ì(«·„Ì“«‰Ì…)‹"
         TabPicture(3)   =   "comptabilite.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Picture4"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "ﬁ«∆„… «·œŒ·"
         TabPicture(4)   =   "comptabilite.frx":0070
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Picture3"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "„Ì“«‰ «·„—«Ã⁄…"
         TabPicture(5)   =   "comptabilite.frx":008C
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "grd3"
         Tab(5).Control(1)=   "grd2"
         Tab(5).Control(2)=   "Picture2"
         Tab(5).ControlCount=   3
         TabCaption(6)   =   "«·œ› — «·ÌÊ„Ì"
         TabPicture(6)   =   "comptabilite.frx":00A8
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Picture5"
         Tab(6).ControlCount=   1
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   7815
            Left            =   -74880
            ScaleHeight     =   7815
            ScaleWidth      =   13935
            TabIndex        =   218
            Top             =   360
            Width           =   13935
            Begin VB.CommandButton Command17 
               Caption         =   "⁄—÷ "
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
               TabIndex        =   220
               Top             =   120
               Width           =   1815
            End
            Begin MSFlexGridLib.MSFlexGrid grd1 
               Height          =   7215
               Left            =   0
               TabIndex        =   219
               Top             =   600
               Width           =   13935
               _ExtentX        =   24580
               _ExtentY        =   12726
               _Version        =   393216
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
            Begin MSComCtl2.DTPicker DT24 
               Height          =   375
               Left            =   9120
               TabIndex        =   221
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
               Format          =   108920833
               CurrentDate     =   41154
            End
            Begin MSComCtl2.DTPicker DT23 
               Height          =   375
               Left            =   11640
               TabIndex        =   222
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
               Format          =   108920833
               CurrentDate     =   41154
            End
            Begin VB.Label Label31 
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
               Index           =   84
               Left            =   10080
               TabIndex        =   224
               Top             =   120
               Width           =   1455
            End
            Begin VB.Label Label31 
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
               Index           =   83
               Left            =   12360
               TabIndex        =   223
               Top             =   120
               Width           =   1455
            End
         End
         Begin VB.PictureBox Picture13 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   7815
            Left            =   -74880
            ScaleHeight     =   7815
            ScaleWidth      =   13935
            TabIndex        =   186
            Top             =   360
            Width           =   13935
            Begin VB.CommandButton Command16 
               Caption         =   " √ﬂÌœ «·„·«ÕŸ…"
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
               TabIndex        =   217
               Top             =   3720
               Width           =   2535
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
               Left            =   960
               TabIndex        =   216
               Top             =   3240
               Width           =   5175
            End
            Begin VB.CommandButton Command15 
               Caption         =   "«·Ê’· €Ì— ”·Ì„"
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
               TabIndex        =   212
               Top             =   2760
               Width           =   2535
            End
            Begin VB.CommandButton Command14 
               Caption         =   "«·Ê’· ”·Ì„"
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
               TabIndex        =   211
               Top             =   2760
               Width           =   2535
            End
            Begin VB.CommandButton Command13 
               Caption         =   "⁄—÷ "
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
               TabIndex        =   198
               Top             =   240
               Width           =   1815
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
               Left            =   2640
               TabIndex        =   188
               Top             =   240
               Width           =   2415
            End
            Begin VB.CommandButton Command1 
               Caption         =   "⁄—÷ "
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
               TabIndex        =   187
               Top             =   120
               Width           =   1815
            End
            Begin MSFlexGridLib.MSFlexGrid grd10 
               Height          =   6735
               Left            =   10800
               TabIndex        =   189
               Top             =   960
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   11880
               _Version        =   393216
               Rows            =   10
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
            Begin MSFlexGridLib.MSFlexGrid grd11 
               Height          =   6735
               Left            =   7200
               TabIndex        =   190
               Top             =   960
               Width           =   3495
               _ExtentX        =   6165
               _ExtentY        =   11880
               _Version        =   393216
               Rows            =   10
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
            Begin MSComCtl2.DTPicker DT22 
               Height          =   375
               Left            =   9120
               TabIndex        =   191
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
               Format          =   108920833
               CurrentDate     =   41154
            End
            Begin MSComCtl2.DTPicker DT21 
               Height          =   375
               Left            =   11640
               TabIndex        =   192
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
               Format          =   108920833
               CurrentDate     =   41154
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·„·«ÕŸ…"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   82
               Left            =   5760
               TabIndex        =   215
               Top             =   3240
               Width           =   1215
            End
            Begin VB.Shape Shape7 
               BorderColor     =   &H00FFFFFF&
               BorderWidth     =   2
               Height          =   1935
               Left            =   120
               Top             =   720
               Width           =   6975
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
               Index           =   81
               Left            =   3720
               TabIndex        =   214
               Top             =   2280
               Width           =   1215
            End
            Begin VB.Label Label58 
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
               Left            =   2520
               TabIndex        =   213
               Top             =   2280
               Width           =   1440
            End
            Begin VB.Label Label57 
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
               Left            =   120
               TabIndex        =   210
               Top             =   1800
               Width           =   5040
            End
            Begin VB.Label Label56 
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
               Left            =   120
               TabIndex        =   209
               Top             =   1320
               Width           =   1440
            End
            Begin VB.Label Label55 
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
               Left            =   3000
               TabIndex        =   208
               Top             =   1320
               Width           =   1440
            End
            Begin VB.Label Label54 
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
               Left            =   5040
               TabIndex        =   207
               Top             =   1320
               Width           =   1440
            End
            Begin VB.Label Label53 
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
               TabIndex        =   206
               Top             =   840
               Width           =   2640
            End
            Begin VB.Label Label30 
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
               Left            =   3840
               TabIndex        =   205
               Top             =   840
               Width           =   1440
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·‘ÂÊ— «·„œ›Ê⁄ ⁄‰Â«"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   80
               Left            =   4680
               TabIndex        =   204
               Top             =   1800
               Width           =   2295
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
               Index           =   79
               Left            =   1320
               TabIndex        =   203
               Top             =   1320
               Width           =   1455
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«”„ «· ·„Ì–"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   78
               Left            =   2520
               TabIndex        =   202
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label Label31 
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
               Index           =   77
               Left            =   3720
               TabIndex        =   201
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label Label31 
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
               Index           =   45
               Left            =   5760
               TabIndex        =   200
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «· ”·”·Ì ·· ·„Ì–"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   5160
               TabIndex        =   199
               Top             =   840
               Width           =   1815
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00FFFFFF&
               BorderWidth     =   2
               X1              =   7080
               X2              =   7080
               Y1              =   0
               Y2              =   7800
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·Ê’·"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   4920
               TabIndex        =   197
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label31 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·√Ê’«· «· Ì  „ «· Õﬁﬁ „‰Â«"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   7200
               TabIndex        =   196
               Top             =   600
               Width           =   3495
            End
            Begin VB.Label Label31 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·√Ê’«· «· Ì  „ œ›⁄Â«"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   11040
               TabIndex        =   195
               Top             =   600
               Width           =   2775
            End
            Begin VB.Label Label31 
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
               Index           =   59
               Left            =   12360
               TabIndex        =   194
               Top             =   120
               Width           =   1455
            End
            Begin VB.Label Label31 
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
               Index           =   66
               Left            =   10080
               TabIndex        =   193
               Top             =   120
               Width           =   1455
            End
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   3735
            Left            =   -74880
            ScaleHeight     =   3735
            ScaleWidth      =   13935
            TabIndex        =   80
            Top             =   4440
            Width           =   13935
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
               ItemData        =   "comptabilite.frx":00C4
               Left            =   5400
               List            =   "comptabilite.frx":00DA
               Style           =   2  'Dropdown List
               TabIndex        =   81
               Top             =   120
               Width           =   3015
            End
            Begin MSFlexGridLib.MSFlexGrid grd4 
               Height          =   2655
               Left            =   6960
               TabIndex        =   82
               Top             =   600
               Width           =   6975
               _ExtentX        =   12303
               _ExtentY        =   4683
               _Version        =   393216
               Rows            =   10
               Cols            =   1
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
            Begin MSFlexGridLib.MSFlexGrid grd5 
               Height          =   2655
               Left            =   120
               TabIndex        =   83
               Top             =   600
               Width           =   6735
               _ExtentX        =   11880
               _ExtentY        =   4683
               _Version        =   393216
               Rows            =   10
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
               Caption         =   "œ› — Õ”«»"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               TabIndex        =   87
               Top             =   120
               Width           =   2055
            End
            Begin VB.Label Label25 
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
               Left            =   4080
               TabIndex        =   86
               Top             =   3360
               Width           =   5655
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·Ã«‰» «·œ«∆‰"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   480
               TabIndex        =   85
               Top             =   120
               Width           =   3975
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·Ã«‰» «·„œÌ‰"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   9840
               TabIndex        =   84
               Top             =   120
               Width           =   3975
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   7815
            Left            =   -74880
            ScaleHeight     =   7815
            ScaleWidth      =   13935
            TabIndex        =   66
            Top             =   360
            Width           =   13935
            Begin VB.PictureBox Picture15 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   6735
               Left            =   120
               ScaleHeight     =   6735
               ScaleWidth      =   13695
               TabIndex        =   136
               Top             =   960
               Visible         =   0   'False
               Width           =   13695
               Begin VB.CommandButton Command12 
                  Caption         =   "<"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
                     Size            =   36
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   0
                  TabIndex        =   151
                  Top             =   0
                  Width           =   495
               End
               Begin MSFlexGridLib.MSFlexGrid grd6 
                  Height          =   2895
                  Left            =   6840
                  TabIndex        =   141
                  Top             =   360
                  Width           =   6855
                  _ExtentX        =   12091
                  _ExtentY        =   5106
                  _Version        =   393216
                  Rows            =   10
                  Cols            =   1
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
               Begin MSFlexGridLib.MSFlexGrid grd7 
                  Height          =   2775
                  Left            =   0
                  TabIndex        =   142
                  Top             =   3840
                  Width           =   6855
                  _ExtentX        =   12091
                  _ExtentY        =   4895
                  _Version        =   393216
                  Rows            =   10
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
               Begin MSFlexGridLib.MSFlexGrid grd8 
                  Height          =   2775
                  Left            =   6840
                  TabIndex        =   143
                  Top             =   3840
                  Width           =   6855
                  _ExtentX        =   12091
                  _ExtentY        =   4895
                  _Version        =   393216
                  Rows            =   10
                  Cols            =   1
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
               Begin MSFlexGridLib.MSFlexGrid grd9 
                  Height          =   2655
                  Left            =   0
                  TabIndex        =   144
                  Top             =   600
                  Width           =   6855
                  _ExtentX        =   12091
                  _ExtentY        =   4683
                  _Version        =   393216
                  Rows            =   10
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
               Begin VB.Line Line1 
                  BorderColor     =   &H00FFFFFF&
                  X1              =   120
                  X2              =   13800
                  Y1              =   3360
                  Y2              =   3360
               End
               Begin VB.Label Label31 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„’—Ê›«  «·„ƒ””…"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
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
                  Left            =   11520
                  TabIndex        =   140
                  Top             =   3435
                  Width           =   2055
               End
               Begin VB.Label Label31 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   " Ê“Ì⁄ «·√—»«Õ ⁄·Ï «·‘—ﬂ«¡"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
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
                  Left            =   3000
                  TabIndex        =   139
                  Top             =   0
                  Width           =   3855
               End
               Begin VB.Label Label31 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„” Õﬁ«  «·√”« –…"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
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
                  Left            =   4680
                  TabIndex        =   138
                  Top             =   3480
                  Width           =   2055
               End
               Begin VB.Label Label31 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ì—«œ«  «· ·«„Ì–"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
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
                  Left            =   11520
                  TabIndex        =   137
                  Top             =   0
                  Width           =   2055
               End
            End
            Begin VB.PictureBox Picture14 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   6735
               Left            =   120
               ScaleHeight     =   6735
               ScaleWidth      =   13695
               TabIndex        =   131
               Top             =   960
               Width           =   13695
               Begin VB.CommandButton Command11 
                  Caption         =   ">"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
                     Size            =   36
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   13200
                  TabIndex        =   150
                  Top             =   0
                  Width           =   495
               End
               Begin VB.PictureBox Picture16 
                  Height          =   3615
                  Left            =   360
                  ScaleHeight     =   3555
                  ScaleWidth      =   1275
                  TabIndex        =   145
                  Top             =   840
                  Visible         =   0   'False
                  Width           =   1335
                  Begin VB.CommandButton Command10 
                     Caption         =   "Command10"
                     Height          =   375
                     Left            =   0
                     TabIndex        =   146
                     Top             =   960
                     Width           =   2535
                  End
                  Begin VB.Label Label2 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "100"
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
                     Left            =   480
                     TabIndex        =   149
                     Top             =   2160
                     Width           =   495
                  End
                  Begin VB.Label Label48 
                     Caption         =   "Label48"
                     Height          =   255
                     Left            =   0
                     TabIndex        =   148
                     Top             =   120
                     Width           =   855
                  End
                  Begin VB.Label Label49 
                     Caption         =   "Label49"
                     Height          =   255
                     Left            =   3600
                     TabIndex        =   147
                     Top             =   120
                     Width           =   1335
                  End
               End
               Begin MSFlexGridLib.MSFlexGrid grd14 
                  Height          =   6495
                  Left            =   720
                  TabIndex        =   132
                  Top             =   120
                  Width           =   12495
                  _ExtentX        =   22040
                  _ExtentY        =   11456
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
                     Size            =   11.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
            End
            Begin VB.Label Label51 
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
               TabIndex        =   135
               Top             =   240
               Width           =   1515
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·—»Õ"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   1680
               TabIndex        =   134
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label Label50 
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
               Left            =   6960
               TabIndex        =   133
               Top             =   480
               Width           =   1515
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "„Ã„Ê⁄ «·œ«Œ·"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   11880
               TabIndex        =   79
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label Label1 
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
               Left            =   10680
               TabIndex        =   78
               Top             =   240
               Width           =   1515
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "„” Õﬁ«  «·√”« –…"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   8760
               TabIndex        =   77
               Top             =   480
               Width           =   1575
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "„Ã„Ê⁄ «·—”Ê„"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   11880
               TabIndex        =   76
               Top             =   480
               Width           =   1815
            End
            Begin VB.Label Label3 
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
               Left            =   10680
               TabIndex        =   75
               Top             =   480
               Width           =   1515
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·«Ì—«œ« "
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   8880
               TabIndex        =   74
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label4 
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
               Left            =   6960
               TabIndex        =   73
               Top             =   240
               Width           =   1515
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "„” Õﬁ«  √”« –… «·‰”»…"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   4680
               TabIndex        =   72
               Top             =   240
               Width           =   2055
            End
            Begin VB.Label Label5 
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
               Left            =   3120
               TabIndex        =   71
               Top             =   240
               Width           =   1515
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "‰’Ì» «·„ƒ””…"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   4920
               TabIndex        =   70
               Top             =   480
               Width           =   1815
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
               Height          =   375
               Left            =   3120
               TabIndex        =   69
               Top             =   480
               Width           =   1515
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·„’—Ê›« "
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   1440
               TabIndex        =   68
               Top             =   240
               Width           =   1455
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
               Height          =   375
               Left            =   240
               TabIndex        =   67
               Top             =   480
               Width           =   1515
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H00FFFFFF&
               Height          =   735
               Left            =   120
               Top             =   120
               Width           =   13695
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   7815
            Left            =   -74880
            ScaleHeight     =   7815
            ScaleWidth      =   13935
            TabIndex        =   30
            Top             =   360
            Width           =   13935
            Begin VB.PictureBox Picture12 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   1095
               Left            =   0
               ScaleHeight     =   1095
               ScaleWidth      =   13935
               TabIndex        =   111
               Top             =   6720
               Visible         =   0   'False
               Width           =   13935
               Begin VB.CommandButton Command9 
                  Caption         =   " √ﬂÌœ › Õ ”‰… œ—«”Ì… ÃœÌœ…"
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
                  TabIndex        =   118
                  Top             =   600
                  Width           =   4935
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
                  ItemData        =   "comptabilite.frx":0124
                  Left            =   7320
                  List            =   "comptabilite.frx":013A
                  Style           =   2  'Dropdown List
                  TabIndex        =   114
                  Top             =   120
                  Width           =   1455
               End
               Begin MSComCtl2.DTPicker DT4 
                  Height          =   375
                  Left            =   3840
                  TabIndex        =   124
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
                  Format          =   108920833
                  CurrentDate     =   41183
               End
               Begin MSComCtl2.DTPicker DT5 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   128
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
                  Format          =   108920833
                  CurrentDate     =   41183
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
                  Left            =   1200
                  TabIndex        =   130
                  Top             =   120
                  Width           =   2535
               End
               Begin VB.Label Label47 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·⁄„·Ì… ﬁœ  √Œ– Êﬁ « ·–«·ﬂ Ì—ÃÏ «·«‰ Ÿ«—..........‹ "
                  BeginProperty Font 
                     Name            =   "Times New Roman"
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
                  TabIndex        =   129
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   8640
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
                  Index           =   36
                  Left            =   5160
                  TabIndex        =   125
                  Top             =   120
                  Width           =   2055
               End
               Begin VB.Label Label31 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·”‰… «·œ—«”Ì… «·ÃœÌœ…"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   375
                  Index           =   73
                  Left            =   8280
                  TabIndex        =   115
                  Top             =   120
                  Width           =   2295
               End
               Begin VB.Label Label26 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "2012-2013"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   375
                  Left            =   10560
                  TabIndex        =   113
                  Top             =   120
                  Width           =   1320
               End
               Begin VB.Label Label31 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·”‰… «·œ—«”Ì… «·„‰’—„…"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   375
                  Index           =   72
                  Left            =   11520
                  TabIndex        =   112
                  Top             =   120
                  Width           =   2295
               End
            End
            Begin VB.CommandButton Command8 
               Caption         =   "› Õ ”‰… œ—«”Ì… ÃœÌœ…"
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
               Left            =   4920
               TabIndex        =   110
               Top             =   6240
               Width           =   4095
            End
            Begin VB.Label Label59 
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
               Left            =   8880
               TabIndex        =   226
               Top             =   2640
               Width           =   1560
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·⁄„«·"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   11640
               TabIndex        =   225
               Top             =   2640
               Width           =   2055
            End
            Begin VB.Label Label46 
               Caption         =   "Label46"
               Height          =   255
               Left            =   12120
               TabIndex        =   127
               Top             =   6240
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.Label Label45 
               Caption         =   "Label45"
               Height          =   255
               Left            =   9240
               TabIndex        =   126
               Top             =   6240
               Visible         =   0   'False
               Width           =   1935
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
               Left            =   2040
               TabIndex        =   123
               Top             =   1560
               Width           =   1560
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
               Left            =   2040
               TabIndex        =   122
               Top             =   2280
               Width           =   1560
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "********"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   75
               Left            =   3840
               TabIndex        =   121
               Top             =   2280
               Width           =   3015
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
               Left            =   8880
               TabIndex        =   120
               Top             =   2280
               Width           =   1560
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·√”« –…"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   74
               Left            =   10680
               TabIndex        =   119
               Top             =   2280
               Width           =   3015
            End
            Begin VB.Label Label41 
               Caption         =   "Label41"
               Height          =   255
               Left            =   2040
               TabIndex        =   117
               Top             =   6240
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label Label40 
               Caption         =   "Label40"
               Height          =   255
               Left            =   120
               TabIndex        =   116
               Top             =   6240
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00FFFFFF&
               BorderWidth     =   5
               Index           =   0
               X1              =   6960
               X2              =   6960
               Y1              =   120
               Y2              =   6000
            End
            Begin VB.Shape Shape2 
               BorderColor     =   &H00FFFFFF&
               Height          =   5895
               Left            =   120
               Top             =   120
               Width           =   13695
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00FFFFFF&
               Index           =   1
               X1              =   10560
               X2              =   10560
               Y1              =   600
               Y2              =   5400
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00FFFFFF&
               Index           =   2
               X1              =   8760
               X2              =   8760
               Y1              =   600
               Y2              =   5400
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
               Height          =   375
               Left            =   8880
               TabIndex        =   65
               Top             =   1200
               Width           =   1560
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00FFFFFF&
               Index           =   0
               X1              =   6960
               X2              =   13800
               Y1              =   600
               Y2              =   600
            End
            Begin VB.Label Label31 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·√’Ê·"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   7080
               TabIndex        =   64
               Top             =   240
               Width           =   6615
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "√’Ê· „ œ«Ê·…"
               BeginProperty Font 
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
               Index           =   14
               Left            =   11640
               TabIndex        =   63
               Top             =   720
               Width           =   2055
            End
            Begin VB.Line Line5 
               BorderColor     =   &H00FFFFFF&
               Index           =   0
               X1              =   12240
               X2              =   13680
               Y1              =   1080
               Y2              =   1080
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·»‰ﬂ"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   11640
               TabIndex        =   62
               Top             =   1200
               Width           =   2055
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·’‰œÊﬁ"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   11640
               TabIndex        =   61
               Top             =   1560
               Width           =   2055
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·‘—ﬂ«¡"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   11640
               TabIndex        =   60
               Top             =   1920
               Width           =   2055
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "≈Ã„«·Ì «·√’Ê· «·„ œ«Ê·…"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   11400
               TabIndex        =   59
               Top             =   3240
               Width           =   2295
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
               Height          =   375
               Left            =   8880
               TabIndex        =   58
               Top             =   1560
               Width           =   1560
            End
            Begin VB.Label Label10 
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
               Left            =   8880
               TabIndex        =   57
               Top             =   1920
               Width           =   1560
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
               Left            =   7080
               TabIndex        =   56
               Top             =   3240
               Width           =   1560
            End
            Begin VB.Line Line6 
               BorderColor     =   &H00FFFFFF&
               Index           =   0
               X1              =   6960
               X2              =   10560
               Y1              =   3720
               Y2              =   3720
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "√’Ê· À«» …"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   11640
               TabIndex        =   55
               Top             =   3840
               Width           =   2055
            End
            Begin VB.Line Line5 
               BorderColor     =   &H00FFFFFF&
               Index           =   1
               X1              =   12240
               X2              =   13680
               Y1              =   4320
               Y2              =   4320
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "********"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   11640
               TabIndex        =   54
               Top             =   4440
               Width           =   2055
            End
            Begin VB.Label Label31 
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
               Index           =   23
               Left            =   8880
               TabIndex        =   53
               Top             =   4440
               Width           =   1575
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "≈Ã„«·Ì «·√’Ê· «·À«» …"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   11400
               TabIndex        =   52
               Top             =   4920
               Width           =   2295
            End
            Begin VB.Label Label31 
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
               Index           =   25
               Left            =   7080
               TabIndex        =   51
               Top             =   4920
               Width           =   1575
            End
            Begin VB.Line Line6 
               BorderColor     =   &H00FFFFFF&
               Index           =   1
               X1              =   120
               X2              =   13800
               Y1              =   5400
               Y2              =   5400
            End
            Begin VB.Label Label31 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·Œ’Ê„ + «·ÕﬁÊﬁ «·„·ﬂÌ…"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   240
               TabIndex        =   50
               Top             =   240
               Width           =   6615
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00FFFFFF&
               Index           =   1
               X1              =   120
               X2              =   6960
               Y1              =   600
               Y2              =   600
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00FFFFFF&
               Index           =   3
               X1              =   1920
               X2              =   1920
               Y1              =   600
               Y2              =   5400
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00FFFFFF&
               Index           =   4
               X1              =   3720
               X2              =   3720
               Y1              =   600
               Y2              =   5400
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·Œ’Ê„"
               BeginProperty Font 
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
               Index           =   27
               Left            =   4800
               TabIndex        =   49
               Top             =   720
               Width           =   2055
            End
            Begin VB.Line Line5 
               BorderColor     =   &H00FFFFFF&
               Index           =   2
               X1              =   5400
               X2              =   6840
               Y1              =   1080
               Y2              =   1080
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·√”« –…"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   4800
               TabIndex        =   48
               Top             =   1200
               Width           =   2055
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "********"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   4800
               TabIndex        =   47
               Top             =   1920
               Width           =   2055
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·⁄„«·"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   4800
               TabIndex        =   46
               Top             =   1560
               Width           =   2055
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "≈Ã„«·Ì «·Œ’Ê„"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   4560
               TabIndex        =   45
               Top             =   2760
               Width           =   2295
            End
            Begin VB.Line Line6 
               BorderColor     =   &H00FFFFFF&
               Index           =   2
               X1              =   120
               X2              =   3720
               Y1              =   3240
               Y2              =   3240
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·ÕﬁÊﬁ «·„·ﬂÌ…"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   4800
               TabIndex        =   44
               Top             =   3360
               Width           =   2055
            End
            Begin VB.Line Line5 
               BorderColor     =   &H00FFFFFF&
               Index           =   3
               X1              =   5400
               X2              =   6840
               Y1              =   3720
               Y2              =   3720
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "—√” «·„«·"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   4800
               TabIndex        =   43
               Top             =   4080
               Width           =   2055
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·√—»«Õ"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   4800
               TabIndex        =   42
               Top             =   4440
               Width           =   2055
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "≈Ã„«·Ì «·Œ’Ê„ Ê «·ÕﬁÊﬁ «·„·ﬂÌ…"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   3840
               TabIndex        =   41
               Top             =   5520
               Width           =   3015
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
               Left            =   2040
               TabIndex        =   40
               Top             =   1200
               Width           =   1560
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
               Left            =   2040
               TabIndex        =   39
               Top             =   1920
               Width           =   1560
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
               Left            =   240
               TabIndex        =   38
               Top             =   2760
               Width           =   1560
            End
            Begin VB.Label Label15 
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
               TabIndex        =   37
               Top             =   4080
               Width           =   1560
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
               Left            =   2040
               TabIndex        =   36
               Top             =   4440
               Width           =   1560
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
               Left            =   240
               TabIndex        =   35
               Top             =   4920
               Width           =   1560
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "≈Ã„«·Ì «·ÕﬁÊﬁ «·„·ﬂÌ…"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   37
               Left            =   4800
               TabIndex        =   34
               Top             =   4920
               Width           =   2055
            End
            Begin VB.Label Label18 
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
               Left            =   240
               TabIndex        =   33
               Top             =   5520
               Width           =   3345
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "≈Ã„«·Ì «·√’Ê·"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   10680
               TabIndex        =   32
               Top             =   5520
               Width           =   3015
            End
            Begin VB.Label Label19 
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
               Left            =   7080
               TabIndex        =   31
               Top             =   5520
               Width           =   3360
            End
         End
         Begin VB.PictureBox Picture7 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   7815
            Left            =   120
            ScaleHeight     =   7815
            ScaleWidth      =   13935
            TabIndex        =   21
            Top             =   360
            Width           =   13935
            Begin TabDlg.SSTab SSTab2 
               Height          =   7575
               Left            =   120
               TabIndex        =   22
               Top             =   120
               Width           =   13695
               _ExtentX        =   24156
               _ExtentY        =   13361
               _Version        =   393216
               Tabs            =   4
               TabsPerRow      =   4
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
               TabCaption(0)   =   "Ã„Ì⁄ «·√ﬁ”«„ ›Ì Ã„Ì⁄ «·√‘Â—"
               TabPicture(0)   =   "comptabilite.frx":0184
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "Picture11"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).ControlCount=   1
               TabCaption(1)   =   "Ã„Ì⁄ «·√ﬁ”«„ ›Ì «·‘Â— «·Ê«Õœ"
               TabPicture(1)   =   "comptabilite.frx":01A0
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "Picture10"
               Tab(1).ControlCount=   1
               TabCaption(2)   =   "«·ﬁ”„ «·Ê«Õœ ›Ì Ã„Ì⁄ «·√‘Â—"
               TabPicture(2)   =   "comptabilite.frx":01BC
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "Picture9"
               Tab(2).ControlCount=   1
               TabCaption(3)   =   "«·ﬁ”„ «·Ê«Õœ ›Ì «·‘Â— «·Ê«Õœ"
               TabPicture(3)   =   "comptabilite.frx":01D8
               Tab(3).ControlEnabled=   0   'False
               Tab(3).Control(0)=   "Picture8"
               Tab(3).ControlCount=   1
               Begin VB.PictureBox Picture11 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   7095
                  Left            =   120
                  ScaleHeight     =   7095
                  ScaleWidth      =   13455
                  TabIndex        =   99
                  Top             =   360
                  Width           =   13455
                  Begin MSChart20Lib.MSChart MSChart4 
                     Height          =   5775
                     Left            =   120
                     OleObjectBlob   =   "comptabilite.frx":01F4
                     TabIndex        =   101
                     Top             =   1200
                     Width           =   4815
                  End
                  Begin VB.CommandButton Command7 
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
                     Left            =   3360
                     TabIndex        =   100
                     Top             =   120
                     Width           =   1335
                  End
                  Begin MSFlexGridLib.MSFlexGrid grd15 
                     Height          =   5775
                     Left            =   5040
                     TabIndex        =   152
                     Top             =   1200
                     Width           =   8295
                     _ExtentX        =   14631
                     _ExtentY        =   10186
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
                        Size            =   11.25
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
                     Caption         =   "«·œŒ·"
                     BeginProperty Font 
                        Name            =   "Times New Roman"
                        Size            =   12
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   375
                     Index           =   71
                     Left            =   12120
                     TabIndex        =   109
                     Top             =   720
                     Width           =   1095
                  End
                  Begin VB.Label Label39 
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
                     Left            =   11040
                     TabIndex        =   108
                     Top             =   720
                     Width           =   1560
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·„” Õﬁ"
                     BeginProperty Font 
                        Name            =   "Times New Roman"
                        Size            =   12
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   375
                     Index           =   70
                     Left            =   9720
                     TabIndex        =   107
                     Top             =   720
                     Width           =   1095
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
                     Left            =   8280
                     TabIndex        =   106
                     Top             =   720
                     Width           =   1560
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·„’—Ê›"
                     BeginProperty Font 
                        Name            =   "Times New Roman"
                        Size            =   12
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   375
                     Index           =   69
                     Left            =   6360
                     TabIndex        =   105
                     Top             =   720
                     Width           =   1695
                  End
                  Begin VB.Label Label28 
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
                     Left            =   5400
                     TabIndex        =   104
                     Top             =   720
                     Width           =   1560
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "’«›Ì «·«Ì—«œ"
                     BeginProperty Font 
                        Name            =   "Times New Roman"
                        Size            =   12
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   375
                     Index           =   68
                     Left            =   3480
                     TabIndex        =   103
                     Top             =   720
                     Width           =   1695
                  End
                  Begin VB.Label Label27 
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
                     Left            =   480
                     TabIndex        =   102
                     Top             =   720
                     Width           =   3480
                  End
                  Begin VB.Shape Shape6 
                     BorderColor     =   &H00FFFFFF&
                     Height          =   495
                     Left            =   120
                     Top             =   600
                     Width           =   13215
                  End
               End
               Begin VB.PictureBox Picture10 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   7095
                  Left            =   -74880
                  ScaleHeight     =   7095
                  ScaleWidth      =   13455
                  TabIndex        =   95
                  Top             =   360
                  Width           =   13455
                  Begin VB.PictureBox Picture17 
                     Height          =   3135
                     Left            =   960
                     ScaleHeight     =   3075
                     ScaleWidth      =   3435
                     TabIndex        =   163
                     Top             =   2520
                     Visible         =   0   'False
                     Width           =   3495
                     Begin VB.TextBox Text4 
                        Height          =   285
                        Left            =   120
                        TabIndex        =   167
                        Text            =   "Text4"
                        Top             =   1800
                        Width           =   1935
                     End
                     Begin VB.TextBox Text3 
                        Height          =   375
                        Left            =   120
                        TabIndex        =   166
                        Text            =   "Text3"
                        Top             =   1200
                        Width           =   1935
                     End
                     Begin MSComCtl2.DTPicker DT11 
                        Height          =   375
                        Left            =   120
                        TabIndex        =   164
                        Top             =   240
                        Width           =   1935
                        _ExtentX        =   3413
                        _ExtentY        =   661
                        _Version        =   393216
                        Format          =   108920833
                        CurrentDate     =   41268
                     End
                     Begin MSComCtl2.DTPicker DT12 
                        Height          =   375
                        Left            =   120
                        TabIndex        =   165
                        Top             =   720
                        Width           =   1935
                        _ExtentX        =   3413
                        _ExtentY        =   661
                        _Version        =   393216
                        Format          =   108920833
                        CurrentDate     =   41268
                     End
                     Begin VB.Label Label61 
                        Caption         =   "Label61"
                        Height          =   855
                        Left            =   2280
                        TabIndex        =   228
                        Top             =   960
                        Width           =   855
                     End
                     Begin VB.Label Label52 
                        Caption         =   "Label52"
                        Height          =   255
                        Left            =   120
                        TabIndex        =   169
                        Top             =   2520
                        Width           =   1335
                     End
                     Begin VB.Label Label38 
                        Caption         =   "Label38"
                        Height          =   375
                        Left            =   120
                        TabIndex        =   168
                        Top             =   2280
                        Width           =   1935
                     End
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
                     ItemData        =   "comptabilite.frx":27FC
                     Left            =   4800
                     List            =   "comptabilite.frx":2812
                     Style           =   2  'Dropdown List
                     TabIndex        =   97
                     Top             =   120
                     Width           =   1215
                  End
                  Begin VB.CommandButton Command6 
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
                     Left            =   3360
                     TabIndex        =   96
                     Top             =   120
                     Width           =   1335
                  End
                  Begin MSChart20Lib.MSChart MSChart3 
                     Height          =   5775
                     Left            =   120
                     OleObjectBlob   =   "comptabilite.frx":285C
                     TabIndex        =   153
                     Top             =   1200
                     Width           =   4815
                  End
                  Begin MSFlexGridLib.MSFlexGrid grd16 
                     Height          =   5775
                     Left            =   5040
                     TabIndex        =   154
                     Top             =   1200
                     Width           =   8295
                     _ExtentX        =   14631
                     _ExtentY        =   10186
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
                        Size            =   11.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin VB.Label Label60 
                     Caption         =   "Label60"
                     Height          =   375
                     Left            =   3960
                     TabIndex        =   227
                     Top             =   1800
                     Width           =   855
                  End
                  Begin VB.Label Label37 
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
                     Left            =   480
                     TabIndex        =   162
                     Top             =   720
                     Width           =   3480
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "’«›Ì «·«Ì—«œ"
                     BeginProperty Font 
                        Name            =   "Times New Roman"
                        Size            =   12
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   375
                     Index           =   65
                     Left            =   3480
                     TabIndex        =   161
                     Top             =   720
                     Width           =   1695
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
                     Left            =   5400
                     TabIndex        =   160
                     Top             =   720
                     Width           =   1560
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·„’—Ê›"
                     BeginProperty Font 
                        Name            =   "Times New Roman"
                        Size            =   12
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   375
                     Index           =   64
                     Left            =   6360
                     TabIndex        =   159
                     Top             =   720
                     Width           =   1695
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
                     Left            =   8280
                     TabIndex        =   158
                     Top             =   720
                     Width           =   1560
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·„” Õﬁ"
                     BeginProperty Font 
                        Name            =   "Times New Roman"
                        Size            =   12
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   375
                     Index           =   62
                     Left            =   9720
                     TabIndex        =   157
                     Top             =   720
                     Width           =   1095
                  End
                  Begin VB.Label Label21 
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
                     Left            =   11040
                     TabIndex        =   156
                     Top             =   720
                     Width           =   1560
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·œŒ·"
                     BeginProperty Font 
                        Name            =   "Times New Roman"
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
                     Left            =   12120
                     TabIndex        =   155
                     Top             =   720
                     Width           =   1095
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
                     Index           =   67
                     Left            =   5640
                     TabIndex        =   98
                     Top             =   120
                     Width           =   975
                  End
                  Begin VB.Shape Shape5 
                     BorderColor     =   &H00FFFFFF&
                     Height          =   495
                     Left            =   120
                     Top             =   600
                     Width           =   13215
                  End
               End
               Begin VB.PictureBox Picture9 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   7095
                  Left            =   -74880
                  ScaleHeight     =   7095
                  ScaleWidth      =   13455
                  TabIndex        =   91
                  Top             =   360
                  Width           =   13455
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
                     ItemData        =   "comptabilite.frx":4E64
                     Left            =   6840
                     List            =   "comptabilite.frx":4E7A
                     Style           =   2  'Dropdown List
                     TabIndex        =   93
                     Top             =   120
                     Width           =   2415
                  End
                  Begin VB.CommandButton Command4 
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
                     Left            =   3360
                     TabIndex        =   92
                     Top             =   120
                     Width           =   1575
                  End
                  Begin MSChart20Lib.MSChart MSChart2 
                     Height          =   5775
                     Left            =   120
                     OleObjectBlob   =   "comptabilite.frx":4EC4
                     TabIndex        =   170
                     Top             =   1200
                     Width           =   4815
                  End
                  Begin MSFlexGridLib.MSFlexGrid grd17 
                     Height          =   5775
                     Left            =   5040
                     TabIndex        =   171
                     Top             =   1200
                     Width           =   8295
                     _ExtentX        =   14631
                     _ExtentY        =   10186
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
                        Size            =   11.25
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
                     Caption         =   "«·œŒ·"
                     BeginProperty Font 
                        Name            =   "Times New Roman"
                        Size            =   12
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   375
                     Index           =   61
                     Left            =   12120
                     TabIndex        =   177
                     Top             =   720
                     Width           =   1095
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
                     Left            =   11040
                     TabIndex        =   176
                     Top             =   720
                     Width           =   1560
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·„” Õﬁ"
                     BeginProperty Font 
                        Name            =   "Times New Roman"
                        Size            =   12
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   375
                     Index           =   60
                     Left            =   9720
                     TabIndex        =   175
                     Top             =   720
                     Width           =   1095
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
                     Left            =   8280
                     TabIndex        =   174
                     Top             =   720
                     Width           =   1560
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·«Ì—«œ«  œÊ‰ «·„’—Ê›« "
                     BeginProperty Font 
                        Name            =   "Times New Roman"
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
                     Left            =   3480
                     TabIndex        =   173
                     Top             =   720
                     Width           =   2655
                  End
                  Begin VB.Label Label29 
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
                     Left            =   480
                     TabIndex        =   172
                     Top             =   720
                     Width           =   3480
                  End
                  Begin VB.Label Label31 
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
                     Index           =   63
                     Left            =   9000
                     TabIndex        =   94
                     Top             =   120
                     Width           =   975
                  End
                  Begin VB.Shape Shape4 
                     BorderColor     =   &H00FFFFFF&
                     Height          =   495
                     Left            =   120
                     Top             =   600
                     Width           =   13215
                  End
               End
               Begin VB.PictureBox Picture8 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   7095
                  Left            =   -74880
                  ScaleHeight     =   7095
                  ScaleWidth      =   13455
                  TabIndex        =   23
                  Top             =   360
                  Width           =   13455
                  Begin MSComCtl2.DTPicker DT3 
                     Height          =   375
                     Left            =   120
                     TabIndex        =   90
                     Top             =   120
                     Visible         =   0   'False
                     Width           =   1335
                     _ExtentX        =   2355
                     _ExtentY        =   661
                     _Version        =   393216
                     Format          =   108920833
                     CurrentDate     =   41176
                  End
                  Begin VB.CommandButton Command3 
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
                     Left            =   3360
                     TabIndex        =   28
                     Top             =   120
                     Width           =   1575
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
                     ItemData        =   "comptabilite.frx":71A3
                     Left            =   5040
                     List            =   "comptabilite.frx":71B9
                     Style           =   2  'Dropdown List
                     TabIndex        =   26
                     Top             =   120
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
                     ItemData        =   "comptabilite.frx":7203
                     Left            =   6840
                     List            =   "comptabilite.frx":7219
                     Style           =   2  'Dropdown List
                     TabIndex        =   24
                     Top             =   120
                     Width           =   2415
                  End
                  Begin MSFlexGridLib.MSFlexGrid grd18 
                     Height          =   5775
                     Left            =   5040
                     TabIndex        =   178
                     Top             =   1200
                     Width           =   8295
                     _ExtentX        =   14631
                     _ExtentY        =   10186
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
                        Size            =   11.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin MSChart20Lib.MSChart MSChart1 
                     Height          =   5775
                     Left            =   120
                     OleObjectBlob   =   "comptabilite.frx":7263
                     TabIndex        =   185
                     Top             =   1200
                     Width           =   4815
                  End
                  Begin VB.Label Label23 
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
                     Left            =   480
                     TabIndex        =   184
                     Top             =   720
                     Width           =   3480
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·«Ì—«œ«  œÊ‰ «·„’—Ê›« "
                     BeginProperty Font 
                        Name            =   "Times New Roman"
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
                     Left            =   3480
                     TabIndex        =   183
                     Top             =   720
                     Width           =   2655
                  End
                  Begin VB.Label Label22 
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
                     Left            =   8280
                     TabIndex        =   182
                     Top             =   720
                     Width           =   1560
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·„” Õﬁ"
                     BeginProperty Font 
                        Name            =   "Times New Roman"
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
                     Left            =   9720
                     TabIndex        =   181
                     Top             =   720
                     Width           =   1095
                  End
                  Begin VB.Label Label20 
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
                     Left            =   11040
                     TabIndex        =   180
                     Top             =   720
                     Width           =   1560
                  End
                  Begin VB.Label Label31 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·œŒ·"
                     BeginProperty Font 
                        Name            =   "Times New Roman"
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
                     Left            =   12120
                     TabIndex        =   179
                     Top             =   720
                     Width           =   1095
                  End
                  Begin VB.Label Label24 
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
                     Left            =   1800
                     TabIndex        =   29
                     Top             =   120
                     Visible         =   0   'False
                     Width           =   840
                  End
                  Begin VB.Shape Shape3 
                     BorderColor     =   &H00FFFFFF&
                     Height          =   495
                     Left            =   120
                     Top             =   600
                     Width           =   13215
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
                     Index           =   50
                     Left            =   5640
                     TabIndex        =   27
                     Top             =   120
                     Width           =   975
                  End
                  Begin VB.Label Label31 
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
                     Index           =   46
                     Left            =   9000
                     TabIndex        =   25
                     Top             =   120
                     Width           =   975
                  End
               End
            End
         End
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   7815
            Left            =   -74880
            ScaleHeight     =   7815
            ScaleWidth      =   13935
            TabIndex        =   9
            Top             =   360
            Width           =   13935
            Begin VB.CommandButton Command2 
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
               TabIndex        =   20
               Top             =   120
               Width           =   1575
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
               Left            =   13560
               MaskColor       =   &H00000000&
               TabIndex        =   18
               Top             =   120
               Value           =   1  'Checked
               Width           =   255
            End
            Begin VB.CommandButton Command5 
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
               Left            =   1800
               TabIndex        =   17
               Top             =   120
               Width           =   1695
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
               Left            =   6720
               TabIndex        =   16
               Top             =   120
               Width           =   1695
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
               ItemData        =   "comptabilite.frx":9542
               Left            =   3600
               List            =   "comptabilite.frx":9552
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   120
               Width           =   855
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
               ItemData        =   "comptabilite.frx":9585
               Left            =   9960
               List            =   "comptabilite.frx":959B
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   120
               Width           =   1695
            End
            Begin MSFlexGridLib.MSFlexGrid grd13 
               Height          =   7095
               Left            =   120
               TabIndex        =   11
               Top             =   600
               Width           =   13695
               _ExtentX        =   24156
               _ExtentY        =   12515
               _Version        =   393216
               Rows            =   10
               Cols            =   21
               FixedCols       =   2
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
            Begin VB.Label Label68 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "≈ŸÂ«— «·√‘Â—"
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
               Left            =   12360
               TabIndex        =   19
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·—”Ê„ «· ”ÃÌ·Ì…"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   7560
               TabIndex        =   14
               Top             =   120
               Width           =   2295
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "⁄œœ √‘Â— «·”‰… «·œ—«”Ì…"
               BeginProperty Font 
                  Name            =   "Times New Roman"
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
               Left            =   4320
               TabIndex        =   13
               Top             =   120
               Width           =   2295
            End
            Begin VB.Label Label31 
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
               Index           =   47
               Left            =   11280
               TabIndex        =   12
               Top             =   120
               Width           =   975
            End
         End
         Begin MSFlexGridLib.MSFlexGrid grd2 
            Height          =   3855
            Left            =   -67800
            TabIndex        =   88
            Top             =   480
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   6800
            _Version        =   393216
            Rows            =   10
            Cols            =   1
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
            Height          =   3855
            Left            =   -74880
            TabIndex        =   89
            Top             =   480
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   6800
            _Version        =   393216
            Rows            =   10
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
   End
   Begin VB.PictureBox Picture19 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   14415
      TabIndex        =   0
      Top             =   120
      Width           =   14415
      Begin VB.CommandButton Command31 
         Caption         =   "⁄—÷ "
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
         TabIndex        =   1
         Top             =   120
         Width           =   7335
      End
      Begin MSComCtl2.DTPicker DT2 
         Height          =   375
         Left            =   7560
         TabIndex        =   2
         Top             =   120
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   108920833
         CurrentDate     =   41154
      End
      Begin MSComCtl2.DTPicker DT1 
         Height          =   375
         Left            =   10920
         TabIndex        =   6
         Top             =   120
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   108920833
         CurrentDate     =   41154
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
         TabIndex        =   5
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label31 
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
         Index           =   54
         Left            =   8760
         TabIndex        =   4
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label31 
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
         Index           =   56
         Left            =   12720
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   720
      OleObjectBlob   =   "comptabilite.frx":95E5
      Top             =   1320
   End
End
Attribute VB_Name = "comptabilite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nbrcla As Double
Public co2 As ADODB.Connection
Public co3 As ADODB.Connection
Public nn As ADODB.Recordset
Public ns As ADODB.Recordset
Public od As ADODB.Recordset
Public ri As ADODB.Recordset
Public ti As ADODB.Recordset
Public cv As ADODB.Recordset
Public ai As ADODB.Recordset
Public jn As ADODB.Recordset
Public bk As ADODB.Recordset
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
Function cont3()
Set co3 = New ADODB.Connection
Set od = New ADODB.Recordset
Set ri = New ADODB.Recordset
Set ti = New ADODB.Recordset
Set cv = New ADODB.Recordset
Set ai = New ADODB.Recordset
Set jn = New ADODB.Recordset
Set bk = New ADODB.Recordset
co3.Provider = "microsoft.jet.oledb.4.0; jet oledb:database password=7346804"
co3.ConnectionString = App.Path & "\" & Combo8.Text & ".mdb"
co3.Open
od.Open "select*from Tcodes", co3, adOpenKeyset, adLockOptimistic
ri.Open "select*from Tseries", co3, adOpenKeyset, adLockOptimistic
ti.Open "select*from Tutilisateurs", co3, adOpenKeyset, adLockOptimistic
cv.Open "select*from Tcoffdevoirs", co3, adOpenKeyset, adLockOptimistic
ai.Open "select*from Tcaisse", co3, adOpenKeyset, adLockOptimistic
jn.Open "select*from Tjournal", co3, adOpenKeyset, adLockOptimistic
bk.Open "select*from Tbank", co3, adOpenKeyset, adLockOptimistic
End Function
Public Sub chargec2()
On Error Resume Next
nbrcla = 1
Call cont
Combo2.Clear
Combo3.Clear
Combo7.Clear
  If cl.RecordCount > 0 Then
 nbrcla = 0
 End If
  Do While Not cl.EOF
  If cl!act = "1" Then
  nbrcla = nbrcla + 1
    Combo2.AddItem cl!cla
    Combo3.AddItem cl!cla
    Combo7.AddItem cl!cla
    End If
      cl.MoveNext
 Loop
End Sub
Public Sub chargec3()
On Error Resume Next
Combo4.Clear
Combo5.Clear
Combo6.Clear
Combo4.AddItem "1"
Combo4.AddItem "2"
Combo4.AddItem "3"
Combo4.AddItem "4"
Combo4.AddItem "5"
Combo4.AddItem "6"
Combo4.AddItem "7"
Combo4.AddItem "8"
Combo4.AddItem "9"
Combo4.AddItem "10"
Combo4.AddItem "11"
Combo4.AddItem "12"
Combo5.AddItem "10"
Combo5.AddItem "11"
Combo5.AddItem "12"
Combo5.AddItem "1"
Combo5.AddItem "2"
Combo5.AddItem "3"
Combo5.AddItem "4"
Combo5.AddItem "5"
Combo5.AddItem "6"
Combo5.AddItem "7"
Combo5.AddItem "8"
Combo5.AddItem "9"
Combo6.AddItem "10"
Combo6.AddItem "11"
Combo6.AddItem "12"
Combo6.AddItem "1"
Combo6.AddItem "2"
Combo6.AddItem "3"
Combo6.AddItem "4"
Combo6.AddItem "5"
Combo6.AddItem "6"
Combo6.AddItem "7"
Combo6.AddItem "8"
Combo6.AddItem "9"
End Sub

Private Sub Check2_Click()
On Error Resume Next
If Check2.Value = 1 Then
grd13.ColWidth(2) = 1000
grd13.ColWidth(3) = 1000
grd13.ColWidth(4) = 1000
grd13.ColWidth(5) = 1000
grd13.ColWidth(6) = 1000
grd13.ColWidth(7) = 1000
grd13.ColWidth(8) = 1000
grd13.ColWidth(9) = 1000
grd13.ColWidth(10) = 1000
grd13.ColWidth(11) = 1000
grd13.ColWidth(12) = 1000
grd13.ColWidth(13) = 1000
grd13.ColWidth(14) = 1000
Else
grd13.ColWidth(2) = 0
grd13.ColWidth(3) = 0
grd13.ColWidth(4) = 0
grd13.ColWidth(5) = 0
grd13.ColWidth(6) = 0
grd13.ColWidth(7) = 0
grd13.ColWidth(8) = 0
grd13.ColWidth(9) = 0
grd13.ColWidth(10) = 0
grd13.ColWidth(11) = 0
grd13.ColWidth(12) = 0
grd13.ColWidth(13) = 0
grd13.ColWidth(14) = 0
End If
End Sub

Private Sub Combo1_Change()
On Error Resume Next
Call chargegrd4_5

End Sub

Private Sub Combo1_Click()
On Error Resume Next
Combo1_Change
End Sub

Private Sub Combo2_Change()
On Error Resume Next
grd13.Clear
grd13.Rows = 1
Text5.Text = ""
Call cont
Do While Not ce.EOF
If Combo2.Text = ce!cla Then
Text5.Text = ce!man
End If
ce.MoveNext
Loop
End Sub

Private Sub Combo2_Click()
On Error Resume Next
Combo2_Change
End Sub

Private Sub Combo3_Change()
On Error Resume Next
grd18.Clear
End Sub

Private Sub Combo3_Click()
On Error Resume Next
Combo3_Change
End Sub

Private Sub Combo4_Change()
On Error Resume Next
grd13.Clear
grd13.Rows = 1

End Sub

Private Sub Combo4_Click()
On Error Resume Next
Combo4_Change
End Sub


Private Sub Combo5_Change()
On Error Resume Next
grd18.Clear

End Sub

Private Sub Combo5_Click()
On Error Resume Next
Combo5_Change
End Sub

Private Sub Combo6_Change()
On Error Resume Next
grd16.Clear
Label21.Caption = "0"
Label35.Caption = "0"
Label36.Caption = "0"
Label37.Caption = "0"
Call date_dt
End Sub

Private Sub Combo6_Click()
On Error Resume Next
Combo6_Change
End Sub

Private Sub Combo7_Change()
On Error Resume Next
grd17.Clear
Label33.Caption = "0"
Label32.Caption = "0"
Label29.Caption = "0"

End Sub

Private Sub Combo7_Click()
On Error Resume Next
Combo7_Change
End Sub

Private Sub Combo8_Change()
On Error Resume Next
Call cont
Do While Not an.EOF
If an!ann = Combo8.Text Then
Label45.Caption = an!an1
Label46.Caption = an!an2
DT4.Year = Label45.Caption
DT4_Change
Exit Sub
End If
an.MoveNext
Loop

End Sub

Private Sub Combo8_Click()
On Error Resume Next
Combo8_Change
End Sub

Private Sub Command1_Click()
On Error Resume Next
grd10.Clear
grd10.Rows = 1
grd10.Visible = False
grd11.Clear
grd11.Rows = 1
grd11.Visible = False
Call chargegrd10_11
grd10.Visible = True
grd11.Visible = True
End Sub

Private Sub Command11_Click()
On Error Resume Next
Picture14.Visible = False
Picture15.Visible = True

End Sub

Private Sub Command12_Click()
On Error Resume Next
Picture15.Visible = False
Picture14.Visible = True
End Sub





Private Sub Command13_Click()
On Error Resume Next
Text1.Text = Trim(Text1.Text)
If Text1.Text = "" Then
MsgBox "ÌÃ» «œŒ«· —ﬁ„ «·Ê’·", vbCritical
Exit Sub
End If
Call cont
Do While Not rc.EOF
If Text1.Text = rc!rec Then
Label30.Caption = rc!ser
Label53.Caption = rc!nom
Label54.Caption = rc!cla
Label55.Caption = rc!num
Label56.Caption = rc!mon
Label57.Caption = rc!mois
Label58.Caption = rc!dat
Exit Sub
End If
rc.MoveNext
Loop
MsgBox "«·Ê’· €Ì— „ÊÃÊœ", vbExclamation
Text1.Text = ""
Text1.SetFocus
End Sub

Private Sub Command14_Click()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim s As String
Text1.Text = Trim(Text1.Text)
If Text1.Text = "" Then
MsgBox "ÌÃ» «œŒ«· —ﬁ„ «·Ê’·", vbCritical
Text1.SetFocus
Exit Sub
End If
If Label30.Caption = "" Then
MsgBox "ÌÃ» ⁄—÷  ›«’Ì· «·Ê’·", vbCritical
Text1.SetFocus
Exit Sub
End If
Call cont
Do While Not cr.EOF
If cr!rec = Text1.Text Then
cr!rec = Text1.Text
cr!mon = Label56.Caption
cr!dat = Label58.Caption
cr!act = "”·Ì„"
cr.Update
n = grd11.Rows
For i = 1 To n - 1
grd11.row = i
grd11.Col = 1
s = grd11.Text
If s = Text1.Text Then
grd11.row = i
grd11.Col = 1
grd11.Text = Text1.Text
grd11.Col = 2
grd11.Text = "”·Ì„"
End If
Next i
Exit Sub
End If
cr.MoveNext
Loop
cr.AddNew
cr!rec = Text1.Text
cr!mon = Label56.Caption
cr!dat = Label58.Caption
cr!act = "”·Ì„"
cr.Update
n = grd11.Rows
grd11.Rows = n + 1
grd11.row = n
grd11.Col = 1
grd11.Text = Text1.Text
grd11.Col = 2
grd11.Text = "”·Ì„"
End Sub

Private Sub Command15_Click()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim s As String
Text1.Text = Trim(Text1.Text)
If Text1.Text = "" Then
MsgBox "ÌÃ» «œŒ«· —ﬁ„ «·Ê’·", vbCritical
Text1.SetFocus
Exit Sub
End If
If Label30.Caption = "" Then
MsgBox "ÌÃ» ⁄—÷  ›«’Ì· «·Ê’·", vbCritical
Text1.SetFocus
Exit Sub
End If
Call cont
Do While Not cr.EOF
If cr!rec = Text1.Text Then
cr!rec = Text1.Text
cr!mon = Label56.Caption
cr!dat = Label58.Caption
cr!act = "€Ì— ”·Ì„"
cr.Update
n = grd11.Rows
For i = 1 To n - 1
grd11.row = i
grd11.Col = 1
s = grd11.Text
If s = Text1.Text Then
grd11.row = i
grd11.Col = 1
grd11.Text = Text1.Text
grd11.Col = 2
grd11.Text = "€Ì— ”·Ì„"
End If
Next i
Exit Sub
End If
cr.MoveNext
Loop
cr.AddNew
cr!rec = Text1.Text
cr!mon = Label56.Caption
cr!dat = Label58.Caption
cr!act = "€Ì— ”·Ì„"
cr.Update
n = grd11.Rows
grd11.Rows = n + 1
grd11.row = n
grd11.Col = 1
grd11.Text = Text1.Text
grd11.Col = 2
grd11.Text = "€Ì— ”·Ì„"

End Sub

Private Sub Command16_Click()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim s As String
Text1.Text = Trim(Text1.Text)
Text2.Text = Trim(Text2.Text)
If Text1.Text = "" Then
MsgBox "ÌÃ» «œŒ«· —ﬁ„ «·Ê’·", vbCritical
Text1.SetFocus
Exit Sub
End If
If Label30.Caption = "" Then
MsgBox "ÌÃ» ⁄—÷  ›«’Ì· «·Ê’·", vbCritical
Text1.SetFocus
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "ÌÃ» «œŒ«· «·„·«ÕŸ…", vbCritical
Text2.SetFocus
Exit Sub
End If
Call cont
Do While Not cr.EOF
If cr!rec = Text1.Text Then
cr!rec = Text1.Text
cr!mon = Label56.Caption
cr!dat = Label58.Caption
cr!act = Text2.Text
cr.Update
n = grd11.Rows
For i = 1 To n - 1
grd11.row = i
grd11.Col = 1
s = grd11.Text
If s = Text1.Text Then
grd11.row = i
grd11.Col = 1
grd11.Text = Text1.Text
grd11.Col = 2
grd11.Text = Text2.Text
End If
Next i
Exit Sub
End If
cr.MoveNext
Loop
cr.AddNew
cr!rec = Text1.Text
cr!mon = Label56.Caption
cr!dat = Label58.Caption
cr!act = Text2.Text
cr.Update
n = grd11.Rows
grd11.Rows = n + 1
grd11.row = n
grd11.Col = 1
grd11.Text = Text1.Text
grd11.Col = 2
grd11.Text = Text2.Text
End Sub

Private Sub Command17_Click()
On Error Resume Next
grd1.Clear
grd1.Rows = 1
grd1.Visible = False
Call chargegrd1_2
grd1.Visible = True
End Sub

Private Sub Command2_Click()
On Error GoTo u
Dim i As Double
Dim j As Double
Dim n As Double
Dim m As Double
Dim k As Double
If Combo2.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·ﬁ”„", vbCritical
Exit Sub
End If
If Text5.Text = "" Then
MsgBox "«œŒ· «·—”Ê„ «· ”ÃÌ·Ì…", vbCritical
Text5.SetFocus
Exit Sub
End If
If Combo4.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— ⁄œœ ‘ÂÊ— «·”‰… «·œ—«”Ì…", vbCritical
Exit Sub
End If
If grd13.Rows < 2 Then
MsgBox "·«  ÊÃœ »Ì«‰« ", vbCritical
Exit Sub
End If
If Check2.Value = 1 Then
FileCopy App.Path & "\Totalpayetuu00000.xls", App.Path & "\Totalpayetudiants.xls"
Command2.Enabled = False
n = grd13.Rows
Set kb = CreateObject("Excel.application")
kb.Workbooks.Open (App.Path & "\Totalpayetudiants.xls")
kb.Visible = True
For i = 0 To n - 1
For j = 1 To 19
k = 20 - j
grd13.row = i
grd13.Col = j
kb.Workbooks("Totalpayetudiants").Sheets(1).Cells(i + 3, k).Value = grd13.Text
Next j
Next i
kb.Workbooks("Totalpayetudiants").Sheets(1).Range("K1").Value = Combo2.Text
'Workbooks("Etudiants").Close savechanges:=False
'Worksheets(1).Activate
Command2.Enabled = True
Exit Sub
Else
FileCopy App.Path & "\Totalpayetuures00000.xls", App.Path & "\TotalpayetudiantsresumÈ.xls"
Command2.Enabled = False
n = grd13.Rows
Set kb = CreateObject("Excel.application")
kb.Workbooks.Open (App.Path & "\TotalpayetudiantsresumÈ.xls")
kb.Visible = True
For i = 0 To n - 1
For j = 1 To 19
If j = 1 Then
k = 6
grd13.row = i
grd13.Col = j
kb.Workbooks("TotalpayetudiantsresumÈ").Sheets(1).Cells(i + 3, k).Value = grd13.Text
Else
k = 20 - j
grd13.row = i
grd13.Col = j
kb.Workbooks("TotalpayetudiantsresumÈ").Sheets(1).Cells(i + 3, k).Value = grd13.Text
End If
If j = 1 Then
j = 14
End If
Next j
Next i
kb.Workbooks("TotalpayetudiantsresumÈ").Sheets(1).Range("E1").Value = Combo2.Text
'Workbooks("Etudiants").Close savechanges:=False
'Worksheets(1).Activate
Command2.Enabled = True
Exit Sub

End If
u:
MsgBox "ÌÃ» «€·«ﬁ ’›Õ… «ﬂ”· «‰ ﬂ«‰  „› ÊÕ…", vbExclamation
Command2.Enabled = True

End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim mm As String
If Combo3.Text = "" Then
MsgBox "ÌÃ» «Œ Ì«— «·ﬁ”„", vbCritical
Exit Sub
End If
If Combo5.Text = "" Then
MsgBox "ÌÃ» «Œ Ì«— «·‘Â—", vbCritical
Exit Sub
End If
grd18.Visible = False
Call chargegrd18_MSChart1
grd18.Visible = True
MSChart1.RowCount = 1
MSChart1.row = 1
MSChart1.Column = 1
MSChart1.data = Label20.Caption
MSChart1.Column = 2
MSChart1.data = Label22.Caption
MSChart1.Column = 3
MSChart1.data = Label23.Caption
mm = ""
MSChart1.RowLabel = mm
MSChart1.Column = 1
MSChart1.ColumnLabel = "«·œ«Œ·"
MSChart1.Column = 2
MSChart1.ColumnLabel = "«·„” Õﬁ"
MSChart1.Column = 3
MSChart1.ColumnLabel = "«·«Ì—«œ"
End Sub

Private Sub Command31_Click()
On Error Resume Next
grd1.Visible = False
Call chargegrd1
Call chargec1
Call chargegrd6_7_8_9
grd14.Visible = False
Call profits_14
grd14.Visible = True
Call balance
grd1.Visible = True
If face.SBB1.Panels(10).Text = "«—‘Ì›" Then
Command8.Enabled = False
Else
Command8.Enabled = True
End If
If face.Caption = "TEST" Then
Command8.Enabled = False
End If
End Sub


Private Sub Command4_Click()
On Error Resume Next
Dim mm As String
If Combo7.Text = "" Then
MsgBox "ÌÃ» «Œ Ì«— «·‘Â—", vbCritical
Exit Sub
End If
grd17.Visible = False
Call chargegrd17_MSChart2
grd17.Visible = True
MSChart2.RowCount = 1
MSChart2.row = 1
MSChart2.Column = 1
MSChart2.data = Label33.Caption
MSChart2.Column = 2
MSChart2.data = Label32.Caption
MSChart2.Column = 3
MSChart2.data = Label29.Caption
mm = ""
MSChart2.RowLabel = mm
MSChart2.Column = 1
MSChart2.ColumnLabel = "«·œ«Œ·"
MSChart2.Column = 2
MSChart2.ColumnLabel = "«·„” Õﬁ"
MSChart2.Column = 3
MSChart2.ColumnLabel = "«·«Ì—«œ"

End Sub

Private Sub Command5_Click()
On Error Resume Next
Text5.Text = Trim(Text5.Text)
'If Combo3.Text = "" Then
'MsgBox "ﬁ„ »«Œ Ì«— «·Õ«·… «· ”ÃÌ·Ì…", vbCritical
'Exit Sub
'End If
If Combo2.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·ﬁ”„", vbCritical
Exit Sub
End If
If Text5.Text = "" Then
MsgBox "«œŒ· «·—”Ê„ «· ”ÃÌ·Ì…", vbCritical
Text5.SetFocus
Exit Sub
End If
If Combo4.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— ⁄œœ ‘ÂÊ— «·”‰… «·œ—«”Ì…", vbCritical
Exit Sub
End If
Command5.Enabled = False
grd13.Visible = False
Call chargegrd13
grd13.Col = 0
grd13.Sort = 1
grd13.Visible = True
Command5.Enabled = True
End Sub

Private Sub Command6_Click()
On Error Resume Next
Dim mm As String
If Combo6.Text = "" Then
MsgBox "ÌÃ» «Œ Ì«— «·‘Â—", vbCritical
Exit Sub
End If
grd16.Visible = False
Call chargegrd16_MSChart3
grd16.Visible = True
MSChart3.RowCount = 1
MSChart3.row = 1
MSChart3.Column = 1
MSChart3.data = Label21.Caption
MSChart3.Column = 2
MSChart3.data = Label35.Caption
MSChart3.Column = 3
MSChart3.data = Label36.Caption
MSChart3.Column = 4
MSChart3.data = Label37.Caption
mm = ""
MSChart3.RowLabel = mm
MSChart3.Column = 1
MSChart3.ColumnLabel = "«·œ«Œ·"
MSChart3.Column = 2
MSChart3.ColumnLabel = "«·„” Õﬁ"
MSChart3.Column = 3
MSChart3.ColumnLabel = "«·„’—Ê›"
MSChart3.Column = 4
MSChart3.ColumnLabel = "«·«Ì—«œ"
End Sub

Private Sub Command7_Click()
On Error Resume Next
Dim mm As String
grd15.Visible = False
Call chargegrd15_MSChart4
grd15.Visible = True
MSChart4.RowCount = 1
MSChart4.row = 1
MSChart4.Column = 1
MSChart4.data = Label39.Caption
MSChart4.Column = 2
MSChart4.data = Label34.Caption
MSChart4.Column = 3
MSChart4.data = Label28.Caption
MSChart4.Column = 4
MSChart4.data = Label27.Caption
mm = ""
MSChart4.RowLabel = mm
MSChart4.Column = 1
MSChart4.ColumnLabel = "«·œ«Œ·"
MSChart4.Column = 2
MSChart4.ColumnLabel = "«·„” Õﬁ"
MSChart4.Column = 3
MSChart4.ColumnLabel = "«·„’—Ê›"
MSChart4.Column = 4
MSChart4.ColumnLabel = "«·«Ì—«œ"
End Sub

Private Sub Command8_Click()
On Error Resume Next
Dim mot0 As String
Dim sb1 As Double
Dim mot1 As String
Dim sb2 As Double
Dim mot2 As String
Dim sb3 As Double
Dim sb4 As Double
Dim sb5 As Double
Dim sb6 As Double
Dim mot3 As String
Dim mot4 As String
Dim mot5 As String
Picture12.Visible = False
mot0 = "≈‰ › Õ √Ì ”‰… œ—«”Ì… ÃœÌœ… Ì⁄‰Ì €·ﬁ «·”‰… «·œ—«”Ì… «·Õ«·Ì… „„« Ì„‰⁄ «„ﬂ«‰Ì… «· ⁄œÌ· ⁄·ÌÂ« „” ﬁ»·« „ÿ·ﬁ«, ·–·ﬂ Ì‰»€Ì „—«Ã⁄… «·¬ Ì: "
sb1 = Label12.Caption
If sb1 > 0 Then
mot1 = "„” Õﬁ«  «·√”« –… : " + Label12.Caption + " Ì‰»€Ì œ›⁄Â«"
End If
sb2 = Label42.Caption
If sb2 > 0 Then
mot1 = "œÌÊ‰ ⁄·Ï «·√”« –… : " + Label42.Caption + " Ì‰»€Ì  ”œÌœÂ«"
End If
sb3 = Label16.Caption
sb4 = Label10.Caption
sb5 = Label13.Caption
If sb3 <= 0 And sb4 > 0 Then
sb3 = sb3 * -1
sb6 = sb3 + sb4
mot4 = sb6
mot3 = "œÌÊ‰ ⁄·Ï «·‘—ﬂ«¡ : " + mot4 + " Ì‰»€Ì  ”œÌœÂ«"
End If
If sb3 >= 0 And sb4 > 0 Then
sb6 = sb3 - sb4
mot4 = sb6
If sb6 > 0 Then
mot3 = "„” Õﬁ«  «·‘—ﬂ«¡ : " + mot4 + " Ì‰»€Ì œ›⁄Â«"
Else
sb6 = sb6 * -1
mot4 = sb6
mot3 = "œÌÊ‰ ⁄·Ï «·‘—ﬂ«¡ : " + mot4 + " Ì‰»€Ì  ”œÌœÂ«"
End If
End If
If sb3 <= 0 And sb5 > 0 Then
sb3 = sb3 * -1
sb6 = sb3 - sb5
mot4 = sb6
If sb6 > 0 Then
mot3 = "œÌÊ‰ ⁄·Ï «·‘—ﬂ«¡ : " + mot4 + " Ì‰»€Ì  ”œÌœÂ«"
Else
sb6 = sb6 * -1
mot4 = sb6
mot3 = "„” Õﬁ«  «·‘—ﬂ«¡ : " + mot4 + " Ì‰»€Ì œ›⁄Â«"
End If
End If
If sb3 >= 0 And sb5 > 0 Then
sb6 = sb3 + sb5
mot4 = sb6
mot3 = "„” Õﬁ«  «·‘—ﬂ«¡ : " + mot4 + " Ì‰»€Ì œ›⁄Â«"
End If
mot5 = mot0 & Chr(10)
If (sb1 + sb2) > 0 Then
mot5 = mot5 & Chr(10) & mot1 & Chr(10)
End If
If sb6 > 0 Then
mot5 = mot5 & Chr(10) & mot3 & Chr(10)
End If
sb7 = sb1 + sb2 + sb6
If sb7 > 0 Then
g = MsgBox(mot5 & Chr(10) & " √„ √‰ﬂ  —Ìœ €÷ «·ÿ—› ⁄‰ –·ﬂ ø", vbInformation + vbYesNo + arabic, "AGEP6")
If g = vbYes Then
Label26.Caption = face.SBB1.Panels(9).Text
Call chargec8
Picture12.Visible = True
Exit Sub
Else
Exit Sub
End If
End If
Label26.Caption = face.SBB1.Panels(9).Text
Call chargec8
Picture12.Visible = True
Exit Sub
End Sub

Private Sub Command9_Click()
On Error Resume Next
Dim k As Double
Dim n As Double
Dim c As Double
Dim au As Double
Dim a As Double
Dim x$
Dim nomar As String
Dim codsp As String
Dim decsp As String
Dim codsd As String
Dim decsd As String
Dim codsr As String
Dim decsr As String
Dim codst As String
Dim decst As String
Dim comp As String
Dim comd As String
Dim comr As String
Dim comt As String
Dim dat1 As Date
Dim dat2 As Date
If Combo8.Text = "" Then
MsgBox "ÌÃ» «œŒ«· «·”‰… «·œ—«”Ì… «·ÃœÌœ…", vbCritical
Exit Sub
End If
If Combo8.Text = Label26.Caption Then
MsgBox "ÌÃ» √‰  Œ ·› «·”‰… «·œ—«”Ì… «·„‰’—„… ⁄‰ «·”‰… «·œ—«”Ì… «·ÃœÌœ…", vbCritical
Exit Sub
End If
DT5.Value = DT4.Value + 364
a = DT4.Year
If a < Val(Label45.Caption) Then
MsgBox "«· «—ÌŒ «·„œŒ· ”«»ﬁ ··”‰… «·œ—«”Ì… " + Combo8.Text, vbCritical
Exit Sub
End If
If a > Val(Label46.Caption) Then
MsgBox "«· «—ÌŒ «·„œŒ· „ √Œ— ⁄‰ «·”‰… «·œ—«”Ì… " + Combo8.Text, vbCritical
Exit Sub
End If
dat1 = DT4.Value
dat2 = DT5.Value
If dat1 > dat2 Then
MsgBox " «—ÌŒ »œ«Ì…«·”‰…«·œ—«”Ì… ÌÃ» √‰ ÌﬂÊ‰ ﬁ»·  «—ÌŒ ‰Â«Ì…«·”‰…«·œ—«”Ì…", vbCritical
Exit Sub
End If
'**** exist and no delete
Call cont2
Do While Not nn.EOF
If Combo8.Text = nn!ann And nn!sup = "0" Then
MsgBox "€Ì— „„ﬂ‰.. ·√‰ «·”‰… «·œ—«”Ì… «·ÃœÌœ… «·„œŒ·… Ã«—Ì «·⁄„· »Â«, ÌÃ» Õ–›Â« √Ê·« √Ê «Œ Ì«— ”‰… ÃœÌœ… √Œ—Ï", vbCritical
Exit Sub
End If
nn.MoveNext
Loop
Label47.Visible = True
'**** no next
If Combo8.Text <> Label41.Caption Then
g = MsgBox("«·”‰… «·œ—«”Ì… «·„‰’—„… ÂÌ " + Label26.Caption + " Ê«·”‰… «·œ—«”Ì… «· Ì ﬁ„  »«œŒ«·Â« ÂÌ " + Combo8.Text + " ›Ì ÕÌ‰ √‰Â ÌÃ» √‰  ﬂÊ‰ «·”‰… «·œ—«”Ì… «·ÃœÌœ… ÂÌ " + Label41.Caption + " ›Â·  —Ìœ €÷ «·ÿ—› ⁄‰ –·ﬂ ø", vbInformation + vbYesNo + arabic, "AGEP6")
'**** no next but accepte
If g = vbYes Then
'**** exist and delete
nomar = ""
k = 0
Call cont2
Do While Not nn.EOF
If Combo8.Text = nn!ann Then
k = 1
nn!act = "1"
nn!sup = "0"
nn.Update
End If
If Label26.Caption = nn!ann Then
nn!act = "0"
nn.Update
End If
nn.MoveNext
Loop
Call cont2
Do While Not ns.EOF
If Combo8.Text = ns!ann Then
nomar = ns!nom
End If
ns.MoveNext
Loop
x$ = Dir$(App.Path & "\" & nomar & ".txt")
If x$ <> "" Then
Kill App.Path & "\" & nomar & ".txt"
End If
'**** no existe
If k = 0 Then
nn.AddNew
nn!ann = Combo8.Text
nn!act = "1"
nn!sup = "0"
nn.Update
End If
'**** exist or no existe
FileCopy App.Path & "\ANNEEVIDE.mdb", App.Path & "\" & Combo8.Text & ".mdb"
FileCopy App.Path & "\CARTESVIDES.mdb", App.Path & "\C" & Combo8.Text & ".mdb"
Call cont3
'**** codes
Call cont
Do While Not cd.EOF
od.AddNew
For i = 0 To 4
od.Fields(i) = cd.Fields(i)
Next i
od.Update
cd.MoveNext
Loop
'**** series
ri.AddNew
For i = 0 To 10
ri.Fields(i) = sr.Fields(i)
If i = 9 Then
ri.Fields(i) = Combo8.Text
End If
Next i
ri.Fields(11) = DT4.Value
ri.Fields(12) = DT5.Value
ri.Update
'**** coffcients
cv.AddNew
For i = 0 To 28
cv.Fields(i) = cf.Fields(i)
Next i
cv.Update
'**** utilisateurs
Call cont
Do While Not ut.EOF
ti.AddNew
For i = 0 To 10
ti.Fields(i) = ut.Fields(i)
Next i
ti.Update
ut.MoveNext
Loop
'**** gestion de comptabilite
Call cont
Do While Not cd.EOF
If cd!der = "—√” «·„«·" Then
codsp = cd!cod
decsp = cd!dec
comp = cd!cas
End If
If cd!der = "«Ìœ«⁄" Then
codsd = cd!cod
decsd = cd!dec
comd = cd!cas
End If
If cd!der = "”Õ»" Then
codsr = cd!cod
decsr = cd!dec
comr = cd!cas
End If
If cd!cas = "Õ”«» «·‘—ﬂ«¡" Then
codst = cd!cod
decst = cd!dec
comt = cd!cas
End If
cd.MoveNext
Loop
'**** capital
Call cont3
'****** caisse
ai.AddNew
au = ai!aun
ai!cod = codsp
ai!dec = decsp
ai!mon = "0"
ai!mem = "—√” „«· „‰ «·”‰…«·œ—«”Ì… " + Label26.Caption
ai!cas = "Œ«—Ã"
ai!heu = Time$
ai!dat = DT4.Value
ai!ger = ""
ai!aut = au
ai!com = comp
ai!act = "1"
ai.Update
'****** journal
c = ri!ord
jn.AddNew
jn!cre = Label15.Caption
jn!deb = "0"
jn!dem = "„‰ Õ‹"
jn!com = "«·»‰ﬂ"
jn!dec = "—√” „«· „‰ «·”‰…«·œ—«”Ì… " + Label26.Caption
jn!ord = c
jn!dat = DT4.Value
jn!heu = Time$
jn!ger = ""
jn.Update
jn.AddNew
jn!cre = "0"
jn!deb = Label15.Caption
jn!dem = "≈·Ï Õ‹"
jn!com = "—√” «·„«·"
jn!dec = "—√” „«· „‰ «·”‰…«·œ—«”Ì… " + Label26.Caption
jn!ord = c
jn!dat = DT4.Value
jn!heu = Time$
jn!ger = ""
jn.Update
ri!ord = c + 1
ri.Update
bk.AddNew
bk!aut = au
bk!dec = "—√” „«· „‰ «·”‰…«·œ—«”Ì… " + Label26.Caption
bk!Mod = "—√” «·„«·"
bk!mon = Label15.Caption
bk!dat = DT4.Value
bk!heu = Time$
bk!ger = ""
bk!act = "1"
bk.Update
'***** deposer
'****** caisse
ai.AddNew
au = ai!aun
ai!cod = codsd
ai!dec = decsd
ai!mon = "0"
ai!mem = "—’Ìœ «·»‰ﬂ „‰ «·”‰… «·œ—«”Ì… " + Label26.Caption
ai!cas = "Œ«—Ã"
ai!heu = Time$
ai!dat = DT4.Value
ai!ger = ""
ai!aut = au
ai!com = comd
ai!act = "1"
ai.Update
'****** journal
c = ri!ord
jn.AddNew
jn!cre = Label8.Caption
jn!deb = "0"
jn!dem = "„‰ Õ‹"
jn!com = "«·»‰ﬂ"
jn!dec = "—’Ìœ «·»‰ﬂ „‰ «·”‰… «·œ—«”Ì… " + Label26.Caption
jn!ord = c
jn!dat = DT4.Value
jn!heu = Time$
jn!ger = ""
jn.Update
jn.AddNew
jn!cre = "0"
jn!deb = Label8.Caption
jn!dem = "≈·Ï Õ‹"
jn!com = "«·’‰œÊﬁ"
jn!dec = "—’Ìœ «·»‰ﬂ „‰ «·”‰… «·œ—«”Ì… " + Label26.Caption
jn!ord = c
jn!dat = DT4.Value
jn!heu = Time$
jn!ger = ""
jn.Update
'****** cascaisse
ri!ord = c + 1
ri.Update
bk.AddNew
bk!aut = au
bk!dec = "—’Ìœ «·»‰ﬂ „‰ «·”‰… «·œ—«”Ì… " + Label26.Caption
bk!Mod = "«Ìœ«⁄"
bk!mon = Label8.Caption
bk!dat = DT4.Value
bk!heu = Time$
bk!ger = ""
bk!act = "1"
bk.Update
'**** retirer
'****** caisse
ai.AddNew
au = ai!aun
ai!cod = codsr
ai!dec = decsr
ai!mon = "0"
ai!mem = "—√” „«· „‰ «·”‰…«·œ—«”Ì… " + Label26.Caption
ai!cas = "œ«Œ·"
ai!heu = Time$
ai!dat = DT4.Value
ai!ger = ""
ai!aut = au
ai!com = comr
ai!act = "1"
ai.Update
'****** journal
c = ri!ord
jn.AddNew
jn!cre = Label15.Caption
jn!deb = "0"
jn!dem = "„‰ Õ‹"
jn!com = "—√” «·„«·"
jn!dec = "—√” „«· „‰ «·”‰…«·œ—«”Ì… " + Label26.Caption
jn!ord = c
jn!dat = DT4.Value
jn!heu = Time$
jn!ger = ""
jn.Update
jn.AddNew
jn!cre = "0"
jn!deb = Label15.Caption
jn!dem = "≈·Ï Õ‹"
jn!com = "«·»‰ﬂ"
jn!dec = "—√” „«· „‰ «·”‰…«·œ—«”Ì… " + Label26.Caption
jn!ord = c
jn!dat = DT4.Value
jn!heu = Time$
jn!ger = ""
jn.Update
ri!ord = c + 1
ri.Update
bk.AddNew
bk!aut = au
bk!dec = "—√” „«· „‰ «·”‰…«·œ—«”Ì… " + Label26.Caption
bk!Mod = "”Õ»"
bk!mon = Label15.Caption
bk!dat = DT4.Value
bk!heu = Time$
bk!ger = ""
bk!act = "1"
bk.Update
'**** caisse
'****** caisse
ai.AddNew
au = ai!aun
ai!cod = "0000"
ai!dec = "«· Ê“Ì⁄"
ai!mon = Label9.Caption
ai!mem = "—’Ìœ «·’‰œÊﬁ „‰ «·”‰…«·œ—«”Ì… " + Label26.Caption
ai!cas = "œ«Œ·"
ai!heu = Time$
ai!dat = DT4.Value
ai!ger = ""
ai!aut = au
ai!com = "«· Ê“Ì⁄"
ai!act = "1"
ai.Update
'****** journal
c = ri!ord
jn.AddNew
jn!cre = Label9.Caption
jn!deb = "0"
jn!dem = "„‰ Õ‹"
jn!com = "«·’‰œÊﬁ"
jn!dec = "—√” „«· „‰ «·”‰…«·œ—«”Ì… " + Label26.Caption
jn!ord = c
jn!dat = DT4.Value
jn!heu = Time$
jn!ger = ""
jn.Update
jn.AddNew
jn!cre = "0"
jn!deb = Label9.Caption
jn!dem = "≈·Ï Õ‹"
jn!com = "«·’‰œÊﬁ"
jn!dec = "—√” „«· „‰ «·”‰…«·œ—«”Ì… " + Label26.Caption
jn!ord = c
jn!dat = DT4.Value
jn!heu = Time$
jn!ger = ""
jn.Update
ri!cca = Label9.Caption
ri!ord = c + 1
ri.Update

'**** end gestion
face.SBB1.Panels(9).Text = Label26.Caption
start.Label1.Caption = Label26.Caption
Call cont
co.Close
Name App.Path & "\" & Label26.Caption & ".mdb" As App.Path & "\R" & Label26.Caption & ".mdb"
face.SBB1.Panels(9).Text = Combo8.Text
start.Label1.Caption = Combo8.Text
MsgBox " „ › Õ ”‰… œ—«”Ì… ÃœÌœ… »‰Ã«Õ , ‰”√· «··Â √‰  ﬂÊ‰ „ﬂ··… »«·‰Ã«Õ", vbInformation
Name App.Path & "\" & Label26.Caption & ".mdb" As App.Path & "\R" & Label26.Caption & ".mdb"
Label47.Visible = False
Call cont2
Do While Not ns.EOF
If Combo8.Text = ns!ann Then
ns.Delete
Unload Me
Exit Sub
End If
ns.MoveNext
Loop
Label47.Visible = False
Unload Me
Exit Sub
End If
Label47.Visible = False
Exit Sub
End If
'**** next
'**** exist and delete
nomar = ""
k = 0
n = 0
Call cont2
Do While Not nn.EOF
If Combo8.Text = nn!ann Then
k = 1
nn!act = "1"
nn!sup = "0"
nn.Update
End If
If Label26.Caption = nn!ann Then
nn!act = "0"
nn.Update
End If
nn.MoveNext
Loop
Call cont2
Do While Not ns.EOF
If Combo8.Text = ns!ann Then
nomar = ns!nom
End If
ns.MoveNext
Loop
x$ = Dir$(App.Path & "\" & nomar & ".txt")
If x$ <> "" Then
Kill App.Path & "\" & nomar & ".txt"
End If
'**** no existe
If k = 0 Then
nn.AddNew
nn!ann = Combo8.Text
nn!act = "1"
nn!sup = "0"
nn.Update
End If
'**** exist or no existe
FileCopy App.Path & "\ANNEEVIDE.mdb", App.Path & "\" & Combo8.Text & ".mdb"
FileCopy App.Path & "\CARTESVIDES.mdb", App.Path & "\C" & Combo8.Text & ".mdb"
Call cont3
'**** codes
Call cont
Do While Not cd.EOF
od.AddNew
For i = 0 To 4
od.Fields(i) = cd.Fields(i)
Next i
od.Update
cd.MoveNext
Loop
'**** series
ri.AddNew
For i = 0 To 10
ri.Fields(i) = sr.Fields(i)
If i = 9 Then
ri.Fields(i) = Combo8.Text
End If
Next i
ri.Fields(11) = DT4.Value
ri.Fields(12) = DT5.Value
ri.Update
'**** coffcients
cv.AddNew
For i = 0 To 28
cv.Fields(i) = cf.Fields(i)
Next i
cv.Update
'**** utilisateurs
Call cont
Do While Not ut.EOF
ti.AddNew
For i = 0 To 10
ti.Fields(i) = ut.Fields(i)
Next i
ti.Update
ut.MoveNext
Loop
'**** gestion de comptabilite
Call cont
Do While Not cd.EOF
If cd!der = "—√” «·„«·" Then
codsp = cd!cod
decsp = cd!dec
comp = cd!cas
End If
If cd!der = "«Ìœ«⁄" Then
codsd = cd!cod
decsd = cd!dec
comd = cd!cas
End If
If cd!der = "”Õ»" Then
codsr = cd!cod
decsr = cd!dec
comr = cd!cas
End If
If cd!cas = "Õ”«» «·‘—ﬂ«¡" Then
codst = cd!cod
decst = cd!dec
comt = cd!cas
End If
cd.MoveNext
Loop
'**** capital
Call cont3
'****** caisse
ai.AddNew
au = ai!aun
ai!cod = codsp
ai!dec = decsp
ai!mon = "0"
ai!mem = "—√” „«· „‰ «·”‰…«·œ—«”Ì… " + Label26.Caption
ai!cas = "Œ«—Ã"
ai!heu = Time$
ai!dat = DT4.Value
ai!ger = ""
ai!aut = au
ai!com = comp
ai!act = "1"
ai.Update
'****** journal
c = ri!ord
jn.AddNew
jn!cre = Label15.Caption
jn!deb = "0"
jn!dem = "„‰ Õ‹"
jn!com = "«·»‰ﬂ"
jn!dec = "—√” „«· „‰ «·”‰…«·œ—«”Ì… " + Label26.Caption
jn!ord = c
jn!dat = DT4.Value
jn!heu = Time$
jn!ger = ""
jn.Update
jn.AddNew
jn!cre = "0"
jn!deb = Label15.Caption
jn!dem = "≈·Ï Õ‹"
jn!com = "—√” «·„«·"
jn!dec = "—√” „«· „‰ «·”‰…«·œ—«”Ì… " + Label26.Caption
jn!ord = c
jn!dat = DT4.Value
jn!heu = Time$
jn!ger = ""
jn.Update
ri!ord = c + 1
ri.Update
bk.AddNew
bk!aut = au
bk!dec = "—√” „«· „‰ «·”‰…«·œ—«”Ì… " + Label26.Caption
bk!Mod = "—√” «·„«·"
bk!mon = Label15.Caption
bk!dat = DT4.Value
bk!heu = Time$
bk!ger = ""
bk!act = "1"
bk.Update
'***** deposer
'****** caisse
ai.AddNew
au = ai!aun
ai!cod = codsd
ai!dec = decsd
ai!mon = "0"
ai!mem = "—’Ìœ «·»‰ﬂ „‰ «·”‰… «·œ—«”Ì… " + Label26.Caption
ai!cas = "Œ«—Ã"
ai!heu = Time$
ai!dat = DT4.Value
ai!ger = ""
ai!aut = au
ai!com = comd
ai!act = "1"
ai.Update
'****** journal
c = ri!ord
jn.AddNew
jn!cre = Label8.Caption
jn!deb = "0"
jn!dem = "„‰ Õ‹"
jn!com = "«·»‰ﬂ"
jn!dec = "—’Ìœ «·»‰ﬂ „‰ «·”‰… «·œ—«”Ì… " + Label26.Caption
jn!ord = c
jn!dat = DT4.Value
jn!heu = Time$
jn!ger = ""
jn.Update
jn.AddNew
jn!cre = "0"
jn!deb = Label8.Caption
jn!dem = "≈·Ï Õ‹"
jn!com = "«·’‰œÊﬁ"
jn!dec = "—’Ìœ «·»‰ﬂ „‰ «·”‰… «·œ—«”Ì… " + Label26.Caption
jn!ord = c
jn!dat = DT4.Value
jn!heu = Time$
jn!ger = ""
jn.Update
'****** cascaisse
ri!ord = c + 1
ri.Update
bk.AddNew
bk!aut = au
bk!dec = "—’Ìœ «·»‰ﬂ „‰ «·”‰… «·œ—«”Ì… " + Label26.Caption
bk!Mod = "«Ìœ«⁄"
bk!mon = Label8.Caption
bk!dat = DT4.Value
bk!heu = Time$
bk!ger = ""
bk!act = "1"
bk.Update
'**** retirer
'****** caisse
ai.AddNew
au = ai!aun
ai!cod = codsr
ai!dec = decsr
ai!mon = "0"
ai!mem = "—√” „«· „‰ «·”‰…«·œ—«”Ì… " + Label26.Caption
ai!cas = "œ«Œ·"
ai!heu = Time$
ai!dat = DT4.Value
ai!ger = ""
ai!aut = au
ai!com = comr
ai!act = "1"
ai.Update
'****** journal
c = ri!ord
jn.AddNew
jn!cre = Label15.Caption
jn!deb = "0"
jn!dem = "„‰ Õ‹"
jn!com = "—√” «·„«·"
jn!dec = "—√” „«· „‰ «·”‰…«·œ—«”Ì… " + Label26.Caption
jn!ord = c
jn!dat = DT4.Value
jn!heu = Time$
jn!ger = ""
jn.Update
jn.AddNew
jn!cre = "0"
jn!deb = Label15.Caption
jn!dem = "≈·Ï Õ‹"
jn!com = "«·»‰ﬂ"
jn!dec = "—√” „«· „‰ «·”‰…«·œ—«”Ì… " + Label26.Caption
jn!ord = c
jn!dat = DT4.Value
jn!heu = Time$
jn!ger = ""
jn.Update
ri!ord = c + 1
ri.Update
bk.AddNew
bk!aut = au
bk!dec = "—√” „«· „‰ «·”‰…«·œ—«”Ì… " + Label26.Caption
bk!Mod = "”Õ»"
bk!mon = Label15.Caption
bk!dat = DT4.Value
bk!heu = Time$
bk!ger = ""
bk!act = "1"
bk.Update
'**** caisse
'****** caisse
ai.AddNew
au = ai!aun
ai!cod = "0000"
ai!dec = "«· Ê“Ì⁄"
ai!mon = Label9.Caption
ai!mem = "—’Ìœ «·’‰œÊﬁ „‰ «·”‰…«·œ—«”Ì… " + Label26.Caption
ai!cas = "œ«Œ·"
ai!heu = Time$
ai!dat = DT4.Value
ai!ger = ""
ai!aut = au
ai!com = "«· Ê“Ì⁄"
ai!act = "1"
ai.Update
'****** journal
c = ri!ord
jn.AddNew
jn!cre = Label9.Caption
jn!deb = "0"
jn!dem = "„‰ Õ‹"
jn!com = "«·’‰œÊﬁ"
jn!dec = "—√” „«· „‰ «·”‰…«·œ—«”Ì… " + Label26.Caption
jn!ord = c
jn!dat = DT4.Value
jn!heu = Time$
jn!ger = ""
jn.Update
jn.AddNew
jn!cre = "0"
jn!deb = Label9.Caption
jn!dem = "≈·Ï Õ‹"
jn!com = "«·’‰œÊﬁ"
jn!dec = "—√” „«· „‰ «·”‰…«·œ—«”Ì… " + Label26.Caption
jn!ord = c
jn!dat = DT4.Value
jn!heu = Time$
jn!ger = ""
jn.Update
ri!cca = Label9.Caption
ri!ord = c + 1
ri.Update

'**** end gestion
face.SBB1.Panels(9).Text = Label26.Caption
start.Label1.Caption = Label26.Caption
Call cont
co.Close
Name App.Path & "\" & Label26.Caption & ".mdb" As App.Path & "\R" & Label26.Caption & ".mdb"
face.SBB1.Panels(9).Text = Combo8.Text
start.Label1.Caption = Combo8.Text
MsgBox " „ › Õ ”‰… œ—«”Ì… ÃœÌœ… »‰Ã«Õ , ‰”√· «··Â √‰  ﬂÊ‰ „ﬂ··… »«·‰Ã«Õ", vbInformation
Label47.Visible = False
Call cont2
Do While Not ns.EOF
If Combo8.Text = ns!ann Then
ns.Delete
Unload Me
Exit Sub
End If
ns.MoveNext
Loop
Label47.Visible = False
Unload Me
End Sub


Private Sub DT4_Change()
On Error Resume Next
DT5.Value = DT4.Value + 364
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
DT1.Value = Date - 30
DT2.Value = Date
Call chargec2
Call chargec3
Call chargec6
Label2.Caption = sr!pou
Label24.Caption = sr!pou
If sr!dat = "Rien" Then
DT1.Value = Date
DT2.Value = Date
Else
DT1.Value = sr!dat
DT2.Value = sr!dtf
End If
Combo4.Text = "9"
Check2.Value = 0
Combo1.Clear
DT11.Value = sr!dat
DT12.Value = sr!dtf
DT21.Value = Date
DT22.Value = Date
DT23.Value = Date
DT24.Value = Date
Text3.Text = DT11.Year
Text4.Text = DT12.Year
End Sub
Private Sub chargegrd1()
On Error Resume Next
'On Error Resume Next
Dim i As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim e As Double
Dim es As Double
Dim s As Double
Dim ss As Double
Dim cre As Double
Dim c1 As Double
Dim sc1 As Double
Dim c2 As Double
Dim sc2 As Double
Dim b1 As Double
Dim sb1 As Double
Dim b2 As Double
Dim sb2 As Double
Dim a1 As Double
Dim sa1 As Double
Dim a2 As Double
Dim sa2 As Double
Dim t1 As Double
Dim st1 As Double
Dim t2 As Double
Dim st2 As Double
Dim f1 As Double
Dim sf1 As Double
Dim f2 As Double
Dim sf2 As Double
Dim d1 As Double
Dim sd1 As Double
Dim d2 As Double
Dim sd2 As Double
Dim p1 As Double
Dim sp1 As Double
Dim p2 As Double
Dim sp2 As Double
Dim w1 As Double
Dim sw1 As Double
Dim w2 As Double
Dim sw2 As Double
Dim o1 As Double
Dim so1 As Double
Dim o2 As Double
Dim so2 As Double
grd2.Clear
grd2.Cols = 3
grd2.Rows = 1
grd2.ColWidth(0) = 2300
grd2.ColWidth(1) = 2300
grd2.ColWidth(2) = 2000
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.row = 0
grd2.Col = 0
grd2.Text = "„Ã„Ê⁄ «·Ã«‰» «·„œÌ‰"
grd2.Col = 1
grd2.Text = "„Ã„Ê⁄ «·Ã«‰» «·œ«∆‰"
grd2.Col = 2
grd2.Text = "«·Õ”«»"
grd3.Clear
grd3.Cols = 3
grd3.Rows = 1
grd3.ColWidth(0) = 2300
grd3.ColWidth(1) = 2300
grd3.ColWidth(2) = 2000
grd3.ColAlignment(0) = 1
grd3.ColAlignment(1) = 1
grd3.ColAlignment(2) = 1
grd3.row = 0
grd3.Col = 0
grd3.Text = "„Ã„Ê⁄ «·√—’œ… «·„œÌ‰…"
grd3.Col = 1
grd3.Text = "„Ã„Ê⁄ «·√—’œ… «·œ«∆‰…"
grd3.Col = 2
grd3.Text = "«·Õ”«»"
grd2.Rows = 11
grd3.Rows = 11
grd2.Col = 2
grd2.row = 1
grd2.Text = "—√” «·„«·"
grd2.row = 2
grd2.Text = "«·»‰ﬂ"
grd2.row = 3
grd2.Text = "«·’‰œÊﬁ"
grd2.row = 4
grd2.Text = "«· ·«„Ì–"
grd2.row = 5
grd2.Text = "«·√”« –…"
grd2.row = 6
grd2.Text = "«·„’—Ê›« "
grd2.row = 7
grd2.Text = "«·‘—ﬂ«¡"
grd2.row = 8
grd2.Text = "«·√—»«Õ"
grd2.row = 9
grd2.Text = "«·⁄„«·"
grd2.row = 10
grd2.Text = "«·„Ã„Ê⁄"
grd3.Col = 2
grd3.row = 1
grd3.Text = "—√” «·„«·"
grd3.row = 2
grd3.Text = "«·»‰ﬂ"
grd3.row = 3
grd3.Text = "«·’‰œÊﬁ"
grd3.row = 4
grd3.Text = "«· ·«„Ì–"
grd3.row = 5
grd3.Text = "«·√”« –…"
grd3.row = 6
grd3.Text = "«·„’—Ê›« "
grd3.row = 7
grd3.Text = "«·‘—ﬂ«¡"
grd3.row = 8
grd3.Text = "«·√—»«Õ"
grd3.row = 9
grd3.Text = "«·⁄„«·"
grd3.row = 10
grd3.Text = "«·„Ã„Ê⁄"
'***** grd1
grd1.Clear
grd1.Cols = 9
grd1.Rows = 1
grd1.ColWidth(0) = 1300
grd1.ColWidth(1) = 1300
grd1.ColWidth(2) = 800
grd1.ColWidth(3) = 1300
grd1.ColWidth(4) = 4000
grd1.ColWidth(5) = 800
grd1.ColWidth(6) = 1300
grd1.ColWidth(7) = 1300
grd1.ColWidth(8) = 1500
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 1
grd1.ColAlignment(6) = 1
grd1.ColAlignment(7) = 1
grd1.ColAlignment(8) = 1
grd1.row = 0
grd1.Col = 0
grd1.Text = "„œÌ‰"
grd1.Col = 1
grd1.Text = "œ«∆‰"
grd1.Col = 2
grd1.Text = "„‰/≈·Ï"
grd1.Col = 3
grd1.Text = "«·Õ”«»"
grd1.Col = 4
grd1.Text = "«·»Ì«‰"
grd1.Col = 5
grd1.Text = "—ﬁ„ «·ﬁÌœ"
grd1.Col = 6
grd1.Text = "«· «—ÌŒ"
grd1.Col = 7
grd1.Text = "«·”«⁄…"
grd1.Col = 8
grd1.Text = "«·„‰›–"
i = 1
es = 0
ss = 0
cre = 0
sc1 = 0
sc2 = 0
sb1 = 0
sb2 = 0
sa1 = 0
sa2 = 0
st1 = 0
st2 = 0
sf1 = 0
sf2 = 0
sd1 = 0
sd2 = 0
sp1 = 0
sp2 = 0
sw1 = 0
sw2 = 0
so1 = 0
so2 = 0
dat1 = DT1.Value
dat2 = DT2.Value
Call cont
grd1.Rows = jr.RecordCount + 5
Do While Not jr.EOF
s = 0
e = 0
dat3 = jr!dat
If dat3 >= dat1 And dat3 <= dat2 Then
grd1.row = i
grd1.Col = 0
grd1.Text = jr!cre
grd1.Col = 1
grd1.Text = jr!deb
grd1.Col = 2
grd1.Text = jr!dem
grd1.Col = 3
grd1.Text = jr!com
grd1.Col = 4
grd1.Text = jr!dec
grd1.Col = 5
grd1.Text = jr!ord
grd1.Col = 6
grd1.Text = jr!dat
grd1.Col = 7
grd1.Text = jr!heu
grd1.Col = 8
grd1.Text = jr!ger
e = jr!cre
es = es + e
s = jr!deb
ss = ss + s
i = i + 1
If jr!com = "—√” «·„«·" Then
c1 = jr!cre
sc1 = sc1 + c1
c2 = jr!deb
sc2 = sc2 + c2
End If
If jr!com = "«·»‰ﬂ" Then
b1 = jr!cre
sb1 = sb1 + b1
b2 = jr!deb
sb2 = sb2 + b2
End If
If jr!com = "«·’‰œÊﬁ" Then
a1 = jr!cre
sa1 = sa1 + a1
a2 = jr!deb
sa2 = sa2 + a2
End If
If jr!com = "«· ·«„Ì–" Then
t1 = jr!cre
st1 = st1 + t1
t2 = jr!deb
st2 = st2 + t2
End If
If jr!com = "«·√”« –…" Then
f1 = jr!cre
sf1 = sf1 + f1
f2 = jr!deb
sf2 = sf2 + f2
End If
If jr!com = "«·„’—Ê›« " Then
d1 = jr!cre
sd1 = sd1 + d1
d2 = jr!deb
sd2 = sd2 + d2
End If
If jr!com = "«·‘—ﬂ«¡" Then
p1 = jr!cre
sp1 = sp1 + p1
p2 = jr!deb
sp2 = sp2 + p2
End If
If jr!com = "«·⁄„«·" Then
o1 = jr!cre
so1 = so1 + o1
o2 = jr!deb
so2 = so2 + o2
End If
If jr!com = "«·√—»«Õ" Or jr!com = "«·Œ”«∆—" Then
w1 = jr!cre
sw1 = sw1 + w1
w2 = jr!deb
sw2 = sw2 + w2
End If
End If
jr.MoveNext
Loop
grd1.row = i
grd1.Col = 0
grd1.Text = es
grd1.Col = 1
grd1.Text = ss
grd1.Col = 4
grd1.Text = "«·„Ã„Ê⁄"
grd1.Rows = i + 1
grd2.row = 1
grd2.Col = 0
grd2.Text = sc1
grd2.Col = 1
grd2.Text = sc2
grd2.row = 2
grd2.Col = 0
grd2.Text = sb1
grd2.Col = 1
grd2.Text = sb2
grd2.row = 3
grd2.Col = 0
grd2.Text = sa1
grd2.Col = 1
grd2.Text = sa2
grd2.row = 4
grd2.Col = 0
grd2.Text = st1
grd2.Col = 1
grd2.Text = st2
grd2.row = 5
grd2.Col = 0
grd2.Text = sf1
grd2.Col = 1
grd2.Text = sf2
grd2.row = 6
grd2.Col = 0
grd2.Text = sd1
grd2.Col = 1
grd2.Text = sd2
grd2.row = 7
grd2.Col = 0
grd2.Text = sp1
grd2.Col = 1
grd2.Text = sp2
grd2.row = 8
grd2.Col = 0
grd2.Text = sw1
grd2.Col = 1
grd2.Text = sw2
grd2.row = 9
grd2.Col = 0
grd2.Text = so1
grd2.Col = 1
grd2.Text = so2
grd2.row = 10
grd2.Col = 0
grd2.Text = sc1 + sb1 + sa1 + st1 + sf1 + sd1 + sp1 + sw1 + so1
grd2.Col = 1
grd2.Text = sc2 + sb2 + sa2 + st2 + sf2 + sd2 + sp2 + sw2 + so2
If (sc1 - sc2) > 0 Then
grd3.row = 1
grd3.Col = 0
grd3.Text = (sc1 - sc2)
grd3.Col = 1
grd3.Text = "0"
Else
grd3.row = 1
grd3.Col = 0
grd3.Text = "0"
grd3.Col = 1
grd3.Text = (sc2 - sc1)
End If
If (sa1 - sa2) > 0 Then
grd3.row = 3
grd3.Col = 0
grd3.Text = (sa1 - sa2)
grd3.Col = 1
grd3.Text = "0"
Else
grd3.row = 3
grd3.Col = 0
grd3.Text = "0"
grd3.Col = 1
grd3.Text = (sa2 - sa1)
End If
If (sb1 - sb2) > 0 Then
grd3.row = 2
grd3.Col = 0
grd3.Text = (sb1 - sb2)
grd3.Col = 1
grd3.Text = "0"
Else
grd3.row = 2
grd3.Col = 0
grd3.Text = "0"
grd3.Col = 1
grd3.Text = (sb2 - sb1)
End If
If (st1 - st2) > 0 Then
grd3.row = 4
grd3.Col = 0
grd3.Text = (st1 - st2)
grd3.Col = 1
grd3.Text = "0"
Else
grd3.row = 4
grd3.Col = 0
grd3.Text = "0"
grd3.Col = 1
grd3.Text = (st2 - st1)
End If
If (sf1 - sf2) > 0 Then
grd3.row = 5
grd3.Col = 0
grd3.Text = (sf1 - sf2)
grd3.Col = 1
grd3.Text = "0"
Else
grd3.row = 5
grd3.Col = 0
grd3.Text = "0"
grd3.Col = 1
grd3.Text = (sf2 - sf1)
End If
If (sd1 - sd2) > 0 Then
grd3.row = 6
grd3.Col = 0
grd3.Text = (sd1 - sd2)
grd3.Col = 1
grd3.Text = "0"
Else
grd3.row = 6
grd3.Col = 0
grd3.Text = "0"
grd3.Col = 1
grd3.Text = (sd2 - sd1)
End If
If (sp1 - sp2) > 0 Then
grd3.row = 7
grd3.Col = 0
grd3.Text = (sp1 - sp2)
grd3.Col = 1
grd3.Text = "0"
Else
grd3.row = 7
grd3.Col = 0
grd3.Text = "0"
grd3.Col = 1
grd3.Text = (sp2 - sp1)
End If
If (sw1 - sw2) > 0 Then
grd3.row = 8
grd3.Col = 0
grd3.Text = (sw1 - sw2)
grd3.Col = 1
grd3.Text = "0"
Else
grd3.row = 8
grd3.Col = 0
grd3.Text = "0"
grd3.Col = 1
grd3.Text = (sw2 - sw1)
End If
If (so1 - so2) > 0 Then
grd3.row = 9
grd3.Col = 0
grd3.Text = (so1 - so2)
grd3.Col = 1
grd3.Text = "0"
Else
grd3.row = 9
grd3.Col = 0
grd3.Text = "0"
grd3.Col = 1
grd3.Text = (so2 - so1)
End If
sc1 = 0
sc2 = 0
For i = 1 To 9
grd3.row = i
grd3.Col = 0
c1 = grd3.Text
grd3.Col = 1
c2 = grd3.Text
sc1 = sc1 + c1
sc2 = sc2 + c2
Next i
grd3.row = 10
grd3.Col = 0
grd3.Text = sc1
grd3.Col = 1
grd3.Text = sc2

End Sub
Private Sub chargegrd1_2()
On Error Resume Next
Dim i As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim e As Double
Dim es As Double
Dim s As Double
Dim ss As Double
Dim cre As Double
Dim c1 As Double
Dim sc1 As Double
Dim c2 As Double
Dim sc2 As Double
Dim b1 As Double
Dim sb1 As Double
Dim b2 As Double
Dim sb2 As Double
Dim a1 As Double
Dim sa1 As Double
Dim a2 As Double
Dim sa2 As Double
Dim t1 As Double
Dim st1 As Double
Dim t2 As Double
Dim st2 As Double
Dim f1 As Double
Dim sf1 As Double
Dim f2 As Double
Dim sf2 As Double
Dim d1 As Double
Dim sd1 As Double
Dim d2 As Double
Dim sd2 As Double
Dim p1 As Double
Dim sp1 As Double
Dim p2 As Double
Dim sp2 As Double
Dim w1 As Double
Dim sw1 As Double
Dim w2 As Double
Dim sw2 As Double
Dim o1 As Double
Dim so1 As Double
Dim o2 As Double
Dim so2 As Double
'***** grd1
grd1.Clear
grd1.Cols = 9
grd1.Rows = 1
grd1.ColWidth(0) = 1300
grd1.ColWidth(1) = 1300
grd1.ColWidth(2) = 800
grd1.ColWidth(3) = 1300
grd1.ColWidth(4) = 4000
grd1.ColWidth(5) = 800
grd1.ColWidth(6) = 1300
grd1.ColWidth(7) = 1300
grd1.ColWidth(8) = 1500
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 1
grd1.ColAlignment(6) = 1
grd1.ColAlignment(7) = 1
grd1.ColAlignment(8) = 1
grd1.row = 0
grd1.Col = 0
grd1.Text = "„œÌ‰"
grd1.Col = 1
grd1.Text = "œ«∆‰"
grd1.Col = 2
grd1.Text = "„‰/≈·Ï"
grd1.Col = 3
grd1.Text = "«·Õ”«»"
grd1.Col = 4
grd1.Text = "«·»Ì«‰"
grd1.Col = 5
grd1.Text = "—ﬁ„ «·ﬁÌœ"
grd1.Col = 6
grd1.Text = "«· «—ÌŒ"
grd1.Col = 7
grd1.Text = "«·”«⁄…"
grd1.Col = 8
grd1.Text = "«·„‰›–"
i = 1
es = 0
ss = 0
cre = 0
sc1 = 0
sc2 = 0
sb1 = 0
sb2 = 0
sa1 = 0
sa2 = 0
st1 = 0
st2 = 0
sf1 = 0
sf2 = 0
sd1 = 0
sd2 = 0
sp1 = 0
sp2 = 0
sw1 = 0
sw2 = 0
so1 = 0
so2 = 0
dat1 = DT23.Value
dat2 = DT24.Value
Call cont
grd1.Rows = jr.RecordCount + 5
Do While Not jr.EOF
s = 0
e = 0
dat3 = jr!dat
If dat3 >= dat1 And dat3 <= dat2 Then
grd1.row = i
grd1.Col = 0
grd1.Text = jr!cre
grd1.Col = 1
grd1.Text = jr!deb
grd1.Col = 2
grd1.Text = jr!dem
grd1.Col = 3
grd1.Text = jr!com
grd1.Col = 4
grd1.Text = jr!dec
grd1.Col = 5
grd1.Text = jr!ord
grd1.Col = 6
grd1.Text = jr!dat
grd1.Col = 7
grd1.Text = jr!heu
grd1.Col = 8
grd1.Text = jr!ger
e = jr!cre
es = es + e
s = jr!deb
ss = ss + s
i = i + 1
If jr!com = "—√” «·„«·" Then
c1 = jr!cre
sc1 = sc1 + c1
c2 = jr!deb
sc2 = sc2 + c2
End If
If jr!com = "«·»‰ﬂ" Then
b1 = jr!cre
sb1 = sb1 + b1
b2 = jr!deb
sb2 = sb2 + b2
End If
If jr!com = "«·’‰œÊﬁ" Then
a1 = jr!cre
sa1 = sa1 + a1
a2 = jr!deb
sa2 = sa2 + a2
End If
If jr!com = "«· ·«„Ì–" Then
t1 = jr!cre
st1 = st1 + t1
t2 = jr!deb
st2 = st2 + t2
End If
If jr!com = "«·√”« –…" Then
f1 = jr!cre
sf1 = sf1 + f1
f2 = jr!deb
sf2 = sf2 + f2
End If
If jr!com = "«·„’—Ê›« " Then
d1 = jr!cre
sd1 = sd1 + d1
d2 = jr!deb
sd2 = sd2 + d2
End If
If jr!com = "«·‘—ﬂ«¡" Then
p1 = jr!cre
sp1 = sp1 + p1
p2 = jr!deb
sp2 = sp2 + p2
End If
If jr!com = "«·⁄„«·" Then
o1 = jr!cre
so1 = so1 + o1
o2 = jr!deb
so2 = so2 + o2
End If
If jr!com = "«·√—»«Õ" Or jr!com = "«·Œ”«∆—" Then
w1 = jr!cre
sw1 = sw1 + w1
w2 = jr!deb
sw2 = sw2 + w2
End If
End If
jr.MoveNext
Loop
grd1.row = i
grd1.Col = 0
grd1.Text = es
grd1.Col = 1
grd1.Text = ss
grd1.Col = 4
grd1.Text = "«·„Ã„Ê⁄"
grd1.Rows = i + 1
End Sub

Private Sub grd1_Click()
On Error Resume Next
grd1.ToolTipText = grd1.Text

End Sub
Public Sub chargec1()
On Error Resume Next
Combo1.Clear
Combo1.AddItem "—√” «·„«·"
Combo1.AddItem "«·»‰ﬂ"
Combo1.AddItem "«·’‰œÊﬁ"
Combo1.AddItem "«· ·«„Ì–"
Combo1.AddItem "«·√”« –…"
Combo1.AddItem "«·⁄„«·"
Combo1.AddItem "«·„’—Ê›« "
Combo1.AddItem "«·‘—ﬂ«¡"
Combo1.AddItem "«·√—»«Õ"
End Sub
Private Sub chargegrd4_5()
On Error Resume Next
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim cre1 As Double
Dim deb1 As Double
Dim cres As Double
Dim debs As Double
Dim dif As Double
Dim i As Double
Dim j As Double
Dim ord1 As Double
Dim ord2 As Double
Dim controle1 As Double
Dim controle2 As Double
grd4.Clear
grd4.Cols = 4
grd4.Rows = 1
grd4.ColWidth(0) = 2000
grd4.ColWidth(1) = 2000
grd4.ColWidth(2) = 800
grd4.ColWidth(3) = 1300
grd4.ColAlignment(0) = 1
grd4.ColAlignment(1) = 1
grd4.ColAlignment(2) = 1
grd4.ColAlignment(3) = 1
grd4.row = 0
grd4.Col = 0
grd4.Text = "«·„»·€"
grd4.Col = 1
grd4.Text = "«·Õ”«»"
grd4.Col = 2
grd4.Text = "«·ﬁÌœ"
grd4.Col = 3
grd4.Text = "«· «—ÌŒ"
grd5.Clear
grd5.Cols = 4
grd5.Rows = 1
grd5.ColWidth(0) = 2000
grd5.ColWidth(1) = 2000
grd5.ColWidth(2) = 800
grd5.ColWidth(3) = 1300
grd5.ColAlignment(0) = 1
grd5.ColAlignment(1) = 1
grd5.ColAlignment(2) = 1
grd5.ColAlignment(3) = 1
grd5.row = 0
grd5.Col = 0
grd5.Text = "«·„»·€"
grd5.Col = 1
grd5.Text = "«·Õ”«»"
grd5.Col = 2
grd5.Text = "«·ﬁÌœ"
grd5.Col = 3
grd5.Text = "«· «—ÌŒ"
dat1 = DT1.Value
dat2 = DT2.Value
i = 1
j = 1
ord2 = "0"
cres = 0
debs = 0
Call cont
grd4.Rows = jr.RecordCount + 3
grd5.Rows = jr.RecordCount + 3
Do While Not jr.EOF
ord1 = jr!ord
dat3 = jr!dat
If dat3 >= dat1 And dat3 <= dat2 Then
'***** Medine
controle1 = jr!cre
If controle1 > 0 Then
If jr!com = Combo1.Text Then
If ord1 = ord2 Then
jr.MovePrevious
grd4.row = i
grd4.Col = 0
grd4.Text = jr!deb
deb1 = jr!deb
debs = debs + deb1
grd4.Col = 1
grd4.Text = jr!com
grd4.Col = 2
grd4.Text = jr!ord
grd4.Col = 3
grd4.Text = jr!dat
i = i + 1
jr.MoveNext
Else
jr.MoveNext
grd4.row = i
grd4.Col = 0
grd4.Text = jr!deb
deb1 = jr!deb
debs = debs + deb1
grd4.Col = 1
grd4.Text = jr!com
grd4.Col = 2
grd4.Text = jr!ord
grd4.Col = 3
grd4.Text = jr!dat
i = i + 1
jr.MovePrevious
End If
End If
End If
'**** end Medine
'***** Daeene
controle2 = jr!deb
If controle2 > 0 Then
If jr!com = Combo1.Text Then
If ord1 = ord2 Then
jr.MovePrevious
grd5.row = j
grd5.Col = 0
grd5.Text = jr!cre
cre1 = jr!cre
cres = cres + cre1
grd5.Col = 1
grd5.Text = jr!com
grd5.Col = 2
grd5.Text = jr!ord
grd5.Col = 3
grd5.Text = jr!dat
j = j + 1
jr.MoveNext
Else
jr.MoveNext
grd5.row = j
grd5.Col = 0
grd5.Text = jr!cre
cre1 = jr!cre
cres = cres + cre1
grd5.Col = 1
grd5.Text = jr!com
grd5.Col = 2
grd5.Text = jr!ord
grd5.Col = 3
grd5.Text = jr!dat
j = j + 1
jr.MovePrevious
End If
End If
End If
'**** end Daeene
End If
ord2 = jr!ord
jr.MoveNext
Loop
dif = cres - debs
If i > j Then
j = i
grd4.Rows = i + 3
grd5.Rows = i + 3
Else
i = j
grd4.Rows = j + 3
grd5.Rows = j + 3
End If
If dif > 0 Then
grd4.row = j
grd4.Col = 0
grd4.Text = dif
grd4.Col = 1
grd4.Text = "—’Ìœ „ÕÊ·"
grd4.row = j + 1
grd4.Col = 0
grd4.Text = dif + debs
grd4.Col = 1
grd4.Text = "«·„Ã„Ê⁄"
grd5.row = i + 1
grd5.Col = 0
grd5.Text = cres
grd5.Col = 1
grd5.Text = "«·„Ã„Ê⁄"
grd5.row = i + 2
grd5.Col = 0
grd5.Text = dif
grd5.Col = 1
grd5.Text = "—’Ìœ „‰ﬁÊ·"
End If
If dif < 0 Then
dif = debs - cres
grd5.row = i
grd5.Col = 0
grd5.Text = dif
grd5.Col = 1
grd5.Text = "—’Ìœ „ÕÊ·"
grd5.row = i + 1
grd5.Col = 0
grd5.Text = dif + cres
grd5.Col = 1
grd5.Text = "«·„Ã„Ê⁄"
grd4.row = j + 1
grd4.Col = 0
grd4.Text = debs
grd4.Col = 1
grd4.Text = "«·„Ã„Ê⁄"
grd4.row = i + 2
grd4.Col = 0
grd4.Text = dif
grd4.Col = 1
grd4.Text = "—’Ìœ „‰ﬁÊ·"
End If
Label25.Caption = "Õ”«» „› ÊÕ"
If dif = 0 Then
grd5.row = i + 1
grd5.Col = 0
grd5.Text = cres
grd5.Col = 1
grd5.Text = "«·„Ã„Ê⁄"
grd4.row = j + 1
grd4.Col = 0
grd4.Text = debs
grd4.Col = 1
grd4.Text = "«·„Ã„Ê⁄"
Label25.Caption = "Õ”«» „€·ﬁ"
End If

End Sub
Private Sub chargegrd6_7_8_9()
On Error Resume Next
'On Error Resume Next
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim i As Double
Dim j As Double
Dim k As Double
Dim l As Double
Dim a As Double
Dim b As Double
Dim c As Double
Dim t As Double
Dim st As Double
Dim d As Double
Dim sd As Double
Dim p As Double
Dim sp As Double
Dim s As Double
Dim ss As Double
grd6.Clear
grd6.Cols = 3
grd6.Rows = 1
grd6.ColWidth(0) = 2000
grd6.ColWidth(1) = 3200
grd6.ColWidth(2) = 1300
grd6.ColAlignment(0) = 1
grd6.ColAlignment(1) = 1
grd6.ColAlignment(2) = 1
grd6.row = 0
grd6.Col = 0
grd6.Text = "«·„»·€"
grd6.Col = 1
grd6.Text = "«·»Ì«‰"
grd6.Col = 2
grd6.Text = "«· «—ÌŒ"
grd7.Clear
grd7.Cols = 3
grd7.Rows = 1
grd7.ColWidth(0) = 2000
grd7.ColWidth(1) = 3200
grd7.ColWidth(2) = 1300
grd7.ColAlignment(0) = 1
grd7.ColAlignment(1) = 1
grd7.ColAlignment(2) = 1
grd7.row = 0
grd7.Col = 0
grd7.Text = "«·„»·€"
grd7.Col = 1
grd7.Text = "«·»Ì«‰"
grd7.Col = 2
grd7.Text = "«· «—ÌŒ"
grd8.Clear
grd8.Cols = 3
grd8.Rows = 1
grd8.ColWidth(0) = 2000
grd8.ColWidth(1) = 3200
grd8.ColWidth(2) = 1300
grd8.ColAlignment(0) = 1
grd8.ColAlignment(1) = 1
grd8.ColAlignment(2) = 1
grd8.row = 0
grd8.Col = 0
grd8.Text = "«·„»·€"
grd8.Col = 1
grd8.Text = "«·»Ì«‰"
grd8.Col = 2
grd8.Text = "«· «—ÌŒ"
grd9.Clear
grd9.Cols = 3
grd9.Rows = 1
grd9.ColWidth(0) = 3200
grd9.ColWidth(1) = 1400
grd9.ColWidth(2) = 1900
grd9.ColAlignment(0) = 1
grd9.ColAlignment(1) = 1
grd9.ColAlignment(2) = 1
grd9.row = 0
grd9.Col = 0
grd9.Text = "«·‘—Ìﬂ"
grd9.Col = 1
grd9.Text = "‰”»… «·‘—«ﬂ…%"
grd9.Col = 2
grd9.Text = "«·‰’Ì»"
i = 1
j = 1
k = 1
l = 1
st = 0
sd = 0
sp = 0
dat1 = DT1.Value
dat2 = DT2.Value
Call cont
grd6.Rows = rc.RecordCount + 3
grd7.Rows = ps.RecordCount + 3
grd8.Rows = dp.RecordCount + 3
grd9.Rows = pa.RecordCount + 3
Do While Not rc.EOF
dat3 = rc!dat
If dat3 >= dat1 And dat3 <= dat2 Then
grd6.row = i
grd6.Col = 0
grd6.Text = rc!mon
t = rc!mon
st = st + t
grd6.Col = 1
grd6.Text = rc!rec + " : " + rc!mois
grd6.Col = 2
grd6.Text = rc!dat
i = i + 1
End If
rc.MoveNext
Loop
'***** presence
If ps.RecordCount > 0 Then
ps.MoveFirst
End If
Do While Not ps.EOF
dat3 = ps!dat
If dat3 >= dat1 And dat3 <= dat2 Then
If ps!cas <> "p" Then
a = ps!tot
b = ps!prm
c = ps!rtr
d = a + b + c
sd = sd + d
grd7.row = j
grd7.Col = 0
grd7.Text = d
grd7.Col = 1
grd7.Text = ps!ser + " : " + ps!nom
grd7.Col = 2
grd7.Text = ps!dat
j = j + 1
End If
End If
ps.MoveNext
Loop
'***** depenses
If dp.RecordCount > 0 Then
dp.MoveFirst
End If
Do While Not dp.EOF
dat3 = dp!dat
If dat3 >= dat1 And dat3 <= dat2 Then
grd8.row = k
grd8.Col = 0
grd8.Text = dp!mon
p = dp!mon
sp = sp + p
grd8.Col = 1
grd8.Text = dp!dec
grd8.Col = 2
grd8.Text = dp!dat
k = k + 1
End If
dp.MoveNext
Loop
'***** partenaires
If pa.RecordCount > 0 Then
pa.MoveFirst
End If
Do While Not pa.EOF
grd9.row = l
grd9.Col = 0
grd9.Text = pa!nom
grd9.Col = 1
grd9.Text = pa!pou
grd9.Col = 2
grd9.Text = ""
l = l + 1
pa.MoveNext
Loop
grd6.Rows = i
grd7.Rows = j
grd8.Rows = k
grd9.Rows = l
'***** tesjil
ss = 0
s = 0
If ce.RecordCount > 0 Then
ce.MoveFirst
End If
Do While Not ce.EOF
dat3 = ce!dat
If dat3 >= dat1 And dat3 <= dat2 Then
If ce!moi = "«· ”ÃÌ·" Then
s = ce!pay
ss = ss + s
End If
End If
ce.MoveNext
Loop
Label1.Caption = st - ss
Label3.Caption = sd
Label4.Caption = (st - sd - ss)
a = Label2.Caption
b = (((st - sd - ss) * a) / 100) + ss
MyNumber = Round(b, 0)
b = MyNumber
Label5.Caption = b
Label6.Caption = sp
Label7.Caption = (b - sp)
b = ((((st - sd - ss) * a) / 100) - sp) + ss
For i = 1 To l - 1
grd9.row = i
grd9.Col = 1
a = grd9.Text
'b = Label7.Caption
c = (b * a) / 100
MyNumber = Round(c, 1)
c = MyNumber
grd9.row = i
grd9.Col = 2
grd9.Text = c
Next i
End Sub
Private Sub balance()
On Error Resume Next
'On Error Resume Next
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim ba1 As Double
Dim ba2 As Double
Dim cs1 As Double
Dim cs2 As Double
Dim pt1 As Double
Dim po1 As Double
Dim pt2 As Double
Dim po2 As Double
Dim ft1 As Double
Dim ft2 As Double
Dim cre As Double
Dim o1 As Double
Dim o2 As Double
Dim s As Double
Dim ct As Double
Dim pi As Double
Dim pps As Double
Dim dph As Double
Dim dpp As Double
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim f As Double
Dim g As Double
Dim h As Double
Dim i As Double
Dim l As Double
Dim o As Double
dat1 = DT1.Value
dat2 = DT2.Value
'**** calcule
'**** Bank et Capital
ba1 = 0
ba2 = 0
ct = 0
Call cont
Do While Not bn.EOF
dat3 = bn!dat
If dat3 >= dat1 And dat3 <= dat2 Then
s = 0
'*** positif
If bn!Mod = "«Ìœ«⁄" Or bn!Mod = "—√” «·„«·" Then
s = bn!mon
ba1 = ba1 + s
End If
'*** negatif
If bn!Mod = "”Õ»" Then
s = bn!mon
ba2 = ba2 + s
End If
'*** capital
If bn!Mod = "—√” «·„«·" Then
s = bn!mon
ct = ct + s
End If
End If
bn.MoveNext
Loop
'**** Caisse
s = 0
cs1 = 0
cs2 = 0
Call cont
Do While Not ca.EOF
dat3 = ca!dat
If dat3 >= dat1 And dat3 <= dat2 Then
'*** positif
If ca!cas = "œ«Œ·" Then
s = ca!mon
cs1 = cs1 + s
End If
'*** negatif
If ca!cas = "Œ«—Ã" Then
s = ca!mon
cs2 = cs2 + s
End If
End If
ca.MoveNext
Loop
'**** partenaires
s = 0
pt1 = 0
pt2 = 0
Call cont
Do While Not pp.EOF
dat3 = pp!dat
If dat3 >= dat1 And dat3 <= dat2 Then
'*** positif
If pp!Mod = "«·’‰œÊﬁ" Then
s = pp!mon
pt1 = pt1 + s
End If
'*** negatif
If pp!Mod = "«·‘—Ìﬂ" Then
s = pp!mon
pt2 = pt2 + s
End If
End If
pp.MoveNext
Loop
'**** fonctionnaires
s = 0
ft1 = 0
ft2 = 0
Call cont
Do While Not pfc.EOF
dat3 = pfc!dat
If dat3 >= dat1 And dat3 <= dat2 Then
'*** positif
If pfc!cas = "œ›⁄ „»·€" Then
s = pfc!mon
ft1 = ft1 + s
End If
'*** negatif
If pfc!cas = " ”œÌœ —« »" Then
s = pfc!mon
ft2 = ft2 + s
End If
End If
pfc.MoveNext
Loop
'**** Payprofesseurs
s = 0
pps = 0
Call cont
Do While Not pf.EOF
dat3 = pf!dat
If dat3 >= dat1 And dat3 <= dat2 Then
s = pf!mon
pps = pps + s
End If
pf.MoveNext
Loop
'**** duprofesseurs H M
dph = Label50.Caption
'**** duprofesseurs H M P
dps = Label5.Caption
'**** Affichage
'****Droite
'****Bank
Label8.Caption = (ba1 - ba2)
'****Caisse
Label9.Caption = (cs1 - cs2)
If (pt1 - pt2) >= 0 Then
Label10.Caption = (pt1 - pt2)
Label31(17).Caption = "«·‘—ﬂ«¡"
Label13.Caption = "0"
Label31(29).Caption = "********"
Else
Label13.Caption = ((pt1 - pt2) * -1)
Label31(29).Caption = "«·‘—ﬂ«¡"
Label10.Caption = "0"
Label31(17).Caption = "********"
End If
'****duprofesseurs
If ((dph + dps) - pps) >= 0 Then
Label12.Caption = ((dph + dps) - pps)
Label31(28).Caption = "«·√”« –…"
Label42.Caption = "0"
Label31(74).Caption = "********"
Else
Label42.Caption = ((dph + dps) - pps) * -1
Label31(74).Caption = "«·√”« –…"
Label12.Caption = "0"
Label31(28).Caption = "********"
End If
'****fonctionnaires
Label44.Caption = "0"
Label59.Caption = (ft1 - ft2)
'****Capital
Label15.Caption = ct
'****Profit
Label16.Caption = Label7.Caption
'*** addition
'*** el oussou
a = Label8.Caption
b = Label9.Caption
c = Label10.Caption
d = Label42.Caption
o = Label59.Caption
'*** toutal el  oussoul el moutedawila
Label11.Caption = (a + b + c + d + o)
Label19.Caption = (a + b + c + d + o)
'*** total Khoussoum
e = Label12.Caption
f = Label13.Caption
g = Label44.Caption
h = Label43.Caption
Label14.Caption = (e + f + g + h)
i = Label15.Caption
l = Label16.Caption
Label17.Caption = (i + l)
'*** total Khoussoum et total el houkouk el molkiya
Label18.Caption = (e + f + g + h) + (i + l)
' **** cas particulier
i = Label19.Caption
l = Label18.Caption
If (l - i) = 0 Then
Label43.Caption = "0"
Label31(75).Caption = "********"
Exit Sub
Else
Label31(75).Caption = "œ«∆‰Ê«  Ê“Ì⁄« "
If (l - i) < 0 Then
Label43.Caption = (i - l)
Else
Label43.Caption = (l - i) * -1
End If
'*** total Khoussoum
e = Label12.Caption
f = Label13.Caption
g = Label44.Caption
h = Label43.Caption
Label14.Caption = (e + f + g + h)
i = Label15.Caption
l = Label16.Caption
Label17.Caption = (i + l)
'*** total Khoussoum et total el houkouk el molkiya
Label18.Caption = (e + f + g + h) + (i + l)
End If
End Sub
Private Sub chargegrd10_11()
On Error Resume Next
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim i As Double
Dim j As Double
Dim k As Double
Dim l As Double
Dim h As Double
Dim mon1 As Double
Dim rec1 As Double
Dim mon2 As Double
Dim rec2 As Double
grd10.Clear
grd10.Cols = 3
grd10.Rows = 1
grd10.ColWidth(0) = 0
grd10.ColWidth(1) = 1300
grd10.ColWidth(2) = 1300
grd10.ColAlignment(0) = 1
grd10.ColAlignment(1) = 1
grd10.ColAlignment(2) = 1
grd10.row = 0
grd10.Col = 0
grd10.Text = ""
grd10.Col = 1
grd10.Text = "«·Ê’·"
grd10.Col = 2
grd10.Text = "«·„»·€"
grd11.Clear
grd11.Cols = 3
grd11.Rows = 1
grd11.ColWidth(0) = 0
grd11.ColWidth(1) = 1200
grd11.ColWidth(2) = 2000
grd11.ColAlignment(0) = 1
grd11.ColAlignment(1) = 1
grd11.ColAlignment(2) = 1
grd11.row = 0
grd11.Col = 0
grd11.Text = ""
grd11.Col = 1
grd11.Text = "«·Ê’·"
grd11.Col = 2
grd11.Text = "«·„·«ÕŸ…"
i = 1
j = 1
k = 1
dat1 = DT21.Value
dat2 = DT22.Value
Call cont
grd10.Rows = rc.RecordCount + 3
Do While Not rc.EOF
dat3 = rc!dat
If dat3 >= dat1 And dat3 <= dat2 Then
grd10.row = i
grd10.Col = 1
grd10.Text = rc!rec
grd10.Col = 2
grd10.Text = rc!mon
i = i + 1
End If
rc.MoveNext
Loop
grd10.Rows = i
If cr.RecordCount > 0 Then
cr.MoveFirst
End If
grd11.Rows = cr.RecordCount + 3
Do While Not cr.EOF
dat3 = cr!dat
If dat3 >= dat1 And dat3 <= dat2 Then
grd11.row = j
grd11.Col = 0
grd11.Text = cr!aut
grd11.Col = 1
grd11.Text = cr!rec
grd11.Col = 2
grd11.Text = cr!act
j = j + 1
End If
cr.MoveNext
Loop
grd11.Rows = j
grd10.Col = 1
grd10.Sort = 1
grd11.Col = 1
grd11.Sort = 1
End Sub

Private Sub grd11_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim tx0 As String
Dim tx1 As String
Dim tx2 As String
i = grd11.row
j = grd11.Col
'***** recu
If j = 1 Then
grd11.row = i
grd11.Col = 0
tx0 = grd11.Text
grd11.Col = 1
tx1 = grd11.Text
grd11.Col = 2
tx2 = grd11.Text
g = InputBox("«œŒ· «·Ê’·", tx2, tx1)
If g = Cancel Then
Exit Sub
End If
If Val(g) <= 0 Then
Exit Sub
End If
Call cont
Do While Not cr.EOF
If Text1.Text = cr!rec And cr!rec <> tx0 Then
MsgBox "—ﬁ„ «·Ê’· «·„œŒ·  „ ÕÃ“Â ”«»ﬁ«", vbCritical
Exit Sub
End If
cr.MoveNext
Loop
grd11.row = i
grd11.Col = 1
grd11.Text = g
'****** controlerecu
Call cont
Do While Not cr.EOF
If cr!aut = tx0 Then
cr!rec = g
cr!dat = Date
cr!act = "1"
cr.Update
cr.MoveLast
End If
cr.MoveNext
Loop
Text2.Text = ""
Text2.SetFocus
grd12.Clear
grd12.Rows = 1
grd12.row = 0
grd12.Col = 1
grd12.Text = "«·Ê’·"
grd12.Col = 2
grd12.Text = "«·Œÿ√"
Call controleurrecus
End If
'**** montants
If j = 2 Then
grd11.row = i
grd11.Col = 0
tx0 = grd11.Text
grd11.Col = 1
tx1 = grd11.Text
grd11.Col = 2
tx2 = grd11.Text
g = InputBox("«œŒ· «·„»·€ «·„œ›Ê⁄", tx1, tx2)
If g = Cancel Then
Exit Sub
End If
If Val(g) <= 0 Then
Exit Sub
End If
grd11.row = i
grd11.Col = 2
grd11.Text = g
'****** controlerecu
Call cont
Do While Not cr.EOF
If cr!aut = tx0 Then
cr!mon = g
cr!dat = Date
cr!act = "1"
cr.Update
cr.MoveLast
End If
cr.MoveNext
Loop
Text2.Text = ""
Text2.SetFocus
grd12.Clear
grd12.Rows = 1
grd12.row = 0
grd12.Col = 1
grd12.Text = "«·Ê’·"
grd12.Col = 2
grd12.Text = "«·Œÿ√"
Call controleurrecus
End If
End Sub

Private Sub Text1_Change()
On Error Resume Next
Label30.Caption = ""
Label53.Caption = ""
Label54.Caption = ""
Label55.Caption = ""
Label56.Caption = ""
Label57.Caption = ""
Label58.Caption = ""

End Sub

Private Sub Text1_Click()
On Error Resume Next
Text1_Change
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii <> 8 Then
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End If

End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Dim i As Double
If Text2.Text <> "" Then
If KeyCode = 13 Then
Text1.Text = Trim(Text1.Text)
Text1.Text = Val(Text1.Text)
If Text1.Text = "" Or Val(Text1.Text) <= 0 Then
MsgBox "Ì—ÃÏ «œŒ«· «·Ê’· »‘ﬂ· ’ÕÌÕ", vbCritical
Exit Sub
End If
Call cont
Do While Not cr.EOF
If Text1.Text = cr!rec Then
MsgBox "—ﬁ„ «·Ê’· «·„œŒ·  „ ÕÃ“Â ”«»ﬁ«", vbCritical
Exit Sub
End If
cr.MoveNext
Loop
'****** controlerecu
cr.AddNew
cr!rec = Text1.Text
cr!mon = Text2.Text
cr!dat = Date
cr!act = "1"
cr.Update
i = grd11.Rows
grd11.Rows = grd11.Rows + 1
grd11.row = i
grd11.Col = 1
grd11.Text = Text1.Text
grd11.Col = 2
grd11.Text = Text2.Text
Text1.Text = Val(Text1.Text) + 1
Text2.Text = ""
Text2.SetFocus
grd12.Clear
grd12.Rows = 1
grd12.row = 0
grd12.Col = 1
grd12.Text = "«·Ê’·"
grd12.Col = 2
grd12.Text = "«·Œÿ√"
Call controleurrecus
End If
End If

End Sub
Private Sub controleurrecus()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim k As Double
Dim l As Double
Dim p As Double
j = grd11.Rows
i = grd10.Rows
'**** grd11.rows >= grd10.rows
If j >= i Then
For k = 1 To j - 1
grd11.row = k
grd11.Col = 1
rec1 = grd11.Text
grd11.Col = 2
mon1 = grd11.Text
For l = 1 To i - 1
grd10.row = l
grd10.Col = 1
rec2 = grd10.Text
grd10.Col = 2
mon2 = grd10.Text
p = 0
'*** rec1 = rec2
If rec1 = rec2 Then
'*** mon1 = mon2
If mon1 = mon2 Then
l = i
p = 1
Else
'*** mon1 =! mon2
h = grd12.Rows
grd12.Rows = grd12.Rows + 1
grd12.row = h
grd12.Col = 1
grd12.Text = rec1
grd12.Col = 2
grd12.Text = "ÌÊÃœ «Œ ·«› ›Ì «·„»·€ »Ì‰ Õ”«» «· ·«„Ì– Ê—ﬂ‰ «· Õﬁﬁ"
l = i
p = 1
End If
Else
If p = 0 Then
p = 2
End If
End If
Next l
If p = 2 Then
'*** mon1 =! mon2
h = grd12.Rows
grd12.Rows = grd12.Rows + 1
grd12.row = h
grd12.Col = 1
grd12.Text = rec1
grd12.Col = 2
grd12.Text = "ÌÊÃœ ›Ì —ﬂ‰ «· Õﬁﬁ Ê·« ÌÊÃœ ›Ì —ﬂ‰ Õ”«» «· ·«„Ì–"
l = i
End If
Next k
'**** grd11.rows < grd10.rows
Else
For k = 1 To i - 1
grd10.row = k
grd10.Col = 1
rec1 = grd10.Text
grd10.Col = 2
mon1 = grd10.Text
For l = 1 To j - 1
grd11.row = l
grd11.Col = 1
rec2 = grd11.Text
grd11.Col = 2
mon2 = grd11.Text
p = 0
'*** rec1 = rec2
If rec1 = rec2 Then
'*** mon1 = mon2
If mon1 = mon2 Then
l = j
p = 1
Else
'*** mon1 =! mon2
h = grd12.Rows
grd12.Rows = grd12.Rows + 1
grd12.row = h
grd12.Col = 1
grd12.Text = rec1
grd12.Col = 2
grd12.Text = "ÌÊÃœ «Œ ·«› ›Ì «·„»·€ »Ì‰ Õ”«» «· ·«„Ì– Ê—ﬂ‰ «· Õﬁﬁ"
l = j
p = 1
End If
Else
If p = 0 Then
p = 2
End If
End If
Next l
If p = 2 Then
'*** mon1 =! mon2
h = grd12.Rows
grd12.Rows = grd12.Rows + 1
grd12.row = h
grd12.Col = 1
grd12.Text = rec1
grd12.Col = 2
grd12.Text = "ÌÊÃœ ›Ì —ﬂ‰ Õ”«» «· ·«„Ì– Ê·« ÌÊÃœ ›Ì —ﬂ‰ «· Õﬁﬁ"
l = j
End If
Next k
End If
h = grd12.Rows
If h = 1 Then
grd12.Rows = grd12.Rows + 1
grd12.row = h
grd12.Col = 1
grd12.Text = "”·Ì„"
grd12.Col = 2
grd12.Text = "·« ÊÃœ √Ì… √Œÿ«¡"
End If

End Sub
Private Sub chargegrd13()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim l As Double
Dim k As Double
Dim n As Double
Dim p As Double
Dim g As Double
Dim num1 As String
Dim num2 As String
Dim a As Double
Dim b As Double
Dim c As Double
Dim m As Double
Dim sm As Double
grd13.Clear
grd13.Cols = 21
grd13.Rows = 1
grd13.ColWidth(0) = 1000
grd13.ColWidth(1) = 4000
If Check2.Value = 1 Then
grd13.ColWidth(2) = 1000
grd13.ColWidth(3) = 1000
grd13.ColWidth(4) = 1000
grd13.ColWidth(5) = 1000
grd13.ColWidth(6) = 1000
grd13.ColWidth(7) = 1000
grd13.ColWidth(8) = 1000
grd13.ColWidth(9) = 1000
grd13.ColWidth(10) = 1000
grd13.ColWidth(11) = 1000
grd13.ColWidth(12) = 1000
grd13.ColWidth(13) = 1000
grd13.ColWidth(14) = 1000
Else
grd13.ColWidth(2) = 0
grd13.ColWidth(3) = 0
grd13.ColWidth(4) = 0
grd13.ColWidth(5) = 0
grd13.ColWidth(6) = 0
grd13.ColWidth(7) = 0
grd13.ColWidth(8) = 0
grd13.ColWidth(9) = 0
grd13.ColWidth(10) = 0
grd13.ColWidth(11) = 0
grd13.ColWidth(12) = 0
grd13.ColWidth(13) = 0
grd13.ColWidth(14) = 0
End If
grd13.ColWidth(15) = 1500
grd13.ColWidth(16) = 800
grd13.ColWidth(17) = 1500
grd13.ColWidth(18) = 1500
grd13.ColWidth(19) = 1500
grd13.ColWidth(20) = 1500
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
grd13.ColAlignment(11) = 1
grd13.ColAlignment(12) = 1
grd13.ColAlignment(13) = 1
grd13.ColAlignment(14) = 1
grd13.ColAlignment(15) = 1
grd13.ColAlignment(16) = 1
grd13.ColAlignment(17) = 1
grd13.ColAlignment(18) = 1
grd13.ColAlignment(19) = 1
grd13.ColAlignment(20) = 1
grd13.row = 0
grd13.Col = 0
grd13.Text = "«·—ﬁ„"
grd13.Col = 1
grd13.Text = "«·≈”„"
grd13.Col = 2
grd13.Text = "«· ”ÃÌ·"
grd13.Col = 3
grd13.Text = "«ﬂ Ê»—"
grd13.Col = 4
grd13.Text = "‰Ê›„»—"
grd13.Col = 5
grd13.Text = "œÌ”„»—"
grd13.Col = 6
grd13.Text = "Ì‰«Ì—"
grd13.Col = 7
grd13.Text = "›»—«Ì—"
grd13.Col = 8
grd13.Text = "„«—”"
grd13.Col = 9
grd13.Text = "«»—Ì·"
grd13.Col = 10
grd13.Text = "„«ÌÊ"
grd13.Col = 11
grd13.Text = "ÌÊ‰ÌÊ"
grd13.Col = 12
grd13.Text = "ÌÊ·ÌÊ"
grd13.Col = 13
grd13.Text = "√€”ÿ”"
grd13.Col = 14
grd13.Text = "”» „»—"
grd13.Col = 15
grd13.Text = "—.«·‘Â—Ì…"
grd13.Col = 16
grd13.Text = "⁄.«·√‘Â—"
grd13.Col = 17
grd13.Text = "«·„” Õﬁ"
grd13.Col = 18
grd13.Text = "«·„œ›Ê⁄"
grd13.Col = 19
grd13.Text = "„ÿ«·» »‹"
grd13.Col = 20
grd13.Text = "«·Õ«·…"
i = 1
Call cont
grd13.Rows = et.RecordCount + 3
Do While Not et.EOF
If Combo2.Text = et!cla Then
grd13.row = i
grd13.Col = 0
grd13.Text = et!num
grd13.Col = 1
grd13.Text = et!nom
grd13.Col = 2
grd13.Text = "0"
grd13.Col = 3
grd13.Text = "0"
grd13.Col = 4
grd13.Text = "0"
grd13.Col = 5
grd13.Text = "0"
grd13.Col = 6
grd13.Text = "0"
grd13.Col = 7
grd13.Text = "0"
grd13.Col = 8
grd13.Text = "0"
grd13.Col = 9
grd13.Text = "0"
grd13.Col = 10
grd13.Text = "0"
grd13.Col = 11
grd13.Text = "0"
grd13.Col = 12
grd13.Text = "0"
grd13.Col = 13
grd13.Text = "0"
grd13.Col = 14
grd13.Text = "0"
grd13.Col = 15
grd13.Text = Text5.Text
grd13.Col = 16
grd13.Text = Combo4.Text
a = Text5.Text
b = Combo4.Text
c = a * b
grd13.Col = 17
grd13.Text = c
grd13.Col = 18
grd13.Text = "0"
grd13.Col = 19
grd13.Text = c
grd13.Col = 20
grd13.Text = "·„ Ìœ›⁄ √Ì ‘Ì¡ »⁄œ"
i = i + 1
End If
et.MoveNext
Loop
grd13.Rows = i
For i = 1 To (grd13.Rows) - 1
sm = 0
k = 0
n = 0
p = 0
grd13.row = i
grd13.Col = 0
num1 = grd13.Text
If ce.RecordCount > 0 Then
ce.MoveFirst
End If
Do While Not ce.EOF
num2 = ce!num
If num1 = num2 And Combo2.Text = ce!cla Then
If ce!cas = "Õ«·… ≈ﬂ„«·" Then
k = k + 1
Else
n = n + 1
End If
'*** 0
If ce!moi = "«· ”ÃÌ·" Then
grd13.row = i
grd13.Col = 2
grd13.Text = ce!pay
m = ce!pay
sm = sm + m
p = 1
g = ce!pay
End If
'*** 10
If ce!moi = "«ﬂ Ê»—" Then
grd13.row = i
grd13.Col = 3
grd13.Text = ce!pay
m = ce!pay
sm = sm + m
End If
'*** 11
If ce!moi = "‰Ê›„»—" Then
grd13.row = i
grd13.Col = 4
grd13.Text = ce!pay
m = ce!pay
sm = sm + m
End If
'*** 12
If ce!moi = "œÌ”„»—" Then
grd13.row = i
grd13.Col = 5
grd13.Text = ce!pay
m = ce!pay
sm = sm + m
End If
'*** 1
If ce!moi = "Ì‰«Ì—" Then
grd13.row = i
grd13.Col = 6
grd13.Text = ce!pay
m = ce!pay
sm = sm + m
End If
'*** 2
If ce!moi = "›»—«Ì—" Then
grd13.row = i
grd13.Col = 7
grd13.Text = ce!pay
m = ce!pay
sm = sm + m
End If
'*** 3
If ce!moi = "„«—”" Then
grd13.row = i
grd13.Col = 8
grd13.Text = ce!pay
m = ce!pay
sm = sm + m
End If
'*** 4
If ce!moi = "«»—Ì·" Then
grd13.row = i
grd13.Col = 9
grd13.Text = ce!pay
m = ce!pay
sm = sm + m
End If
'*** 5
If ce!moi = "„«ÌÊ" Then
grd13.row = i
grd13.Col = 10
grd13.Text = ce!pay
m = ce!pay
sm = sm + m
End If
'*** 6
If ce!moi = "ÌÊ‰ÌÊ" Then
grd13.row = i
grd13.Col = 11
grd13.Text = ce!pay
m = ce!pay
sm = sm + m
End If
'*** 7
If ce!moi = "ÌÊ·ÌÊ" Then
grd13.row = i
grd13.Col = 12
grd13.Text = ce!pay
m = ce!pay
sm = sm + m
End If
'*** 8
If ce!moi = "√€”ÿ”" Then
grd13.row = i
grd13.Col = 13
grd13.Text = ce!pay
m = ce!pay
sm = sm + m
End If
'*** 9
If ce!moi = "”» „»—" Then
grd13.row = i
grd13.Col = 14
grd13.Text = ce!pay
m = ce!pay
sm = sm + m
End If
'*** mansuelle
grd13.row = i
grd13.Col = 15
grd13.Text = Text5.Text
'*** nbr mois
If n = 0 Then
b = Combo4.Text
Else
b = k - p
End If
grd13.row = i
grd13.Col = 16
grd13.Text = b
'*** dus
a = Text5.Text
c = a * b
grd13.row = i
grd13.Col = 17
grd13.Text = c + g
'*** pay
grd13.row = i
grd13.Col = 18
grd13.Text = sm
'*** res
grd13.row = i
grd13.Col = 19
grd13.Text = ((c + g) - sm)
'*** cas
grd13.row = i
grd13.Col = 20
grd13.Text = ce!cas

'**** end if num1=num2
End If
ce.MoveNext
Loop
Next i
grd13.Col = 20
grd13.Sort = 1
j = grd13.Rows
grd13.Rows = j + 1
For l = 2 To 19
a = 0
b = 0
If l <> 16 And l <> 15 Then
For i = 1 To j - 1
grd13.Col = l
grd13.row = i
a = grd13.Text
b = b + a
Next i
grd13.row = i
grd13.Col = l
grd13.Text = b
End If
Next l
grd13.row = j
grd13.Col = 0
grd13.Text = "«·„Ã„Ê⁄"
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If Text1.Text <> "" Then
If KeyCode = 13 Then
Command13_Click
End If
End If
End Sub

Private Sub Text5_Change()
On Error Resume Next
grd13.Clear
grd13.Rows = 1
End Sub

Private Sub Text5_Click()
On Error Resume Next
Text5_Change
End Sub
Public Sub chargec8()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim k As Double
Combo8.Clear
Label41.Caption = "2030-2031"
Call cont
Do While Not an.EOF
Combo8.AddItem an!ann
If an!ann = Label26.Caption Then
Label40.Caption = an!an2
End If
If an!an1 = Label40.Caption Then
Label41.Caption = an!ann
'Combo8.Text = an!ann
End If
an.MoveNext
Loop
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
Private Sub profits_14()
On Error Resume Next
'On Error Resume Next
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
Dim g As Double
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
Dim t As Integer
Dim ss As Double
grd14.Rows = 1
grd14.Cols = 8
grd14.ColWidth(0) = 0
grd14.ColWidth(1) = 1200
grd14.ColWidth(2) = 1900
grd14.ColWidth(3) = 1900
grd14.ColWidth(4) = 1900
grd14.ColWidth(5) = 1200
grd14.ColWidth(6) = 1900
grd14.ColWidth(7) = 1900
grd14.ColAlignment(0) = 1
grd14.ColAlignment(1) = 1
grd14.ColAlignment(2) = 1
grd14.ColAlignment(3) = 1
grd14.ColAlignment(4) = 1
grd14.ColAlignment(5) = 1
grd14.ColAlignment(6) = 1
grd14.ColAlignment(7) = 1
grd14.row = 0
grd14.Col = 0
grd14.Text = "«·‘Â—"
grd14.Col = 1
grd14.Text = "«·ﬁ”„"
grd14.Col = 2
grd14.Text = "«· ·«„Ì–"
grd14.Col = 3
grd14.Text = "«·√”« –…"
grd14.Col = 4
grd14.Text = "«·»«ﬁÌ"
grd14.Col = 5
grd14.Text = "% «·„ƒ””…"
grd14.Col = 6
grd14.Text = "«·„ƒ””…"
grd14.Col = 7
grd14.Text = "√”« –… «·‰”»…"
i = 1
cla1 = ""
mois1 = ""
sh = 0
sm = 0
sp = 0
dat1 = DT1.Value
dat2 = DT2.Value
'**** charge mois, classes et nbr
Call cont
grd14.Rows = cl.RecordCount + 5
Do While Not cl.EOF
t = 0
cla1 = cl!cla
If ps.RecordCount > 0 Then
ps.MoveFirst
End If
Do While Not ps.EOF
'**** grd14
If ps!cla = cla1 And ps!cas = "p" Then
t = 1
grd14.row = i
grd14.Col = 0
grd14.Text = ps!mois
grd14.Col = 1
grd14.Text = ps!cla
grd14.Col = 2
grd14.Text = "0"
grd14.Col = 3
grd14.Text = "0"
grd14.Col = 4
grd14.Text = "0"
grd14.Col = 5
grd14.Text = Label2.Caption
grd14.Col = 6
grd14.Text = "0"
i = i + 1
ps.MoveLast
End If
ps.MoveNext
Loop
If t = 0 Then
grd14.row = i
grd14.Col = 0
grd14.Text = "0"
grd14.Col = 1
grd14.Text = cl!cla
grd14.Col = 2
grd14.Text = "0"
grd14.Col = 3
grd14.Text = "0"
grd14.Col = 4
grd14.Text = "0"
grd14.Col = 5
grd14.Text = "100"
grd14.Col = 6
grd14.Text = "0"
i = i + 1
End If
cl.MoveNext
Loop
grd14.Rows = i
'Exit Sub

'**** professeurs pourcentage
n = grd14.Rows
If n > 1 Then
'**** charge montants etudiants
c = 0
Call cont
Do While Not ce.EOF
For i = 1 To n - 1
grd14.row = i
grd14.Col = 0
mois1 = grd14.Text
grd14.Col = 1
cla1 = grd14.Text
grd14.Col = 2
a = grd14.Text
If ce!cla = cla1 And ce!moi <> "«· ”ÃÌ·" Then
b = ce!pay
c = a + b
grd14.row = i
grd14.Col = 2
grd14.Text = c
End If
Next i
ce.MoveNext
Loop
'Exit Sub
'**** charge montants professeurs secondaires
c = 0
r = 0
f = 0
Call cont
Do While Not ps.EOF
For i = 1 To n - 1
grd14.row = i
grd14.Col = 0
mois1 = grd14.Text
grd14.Col = 1
cla1 = grd14.Text
grd14.Col = 3
a = grd14.Text
If ps!cla = cla1 And ps!cas <> "p" Then
e = ps!tot
f = a + e
e = ps!prm
f = f + e
e = ps!rtr
f = f + e
grd14.row = i
grd14.Col = 3
grd14.Text = f
End If
Next i
ps.MoveNext
Loop
'Exit Sub
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
grd14.row = i
grd14.Col = 2
a = grd14.Text
grd14.Col = 3
b = grd14.Text
grd14.Col = 5
Label48.Caption = grd14.Text
c = a - b
d = Label48.Caption
e = (c * d) / 100
MyNumber = Round(e, 0)
e = MyNumber
f = c - e
grd14.row = i
grd14.Col = 4
grd14.Text = c
grd14.Col = 6
grd14.Text = e
grd14.Col = 7
grd14.Text = f
Next i
'grd14.Rows = 70
End If
ss = 0
s = 0
Call cont
Do While Not ce.EOF
dat3 = ce!dat
If dat3 >= dat1 And dat3 <= dat2 Then
If ce!moi = "«· ”ÃÌ·" Then
s = ce!pay
ss = ss + s
End If
End If
ce.MoveNext
Loop
n = grd14.Rows
a = 0
b = 0
c = 0
d = 0
e = 0
f = 0
g = 0
h = 0
r = 0
p = 0
For i = 1 To n - 1
grd14.row = i
grd14.Col = 2
a = grd14.Text
b = b + a
grd14.Col = 3
c = grd14.Text
d = d + c
grd14.Col = 4
e = grd14.Text
f = f + e
grd14.Col = 6
g = grd14.Text
h = h + g
grd14.Col = 7
r = grd14.Text
p = p + r
Next i
grd14.Rows = n + 5
grd14.row = n
grd14.Col = 1
grd14.Text = ""
grd14.Col = 2
grd14.Text = "----------"
grd14.Col = 3
grd14.Text = "----------"
grd14.Col = 4
grd14.Text = "----------"
grd14.Col = 6
grd14.Text = "----------"
grd14.Col = 7
grd14.Text = "----------"
grd14.row = n + 1
grd14.Col = 1
grd14.Text = "«·„Ã„Ê⁄"
grd14.Col = 2
grd14.Text = b
grd14.Col = 3
grd14.Text = d
grd14.Col = 4
grd14.Text = f
grd14.Col = 6
grd14.Text = h
grd14.Col = 7
grd14.Text = p
grd14.row = n + 2
grd14.Col = 5
grd14.Text = "«·—”Ê„"
grd14.Col = 6
grd14.Text = ss
grd14.row = n + 3
grd14.Col = 6
grd14.Text = "----------"
grd14.row = n + 4
grd14.Col = 5
grd14.Text = "«·„ƒ””…"
grd14.Col = 6
grd14.Text = (h + ss)
'grd14.Rows = 50
a = 0
sp = 0
'***** depenses
If dp.RecordCount > 0 Then
dp.MoveFirst
End If
Do While Not dp.EOF
dat3 = dp!dat
If dat3 >= dat1 And dat3 <= dat2 Then
a = dp!mon
sp = sp + a
End If
dp.MoveNext
Loop
Label1.Caption = (b + ss)
Label49.Caption = ss
Label3.Caption = ss
Label4.Caption = b
Label50.Caption = d
Label5.Caption = p
Label6.Caption = (h + ss)
Label51.Caption = sp
a = 0
a = (h + ss)
b = (a - sp)
Label7.Caption = b
End Sub
Private Sub chargegrd15_MSChart4()
On Error Resume Next
'On Error Resume Next
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
Dim g As Double
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
Dim t As Integer
Dim ss As Double
grd15.Rows = 1
grd15.Cols = 8
grd15.ColWidth(0) = 0
grd15.ColWidth(1) = 800
grd15.ColWidth(2) = 1200
grd15.ColWidth(3) = 1200
grd15.ColWidth(4) = 1200
grd15.ColWidth(5) = 1100
grd15.ColWidth(6) = 1200
grd15.ColWidth(7) = 1200
grd15.ColAlignment(0) = 1
grd15.ColAlignment(1) = 1
grd15.ColAlignment(2) = 1
grd15.ColAlignment(3) = 1
grd15.ColAlignment(4) = 1
grd15.ColAlignment(5) = 1
grd15.ColAlignment(6) = 1
grd15.ColAlignment(7) = 1
grd15.row = 0
grd15.Col = 0
grd15.Text = "«·‘Â—"
grd15.Col = 1
grd15.Text = "«·ﬁ”„"
grd15.Col = 2
grd15.Text = "«· ·«„Ì–"
grd15.Col = 3
grd15.Text = "«·√”« –…"
grd15.Col = 4
grd15.Text = "«·»«ﬁÌ"
grd15.Col = 5
grd15.Text = "% «·„ƒ””…"
grd15.Col = 6
grd15.Text = "«·„ƒ””…"
grd15.Col = 7
grd15.Text = "√”« –… «·‰”»…"
i = 1
cla1 = ""
mois1 = ""
sh = 0
sm = 0
sp = 0
dat1 = DT1.Value
dat2 = DT2.Value
'**** charge mois, classes et nbr
Call cont
grd15.Rows = cl.RecordCount + 5
Do While Not cl.EOF
t = 0
cla1 = cl!cla
If ps.RecordCount > 0 Then
ps.MoveFirst
End If
Do While Not ps.EOF
'**** grd15
If ps!cla = cla1 And ps!cas = "p" Then
t = 1
grd15.row = i
grd15.Col = 0
grd15.Text = ps!mois
grd15.Col = 1
grd15.Text = ps!cla
grd15.Col = 2
grd15.Text = "0"
grd15.Col = 3
grd15.Text = "0"
grd15.Col = 4
grd15.Text = "0"
grd15.Col = 5
grd15.Text = Label2.Caption
grd15.Col = 6
grd15.Text = "0"
i = i + 1
ps.MoveLast
End If
ps.MoveNext
Loop
If t = 0 Then
grd15.row = i
grd15.Col = 0
grd15.Text = "0"
grd15.Col = 1
grd15.Text = cl!cla
grd15.Col = 2
grd15.Text = "0"
grd15.Col = 3
grd15.Text = "0"
grd15.Col = 4
grd15.Text = "0"
grd15.Col = 5
grd15.Text = "100"
grd15.Col = 6
grd15.Text = "0"
i = i + 1
End If
cl.MoveNext
Loop
grd15.Rows = i
'Exit Sub

'**** professeurs pourcentage
n = grd15.Rows
If n > 1 Then
'**** charge montants etudiants
c = 0
Call cont
Do While Not ce.EOF
For i = 1 To n - 1
grd15.row = i
grd15.Col = 0
mois1 = grd15.Text
grd15.Col = 1
cla1 = grd15.Text
grd15.Col = 2
a = grd15.Text
If ce!cla = cla1 And ce!moi <> "«· ”ÃÌ·" Then
b = ce!pay
c = a + b
grd15.row = i
grd15.Col = 2
grd15.Text = c
End If
Next i
ce.MoveNext
Loop
'Exit Sub
'**** charge montants professeurs secondaires
c = 0
r = 0
f = 0
Call cont
Do While Not ps.EOF
For i = 1 To n - 1
grd15.row = i
grd15.Col = 0
mois1 = grd15.Text
grd15.Col = 1
cla1 = grd15.Text
grd15.Col = 3
a = grd15.Text
If ps!cla = cla1 And ps!cas <> "p" Then
e = ps!tot
f = a + e
e = ps!prm
f = f + e
e = ps!rtr
f = f + e
grd15.row = i
grd15.Col = 3
grd15.Text = f
End If
Next i
ps.MoveNext
Loop
'Exit Sub
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
grd15.row = i
grd15.Col = 2
a = grd15.Text
grd15.Col = 3
b = grd15.Text
grd15.Col = 5
Label48.Caption = grd15.Text
c = a - b
d = Label48.Caption
e = (c * d) / 100
MyNumber = Round(e, 0)
e = MyNumber
f = c - e
grd15.row = i
grd15.Col = 4
grd15.Text = c
grd15.Col = 6
grd15.Text = e
grd15.Col = 7
grd15.Text = f
Next i
'grd15.Rows = 70
End If
ss = 0
s = 0
Call cont
Do While Not ce.EOF
dat3 = ce!dat
If dat3 >= dat1 And dat3 <= dat2 Then
If ce!moi = "«· ”ÃÌ·" Then
s = ce!pay
ss = ss + s
End If
End If
ce.MoveNext
Loop
n = grd15.Rows
a = 0
b = 0
c = 0
d = 0
e = 0
f = 0
g = 0
h = 0
r = 0
p = 0
For i = 1 To n - 1
grd15.row = i
grd15.Col = 2
a = grd15.Text
b = b + a
grd15.Col = 3
c = grd15.Text
d = d + c
grd15.Col = 4
e = grd15.Text
f = f + e
grd15.Col = 6
g = grd15.Text
h = h + g
grd15.Col = 7
r = grd15.Text
p = p + r
Next i
grd15.Rows = n + 5
grd15.row = n
grd15.Col = 1
grd15.Text = ""
grd15.Col = 2
grd15.Text = "----------"
grd15.Col = 3
grd15.Text = "----------"
grd15.Col = 4
grd15.Text = "----------"
grd15.Col = 6
grd15.Text = "----------"
grd15.Col = 7
grd15.Text = "----------"
grd15.row = n + 1
grd15.Col = 1
grd15.Text = "«·„Ã„Ê⁄"
grd15.Col = 2
grd15.Text = b
grd15.Col = 3
grd15.Text = d
grd15.Col = 4
grd15.Text = f
grd15.Col = 6
grd15.Text = h
grd15.Col = 7
grd15.Text = p
grd15.row = n + 2
grd15.Col = 5
grd15.Text = "«·—”Ê„"
grd15.Col = 6
grd15.Text = ss
grd15.row = n + 3
grd15.Col = 6
grd15.Text = "----------"
grd15.row = n + 4
grd15.Col = 5
grd15.Text = "«·„ƒ””…"
grd15.Col = 6
grd15.Text = (h + ss)
'grd15.Rows = 50
a = 0
sp = 0
'***** depenses
If dp.RecordCount > 0 Then
dp.MoveFirst
End If
Do While Not dp.EOF
dat3 = dp!dat
If dat3 >= dat1 And dat3 <= dat2 Then
a = dp!mon
sp = sp + a
End If
dp.MoveNext
Loop
Label39.Caption = (b + ss)
Label34.Caption = (d + p)
Label28.Caption = sp
Label27.Caption = ((b + ss) - ((d + p) + sp))
End Sub

Private Sub chargegrd16_MSChart3()
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
Dim g As Double
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
Dim t As Integer
Dim v As Double
Dim ss As Double
grd16.Rows = 1
grd16.Cols = 8
grd16.ColWidth(0) = 600
grd16.ColWidth(1) = 900
grd16.ColWidth(2) = 1100
grd16.ColWidth(3) = 1100
grd16.ColWidth(4) = 1100
grd16.ColWidth(5) = 1100
grd16.ColWidth(6) = 1100
grd16.ColWidth(7) = 1100
grd16.ColAlignment(0) = 1
grd16.ColAlignment(1) = 1
grd16.ColAlignment(2) = 1
grd16.ColAlignment(3) = 1
grd16.ColAlignment(4) = 1
grd16.ColAlignment(5) = 1
grd16.ColAlignment(6) = 1
grd16.ColAlignment(7) = 1
grd16.row = 0
grd16.Col = 0
grd16.Text = "«·‘Â—"
grd16.Col = 1
grd16.Text = "«·ﬁ”„"
grd16.Col = 2
grd16.Text = "«· ·«„Ì–"
grd16.Col = 3
grd16.Text = "«·√”« –…"
grd16.Col = 4
grd16.Text = "«·»«ﬁÌ"
grd16.Col = 5
grd16.Text = "% «·„ƒ””…"
grd16.Col = 6
grd16.Text = "«·„ƒ””…"
grd16.Col = 7
grd16.Text = "√”« –… «·‰”»…"
i = 1
cla1 = ""
mois1 = ""
sh = 0
sm = 0
sp = 0
dat1 = DT11.Value
dat2 = DT12.Value
'**** charge mois, classes et nbr
Call cont
grd16.Rows = cl.RecordCount + 5
Do While Not cl.EOF
t = 0
cla1 = cl!cla
If ps.RecordCount > 0 Then
ps.MoveFirst
End If
Do While Not ps.EOF
'**** grd16
If ps!cla = cla1 And ps!cas = "p" And Combo6.Text = ps!moi Then
t = 1
grd16.row = i
grd16.Col = 0
grd16.Text = ps!mois
grd16.Col = 1
grd16.Text = ps!cla
grd16.Col = 2
grd16.Text = "0"
grd16.Col = 3
grd16.Text = "0"
grd16.Col = 4
grd16.Text = "0"
grd16.Col = 5
grd16.Text = Label2.Caption
grd16.Col = 6
grd16.Text = "0"
i = i + 1
ps.MoveLast
End If
ps.MoveNext
Loop
If t = 0 Then
grd16.row = i
grd16.Col = 0
grd16.Text = Label38.Caption
grd16.Col = 1
grd16.Text = cl!cla
grd16.Col = 2
grd16.Text = "0"
grd16.Col = 3
grd16.Text = "0"
grd16.Col = 4
grd16.Text = "0"
grd16.Col = 5
grd16.Text = "100"
grd16.Col = 6
grd16.Text = "0"
i = i + 1
End If
cl.MoveNext
Loop
grd16.Rows = i
'Exit Sub

'**** professeurs pourcentage
n = grd16.Rows
If n > 1 Then
'**** charge montants etudiants
c = 0
Call cont
Do While Not ce.EOF
For i = 1 To n - 1
grd16.row = i
grd16.Col = 0
mois1 = grd16.Text
grd16.Col = 1
cla1 = grd16.Text
grd16.Col = 2
a = grd16.Text
If ce!mois = mois1 And ce!cla = cla1 And ce!moi <> "«· ”ÃÌ·" Then
b = ce!pay
c = a + b
grd16.row = i
grd16.Col = 2
grd16.Text = c
End If
Next i
ce.MoveNext
Loop
'Exit Sub
'**** charge montants professeurs secondaires
c = 0
r = 0
f = 0
Call cont
Do While Not ps.EOF
For i = 1 To n - 1
grd16.row = i
grd16.Col = 0
mois1 = grd16.Text
grd16.Col = 1
cla1 = grd16.Text
grd16.Col = 3
a = grd16.Text
If ps!mois = mois1 And ps!cla = cla1 And ps!cas <> "p" Then
e = ps!tot
f = a + e
e = ps!prm
f = f + e
e = ps!rtr
f = f + e
grd16.row = i
grd16.Col = 3
grd16.Text = f
End If
Next i
ps.MoveNext
Loop
'Exit Sub
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
grd16.row = i
grd16.Col = 2
a = grd16.Text
grd16.Col = 3
b = grd16.Text
grd16.Col = 5
Label52.Caption = grd16.Text
c = a - b
d = Label52.Caption
e = (c * d) / 100
MyNumber = Round(e, 0)
e = MyNumber
f = c - e
grd16.row = i
grd16.Col = 4
grd16.Text = c
grd16.Col = 6
grd16.Text = e
grd16.Col = 7
grd16.Text = f
Next i
'grd16.Rows = 70
End If
ss = 0
s = 0
Call cont
Do While Not ce.EOF
dat3 = ce!dat
If dat3 >= dat1 And dat3 <= dat2 Then
If ce!moi = "«· ”ÃÌ·" Then
s = ce!pay
ss = ss + s
End If
End If
ce.MoveNext
Loop
'Label25.Caption = ss
n = grd16.Rows
a = 0
b = 0
c = 0
d = 0
e = 0
f = 0
g = 0
h = 0
r = 0
p = 0
For i = 1 To n - 1
grd16.row = i
grd16.Col = 2
a = grd16.Text
b = b + a
grd16.Col = 3
c = grd16.Text
d = d + c
grd16.Col = 4
e = grd16.Text
f = f + e
grd16.Col = 6
g = grd16.Text
h = h + g
grd16.Col = 7
r = grd16.Text
p = p + r
Next i
grd16.Rows = n + 5
grd16.row = n
grd16.Col = 1
grd16.Text = ""
grd16.Col = 2
grd16.Text = "----------"
grd16.Col = 3
grd16.Text = "----------"
grd16.Col = 4
grd16.Text = "----------"
grd16.Col = 6
grd16.Text = "----------"
grd16.Col = 7
grd16.Text = "----------"
grd16.row = n + 1
grd16.Col = 1
grd16.Text = "«·„Ã„Ê⁄"
grd16.Col = 2
grd16.Text = b
grd16.Col = 3
grd16.Text = d
grd16.Col = 4
grd16.Text = f
grd16.Col = 6
grd16.Text = h
grd16.Col = 7
grd16.Text = p
grd16.row = n + 2
grd16.Col = 5
grd16.Text = "«·—”Ê„"
grd16.Col = 6
grd16.Text = ss
grd16.row = n + 3
grd16.Col = 6
grd16.Text = "----------"
grd16.row = n + 4
grd16.Col = 5
grd16.Text = "«·„ƒ””…"
grd16.Col = 6
grd16.Text = (h + ss)
'grd16.Rows = 50
v = 0
sp = 0
'***** depenses
Call cont
If dp.RecordCount > 0 Then
dp.MoveFirst
End If
Do While Not dp.EOF
dat3 = dp!dat
If dat3 >= dat1 And dat3 <= dat2 Then
v = dp!mon
sp = sp + v
End If
dp.MoveNext
Loop
Label21.Caption = (b + ss)
Label35.Caption = (d + p)
Label36.Caption = sp
Label37.Caption = ((b + ss) - ((d + p) + sp))
End Sub

Private Sub date_dt()
On Error Resume Next
Dim k As Double
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim f As Date
Dim g As Date
k = Text4.Text
If Combo6.Text = "«ﬂ Ê»—" Then
c = "10"
d = "31"
b = Text3.Text
ElseIf Combo6.Text = "‰Ê›„»—" Then
c = "11"
d = "30"
b = Text3.Text
ElseIf Combo6.Text = "œÌ”„»—" Then
c = "12"
d = "31"
b = Text3.Text
ElseIf Combo6.Text = "Ì‰«Ì—" Then
c = "01"
d = "31"
b = Text4.Text
ElseIf Combo6.Text = "›»—«Ì—" Then
c = "02"
If k Mod 4 = 0 Then
d = "29"
Else
d = "28"
End If
b = Text4.Text
ElseIf Combo6.Text = "„«—”" Then
c = "03"
d = "31"
b = Text4.Text
ElseIf Combo6.Text = "«»—Ì·" Then
c = "04"
d = "30"
b = Text4.Text
ElseIf Combo6.Text = "„«ÌÊ" Then
c = "05"
d = "31"
b = Text4.Text
ElseIf Combo6.Text = "ÌÊ‰ÌÊ" Then
c = "06"
d = "30"
b = Text4.Text
ElseIf Combo6.Text = "ÌÊ·ÌÊ" Then
c = "07"
d = "31"
b = Text4.Text
ElseIf Combo6.Text = "√€”ÿ”" Then
c = "08"
d = "31"
b = Text4.Text
ElseIf Combo6.Text = "”» „»—" Then
c = "09"
d = "30"
b = Text4.Text
End If
a = "01/" + c + "/" + b
e = d + "/" + c + "/" + b
f = a
g = e
Label38.Caption = Val(c)
DT11.Value = f
DT12.Value = g
End Sub
Private Sub chargec6()
On Error Resume Next
Combo6.Clear
Combo6.AddItem "«ﬂ Ê»—"
Combo6.AddItem "‰Ê›„»—"
Combo6.AddItem "œÌ”„»—"
Combo6.AddItem "Ì‰«Ì—"
Combo6.AddItem "›»—«Ì—"
Combo6.AddItem "„«—”"
Combo6.AddItem "«»—Ì·"
Combo6.AddItem "„«ÌÊ"
Combo6.AddItem "ÌÊ‰ÌÊ"
Combo6.AddItem "ÌÊ·ÌÊ"
Combo6.AddItem "√€”ÿ”"
Combo6.AddItem "”» „»—"

End Sub
Private Sub chargegrd17_MSChart2()
On Error Resume Next
'On Error Resume Next
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
Dim g As Double
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
Dim t As Integer
Dim v As Double
Dim ss As Double
grd17.Rows = 1
grd17.Cols = 8
grd17.ColWidth(0) = 600
grd17.ColWidth(1) = 900
grd17.ColWidth(2) = 1100
grd17.ColWidth(3) = 1100
grd17.ColWidth(4) = 1100
grd17.ColWidth(5) = 1100
grd17.ColWidth(6) = 1100
grd17.ColWidth(7) = 1100
grd17.ColAlignment(0) = 1
grd17.ColAlignment(1) = 1
grd17.ColAlignment(2) = 1
grd17.ColAlignment(3) = 1
grd17.ColAlignment(4) = 1
grd17.ColAlignment(5) = 1
grd17.ColAlignment(6) = 1
grd17.ColAlignment(7) = 1
grd17.row = 0
grd17.Col = 0
grd17.Text = "«·‘Â—"
grd17.Col = 1
grd17.Text = "«·ﬁ”„"
grd17.Col = 2
grd17.Text = "«· ·«„Ì–"
grd17.Col = 3
grd17.Text = "«·√”« –…"
grd17.Col = 4
grd17.Text = "«·»«ﬁÌ"
grd17.Col = 5
grd17.Text = "% «·„ƒ””…"
grd17.Col = 6
grd17.Text = "«·„ƒ””…"
grd17.Col = 7
grd17.Text = "√”« –… «·‰”»…"
i = 1
cla1 = ""
mois1 = ""
sh = 0
sm = 0
sp = 0
dat1 = DT1.Value
dat2 = DT2.Value
t = 9
i = 1
grd17.Rows = 13
For i = 1 To 12
If i < 4 Then
t = t + 1
Else
If t = 12 Then
t = 0
End If
t = t + 1
End If
grd17.row = i
grd17.Col = 0
grd17.Text = t
grd17.Col = 1
grd17.Text = Combo7.Text
grd17.Col = 2
grd17.Text = "0"
grd17.Col = 3
grd17.Text = "0"
grd17.Col = 4
grd17.Text = "0"
grd17.Col = 5
grd17.Text = "0"
grd17.Col = 6
grd17.Text = "0"
Next i
t = 0
'**** charge mois, classes et nbr
Call cont
Do While Not ps.EOF
'grd17.Rows = ps.RecordCount + 5
'**** grd17
If ps!cla = Combo7.Text And ps!cas = "p" Then
t = 1
For i = 1 To 12
grd17.row = i
grd17.Col = 5
grd17.Text = Label2.Caption
Next i
ps.MoveLast
End If
ps.MoveNext
Loop
If t = 0 Then
For i = 1 To 12
grd17.row = i
grd17.Col = 5
grd17.Text = "100"
Next i
End If
'Exit Sub
'**** professeurs pourcentage
n = grd17.Rows
If n > 1 Then
'**** charge montants etudiants
c = 0
Call cont
Do While Not ce.EOF
For i = 1 To n - 2
grd17.row = i
grd17.Col = 0
mois1 = grd17.Text
grd17.Col = 1
cla1 = grd17.Text
grd17.Col = 2
a = grd17.Text
If ce!cla = cla1 And ce!mois = mois1 And ce!moi <> "«· ”ÃÌ·" Then
b = ce!pay
c = a + b
grd17.row = i
grd17.Col = 2
grd17.Text = c
End If
Next i
ce.MoveNext
Loop
'Exit Sub
'**** charge montants professeurs secondaires
c = 0
r = 0
f = 0
Call cont
Do While Not ps.EOF
For i = 1 To n - 1
grd17.row = i
grd17.Col = 0
mois1 = grd17.Text
grd17.Col = 1
cla1 = grd17.Text
grd17.Col = 3
a = grd17.Text
If ps!cla = cla1 And ps!mois = mois1 And ps!cas <> "p" Then
e = ps!tot
f = a + e
e = ps!prm
f = f + e
e = ps!rtr
f = f + e
grd17.row = i
grd17.Col = 3
grd17.Text = f
End If
Next i
ps.MoveNext
Loop
'Exit Sub
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
grd17.row = i
grd17.Col = 2
a = grd17.Text
grd17.Col = 3
b = grd17.Text
grd17.Col = 5
Label52.Caption = grd17.Text
c = a - b
d = Label52.Caption
e = (c * d) / 100
MyNumber = Round(e, 0)
e = MyNumber
f = c - e
grd17.row = i
grd17.Col = 4
grd17.Text = c
grd17.Col = 6
grd17.Text = e
grd17.Col = 7
grd17.Text = f
Next i
'grd17.Rows = 70
End If
ss = 0
s = 0
Call cont
Do While Not ce.EOF
dat3 = ce!dat
If dat3 >= dat1 And dat3 <= dat2 Then
If ce!moi = "«· ”ÃÌ·" And ce!cla = Combo7.Text Then
s = ce!pay
ss = ss + s
End If
End If
ce.MoveNext
Loop
'Label25.Caption = ss
n = grd17.Rows
a = 0
b = 0
c = 0
d = 0
e = 0
f = 0
g = 0
h = 0
r = 0
p = 0
For i = 1 To n - 1
grd17.row = i
grd17.Col = 2
a = grd17.Text
b = b + a
grd17.Col = 3
c = grd17.Text
d = d + c
grd17.Col = 4
e = grd17.Text
f = f + e
grd17.Col = 6
g = grd17.Text
h = h + g
grd17.Col = 7
r = grd17.Text
p = p + r
Next i
grd17.Rows = n + 5
grd17.row = n
grd17.Col = 1
grd17.Text = ""
grd17.Col = 2
grd17.Text = "----------"
grd17.Col = 3
grd17.Text = "----------"
grd17.Col = 4
grd17.Text = "----------"
grd17.Col = 6
grd17.Text = "----------"
grd17.Col = 7
grd17.Text = "----------"
grd17.row = n + 1
grd17.Col = 1
grd17.Text = "«·„Ã„Ê⁄"
grd17.Col = 2
grd17.Text = b
grd17.Col = 3
grd17.Text = d
grd17.Col = 4
grd17.Text = f
grd17.Col = 6
grd17.Text = h
grd17.Col = 7
grd17.Text = p
grd17.row = n + 2
grd17.Col = 5
grd17.Text = "«·—”Ê„"
grd17.Col = 6
grd17.Text = ss
grd17.row = n + 3
grd17.Col = 6
grd17.Text = "----------"
grd17.row = n + 4
grd17.Col = 5
grd17.Text = "«·„ƒ””…"
grd17.Col = 6
grd17.Text = (h + ss)
Label33.Caption = (b + ss)
Label32.Caption = (d + p)
Label29.Caption = ((b + ss) - (d + p))
End Sub
Private Sub chargegrd18_MSChart1()
On Error Resume Next
'On Error Resume Next
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
Dim g As Double
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
Dim t As Integer
Dim v As Double
Dim ss As Double
grd18.Rows = 1
grd18.Cols = 8
grd18.ColWidth(0) = 600
grd18.ColWidth(1) = 900
grd18.ColWidth(2) = 1100
grd18.ColWidth(3) = 1100
grd18.ColWidth(4) = 1100
grd18.ColWidth(5) = 1100
grd18.ColWidth(6) = 1100
grd18.ColWidth(7) = 1100
grd18.ColAlignment(0) = 1
grd18.ColAlignment(1) = 1
grd18.ColAlignment(2) = 1
grd18.ColAlignment(3) = 1
grd18.ColAlignment(4) = 1
grd18.ColAlignment(5) = 1
grd18.ColAlignment(6) = 1
grd18.ColAlignment(7) = 1
grd18.row = 0
grd18.Col = 0
grd18.Text = "«·‘Â—"
grd18.Col = 1
grd18.Text = "«·ﬁ”„"
grd18.Col = 2
grd18.Text = "«· ·«„Ì–"
grd18.Col = 3
grd18.Text = "«·√”« –…"
grd18.Col = 4
grd18.Text = "«·»«ﬁÌ"
grd18.Col = 5
grd18.Text = "% «·„ƒ””…"
grd18.Col = 6
grd18.Text = "«·„ƒ””…"
grd18.Col = 7
grd18.Text = "√”« –… «·‰”»…"
i = 1
cla1 = ""
mois1 = ""
sh = 0
sm = 0
sp = 0
dat1 = DT1.Value
dat2 = DT2.Value
t = 0
i = 1
'**** charge mois, classes et nbr
Call cont
Do While Not ps.EOF
grd18.Rows = ps.RecordCount + 5
'**** grd18
If ps!cla = Combo3.Text And ps!mois = Combo5.Text And ps!cas = "p" Then
t = 1
grd18.row = i
grd18.Col = 0
grd18.Text = ps!mois
grd18.Col = 1
grd18.Text = ps!cla
grd18.Col = 2
grd18.Text = "0"
grd18.Col = 3
grd18.Text = "0"
grd18.Col = 4
grd18.Text = "0"
grd18.Col = 5
grd18.Text = Label2.Caption
grd18.Col = 6
grd18.Text = "0"
i = i + 1
ps.MoveLast
End If
ps.MoveNext
Loop
If t = 0 Then
grd18.row = i
grd18.Col = 0
grd18.Text = Combo5.Text
grd18.Col = 1
grd18.Text = Combo3.Text
grd18.Col = 2
grd18.Text = "0"
grd18.Col = 3
grd18.Text = "0"
grd18.Col = 4
grd18.Text = "0"
grd18.Col = 5
grd18.Text = "100"
grd18.Col = 6
grd18.Text = "0"
i = i + 1
End If
grd18.Rows = i
'Exit Sub

'**** professeurs pourcentage
n = grd18.Rows
If n > 1 Then
'**** charge montants etudiants
c = 0
Call cont
Do While Not ce.EOF
For i = 1 To n - 1
grd18.row = i
grd18.Col = 0
mois1 = grd18.Text
grd18.Col = 1
cla1 = grd18.Text
grd18.Col = 2
a = grd18.Text
If ce!cla = cla1 And ce!mois = mois1 And ce!moi <> "«· ”ÃÌ·" Then
b = ce!pay
c = a + b
grd18.row = i
grd18.Col = 2
grd18.Text = c
End If
Next i
ce.MoveNext
Loop
'Exit Sub
'**** charge montants professeurs secondaires
c = 0
r = 0
f = 0
Call cont
Do While Not ps.EOF
For i = 1 To n - 1
grd18.row = i
grd18.Col = 0
mois1 = grd18.Text
grd18.Col = 1
cla1 = grd18.Text
grd18.Col = 3
a = grd18.Text
If ps!cla = cla1 And ps!mois = mois1 And ps!cas <> "p" Then
e = ps!tot
f = a + e
e = ps!prm
f = f + e
e = ps!rtr
f = f + e
grd18.row = i
grd18.Col = 3
grd18.Text = f
End If
Next i
ps.MoveNext
Loop
'Exit Sub
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
grd18.row = i
grd18.Col = 2
a = grd18.Text
grd18.Col = 3
b = grd18.Text
grd18.Col = 5
Label52.Caption = grd18.Text
c = a - b
d = Label52.Caption
e = (c * d) / 100
MyNumber = Round(e, 0)
e = MyNumber
f = c - e
grd18.row = i
grd18.Col = 4
grd18.Text = c
grd18.Col = 6
grd18.Text = e
grd18.Col = 7
grd18.Text = f
Next i
'grd18.Rows = 70
End If
ss = 0
s = 0
Call cont
Do While Not ce.EOF
dat3 = ce!dat
If dat3 >= dat1 And dat3 <= dat2 Then
If ce!moi = "«· ”ÃÌ·" And ce!cla = Combo3.Text And ce!mois = Combo5.Text Then
s = ce!pay
ss = ss + s
End If
End If
ce.MoveNext
Loop
'Label25.Caption = ss
n = grd18.Rows
a = 0
b = 0
c = 0
d = 0
e = 0
f = 0
g = 0
h = 0
r = 0
p = 0
For i = 1 To n - 1
grd18.row = i
grd18.Col = 2
a = grd18.Text
b = b + a
grd18.Col = 3
c = grd18.Text
d = d + c
grd18.Col = 4
e = grd18.Text
f = f + e
grd18.Col = 6
g = grd18.Text
h = h + g
grd18.Col = 7
r = grd18.Text
p = p + r
Next i
grd18.Rows = n + 5
grd18.row = n
grd18.Col = 1
grd18.Text = ""
grd18.Col = 2
grd18.Text = "----------"
grd18.Col = 3
grd18.Text = "----------"
grd18.Col = 4
grd18.Text = "----------"
grd18.Col = 6
grd18.Text = "----------"
grd18.Col = 7
grd18.Text = "----------"
grd18.row = n + 1
grd18.Col = 1
grd18.Text = "«·„Ã„Ê⁄"
grd18.Col = 2
grd18.Text = b
grd18.Col = 3
grd18.Text = d
grd18.Col = 4
grd18.Text = f
grd18.Col = 6
grd18.Text = h
grd18.Col = 7
grd18.Text = p
grd18.row = n + 2
grd18.Col = 5
grd18.Text = "«·—”Ê„"
grd18.Col = 6
grd18.Text = ss
grd18.row = n + 3
grd18.Col = 6
grd18.Text = "----------"
grd18.row = n + 4
grd18.Col = 5
grd18.Text = "«·„ƒ””…"
grd18.Col = 6
grd18.Text = (h + ss)
Label20.Caption = (b + ss)
Label22.Caption = (d + p)
Label23.Caption = ((b + ss) - (d + p))
End Sub


