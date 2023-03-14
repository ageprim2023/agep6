VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ActiveSkin.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form classes 
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
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Ã·» «”„«¡ «· ·«„Ì– „‰ ’›Õ… «ﬂ”·"
      TabPicture(0)   =   "classes.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "«·ÃœÊ· «·“„‰Ì"
      TabPicture(1)   =   "classes.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture11"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "«·„Ê«œ"
      TabPicture(2)   =   "classes.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Skin1"
      Tab(2).Control(1)=   "Picture3"
      Tab(2).Control(2)=   "Picture8"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "«·√ﬁ”«„"
      TabPicture(3)   =   "classes.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Picture1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Picture7"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin VB.PictureBox Picture11 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8775
         Left            =   -74880
         ScaleHeight     =   8775
         ScaleWidth      =   14295
         TabIndex        =   57
         Top             =   360
         Width           =   14295
         Begin VB.CommandButton Command16 
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
            Left            =   120
            TabIndex        =   63
            Top             =   120
            Width           =   2655
         End
         Begin VB.CommandButton Command15 
            Caption         =   "”Õ» «·ÃœÊ·"
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
            Left            =   11520
            TabIndex        =   62
            Top             =   120
            Width           =   2655
         End
         Begin VB.CommandButton Command14 
            Caption         =   "⁄—÷ «·ÃœÊ·"
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
            TabIndex        =   59
            Top             =   120
            Visible         =   0   'False
            Width           =   2655
         End
         Begin MSFlexGridLib.MSFlexGrid grd8 
            Height          =   8295
            Left            =   120
            TabIndex        =   58
            Top             =   480
            Width           =   14055
            _ExtentX        =   24791
            _ExtentY        =   14631
            _Version        =   393216
            Cols            =   3
            FixedCols       =   2
            BackColor       =   0
            ForeColor       =   0
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
      End
      Begin VB.PictureBox Picture10 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8775
         Left            =   -74880
         ScaleHeight     =   8775
         ScaleWidth      =   14295
         TabIndex        =   46
         Top             =   360
         Width           =   14295
         Begin VB.CommandButton Command13 
            Caption         =   "„‰ ÃœÌœ"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   56
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Õ›Ÿ «·»Ì«‰« "
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1320
            TabIndex        =   53
            Top             =   120
            Width           =   2295
         End
         Begin VB.CommandButton Command11 
            Caption         =   "«· Õﬁﬁ „‰ «·√Œÿ«¡"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3720
            TabIndex        =   52
            Top             =   120
            Width           =   2295
         End
         Begin VB.CommandButton Command10 
            Caption         =   "› Õ ’›Õ… «ﬂ”·"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   8520
            TabIndex        =   50
            Top             =   120
            Width           =   2295
         End
         Begin VB.CommandButton Command9 
            Caption         =   "⁄—÷ „Õ ÊÏ ’›Õ… «ﬂ”·"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6120
            TabIndex        =   49
            Top             =   120
            Width           =   2295
         End
         Begin MSFlexGridLib.MSFlexGrid grd6 
            Height          =   8055
            Left            =   10920
            TabIndex        =   47
            Top             =   600
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   14208
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
         Begin MSFlexGridLib.MSFlexGrid grd7 
            Height          =   7695
            Left            =   120
            TabIndex        =   51
            Top             =   1080
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   13573
            _Version        =   393216
            Cols            =   3
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
         Begin VB.Label Label21 
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
            Left            =   2880
            TabIndex        =   55
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "„Ã„Ê⁄ «· ·«„Ì–"
            BeginProperty Font 
               Name            =   "Times New Roman"
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
            TabIndex        =   54
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·√ﬁ”«„ «·„ÊÃÊœ…"
            BeginProperty Font 
               Name            =   "Times New Roman"
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
            TabIndex        =   48
            Top             =   120
            Width           =   3375
         End
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8655
         Left            =   -74880
         ScaleHeight     =   8655
         ScaleWidth      =   9495
         TabIndex        =   34
         Top             =   360
         Width           =   9495
         Begin VB.PictureBox Picture9 
            Height          =   5775
            Left            =   240
            ScaleHeight     =   5715
            ScaleWidth      =   8595
            TabIndex        =   35
            Top             =   960
            Visible         =   0   'False
            Width           =   8655
            Begin VB.PictureBox Picture5 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   2775
               Left            =   960
               ScaleHeight     =   2775
               ScaleWidth      =   6735
               TabIndex        =   36
               Top             =   1560
               Width           =   6735
               Begin VB.CommandButton Command7 
                  Caption         =   "”Õ» «·ÃœÊ· «·“„‰Ì"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   4800
                  TabIndex        =   41
                  Top             =   8040
                  Width           =   4815
               End
               Begin VB.PictureBox Picture6 
                  Height          =   3735
                  Left            =   1440
                  ScaleHeight     =   3675
                  ScaleWidth      =   4275
                  TabIndex        =   38
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   4335
                  Begin VB.Timer Timer3 
                     Enabled         =   0   'False
                     Interval        =   50
                     Left            =   480
                     Top             =   480
                  End
                  Begin VB.Label Label10 
                     Height          =   375
                     Left            =   120
                     TabIndex        =   40
                     Top             =   1080
                     Width           =   1095
                  End
                  Begin VB.Label Label9 
                     Height          =   375
                     Left            =   120
                     TabIndex        =   39
                     Top             =   120
                     Width           =   1455
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
                  Left            =   4800
                  Style           =   2  'Dropdown List
                  TabIndex        =   37
                  Top             =   120
                  Width           =   3615
               End
               Begin MSFlexGridLib.MSFlexGrid grd3 
                  Height          =   5895
                  Left            =   360
                  TabIndex        =   42
                  Top             =   600
                  Width           =   10935
                  _ExtentX        =   19288
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
               Begin VB.Label Label13 
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
                  TabIndex        =   43
                  Top             =   120
                  Width           =   1935
               End
            End
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "2012-2013"
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
            Left            =   120
            TabIndex        =   45
            Top             =   6960
            Width           =   9255
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "„Ã„⁄ „œ«—” «· ﬁÊÏ «·Õ—…"
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
            Height          =   6735
            Left            =   120
            TabIndex        =   44
            Top             =   120
            Width           =   9255
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8655
         Left            =   -65280
         ScaleHeight     =   8655
         ScaleWidth      =   4695
         TabIndex        =   14
         Top             =   360
         Width           =   4695
         Begin VB.PictureBox Picture4 
            Height          =   3735
            Left            =   1680
            ScaleHeight     =   3675
            ScaleWidth      =   1635
            TabIndex        =   23
            Top             =   3720
            Visible         =   0   'False
            Width           =   1695
            Begin VB.Timer Timer2 
               Enabled         =   0   'False
               Interval        =   50
               Left            =   480
               Top             =   480
            End
            Begin VB.Label Label4 
               Height          =   375
               Left            =   120
               TabIndex        =   27
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Label Label5 
               Height          =   375
               Left            =   120
               TabIndex        =   26
               Top             =   120
               Width           =   1455
            End
            Begin VB.Label Label17 
               Caption         =   "Label17"
               Height          =   255
               Left            =   240
               TabIndex        =   25
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label Label18 
               Caption         =   "Label18"
               Height          =   255
               Left            =   0
               TabIndex        =   24
               Top             =   960
               Width           =   1575
            End
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
            Left            =   120
            TabIndex        =   22
            Top             =   600
            Width           =   3615
         End
         Begin VB.CommandButton Command4 
            Caption         =   " ÕœÌÀ «·»Ì«‰« "
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
            TabIndex        =   21
            Top             =   2520
            Width           =   2055
         End
         Begin VB.CommandButton Command5 
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
            Left            =   2520
            TabIndex        =   20
            Top             =   2520
            Width           =   2055
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
            Left            =   840
            TabIndex        =   19
            Top             =   2040
            Width           =   3015
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
            Left            =   2160
            TabIndex        =   18
            Top             =   1080
            Width           =   1575
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
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   120
            Width           =   1575
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
            Left            =   2160
            TabIndex        =   16
            Top             =   1560
            Width           =   1575
         End
         Begin VB.CommandButton Command8 
            Caption         =   " €ÌÌ— ⁄œœ «·„Ê«œ"
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
            TabIndex        =   15
            Top             =   1560
            Width           =   1935
         End
         Begin MSComctlLib.ProgressBar ProgressBar2 
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   3000
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
         End
         Begin MSFlexGridLib.MSFlexGrid grd2 
            Height          =   5055
            Left            =   120
            TabIndex        =   29
            Top             =   3480
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   8916
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
         Begin VB.Label Label6 
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
            Left            =   2640
            TabIndex        =   33
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label7 
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
            Left            =   2640
            TabIndex        =   32
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·÷«—»"
            BeginProperty Font 
               Name            =   "Times New Roman"
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
            TabIndex        =   31
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·„Ê«œ"
            BeginProperty Font 
               Name            =   "Times New Roman"
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
            TabIndex        =   30
            Top             =   1560
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8655
         Left            =   120
         ScaleHeight     =   8655
         ScaleWidth      =   9495
         TabIndex        =   11
         Top             =   360
         Width           =   9495
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "„Ã„⁄ „œ«—” «· ﬁÊÏ «·Õ—…"
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
            Height          =   6735
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   9255
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "2012-2013"
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
            Left            =   120
            TabIndex        =   12
            Top             =   6960
            Width           =   9255
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8655
         Left            =   9720
         ScaleHeight     =   8655
         ScaleWidth      =   4695
         TabIndex        =   1
         Top             =   360
         Width           =   4695
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
            Left            =   2400
            TabIndex        =   61
            Top             =   120
            Width           =   1575
         End
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
            Left            =   1800
            TabIndex        =   8
            Top             =   600
            Width           =   2775
         End
         Begin VB.CommandButton Command3 
            Caption         =   " ÕœÌÀ «·»Ì«‰« "
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
            TabIndex        =   7
            Top             =   600
            Width           =   1575
         End
         Begin VB.PictureBox Picture2 
            Height          =   3735
            Left            =   1680
            ScaleHeight     =   3675
            ScaleWidth      =   1635
            TabIndex        =   2
            Top             =   3720
            Visible         =   0   'False
            Width           =   1695
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
               ItemData        =   "classes.frx":0070
               Left            =   120
               List            =   "classes.frx":0086
               Style           =   2  'Dropdown List
               TabIndex        =   60
               Top             =   1320
               Width           =   1575
            End
            Begin VB.Timer Timer1 
               Enabled         =   0   'False
               Interval        =   50
               Left            =   480
               Top             =   480
            End
            Begin VB.CommandButton Command2 
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
               Left            =   -240
               TabIndex        =   3
               Top             =   2160
               Width           =   2055
            End
            Begin VB.Label Label2 
               Height          =   375
               Left            =   120
               TabIndex        =   5
               Top             =   120
               Width           =   1455
            End
            Begin VB.Label Label3 
               Caption         =   "Label3"
               Height          =   375
               Left            =   120
               TabIndex        =   4
               Top             =   1080
               Width           =   1095
            End
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   1080
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
         End
         Begin MSFlexGridLib.MSFlexGrid grd1 
            Height          =   6975
            Left            =   120
            TabIndex        =   9
            Top             =   1560
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   12303
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
         Begin VB.Label Label1 
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
            Left            =   2640
            TabIndex        =   10
            Top             =   120
            Width           =   1935
         End
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   -73080
         OleObjectBlob   =   "classes.frx":00A2
         Top             =   5220
      End
   End
End
Attribute VB_Name = "classes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public co2 As ADODB.Connection
Public cr2 As ADODB.Recordset
Public be As ADODB.Recordset
Public ce As ADODB.Recordset
Dim data As New Access.Application
Dim anes As String
Function cont2()
Set co2 = New ADODB.Connection
Set cr2 = New ADODB.Recordset
Set be = New ADODB.Recordset
Set ce = New ADODB.Recordset
co2.Provider = "microsoft.jet.oledb.4.0; jet oledb:database password=7346804"
anes = "C" + face.SBB1.Panels(9).Text
co2.ConnectionString = App.Path & "\" & anes & ".mdb"
co2.Open
cr2.Open "select*from Tcarts", co2, adOpenKeyset, adLockOptimistic
be.Open "select*from Tbulletin", co2, adOpenKeyset, adLockOptimistic
ce.Open "select*from Tcartes", co2, adOpenKeyset, adLockOptimistic
End Function

Private Sub Combo1_Change()
On Error Resume Next
grd2.Visible = False
grd2.Clear
grd2.Rows = 1
Text2.SetFocus
Text4.Text = ""
Command8.Enabled = False
Text4.Enabled = True
Call chargegrd2
grd2.Visible = True
Call cont
Do While Not cl.EOF
If Combo1.Text = cl!cla Then
Label17.Caption = cl!num
Exit Sub
End If
cl.MoveNext
Loop
End Sub

Private Sub Combo1_Click()
On Error Resume Next
Combo1_Change
End Sub

Private Sub Combo2_Change()
On Error Resume Next
grd3.Visible = False
Call chargetimes
grd3.Visible = True
End Sub

Private Sub Combo2_Click()
On Error Resume Next
Combo2_Change
End Sub

Private Sub Combo3_Change()
Text1.Text = Combo3.Text
End Sub

Private Sub Combo3_Click()
Combo3_Change
End Sub
Private Sub charge_combo3()
Combo3.Clear
Combo3.AddItem "5C"
Combo3.AddItem "5D"
Combo3.AddItem "6C"
Combo3.AddItem "6D"
Combo3.AddItem "7C"
Combo3.AddItem "7D"
End Sub
Private Sub Command1_Click()
'On Error Resume Next
Dim tx1 As String
Dim tx2 As String
Dim a As Double
Text1.Text = Trim(Text1.Text)
If Text1.Text = "" Then
MsgBox "«œŒ· «·ﬁ”„", vbCritical
Text1.SetFocus
Exit Sub
End If
Command11.Enabled = True
Command12.Enabled = False
tx2 = Text1.Text
If Label2.Caption <> "" Then
tx1 = Label3.Caption
g = MsgBox("·ﬁœ ﬁ„  » ⁄œÌ· «”„ «·ﬁ”„ " + tx1 + " ≈·Ï " + tx2 + " Â·  —Ìœ «·«” „—«—ø", vbInformation + vbYesNo + arabic, "AGEP6")
If g = vbYes Then
Call cont
Do While Not cl.EOF
If Label2.Caption <> cl!aut And Text1.Text = cl!cla Then
MsgBox "€Ì— „„ﬂ‰ ·ﬁœ  „ ÕÃ“ «”„ " + tx2 + " ·ﬁ”„ ¬Œ— ”«»ﬁ«", vbExclamation + arabic
Exit Sub
End If
cl.MoveNext
Loop
Call cont
Do While Not cl.EOF
If Label2.Caption = cl!aut Then
cl!cla = Text1.Text
cl.Update
Timer1.Enabled = True
Exit Sub
End If
cl.MoveNext
Loop
End If
Exit Sub
End If
Call cont
Do While Not cl.EOF
If Label2.Caption <> cl!aut And Text1.Text = cl!cla Then
MsgBox "€Ì— „„ﬂ‰ ·ﬁœ  „ ÕÃ“ «”„ " + tx2 + " ·ﬁ”„ ¬Œ— ”«»ﬁ«", vbExclamation + arabic
Exit Sub
End If
cl.MoveNext
Loop
cl.AddNew
a = cl!aut
cl!cla = Text1.Text
cl!num = a
cl!act = "1"
cl.Update
Timer1.Enabled = True

End Sub

Private Sub Command10_Click()
On Error Resume Next
Command10.Enabled = False
FileCopy App.Path & "\NomVide.xls", App.Path & "\NomEtudiants.xls"
Set kb = CreateObject("Excel.application")
kb.Workbooks.Open (App.Path & "\NomEtudiants.xls")
kb.Visible = True
Command10.Enabled = False
Command9.Enabled = True

End Sub

Private Sub Command11_Click()
On Error GoTo u
Dim i As Double
Dim n As Double
Dim j As Double
Dim m As Double
Dim k As Double
Dim b As Double
Dim c As Double
Dim tx1 As String
Dim tx2 As String
Dim tx3 As String
Dim tx4 As String
Dim tx5 As String
Dim tx6 As String
Dim tx7 As String
Dim tx8 As String
n = grd6.Rows
m = grd7.Rows
c = 0
b = 0
Command11.Enabled = False
'**** verify class
For j = 1 To m - 1
grd7.row = j
grd7.Col = 0
tx1 = grd7.Text
grd7.CellBackColor = &H0&
For i = 1 To n - 1
grd6.row = i
grd6.Col = 0
tx2 = grd6.Text
k = 0
If tx1 = tx2 Then
k = 1
i = n
End If
Next i
If k = 0 Then
c = 1
grd7.row = j
grd7.Col = 0
grd7.CellBackColor = &HFF&
End If
Next j
If c = 1 Then
MsgBox " „ «ﬂ ‘«› «ﬁ”«„ €Ì— „ ÿ«»ﬁ… „⁄ «·√ﬁ”«„ «·„ÊÃÊœ… ›Ì «·ﬁ«∆„… »«·Ì„Ì‰ ,  „  ·ÊÌ‰Â« »«··Ê‰ «·√Õ„—, ÌÃ»  ’ÕÌÕ «·Ê÷⁄Ì… Õ Ï Ì ”‰Ï Õ›Ÿ «·»Ì«‰«  »«·ﬂ«„·", vbExclamation + arabic
Command11.Enabled = True
Exit Sub
End If
'**** verify Etudiants
grd7.Visible = False
For j = 1 To m - 1
grd7.row = j
grd7.Col = 0
grd7.CellBackColor = &H0&
grd7.Col = 1
grd7.CellBackColor = &H0&
grd7.Col = 2
grd7.CellBackColor = &H0&
Next j
b = 0
c = 0
Call cont
Do While Not et.EOF
tx1 = et!cla
tx2 = et!num
tx3 = et!nom
tx7 = Val(et!ser)
For j = 1 To m - 1
grd7.row = j
grd7.Col = 0
tx4 = grd7.Text
grd7.Col = 1
tx5 = grd7.Text
grd7.Col = 2
tx6 = grd7.Text
grd7.Col = 4
tx8 = Val(grd7.Text)
If tx7 = tx8 Then
b = 1
grd7.row = j
grd7.Col = 4
grd7.CellBackColor = &HFF&
End If
If tx1 = tx4 And tx2 = tx5 And tx3 = tx6 Then
c = 1
grd7.row = j
grd7.Col = 0
grd7.CellBackColor = &HFF&
grd7.Col = 1
grd7.CellBackColor = &HFF&
grd7.Col = 2
grd7.CellBackColor = &HFF&
End If
Next j
et.MoveNext
Loop
grd7.Visible = True
If b = 1 Then
MsgBox " „ «ﬂ ‘«› —ﬁ„  ”·”·Ì „ ÿ«»ﬁ „⁄ —ﬁ„  ”·”·Ì „ÊÃÊœ ›Ì ﬁ«⁄œ… «·»Ì«‰«  ,  „  ·ÊÌ‰Â »«··Ê‰ «·√Õ„—, ÌÃ»  ’ÕÌÕ «·Ê÷⁄Ì… Õ Ï Ì ”‰Ï Õ›Ÿ «·»Ì«‰«  »«·ﬂ«„·", vbExclamation + arabic
Command11.Enabled = True
Exit Sub
End If
If c = 1 Then
MsgBox " „ «ﬂ ‘«› «ﬁ”«„ Ê√”„«¡ Ê√—ﬁ«„  ·«„Ì– „ ÿ«»ﬁ… „⁄ √ﬁ”«„ Ê√”„«¡ Ê√—ﬁ«„  ·«„Ì– „ÊÃÊœ… ›Ì ﬁ«⁄œ… «·»Ì«‰«  ,  „  ·ÊÌ‰Â« »«··Ê‰ «·√Õ„—, ÌÃ»  ’ÕÌÕ «·Ê÷⁄Ì… Õ Ï Ì ”‰Ï Õ›Ÿ «·»Ì«‰«  »«·ﬂ«„·", vbExclamation + arabic
Command11.Enabled = True
Exit Sub
End If
'*** meme code en grd7
b = 0
grd7.Col = 4
grd7.Sort = 1
n = grd7.Rows
For i = 1 To n - 1
grd7.row = i
grd7.Col = 1
k = grd7.Text
grd7.Col = 4
m = grd7.Text
grd7.Col = 4
grd7.row = i - 1
tx7 = Val(grd7.Text)
grd7.Col = 4
grd7.row = i
tx8 = Val(grd7.Text)
If tx7 = tx8 Then
b = 1
grd7.row = i - 1
grd7.Col = 4
grd7.CellBackColor = &HFF&
grd7.row = i
grd7.Col = 4
grd7.CellBackColor = &HFF&
End If
If b = 1 Then
MsgBox " „ «ﬂ ‘«› —ﬁ„Ì  ”·”· „ ÿ«»ﬁÌ‰ ,  „  ·ÊÌ‰Â„« »«··Ê‰ «·√Õ„—, ÌÃ»  ’ÕÌÕ «·Ê÷⁄Ì… Õ Ï Ì ”‰Ï Õ›Ÿ «·»Ì«‰«  »«·ﬂ«„·", vbExclamation + arabic
Command11.Enabled = True
Exit Sub
End If
Next i
grd7.Visible = True
Command11.Enabled = False
Command12.Enabled = True
Exit Sub
u:
MsgBox "ÌÊÃœ Œÿ√ , ÌÃ» «· Õﬁﬁ „‰ «·√—ﬁ«„ «· ”·”·Ì… √Ê «·√—ﬁ«„ œ«Œ· «·ﬁ”„", vbExclamation
Command11.Enabled = True
End Sub

Private Sub Command12_Click()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim a As Double
Dim b As String
Dim c As Double
n = grd7.Rows
a = 0
Call cont
For i = 1 To n - 1
et.AddNew
grd7.row = i
grd7.Col = 0
et!cla = grd7.Text
grd7.Col = 1
et!num = grd7.Text
grd7.Col = 2
et!nom = grd7.Text
grd7.Col = 3
et!tel = grd7.Text
grd7.Col = 4
et!ser = grd7.Text
et!dat = Date
et!sex = ""
et!pho = "01"
et!adr = ""
et!act = "1"
et.Update
Next i
a = 0
Call cont
Do While Not et.EOF
b = et!ser
If b > a Then
a = b
End If
et.MoveNext
Loop
sr!num = a + 1
sr.Update
MsgBox " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ", vbInformation
Command13_Click
End Sub

Private Sub Command13_Click()
On Error Resume Next
Command10.Enabled = True
Command9.Enabled = False
Command11.Enabled = False
Command12.Enabled = False
Label21.Caption = "0"
grd7.Clear
grd7.Rows = 1
grd7.Cols = 3
Set kb = CreateObject("Excel.application")
kb.Workbooks.Open (App.Path & "\NomEtudiants.xls")
kb.Workbooks("NomEtudiants").Close savechanges:=False
End Sub

Private Sub Command14_Click()
'On Error Resume Next
Dim c As Double
Dim r As Double
Dim i As Double
Dim nr As Double
Dim nc As Double
Dim cl1 As String
Dim cl2 As String
nr = grd8.Rows
nc = grd8.Cols
c = 2
Call cont
Do While Not em.EOF
cl1 = em!cla
For i = 2 To nc - 1
grd8.Col = i
grd8.row = 0
cl2 = grd8.Text
If cl1 = cl2 Then
grd8.Col = i
grd8.row = 1
grd8.Text = em!col11
grd8.row = 2
grd8.Text = em!col12
grd8.row = 3
grd8.Text = em!col13
grd8.row = 4
grd8.Text = em!col14
grd8.row = 5
grd8.Text = em!col21
grd8.row = 6
grd8.Text = em!col22
grd8.row = 7
grd8.Text = em!col23
grd8.row = 8
grd8.Text = em!col24
grd8.row = 9
grd8.Text = em!col31
grd8.row = 10
grd8.Text = em!col32
grd8.row = 11
grd8.Text = em!col33
grd8.row = 12
grd8.Text = em!col34
grd8.row = 13
grd8.Text = em!col41
grd8.row = 14
grd8.Text = em!col42
grd8.row = 15
grd8.Text = em!col43
grd8.row = 16
grd8.Text = em!col44
grd8.row = 17
grd8.Text = em!col51
grd8.row = 18
grd8.Text = em!col52
grd8.row = 19
grd8.Text = em!col53
grd8.row = 20
grd8.Text = em!col54
grd8.row = 21
grd8.Text = em!col61
grd8.row = 22
grd8.Text = em!col62
grd8.row = 23
grd8.Text = em!col63
grd8.row = 24
grd8.Text = em!col64
grd8.row = 25
grd8.Text = em!col71
grd8.row = 26
grd8.Text = em!col72
grd8.row = 27
grd8.Text = em!col73
grd8.row = 28
grd8.Text = em!col74
i = nc
End If
Next i
em.MoveNext
Loop
End Sub

Private Sub Command15_Click()
'On Error GoTo u
Dim i As Double
Dim j As Double
Dim n As Double
Dim k As Double
Dim d As Double
Dim sd As Double
FileCopy App.Path & "\emplois010.xls", App.Path & "\Emplois de Temps.xls"
Command15.Enabled = False
n = grd8.Cols
Set kb = CreateObject("Excel.application")
kb.Workbooks.Open (App.Path & "\Emplois de Temps.xls")
kb.Visible = True
For i = 0 To 28
For j = 0 To n - 1
grd8.row = i
grd8.Col = j
kb.Workbooks("Emplois de Temps").Sheets(1).Cells(i + 3, j + 1).Value = grd8.Text
Next j
Next i
'kb.Workbooks("ReportAnuelle").Sheets(1).Range("L3").Value = Combo5.Text
'kb.Workbooks("fiche de presences").Sheets(1).Range("B5").Value = DT11.Value
'Workbooks("Etudiants").Close savechanges:=False
'Worksheets(1).Activate
Command15.Enabled = True
Exit Sub
u:
MsgBox "ÌÃ» «€·«ﬁ ’›Õ… «ﬂ”· «‰ ﬂ«‰  „› ÊÕ…", vbExclamation
Command10.Enabled = True
End Sub

Private Sub Command16_Click()
Dim n As Double
Dim i As Double
Command16.Enabled = False
Call cont
Do While Not em.EOF
em.Delete
em.MoveNext
Loop
n = grd8.Cols
For i = 2 To n - 1
em.AddNew
grd8.Col = i
grd8.row = 0
em!cla = grd8.Text
grd8.row = 1
em!col11 = grd8.Text
grd8.row = 2
em!col12 = grd8.Text
grd8.row = 3
em!col13 = grd8.Text
grd8.row = 4
em!col14 = grd8.Text
grd8.row = 5
em!col21 = grd8.Text
grd8.row = 6
em!col22 = grd8.Text
grd8.row = 7
em!col23 = grd8.Text
grd8.row = 8
em!col24 = grd8.Text
grd8.row = 9
em!col31 = grd8.Text
grd8.row = 10
em!col32 = grd8.Text
grd8.row = 11
em!col33 = grd8.Text
grd8.row = 12
em!col34 = grd8.Text
grd8.row = 13
em!col41 = grd8.Text
grd8.row = 14
em!col42 = grd8.Text
grd8.row = 15
em!col43 = grd8.Text
grd8.row = 16
em!col44 = grd8.Text
grd8.row = 17
em!col51 = grd8.Text
grd8.row = 18
em!col52 = grd8.Text
grd8.row = 19
em!col53 = grd8.Text
grd8.row = 20
em!col54 = grd8.Text
grd8.row = 21
em!col61 = grd8.Text
grd8.row = 22
em!col62 = grd8.Text
grd8.row = 23
em!col63 = grd8.Text
grd8.row = 24
em!col64 = grd8.Text
grd8.row = 25
em!col71 = grd8.Text
grd8.row = 26
em!col72 = grd8.Text
grd8.row = 27
em!col73 = grd8.Text
grd8.row = 28
em!col74 = grd8.Text
em.Update
Next i
MsgBox " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ", vbInformation
Command16.Enabled = True
End Sub

Private Sub Command2_Click()
On Error Resume Next
If Label2.Caption = "" Then
MsgBox "ﬁ„ »«·÷€ÿ ⁄·Ï «·»Ì«‰«  «·„—«œ Õ–›Â«", vbCritical
Exit Sub
End If
g = MsgBox("Â·  —Ìœ «·Õ–› Õﬁ« ", vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
Call cont
Do While Not cl.EOF
If Label2.Caption = cl!aut Then
cl.Delete
Timer1.Enabled = True
Exit Sub
End If
cl.MoveNext
Loop
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
Text1.Text = ""
Text1.SetFocus
Label2.Caption = ""
grd1.Visible = False
grd1.Clear
grd1.Rows = 1
Call chargegrd1
Call chargec1
grd1.Visible = True
ProgressBar1.Value = 0
Timer1.Enabled = False
Call charge_combo3

End Sub

Private Sub Command4_Click()
On Error Resume Next
Text3.Text = ""
Text2.Text = ""
Text2.SetFocus
Label4.Caption = ""
grd2.Visible = False
grd2.Clear
grd2.Rows = 1
Call chargegrd2
grd2.Visible = True
ProgressBar2.Value = 0
Timer2.Enabled = False

End Sub

Private Sub Command5_Click()
On Error Resume Next
If Label4.Caption = "" Then
MsgBox "ﬁ„ »«·÷€ÿ ⁄·Ï «·»Ì«‰«  «·„—«œ Õ–›Â«", vbCritical
Exit Sub
End If
g = MsgBox("Â·  —Ìœ «·Õ–› Õﬁ« ", vbInformation + vbYesNo, "AGEP6")
If g = vbYes Then
Call cont
Do While Not mt.EOF
If Label4.Caption = mt!aut Then
mt.Delete
Timer2.Enabled = True
Exit Sub
End If
mt.MoveNext
Loop
End If

End Sub

Private Sub Command6_Click()
On Error Resume Next
Text3.Text = Trim(Text3.Text)
Text2.Text = Trim(Text2.Text)
If Combo1.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·ﬁ”„", vbCritical
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "«œŒ· «·„«œ…", vbCritical
Text2.SetFocus
Exit Sub
End If
If Text3.Text = "" Then
MsgBox "«œŒ· «·÷«—»", vbCritical
Text3.SetFocus
Exit Sub
End If
If Text4.Text = "" Then
MsgBox "«œŒ· ⁄œœ „Ê«œ Â–« «·ﬁ”„", vbCritical
Text4.SetFocus
Exit Sub
End If
If Label4.Caption = "" Then
If grd2.Rows > Val(Text4.Text) Then
MsgBox "⁄œœ «·„Ê«œ «·„”„ÊÕ »Â «‰ ÂÏ", vbCritical
Exit Sub
End If
End If
If Val(Text4.Text) > 20 Then
MsgBox "⁄œœ «·„Ê«œ «·„”„ÊÕ »Â ··ﬁ”„ «·Ê«Õœ ÂÊ 20 „«œ… ›ﬁÿ", vbCritical
Exit Sub
End If
If Label4.Caption <> "" Then
Call cont
Do While Not mt.EOF
If Label4.Caption <> mt!aut And Combo1.Text = mt!cla And Text2.Text = mt!mat Then
MsgBox " „ ÕÃ“ Â–Â «·„«œ… ”«»ﬁ«", vbCritical
Exit Sub
End If
mt.MoveNext
Loop
If Label4.Caption <> "" Then
Call cont2
Do While Not be.EOF
If be!cla = Combo1.Text Then
If be!mat1 = Label18.Caption Then
be!mat1 = Text2.Text
be.Update
ElseIf be!mat2 = Label18.Caption Then
be!mat2 = Text2.Text
be.Update
ElseIf be!mat3 = Label18.Caption Then
be!mat3 = Text2.Text
be.Update
ElseIf be!mat4 = Label18.Caption Then
be!mat4 = Text2.Text
be.Update
ElseIf be!mat5 = Label18.Caption Then
be!mat5 = Text2.Text
be.Update
ElseIf be!mat6 = Label18.Caption Then
be!mat6 = Text2.Text
be.Update
ElseIf be!mat7 = Label18.Caption Then
be!mat7 = Text2.Text
be.Update
ElseIf be!mat8 = Label18.Caption Then
be!mat8 = Text2.Text
be.Update
ElseIf be!mat9 = Label18.Caption Then
be!mat9 = Text2.Text
be.Update
ElseIf be!mat10 = Label18.Caption Then
be!mat10 = Text2.Text
be.Update
ElseIf be!mat11 = Label18.Caption Then
be!mat11 = Text2.Text
be.Update
ElseIf be!mat12 = Label18.Caption Then
be!mat12 = Text2.Text
be.Update
ElseIf be!mat13 = Label18.Caption Then
be!mat13 = Text2.Text
be.Update
ElseIf be!mat14 = Label18.Caption Then
be!mat14 = Text2.Text
be.Update
ElseIf be!mat15 = Label18.Caption Then
be!mat15 = Text2.Text
be.Update
ElseIf be!mat16 = Label18.Caption Then
be!mat16 = Text2.Text
be.Update
ElseIf be!mat17 = Label18.Caption Then
be!mat17 = Text2.Text
be.Update
ElseIf be!mat18 = Label18.Caption Then
be!mat18 = Text2.Text
be.Update
ElseIf be!mat19 = Label18.Caption Then
be!mat19 = Text2.Text
be.Update
ElseIf be!mat20 = Label18.Caption Then
be!mat20 = Text2.Text
be.Update
End If
End If
be.MoveNext
Loop
End If
Call cont
Do While Not mt.EOF
If Label4.Caption = mt!aut Then
mt!cla = Combo1.Text
mt!mat = Text2.Text
mt!cof = Text3.Text
mt!nbr = Text4.Text
mt.Update
Timer2.Enabled = True
Exit Sub
End If
mt.MoveNext
Loop
End If
Call cont
Do While Not mt.EOF
If Combo1.Text = mt!cla And Text2.Text = mt!mat Then
MsgBox " „ ÕÃ“ Â–Â «·„«œ… ”«»ﬁ«", vbCritical
Exit Sub
End If
mt.MoveNext
Loop
mt.AddNew
mt!cla = Combo1.Text
mt!mat = Text2.Text
mt!cof = Text3.Text
mt!nbr = Text4.Text
mt!nme = Label17.Caption
mt.Update
Timer2.Enabled = True

End Sub

Private Sub Command7_Click()
On Error GoTo u
Dim i As Double
Dim j As Double
Dim k As Double
Dim r As Double
Dim n As Double
Dim m As Double
Dim row As Integer
Dim ro As String
Dim roo As String
Dim r1 As Double
'On Error Resume Next
If Combo2.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·ﬁ”„", vbCritical
Exit Sub
End If
If grd3.Rows < 2 Then
MsgBox "·«  ÊÃœ »Ì«‰« ", vbCritical
Exit Sub
End If
Command7.Enabled = False
FileCopy App.Path & "\emp00000.doc", App.Path & "\Emplois.doc"
k = grd3.Rows - 1
r = "1"
Set kb = CreateObject("word.application")
'kb.Visible = True
kb.Documents.Open (App.Path & "\Emplois.doc")
kb.ActiveDocument.Bookmarks("etab").Select
kb.Selection.InsertAfter Label15.Caption
kb.ActiveDocument.Bookmarks("ann").Select
kb.Selection.InsertAfter Label12.Caption
kb.ActiveDocument.Bookmarks("cla").Select
kb.Selection.InsertAfter Combo2.Text
n = 0
i = 0
'For j = 1 To 3
'n = n + 1
'grd9.row = i
'grd9.col = j
'kb.ActiveDocument.Bookmarks(a + n).Select
'kb.Selection.InsertAfter grd9.Text
'Next j
n = 0
For i = 1 To k
For j = 0 To 7
n = n + 1
grd3.row = i
grd3.Col = j
kb.ActiveDocument.Bookmarks(a + n).Select
kb.Selection.InsertAfter grd3.Text
Next j
'n = n + 3
Next i
kb.Visible = True
Command7.Enabled = True
Exit Sub
u:
MsgBox "ÌÃ» «€·«ﬁ ’›Õ… ÊÊ—œ «‰ ﬂ«‰  „› ÊÕ…", vbExclamation
Command7.Enabled = True


End Sub

Private Sub Command8_Click()
On Error Resume Next
If Combo1.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·ﬁ”„", vbCritical
Exit Sub
End If
g = InputBox("«œŒ· «·⁄œœ", "⁄œœ „Ê«œ «·ﬁ”„")
If g = Cancel Or Val(g) = 0 Or Val(g) < 0 Then
Exit Sub
End If
If Val(g) < (Val(grd2.Rows) - 1) Then
MsgBox "·· „ﬂ‰ „‰ Ê÷⁄ Â–« «·⁄œœ ÌÃ» Õ–› »⁄÷ «·„Ê«œ «·„ÊÃÊœ…", vbCritical
Exit Sub
End If
If Val(g) > 20 Then
MsgBox "⁄œœ «·„Ê«œ «·„”„ÊÕ »Â ··ﬁ”„ «·Ê«Õœ ÂÊ 20 „«œ… ›ﬁÿ", vbCritical
Exit Sub
End If
Call cont
Do While Not mt.EOF
If Combo1.Text = mt!cla Then
mt!nbr = g
mt.Update
End If
mt.MoveNext
Loop
Text4.Text = g
End Sub

Private Sub Command9_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim n As Double
Dim k As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim d As String
Dim m As String
Dim y As String
Dim x As String
j = 0
grd7.Clear
grd7.Rows = 1
grd7.Cols = 5
Set kb = CreateObject("Excel.application")
kb.Workbooks.Open (App.Path & "\NomEtudiants.xls")
'j = kb.Workbooks("NomEtudiants").Sheets(1).Cells(3, 7)
'Label21.Caption = j
'j = kb.Workbooks("NomEtudiants").Sheets(1).Cells(4, 7)
'Label22.Caption = j
'Exit Sub
For i = 3 To 65534
j = kb.Workbooks("NomEtudiants").Sheets(1).Cells(i, 9)
If j = 0 Then
Label21.Caption = (i - 3)
i = 65534
End If
Next i
n = Label21.Caption
Command9.Enabled = False
grd7.Clear
grd7.Rows = n + 5
grd7.Cols = 5
grd7.Visible = False
grd7.ColWidth(0) = 1600
grd7.ColWidth(1) = 1600
grd7.ColWidth(2) = 4000
grd7.ColWidth(3) = 1600
grd7.ColWidth(4) = 1600
grd7.ColAlignment(0) = 3
grd7.ColAlignment(1) = 3
grd7.ColAlignment(2) = 3
grd7.ColAlignment(3) = 3
grd7.ColAlignment(4) = 3
j = -1
For i = 2 To n + 3
j = j + 1
For k = 0 To 4
grd7.row = j
grd7.Col = k
grd7.Text = kb.Workbooks("NomEtudiants").Sheets(1).Cells(i, k + 2)
Next k
Next i
grd7.Rows = j
kb.Workbooks("NomEtudiants").Close savechanges:=False
grd7.Visible = True
Command9.Enabled = False
Command11.Enabled = True

End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Top = 100
Skin1.LoadSkin App.Path & "\green.skn"
Skin1.ApplySkin Me.hWnd
grd1.Clear
grd1.Cols = 3
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 3000
grd1.ColWidth(2) = 1000
Call chargegrd1
Call chargec1
grd2.Clear
grd2.Cols = 3
grd2.Rows = 1
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 3000
grd2.ColWidth(2) = 1000
Call chargegrd8
End Sub
Private Sub chargegrd1()
'On Error Resume Next
Dim i As Double
Dim tx As String
Dim j As Double
grd1.row = 0
grd1.Col = 0
grd1.Text = ""
grd1.Col = 1
grd1.Text = "«·√ﬁ”«„"
grd1.Col = 2
grd1.Text = "«·Õ«·…"
i = 1
grd6.Cols = 1
grd6.Rows = 1
grd6.row = 0
grd6.Col = 0
grd6.Text = "«·√ﬁ”«„"
grd8.Clear
grd8.Cols = 2
grd8.Rows = 29
grd8.ColWidth(0) = 1200
grd8.ColWidth(1) = 800
'**** «·≈À‰Ì‰
grd8.row = 1
grd8.Col = 0
grd8.Text = "Lundi"
grd8.Col = 1
grd8.Text = "08-10"
grd8.CellBackColor = &H80FF&
grd8.row = 2
grd8.Col = 0
grd8.Text = """"
grd8.Col = 1
grd8.Text = "10-12"
grd8.CellBackColor = &HC000&
grd8.row = 3
grd8.Col = 0
grd8.Text = """"
grd8.Col = 1
grd8.Text = "12-14"
grd8.CellBackColor = &HFF80FF
grd8.row = 4
grd8.Col = 0
grd8.Text = """"
grd8.Col = 1
grd8.Text = "14-16"
grd8.CellBackColor = &HC000&
'**** «·À·«À«¡
grd8.row = 5
grd8.Col = 0
grd8.Text = "Mardi"
grd8.Col = 1
grd8.Text = "08-10"
grd8.CellBackColor = &H80FF&
grd8.row = 6
grd8.Col = 0
grd8.Text = """"
grd8.Col = 1
grd8.Text = "10-12"
grd8.CellBackColor = &HC000&
grd8.row = 7
grd8.Col = 0
grd8.Text = """"
grd8.Col = 1
grd8.Text = "12-14"
grd8.CellBackColor = &HFF80FF
grd8.row = 8
grd8.Col = 0
grd8.Text = """"
grd8.Col = 1
grd8.Text = "14-16"
grd8.CellBackColor = &HC000&
'**** «·√—»⁄«¡
grd8.row = 9
grd8.Col = 0
grd8.Text = "Mercredi"
grd8.Col = 1
grd8.Text = "08-10"
grd8.CellBackColor = &H80FF&
grd8.row = 10
grd8.Col = 0
grd8.Text = """"
grd8.Col = 1
grd8.Text = "10-12"
grd8.CellBackColor = &HC000&
grd8.row = 11
grd8.Col = 0
grd8.Text = """"
grd8.Col = 1
grd8.Text = "12-14"
grd8.CellBackColor = &HFF80FF
grd8.row = 12
grd8.Col = 0
grd8.Text = """"
grd8.Col = 1
grd8.Text = "14-16"
grd8.CellBackColor = &HC000&
'**** «·Œ„Ì”
grd8.row = 13
grd8.Col = 0
grd8.Text = "Jeudi"
grd8.Col = 1
grd8.Text = "08-10"
grd8.CellBackColor = &H80FF&
grd8.row = 14
grd8.Col = 0
grd8.Text = """"
grd8.Col = 1
grd8.Text = "10-12"
grd8.CellBackColor = &HC000&
grd8.row = 15
grd8.Col = 0
grd8.Text = """"
grd8.Col = 1
grd8.Text = "12-14"
grd8.CellBackColor = &HFF80FF
grd8.row = 16
grd8.Col = 0
grd8.Text = """"
grd8.Col = 1
grd8.Text = "14-16"
grd8.CellBackColor = &HC000&
'**** «·Ã„⁄…
grd8.row = 17
grd8.Col = 0
grd8.Text = "Vendredi"
grd8.Col = 1
grd8.Text = "08-10"
grd8.CellBackColor = &H80FF&
grd8.row = 18
grd8.Col = 0
grd8.Text = """"
grd8.Col = 1
grd8.Text = "10-12"
grd8.CellBackColor = &HC000&
grd8.row = 19
grd8.Col = 0
grd8.Text = """"
grd8.Col = 1
grd8.Text = "12-14"
grd8.CellBackColor = &HFF80FF
grd8.row = 20
grd8.Col = 0
grd8.Text = """"
grd8.Col = 1
grd8.Text = "14-16"
grd8.CellBackColor = &HC000&
'**** «·”» 
grd8.row = 21
grd8.Col = 0
grd8.Text = "Samedi"
grd8.Col = 1
grd8.Text = "08-10"
grd8.CellBackColor = &H80FF&
grd8.row = 22
grd8.Col = 0
grd8.Text = """"
grd8.Col = 1
grd8.Text = "10-12"
grd8.CellBackColor = &HC000&
grd8.row = 23
grd8.Col = 0
grd8.Text = """"
grd8.Col = 1
grd8.Text = "12-14"
grd8.CellBackColor = &HFF80FF
grd8.row = 24
grd8.Col = 0
grd8.Text = """"
grd8.Col = 1
grd8.Text = "14-16"
grd8.CellBackColor = &HC000&
'**** «·√Õœ
grd8.row = 25
grd8.Col = 0
grd8.Text = "Dimanche"
grd8.Col = 1
grd8.Text = "08-10"
grd8.CellBackColor = &H80FF&
grd8.row = 26
grd8.Col = 0
grd8.Text = """"
grd8.Col = 1
grd8.Text = "10-12"
grd8.CellBackColor = &HC000&
grd8.row = 27
grd8.Col = 0
grd8.Text = """"
grd8.Col = 1
grd8.Text = "12-14"
grd8.CellBackColor = &HFF80FF
grd8.row = 28
grd8.Col = 0
grd8.Text = """"
grd8.Col = 1
grd8.Text = "14-16"
grd8.CellBackColor = &HC000&
'For j = 1 To 4
Call cont
grd1.Rows = cl.RecordCount + 2
grd6.Rows = cl.RecordCount + 2
grd8.Cols = cl.RecordCount + 2
Do While Not cl.EOF
grd1.row = i
grd1.Col = 0
grd1.Text = cl!aut
grd1.Col = 1
grd1.Text = cl!cla
If cl!act = "1" Then
grd1.Col = 2
grd1.Text = " ‘€Ì·"
Else
grd1.Col = 2
grd1.Text = " ⁄ÿÌ·"
End If
grd6.row = i
grd6.Col = 0
grd6.Text = cl!cla
grd8.Col = i + 1
grd8.ColWidth(i + 1) = 700
grd8.row = 0
grd8.Text = cl!cla
For j = 1 To 28
grd8.row = j
grd8.Col = i
tx = grd8.CellBackColor
grd8.row = j
grd8.Col = i + 1
grd8.CellBackColor = tx
Next j
i = i + 1
cl.MoveNext
Loop
Command14_Click
'Next j
grd1.Rows = i
grd1.ColAlignment(1) = 1
grd1.Col = 1
grd1.Sort = 1
grd6.Rows = i
grd6.ColAlignment(0) = 3
grd6.Col = 1
grd6.Sort = 1
grd8.row = 1
grd6.Sort = 1
End Sub
Private Sub chargegrd2()
On Error Resume Next
Dim i As Double
grd2.row = 0
grd2.Col = 0
grd2.Text = ""
grd2.Col = 1
grd2.Text = "«·„«œ…"
grd2.Col = 2
grd2.Text = "«·÷«—»"
i = 1
'For j = 1 To 4
Call cont
grd2.Rows = mt.RecordCount + 2
Do While Not mt.EOF
If mt!cla = Combo1.Text Then
Text4.Text = mt!nbr
Text4.Enabled = False
Command8.Enabled = True
grd2.row = i
grd2.Col = 0
grd2.Text = mt!aut
grd2.Col = 1
grd2.Text = mt!mat
grd2.Col = 2
grd2.Text = mt!cof
i = i + 1
End If
mt.MoveNext
Loop
'Next j
grd2.Rows = i
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.Col = 1
grd2.Sort = 1
End Sub

Private Sub grd1_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim tx As String
i = grd1.row
j = grd1.Col
grd1.row = i
grd1.Col = 0
Label2.Caption = grd1.Text
grd1.Col = 1
Text1.Text = grd1.Text
grd1.Col = 1
Label3.Caption = grd1.Text
If j = 2 Then
grd1.row = i
grd1.Col = 2
tx = grd1.Text
g = MsgBox("Â·  —Ìœ Õﬁ« ≈·€«¡ " + tx + " «·ﬁ”„ " + Text1.Text, vbInformation + vbYesNo + arabic, "AGEP6")
If g = vbYes Then
Call cont
Do While Not et.EOF
If et!cla = Text1.Text Then
If tx = " ‘€Ì·" Then
et!act = "0"
et.Update
Else
et!act = "1"
et.Update
End If
End If
et.MoveNext
Loop
Call cont
Do While Not cl.EOF
If Label2.Caption = cl!aut Then
If tx = " ‘€Ì·" Then
cl!act = "0"
cl.Update
Else
cl!act = "1"
cl.Update
End If
Timer1.Enabled = True
Exit Sub
End If
cl.MoveNext
Loop
Else
Text1.Text = ""
Text1.SetFocus
Label2.Caption = ""
End If
End If
End Sub

Private Sub grd2_Click()
On Error Resume Next
Dim i As Double
i = grd2.row
grd2.row = i
grd2.Col = 0
Label4.Caption = grd2.Text
grd2.Col = 1
Text2.Text = grd2.Text
grd2.Col = 1
Label18.Caption = grd2.Text
grd2.Col = 2
Text3.Text = grd2.Text

End Sub

Private Sub grd3_Click()
On Error Resume Next
Dim c As Double
Dim r As Double
Dim n As Double
If Combo2.Text = "" Then
MsgBox "ﬁ„ »«Œ Ì«— «·ﬁ”„", vbCritical
Exit Sub
End If
c = grd3.Col
r = grd3.row
If c > 0 And r > 0 Then
g = InputBox("«œŒ· «·„«œ…", "«œŒ«· «·„«œ…")
If g = Cancel Then
Exit Sub
End If
grd3.Col = c
grd3.row = r
grd3.Text = g
Call cont
Do While Not em.EOF
If em!cla = Combo2.Text Then
grd3.row = 1
grd3.Col = 1
em!j18 = grd3.Text
grd3.Col = 2
em!j28 = grd3.Text
grd3.Col = 3
em!j38 = grd3.Text
grd3.Col = 4
em!j48 = grd3.Text
grd3.Col = 5
em!j58 = grd3.Text
grd3.Col = 6
em!j68 = grd3.Text
grd3.Col = 7
em!j78 = grd3.Text
grd3.row = 2
grd3.Col = 1
em!j19 = grd3.Text
grd3.Col = 2
em!j29 = grd3.Text
grd3.Col = 3
em!j39 = grd3.Text
grd3.Col = 4
em!j49 = grd3.Text
grd3.Col = 5
em!j59 = grd3.Text
grd3.Col = 6
em!j69 = grd3.Text
grd3.Col = 7
em!j79 = grd3.Text
grd3.row = 3
grd3.Col = 1
em!j110 = grd3.Text
grd3.Col = 2
em!j210 = grd3.Text
grd3.Col = 3
em!j310 = grd3.Text
grd3.Col = 4
em!j410 = grd3.Text
grd3.Col = 5
em!j510 = grd3.Text
grd3.Col = 6
em!j610 = grd3.Text
grd3.Col = 7
em!j710 = grd3.Text
grd3.row = 4
grd3.Col = 1
em!j111 = grd3.Text
grd3.Col = 2
em!j211 = grd3.Text
grd3.Col = 3
em!j311 = grd3.Text
grd3.Col = 4
em!j411 = grd3.Text
grd3.Col = 5
em!j511 = grd3.Text
grd3.Col = 6
em!j611 = grd3.Text
grd3.Col = 7
em!j711 = grd3.Text
grd3.row = 5
grd3.Col = 1
em!j112 = grd3.Text
grd3.Col = 2
em!j212 = grd3.Text
grd3.Col = 3
em!j312 = grd3.Text
grd3.Col = 4
em!j412 = grd3.Text
grd3.Col = 5
em!j512 = grd3.Text
grd3.Col = 6
em!j612 = grd3.Text
grd3.Col = 7
em!j712 = grd3.Text
grd3.row = 6
grd3.Col = 1
em!j113 = grd3.Text
grd3.Col = 2
em!j213 = grd3.Text
grd3.Col = 3
em!j313 = grd3.Text
grd3.Col = 4
em!j413 = grd3.Text
grd3.Col = 5
em!j513 = grd3.Text
grd3.Col = 6
em!j613 = grd3.Text
grd3.Col = 7
em!j713 = grd3.Text
grd3.row = 7
grd3.Col = 1
em!j114 = grd3.Text
grd3.Col = 2
em!j214 = grd3.Text
grd3.Col = 3
em!j314 = grd3.Text
grd3.Col = 4
em!j414 = grd3.Text
grd3.Col = 5
em!j514 = grd3.Text
grd3.Col = 6
em!j614 = grd3.Text
grd3.Col = 7
em!j714 = grd3.Text
grd3.row = 8
grd3.Col = 1
em!j115 = grd3.Text
grd3.Col = 2
em!j215 = grd3.Text
grd3.Col = 3
em!j315 = grd3.Text
grd3.Col = 4
em!j415 = grd3.Text
grd3.Col = 5
em!j515 = grd3.Text
grd3.Col = 6
em!j615 = grd3.Text
grd3.Col = 7
em!j715 = grd3.Text
grd3.row = 9
grd3.Col = 1
em!j116 = grd3.Text
grd3.Col = 2
em!j216 = grd3.Text
grd3.Col = 3
em!j316 = grd3.Text
grd3.Col = 4
em!j416 = grd3.Text
grd3.Col = 5
em!j516 = grd3.Text
grd3.Col = 6
em!j616 = grd3.Text
grd3.Col = 7
em!j716 = grd3.Text
grd3.row = 10
grd3.Col = 1
em!j117 = grd3.Text
grd3.Col = 2
em!j217 = grd3.Text
grd3.Col = 3
em!j317 = grd3.Text
grd3.Col = 4
em!j417 = grd3.Text
grd3.Col = 5
em!j517 = grd3.Text
grd3.Col = 6
em!j617 = grd3.Text
grd3.Col = 7
em!j717 = grd3.Text
grd3.row = 11
grd3.Col = 1
em!j118 = grd3.Text
grd3.Col = 2
em!j218 = grd3.Text
grd3.Col = 3
em!j318 = grd3.Text
grd3.Col = 4
em!j418 = grd3.Text
grd3.Col = 5
em!j518 = grd3.Text
grd3.Col = 6
em!j618 = grd3.Text
grd3.Col = 7
em!j718 = grd3.Text
grd3.row = 12
grd3.Col = 1
em!j119 = grd3.Text
grd3.Col = 2
em!j219 = grd3.Text
grd3.Col = 3
em!j319 = grd3.Text
grd3.Col = 4
em!j419 = grd3.Text
grd3.Col = 5
em!j519 = grd3.Text
grd3.Col = 6
em!j619 = grd3.Text
grd3.Col = 7
em!j719 = grd3.Text
grd3.row = 13
grd3.Col = 1
em!j120 = grd3.Text
grd3.Col = 2
em!j220 = grd3.Text
grd3.Col = 3
em!j320 = grd3.Text
grd3.Col = 4
em!j420 = grd3.Text
grd3.Col = 5
em!j520 = grd3.Text
grd3.Col = 6
em!j620 = grd3.Text
grd3.Col = 7
em!j720 = grd3.Text
grd3.row = 14
grd3.Col = 1
em!j121 = grd3.Text
grd3.Col = 2
em!j221 = grd3.Text
grd3.Col = 3
em!j321 = grd3.Text
grd3.Col = 4
em!j421 = grd3.Text
grd3.Col = 5
em!j521 = grd3.Text
grd3.Col = 6
em!j621 = grd3.Text
grd3.Col = 7
em!j721 = grd3.Text
grd3.row = 15
grd3.Col = 1
em!j122 = grd3.Text
grd3.Col = 2
em!j222 = grd3.Text
grd3.Col = 3
em!j322 = grd3.Text
grd3.Col = 4
em!j422 = grd3.Text
grd3.Col = 5
em!j522 = grd3.Text
grd3.Col = 6
em!j622 = grd3.Text
grd3.Col = 7
em!j722 = grd3.Text
grd3.row = 16
grd3.Col = 1
em!j123 = grd3.Text
grd3.Col = 2
em!j223 = grd3.Text
grd3.Col = 3
em!j323 = grd3.Text
grd3.Col = 4
em!j423 = grd3.Text
grd3.Col = 5
em!j523 = grd3.Text
grd3.Col = 6
em!j623 = grd3.Text
grd3.Col = 7
em!j723 = grd3.Text
em.Update
Exit Sub
End If
em.MoveNext
Loop
em.AddNew
em!cla = Combo2.Text
grd3.row = 1
grd3.Col = 1
em!j18 = grd3.Text
grd3.Col = 2
em!j28 = grd3.Text
grd3.Col = 3
em!j38 = grd3.Text
grd3.Col = 4
em!j48 = grd3.Text
grd3.Col = 5
em!j58 = grd3.Text
grd3.Col = 6
em!j68 = grd3.Text
grd3.Col = 7
em!j78 = grd3.Text
grd3.row = 2
grd3.Col = 1
em!j19 = grd3.Text
grd3.Col = 2
em!j29 = grd3.Text
grd3.Col = 3
em!j39 = grd3.Text
grd3.Col = 4
em!j49 = grd3.Text
grd3.Col = 5
em!j59 = grd3.Text
grd3.Col = 6
em!j69 = grd3.Text
grd3.Col = 7
em!j79 = grd3.Text
grd3.row = 3
grd3.Col = 1
em!j110 = grd3.Text
grd3.Col = 2
em!j210 = grd3.Text
grd3.Col = 3
em!j310 = grd3.Text
grd3.Col = 4
em!j410 = grd3.Text
grd3.Col = 5
em!j510 = grd3.Text
grd3.Col = 6
em!j610 = grd3.Text
grd3.Col = 7
em!j710 = grd3.Text
grd3.row = 4
grd3.Col = 1
em!j111 = grd3.Text
grd3.Col = 2
em!j211 = grd3.Text
grd3.Col = 3
em!j311 = grd3.Text
grd3.Col = 4
em!j411 = grd3.Text
grd3.Col = 5
em!j511 = grd3.Text
grd3.Col = 6
em!j611 = grd3.Text
grd3.Col = 7
em!j711 = grd3.Text
grd3.row = 5
grd3.Col = 1
em!j112 = grd3.Text
grd3.Col = 2
em!j212 = grd3.Text
grd3.Col = 3
em!j312 = grd3.Text
grd3.Col = 4
em!j412 = grd3.Text
grd3.Col = 5
em!j512 = grd3.Text
grd3.Col = 6
em!j612 = grd3.Text
grd3.Col = 7
em!j712 = grd3.Text
grd3.row = 6
grd3.Col = 1
em!j113 = grd3.Text
grd3.Col = 2
em!j213 = grd3.Text
grd3.Col = 3
em!j313 = grd3.Text
grd3.Col = 4
em!j413 = grd3.Text
grd3.Col = 5
em!j513 = grd3.Text
grd3.Col = 6
em!j613 = grd3.Text
grd3.Col = 7
em!j713 = grd3.Text
grd3.row = 7
grd3.Col = 1
em!j114 = grd3.Text
grd3.Col = 2
em!j214 = grd3.Text
grd3.Col = 3
em!j314 = grd3.Text
grd3.Col = 4
em!j414 = grd3.Text
grd3.Col = 5
em!j514 = grd3.Text
grd3.Col = 6
em!j614 = grd3.Text
grd3.Col = 7
em!j714 = grd3.Text
grd3.row = 8
grd3.Col = 1
em!j115 = grd3.Text
grd3.Col = 2
em!j215 = grd3.Text
grd3.Col = 3
em!j315 = grd3.Text
grd3.Col = 4
em!j415 = grd3.Text
grd3.Col = 5
em!j515 = grd3.Text
grd3.Col = 6
em!j615 = grd3.Text
grd3.Col = 7
em!j715 = grd3.Text
grd3.row = 9
grd3.Col = 1
em!j116 = grd3.Text
grd3.Col = 2
em!j216 = grd3.Text
grd3.Col = 3
em!j316 = grd3.Text
grd3.Col = 4
em!j416 = grd3.Text
grd3.Col = 5
em!j516 = grd3.Text
grd3.Col = 6
em!j616 = grd3.Text
grd3.Col = 7
em!j716 = grd3.Text
grd3.row = 10
grd3.Col = 1
em!j117 = grd3.Text
grd3.Col = 2
em!j217 = grd3.Text
grd3.Col = 3
em!j317 = grd3.Text
grd3.Col = 4
em!j417 = grd3.Text
grd3.Col = 5
em!j517 = grd3.Text
grd3.Col = 6
em!j617 = grd3.Text
grd3.Col = 7
em!j717 = grd3.Text
grd3.row = 11
grd3.Col = 1
em!j118 = grd3.Text
grd3.Col = 2
em!j218 = grd3.Text
grd3.Col = 3
em!j318 = grd3.Text
grd3.Col = 4
em!j418 = grd3.Text
grd3.Col = 5
em!j518 = grd3.Text
grd3.Col = 6
em!j618 = grd3.Text
grd3.Col = 7
em!j718 = grd3.Text
grd3.row = 12
grd3.Col = 1
em!j119 = grd3.Text
grd3.Col = 2
em!j219 = grd3.Text
grd3.Col = 3
em!j319 = grd3.Text
grd3.Col = 4
em!j419 = grd3.Text
grd3.Col = 5
em!j519 = grd3.Text
grd3.Col = 6
em!j619 = grd3.Text
grd3.Col = 7
em!j719 = grd3.Text
grd3.row = 13
grd3.Col = 1
em!j120 = grd3.Text
grd3.Col = 2
em!j220 = grd3.Text
grd3.Col = 3
em!j320 = grd3.Text
grd3.Col = 4
em!j420 = grd3.Text
grd3.Col = 5
em!j520 = grd3.Text
grd3.Col = 6
em!j620 = grd3.Text
grd3.Col = 7
em!j720 = grd3.Text
grd3.row = 14
grd3.Col = 1
em!j121 = grd3.Text
grd3.Col = 2
em!j221 = grd3.Text
grd3.Col = 3
em!j321 = grd3.Text
grd3.Col = 4
em!j421 = grd3.Text
grd3.Col = 5
em!j521 = grd3.Text
grd3.Col = 6
em!j621 = grd3.Text
grd3.Col = 7
em!j721 = grd3.Text
grd3.row = 15
grd3.Col = 1
em!j122 = grd3.Text
grd3.Col = 2
em!j222 = grd3.Text
grd3.Col = 3
em!j322 = grd3.Text
grd3.Col = 4
em!j422 = grd3.Text
grd3.Col = 5
em!j522 = grd3.Text
grd3.Col = 6
em!j622 = grd3.Text
grd3.Col = 7
em!j722 = grd3.Text
grd3.row = 16
grd3.Col = 1
em!j123 = grd3.Text
grd3.Col = 2
em!j223 = grd3.Text
grd3.Col = 3
em!j323 = grd3.Text
grd3.Col = 4
em!j423 = grd3.Text
grd3.Col = 5
em!j523 = grd3.Text
grd3.Col = 6
em!j623 = grd3.Text
grd3.Col = 7
em!j723 = grd3.Text
em.Update
End If
End Sub

Private Sub grd8_Click()
'On Error Resume Next
Dim c As Double
Dim r As Double
c = grd8.Col
r = grd8.row
If c > 1 And r > 0 And r < 29 Then
g = InputBox("«œŒ· «·„«œ…", "«œŒ«· «·„«œ…")
If g = Cancel Then
Exit Sub
End If
grd8.Col = c
grd8.row = r
grd8.Text = g
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim i As Integer
Dim j As Integer
Dim n As Double
Dim vg As String
Text3.Text = Trim(Text3.Text)
n = Len(Text3.Text)
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
vg = Mid$(Text3.Text, i, 1)
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

Private Sub Text4_KeyPress(KeyAscii As Integer)
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
Text1.Text = ""
Text1.SetFocus
Label2.Caption = ""
grd1.Visible = False
grd1.Clear
grd1.Rows = 1
Call chargegrd1
Call chargec1
grd1.Visible = True
ProgressBar1.Value = 0
Timer1.Enabled = False
Call charge_combo3
End If

End Sub
Public Sub chargec1()
On Error Resume Next
Call cont
Label15.Caption = sr!eco
Label11.Caption = sr!eco
Label14.Caption = sr!ann
Label12.Caption = sr!ann
Combo1.Clear
Combo2.Clear
  Do While Not cl.EOF
  If cl!act = "1" Then
    Combo1.AddItem cl!cla
Combo2.AddItem cl!cla
End If
cl.MoveNext
  Loop
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
ProgressBar2.Value = ProgressBar2.Value + 8
If ProgressBar2.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
Text3.Text = ""
Text2.Text = ""
Text2.SetFocus
Label4.Caption = ""
grd2.Visible = False
grd2.Clear
grd2.Rows = 1
Call chargegrd2
grd2.Visible = True
ProgressBar2.Value = 0
Timer2.Enabled = False
End If

End Sub
Private Sub chargetimes()
'On Error Resume Next
grd3.Clear
grd3.Rows = 17
grd3.Cols = 8
grd3.row = 0
grd3.Col = 0
grd3.Text = "«· ÊﬁÌ "
grd3.Col = 1
grd3.Text = "«·√Õœ"
grd3.Col = 2
grd3.Text = "«·≈À‰Ì‰"
grd3.Col = 3
grd3.Text = "«·À·«À«¡"
grd3.Col = 4
grd3.Text = "«·√—»⁄«¡"
grd3.Col = 5
grd3.Text = "«·Œ„Ì”"
grd3.Col = 6
grd3.Text = "«·Ã„⁄…"
grd3.Col = 7
grd3.Text = "«·”» "
grd3.Col = 0
grd3.row = 1
grd3.Text = "08-09"
grd3.row = 2
grd3.Text = "09-10"
grd3.row = 3
grd3.Text = "10-11"
grd3.row = 4
grd3.Text = "11-12"
grd3.row = 5
grd3.Text = "12-13"
grd3.row = 6
grd3.Text = "13-14"
grd3.row = 7
grd3.Text = "14-15"
grd3.row = 8
grd3.Text = "15-16"
grd3.row = 9
grd3.Text = "16-17"
grd3.row = 10
grd3.Text = "17-18"
grd3.row = 11
grd3.Text = "18-19"
grd3.row = 12
grd3.Text = "19-20"
grd3.row = 13
grd3.Text = "20-21"
grd3.row = 14
grd3.Text = "21-22"
grd3.row = 15
grd3.Text = "22-23"
grd3.row = 16
grd3.Text = "23-00"
Call cont
Do While Not em.EOF
If Combo2.Text = em!cla Then
grd3.row = 1
grd3.Col = 1
grd3.Text = em!j18
grd3.Col = 2
grd3.Text = em!j28
grd3.Col = 3
grd3.Text = em!j38
grd3.Col = 4
grd3.Text = em!j48
grd3.Col = 5
grd3.Text = em!j58
grd3.Col = 6
grd3.Text = em!j68
grd3.Col = 7
grd3.Text = em!j78
grd3.row = 2
grd3.Col = 1
grd3.Text = em!j19
grd3.Col = 2
grd3.Text = em!j29
grd3.Col = 3
grd3.Text = em!j39
grd3.Col = 4
grd3.Text = em!j49
grd3.Col = 5
grd3.Text = em!j59
grd3.Col = 6
grd3.Text = em!j69
grd3.Col = 7
grd3.Text = em!j79
grd3.row = 3
grd3.Col = 1
grd3.Text = em!j110
grd3.Col = 2
grd3.Text = em!j210
grd3.Col = 3
grd3.Text = em!j310
grd3.Col = 4
grd3.Text = em!j410
grd3.Col = 5
grd3.Text = em!j510
grd3.Col = 6
grd3.Text = em!j610
grd3.Col = 7
grd3.Text = em!j710
grd3.row = 4
grd3.Col = 1
grd3.Text = em!j111
grd3.Col = 2
grd3.Text = em!j211
grd3.Col = 3
grd3.Text = em!j311
grd3.Col = 4
grd3.Text = em!j411
grd3.Col = 5
grd3.Text = em!j511
grd3.Col = 6
grd3.Text = em!j611
grd3.Col = 7
grd3.Text = em!j711
grd3.row = 5
grd3.Col = 1
grd3.Text = em!j112
grd3.Col = 2
grd3.Text = em!j212
grd3.Col = 3
grd3.Text = em!j312
grd3.Col = 4
grd3.Text = em!j412
grd3.Col = 5
grd3.Text = em!j512
grd3.Col = 6
grd3.Text = em!j612
grd3.Col = 7
grd3.Text = em!j712
grd3.row = 6
grd3.Col = 1
grd3.Text = em!j113
grd3.Col = 2
grd3.Text = em!j213
grd3.Col = 3
grd3.Text = em!j313
grd3.Col = 4
grd3.Text = em!j413
grd3.Col = 5
grd3.Text = em!j513
grd3.Col = 6
grd3.Text = em!j613
grd3.Col = 7
grd3.Text = em!j713
grd3.row = 7
grd3.Col = 1
grd3.Text = em!j114
grd3.Col = 2
grd3.Text = em!j214
grd3.Col = 3
grd3.Text = em!j314
grd3.Col = 4
grd3.Text = em!j414
grd3.Col = 5
grd3.Text = em!j514
grd3.Col = 6
grd3.Text = em!j614
grd3.Col = 7
grd3.Text = em!j714
grd3.row = 8
grd3.Col = 1
grd3.Text = em!j115
grd3.Col = 2
grd3.Text = em!j215
grd3.Col = 3
grd3.Text = em!j315
grd3.Col = 4
grd3.Text = em!j415
grd3.Col = 5
grd3.Text = em!j515
grd3.Col = 6
grd3.Text = em!j615
grd3.Col = 7
grd3.Text = em!j715
grd3.row = 9
grd3.Col = 1
grd3.Text = em!j116
grd3.Col = 2
grd3.Text = em!j216
grd3.Col = 3
grd3.Text = em!j316
grd3.Col = 4
grd3.Text = em!j416
grd3.Col = 5
grd3.Text = em!j516
grd3.Col = 6
grd3.Text = em!j616
grd3.Col = 7
grd3.Text = em!j716
grd3.row = 10
grd3.Col = 1
grd3.Text = em!j117
grd3.Col = 2
grd3.Text = em!j217
grd3.Col = 3
grd3.Text = em!j317
grd3.Col = 4
grd3.Text = em!j417
grd3.Col = 5
grd3.Text = em!j517
grd3.Col = 6
grd3.Text = em!j617
grd3.Col = 7
grd3.Text = em!j717
grd3.row = 11
grd3.Col = 1
grd3.Text = em!j118
grd3.Col = 2
grd3.Text = em!j218
grd3.Col = 3
grd3.Text = em!j318
grd3.Col = 4
grd3.Text = em!j418
grd3.Col = 5
grd3.Text = em!j518
grd3.Col = 6
grd3.Text = em!j618
grd3.Col = 7
grd3.Text = em!j718
grd3.row = 12
grd3.Col = 1
grd3.Text = em!j119
grd3.Col = 2
grd3.Text = em!j219
grd3.Col = 3
grd3.Text = em!j319
grd3.Col = 4
grd3.Text = em!j419
grd3.Col = 5
grd3.Text = em!j519
grd3.Col = 6
grd3.Text = em!j619
grd3.Col = 7
grd3.Text = em!j719
grd3.row = 13
grd3.Col = 1
grd3.Text = em!j120
grd3.Col = 2
grd3.Text = em!j220
grd3.Col = 3
grd3.Text = em!j320
grd3.Col = 4
grd3.Text = em!j420
grd3.Col = 5
grd3.Text = em!j520
grd3.Col = 6
grd3.Text = em!j620
grd3.Col = 7
grd3.Text = em!j720
grd3.row = 14
grd3.Col = 1
grd3.Text = em!j121
grd3.Col = 2
grd3.Text = em!j221
grd3.Col = 3
grd3.Text = em!j321
grd3.Col = 4
grd3.Text = em!j421
grd3.Col = 5
grd3.Text = em!j521
grd3.Col = 6
grd3.Text = em!j621
grd3.Col = 7
grd3.Text = em!j721
grd3.row = 15
grd3.Col = 1
grd3.Text = em!j122
grd3.Col = 2
grd3.Text = em!j222
grd3.Col = 3
grd3.Text = em!j322
grd3.Col = 4
grd3.Text = em!j422
grd3.Col = 5
grd3.Text = em!j522
grd3.Col = 6
grd3.Text = em!j622
grd3.Col = 7
grd3.Text = em!j722
grd3.row = 16
grd3.Col = 1
grd3.Text = em!j123
grd3.Col = 2
grd3.Text = em!j223
grd3.Col = 3
grd3.Text = em!j323
grd3.Col = 4
grd3.Text = em!j423
grd3.Col = 5
grd3.Text = em!j523
grd3.Col = 6
grd3.Text = em!j623
grd3.Col = 7
grd3.Text = em!j723
Exit Sub
End If
em.MoveNext
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
Private Sub chargegrd8()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim tx1 As String
Dim tx2 As String
n = grd8.Cols
Call cont
Do While Not em.EOF
tx1 = em!cla
For i = 2 To n - 1
grd8.Col = i
grd8.row = 0
tx2 = grd8.Text
If tx1 = tx2 Then
grd8.Col = i
grd8.row = 1
grd8.Text = em!sa1
grd8.row = 2
grd8.Text = em!sa2
grd8.row = 3
grd8.Text = em!sa3
grd8.row = 4
grd8.Text = em!di1
grd8.row = 5
grd8.Text = em!di2
grd8.row = 6
grd8.Text = em!di3
grd8.row = 7
grd8.Text = em!lu1
grd8.row = 8
grd8.Text = em!lu2
grd8.row = 9
grd8.Text = em!lu3
grd8.row = 10
grd8.Text = em!ma1
grd8.row = 11
grd8.Text = em!ma2
grd8.row = 12
grd8.Text = em!ma3
grd8.row = 13
grd8.Text = em!me1
grd8.row = 14
grd8.Text = em!me2
grd8.row = 15
grd8.Text = em!me3
grd8.row = 16
grd8.Text = em!je1
grd8.row = 17
grd8.Text = em!je2
grd8.row = 18
grd8.Text = em!je3
End If
Next i
em.MoveNext
Loop


End Sub
