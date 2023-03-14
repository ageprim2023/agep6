VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CODE 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AGEP"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8085
   Icon            =   "codes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   4455
      Left            =   1320
      ScaleHeight     =   4395
      ScaleWidth      =   5235
      TabIndex        =   22
      Top             =   6480
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   360
         TabIndex        =   25
         Text            =   "Text3"
         Top             =   600
         Width           =   4695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   360
         TabIndex        =   23
         Text            =   "Text2"
         Top             =   960
         Width           =   4695
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   4335
      End
      Begin VB.Label Label19 
         Caption         =   "Label19"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   2040
         Width           =   4695
      End
      Begin VB.Label Label18 
         Caption         =   "Label18"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1680
         Width           =   4695
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   -120
      OleObjectBlob   =   "codes.frx":324A
      Top             =   3000
   End
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7815
      Begin VB.PictureBox Picture1 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   0  'None
         Height          =   6255
         Left            =   240
         ScaleHeight     =   6255
         ScaleWidth      =   7335
         TabIndex        =   2
         Top             =   240
         Width           =   7335
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   720
            ScaleHeight     =   375
            ScaleWidth      =   5895
            TabIndex        =   29
            Top             =   5640
            Visible         =   0   'False
            Width           =   5895
            Begin VB.TextBox Text4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               IMEMode         =   3  'DISABLE
               Left            =   120
               PasswordChar    =   "*"
               TabIndex        =   30
               Top             =   0
               Width           =   3135
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "ﬂ·„… «·”— «·Œ«’… »«·„»—„Ã"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3120
               TabIndex        =   31
               Top             =   0
               Width           =   2655
            End
         End
         Begin VB.CommandButton Command3 
            Caption         =   "«·”„«Õ »«” Œœ«„ «·»—‰«„Ã ⁄·Ï Â–« «·ÃÂ«“"
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   15.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   26
            Top             =   5640
            Width           =   6855
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Œ—ÊÃ"
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   15.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2640
            TabIndex        =   6
            Top             =   5040
            Width           =   2055
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   240
            TabIndex        =   0
            Top             =   1680
            Width           =   6855
         End
         Begin VB.CommandButton Command1 
            Caption         =   " ›⁄Ì·"
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   15.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2640
            TabIndex        =   3
            Top             =   2160
            Width           =   2055
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Version 6.0"
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   4440
            TabIndex        =   21
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·‰”Œ… 6.0"
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   1320
            TabIndex        =   20
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "(00222)22660920"
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   3720
            TabIndex        =   19
            Top             =   3000
            Width           =   2535
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "›Ì Õ«·… «„ ·«ﬂﬂ„ ·–·ﬂ «·ﬂÊœ «·—Ã«¡ «œŒ«·Â ›Ì «·Œ«‰… «”›·Â À„ «·÷€ÿ ⁄·Ï  ›⁄Ì·"
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   1320
            Width           =   6735
         End
         Begin VB.Shape Shape1 
            BorderWidth     =   3
            Height          =   735
            Left            =   240
            Top             =   3720
            Width           =   6855
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   " ·Ì“Êœﬂ„ „‰ Œ·«·Â »ﬂÊœ «·”„«Õ »«” Œœ«„ «·»—‰«„Ã ⁄·Ï Â–« «·ÃÂ«“  Ê‘ﬂ—«"
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   4560
            Width           =   6855
         End
         Begin VB.Label Label13 
            Caption         =   "Label13"
            Height          =   255
            Left            =   360
            TabIndex        =   16
            Top             =   6720
            Width           =   4335
         End
         Begin VB.Label Label12 
            Caption         =   "Label12"
            Height          =   255
            Left            =   360
            TabIndex        =   15
            Top             =   6480
            Width           =   4335
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            Height          =   255
            Left            =   4680
            TabIndex        =   14
            Top             =   6720
            Width           =   2415
         End
         Begin VB.Label Label10 
            Caption         =   "Label10"
            Height          =   255
            Left            =   360
            TabIndex        =   13
            Top             =   6240
            Width           =   4335
         End
         Begin VB.Label Label99 
            Caption         =   "Label9"
            Height          =   255
            Left            =   4680
            TabIndex        =   12
            Top             =   6480
            Width           =   2415
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   " À„ «⁄ÿ«∆Â «·—„“ «·Ÿ«Â— «„«„ﬂ„ «”›·Â"
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   3360
            Width           =   6855
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "(00222)33440920"
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1080
            TabIndex        =   10
            Top             =   3000
            Width           =   2535
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   20.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   495
            Left            =   240
            TabIndex        =   9
            Top             =   3840
            Width           =   6855
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "›Ì Õ«·… ⁄œ„ «„ ·«ﬂﬂ„ ·–·ﬂ «·ﬂÊœ «Ê ÷Ì«⁄Â „‰ﬂ„ Ì—ÃÏ «·« ’«· »√Õœ «·√—ﬁ«„ «· «·Ì…"
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   2640
            Width           =   6735
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "»—‰«„Ã  ”ÌÌ— «·„ƒ””«  «·Œ«’… («·„œ«—” «·Õ—…)‹ "
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   20.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   7095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "AGEP"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1680
            TabIndex        =   5
            Top             =   -120
            Width           =   3975
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   " ÌÃ» «„ ·«ﬂﬂ„ ·ﬂÊœ «·”„«Õ »«” Œœ«„ «·»—‰«„Ã ⁄·Ï Â–« «·ÃÂ«“"
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   960
            Width           =   6735
         End
      End
   End
End
Attribute VB_Name = "CODE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nbdec As Double
Dim nbhex As String
Dim op As Double
Private Function Hex2Dec(ByVal paraHexValue As String) As String

    Dim LC As Long
    Dim curHexValue As String
    Dim curChar As String
    Dim curValue As String
    
    curHexValue = UCase(Trim(paraHexValue))
    If Left(curHexValue, 2) = "&H" Then curHexValue = Mid(curHexValue, 3)
    For LC = 1 To Len(curHexValue)
        curChar = Mid(curHexValue, LC, 1)
        If InStr(1, nptodec, curChar) <> 0 Then curChar = Asc(curChar) - 55
        curValue = Val(curValue) + (Val(curChar) * (16 ^ (Len(curHexValue) - LC)))
    Next LC
    Hex2Dec = Val(curValue)
End Function

Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "«·—Ã«¡ «œŒ«· —ﬁ„ «· ›⁄Ì·", vbCritical
Text1.SetFocus
Exit Sub
End If
AHMED$ = "1"
x = Text1.Text
y = UCase(Left(x, Len(x)))
Text1.Text = y
' «· ⁄—› ⁄·Ï «·ﬂÊœ «·„œŒ· «·ÂÌﬂ”Ì œ”Ì „«·
If Text1.Text = Label19.Caption Then
AHMED$ = Label18.Caption
End If
If AHMED$ = Label18.Caption Then
'√Œ– ﬁÌ„… «·ÃÂ«“ «·„Œ“‰…
SaveSetting "A", "0", "RunCount", AHMED$
retvalue = GetSetting("A", "0", "Runcount")
MsgBox "«·ﬂÊœ ’ÕÌÕ ... „ «·”„«Õ ·ﬂ„ »«” Œœ«„ «·»—‰«„Ã ⁄·Ï Â–« «·ÃÂ«“", vbInformation
'MDIForm1.Caption = "  »—‰«„Ã  ”ÌÌ— «·„œ«—” «·Õ—…  «·‰”Œ… 5.0 "
start.Show
Unload Me
Exit Sub
Else
MsgBox "«·ﬂÊœ «·–Ì «œŒ·  €Ì— ’ÕÌÕ", vbCritical
Text1.Text = ""
Text1.SetFocus
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Picture3.Visible = True
Text4.SetFocus
End Sub

Private Sub Form_Load()
'On Error Resume Next
Dim l As Double
Dim i As Double
Dim k As Double
Dim a As String
Dim obj_FSO As Object, obj_Drive As Object
 Set obj_FSO = CreateObject("Scripting.FileSystemObject")
 Set obj_Drive = obj_FSO.GetDrive("c:\")
 '√Œ– ”Ì—Ì«· ”Ì
Text3.Text = obj_Drive.SerialNumber
'  ÕÊÌ· ”Ì—Ì«· «·Ï ﬁÌ„…
Text2.Text = Val(Text3.Text)
' ÕÊÌ· ”Ì—Ì«· «·Ï ﬁÌ„… «ÌÃ«»Ì…
If Val(Text2.Text) < 0 Then
Text2.Text = Val(Text2.Text) * -1
Else
Text2.Text = Val(Text2.Text)
End If
'  ÕœÌœ ÿÊ· ”Ì—Ì«·
l = Len(Text2.Text)
' √Œ– 5 «—ﬁ«„ ›ﬁÿ
If l > 5 Then
For i = 1 To l
k = l - i
If k = 5 Then
a = Mid$(Text2.Text, i + 1, 5)
i = l
Label18.Caption = a
End If
Next i
Else
Label18.Caption = Text2.Text
End If
' √Œ– ﬁÌ„… Œ„”… √—ﬁ«„
nbdec = Label18.Caption
' √Ã—«¡ ⁄„·Ì… Õ”«»Ì… ⁄·Ï ﬁÌ„… 5 «—ﬁ«„
op = (((nbdec + 1) * 2) - 3)
' ÕÊÌ· «· «Ã «·Ï ÂÌﬂ”Ì œÌ”Ì„«·
nbhex = Hex(op)
' ⁄—÷ ÂÌﬂ”Ì œÌ” „«·
Label19.Caption = nbhex
' ⁄—÷ œÌ”Ì „«·
Label9.Caption = nbdec
retvalue = GetSetting("A", "0", "Runcount")
'√Œ– ﬁÌ„… «·ÃÂ«“ «·„Œ“‰…
AHMED$ = Val(retvalue)
SaveSetting "A", "0", "RunCount", AHMED$
Me.Left = 5000
Me.Top = 2000
Skin1.LoadSkin App.Path & "\green.skn"
Skin1.ApplySkin Me.hWnd
If AHMED$ = Label18.Caption Then
'MDIForm1.Caption = "  »—‰«„Ã  ”ÌÌ— «·„œ«—” «·Õ—…  «·‰”Œ… 5.0 "
start.Show
Unload Me
Exit Sub
End If
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If Text1.Text <> "" Then
If KeyCode = 13 Then
Command1_Click
End If
End If
End Sub

Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
Dim g As String
If Text4.Text <> "" Then
If KeyCode = 13 Then
g = Text4.Text
If g = "7415963153987654185274157415963" Then
AHMED$ = "1"
Text1.Text = Label19.Caption
x = Text1.Text
y = UCase(Left(x, Len(x)))
Text1.Text = y
AHMED$ = Label18.Caption
'√Œ– ﬁÌ„… «·ÃÂ«“ «·„Œ“‰…
SaveSetting "A", "0", "RunCount", AHMED$
retvalue = GetSetting("A", "0", "Runcount")
'MsgBox "«·ﬂÊœ ’ÕÌÕ ... „ «·”„«Õ ·ﬂ„ »«” Œœ«„ «·»—‰«„Ã ⁄·Ï Â–« «·ÃÂ«“", vbInformation
'MDIForm1.Caption = " «·‰”Œ… 5.0 »—‰«„Ã  ”ÌÌ— «·„œ«—” «·Õ—…"
start.Show
Unload Me
Exit Sub
End If
MsgBox "ﬂ·„… «·”— €Ì— ’ÕÌÕ…", vbExclamation
Text4.Text = ""
End If
End If
End Sub
