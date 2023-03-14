VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ActiveSkin.ocx"
Begin VB.Form message 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1110
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   13065
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   13065
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4320
      OleObjectBlob   =   "message.frx":0000
      Top             =   360
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   12855
      TabIndex        =   1
      Top             =   120
      Width           =   12855
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "⁄›Ê« ...Â–« «·„” Œœ„ €Ì— „”„ÊÕ ·Â »œŒÊ· Â–Â «·’›Õ…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   -240
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   12975
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   2160
      Width           =   2655
   End
End
Attribute VB_Name = "message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
Me.Top = 3900
Skin1.LoadSkin App.Path & "\green.skn"
Skin1.ApplySkin Me.hWnd
End Sub

