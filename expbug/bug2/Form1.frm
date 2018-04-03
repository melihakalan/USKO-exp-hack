VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "kojd expdupe 1803"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4185
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2160
      Top             =   1560
   End
   Begin VB.ListBox List1 
      Height          =   2400
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0028
      TabIndex        =   7
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "stop"
      Height          =   315
      Left            =   3120
      TabIndex        =   6
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "start"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Text            =   "100"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ok"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Text            =   "knight online client"
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "kojd@windowslive.com"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "KOJD"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "www.onlinehile.com"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "interval"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
LoadOffsets
If AttachKO = False Then
Exit Sub
End If
Me.Show
KO_ADR_CHR = ReadLong(KO_PTR_CHR)
KO_ADR_DLG = ReadLong(KO_PTR_DLG)
End Sub

Private Sub Command2_Click()
Timer1.Interval = CInt(Text2)
Timer1.Enabled = True
End Sub

Private Sub Command3_Click()
Timer1.Enabled = False
End Sub

Private Sub List1_Click()
KO_SND_FNC = List1.Text
Notice "&&H" & Hex(KO_SND_FNC)
End Sub

Private Sub Timer1_Timer()
Dim pBytes() As Byte
Dim pStr As String
pStr = "2001232BFFFFFFFF"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes()
pStr = "640770080000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes()
pStr = "55001332343430365F4775617264736D616E2E6C7561FF"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes()
End Sub
