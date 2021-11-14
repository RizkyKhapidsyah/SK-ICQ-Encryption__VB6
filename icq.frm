VERSION 5.00
Begin VB.Form main 
   Caption         =   "ICQ V5 Encryption Method"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Information"
      Height          =   2175
      Left            =   0
      TabIndex        =   6
      Top             =   2520
      Width           =   6855
      Begin VB.Label Label5 
         Caption         =   $"icq.frx":0000
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   6615
      End
      Begin VB.Label Label4 
         Caption         =   $"icq.frx":00FA
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   6615
         WordWrap        =   -1  'True
      End
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Text            =   "Decrypted text appear here!"
      Top             =   2040
      Width           =   6735
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Text            =   "Encrypted text appear here!"
      Top             =   1320
      Width           =   6735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Enter text here!"
      Top             =   360
      Width           =   6735
   End
   Begin VB.Label Label3 
      Caption         =   "Output -- Decrypted Text from the encrypted packet"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1800
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "Output -- Encrypted Text:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Input Text"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()
    Dim TempTxt As String
    TempTxt = EncText(Text1.Text)
    Text2.Text = TempTxt
    Text3.Text = DecText(TempTxt)
End Sub
