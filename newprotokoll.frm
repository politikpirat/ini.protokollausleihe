VERSION 5.00
Begin VB.Form newprotokoll 
   Caption         =   "Protokolle hinzufügen."
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   960
      TabIndex        =   7
      Text            =   "BSc"
      Top             =   840
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   6
      Text            =   "Dähne"
      Top             =   480
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Text            =   "Experimantalphysik"
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Datensatz einfügen..."
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Typ:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Prüfer:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   495
   End
   Begin VB.Label label1 
      Caption         =   "Fach:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "newprotokoll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 
End Sub

Private Sub label1_Click()

End Sub

Private Sub Text3_Change()

End Sub
