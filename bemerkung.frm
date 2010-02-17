VERSION 5.00
Begin VB.Form bemerkung 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bemerkung bearbeiten..."
   ClientHeight    =   2025
   ClientLeft      =   6015
   ClientTop       =   4950
   ClientWidth     =   6105
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "bemerkung.frx":0000
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "bemerkung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
 bearbeiten.bemerkungs = ""
 Unload bemerkung
End Sub

Private Sub Form_Load()
 Text1.Text = bearbeiten.bemerkungs
End Sub

Private Sub OKButton_Click()
 bearbeiten.bemerkungs = Text1.Text
 bearbeiten.write_bemerkung
 Unload bemerkung
End Sub
