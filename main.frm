VERSION 5.00
Begin VB.Form haupt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Protokollausleihe der freundlichen Ini-Physik"
   ClientHeight    =   1680
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   2175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   2175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Beenden"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ausleihe bearbeiten"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Neue Ausleihe"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "haupt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub Command1_Click()
 ausleihe.Left = haupt.Left + haupt.Width + 500
 ausleihe.Top = haupt.Top
 ausleihe.Show
 If bearbeiten.Visible Then Unload bearbeiten
End Sub


Private Sub Command2_Click()
 bearbeiten.Left = haupt.Left + haupt.Width + 500
 bearbeiten.Top = haupt.Top
 bearbeiten.Show
 If ausleihe.Visible Then Unload ausleihe
End Sub


Private Sub Command3_Click()
 If ausleihe.Visible Then Unload ausleihe
 If bearbeiten.Visible Then Unload bearbeiten
 End
End Sub

Private Sub Form_Resize()
 If ausleihe.Visible Then
 ausleihe.Left = haupt.Left + haupt.Width + 500
 ausleihe.Top = haupt.Top
 End If
 If bearbeiten.Visible Then
 bearbeiten.Left = haupt.Left + haupt.Width + 500
 bearbeiten.Top = haupt.Top
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 If ausleihe.Visible Then Unload ausleihe
 If bearbeiten.Visible Then Unload bearbeiten
 End
End Sub
