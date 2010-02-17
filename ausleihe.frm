VERSION 5.00
Begin VB.Form ausleihe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Neue Ausliehe"
   ClientHeight    =   4215
   ClientLeft      =   5265
   ClientTop       =   1440
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   11190
   Begin VB.CommandButton Command3 
      Caption         =   "neue FP-Gruppe"
      Height          =   375
      Left            =   9120
      TabIndex        =   17
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   855
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   1560
      Width           =   7335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   9120
      TabIndex        =   14
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ausleihe eintragen"
      Height          =   375
      Left            =   9120
      TabIndex        =   13
      Top             =   120
      Width           =   1695
   End
   Begin VB.ComboBox Combo4 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6000
      TabIndex        =   10
      Text            =   "max@mustermann.de"
      Top             =   600
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6000
      TabIndex        =   9
      Text            =   "Max Mustermann"
      Top             =   120
      Width           =   2895
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   1080
      Width           =   3255
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   360
      TabIndex        =   0
      Top             =   2640
      Width           =   10455
   End
   Begin VB.Label Label5 
      Caption         =   "Bemerkungen"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "FP-Gruppe"
      Height          =   255
      Left            =   5040
      TabIndex        =   12
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "eMail"
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Name"
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Protokoll"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Fach"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   855
   End
   Begin VB.Label T 
      Caption         =   "Typ"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "ausleihe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conn As ADODB.Connection
Dim serverconn, servernr, UID, PWD, DBname, S As String
Dim typs, fachs, prüfs, fps, fpname, fpemail As String
Dim typid, fachid, prtid, fpid, ausid, fpakt As Integer

Sub cnct_conn()
 
 Set Conn = New ADODB.Connection
    servernr = "188.40.97.252"
    DBname = "paadb"
    UID = "paauser"
    PWD = "paauser.190761"
    serverconn = "Server=" + servernr + "; UID=" + UID + "; PWD=" + PWD + ";" + "database=" + DBname + "; Option=16387"
    Conn.ConnectionString = "Provider=MSDASQL;" + "DRIVER={MySQL ODBC 5.1 Driver};" + serverconn
End Sub

Sub combos()
 
 Dim paafach, paafach2, paatyp, paaprot, paaprot2, paafp As ADODB.Recordset
 Set paatyp = New ADODB.Recordset
 
 typid = 0
 
 If Combo1.Text <> "" Then
 paatyp.Open "SELECT * FROM paa_typ WHERE (Name = '" + Combo1.Text + "')", Conn
 paatyp.MoveFirst
 typid = paatyp!lfdTypNr
 If Combo1.Text = "FP" Then Combo4.Enabled = True Else Combo4.Enabled = False
 End If
 
 
 Set paafach = New ADODB.Recordset
 If typid <> 0 Then
 paafach.Open "SELECT * FROM paa_fach WHERE Typ =" + Str(typid), Conn
 
 If Not paafach.EOF Then
  paafach.MoveFirst
  If Combo2.Text = "" Then
  Combo2.Clear
  While Not paafach.EOF
   Combo2.AddItem (paafach!Name)
   paafach.MoveNext
  Wend
  End If
 Else
  Combo2.Clear
  Combo2.AddItem ("Keine Fächer vorhanden.")
 End If
 End If
 
 fachs = Combo2.Text
 prüfs = Combo3.Text
  
 If fachs <> "" Then
  Set paafach2 = New ADODB.Recordset
  paafach2.Open "SELECT * FROM paa_fach WHERE (Name = '" + fachs + "') AND (typ = " + Str(typid) + ")", Conn
  If Not paafach2.EOF Then
  paafach2.MoveFirst
  fachid = paafach2!lfdFachNr
  paafach2.Close
 
 
 Set paaprot = New ADODB.Recordset
 paaprot.Open "SELECT * FROM paa_protokolle WHERE Fach =" + Str(fachid), Conn
 
 If paaprot.EOF Then
 Combo3.AddItem ("Keine Protokolle vorhanden.")
 Else
  paaprot.MoveFirst
  If Combo3.Text = "" Then
  Combo3.Clear
  While Not paaprot.EOF
   Combo3.AddItem (paaprot!prüfer)
   paaprot.MoveNext
  Wend
  End If
  End If
  End If
 End If
 
 typs = Combo1.Text
 
 If prüfs <> "" Then
  Set paaprot2 = New ADODB.Recordset
  paaprot2.Open "SELECT * FROM paa_protokolle WHERE (fach = " + Str(fachid) + ") AND (prüfer = '" + prüfs + "') AND (typ = " + Str(typid) + ")", Conn
  If Not paaprot2.EOF Then
  paaprot2.MoveFirst
  prtid = paaprot2!lfdProtokollNr
 End If
 'MsgBox (Str(prtid))
 paaprot2.Close
 End If
 
 fps = Combo4.Text
End Sub

Sub show_prot()
Dim slct_str As String
Dim paaprot, paaaus As ADODB.Recordset
Set paaprot = New ADODB.Recordset

slct_str = "SELECT * FROM paa_protokolle WHERE "

If typid <> 0 Then slct_str = slct_str + "(typ = " + Str(typid) + ") AND "
If fachid <> 0 Then slct_str = slct_str + "(Fach = " + Str(fachid) + ") AND "
If Combo3.Text <> "" Then slct_str = slct_str + "(prüfer = '" + Combo3.Text + "') AND "

slct_str = Mid(slct_str, 1, Len(slct_str) - 5)
'MsgBox (slct_str)

paaprot.Open slct_str, Conn, adOpenDynamic, adLockBatchOptimistic
If Not paaprot.EOF Then
 paaprot.MoveFirst
 List1.Clear
 While Not paaprot.EOF
 If paaprot!aktAusleihe = 0 Then
  List1.AddItem ("Protokollordner " + paaprot!prüfer + " " + fachs + " - nicht ausgeliehen.")
 Else
  Set paaaus = New ADODB.Recordset
  paaaus.Open "SELECT * FROM paa_ausleihe WHERE lfdAusleiheNr = " + Str(paaprot!aktAusleihe), Conn
  paaaus.MoveFirst
  List1.AddItem ("Protokollordner " + paaprot!prüfer + " " + fachs + " Ausgeliehen am " + Str(paaaus!Datum))
  paaaus.Close
 End If
  paaprot.MoveNext
 Wend
 Else
  List1.Clear
  List1.AddItem ("Keine Protokolle vorhanden.")
 End If
 
End Sub
Private Sub Combo1_Click()
 fachid = 0
 prtid = 0
 
 Combo2.Clear
 Combo3.Clear
 combos
 show_prot
End Sub


Private Sub Combo2_Click()
 prtid = 0
 Combo3.Clear
 combos
 show_prot
End Sub

Private Sub Combo3_Click()
 combos
 show_prot
End Sub

Sub fp_ausleihe()
 Dim paafp, paaprot, paaaus As ADODB.Recordset
 Set paafp = New ADODB.Recordset
 
 If Combo4.Text <> "" Then
    fpid = Combo4.ItemData(Combo4.ListIndex)
    paafp.Open "SELECT * FROM paa_fp WHERE lfdFpNr = " + Str(fpid), Conn, adOpenKeyset, adLockBatchOptimistic
    If Not paafp.EOF Then
     paafp.MoveFirst
     fpname = paafp!Name
     fpemail = paafp!email
     fpakt = paafp!aktAusleihe
    End If
    paafp.Close
    'MsgBox (fpname + " " + fpemail)
 Else
    fpakt = 0
    fpid = 0
 End If
 
 Set paaaus = New ADODB.Recordset
 paaaus.Open "SELECT * FROM paa_ausleihe ORDER BY  'lfdAusleiheNr'", Conn, adOpenKeyset, adLockBatchOptimistic
 
 Set paaprot = New ADODB.Recordset
 paaprot.Open "SELECT * FROM paa_protokolle WHERE lfdProtokollNr = " + Str(prtid), Conn, adOpenDynamic, adLockBatchOptimistic
 
 If Not paaprot.EOF Then
  paaprot.MoveFirst
  List1.AddItem ("Protokoll gefunden. Nr.: " + Str(paaprot!lfdProtokollNr))
  If paaprot!aktAusleihe <> 0 Or fpakt <> 0 Then MsgBox ("Das gewählte Protokoll wurde bereits am datum ausgeliehen.")
  If (paaprot!aktAusleihe = 0) And fpakt = 0 Then
   List1.AddItem ("Protokoll noch vorhanden.")
    If fpname <> "" And fpemail <> "" Then
     'List1.AddItem ("Name und email vorhanden")
     paaaus.MoveLast
     paaaus.AddNew
     paaaus!fpgruppe = fpid
     paaaus!Name = fpname
     paaaus!email = fpemail
     paaaus!Datum = Now
     paaaus!protokollnr = paaprot!lfdProtokollNr 'prtid
     paaaus!bemerkung = Text3.Text
     paaaus.UpdateBatch
     paaaus.Close
     
     Set paaaus = New ADODB.Recordset
     paaaus.Open "SELECT * FROM paa_ausleihe WHERE (protokollnr = " + Str(prtid) + ") AND (zurück = 0) ORDER BY  'lfdAusleiheNr'", Conn, adOpenKeyset, adLockBatchOptimistic
     paaaus.MoveLast
     ausid = paaaus!lfdAusleiheNr
     List1.AddItem ("Ausleihe hinzugefügt. Ausleihenummer ist " + Str(paaaus!lfdAusleiheNr))
     
     
     paaprot.Update
     paaprot!aktAusleihe = paaaus!lfdAusleiheNr
     paaprot.UpdateBatch
     paaaus.Close
     
     paafp.Open
     paafp.MoveFirst
     paafp.Update
     paafp!aktAusleihe = ausid
     paafp.UpdateBatch
     paafp.Close
     
    Else
     MsgBox ("Ungültiger Name oder eMail.")
    End If
   End If
 Else: MsgBox ("Es konnte kein Protokoll zum Ausleihen gefunden werden.")
 End If
 
 paaprot.Close
 
End Sub


Private Sub Command1_Click()
List1.Clear
If typid = 3 Then
    fp_ausleihe
Else: subausleihe
End If
'show_prot
End Sub

Sub subausleihe()
 Dim paaaus, paaprot As ADODB.Recordset
 
 Set paaaus = New ADODB.Recordset
 paaaus.Open "SELECT * FROM paa_ausleihe", Conn, adOpenKeyset, adLockBatchOptimistic
 
 Set paaprot = New ADODB.Recordset
 paaprot.Open "SELECT * FROM paa_protokolle WHERE lfdProtokollNr = " + Str(prtid), Conn, adOpenDynamic, adLockBatchOptimistic
 
 If Not paaprot.EOF Then
  paaprot.MoveFirst
  List1.AddItem ("Protokoll gefunden. Nr.: " + Str(paaprot!lfdProtokollNr))
  If paaprot!aktAusleihe <> 0 Then MsgBox ("Das gewählte Protokoll wurde bereits am datum ausgeliehen.")
  If (paaprot!aktAusleihe = 0) Then
   List1.AddItem ("Protokoll noch vorhanden.")
   'MsgBox (Str(typid))
    If Text1.Text <> "" And Text2.Text <> "" Then
     'List1.AddItem ("Name und email vorhanden")
     paaaus.AddNew
     paaaus!fpgruppe = fpid
     paaaus!Name = Text1.Text
     paaaus!email = Text2.Text
     paaaus!Datum = Now
     paaaus!protokollnr = paaprot!lfdProtokollNr 'prtid
     paaaus!bemerkung = Text3.Text
     paaaus.UpdateBatch
     paaaus.Close
     
     Set paaaus = New ADODB.Recordset
     paaaus.Open "SELECT * FROM paa_ausleihe WHERE (protokollnr = " + Str(prtid) + ") AND (zurück = 0) ORDER BY  'lfdAusleiheNr'", Conn, adOpenKeyset, adLockBatchOptimistic
     paaaus.MoveLast
     ausid = paaaus!lfdAusleiheNr
     List1.AddItem ("Ausleihe hinzugefügt. Ausleihenummer ist " + Str(paaaus!lfdAusleiheNr))
     
     paaprot.Update
     paaprot!aktAusleihe = paaaus!lfdAusleiheNr
     paaprot.UpdateBatch
    Else
     MsgBox ("Ungültiger Name oder eMail.")
    End If
   End If
 Else: MsgBox ("Es konnte kein Protokoll zum Ausleihen gefunden werden.")
 End If
 
 paaaus.Close
 paaprot.Close
End Sub

Private Sub Command2_Click()
 Unload ausleihe
End Sub

Private Sub Command3_Click()
 Dim paafp As ADODB.Recordset
 Set paafp = New ADODB.Recordset
 
 fpakt = 0
 paafp.Open "SELECT * FROM paa_fp", Conn, adOpenDynamic, adLockBatchOptimistic
 
 If Combo4.Text = "" Then
  paafp.AddNew
  paafp!Name = Text1.Text
  paafp!email = Text2.Text
  paafp.UpdateBatch
 End If
 paafp.Close
 
 paafp.Open
 If Not paafp.EOF Then
 paafp.MoveFirst
 Combo4.Clear
 While Not paafp.EOF
  Combo4.AddItem (Str(paafp!lfdFpNr) + " - " + paafp!Name)
  Combo4.ItemData(Combo4.NewIndex) = paafp!lfdFpNr
  paafp.MoveNext
 Wend
 Else
 Combo4.Clear
 Combo4.AddItem ("keine FP-Gruppen vorhanden.")
 End If
 paafp.Close
End Sub

Private Sub Form_Load()

fpid = 0
cnct_conn

Conn.Open

Dim paatyp As ADODB.Recordset
 Set paatyp = New ADODB.Recordset
 
 paatyp.Open "SELECT * FROM paa_typ", Conn, adOpenDynamic, adLockBatchOptimistic
 paatyp.MoveFirst
 Combo1.Clear
 While Not paatyp.EOF
  Combo1.AddItem (paatyp!Name)
  paatyp.MoveNext
 Wend
 
Dim paafp As ADODB.Recordset
 Set paafp = New ADODB.Recordset
 
 paafp.Open "SELECT * FROM paa_fp", Conn, adOpenDynamic, adLockBatchOptimistic
 
 If Not paafp.EOF Then
 paafp.MoveFirst
 Combo4.Clear
 While Not paafp.EOF
  Combo4.AddItem (Str(paafp!lfdFpNr) + " - " + paafp!Name)
  Combo4.ItemData(Combo4.NewIndex) = paafp!lfdFpNr
  paafp.MoveNext
 Wend
 Else
 Combo4.Clear
 Combo4.AddItem ("keine FP-Gruppen vorhanden.")
 End If
 paafp.Close
End Sub


Private Sub Form_Unload(Cancel As Integer)
 Conn.Close
End Sub

