VERSION 5.00
Begin VB.Form bearbeiten 
   Caption         =   "Ausleihe bearbeiten"
   ClientHeight    =   5670
   ClientLeft      =   7140
   ClientTop       =   6075
   ClientWidth     =   9375
   DrawMode        =   1  'Blackness
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   9375
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   375
      Left            =   5760
      TabIndex        =   15
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Bemerkung"
      Height          =   375
      Left            =   7680
      TabIndex        =   14
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Schließen"
      Height          =   375
      Left            =   7680
      TabIndex        =   13
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "letztes FP"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "eigenes Protokoll"
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ordner zurück"
      Height          =   375
      Index           =   0
      Left            =   3840
      TabIndex        =   10
      Top             =   4560
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   3150
      Left            =   3840
      TabIndex        =   9
      Top             =   1320
      Width           =   5415
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   4125
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   3615
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   720
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   6000
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label6 
      Caption         =   "FP-Gruppe"
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   720
      Width           =   855
   End
   Begin VB.Label T 
      Caption         =   "Typ"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Fach"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Protokoll"
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "bearbeiten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conn As ADODB.Connection
Dim serverconn, servernr, UID, PWD, DBname, S As String
Dim typs, fachs, prüfs As String
Dim typid, fachid, prtid, ausid, fpid As Integer
Dim idx, anr(1000) As Integer
Public bemerkungs As String

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
 
 Dim paafach, paafach2, paatyp, paaprot, paaprot2 As ADODB.Recordset
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
  paafach2.Open "SELECT * FROM paa_fach WHERE (Name = '" + fachs + "')", Conn
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
End Sub

Sub combo_init()
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
 fpid = 0
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
 
 Dim paafach As ADODB.Recordset
 Set paafach = New ADODB.Recordset
 
 paafach.Open "SELECT * FROM paa_fach", Conn
 
 If Not paafach.EOF Then
 paafach.MoveFirst
 Combo2.Clear
 While Not paafach.EOF
  Combo2.AddItem (paafach!Name)
  paafach.MoveNext
 Wend
 Else
 Combo2.Clear
 Combo2.AddItem ("keine Fächer vorhanden.")
 End If
 paafach.Close
 
 Dim paaprot As ADODB.Recordset
 Set paaprot = New ADODB.Recordset
 
 paaprot.Open "SELECT * FROM paa_protokolle", Conn
 
 If Not paaprot.EOF Then
 paaprot.MoveFirst
 Combo3.Clear
 While Not paaprot.EOF
  Combo3.AddItem (paaprot!prüfer)
  paaprot.MoveNext
 Wend
 Else
 Combo3.Clear
 Combo3.AddItem ("keine Protokollordner vorhanden.")
 End If
 paaprot.Close

End Sub

Private Sub Combo4_Click()
 fpid = Combo4.ItemData(Combo4.ListIndex)
 showausleihe
End Sub

Private Sub Command1_Click(Index As Integer)
 Dim paaaus, paaprot, paafp As ADODB.Recordset
 
 aus_anzeige
 
 If typid = 3 Then
    Set paafp = New ADODB.Recordset
    paafp.Open "SELECT * FROM paa_fp WHERE aktAusleihe=" + Str(ausid), Conn, adOpenKeyset, adLockBatchOptimistic
    If Not paafp.EOF Then
     paafp.MoveFirst
     paafp.Update
     paafp!aktAusleihe = 0
     paafp.UpdateBatch
    Else
     MsgBox ("Es die Ausleihe wurde schon am <Datum> zurückgegeben.")
    End If
    paafp.Close
  End If
  
  Set paaaus = New ADODB.Recordset
  paaaus.Open "SELECT * FROM paa_ausleihe WHERE lfdAusleiheNr = " + Str(ausid), Conn, adOpenKeyset, adLockBatchOptimistic
  If Not paaaus.EOF Then
   paaaus.MoveFirst
   prtid = paaaus!protokollnr
   If paaaus!zurück <> 0 Then
    MsgBox ("Der Protokollordner wurde am " + Str(paaaus!zurück) + " zurückgegeben.")
   Else
    paaaus.Update
    paaaus!zurück = Now
    paaaus.UpdateBatch
   End If
  End If
  paaaus.Close
  
  Set paaprot = New ADODB.Recordset
  paaprot.Open "SELECT * FROM paa_protokolle WHERE lfdProtokollNr = " + Str(prtid), Conn, adOpenKeyset, adLockBatchOptimistic
  If Not paaprot.EOF Then
   paaprot.MoveFirst
   If paaprot!aktAusleihe = ausid Then
    paaprot.Update
    paaprot!aktAusleihe = 0
    paaprot.UpdateBatch
   Else
    MsgBox ("Das Protokoll wurde schon zurückgegeben ...")
   End If
  End If
  paaprot.Close
  
  aus_anzeige
End Sub

Private Sub Command2_Click()
  Dim paaaus, paafp As ADODB.Recordset
 
 aus_anzeige
 
 Set paaaus = New ADODB.Recordset
  paaaus.Open "SELECT * FROM paa_ausleihe WHERE lfdAusleiheNr = " + Str(ausid), Conn, adOpenKeyset, adLockBatchOptimistic
  If Not paaaus.EOF Then
   paaaus.MoveFirst
   fpid = paaaus!fpgruppe
   
    If typid = 3 Then
        Set paafp = New ADODB.Recordset
        paafp.Open "SELECT * FROM paa_fp WHERE lfdFpNr = " + Str(fpid), Conn, adOpenKeyset, adLockBatchOptimistic
        If Not paafp.EOF Then
        paafp.MoveFirst
        paafp.Update
        paafp!eigenes = Now
        paafp.UpdateBatch
        End If
        paafp.Close
    End If
  
  If paaaus!zurück <> 0 Then
   If paaaus!eigenes <> 0 Then
    MsgBox ("Ein neues Protokoll zu diesem Ordner wurde am " + Str(paaaus!zurück) + " abgegeben.")
   Else
    paaaus.Update
    paaaus!eigenes = Now
    paaaus.UpdateBatch
   End If
  Else
   MsgBox ("Dieser Protokollordner wurde noch nicht zurückgegeben. Ein eigenes Protokoll sollte erst anerkannt werden wenn der ausgeliehene Ordner zurückgegeben wurde.")
  End If
  End If
  paaaus.Close
  
  aus_anzeige
End Sub

Private Sub Command4_Click()
 Unload bearbeiten
End Sub

Private Sub Command5_Click()
 Dim paaaus As ADODB.Recordset
 Set paaaus = New ADODB.Recordset
 
 aus_anzeige
 If ausid <> 0 Then
 paaaus.Open "SELECT * FROM paa_ausleihe WHERE lfdAusleiheNr = " + Str(ausid), Conn, adOpenKeyset, adLockBatchOptimistic
 If Not paaaus.EOF Then
  paaaus.MoveFirst
  bemerkungs = paaaus!bemerkung
 End If
 paaaus.Close
 
 bemerkung.Show
 Else
  MsgBox ("Es wurde keine Ausleihe ausgewählt")
 End If
End Sub

Public Sub write_bemerkung()
 Dim paaaus As ADODB.Recordset
 Set paaaus = New ADODB.Recordset
 
 paaaus.Open "SELECT * FROM paa_ausleihe WHERE lfdAusleiheNr = " + Str(ausid), Conn, adOpenKeyset, adLockBatchOptimistic
 If Not paaaus.EOF Then
  paaaus.MoveFirst
  paaaus.Update
  paaaus!bemerkung = Me.bemerkungs
  paaaus.UpdateBatch
 End If
 paaaus.Close
 
 aus_anzeige
End Sub

Private Sub Form_Load()
bemerkungs = ""
List2.BackColor = bearbeiten.BackColor
cnct_conn

Conn.Open

combo_init
End Sub

Private Sub Combo1_Click()
 fachid = 0
 prtid = 0
 ausid = 0
 Combo2.Clear
 Combo3.Clear
 combos
 showausleihe
End Sub


Private Sub Combo2_Click()
 Combo3.Clear
 combos
 showausleihe
End Sub

Private Sub Combo3_Click()
 combos
 showausleihe
End Sub

Sub showausleihe()
Dim schalt_ausleihe As Boolean
 Dim slct_str, aus_str As String
Dim paaprot, paaaus As ADODB.Recordset
Set paaprot = New ADODB.Recordset

schalt_ausleihe = False
idx = 0
slct_str = "SELECT * FROM paa_protokolle WHERE "
aus_str = "SELECT * FROM paa_ausleihe WHERE "

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
  aus_str = "SELECT * FROM paa_ausleihe WHERE "
  Set paaaus = New ADODB.Recordset
  aus_str = aus_str + "(protokollnr = " + Str(paaprot!lfdProtokollNr) + ") "
  If fpid <> 0 And typid = 3 Then aus_str = aus_str + "AND (fpgruppe = " + Str(fpid) + ") "
  paaaus.Open aus_str, Conn
  If Not paaaus.EOF Then
  schalt_ausleihe = True
  paaaus.MoveFirst
  While Not paaaus.EOF
   
   If paaaus!zurück <> 0 Then
        anr(idx) = paaaus!lfdAusleiheNr
        List1.AddItem ("#" + Str(paaaus!lfdAusleiheNr) + " - " + paaprot!prüfer + " - " + paaaus!Name + " - Zurückgegeben")
        List1.ItemData(List1.NewIndex) = paaaus!lfdAusleiheNr
        idx = idx + 1
   ElseIf (CDbl(Now) - CDbl(paaaus!Datum)) > 30 Then
        anr(idx) = paaaus!lfdAusleiheNr
        List1.AddItem ("#" + Str(paaaus!lfdAusleiheNr) + " - " + paaprot!prüfer + " - " + paaaus!Name + " - zu spät")
        List1.ItemData(List1.NewIndex) = paaaus!lfdAusleiheNr
        idx = idx + 1
   Else
        anr(idx) = paaaus!lfdAusleiheNr
        List1.AddItem ("#" + Str(paaaus!lfdAusleiheNr) + " - " + paaprot!prüfer + " - " + paaaus!Name + " - Ausgeliehen")
        List1.ItemData(List1.NewIndex) = paaaus!lfdAusleiheNr
        idx = idx + 1
   End If
   paaaus.MoveNext
  Wend
  'Else
   ' schalt_ausleihe = False
  End If
  paaaus.Close
  paaprot.MoveNext
 Wend
 End If
  
 
 If Not schalt_ausleihe Then
  List1.Clear
  List1.AddItem ("Keine Ausleihen vorhanden.")
 End If
 
End Sub
Private Sub Form_Unload(Cancel As Integer)
 Conn.Close
End Sub

Sub aus_anzeige()
 Dim paaaus, paaprot, paafach, paatyp As ADODB.Recordset
 Set paaaus = New ADODB.Recordset
 
 paaaus.Open "SELECT * FROM paa_ausleihe WHERE lfdAusleiheNr = " + Str(ausid), Conn
 List2.Clear
 
 If Not paaaus.EOF Then
  paaaus.MoveFirst
  
  Set paaprot = New ADODB.Recordset
  paaprot.Open "SELECT * FROM paa_protokolle WHERE lfdProtokollNr = " + Str(paaaus!protokollnr), Conn
  
  List2.AddItem ("Zusammenfassung der Ausleihe")
  List2.AddItem (" ")
  List2.AddItem ("Ausleihenummer:" + Str(paaaus!lfdAusleiheNr))
  List2.AddItem ("FP-Gruppennummer: " + Str(paaaus!fpgruppe))
  List2.AddItem ("Name: " + paaaus!Name)
  List2.AddItem ("eMail-Adresse: " + paaaus!email)
  List2.AddItem ("Ausgeliehen am      : " + Str(paaaus!Datum))

  If paaaus!zurück <> 0 Then
   List2.AddItem ("Zurückgegeben am    : " + Str(paaaus!zurück))
  Else
    List2.AddItem ("Zurückgegeben am    : Noch nicht zurückgegeben!")
  End If
  If paaaus!eigenes <> 0 Then
    List2.AddItem ("Eigenes abgegeben am: " + Str(paaaus!eigenes))
  Else
    List2.AddItem ("Eigenes abgegeben am: Noch kein eigenes Protokoll!")
  End If
  List2.AddItem ("Bemerkungen         : " + paaaus!bemerkung)
  List2.AddItem (" ")
  
  List2.AddItem ("Protokolldaten")
  List2.AddItem (" ")
  If Not paaprot.EOF Then
    Set paafach = New ADODB.Recordset
    paafach.Open "SELECT * FROM paa_fach WHERE lfdFachNr = " + Str(paaprot!fach), Conn
    paafach.MoveFirst
    
    Set paatyp = New ADODB.Recordset
    paatyp.Open "SELECT * FROM paa_typ WHERE lfdTypNr = " + Str(paaprot!typ), Conn
    paatyp.MoveFirst
    
    paaprot.MoveFirst
    List2.AddItem ("Protokollnummer   : " + Str(paaprot!lfdProtokollNr))
    List2.AddItem ("Fach              : " + paafach!Name)
    List2.AddItem ("Protokollordner   : " + paaprot!prüfer)
    List2.AddItem ("Protokolltyp      : " + paatyp!Name)
    List2.AddItem ("akt. Ausleihe Nr. : " + Str(paaprot!aktAusleihe))
    List2.AddItem (" ")
  Else
    List2.AddItem ("Keine Protokolldaten vorhanden.")
  End If
  paafach.Close
  paaprot.Close
 Else
    List2.AddItem ("Keine Ausleihedaten vorhanden")
 End If
 paaaus.Close
 'paaprot.Close
End Sub

Private Sub List1_Click()
 ausid = List1.ItemData(List1.ListIndex)
 'ausid = anr(List1.ListIndex)
 aus_anzeige
End Sub
