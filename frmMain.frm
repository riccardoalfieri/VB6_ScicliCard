VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "ScicliCard"
   ClientHeight    =   6600
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11280
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   1535
      ButtonWidth     =   1984
      ButtonHeight    =   1429
      AllowCustomize  =   0   'False
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "La Cassa"
            Object.Tag             =   ""
            ImageIndex      =   27
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Impostazioni"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Connettiti"
            Object.Tag             =   ""
            ImageIndex      =   23
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Esci"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6225
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "www.erreasoft.com"
            TextSave        =   "www.erreasoft.com"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1940
            MinWidth        =   1940
            Text            =   "ScicliCard"
            TextSave        =   "ScicliCard"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   12965
            MinWidth        =   5292
            Text            =   "support@erreasoft.com"
            TextSave        =   "support@erreasoft.com"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   1234
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   28
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":59424
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5A076
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5ACC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5B91A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5C56C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5D1BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5DE10
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5EA62
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5F6B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":60306
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":60F58
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":61BAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":627FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":6344E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":640A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":64CF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":65944
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":66596
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":671E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":67E3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":68A8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":696DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":6A330
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":6AF82
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":6BBD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":6C826
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":6D478
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":6E0CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnusocietà 
         Caption         =   "&Impostazioni"
      End
      Begin VB.Menu mnucard 
         Caption         =   "Stampa card"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Esci"
      End
   End
   Begin VB.Menu mnucaisse 
      Caption         =   "&Cassa"
      Begin VB.Menu mnupos 
         Caption         =   "&Transazioni"
      End
      Begin VB.Menu mnuinfo 
         Caption         =   "Info Carta"
      End
   End
   Begin VB.Menu mnuconnect 
      Caption         =   "Credito"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&?"
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&Info"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const EM_UNDO = &HC7
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Sub Command1_Click()

End Sub

Private Sub frmtires_Click()

End Sub

Private Sub frmstampadifferite_Click()
frmstampadiff.Show
End Sub


Private Sub MDIForm_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    LoadNewDoc
    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False;"
   cn.Open
   lingue
 
 
End Sub


Private Sub LoadNewDoc()
    Static lDocumentCount As Long
    Dim frmD As frmDocument
    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmDocument
    frmD.Caption = "Document " & lDocumentCount
     frmD.Picture = LoadPicture("desktop.jpg")

    frmD.Show
    
   

End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    cn.Close
End Sub

Private Sub mnuAggiornaFornitore_Click()
frmAggiornaFornitore.Show
End Sub

Private Sub mnuajoutercontrat_Click()
frmAggiornacontratti.Show
End Sub

Private Sub mnuArchiviArticoli_Click()
FRMARTICO.Show vbModal
End Sub

Private Sub mnuArchivioClienti_Click()
frmCustomer.Show
End Sub

Private Sub mnuArchivioFornitori_Click()
frmFornitori.Show
End Sub

Private Sub mnuArchiviOggetti_Click()
frmoggetti.Show
End Sub

Private Sub mnuArchivioIva_Click()
frmSize.Show
End Sub

Private Sub mnuArchivioMisure_Click()
frmPattern.Show
End Sub



Private Sub mnuArchivioPagamenti_Click()
frmArchivioPagamenti.Show

End Sub


Private Sub mnucard_Click()
frmetich.Show
End Sub

Private Sub mnuconnect_Click()
frmcredit.Show vbModal
End Sub

Private Sub mnuisdn_Click()

End Sub

Private Sub mnuinfo_Click()
frminfo.Show
End Sub

Private Sub mnupos_Click()
frmpos.Show
End Sub



Private Sub mnusocietà_Click()
frmAzienda.Show
End Sub



Private Sub mnuHelpAbout_Click()
    MsgBox "ScicliCard support@erreasoft.com  Ver. " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer


    'Se il progetto non include un file della Guida, visualizza un messaggio per
    'l'utente. È possibile impostare il file della Guida per l'applicazione nella
    'finestra di dialogo Proprietà progetto.
    If Len(App.HelpFile) = 0 Then
        MsgBox "Impossibile visualizzare il Sommario della Guida. Nessun file della Guida associato al progetto.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer


    'Se il progetto non include un file della Guida, visualizza un messaggio per
    'l'utente. È possibile impostare il file della Guida per l'applicazione nella
    'finestra di dialogo Proprietà progetto.
    If Len(App.HelpFile) = 0 Then
        MsgBox "Impossibile visualizzare il Sommario della Guida. Nessun file della Guida associato al progetto.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub



Private Sub mnuViewWebBrowser_Click()
    'Da fare: Aggiunge il codice per 'mnuViewWebBrowser_Click'.
    MsgBox "Aggiunge il codice per 'mnuViewWebBrowser_Click'."
End Sub

Private Sub mnuViewOptions_Click()
    'Da fare: Aggiunge il codice per 'mnuViewOptions_Click'.
    MsgBox "Aggiunge il codice per 'mnuViewOptions_Click'."
End Sub

Private Sub mnuViewRefresh_Click()
    'Da fare: Aggiunge il codice per 'mnuViewRefresh_Click'.
    MsgBox "Aggiunge il codice per 'mnuViewRefresh_Click'."
End Sub

Private Sub mnuvettori_Click()
frmVettori.Show
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    StatusBar1.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    Toolbar1.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuEditPasteSpecial_Click()
    'Da fare: Aggiunge il codice per 'mnuEditPasteSpecial_Click'.
    MsgBox "Aggiunge il codice per 'mnuEditPasteSpecial_Click'."
End Sub

Private Sub mnuEditPaste_Click()
    On Error Resume Next
    ActiveForm.rtfText.SelRTF = Clipboard.GetText

End Sub

Private Sub mnuEditCopy_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtfText.SelRTF

End Sub

Private Sub mnuEditCut_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtfText.SelRTF
    ActiveForm.rtfText.SelText = vbNullString

End Sub

Private Sub mnuEditUndo_Click()
    'Da fare: Aggiunge il codice per 'mnuEditUndo_Click'.
    MsgBox "Aggiunge il codice per 'mnuEditUndo_Click'."
End Sub


Private Sub mnuFileExit_Click()
    'Scarica il form.
    Unload Me

End Sub

Private Sub mnuFileSend_Click()
    'Da fare: Aggiunge il codice per 'mnuFileSend_Click'.
    MsgBox "Aggiunge il codice per 'mnuFileSend_Click'."
End Sub

Private Sub mnuFilePrint_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    

    With dlgCommonDialog
        .DialogTitle = "Stampa"
        .CancelError = True
        .flags = cdlPDReturnDC + cdlPDNoPageNums
        If ActiveForm.rtfText.SelLength = 0 Then
            .flags = .flags + cdlPDAllPages
        Else
            .flags = .flags + cdlPDSelection
        End If
        .ShowPrinter
        If Err <> MSComDlg.cdlCancel Then
            ActiveForm.rtfText.SelPrint .hDC
        End If
    End With

End Sub

Private Sub mnuFilePrintPreview_Click()
    'Da fare: Aggiunge il codice per 'mnuFilePrintPreview_Click'.
    MsgBox "Aggiunge il codice per 'mnuFilePrintPreview_Click'."
End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "Imposta pagina"
        .CancelError = True
        .ShowPrinter
    End With

End Sub

Private Sub mnuFileProperties_Click()
    'Da fare: Aggiunge il codice per 'mnuFileProperties_Click'.
    MsgBox "Aggiunge il codice per 'mnuFileProperties_Click'."
End Sub

Private Sub mnuFileSaveAll_Click()
    'Da fare: Aggiunge il codice per 'mnuFileSaveAll_Click'.
    MsgBox "Aggiunge il codice per 'mnuFileSaveAll_Click'."
End Sub

Private Sub mnuFileSaveAs_Click()
    Dim sFile As String
    

    If ActiveForm Is Nothing Then Exit Sub
    

    With dlgCommonDialog
        .DialogTitle = "Salva con nome"
        .CancelError = False
        'Da fare: impostare i flag e gli attributi del controllo CommonDialog.
        .Filter = "Tutti i file (*.*)|*.*"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    ActiveForm.Caption = sFile
    ActiveForm.rtfText.SaveFile sFile

End Sub

Private Sub mnuFileSave_Click()
    Dim sFile As String
    If Left$(ActiveForm.Caption, 8) = "Document" Then
        With dlgCommonDialog
            .DialogTitle = "Salva"
            .CancelError = False
            'Da fare: impostare i flag e gli attributi del controllo CommonDialog.
            .Filter = "Tutti i file (*.*)|*.*"
            .ShowSave
            If Len(.FileName) = 0 Then
                Exit Sub
            End If
            sFile = .FileName
        End With
        ActiveForm.rtfText.SaveFile sFile
    Else
        sFile = ActiveForm.Caption
        ActiveForm.rtfText.SaveFile sFile
    End If

End Sub

Private Sub mnuFileClose_Click()
    'Da fare: Aggiunge il codice per 'mnuFileClose_Click'.
    MsgBox "Aggiunge il codice per 'mnuFileClose_Click'."
End Sub


Private Sub mnuFileNew_Click()
    LoadNewDoc
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Index
Case 1:
frmpos.Show
Case 2:
frmAzienda.Show
Case 3:
Form1.Show vbModal
Case 4:
Unload Me


End Select
End Sub

Public Sub lingue()

Select Case LINGUA
  
 Case Is = "1F"
 messaggi(7) = " Insérer un numéro "
messaggi(8) = "Attention"
 messaggi(9) = "Effacer la ligne?"
 messaggi(10) = " Impostare la Commport"
 messaggi(11) = "PizzaEnLigne.net - www.pizzaenligne.net support@erreasoft.com Ver. " & App.Major & "." & App.Minor & "." & App.Revision
 messaggi(12) = "Total il est 0"

 
   Exit Sub
   
   Case Is = "2I"
   
Me.Caption = "Ristorante & Pizzeria Online"
Toolbar1.Buttons(1).Caption = "Cassa"
Toolbar1.Buttons(2).Caption = "Prodotti"

Toolbar1.Buttons(3).Caption = "Clienti"
Toolbar1.Buttons(4).Caption = "Gestione Tavoli"
Toolbar1.Buttons(5).Caption = "Preventivi"
Toolbar1.Buttons(6).Caption = "Gestione Comande"
Toolbar1.Buttons(7).Caption = "Fatture "
Toolbar1.Buttons(8).Caption = "Incassi"
Toolbar1.Buttons(9).Caption = "Conto Clienti"
Toolbar1.Buttons(10).Caption = "Esci"

mnuFile.Caption = "File"
mnusocietà.Caption = "Società"
mnuisdn.Caption = "ISDN CAPI"
mnuposset.Caption = "Settaggio Tasti"
mnuticket.Caption = "Settaggio Ticket"
mnuplanset.Caption = "Situazione Tavoli"
mnuFileExit.Caption = "Esci"
mnuStampe.Caption = "Stampe"
mnuStampeOrdini.Caption = "Stampa Preventivi"
mnuStampeFatture.Caption = "Stampa Fatture"
mnuristampe.Caption = "Ristampa Fatture"
'mnustampacontratti.Caption = "Stampa Un DDT"
'mnufatturecontratti.Caption = "Stampa Fatture da DDT"
mnuView.Caption = "Visualizza"
mnuViewToolbar.Caption = "Barra degli Strumenti"
mnuViewStatusBar.Caption = "Barra di stato"
mnuArchivi.Caption = "Archivi"
mnuArchiviArticoli.Caption = "Prodotti"
mnuArchivioClienti.Caption = "Clienti"
mnuArchivioFornitori.Caption = "Fornitori"
'mnuArchivioMisure.Caption = "Unità Misura"
mnuArchivioIva.Caption = "Iva"
mnuArchivioreparti.Caption = "Categorie"
mnuvettori.Caption = "Vettori"
mnuArchivioPagamenti.Caption = "Condizioni di Pagamento"
mnuvendite.Caption = "Vendite"
mnuvendite1.Caption = "Stampa Ordini"
mnuOrdiniNuovi.Caption = "Inserimento Ordine"
mnuOrdiniModifica.Caption = "Revisione Ordine"
'mnucodebarre.Caption = "Vendita con Codici a Barre"
mnuschede.Caption = "Schede Contabili"
mnuSchedeclienti.Caption = "Schede Clienti"
mnuSchedeFornitori.Caption = "Schede Fornitori"
mnuschedescadenziario.Caption = "Scadenzario Fornitori"

mnuAcquisti.Caption = "Acquisti"
mnuOrdineFornitore.Caption = "Ordine a Fornitore"
mnuAggiornaFornitore.Caption = "Aggorna Ordine Fornitore"
mnuFattureFornitori.Caption = "Inserimento Fatture Fornitori"
mnumagazzino.Caption = "Magazzino"
mnumagazzinoinventario.Caption = "Inventario"
mnumagazzinosottoscorta.Caption = "Prodotti in riordino"
mnumagazzinofornitore.Caption = "Prodotti per fornitore"
mnucommandesf.Caption = "Elenco Ordini a Fornitore"
mnucaisse.Caption = "La Cassa"
mnupos.Caption = "La Cassa"
mnutojour.Caption = "Incassi di Oggi"
mnuencaissement.Caption = "Incassi per Periodo"
mnucassa.Caption = "Modifica Movimenti di Cassa"
mnuHelp.Caption = "?"
mnuHelpAbout.Caption = "RistoPizza Facile"


messaggi(7) = "Inserire un numero [0 - 10000]"
messaggi(8) = "Attenzione"
messaggi(9) = "Cancellare la riga?"
messaggi(10) = " Impostare la Commport"
messaggi(11) = "Ristorante & Pizzeria Online PizzaEnLigne.net - www.pizzaenligne.net support@erreasoft.com Ver. " & App.Major & "." & App.Minor & "." & App.Revision
 messaggi(12) = "Il totale è 0"




 Case Is = "3G"
   
   
Me.Caption = "Ristorante Tattile Facile"
Toolbar1.Buttons(1).Caption = "Cassa"
Toolbar1.Buttons(2).Caption = "Prodotti"

Toolbar1.Buttons(3).Caption = "Clienti"
Toolbar1.Buttons(4).Caption = "Gestione Tavoli"
Toolbar1.Buttons(5).Caption = "Inventario"
Toolbar1.Buttons(6).Caption = "Gestione Comande"
Toolbar1.Buttons(7).Caption = "Fatture "
Toolbar1.Buttons(8).Caption = "Incassi"
Toolbar1.Buttons(9).Caption = "Conto Clienti"
Toolbar1.Buttons(10).Caption = "Esci"

mnusocietà.Caption = "Società"
mnuisdn.Caption = "ISDN CAPI"
mnuposset.Caption = "Settaggio Tasti"
mnuticket.Caption = "Settaggio Ticket"
mnuplanset.Caption = "Settaggio Tavoli"
mnuFileExit.Caption = "Esci"
mnuStampe.Caption = "Stampe"
mnuStampeOrdini.Caption = "Stampa Preventivi"
mnuStampeFatture.Caption = "Stampa Fatture"
mnuristampe.Caption = "Ristampa Fatture"
'mnustampacontratti.Caption = "Stampa Un DDT"
'mnufatturecontratti.Caption = "Stampa Fatture da DDT"
mnuView.Caption = "Visualizza"
mnuViewToolbar.Caption = "Barra degli Strumenti"
mnuViewStatusBar.Caption = "Barra di stato"
mnuArchivi.Caption = "Archivi"
mnuArchiviArticoli.Caption = "Prodotti"
mnuArchivioClienti.Caption = "Clienti"
mnuArchivioFornitori.Caption = "Fornitori"
'mnuArchivioMisure.Caption = "Unità Misura"
mnuArchivioIva.Caption = "Iva"
mnuArchivioreparti.Caption = "Categorie"
mnuvettori.Caption = "Vettori"
mnuArchivioPagamenti.Caption = "Condizioni di Pagamento"
mnuvendite.Caption = "Vendite"
mnuvendite1.Caption = "Stampa Ordini"
mnuOrdiniNuovi.Caption = "Inserimento Ordine"
mnuOrdiniModifica.Caption = "Revisione Ordine"
'mnucodebarre.Caption = "Vendita con Codici a Barre"
mnuschede.Caption = "Schede Contabili"
mnuSchedeclienti.Caption = "Schede Clienti"
mnuSchedeFornitori.Caption = "Schede Fornitori"
mnuschedescadenziario.Caption = "Scadenzario Fornitori"

mnuAcquisti.Caption = "Acquisti"
mnuOrdineFornitore.Caption = "Ordine a Fornitore"
mnuAggiornaFornitore.Caption = "Aggorna Ordine Fornitore"
mnuFattureFornitori.Caption = "Inserimento Fatture Fornitori"
mnumagazzino.Caption = "Magazzino"
mnumagazzinoinventario.Caption = "Inventario"
mnumagazzinosottoscorta.Caption = "Prodotti in riordino"
mnumagazzinofornitore.Caption = "Prodotti per fornitore"
mnucommandesf.Caption = "Elenco Ordini a Fornitore"
mnucaisse.Caption = "La Cassa"
mnupos.Caption = "La Cassa"
mnutojour.Caption = "Incassi di Oggi"
mnuencaissement.Caption = "Incassi per Periodo"
mnucassa.Caption = "Modifica Movimenti di Cassa"
mnuHelp.Caption = "?"
mnuHelpAbout.Caption = "Ristorante Tattile"

messaggi(7) = "Insert a number [0 - 10000]"
messaggi(8) = "Warning"
messaggi(9) = "Delete the line?"
messaggi(10) = " Impostare la Commport"
messaggi(11) = "Easy Salon - www.alfierisoftware.net support@alfierisoftware.net Ver. " & App.Major & "." & App.Minor & "." & App.Revision
messaggi(12) = "Total it is 0"
Case Is = "4S"
   
   
Me.Caption = "Ristorante Tattile Facile"
Toolbar1.Buttons(1).Caption = "Cassa"
Toolbar1.Buttons(2).Caption = "Prodotti"

Toolbar1.Buttons(3).Caption = "Clienti"
Toolbar1.Buttons(4).Caption = "Gestione Tavoli"
Toolbar1.Buttons(5).Caption = "Stampa Ddt"
Toolbar1.Buttons(6).Caption = "Gestione Comande"
Toolbar1.Buttons(7).Caption = "Fatture "
Toolbar1.Buttons(8).Caption = "Incassi"
Toolbar1.Buttons(9).Caption = "Conto Clienti"
Toolbar1.Buttons(10).Caption = "Esci"

mnuFile.Caption = "File"
mnusocietà.Caption = "Società"
mnuisdn.Caption = "ISDN CAPI"
mnuposset.Caption = "Settaggio Tasti"
mnuticket.Caption = "Settaggio Ticket"
mnuplanset.Caption = "Settaggio Tavoli"
mnuFileExit.Caption = "Esci"
mnuStampe.Caption = "Stampe"
mnuStampeOrdini.Caption = "Stampa Preventivi"
mnuStampeFatture.Caption = "Stampa Fatture"
mnuristampe.Caption = "Ristampa Fatture"
'mnustampacontratti.Caption = "Stampa Un DDT"
'mnufatturecontratti.Caption = "Stampa Fatture da DDT"
mnuView.Caption = "Visualizza"
mnuViewToolbar.Caption = "Barra degli Strumenti"
mnuViewStatusBar.Caption = "Barra di stato"
mnuArchivi.Caption = "Archivi"
mnuArchiviArticoli.Caption = "Prodotti"
mnuArchivioClienti.Caption = "Clienti"
mnuArchivioFornitori.Caption = "Fornitori"
'mnuArchivioMisure.Caption = "Unità Misura"
mnuArchivioIva.Caption = "Iva"
mnuArchivioreparti.Caption = "Categorie"
mnuvettori.Caption = "Vettori"
mnuArchivioPagamenti.Caption = "Condizioni di Pagamento"
mnuvendite.Caption = "Vendite"
mnuvendite1.Caption = "Stampa Ordini"
mnuOrdiniNuovi.Caption = "Inserimento Ordine"
mnuOrdiniModifica.Caption = "Revisione Ordine"
'mnucodebarre.Caption = "Vendita con Codici a Barre"
mnuschede.Caption = "Schede Contabili"
mnuSchedeclienti.Caption = "Schede Clienti"
mnuSchedeFornitori.Caption = "Schede Fornitori"
mnuschedescadenziario.Caption = "Scadenzario Fornitori"

mnuAcquisti.Caption = "Acquisti"
mnuOrdineFornitore.Caption = "Ordine a Fornitore"
mnuAggiornaFornitore.Caption = "Aggorna Ordine Fornitore"
mnuFattureFornitori.Caption = "Inserimento Fatture Fornitori"
mnumagazzino.Caption = "Magazzino"
mnumagazzinoinventario.Caption = "Inventario"
mnumagazzinosottoscorta.Caption = "Prodotti in riordino"
mnumagazzinofornitore.Caption = "Prodotti per fornitore"
mnucommandesf.Caption = "Elenco Ordini a Fornitore"
mnucaisse.Caption = "La Cassa"
mnupos.Caption = "La Cassa"
mnutojour.Caption = "Incassi di Oggi"
mnuencaissement.Caption = "Incassi per Periodo"
mnucassa.Caption = "Modifica Movimenti di Cassa"
mnuHelp.Caption = "?"
mnuHelpAbout.Caption = "Ristorante Tattile"


messaggi(7) = "Insertar un número [0 - 10000]"
messaggi(8) = "Atención"
messaggi(9) = "¿borrar la línea?"
messaggi(10) = " Impostare la Commport"
messaggi(12) = "Total es 0"
messaggi(11) = "Salón Fácil - www.alfierisoftware.net support@alfierisoftware.net Ver. " & App.Major & "." & App.Minor & "." & App.Revision
  Case Else
  End Select
  


End Sub

