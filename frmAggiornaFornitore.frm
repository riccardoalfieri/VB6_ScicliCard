VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmAggiornaFornitore 
   BackColor       =   &H00FFFF00&
   Caption         =   "Ajouter Commandes Fournisseur et Stock"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ajourner Stock"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      Picture         =   "frmAggiornaFornitore.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H0080FF80&
      Caption         =   "Quitter"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10560
      Picture         =   "frmAggiornaFornitore.frx":1CFA
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Supprimer "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8640
      Picture         =   "frmAggiornaFornitore.frx":3A3C
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   6120
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataSource      =   "Adodc4"
      Height          =   375
      Left            =   840
      TabIndex        =   55
      Top             =   7440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdmemo 
      BackColor       =   &H00FF8080&
      Caption         =   "Valider"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      Picture         =   "frmAggiornaFornitore.frx":3B86
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "...."
      Height          =   735
      Left            =   3480
      Picture         =   "frmAggiornaFornitore.frx":5F78
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtcodart 
      Height          =   375
      Left            =   120
      TabIndex        =   47
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtsconto3 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6360
      TabIndex        =   42
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtsconto2 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5400
      TabIndex        =   41
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtsconto1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4440
      TabIndex        =   40
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtprezzolistino 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   7320
      TabIndex        =   39
      Top             =   2760
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmAggiornaFornitore.frx":7C8A
      Height          =   315
      Left            =   5160
      TabIndex        =   38
      Top             =   360
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      Locked          =   -1  'True
      ListField       =   "Pattern"
      Text            =   ""
   End
   Begin VB.PictureBox Picture1 
      Height          =   1530
      Left            =   8760
      Picture         =   "frmAggiornaFornitore.frx":7C9F
      ScaleHeight     =   1470
      ScaleWidth      =   3225
      TabIndex        =   37
      Top             =   0
      Width           =   3285
   End
   Begin VB.TextBox txtdesart 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   120
      TabIndex        =   19
      Top             =   3360
      Width           =   3255
   End
   Begin VB.TextBox txtclicode 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   240
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox txtSupplier 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox txtpezzi 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3720
      TabIndex        =   16
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtconsegna 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   15
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtdtaordine 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtordine 
      Height          =   375
      Left            =   1200
      TabIndex        =   13
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000B&
      Caption         =   "......"
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox txtimportoriga 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "NumeroMassimo"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   1
      EndProperty
      DataSource      =   "Adodc3"
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtaltezza 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   7320
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtbase 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtprezzo 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   4800
      TabIndex        =   7
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txttotiva 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox txttotordine 
      Alignment       =   1  'Right Justify
      DataSource      =   "Adodc4"
      Height          =   375
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   6360
      Width           =   1335
   End
   Begin VB.TextBox txtaliquota 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtquantità 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtimportoiva 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox txttotaleconiva 
      Alignment       =   1  'Right Justify
      DataSource      =   "Adodc4"
      Height          =   375
      Left            =   8160
      TabIndex        =   1
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox txttotimponibile 
      Alignment       =   1  'Right Justify
      DataSource      =   "Adodc4"
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   6360
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   495
      Left            =   6840
      Top             =   5640
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from dettaglioAcquisti WHERE [NUMERO ORDINE]="""""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid5 
      Bindings        =   "frmAggiornaFornitore.frx":17E21
      Height          =   2415
      Left            =   120
      TabIndex        =   20
      Top             =   3720
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   4260
      _Version        =   393216
      AllowUpdate     =   -1  'True
      ColumnHeaders   =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Commandes"
      ColumnCount     =   39
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "id"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Articolo"
         Caption         =   "Articolo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Fornitore"
         Caption         =   "Fornitore"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Numero Ordine"
         Caption         =   "Numero Ordine"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Data Ordine"
         Caption         =   "Data Ordine"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Data Consegna"
         Caption         =   "Data Consegna"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Numero Pezzi"
         Caption         =   "Numero Pezzi"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Prezzo"
         Caption         =   "Prezzo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "PrezzoListino"
         Caption         =   "PrezzoListino"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Zincatura"
         Caption         =   "Zincatura"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "Verniciatura"
         Caption         =   "Verniciatura"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "UM"
         Caption         =   "UM"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "Aliquota"
         Caption         =   "Aliquota"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "Data1"
         Caption         =   "Data1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "Destinazione"
         Caption         =   "Destinazione"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "Riferimento"
         Caption         =   "Riferimento"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "Importo"
         Caption         =   "Importo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column17 
         DataField       =   "iva"
         Caption         =   "iva"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column18 
         DataField       =   "CodiceFornitore"
         Caption         =   "CodiceFornitore"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column19 
         DataField       =   "vuoto"
         Caption         =   "vuoto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column20 
         DataField       =   "Codice prodotto"
         Caption         =   "Codice prodotto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column21 
         DataField       =   "Ubicazione"
         Caption         =   "Ubicazione"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column22 
         DataField       =   "flagstampato"
         Caption         =   "flagstampato"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column23 
         DataField       =   "flagcancellato"
         Caption         =   "flagcancellato"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column24 
         DataField       =   "flagfatturato"
         Caption         =   "flagfatturato"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column25 
         DataField       =   "colore"
         Caption         =   "colore"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column26 
         DataField       =   "base"
         Caption         =   "base"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column27 
         DataField       =   "altezza"
         Caption         =   "altezza"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column28 
         DataField       =   "quantità"
         Caption         =   "quantità"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column29 
         DataField       =   "totaledocumento"
         Caption         =   "totaledocumento"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column30 
         DataField       =   "Numero Fattura"
         Caption         =   "Numero Fattura"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column31 
         DataField       =   "Data Fattura"
         Caption         =   "Data Fattura"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column32 
         DataField       =   "progressivo"
         Caption         =   "progressivo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column33 
         DataField       =   "Corpo"
         Caption         =   "Corpo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column34 
         DataField       =   "Numero DDT"
         Caption         =   "Numero DDT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column35 
         DataField       =   "Data DDT"
         Caption         =   "Data DDT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column36 
         DataField       =   "sconto1"
         Caption         =   "sconto1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column37 
         DataField       =   "sconto2"
         Caption         =   "sconto2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column38 
         DataField       =   "sconto3"
         Caption         =   "sconto3"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   0   'False
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1244,976
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   1184,882
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   1244,976
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   1124,787
         EndProperty
         BeginProperty Column25 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column26 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column27 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column28 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column29 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column30 
            ColumnWidth     =   1409,953
         EndProperty
         BeginProperty Column31 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column32 
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column33 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column34 
            ColumnWidth     =   1184,882
         EndProperty
         BeginProperty Column35 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column36 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column37 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column38 
            ColumnWidth     =   2085,166
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   495
      Left            =   840
      Top             =   1320
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select max(progressivo) as NumeroMassimo FROM dettaglio"
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   375
      Left            =   3600
      Top             =   2760
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Patterns"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "frmAggiornaFornitore.frx":17E36
      Height          =   375
      Left            =   3720
      TabIndex        =   48
      Top             =   7560
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   39
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "id"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Articolo"
         Caption         =   "Articolo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Cliente"
         Caption         =   "Cliente"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Numero Ordine"
         Caption         =   "Numero Ordine"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Data Ordine"
         Caption         =   "Data Ordine"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Data Consegna"
         Caption         =   "Data Consegna"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Numero Pezzi"
         Caption         =   "Numero Pezzi"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Prezzo"
         Caption         =   "Prezzo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Sabbiatura"
         Caption         =   "Sabbiatura"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Zincatura"
         Caption         =   "Zincatura"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "Verniciatura"
         Caption         =   "Verniciatura"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "UM"
         Caption         =   "UM"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "Aliquota"
         Caption         =   "Aliquota"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "Data1"
         Caption         =   "Data1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "Destinazione"
         Caption         =   "Destinazione"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "Riferimento"
         Caption         =   "Riferimento"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "Importo"
         Caption         =   "Importo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column17 
         DataField       =   "iva"
         Caption         =   "iva"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column18 
         DataField       =   "CodiceCliente"
         Caption         =   "CodiceCliente"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column19 
         DataField       =   "vuoto"
         Caption         =   "vuoto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column20 
         DataField       =   "Codice prodotto"
         Caption         =   "Codice prodotto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column21 
         DataField       =   "Ubicazione"
         Caption         =   "Ubicazione"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column22 
         DataField       =   "flagstampato"
         Caption         =   "flagstampato"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column23 
         DataField       =   "flagcancellato"
         Caption         =   "flagcancellato"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column24 
         DataField       =   "flagfatturato"
         Caption         =   "flagfatturato"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column25 
         DataField       =   "colore"
         Caption         =   "colore"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column26 
         DataField       =   "base"
         Caption         =   "base"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column27 
         DataField       =   "altezza"
         Caption         =   "altezza"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column28 
         DataField       =   "quantità"
         Caption         =   "quantità"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column29 
         DataField       =   "totaledocumento"
         Caption         =   "totaledocumento"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column30 
         DataField       =   "Numero Fattura"
         Caption         =   "Numero Fattura"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column31 
         DataField       =   "Data Fattura"
         Caption         =   "Data Fattura"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column32 
         DataField       =   "progressivo"
         Caption         =   "progressivo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column33 
         DataField       =   "Corpo"
         Caption         =   "Corpo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column34 
         DataField       =   "Numero DDT"
         Caption         =   "Numero DDT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column35 
         DataField       =   "Data DDT"
         Caption         =   "Data DDT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column36 
         DataField       =   "sconto1"
         Caption         =   "sconto1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column37 
         DataField       =   "sconto2"
         Caption         =   "sconto2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column38 
         DataField       =   "sconto3"
         Caption         =   "sconto3"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1454,74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   929,764
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   1425,26
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   1544,882
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   1305,071
         EndProperty
         BeginProperty Column25 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column26 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column27 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column28 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column29 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column30 
            ColumnWidth     =   1665,071
         EndProperty
         BeginProperty Column31 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column32 
            ColumnWidth     =   1454,74
         EndProperty
         BeginProperty Column33 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column34 
            ColumnWidth     =   1454,74
         EndProperty
         BeginProperty Column35 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column36 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column37 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column38 
            ColumnWidth     =   2775,118
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   495
      Left            =   5280
      Top             =   7320
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM Dettaglio WHERE [numero ordine]=''"
      Caption         =   "Adodc4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TVA"
      Height          =   255
      Left            =   2760
      TabIndex        =   54
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      Height          =   255
      Left            =   3960
      TabIndex        =   53
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Montant"
      Height          =   255
      Left            =   1320
      TabIndex        =   52
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   840
      TabIndex        =   51
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblsconto3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ecompte 3"
      Height          =   255
      Left            =   6360
      TabIndex        =   46
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblsconto2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ecompte 2"
      Height          =   255
      Left            =   5400
      TabIndex        =   45
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblsconto1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ecompte 1"
      Height          =   255
      Left            =   4440
      TabIndex        =   44
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblprezzolistino 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Prix d'Achat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   43
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Produit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   90
      TabIndex        =   36
      Top             =   3120
      Width           =   3645
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Codice Fornitore"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Fournisseur"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1275
      TabIndex        =   34
      Top             =   720
      Width           =   1140
   End
   Begin VB.Label lblpezzi 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Qté"
      Height          =   255
      Left            =   3720
      TabIndex        =   33
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label lblconsegna 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   8880
      TabIndex        =   32
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lbldtaordine 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lbum 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unit"
      Height          =   255
      Left            =   4800
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblordine 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No"
      Height          =   255
      Left            =   1200
      TabIndex        =   29
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblimporto 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Montant"
      Height          =   255
      Left            =   5760
      TabIndex        =   28
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label lblaltezza 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Altezza"
      Height          =   255
      Left            =   7320
      TabIndex        =   27
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblbase 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Base"
      Height          =   255
      Left            =   6480
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lbliva 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tax"
      Height          =   255
      Left            =   5880
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprezzo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Achat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   24
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblquantità 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Qty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblimportoiva 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TVA"
      Height          =   255
      Left            =   7440
      TabIndex        =   22
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label lbltotaleconiva 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      Height          =   255
      Left            =   8160
      TabIndex        =   21
      Top             =   3120
      Width           =   1455
   End
End
Attribute VB_Name = "frmAggiornaFornitore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdmemo_Click()
If Val(txtpezzi) = 0 Then GoTo dopo


DataGrid5.Columns(1) = txtdesart
DataGrid5.Columns(2) = txtSupplier
DataGrid5.Columns(3) = txtordine
DataGrid5.Columns(4) = txtdtaordine
DataGrid5.Columns(6) = txtpezzi
DataGrid5.Columns(8) = txtprezzolistino
DataGrid5.Columns(7) = txtprezzo
DataGrid5.Columns(11) = DataCombo1
DataGrid5.Columns(5) = txtconsegna
DataGrid5.Columns(14) = txtdestino
DataGrid5.Columns(20) = txtcodart
DataGrid5.Columns(16) = txtimportoriga
DataGrid5.Columns(15) = txtrif
DataGrid5.Columns(18) = txtclicode
'DataGrid5.Columns(9) = Text(1)
'DataGrid5.Columns(10) = Text(0)
DataGrid5.Columns(17) = txtimportoiva
DataGrid5.Columns(25) = txtcolore
DataGrid5.Columns(12) = txtaliquota
DataGrid5.Columns(26) = txtbase
DataGrid5.Columns(27) = txtaltezza
DataGrid5.Columns(28) = txtquantità
DataGrid5.Columns(29) = txttotaleconiva
DataGrid5.Columns(36) = txtsconto1
DataGrid5.Columns(37) = txtsconto2
DataGrid5.Columns(38) = txtsconto3

Adodc5.Recordset.Update

dopo:
  txtdesart = ""
 ' txtordine = ""
' txtdtaordine = ""
 txtpezzi = ""
 txtbase = ""
 txtaltezza = ""
 

 
 
 txtcolore = ""
 txtprezzo = ""
 DataCombo1 = ""
txtimportoiva = ""
txttotaleconiva = ""
 txtcodart = ""
 txtimporto = ""
' txtrif = ""
txttotimponibile = ""
txttotiva = ""
txttotordine = ""


End Sub

Private Sub Command1_Click()
txtdesart = ""
  txtordine = ""
 txtdtaordine = ""
 txtpezzi = ""
 txtbase = ""
 txtaltezza = ""
 
 
 txtprezzo = ""
 DataCombo1 = ""
 txtconsegna = ""
 txtdestino = ""
 txtcodart = ""
 txtimportoriga = ""
 txtrif = ""


Form7.Show vbModal


Dim stringa As String
 stringa = "SELECT * FROM DettaglioAcquisti WHERE codiceFornitore = '" & variabile1 & "' AND flagfatturato=false"

   With Adodc5
    .RecordSource = stringa
    .Refresh
    End With
    
    With DataGrid5
     .ClearFields
     .HoldFields
     .ReBind
     End With
txtclicode = variabile1
txtSupplier = variabile2

 If Adodc5.Recordset.RecordCount > 0 Then
 DataGrid5.Enabled = True
 Else
 DataGrid5.Enabled = False
 End If
 
End Sub

Private Sub Command2_Click()
Form6.Show vbModal
txtcodart = art1
txtdesart = art2
'txtgiorni = art3
txtprezzo = art7
DataCombo1 = art5
End Sub

Private Sub Command3_Click()
response = MsgBox(messaggi(9), vbOKCancel + vbCancel, "Attenzione")
Select Case response
 Case 6
 Adodc5.Recordset.delete
 Case Else
 
 End Select

txtdesart = ""
txtcodart = ""

 txtpezzi = ""
 txtbase = ""
 txtaltezza = ""
  txtprezzo = ""
 DataCombo1 = ""


 txtcodart = ""
 txtimportoriga = ""



End Sub

Private Sub Command4_Click()
Dim stringa1 As String


  stringa1 = "update tires, dettaglioacquisti set giacenza = giacenza + dettaglioacquisti.[numero pezzi],tires.dataultimocarico = dettaglioacquisti.[data ordine],tires.fornitore=dettaglioacquisti.fornitore WHERE tires.codiceinterno=dettaglioacquisti.[codice prodotto] and dettaglioacquisti.[numero ordine]= '" & txtordine & "'"
  
 rs.Open stringa1, cn
  stringa1 = "update dettaglioacquisti set flagfatturato=true WHERE dettaglioacquisti.[codice prodotto] and dettaglioacquisti.[numero ordine]= '" & txtordine & "'"
  
 rs.Open stringa1, cn
 Set f = New frmAggiornaFornitore
f.Show

Unload Me

End Sub

Private Sub Command5_Click()

End Sub

Private Sub DataGrid5_Click()

On Error Resume Next

'Command1.Enabled = False
txtsconto1 = DataGrid5.Columns(36)
txtsconto2 = DataGrid5.Columns(37)
txtsconto3 = DataGrid5.Columns(38)
  
 txtdesart = DataGrid5.Columns(1)
 txtSupplier = DataGrid5.Columns(2)
 txtordine = DataGrid5.Columns(3)
 txtdtaordine = DataGrid5.Columns(4)
 txtconsegna = DataGrid5.Columns(5)
 txtpezzi = DataGrid5.Columns(6)
 txtprezzo = DataGrid5.Columns(7)
txtprezzolistino = Format(DataGrid5.Columns(8), "###,##0.00")

DataCombo1.Text = DataGrid5.Columns(11)
 txtaliquota = DataGrid5.Columns(12)
 txtdestino = DataGrid5.Columns(14)
  txtrif = DataGrid5.Columns(15)
txtimportoriga = Format(DataGrid5.Columns(16), "###,###.00")
  txtimportoiva = Format(DataGrid5.Columns(17), "###,###.00")
    txttotaleconiva = Format(Val(txtimporto) + Val(txtimportoiva), "###,###.00")
  txtclicode = DataGrid5.Columns(18)
     txtcodart = DataGrid5.Columns(20)

txtbase = DataGrid5.Columns(26)
txtaltezza = DataGrid5.Columns(27)
txtquantità = Format(DataGrid5.Columns(28), "###,###.00")
txttotaleconiva = Format(DataGrid5.Columns(29), "###,###.00")


 ' aggiorna Ydummy per indirizzare le stampe
rs.Open "SELECT * FROM Ydummy", cn, adOpenDynamic, adLockOptimistic
 rs![order code] = DataGrid5.Columns(0)
 rs!cliente = DataGrid5.Columns(18)
 rs!campo1 = DataGrid5.Columns(15) 'riferimento
  rs!campo2 = DataGrid5.Columns(14) 'destinazione
   rs![Numero Ordine] = DataGrid5.Columns(3)
    rs!Date = DataGrid5.Columns(4)
  'rs1!numerorighe = Adodc4.Recordset.RecordCount - 1
  rs.Update
   
    rs.Close
    
Command2.Enabled = True

Dim stringa As String
 stringa = "SELECT * FROM Dettaglioacquisti WHERE [numero ordine]='" & txtordine & "'"


   With Adodc4
    .RecordSource = stringa
    .Refresh
    End With
      With DataGrid3
     .ClearFields
     .HoldFields
     .ReBind
     End With
   
a = 0: b = 0: c = 0
Adodc4.Recordset.MoveFirst
 For I = 0 To Adodc4.Recordset.RecordCount - 1
  a = a + DataGrid3.Columns(16)
  b = b + DataGrid3.Columns(17)
    c = c + DataGrid3.Columns(29)
     Adodc4.Recordset.MoveNext
   Next I
    txttotimponibile = Format((a), "###,###.00")
    txttotiva = Format((b), "###,###.00")
     txttotordine = Format((c), "###,###.00")
     
End Sub

Private Sub DataGrid5_DblClick()
Dim stringa As String
 stringa = "SELECT * FROM DettaglioAcquisti WHERE codicefornitore = '" & variabile1 & "' AND [Numero Ordine]='" & txtordine & "'"

   With Adodc5
    .RecordSource = stringa
    .Refresh
    End With
    
    With DataGrid5
     .ClearFields
     .HoldFields
     .ReBind
     End With
     
End Sub

Private Sub Form_Load()
    txtSupplier.Enabled = False
    txtclicode.Enabled = False
    
cmdmemo.Enabled = False
DataGrid5.Enabled = False
lingue

End Sub
Private Sub DataCombo1_Change()

calcola
End Sub

Private Sub datacombo1_Validate(Cancel As Boolean)
calcola
End Sub



Private Sub cmdBack_Click()
'cn.Close

Unload Me
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)



  calcola
End Sub

Private Sub txtimportoiva_Validate(Cancel As Boolean)
calcola
End Sub

Private Sub txtpezzi_Change()
cmdmemo.Enabled = True
End Sub

Private Sub txtpezzi_Validate(Cancel As Boolean)
calcola
End Sub

Private Sub txtprezzo_Validate(Cancel As Boolean)
calcola

End Sub
Public Sub calcola()
Dim sconto1, sconto2, sconto3 As Single



 




On Error Resume Next
txtquantità = ""
varia5 = CDbl(txtprezzo)




    varia1 = CDbl(txtpezzi)
        varia2 = varia1 * varia5
     
         
     
         txtdesart.SetFocus
          
   varia3 = CDbl(txtimportoiva)
         
         
      varia4 = (varia3 * varia2) / 100 + varia2
      
        txtimportoriga = Format(varia2, "###,###,##0.00")
        txttotaleconiva = Format(varia4, "###,###,##0.00")
        txtimportoiva = Format(varia3, "###,###,##0.00")
        
        
       
End Sub
Public Sub lingue()

Select Case lingua
  
 Case Is = "1F"
   Exit Sub
   
   Case Is = "2I"
 frmAggiornaFornitore.Caption = "Aggiornamento Ordini Fornitori"
 
  Label2 = "Fornitore"
lbldtaordine = "Data"
lblordine = "Ordine n."
lbldestino = "Note"
lblrif = "Ref."
Label13 = "Prodotto"
lblimporto = "Totale riga"
Label1 = "Totale Ordine"
lblpezzi = "Qtà"
DataGrid5.Caption = "Dettaglio ordine"
lblprezzo = "Prezzo"
lblimportoiva = "IVA"
lbltotaleconiva = "Totale"
lblsconto1 = "Sconto 1"
lblsconto2 = "Sconto 2"
lblsconto3 = "Sconto 3"
lblprezzolistino = "Prezzo di listino"
Label4 = "Imponibile"
Label7 = "Iva"
Label6 = "Totale "

cmdmemo.Caption = "Salva"
cmdback.Caption = "Esci"
Command3.Caption = "Cancella riga"
Command4.Caption = "Aggiorna Stock"

 Case Is = "3G"
 frmAggiornaFornitore.Caption = "Update Vendors Orders and Stock"
 
  Label2 = "Vendor"
lbldtaordine = "Date"
lblordine = "Order no."
lbldestino = "Note"
lblrif = "Ref."
Label13 = "Product"
lblimporto = "Total line"
Label1 = "Total Order"
lblpezzi = "Qty"
DataGrid5.Caption = "Details"
lblprezzo = "Price"
lblimportoiva = "TAX"
lbltotaleconiva = "Total "
lblsconto1 = "Discount 1"
lblsconto2 = "Discount 2"
lblsconto3 = "Discount 3"
lblprezzolistino = "Original Price"
Label4 = "Subtotal"
Label7 = "Taxes"
Label6 = "Total  "

cmdmemo.Caption = "Save"
cmdback.Caption = "Exit"
Command3.Caption = "Delete row"
Command4.Caption = "Stock Update"

Case Is = "4S"
 frmAggiornaFornitore.Caption = "Actualiza Pedidos a Proveedor y Movimientos de Stock"
 
  Label2 = "Proveedor"
lbldtaordine = "Fecha"
lblordine = "Pedido n."
lbldestino = "Note"
lblrif = "Ref."
Label13 = "Producto"
lblimporto = "Total"
Label1 = "Total Pedido"
lblpezzi = "Cantidad"
DataGrid5.Caption = "Details"
lblprezzo = "Precio"
lblimportoiva = "Iva"
lbltotaleconiva = "Total "
lblsconto1 = "Descuento 1"
lblsconto2 = "Descuento 2"
lblsconto3 = "Descuento 3"
lblprezzolistino = "Precio Original"
Label4 = "Subtotal"
Label7 = "Iva"
Label6 = "Total  "

cmdmemo.Caption = "Guarda"
cmdback.Caption = "Salir"
Command3.Caption = "Borra Linea"
Command4.Caption = "Actualiza Stock"
Case Else
End Select
 
End Sub

