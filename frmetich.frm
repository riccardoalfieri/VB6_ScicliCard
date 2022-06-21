VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A13954C6-163B-4D45-990C-1414B867D177}#182.0#0"; "MCPrnEt.ocx"
Begin VB.Form frmetich 
   BackColor       =   &H8000000D&
   Caption         =   "Stampa Card"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12075
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   12075
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Articoli"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   17
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
      Begin VB.CheckBox Check2 
         Caption         =   "Usati"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Nuovi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1455
      End
   End
   Begin MCPrnEt.MCEtichette et 
      Left            =   3360
      Top             =   6720
      _ExtentX        =   873
      _ExtentY        =   873
      SpazioOrizzontale=   1
      SpazioVerticale =   1
      AltezzaEtichetta=   30
      LarghezzaEtichetta=   65
   End
   Begin VB.CommandButton stampa 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Stampa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      Picture         =   "frmetich.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
      Caption         =   "Esci"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6240
      Picture         =   "frmetich.frx":1D12
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton anteprima 
      BackColor       =   &H00FFFF80&
      Caption         =   "Anteprima"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      Picture         =   "frmetich.frx":3A54
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton cmddel 
      BackColor       =   &H000080FF&
      Caption         =   "Cancella Elenco"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   4920
      Picture         =   "frmetich.frx":431E
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdmemo 
      BackColor       =   &H00FF8080&
      Caption         =   "Salva"
      Height          =   735
      Left            =   5160
      Picture         =   "frmetich.frx":4468
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "...."
      Height          =   735
      Left            =   4320
      Picture         =   "frmetich.frx":685A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmetich.frx":856C
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1085
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "altezzafoglio"
         Caption         =   "Altezza Pagina"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "larghezzafoglio"
         Caption         =   "Larghezza Pagina"
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
         DataField       =   "altezzaetichetta"
         Caption         =   "Altezza Etich."
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
         DataField       =   "larghezzaetichetta"
         Caption         =   "Larghezza Etich."
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
         DataField       =   "marginealto"
         Caption         =   "Margine Sup."
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
         DataField       =   "marginebasso"
         Caption         =   "Margine Sub."
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
         DataField       =   "marginedestro"
         Caption         =   "Margine Dx"
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
         DataField       =   "marginesinistro"
         Caption         =   "Margine Sx"
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
         DataField       =   "spazioorizzontale"
         Caption         =   "Distanza Orizz."
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
         DataField       =   "spazioverticale"
         Caption         =   "Distanza Vert."
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
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1124,787
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1184,882
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1365,165
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1110,047
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1275,024
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1124,787
         EndProperty
      EndProperty
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
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox txtcodart 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   -120
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   225
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   495
      Index           =   0
      Left            =   7200
      Top             =   4440
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
      RecordSource    =   "select * from etichette"
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
      Bindings        =   "frmetich.frx":8581
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   4260
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Caption         =   "Elenco"
      ColumnCount     =   14
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
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
         DataField       =   "codicearticolo"
         Caption         =   "codicearticolo"
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
         DataField       =   "articolo"
         Caption         =   "articolo"
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
         DataField       =   "numeropezzi"
         Caption         =   "numeropezzi"
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
         DataField       =   "reparto"
         Caption         =   "reparto"
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
         DataField       =   "funzione"
         Caption         =   "funzione"
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
         DataField       =   "taglia"
         Caption         =   "taglia"
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
      BeginProperty Column08 
         DataField       =   "codicefornitore"
         Caption         =   "codicefornitore"
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
         DataField       =   "fornitore"
         Caption         =   "fornitore"
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
         DataField       =   "campolibero"
         Caption         =   "campolibero"
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
         DataField       =   "prezzo"
         Caption         =   "Numero Card"
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
         DataField       =   "anno"
         Caption         =   "anno"
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
         DataField       =   "costo"
         Caption         =   "costo"
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
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
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
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
            ColumnWidth     =   2085,166
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   5040
      Top             =   360
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
      RecordSource    =   "select * from impostazioni"
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   495
      Index           =   2
      Left            =   9480
      Top             =   720
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
      RecordSource    =   "select * from tires"
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   495
      Index           =   3
      Left            =   9480
      Top             =   1200
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
      RecordSource    =   "select * from carico"
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "Stampa Card"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   10
      Top             =   720
      Width           =   8055
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descrizione"
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
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label lblpezzi 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "N.Posiz.Vuote"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   5640
      Width           =   1095
   End
End
Attribute VB_Name = "frmetich"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Anteprima_Click()
et.anteprima = True
  
  StampaEtichette
End Sub

Private Sub Check1_Click()
Select Case Check1.Value
    Case 1
     Check2.Value = 0
     Case Else
      Check2.Value = 1
     End Select
     
  
    
End Sub

Private Sub Check2_Click()
 Select Case Check2.Value
    Case 1
     Check1.Value = 0
     Case Else
      Check1.Value = 1
    End Select
End Sub

Private Sub cmddel_Click(Index As Integer)
Dim stringa As String
stringa = "delete * from etichette"
rs.Open stringa, cn

 stringa = "select * from etichette"


   With Adodc5(0)
    .RecordSource = stringa
    .Refresh
    End With
    
    With DataGrid5
     .ClearFields
     .HoldFields
     .ReBind
     End With
     
     Unload Me

End Sub

Private Sub cmdmemo_Click()
Adodc5(0).Recordset.AddNew
DataGrid5.Columns(1) = art1 'codice
DataGrid5.Columns(2) = art2 ' articolo
DataGrid5.Columns(5) = art12 ' cod.ean
'DataGrid5.Columns(4) = art9 ' reparto
'DataGrid5.Columns(6) = txt(1)
DataGrid5.Columns(9) = art11 'fornitore
'DataGrid5.Columns(10) = txt(3)
'DataGrid5.Columns(13) = art8 'costo
DataGrid5.Columns(11) = Text2 'prezzo
'DataGrid5.Columns(12) = art11 'costo





Adodc5(0).Recordset.Update

End Sub

Private Sub Command2_Click()
Dim stringa As String
variabile3 = 1 ' prezzo 1
If Check1.Value = 1 Then
Form6.Show vbModal
Else
variabile3 = 1 ' prezzo iva inclusa
Form8.Show vbModal
End If

On Error Resume Next

Text1 = art1
txtdesart = art2
'art3 = flag distinta
Text2 = Format(art4, "###,###,###.00")


txtpezzi.SetFocus
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Dim stringa As String
 stringa = "delete * from etichette"
 rs.Open stringa, cn
 stringa = " insert into etichette select articolo,[codice interno] as codicearticolo,taglia,colore,[numero pezzi] as numeropezzi from carico where carico.numerodocumento='" & DataCombo2 & "'"
 rs.Open stringa, cn
 stringa = "update etichette,tires set etichette.reparto=tires.reparto, etichette.funzione=tires.funzione,etichette.fornitore=tires.fornitore,etichette.prezzo=tires.listino1,etichette.anno=tires.per2,etichette.costo=tires.pallet WHERE etichette.codicearticolo=tires.[codiceinterno]"
 rs.Open stringa, cn
 stringa = " update etichette,tabellacolori set etichette.campolibero=tabellacolori.colore where etichette.colore=tabellacolori.[codice colore]"
 rs.Open stringa, cn
 
  With DataGrid5
     .ClearFields
     .HoldFields
     .ReBind
     End With
 
End Sub

Private Sub DataCombo1_Change()
txt(3) = DataCombo1.BoundText
End Sub

Private Sub Form_Load()
  pageId = 0
End Sub

Private Sub stampa_Click()
  ' esegui stampa sulla stampante dei default
  et.anteprima = False
  et.PrnDev = Printer
  
  StampaEtichette
End Sub

Private Sub StampaEtichette()
On Error Resume Next


et.AltezzaEtichetta = DataGrid1.Columns(2)
et.AltezzaFoglio = DataGrid1.Columns(0)
et.LarghezzaEtichetta = DataGrid1.Columns(3)
et.LarghezzaFoglio = DataGrid1.Columns(1)
et.MargineAlto = DataGrid1.Columns(4)
et.MargineBasso = DataGrid1.Columns(5)
et.MargineDx = DataGrid1.Columns(6)
et.MargineSx = DataGrid1.Columns(7)
et.SpazioOrizzontale = DataGrid1.Columns(8)
et.SpazioVerticale = DataGrid1.Columns(9)
'et.NEtichetteOrizzontali = DataGrid1.Columns(10)
'et.NEtichetteVerticali = DataGrid1.Columns(11)

  Dim dev As Object
  
  Set dev = et.PrnDev
  
  For e = 1 To Val(txt(0))
   et.BordoEt = False
   et.NuovaEtichetta
  
  Next e

  
Adodc5(0).Recordset.MoveFirst
  
  Dim i As Integer
  i = o
  Do
  
  i = i + 1
  et.BordoEt = True
    et.NuovaEtichetta
   et.PrnDev.FontBold = True
   et.PrnDev.FontSize = 14
   
      ' et.PrintText mcCentra, 10, "Etichetta n." & Format(i)
       et.PrintText mcCentra, 1, "Scicli      Card"
       et.PrnDev.FontBold = False
       et.PrnDev.FontSize = 5
       
      et.PrintBarCode mcCentra, 1, DataGrid5.Columns(5), 3, 10
    ' et.PrintText mcCentra, 20, DataGrid5.Columns(1) + " " + DataGrid5.Columns(5) + " " + DataGrid5.Columns(9) ' CODICE ARTICOLO E FORNITORE
     et.PrnDev.FontBold = True
   et.PrnDev.FontSize = 12
   
       et.PrintText mcCentra, 11, DataGrid5.Columns(11)
   



    
    
       
       
   Adodc5(0).Recordset.MoveNext
  Loop Until Adodc5(0).Recordset.EOF
  
  et.FineStampa
End Sub

