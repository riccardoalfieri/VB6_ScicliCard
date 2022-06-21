VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmcarico 
   Caption         =   "Carico Merci in Magazzino"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      DataField       =   "totalecarico"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   1
      EndProperty
      DataSource      =   "Adodc6"
      Height          =   360
      Left            =   3120
      TabIndex        =   41
      Top             =   6840
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "importocarico"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   1
      EndProperty
      DataSource      =   "Adodc6"
      Height          =   360
      Left            =   3120
      TabIndex        =   40
      Top             =   6480
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      DataField       =   "descrizione"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   2640
      TabIndex        =   35
      Text            =   "Text2"
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6840
      TabIndex        =   1
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtDataDoc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      IMEMode         =   3  'DISABLE
      Left            =   6840
      TabIndex        =   3
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox txtNumDoc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      IMEMode         =   3  'DISABLE
      Left            =   6840
      TabIndex        =   2
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txttotaleconiva 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   1
      EndProperty
      Height          =   360
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txtiva 
      Alignment       =   1  'Right Justify
      DataField       =   "Iva"
      DataSource      =   "Adodc2"
      Height          =   360
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox txttotale 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   1
      EndProperty
      Height          =   360
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   3240
      Width           =   1215
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
      Left            =   240
      TabIndex        =   23
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtcodiceean 
      DataField       =   "codiceEAN"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   7440
      TabIndex        =   22
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Memorizza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10440
      TabIndex        =   9
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox txtpezzi 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Annulla Movimenti"
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
      Left            =   8640
      TabIndex        =   11
      Top             =   6360
      Width           =   1575
   End
   Begin VB.TextBox txtprezzo 
      Alignment       =   1  'Right Justify
      DataField       =   "PrezzoAcquisto"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   0
      EndProperty
      DataSource      =   "Adodc2"
      Height          =   360
      Left            =   5520
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   1530
      Left            =   8400
      Picture         =   "frmcarico.frx":0000
      ScaleHeight     =   1470
      ScaleWidth      =   3225
      TabIndex        =   8
      Top             =   0
      Width           =   3285
   End
   Begin VB.TextBox txtinterno 
      BackColor       =   &H00FFFFFF&
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
      IMEMode         =   3  'DISABLE
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Carica Magazzino"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   10
      Top             =   6360
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   6600
      Top             =   600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Caption         =   "Adodc2"
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmcarico.frx":10182
      Height          =   975
      Left            =   6480
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Archivio Articoli"
      ColumnCount     =   14
      BeginProperty Column00 
         DataField       =   "descrizione"
         Caption         =   "descrizione"
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
         DataField       =   "Varieta"
         Caption         =   "Varieta"
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
         DataField       =   "Quantity"
         Caption         =   "Quantity"
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
         DataField       =   "Min"
         Caption         =   "Min"
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
         DataField       =   "Max"
         Caption         =   "Max"
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
         DataField       =   "Iva"
         Caption         =   "Iva"
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
         DataField       =   "listino1"
         Caption         =   "listino1"
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
         DataField       =   "pallet"
         Caption         =   "pallet"
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
         DataField       =   "misura"
         Caption         =   "misura"
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
         DataField       =   "Price"
         Caption         =   "Price"
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
         DataField       =   "Tire Code"
         Caption         =   "Tire Code"
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
         DataField       =   "Pattern"
         Caption         =   "Pattern"
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
         DataField       =   "Size"
         Caption         =   "Size"
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
         DataField       =   "giorni"
         Caption         =   "giorni"
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
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
         EndProperty
         BeginProperty Column13 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   495
      Left            =   9480
      Top             =   5760
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
      RecordSource    =   "select * from carico WHERE progressivo =0"
      Caption         =   "Adodc5"
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
      Bindings        =   "frmcarico.frx":10197
      Height          =   2775
      Left            =   0
      TabIndex        =   13
      Top             =   3600
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   21
      FormatLocked    =   -1  'True
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
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Carico"
      ColumnCount     =   21
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
         DataField       =   "Numero Progressivo"
         Caption         =   "Numero Progressivo"
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
         DataField       =   "Data"
         Caption         =   "Data"
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
      BeginProperty Column04 
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
      BeginProperty Column05 
         DataField       =   "Prezzo Acquisto"
         Caption         =   "Prezzo Acquisto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column06 
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
      BeginProperty Column07 
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
      BeginProperty Column08 
         DataField       =   "Importo"
         Caption         =   "Importo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column09 
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
      BeginProperty Column10 
         DataField       =   "Codice interno"
         Caption         =   "Codice interno"
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
      BeginProperty Column12 
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
      BeginProperty Column13 
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
      BeginProperty Column14 
         DataField       =   "codice EAN"
         Caption         =   "codice EAN"
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
      BeginProperty Column16 
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
      BeginProperty Column17 
         DataField       =   "Tipodocumento"
         Caption         =   "Tipodocumento"
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
         DataField       =   "Numerodocumento"
         Caption         =   "Numerodocumento"
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
         DataField       =   "Datadocumento"
         Caption         =   "Datadocumento"
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
         DataField       =   "Totaleconiva"
         Caption         =   "Totaleconiva"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   0   'False
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column14 
            Object.Visible         =   0   'False
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column16 
            Object.Visible         =   0   'False
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column20 
            Alignment       =   1
            ColumnWidth     =   2775.118
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmcarico.frx":101AC
      DataField       =   "Fornitore"
      DataSource      =   "datPrimaryRS"
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   1200
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Fornitore"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3360
      Top             =   840
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
      RecordSource    =   "SELECT * FROM Fornitori ORDER BY Fornitore"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   495
      Left            =   840
      Top             =   0
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
      RecordSource    =   "select max(progressivo) as NumeroMassimo FROM carico"
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "frmcarico.frx":101C1
      Height          =   315
      Left            =   1080
      TabIndex        =   5
      Top             =   3240
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "descrizione"
      BoundColumn     =   "CodiceInterno"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   1080
      Top             =   2520
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      RecordSource    =   "select * from tires "
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   495
      Left            =   0
      Top             =   6480
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
      RecordSource    =   "select sum(importo) as importocarico, sum(totaleconiva) as totalecarico from carico where progressivo=9999"
      Caption         =   "Adodc6"
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
   Begin VB.Label lblLabels 
      Caption         =   "Totale Documento"
      Height          =   360
      Index           =   4
      Left            =   1680
      TabIndex        =   43
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      Caption         =   "Totale Imponibile"
      Height          =   360
      Index           =   3
      Left            =   1680
      TabIndex        =   42
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      Caption         =   "Città"
      Height          =   360
      Index           =   2
      Left            =   120
      TabIndex        =   39
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Caption         =   "Indirizzo"
      Height          =   360
      Index           =   1
      Left            =   120
      TabIndex        =   38
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "città"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1200
      TabIndex        =   37
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label lblindirizzo 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "Indirizzo"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1200
      TabIndex        =   36
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Data Doc."
      Height          =   255
      Left            =   5520
      TabIndex        =   34
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Numero Doc."
      Height          =   255
      Left            =   5520
      TabIndex        =   33
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      Caption         =   "Tipo Documento"
      Height          =   480
      Index           =   0
      Left            =   5520
      TabIndex        =   32
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lbltotaleconiva 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Totale con Iva"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   31
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lbliva 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "% Iva"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   29
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lblTotale 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Totale"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   27
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lbldtaordine 
      Caption         =   "Data "
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblordine 
      Caption         =   "Num.Progr."
      Height          =   255
      Left            =   1560
      TabIndex        =   24
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "CARICO DI MAGAZZINO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   21
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fornitore"
      Height          =   480
      Index           =   7
      Left            =   120
      TabIndex        =   20
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descrizione Articolo"
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
      Left            =   1080
      TabIndex        =   19
      Top             =   2880
      Width           =   3315
   End
   Begin VB.Label lblpezzi 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Num.Pezzi"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   18
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lblprezzo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Prezzo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   17
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblinterno 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codice"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   960
   End
   Begin VB.Label lbldtascontrino 
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblscontrino 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "frmcarico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim valore, val1, val2 As Single
Dim f As New frmcarico



Private Sub cmdBack_Click()
Dim response As Integer
Dim stringa As String

response = MsgBox("Vuoi annullare questi movimenti?", vbOKCancel + vbCancel, "Attenzione")

Select Case response
 Case 6
 stringa = "DELETE * FROM carico WHERE carico.progressivo= " & lblscontrino
  rs.Open stringa, cn
 Case Else
 
 End Select



 





Unload Me
End Sub

Private Sub Command1_Click()

 Select Case Val(txtpezzi)
   Case Is = 0
    GoTo dopo
     Case Is > 0
 Adodc5.Recordset.AddNew
  DataGrid5.Columns(16) = lblscontrino
 DataGrid5.Columns(2) = Date
 DataGrid5.Columns(3) = DataCombo2
 DataGrid5.Columns(4) = txtpezzi
 DataGrid5.Columns(5) = txtprezzo
 DataGrid5.Columns(15) = DataCombo1
 DataGrid5.Columns(10) = txtinterno
 DataGrid5.Columns(14) = txtcodiceean
 DataGrid5.Columns(7) = txtiva
  DataGrid5.Columns(8) = txttotale
   DataGrid5.Columns(17) = Combo1.Text
    DataGrid5.Columns(18) = txtNumDoc
     DataGrid5.Columns(19) = txtDataDoc
      DataGrid5.Columns(20) = txttotaleconiva
      
 Adodc5.Recordset.Update

 txtinterno = ""
 txtpezzi = ""
 txtprezzo = ""
  txttotale = ""
   txtiva = ""
    txttotaleconiva = ""
    Text2 = ""
  txtinterno.SetFocus
        Case Else
        End Select
        
dopo:
  
 Dim stringa As String
 stringa = "select sum(importo) as importocarico, sum(totaleconiva) as totalecarico from carico where carico.progressivo = " & lblscontrino


   With Adodc6
    .RecordSource = stringa
    .Refresh
    End With
    
 

 
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
Dim stringa1 As String


   stringa1 = "update tires, carico set giacenza = giacenza + carico.[numero pezzi]," _
   & "tires.prezzoacquisto=carico.[prezzo acquisto],tires.dataultimocarico=carico.data " _
   & "WHERE tires.codiceinterno=carico.[codice interno] and carico.[progressivo]= " & lblscontrino
  

 
   rs.Open stringa1, cn
   
    


Set f = New frmcarico

f.Show

Unload Me
End Sub

Private Sub DataCombo1_Change()
Adodc1.Recordset.MoveFirst
  Adodc1.Recordset.Find "Fornitore = '" & DataCombo1.Text & "'", 0, adSearchForward


End Sub

Private Sub DataCombo1_LostFocus()
Dim stringa As String
 stringa = "SELECT * FROM tires WHERE [fornitore]='" & DataCombo1.Text & "'"

   With Adodc4
    .RecordSource = stringa
    .Refresh
    End With
End Sub

Private Sub DataCombo2_Change()
txtinterno.Text = DataCombo2.BoundText
txtinterno_change
End Sub

Private Sub Form_Load()
On Error Resume Next

Combo1.AddItem "Fattura Acquisto"
Combo1.AddItem "Ddt"
Combo1.AddItem "Documento interno"


lbldtascontrino = Date
lblscontrino = Val(Text1) + 1

Dim stringa As String
 stringa = "SELECT * FROM carico WHERE [progressivo]=" & lblscontrino


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



Private Sub txtDataDoc_lostfocus()
 On Error Resume Next
    txtDataDoc.Text = Format$(CDate(txtDataDoc.Text), "short date")
End Sub

Private Sub txtinterno_change()
 
Adodc2.Recordset.MoveFirst
  Adodc2.Recordset.Find "codiceinterno = '" & txtinterno.Text & "'", 0, adSearchForward
  DataCombo2 = Text2
   
End Sub




Private Sub txtpezzi_lostfocus()

  If Not IsNumeric(txtpezzi.Text) Then
              MsgBox "Introdurre un valore numerico ", vbExclamation
    End If

calcola
     
End Sub

Public Sub calcola()

val1 = Val(txtprezzo.Text)
val2 = Val(txtpezzi)
valore = val1 * val2
txttotale = Format(valore, "####0.00")

If Val(txtiva) > 0 Then
txttotaleconiva = Format(txttotale + ((txttotale * txtiva) / 100), "####0.00")
  Else


   End If
   
End Sub

Private Sub txtprezzo_lostfocus()

If Not IsNumeric(txtprezzo.Text) Then
              MsgBox "Introdurre un valore numerico ", vbExclamation
    End If


calcola
End Sub
