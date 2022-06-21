VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmAggiorna 
   BackColor       =   &H80000003&
   Caption         =   "Aggiornamento Ordini"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Elimina Riga"
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
      Left            =   7200
      Picture         =   "frmAggiorna.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton cmdback 
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
      Left            =   9120
      Picture         =   "frmAggiorna.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   7440
      Width           =   1335
   End
   Begin VB.TextBox txtimportoriga 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   0
      TabIndex        =   49
      Top             =   6600
      Width           =   1575
   End
   Begin VB.TextBox txttotiva 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6840
      TabIndex        =   48
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox txttotordine 
      Alignment       =   1  'Right Justify
      DataSource      =   "Adodc4"
      Height          =   375
      Left            =   8040
      TabIndex        =   47
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox txtimportoiva 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1560
      TabIndex        =   46
      Top             =   6600
      Width           =   1455
   End
   Begin VB.TextBox txttotaleconiva 
      Alignment       =   1  'Right Justify
      DataSource      =   "Adodc4"
      Height          =   375
      Left            =   3000
      TabIndex        =   45
      Top             =   6600
      Width           =   1455
   End
   Begin VB.TextBox txttotimponibile 
      Alignment       =   1  'Right Justify
      DataSource      =   "Adodc4"
      Height          =   375
      Left            =   5400
      TabIndex        =   44
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton cmdmemo 
      BackColor       =   &H00FF8080&
      Caption         =   "Memorizza"
      Height          =   735
      Left            =   9840
      Picture         =   "frmAggiorna.frx":1E8C
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "...."
      Height          =   735
      Left            =   3600
      Picture         =   "frmAggiorna.frx":427E
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   720
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   4920
      TabIndex        =   32
      Top             =   960
      Width           =   255
   End
   Begin MSDataGridLib.DataGrid DataGrid4 
      Bindings        =   "frmAggiorna.frx":5F90
      Height          =   855
      Left            =   600
      TabIndex        =   31
      Top             =   7680
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Giacenza"
         Caption         =   "GIACENZA"
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
         DataField       =   ""
         Caption         =   ""
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
      EndProperty
   End
   Begin VB.TextBox txtcodart 
      Height          =   285
      Left            =   0
      TabIndex        =   30
      Top             =   3360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtoldval 
      Height          =   285
      Left            =   960
      TabIndex        =   29
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmAggiorna.frx":5FA5
      Height          =   315
      Left            =   3720
      TabIndex        =   27
      Top             =   3360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Pattern"
      Text            =   ""
   End
   Begin VB.PictureBox Picture1 
      Height          =   1530
      Left            =   8640
      Picture         =   "frmAggiorna.frx":5FBA
      ScaleHeight     =   1470
      ScaleWidth      =   3225
      TabIndex        =   26
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
      Left            =   0
      TabIndex        =   15
      Top             =   3360
      Width           =   3375
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
      TabIndex        =   14
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
      TabIndex        =   13
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox txtpezzi 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4560
      TabIndex        =   12
      Top             =   3360
      Width           =   975
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
      TabIndex        =   11
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
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
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtordine 
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000B&
      Caption         =   "......"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox txtdestino 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   2760
      Width           =   4095
   End
   Begin VB.TextBox txtrif 
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   2760
      Width           =   4575
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
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtaltezza 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtbase 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtprezzo 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   8400
      TabIndex        =   3
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtaliquota 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   9360
      TabIndex        =   2
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtquantità 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3360
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   8160
      Top             =   1680
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
      RecordSource    =   "SELECT * FROM Customers ORDER BY Cliente"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   7800
      Top             =   1320
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
      RecordSource    =   "select * from tires ORDER BY Descrizione"
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmAggiorna.frx":1613C
      Height          =   1935
      Left            =   8400
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3413
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
      RecordSource    =   "select * from dettaglio WHERE [NUMERO ORDINE]=codicecliente"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAggiorna.frx":16151
      Height          =   1935
      Left            =   7920
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3413
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483634
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
      Caption         =   "Anagrafica Clienti"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "Customer Code"
         Caption         =   "Customer Code"
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
      BeginProperty Column02 
         DataField       =   "Indirizzo"
         Caption         =   "Indirizzo"
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
         DataField       =   "NumeroTelefonico"
         Caption         =   "NumeroTelefonico"
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
         DataField       =   "città"
         Caption         =   "città"
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
         DataField       =   "PartitaIva"
         Caption         =   "PartitaIva"
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
         DataField       =   "CodiceFiscale"
         Caption         =   "CodiceFiscale"
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
         DataField       =   "listino"
         Caption         =   "listino"
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
      Left            =   3840
      Top             =   2640
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
      Bindings        =   "frmAggiorna.frx":16166
      Height          =   375
      Left            =   2520
      TabIndex        =   28
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
      Left            =   4080
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
   Begin MSDataGridLib.DataGrid DataGrid5 
      Bindings        =   "frmAggiorna.frx":1617B
      Height          =   2415
      Left            =   0
      TabIndex        =   34
      Top             =   3720
      Width           =   13335
      _ExtentX        =   23521
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
      Caption         =   "Commande"
      ColumnCount     =   33
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
         Caption         =   "Produit"
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
         Caption         =   "Numèro Devis"
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
         Caption         =   "Qtè"
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
         Caption         =   "Prix"
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
         Caption         =   "TVA %"
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
         Caption         =   "Montant"
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
         Caption         =   "TVA"
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
         Caption         =   "Qtè"
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
         Caption         =   "Montant"
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
         Caption         =   "Numèro"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column14 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column15 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column16 
         EndProperty
         BeginProperty Column17 
         EndProperty
         BeginProperty Column18 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column19 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column20 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column21 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column22 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column23 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column24 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column25 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column26 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column27 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column28 
         EndProperty
         BeginProperty Column29 
         EndProperty
         BeginProperty Column30 
         EndProperty
         BeginProperty Column31 
         EndProperty
         BeginProperty Column32 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblimporto 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Imponibile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   55
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label lbltotiva 
      Alignment       =   2  'Center
      Caption         =   "Totale Iva"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   54
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label lbltotordine 
      Alignment       =   2  'Center
      Caption         =   "Totale Generale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   53
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label lblimportoiva 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Iva"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   52
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label lbltotaleconiva 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Totale Riga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   51
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Tot.Imponibile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   50
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Totale"
      Height          =   255
      Left            =   7440
      TabIndex        =   42
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblpezzi 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Qtà"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   41
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descrizione"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   40
      Top             =   3120
      Width           =   1320
   End
   Begin VB.Label lbldestino 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Destinazione"
      Height          =   255
      Left            =   0
      TabIndex        =   39
      Top             =   2520
      Width           =   4095
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "N."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   37
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   36
      Top             =   720
      Width           =   3480
   End
   Begin VB.Label Label3 
      Caption         =   "Cambia Cliente"
      Height          =   255
      Left            =   5160
      TabIndex        =   33
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Codice Cliente"
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
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Label lblconsegna 
      Caption         =   "Data Consegna"
      Height          =   255
      Left            =   8880
      TabIndex        =   24
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lbum 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UM"
      Height          =   255
      Left            =   3720
      TabIndex        =   23
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lblrif 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Riferimenti"
      Height          =   255
      Left            =   4080
      TabIndex        =   22
      Top             =   2520
      Width           =   4335
   End
   Begin VB.Label lblaltezza 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6480
      TabIndex        =   21
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblbase 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5520
      TabIndex        =   20
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lbliva 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "IVA"
      Height          =   255
      Left            =   9360
      TabIndex        =   19
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label lblprezzo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Prezzo"
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
      Left            =   8400
      TabIndex        =   18
      Top             =   3120
      Width           =   975
   End
End
Attribute VB_Name = "frmAggiorna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdmemo_Click()
If Val(txtpezzi) = 0 Then GoTo dopo
                If Check1.Value > 0 Then
                
  response = MsgBox("Vuoi cambiare il cliente di questo ordine?", vbOKCancel + vbCancel, "Attenzione")
     Select Case response
         Case 6
           
     stringa1 = "update dettaglio set cliente = '" & txtSupplier & "' ,codicecliente='" & txtclicode & "' WHERE dettaglio.[numero ordine] = '" & txtordine & "'"
       rs.Open stringa1, cn
       Check1.Value = 0
        GoTo dopo
        
           Case Else
             GoTo dopo
              End Select

                End If


DataGrid5.Columns(1) = txtdesart
DataGrid5.Columns(2) = txtSupplier
DataGrid5.Columns(3) = txtordine
DataGrid5.Columns(4) = txtdtaordine
DataGrid5.Columns(6) = txtpezzi

DataGrid5.Columns(7) = txtprezzo
DataGrid5.Columns(11) = DataCombo1
DataGrid5.Columns(5) = txtconsegna
DataGrid5.Columns(14) = txtdestino
DataGrid5.Columns(20) = txtcodart
DataGrid5.Columns(16) = txtimportoriga
DataGrid5.Columns(15) = txtrif
DataGrid5.Columns(18) = txtclicode

DataGrid5.Columns(17) = txtimportoiva
DataGrid5.Columns(25) = txtcolore
DataGrid5.Columns(12) = txtaliquota
DataGrid5.Columns(26) = txtbase
DataGrid5.Columns(27) = txtaltezza
DataGrid5.Columns(28) = txtquantità
DataGrid5.Columns(29) = txttotaleconiva

Adodc5.Recordset.Update

'aggiorna giacenza
 DataGrid4.Columns(0) = DataGrid4.Columns(0) + Val(txtoldval) - Val(txtpezzi)
Adodc2.Recordset.Update

dopo:
  txtdesart = ""
 txtpezzi = ""
 txtbase = ""
 txtaltezza = ""
 
 txtcolore = ""
 txtprezzo = ""
 DataCombo1 = ""
 txtcodart = ""
 txtimporto = ""

txttotimponibile = ""
txttotiva = ""
txttotordine = ""


End Sub

Private Sub Command1_Click()



Form4.Show vbModal

If Check1.Value = 0 Then
Dim stringa As String
 stringa = "SELECT * FROM Dettaglio WHERE codiceCliente = '" & variabile1 & "' AND [Numero Fattura]=0  and [numero DDT]=0 ORDER BY [DATA ORDINE]"

   With Adodc5
    .RecordSource = stringa
    .Refresh
    End With
    
    With DataGrid5
     .ClearFields
     .ReBind
     End With
     
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
     
     End If
     
txtclicode = variabile1
txtSupplier = variabile2
'variabile3 = 1 ' PREZZO DI LISTINO 1

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

txtprezzo = Format(art4, "###,###,##0.00")

DataCombo1 = art5
End Sub

Private Sub Command3_Click()
response = MsgBox("Vuoi cancellare questa riga?", vbOKCancel + vbCancel, "Attenzione")
Select Case response
 Case 6
 Adodc5.Recordset.delete
 DataGrid4.Columns(0) = DataGrid4.Columns(0) + Val(txtoldval)
Adodc2.Recordset.Update
 Case Else
 
 End Select

txtdesart = ""
txtcodart = ""

 txtpezzi = ""
 txtbase = ""
 txtaltezza = ""
  txtprezzo = ""
 DataCombo1 = ""

txtimportoiva = ""
txttotaleconiva = ""

 txtcodart = ""
 txtimportoriga = ""
 
Dim a, b, c As Integer
a = 0: b = 0: c = 0
Adodc5.Recordset.MoveFirst
 For I = 0 To Adodc5.Recordset.RecordCount - 1
  a = a + DataGrid5.Columns(16)
   b = b + DataGrid5.Columns(17)
    c = c + DataGrid5.Columns(29)
     Adodc5.Recordset.MoveNext
   Next I
   
      txttotimponibile = Format((a), "###,##0.00")
    txttotiva = Format((b), "###,##0.00")
     txttotordine = Format((c), "###,##0.00")
     
     
     
     txtoldval = txtpezzi


End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()

End Sub

Private Sub DataGrid5_Click()
'Command1.Enabled = False
On Error Resume Next

Check1.Enabled = True
  
 txtdesart = DataGrid5.Columns(1)
 txtSupplier = DataGrid5.Columns(2)
 txtordine = DataGrid5.Columns(3)
 txtdtaordine = DataGrid5.Columns(4)
 txtconsegna = DataGrid5.Columns(5)
 txtpezzi = DataGrid5.Columns(6)
 txtprezzo = DataGrid5.Columns(7)


DataCombo1.Text = DataGrid5.Columns(11)
 txtaliquota = DataGrid5.Columns(12)
 txtdestino = DataGrid5.Columns(14)
  txtrif = DataGrid5.Columns(15)
txtimportoriga = Format(DataGrid5.Columns(16), "###,##0.00")
  txtimportoiva = Format(DataGrid5.Columns(17), "###,##0.00")
    txttotaleconiva = Format(Val(txtimporto) + Val(txtimportoiva), "###,##0.00")
  txtclicode = DataGrid5.Columns(18)
     txtcodart = DataGrid5.Columns(20)
     
'aggancia articolo per giacenza
 Adodc2.Recordset.MoveFirst
  Adodc2.Recordset.Find "codiceinterno = '" & txtcodart & "'", 0, adSearchForward

txtbase = DataGrid5.Columns(26)
txtaltezza = DataGrid5.Columns(27)
txtquantità = Format(DataGrid5.Columns(28), "###,##0.00")
txttotaleconiva = Format(DataGrid5.Columns(29), "###,##0.00")

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
 stringa = "SELECT * FROM Dettaglio WHERE [numero ordine]='" & txtordine & "'"


   With Adodc4
    .RecordSource = stringa
    .Refresh
    End With
      With DataGrid3
     .ClearFields
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
   
      txttotimponibile = Format((a), "###,##0.00")
    txttotiva = Format((b), "###,##0.00")
     txttotordine = Format((c), "###,##0.00")
     
     
     
     txtoldval = txtpezzi
End Sub

Private Sub DataGrid5_DblClick()
Dim stringa As String
 stringa = "SELECT * FROM Dettaglio WHERE codiceCliente = '" & variabile1 & "' AND [Numero Ordine]='" & txtordine & "'"

   With Adodc5
    .RecordSource = stringa
    .Refresh
    End With
    
    With DataGrid5
     .ClearFields
     .ReBind
     End With
     
End Sub

Private Sub Form_Load()
    txtSupplier.Enabled = False
    txtclicode.Enabled = False
    Check1.Enabled = False
    
cmdmemo.Enabled = False
DataGrid5.Enabled = False


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

Private Sub txtdtaordine_Validate(Cancel As Boolean)
' txtdtaordine.Text = Format$(CDate(txtdtaordine.Text), "short date")

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
txtquantità = ""

On Error Resume Next

varia5 = txtprezzo



Select Case DataCombo1.Text
  
    Case "PZ", "ML", "MQ", "NR"
    
    varia1 = Val(txtpezzi)
        varia2 = varia1 * varia5
       varia4 = (varia2 / 100) * (txtaliquota)
          varia3 = varia2 + varia4
         
        txtbase.Enabled = False
        txtaltezza.Enabled = False
        txtdesart.SetFocus
          

          
          
    '  Case "ML"
    '   varia1 = Val(txtpezzi) * Val(txtaltezza) / 100
    'varia2 = varia1 * varia5
    '    varia4 = (varia2 / 100) * Val(txtaliquota)
    '     varia3 = varia2 + varia4
     '   txtbase.Enabled = False
    '      txtaltezza.Enabled = True
    '                txtaltezza.SetFocus
         
     ' Case "MQ"
     ' varia1 = Val(txtpezzi) * Val(txtbase) / 100 * Val(txtaltezza) / 100
     '   varia2 = varia1 * varia5
     '      varia4 = (varia2 / 100) * Val(txtaliquota)
     '     varia3 = varia2 + varia4
     '    txtbase.Enabled = True
      '  txtaltezza.Enabled = True
      '   txtbase.SetFocus
         
         
        Case Else
        varia1 = Val(txtpezzi)
        varia2 = varia1 * varia5
       varia4 = (varia2 / 100) * (txtaliquota)
          varia3 = varia2 + varia4
         
        txtbase.Enabled = False
        txtaltezza.Enabled = False
         txtdesart.SetFocus
        End Select
        txtquantità = varia1
        
       ' txtprezzo = varia5

  txtimportoriga = Format(varia2, "###,###,##0.00")
        txttotaleconiva = Format(varia3, "###,###,##0.00")
        txtimportoiva = Format(varia4, "###,###,##0.00")


End Sub

