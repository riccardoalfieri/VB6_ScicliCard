VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmarticoli 
   Caption         =   "Archivio Articoli"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   14633.86
   ScaleMode       =   0  'User
   ScaleWidth      =   29307.7
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13095
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmarticoli.frx":0000
         Height          =   975
         Left            =   240
         TabIndex        =   40
         Top             =   8040
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   1720
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
      Begin VB.CommandButton Command1 
         Caption         =   "RIC ERCA"
         Height          =   375
         Left            =   1680
         TabIndex        =   39
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   840
         TabIndex        =   38
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text5 
         DataField       =   "DataUltimoCarico"
         DataSource      =   "Adodc4"
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
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   7560
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         DataField       =   "fornitore"
         DataSource      =   "Adodc4"
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
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   7200
         Width           =   3495
      End
      Begin VB.TextBox Text3 
         DataField       =   "PrezzoAcquisto"
         DataSource      =   "Adodc4"
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
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   6840
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmarticoli.frx":0015
         DataField       =   "misura"
         DataSource      =   "Adodc4"
         Height          =   315
         Left            =   4680
         TabIndex        =   7
         Top             =   4080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Pattern"
         Text            =   ""
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         DataField       =   "codiceEAN"
         DataSource      =   "Adodc4"
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
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtprice3 
         DataField       =   "listino1"
         DataSource      =   "Adodc4"
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
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox txtprice2 
         DataField       =   "giorni"
         DataSource      =   "Adodc4"
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
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   3000
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         DataField       =   "descrizione"
         DataSource      =   "Adodc4"
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
         Left            =   4680
         TabIndex        =   3
         Top             =   2640
         Width           =   3495
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H00E0E0E0&
         DataField       =   "CodiceInterno"
         DataSource      =   "Adodc4"
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
         Left            =   4680
         TabIndex        =   2
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "CHIUDI"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11400
         TabIndex        =   21
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Elimina"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11400
         TabIndex        =   20
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancella"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11400
         TabIndex        =   19
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Salva"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11400
         TabIndex        =   18
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Aggiorna"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11400
         TabIndex        =   17
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Nuovo"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11400
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtMax 
         BackColor       =   &H00E0E0E0&
         DataField       =   "LottoRiordino"
         DataSource      =   "Adodc4"
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
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   6480
         Width           =   1935
      End
      Begin VB.TextBox txtQuantity 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Giacenza"
         DataSource      =   "Adodc4"
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
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   5760
         Width           =   1935
      End
      Begin VB.TextBox txtMin 
         BackColor       =   &H00E0E0E0&
         DataField       =   "ScortaMinima"
         DataSource      =   "Adodc4"
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
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   6120
         Width           =   1935
      End
      Begin VB.TextBox txtPrice1 
         DataField       =   "pallet"
         DataSource      =   "Adodc4"
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
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   3720
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   495
         Left            =   7920
         Top             =   6120
         Visible         =   0   'False
         Width           =   1440
         _ExtentX        =   2540
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
         RecordSource    =   "SELECT * FROM Tires ORDER BY Descrizione"
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   480
         Top             =   3960
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   480
         Top             =   4320
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
         RecordSource    =   "Sizes"
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
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frmarticoli.frx":002A
         DataField       =   "Iva"
         DataSource      =   "Adodc4"
         Height          =   315
         Left            =   4680
         TabIndex        =   8
         Top             =   4440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Size"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   375
         Left            =   480
         Top             =   4680
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
         RecordSource    =   "reparti"
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
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "frmarticoli.frx":003F
         DataField       =   "Varieta"
         DataSource      =   "Adodc4"
         Height          =   315
         Left            =   4680
         TabIndex        =   9
         Top             =   4800
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Size"
         Text            =   ""
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Ultimo Carico"
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
         Left            =   3240
         TabIndex        =   37
         Top             =   7560
         Width           =   1410
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Fornitore"
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
         Left            =   3480
         TabIndex        =   36
         Top             =   7200
         Width           =   1245
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Prezzo di Acquisto"
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
         Left            =   2880
         TabIndex        =   35
         Top             =   6840
         Width           =   1815
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Codice EAN"
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
         Left            =   3480
         TabIndex        =   34
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label LBLPALLET 
         AutoSize        =   -1  'True
         Caption         =   "Confezioni"
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
         Left            =   3480
         TabIndex        =   33
         Top             =   3720
         Width           =   1245
      End
      Begin VB.Label lblprice3 
         AutoSize        =   -1  'True
         Caption         =   "Prezzo di Vendita "
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
         Left            =   2880
         TabIndex        =   32
         Top             =   3360
         Width           =   2025
      End
      Begin VB.Label lblprice2 
         AutoSize        =   -1  'True
         Caption         =   "Giorni Approvviggionamento"
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
         Left            =   1920
         TabIndex        =   31
         Top             =   3000
         Width           =   2970
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Descrizione Prodotto"
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
         Left            =   2520
         TabIndex        =   30
         Top             =   2640
         Width           =   2280
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codice Articolo"
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
         Left            =   3120
         TabIndex        =   29
         Top             =   2280
         Width           =   1605
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Giacenza"
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
         Left            =   3720
         TabIndex        =   28
         Top             =   5760
         Width           =   1005
      End
      Begin VB.Label lblprice1 
         AutoSize        =   -1  'True
         Caption         =   "Prezzo di Vendita 1"
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
         Left            =   240
         TabIndex        =   27
         Top             =   -360
         Width           =   1905
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Unità di Misura"
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
         Left            =   3240
         TabIndex        =   26
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Aliquota"
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
         Left            =   3840
         TabIndex        =   25
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Reparto"
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
         Left            =   3840
         TabIndex        =   24
         Top             =   4800
         Width           =   885
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   4680
         Picture         =   "frmarticoli.frx":0054
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Scorta Minima"
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
         Left            =   3120
         TabIndex        =   23
         Top             =   6120
         Width           =   1755
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Lotto di riordino"
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
         Left            =   2880
         TabIndex        =   22
         Top             =   6480
         Width           =   2160
      End
   End
End
Attribute VB_Name = "frmarticoli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
Unload Me
End Sub

Private Sub Text6_Change()
Adodc4.Recordset.MoveFirst
  Adodc4.Recordset.Find "codiceinterno = '" & Text6.Text & "'", 0, adSearchForward
 
 
End Sub
