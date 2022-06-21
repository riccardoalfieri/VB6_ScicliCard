VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMARTICO 
   BackColor       =   &H80000003&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Produits"
   ClientHeight    =   9060
   ClientLeft      =   1095
   ClientTop       =   435
   ClientWidth     =   15630
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   15630
   Begin VB.CommandButton Command8 
      BackColor       =   &H0080C0FF&
      Caption         =   "       Publier    sur le net"
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
      Left            =   5880
      Picture         =   "frmARTICO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   7920
      Width           =   1815
   End
   Begin VB.TextBox txtcolor 
      Alignment       =   1  'Right Justify
      DataField       =   "bcolor"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   0
      EndProperty
      DataSource      =   "datPrimaryRS"
      Height          =   405
      Left            =   2040
      TabIndex        =   68
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "...."
      Height          =   615
      Left            =   2280
      Picture         =   "frmARTICO.frx":1D42
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   7080
      Width           =   375
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check3"
      DataField       =   "ingrediente"
      DataSource      =   "datPrimaryRS"
      Height          =   255
      Left            =   2040
      TabIndex        =   65
      Top             =   7200
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      DataField       =   "visibile"
      DataSource      =   "datPrimaryRS"
      Height          =   255
      Left            =   2040
      TabIndex        =   63
      Top             =   6720
      Width           =   255
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0000FFFF&
      Caption         =   "Categories"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      MaskColor       =   &H0080FFFF&
      Picture         =   "frmARTICO.frx":3A54
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Color Set"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   6240
      Width           =   1695
   End
   Begin VB.TextBox txtpic 
      DataField       =   "immagine"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   6360
      TabIndex        =   59
      Text            =   "Text2"
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text3 
      DataField       =   "Itemdesc"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   4680
      TabIndex        =   58
      Top             =   2640
      Width           =   2535
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "frmARTICO.frx":431E
      DataField       =   "categorie"
      DataSource      =   "datPrimaryRS"
      Height          =   315
      Left            =   2040
      TabIndex        =   57
      Top             =   5760
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Submenudesc"
      BoundColumn     =   "Submenudesc"
      Text            =   ""
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Mémorise Image"
      Height          =   495
      Left            =   5400
      TabIndex        =   55
      Top             =   6600
      Width           =   2535
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   5400
      TabIndex        =   54
      Top             =   5280
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   5400
      TabIndex        =   53
      Top             =   4080
      Width           =   2535
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   5400
      TabIndex        =   52
      Top             =   3840
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      DragMode        =   1  'Automatic
      Height          =   1875
      Left            =   7920
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      Picture         =   "frmARTICO.frx":4348
      ScaleHeight     =   1815
      ScaleWidth      =   1530
      TabIndex        =   51
      Top             =   0
      Width           =   1590
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      DataField       =   "service"
      DataSource      =   "datPrimaryRS"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Nouvelle Recherche"
      Height          =   615
      Left            =   3360
      Picture         =   "frmARTICO.frx":4FC6
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "...."
      Height          =   735
      Left            =   5400
      Picture         =   "frmARTICO.frx":5110
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "...."
      Height          =   1095
      Left            =   5400
      Picture         =   "frmARTICO.frx":6E22
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "PrezzoAcquistoNetto"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   1
      EndProperty
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   15
      Left            =   2040
      TabIndex        =   16
      Top             =   4680
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "listino3"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   1
      EndProperty
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   14
      Left            =   2040
      TabIndex        =   8
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "per3"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   1
      EndProperty
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   13
      Left            =   4560
      TabIndex        =   9
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "per1"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   1
      EndProperty
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   9
      Left            =   4560
      TabIndex        =   5
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "per2"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   1
      EndProperty
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   12
      Left            =   4560
      TabIndex        =   7
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "listino2"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   1
      EndProperty
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   6
      Top             =   1920
      Width           =   735
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmARTICO.frx":8B34
      DataField       =   "fornitore"
      DataSource      =   "datPrimaryRS"
      Height          =   315
      Left            =   2040
      TabIndex        =   13
      Top             =   3600
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Fornitore"
      BoundColumn     =   "CodiceFornitore"
      Text            =   ""
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3360
      TabIndex        =   19
      Top             =   6720
      Width           =   1935
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   0
      ScaleHeight     =   900
      ScaleWidth      =   15630
      TabIndex        =   33
      Top             =   7830
      Width           =   15630
      Begin VB.CommandButton cmdClose 
         Caption         =   "Quitter"
         Height          =   900
         Left            =   4320
         Picture         =   "frmARTICO.frx":8B49
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Ajourner"
         Height          =   900
         Left            =   3240
         Picture         =   "frmARTICO.frx":8F8B
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Supprimer"
         Height          =   900
         Left            =   2160
         Picture         =   "frmARTICO.frx":90D5
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Valider"
         Height          =   900
         Left            =   1080
         Picture         =   "frmARTICO.frx":921F
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Nouvelle"
         Height          =   900
         Left            =   0
         Picture         =   "frmARTICO.frx":9369
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "ScortaMinima"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   11
      Left            =   2040
      TabIndex        =   17
      Top             =   5040
      Width           =   615
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "PrezzoAcquisto"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   1
      EndProperty
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   10
      Left            =   2040
      TabIndex        =   15
      Top             =   4320
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "LottoRiordino"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   1
      EndProperty
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   8
      Left            =   4080
      TabIndex        =   18
      Top             =   5040
      Width           =   615
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "DataUltimoCarico"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   12
      Top             =   3300
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "DataUltimaVendita"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   11
      Top             =   2985
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "listino1"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   1
      EndProperty
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "Giacenza"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   1
      EndProperty
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   20
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox txtFields 
      DataField       =   "descrizione"
      DataSource      =   "datPrimaryRS"
      Height          =   750
      Index           =   2
      Left            =   2040
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   700
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CodiceInterno"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   2
      Top             =   360
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "codiceEAN"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   60
      Width           =   3375
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   8730
      Width           =   15630
      _ExtentX        =   27570
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
      Connect         =   "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;"
      OLEDBString     =   "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tires"
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Left            =   6000
      Top             =   3000
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   5520
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "frmARTICO.frx":B75B
      DataField       =   "misura"
      DataSource      =   "datPrimaryRS"
      Height          =   315
      Left            =   2040
      TabIndex        =   14
      Top             =   3960
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Pattern"
      Text            =   "PZ"
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "frmARTICO.frx":B770
      DataField       =   "iva"
      DataSource      =   "datPrimaryRS"
      Height          =   315
      Left            =   4080
      TabIndex        =   22
      Top             =   5400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Size"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   4200
      Top             =   5400
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
   Begin MSAdodcLib.Adodc Adosubmenu 
      Height          =   375
      Left            =   4200
      Top             =   5760
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
      RecordSource    =   "SELECT * FROM submenu"
      Caption         =   "AdoSubmenu"
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
   Begin MSComDlg.CommonDialog cdl 
      Left            =   120
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ingrédients"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   240
      TabIndex        =   66
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Visible"
      Height          =   255
      Index           =   22
      Left            =   240
      TabIndex        =   64
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ItemDesc"
      Height          =   255
      Index           =   21
      Left            =   2880
      TabIndex        =   62
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Categorie"
      Height          =   255
      Index           =   20
      Left            =   120
      TabIndex        =   56
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pizza"
      Height          =   255
      Index           =   19
      Left            =   120
      TabIndex        =   50
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Prix d'Achat Escompté"
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   41
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vente 3 HTC"
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   40
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Coeff.pur ottenir vente 3"
      Height          =   255
      Index           =   16
      Left            =   2760
      TabIndex        =   39
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Coeff.pur ottenir vente 1"
      Height          =   255
      Index           =   15
      Left            =   2760
      TabIndex        =   38
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Coeff.pur ottenir vente 2"
      Height          =   255
      Index           =   14
      Left            =   2760
      TabIndex        =   37
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TVA"
      Height          =   255
      Index           =   13
      Left            =   3360
      TabIndex        =   36
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vente 2 HTC"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   35
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Rechercher par codebarre"
      Height          =   255
      Left            =   3360
      TabIndex        =   34
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stock Maxi"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   32
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Prix d'Achat "
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   31
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UM"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   30
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Stock Mini"
      Height          =   255
      Index           =   8
      Left            =   2760
      TabIndex        =   29
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fournisseur"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   28
      Top             =   3615
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date Dernière Entrée"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   27
      Top             =   3300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date Dernière Sortie"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   26
      Top             =   2985
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vente 1 TTC"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   25
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Qté en Stock"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   24
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Designation avec Ingrédients"
      Height          =   735
      Index           =   2
      Left            =   120
      TabIndex        =   23
      Top             =   705
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Code Produit"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   21
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codes Barre"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "FRMARTICO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub Command1_Click()
Picture1.Picture = Nothing
Form6.Show vbModal
'On Error Resume Next
               datPrimaryRS.Recordset.MoveFirst
  datPrimaryRS.Recordset.Find "codiceinterno = '" & art1 & "'", 0, adSearchForward
 


End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text1.SetFocus

End Sub

Private Sub Command3_Click()
'On Error Resume Next
frmricerca.Show vbModal
        datPrimaryRS.Recordset.MoveFirst
  datPrimaryRS.Recordset.Find "codiceinterno = '" & art1 & "'", 0, adSearchForward


End Sub



Private Sub Command4_Click()
On Error GoTo dopo
Dim stringa As String
stringa = txtFields(1) & ".jpg"

SavePicture Picture1.Picture, stringa
txtpic = stringa
dopo:
End Sub

Private Sub Command5_Click()
On Error GoTo 100
    cdl.CancelError = True
    cdl.flags = cdlCCRGBInit
   ' cdl.Color = lbcolor.BackColor
    cdl.ShowColor
    Col = cdl.Color
    Command5.BackColor = cdl.Color
    
  txtcolor = cdl.Color
10
    Exit Sub
100
    Resume 10
End Sub

Private Sub Command6_Click()
On Error Resume Next
Form5.Show vbModal
'On Error Resume Next
               datPrimaryRS.Recordset.MoveFirst
  datPrimaryRS.Recordset.Find "codiceinterno = '" & art1 & "'", 0, adSearchForward
 

End Sub

Private Sub Command7_Click()
On Error Resume Next
frmSubmenu.Show vbModal
End Sub

Private Sub Command8_Click()
campo(0) = Text3     ' descrizione breve
campo(1) = txtFields(1)  'codice
campo(2) = txtFields(2)   ' descrizione lunga
campo(3) = DataCombo3   ' iva
campo(4) = DataCombo4 '   reparto
 valore1 = txtFields(4) ' prezzo
 art10 = txtcolor  'colore
 
If Check1.Value = 1 Then   ' pizza
 flags(1) = True
  Else: flags(1) = False
  End If
  
  If Check3.Value = 1 Then   'ingrediente
   flags(2) = True
  Else: flags(2) = False
  End If
 
  
If azienda <> "" Then
frm_on_net.Show vbModal

Else
MsgBox messaggi(1), vbOKOnly + vbExclamation, messaggi(2)
End If

End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
'rtb.Text = ""
Picture1.Cls

File1.Pattern = "*.jpg;*.bmp;*.gif"

End Sub

Private Sub Drive1_Change()
On Error GoTo errhnd
Dir1.Path = Drive1.Drive
errhnd:
Select Case Err.Number
Dim msg As String
Case 68
msg = "your this drive are not resource" & vbNewLine
msg = msg + "1. try or look your cdrom or floppy disk drive" & vbNewLine
msg = msg + " 2. resource are not available " & vbNewLine
msg = msg + " 3. we are set drive default is c:\"
MsgBox msg, vbOKOnly + vbExclamation, "please check"
 Drive1.Drive = "C:\"
 End Select
End Sub

Private Sub File1_Click()
On Error Resume Next
Picture1.Picture = LoadPicture(Dir1.Path & "\" & File1.FileName)

End Sub

Private Sub Form_Activate()
'txtFields(16) = variabile1
File1.Pattern = "*.jpg;*.bmp;*.gif"

lingue
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
  File1.Pattern = "*.jpg;*.bmp;*.gif"

End Sub

Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'Posizione in cui inserire il codice per la gestione degli errori
  'Per ignorare gli errori, impostare come commento la riga seguente
  'Per intercettare gli errori, inserire il codice per la gestione degli errori in questa posizione
  MsgBox "Data error event hit err:" & Description
End Sub

Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Visualizza la posizione del record corrente per questo gruppo di record
 ' datPrimaryRS.Caption = "Record: " & CStr(datPrimaryRS.Recordset.AbsolutePosition)
End Sub

Private Sub datPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Posizione in cui inserire il codice per la convalida
  'L'evento viene richiamato in seguito alle seguenti azioni
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  datPrimaryRS.Recordset.AddNew
 
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  Dim response As String

response = MsgBox("Elimino l'articolo?", vbOKCancel + vbCancel, "Attenzione")
Select Case response
 Case 6
 With datPrimaryRS.Recordset
    .delete
    .MoveNext
    If .EOF Then .MoveLast
  End With

 Case Else
End Select
  
   
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'Necessario solo per applicazioni multiutente
  On Error GoTo RefreshErr
  datPrimaryRS.Refresh
  adoitems.Refresh
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  datPrimaryRS.Recordset.UpdateBatch adAffectAll
 
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Text1_KeyDown( _
           KeyCode As Integer, Shift As Integer)
           
On Error Resume Next

     Select Case KeyCode
     Case vbKeyReturn:
               datPrimaryRS.Recordset.MoveFirst
  datPrimaryRS.Recordset.Find "codiceean = '" & Text1.Text & "'", 0, adSearchForward




     End Select

End Sub


Private Sub txtFields_Change(Index As Integer)
Dim numero1, numero2, numero3, numero4 As Single



On Error Resume Next

     
  numero1 = (txtFields(10))
  numero2 = (txtFields(12))
  numero3 = (txtFields(13))
  numero4 = (txtFields(9))
  
  
 Select Case Index
         Case 0, 1, 2
          stringa = txtFields(1) & ".jpg"
 Picture1.Picture = LoadPicture(stringa)
  
               
  
        Case 10, 9, 12, 13
        
         If txtFields(12) > 0 Then
         txtFields(7) = Format(((numero1 * (numero2) / 100)) + (numero1), "####0.00")
             End If
             
              If txtFields(13) > 0 Then
         txtFields(14) = Format(((numero1 * (numero3) / 100)) + (numero1), "####0.00")
             End If
             
              If txtFields(9) > 0 Then
         txtFields(4) = Format(((numero1 * (numero4) / 100)) + (numero1), "####0.00")
             End If
             
     Case Else
     End Select
    
End Sub


Public Sub lingue()

Select Case LINGUA
  
 Case Is = "1F"
 
messaggi(1) = "Connexion au server n'établi pas "
messaggi(2) = "Attention"
   
   Case Is = "2I"
lblLabels(0) = "Barcode"
lblLabels(1) = "Codice"
lblLabels(2) = "Prodotto"
lblLabels(3) = "Qtà in Stock"
lblLabels(4) = "Prezzo di Vendita"
lblLabels(12) = "Prezzo di Vendita 2"
lblLabels(15) = "% di ricarico"
lblLabels(17) = "Servizio"
lblLabels(16) = "Carta Fedeltà"
lblLabels(5) = "Data Ultima Vendita"
lblLabels(6) = "Data Ultimo Acquisto"
lblLabels(7) = "Fornitore"
lblLabels(8) = "Qtà Riordino"
lblLabels(9) = "Unità di Misura"
lblLabels(10) = "Costo"
lblLabels(18) = "Costo scontato"
lblLabels(11) = "Qtà Minima"
lblLabels(13) = "IVA"
Label1 = "Ricerca per Barcode"
lblLabels(20) = "Categorie"
lblLabels(22) = "Visibile"
lblLabels(23) = "Ingredienti"

Command2.Caption = "Nuova Ricerca"
cmdAdd.Caption = "Nuovo"
cmdUpdate.Caption = "Salva"
cmdDelete.Caption = "Elimina"
cmdRefresh.Caption = "Aggiorna"
cmdClose.Caption = "Esci"

Command4.Caption = " Memorizza Immagine"
FRMARTICO.Caption = "Archivo Articoli"
Command7.Caption = "Categorie"

messaggi(1) = "Connessione al server non attiva"
  messaggi(2) = "Attenzione"
  

 Case Is = "3G"
 lblLabels(0) = "Barcode"
lblLabels(1) = "Code"
lblLabels(2) = "Product"
lblLabels(3) = "Qty in Stock"
lblLabels(4) = "Sell Price"
lblLabels(15) = "% on cost"
lblLabels(17) = "Service"
lblLabels(16) = "Fidelity Card"
lblLabels(5) = "Last Sell Date"
lblLabels(6) = "Last Entry Date"
lblLabels(7) = "Vendor"
lblLabels(8) = "Reorder Qty"
lblLabels(9) = "Misure Unit"
lblLabels(10) = "Cost"
lblLabels(18) = "Discount Cost"
lblLabels(11) = "Reorder Level"
lblLabels(13) = "Tax"
Label1 = "Barcode Search"

Command2.Caption = "New Search"
cmdAdd.Caption = "New"
cmdUpdate.Caption = "Save"
cmdDelete.Caption = "Delete"
cmdRefresh.Caption = "Update"
cmdClose.Caption = "Exit"

Command4.Caption = "Save Image"
FRMARTICO.Caption = "Products"

Case Is = "4S"
 lblLabels(0) = "Barcode"
lblLabels(1) = "Codigo"
lblLabels(2) = "Producto"
lblLabels(3) = "Cantidad en Stock"
lblLabels(4) = "Precio"
lblLabels(15) = "% de recarga"
lblLabels(17) = "Servicio"
lblLabels(16) = "Tarjeta Fidelidad"
lblLabels(5) = "Fecha de Ultima Venta"
lblLabels(6) = "Fecha de Ultimo Ingreso"
lblLabels(7) = "Proveedor"
lblLabels(8) = "Cantidad de Pedido"
lblLabels(9) = "Unidad"
lblLabels(10) = "Coste"
lblLabels(18) = "Coste Descontado"
lblLabels(11) = "Existencia Minima"
lblLabels(13) = "Iva"
Label1 = "Busqueda por Codigo Barreado"

Command2.Caption = "Nueva Busqueda"
cmdAdd.Caption = "Nuevo"
cmdUpdate.Caption = "Guarda"
cmdDelete.Caption = "Borra"
cmdRefresh.Caption = "Actualiza"
cmdClose.Caption = "Salir"

Command4.Caption = "Guarda Imagen"
FRMARTICO.Caption = "Productos"

Case Else
End Select
 
End Sub


