VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMARTICO2 
   BackColor       =   &H80000002&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Produits"
   ClientHeight    =   7305
   ClientLeft      =   1095
   ClientTop       =   435
   ClientWidth     =   7590
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   7590
   Begin VB.CommandButton Command2 
      Caption         =   "Nouvelle Recherche"
      Height          =   615
      Left            =   5640
      Picture         =   "frmARTICO2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "...."
      Height          =   735
      Left            =   5400
      Picture         =   "frmARTICO2.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "...."
      Height          =   735
      Left            =   5400
      Picture         =   "frmARTICO2.frx":1E5C
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   6480
      TabIndex        =   42
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "recupera articoli "
      Height          =   495
      Left            =   6360
      TabIndex        =   41
      Top             =   480
      Visible         =   0   'False
      Width           =   735
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
      TabIndex        =   17
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
      TabIndex        =   9
      Top             =   2040
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
      Left            =   5160
      TabIndex        =   10
      Top             =   2040
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
      Left            =   5160
      TabIndex        =   6
      Top             =   1320
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
      Left            =   5160
      TabIndex        =   8
      Top             =   1680
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
      TabIndex        =   7
      Top             =   1680
      Width           =   735
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmARTICO2.frx":3B6E
      DataField       =   "fornitore"
      DataSource      =   "datPrimaryRS"
      Height          =   315
      Left            =   2040
      TabIndex        =   13
      Top             =   3240
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "CodiceFornitore"
      BoundColumn     =   "CodiceFornitore"
      Text            =   ""
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5640
      TabIndex        =   20
      Top             =   4920
      Width           =   1815
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   0
      ScaleHeight     =   900
      ScaleWidth      =   7590
      TabIndex        =   32
      Top             =   6075
      Width           =   7590
      Begin VB.CommandButton cmdClose 
         Caption         =   "Quitter"
         Height          =   900
         Left            =   4320
         Picture         =   "frmARTICO2.frx":3B83
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Ajouter"
         Height          =   900
         Left            =   3240
         Picture         =   "frmARTICO2.frx":3FC5
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Supprimer"
         Height          =   900
         Left            =   2160
         Picture         =   "frmARTICO2.frx":410F
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Valider"
         Height          =   900
         Left            =   1080
         Picture         =   "frmARTICO2.frx":4259
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Nouvelle"
         Height          =   900
         Left            =   0
         Picture         =   "frmARTICO2.frx":43A3
         Style           =   1  'Graphical
         TabIndex        =   43
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
      TabIndex        =   18
      Top             =   5040
      Width           =   3375
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
      TabIndex        =   16
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
      Left            =   2040
      TabIndex        =   14
      Top             =   3585
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "DataUltimoCarico"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   12
      Top             =   2940
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
      Top             =   2625
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
      TabIndex        =   5
      Top             =   1320
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
      TabIndex        =   4
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "descrizione"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   2040
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
      Top             =   380
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
      Visible         =   0   'False
      Width           =   3375
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   6975
      Width           =   7590
      _ExtentX        =   13388
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
      RecordSource    =   $"frmARTICO2.frx":6795
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
      Left            =   5760
      Top             =   3240
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
      Bindings        =   "frmARTICO2.frx":6877
      DataField       =   "misura"
      DataSource      =   "datPrimaryRS"
      Height          =   315
      Left            =   2040
      TabIndex        =   15
      Top             =   3960
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Pattern"
      Text            =   "PZ"
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "frmARTICO2.frx":688C
      DataField       =   "iva"
      DataSource      =   "datPrimaryRS"
      Height          =   315
      Left            =   2040
      TabIndex        =   19
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
   Begin VB.Label lblLabels 
      Caption         =   "Prix d'Achat Escompté"
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   40
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Vente 3"
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   39
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Coeff.pur ottenir vente 3"
      Height          =   255
      Index           =   16
      Left            =   3240
      TabIndex        =   38
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Coeff.pur ottenir vente 1"
      Height          =   255
      Index           =   15
      Left            =   3240
      TabIndex        =   37
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Coeff.pur ottenir vente 2"
      Height          =   255
      Index           =   14
      Left            =   3240
      TabIndex        =   36
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TVA"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   35
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Vente 2"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   34
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Rechercher"
      Height          =   255
      Left            =   5640
      TabIndex        =   33
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Stock Maxi"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   31
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Prix d'Achat "
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   30
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "UM"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   29
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Stock Mini"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   28
      Top             =   3585
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fournisseur"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   27
      Top             =   3255
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Date Dernière Entrée"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   26
      Top             =   2940
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Date Dernière Sortie"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   25
      Top             =   2625
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Vente 1"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   24
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Qté en Stock"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   23
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Designation"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   22
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Code Produit"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   21
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codes Barre"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "FRMARTICO2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form6.Show vbModal
On Error Resume Next
               datPrimaryRS.Recordset.MoveFirst
  datPrimaryRS.Recordset.Find "codiceinterno = '" & art1 & "'", 0, adSearchForward

End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text1.SetFocus

End Sub

Private Sub Command3_Click()
On Error Resume Next
frmricerca.Show vbModal
        datPrimaryRS.Recordset.MoveFirst
  datPrimaryRS.Recordset.Find "codiceinterno = '" & art1 & "'", 0, adSearchForward

End Sub

Private Sub Command4_Click()

On Error Resume Next


Dim b As String
Dim a As String * 330

 fnum = FreeFile()
   
Open "c:\magazzino.txt" For Input As #2
 rs.Open "select * from tires", cn, adOpenDynamic, adLockOptimistic
  
 'Input #1, a, b, c, d, e, f, g, h
 
 For i = 1 To LOF(2) / 330
 Line Input #2, a

 rs.AddNew
 rs!codiceinterno = Mid(a, 3, 13)
 rs!descrizione = Mid(a, 18, 30)
 rs!misura = Mid(a, 127, 2)
 rs!fornitore = Mid(a, 139, 30)
 rs!iva = Mid(a, 169, 2)
 rs!giacenza = Val(Mid(a, 119, 5))
 
 If Mid(a, 173, 1) <> " " Then
  b = Mid(a, 173, 3) & "," & Mid(a, 177, 2)
        Else
    If Mid(a, 174, 1) <> " " Then
    b = Mid(a, 174, 2) & "," & Mid(a, 177, 2)
       Else
 b = Mid(a, 175, 1) & "," & Mid(a, 177, 2)
 
       End If
       End If
       
 rs!prezzoacquisto = b
 
 If Mid(a, 183, 1) <> " " Then
  b = Mid(a, 183, 3) & "," & Mid(a, 187, 2)
        Else
    If Mid(a, 184, 1) <> " " Then
    b = Mid(a, 184, 2) & "," & Mid(a, 187, 2)
       Else
 b = Mid(a, 185, 1) & "," & Mid(a, 187, 2)
 
       End If
       End If
 rs!listino1 = b / 1.2
 
 If Mid(a, 193, 1) <> " " Then
  b = Mid(a, 193, 3) & "," & Mid(a, 197, 2)
        Else
    If Mid(a, 194, 1) <> " " Then
    b = Mid(a, 194, 2) & "," & Mid(a, 197, 2)
       Else
 b = Mid(a, 195, 1) & "," & Mid(a, 197, 2)
 
       End If
       End If
       
       rs!listino2 = b / 1.2
 
rs.Update


Next i

Close #2

End Sub

Private Sub Command5_Click()
rs.Open "update tires set listino1=listino1/1.2, listino2=listino2/1.2,listino3=listino3/1.2", cn, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'Posizione in cui inserire il codice per la gestione degli errori
  'Per ignorare gli errori, impostare come commento la riga seguente
  'Per intercettare gli errori, inserire il codice per la gestione degli errori in questa posizione
  MsgBox "Data error event hit err:" & Description
End Sub

Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Visualizza la posizione del record corrente per questo gruppo di record
  datPrimaryRS.Caption = "Record: " & CStr(datPrimaryRS.Recordset.AbsolutePosition)
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
  With datPrimaryRS.Recordset
    .delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'Necessario solo per applicazioni multiutente
  On Error GoTo RefreshErr
  datPrimaryRS.Refresh
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

     Select Case KeyCode
     Case vbKeyReturn:
               datPrimaryRS.Recordset.MoveFirst
  datPrimaryRS.Recordset.Find "codiceinterno = '" & Text1.Text & "'", 0, adSearchForward

     End Select

End Sub


Private Sub txtFields_Change(Index As Integer)

Dim numero1, numero2, numero3, numero4 As Single

On Error Resume Next
 Select Case DataCombo3
  Case "20"
   valoreiva% = 20
    valoreinverso% = 1.2
    Case "10"
     valoreiva% = 10
    valoreinverso% = 1.1
      Case "4", "04"
      valoreiva% = 4
    valoreinverso% = 1.04
     Case Else
     End Select
     
  numero1 = (txtFields(10))
  numero2 = (txtFields(12))
  numero3 = (txtFields(13))
  numero4 = (txtFields(9))
  
  
 Select Case Index
 
  
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
