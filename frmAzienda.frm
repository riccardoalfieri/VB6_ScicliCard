VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAzienda 
   BackColor       =   &H00808080&
   Caption         =   "Gestione sconto"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10785
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   10785
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5535
      Begin VB.TextBox Text6 
         DataField       =   "listino"
         DataSource      =   "Adodc1"
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   8760
         PasswordChar    =   "."
         TabIndex        =   25
         Text            =   "lingua"
         Top             =   7320
         Width           =   615
      End
      Begin VB.CommandButton cmddel 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Supprimer Factures"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   0
         Left            =   360
         Picture         =   "frmAzienda.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   5400
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmddel 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Supprimer Solde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   1
         Left            =   1800
         Picture         =   "frmAzienda.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   5400
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmddel 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Supprimer Caisse"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   2
         Left            =   3240
         Picture         =   "frmAzienda.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   5400
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         DataField       =   "menu"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   2520
         TabIndex        =   21
         Text            =   "Text3"
         Top             =   6000
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         DataField       =   "colonne"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Text            =   "Text3"
         Top             =   6000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         DataField       =   "righe"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   600
         TabIndex        =   16
         Text            =   "Text2"
         Top             =   6000
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Modifier Desktop.jpg"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   7200
         Picture         =   "frmAzienda.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5400
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Modifier Logo.bmp"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   5400
         Picture         =   "frmAzienda.frx":2120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5400
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdBack 
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
         Height          =   855
         Left            =   3000
         Picture         =   "frmAzienda.frx":3E62
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmddelete 
         Caption         =   "Supprimer"
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
         Left            =   7920
         Picture         =   "frmAzienda.frx":42A4
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Annuller"
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
         Height          =   615
         Left            =   7920
         Picture         =   "frmAzienda.frx":43EE
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2280
         Visible         =   0   'False
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
         Height          =   615
         Left            =   3000
         Picture         =   "frmAzienda.frx":4830
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2160
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
         Height          =   735
         Left            =   3000
         Picture         =   "frmAzienda.frx":497A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Nouvelle"
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
         Left            =   7920
         Picture         =   "frmAzienda.frx":4AC4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtiva 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Azienda"
         DataSource      =   "Adodc1"
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
         Left            =   1080
         TabIndex        =   1
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H8000000E&
         DataField       =   "Azienda"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         IMEMode         =   3  'DISABLE
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   6720
         Visible         =   0   'False
         Width           =   3495
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   240
         Top             =   7440
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
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
         RecordSource    =   "azienda"
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
         Bindings        =   "frmAzienda.frx":6EB6
         Height          =   1455
         Left            =   0
         TabIndex        =   3
         Top             =   6600
         Visible         =   0   'False
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   2566
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         ColumnCount     =   15
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
            DataField       =   "Azienda"
            Caption         =   "Azienda"
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
            DataField       =   "riga3"
            Caption         =   "riga3"
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
            DataField       =   "path"
            Caption         =   "path"
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
            DataField       =   "riga2"
            Caption         =   "riga2"
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
         BeginProperty Column08 
            DataField       =   "cap"
            Caption         =   "cap"
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
            DataField       =   "Pagamento"
            Caption         =   "Pagamento"
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
            DataField       =   "Provincia"
            Caption         =   "Provincia"
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
            DataField       =   "server"
            Caption         =   "server"
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
            DataField       =   "client"
            Caption         =   "client"
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
            DataField       =   "righe"
            Caption         =   "righe"
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
            DataField       =   "colonne"
            Caption         =   "colonne"
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
               ColumnWidth     =   2429,858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2429,858
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2429,858
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2429,858
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2429,858
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2429,858
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   2429,858
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   750,047
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   2429,858
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   2429,858
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   959,811
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   2429,858
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   2429,858
            EndProperty
         EndProperty
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Per modificare lo sconto: premi AGGIORNA, modifica la percentuale e conferma premendo SALVA"
         Height          =   615
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label Label11 
         Caption         =   "Label10"
         Height          =   255
         Left            =   8160
         TabIndex        =   26
         Top             =   7320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Columns"
         Height          =   255
         Left            =   2520
         TabIndex        =   20
         Top             =   5760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Menu"
         Height          =   255
         Left            =   2520
         TabIndex        =   19
         Top             =   5520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caisse Objets"
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   5520
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Columns"
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Top             =   5760
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rows"
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   5760
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000B&
         Caption         =   "Sconto praticato"
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
         Left            =   1080
         TabIndex        =   5
         Top             =   1320
         Width           =   1845
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "Intestazione Stampe"
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
         Index           =   1
         Left            =   -120
         TabIndex        =   4
         Top             =   7440
         Visible         =   0   'False
         Width           =   3465
      End
   End
End
Attribute VB_Name = "frmAzienda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mb





Private Sub cmdAdd_Click()
Adodc1.Recordset.AddNew

addupdate


End Sub

Private Sub cmdBack_Click()

Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me

End Sub








Private Sub cmddel_Click(Index As Integer)
Dim stringa, stringa1, stringa2, stringa3, response As String

response = MsgBox("Cancello gli Archivi?", vbOKCancel + vbCancel, "Attenzione")
Select Case response
 Case 6
  
    Select Case Index
    Case 0
     stringa = "dettaglio"
     stringa2 = "dettaglio"
     Case 1
      stringa = "pagamenti"
      stringa2 = "scadenzario"
     Case 2
      stringa = "cassa"
      stringa2 = "anamnesi"
     Case Else
     End Select
     
      stringa1 = "delete * from " & stringa
      rs.Open stringa1, cn
      
  stringa3 = "delete * from " & stringa2
      rs.Open stringa3, cn
 
     
 MsgBox "Archivi Cancellati ", vbExclamation

 Case Else
 
 End Select
End Sub

Private Sub cmdSave_Click()

SCONTO = txtiva
    Adodc1.Recordset.UpdateBatch
    savecancel
    delete

End Sub

Private Sub cmdUpdate_Click()
addupdate
End Sub

Private Sub Command1_Click()
On Error Resume Next
Shell "c:\windows\system32\mspaint.exe desktop.jpg", vbNormalFocus

End Sub

Private Sub Command2_Click()
On Error Resume Next
Shell "c:\windows\system32\mspaint.exe logo.bmp", vbNormalFocus

End Sub

Private Sub Form_Activate()
lingue
delete
End Sub

Private Function addupdate()
cmdAdd.Enabled = False
cmdUpdate.Enabled = False
cmdSave.Enabled = True
cmdcancel.Enabled = True
cmddelete.Enabled = False
'cmdReport.Enabled = False

txtName.Locked = False

'txtName.SetFocus


Adodc1.Enabled = False
DataGrid1.Enabled = False
End Function

Private Function savecancel()
DataGrid1.Refresh

cmdAdd.Enabled = True
cmdSave.Enabled = False
cmdcancel.Enabled = False


End Function

Private Function delete()
DataGrid1.Refresh

If Adodc1.Recordset.RecordCount = 0 Then
    Adodc1.Enabled = False
    DataGrid1.Enabled = False
    
    cmddelete.Enabled = False
    cmdUpdate.Enabled = False
 '   cmdReport.Enabled = False
Else
    Adodc1.Enabled = True
    DataGrid1.Enabled = True
    
    cmddelete.Enabled = True
    cmdUpdate.Enabled = True
   ' cmdReport.Enabled = True
End If

End Function

Private Sub Form_Unload(Cancel As Integer)
Adodc1.Recordset.CancelBatch
End Sub

Private Sub txtfiscale_LostFocus()
'controllaCF (txtfiscale)
End Sub



Private Sub txtiva_LostFocus()
If Val(txtiva) > 100 Or Val(txtiva) <= 0 Then
 mb = MsgBox("Errore percentuale sconto ", vbCritical, "Attentione")
    txtiva.SetFocus

End If

End Sub

Private Sub txtlistino_Validate(Cancel As Boolean)
End Sub

Public Sub lingue()


End Sub



