VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_on_net 
   BackColor       =   &H00FF8080&
   Caption         =   "Menu sur le net"
   ClientHeight    =   9345
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   12060
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frm_on_net.frx":0000
      Height          =   7935
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   13996
      _Version        =   393216
      AllowUpdate     =   0   'False
      ColumnHeaders   =   0   'False
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
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "azienda"
         Caption         =   "azienda"
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
         DataField       =   "codice"
         Caption         =   "codice"
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
         DataField       =   "itemdesc"
         Caption         =   "itemdesc"
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
      BeginProperty Column04 
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
      BeginProperty Column05 
         DataField       =   "pizza"
         Caption         =   "pizza"
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
         DataField       =   "ingrediente"
         Caption         =   "ingrediente"
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
         DataField       =   "bcolor"
         Caption         =   "bcolor"
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
      BeginProperty Column09 
         DataField       =   "categoria"
         Caption         =   "categoria"
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
         DataField       =   "chiave"
         Caption         =   "chiave"
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
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
            ColumnWidth     =   915,024
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "BColor"
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
      Height          =   285
      Index           =   16
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Menu complet"
      Height          =   855
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FF80FF&
      Caption         =   "Supprimer"
      Height          =   900
      Left            =   4920
      Picture         =   "frm_on_net.frx":0015
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7920
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      DataField       =   "chiave"
      DataSource      =   "Adodc2"
      Height          =   495
      Left            =   9600
      TabIndex        =   15
      Top             =   7440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Quitter"
      Height          =   900
      Left            =   8400
      Picture         =   "frm_on_net.frx":015F
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00FFFF80&
      Caption         =   "Sauver sur le net"
      Height          =   900
      Left            =   3000
      Picture         =   "frm_on_net.frx":05A1
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7920
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      DataField       =   "Itemdesc"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   7920
      TabIndex        =   11
      Top             =   5160
      Width           =   2535
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CodiceInterno"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   5
      Top             =   5160
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "descrizione"
      DataSource      =   "datPrimaryRS"
      Height          =   750
      Index           =   2
      Left            =   2040
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   5520
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "listino1"
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
      Index           =   4
      Left            =   2040
      TabIndex        =   3
      Top             =   6360
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      DataField       =   "service"
      DataSource      =   "datPrimaryRS"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   7200
      Width           =   255
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check3"
      DataField       =   "ingrediente"
      DataSource      =   "datPrimaryRS"
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   6840
      Width           =   255
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm_on_net.frx":06EB
      Height          =   855
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   1508
      _Version        =   393216
      AllowUpdate     =   0   'False
      ColumnHeaders   =   0   'False
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
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "azienda"
         Caption         =   "azienda"
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
         DataField       =   "codice"
         Caption         =   "codice"
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
         DataField       =   "itemdesc"
         Caption         =   "itemdesc"
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
      BeginProperty Column04 
         DataField       =   "listino1"
         Caption         =   "listino1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "pizza"
         Caption         =   "pizza"
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
         DataField       =   "ingrediente"
         Caption         =   "ingrediente"
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
         DataField       =   "bcolor"
         Caption         =   "bcolor"
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
      BeginProperty Column09 
         DataField       =   "categoria"
         Caption         =   "categoria"
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
         DataField       =   "chiave"
         Caption         =   "chiave"
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
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
            ColumnWidth     =   915,024
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   9015
      Visible         =   0   'False
      Width           =   12060
      _ExtentX        =   21273
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
      Connect         =   $"frm_on_net.frx":0700
      OLEDBString     =   $"frm_on_net.frx":078D
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from menu "
      Caption         =   " adodc2"
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
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "frm_on_net.frx":081A
      DataField       =   "categorie"
      DataSource      =   "datPrimaryRS"
      Height          =   315
      Left            =   7920
      TabIndex        =   18
      Top             =   6360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      Locked          =   -1  'True
      ListField       =   "Submenudesc"
      BoundColumn     =   "Submenudesc"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "frm_on_net.frx":0844
      DataField       =   "iva"
      DataSource      =   "datPrimaryRS"
      Height          =   315
      Left            =   7920
      TabIndex        =   19
      Top             =   5880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      _Version        =   393216
      Locked          =   -1  'True
      ListField       =   "Size"
      Text            =   ""
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Color"
      Height          =   255
      Index           =   0
      Left            =   6120
      TabIndex        =   23
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TVA"
      Height          =   255
      Index           =   13
      Left            =   6120
      TabIndex        =   21
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Categorie"
      Height          =   375
      Index           =   20
      Left            =   6120
      TabIndex        =   20
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ItemDesc"
      Height          =   255
      Index           =   21
      Left            =   6120
      TabIndex        =   12
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Code Produit"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   5175
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Designation avec Ingrédients"
      Height          =   735
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Prix de vente"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pizza"
      Height          =   255
      Index           =   19
      Left            =   120
      TabIndex        =   7
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFFF00&
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
      Left            =   120
      TabIndex        =   6
      Top             =   6840
      Width           =   1695
   End
End
Attribute VB_Name = "frm_on_net"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
 On Error GoTo DeleteErr
  Dim response As String

response = MsgBox("Supprimer ce article?", vbOKCancel + vbCancel, "Attention")
Select Case response
 Case 6
 With Adodc2.Recordset
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

Private Sub cmdUpdate_Click()
If Text1.Text = "" Then
Adodc2.Recordset.AddNew
End If

DataGrid1.Columns(0) = azienda
DataGrid1.Columns(1) = txtFields(1)
DataGrid1.Columns(2) = Text3
DataGrid1.Columns(3) = txtFields(2)
DataGrid1.Columns(4) = valore1
DataGrid1.Columns(5) = Check1.Value
DataGrid1.Columns(6) = Check3.Value
DataGrid1.Columns(7) = art10
DataGrid1.Columns(8) = campo(3)
DataGrid1.Columns(9) = campo(4)

Adodc2.Recordset.Update


End Sub

Private Sub Command1_Click()
Dim stringa As String

DataGrid2.Visible = True
cmdUpdate.Visible = False

stringa = "select * from menu where azienda='" & azienda & "' order by categoria"

With Adodc2
 .RecordSource = stringa
 .Refresh
 End With
 
 With DataGrid2
  .ReBind
  End With
  
End Sub

Private Sub Form_Load()
Dim stringa As String

On Error GoTo sotto
DataGrid2.Visible = False

stringa = "select * from menu where codice='" & campo(1) & "' and azienda='" & azienda & "'"

With Adodc2
 .RecordSource = stringa
 .Refresh
 End With
 
 With DataGrid1
  .ReBind
  End With
  

 Text3 = campo(0)
 txtFields(1) = campo(1)
txtFields(2) = campo(2)
DataCombo3 = campo(3)
DataCombo4 = campo(4)

If flags(1) = True Then
Check1.Value = 1
  Else: Check1.Value = 0
  End If
  
 If flags(2) = True Then
Check3.Value = 1
  Else: Check3.Value = 0
  End If
  txtFields(4) = valore1
   txtFields(16) = art10


sotto:

End Sub
