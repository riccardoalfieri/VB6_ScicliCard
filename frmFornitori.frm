VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmFornitori 
   BackColor       =   &H00808080&
   Caption         =   "Fournisseurs"
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
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   0
      TabIndex        =   10
      Top             =   -120
      Width           =   9975
      Begin VB.CommandButton Command1 
         Caption         =   "...."
         Height          =   615
         Left            =   3840
         Picture         =   "frmFornitori.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   720
         Width           =   615
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
         Left            =   7800
         Picture         =   "frmFornitori.frx":1D12
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Ajourner"
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
         Left            =   7800
         Picture         =   "frmFornitori.frx":4104
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Valider"
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
         Left            =   7800
         Picture         =   "frmFornitori.frx":424E
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1560
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
         Left            =   7800
         Picture         =   "frmFornitori.frx":4398
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
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
         Left            =   7800
         Picture         =   "frmFornitori.frx":47DA
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdBack 
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
         Left            =   7800
         Picture         =   "frmFornitori.frx":4924
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtsconto1 
         Alignment       =   1  'Right Justify
         DataField       =   "sconto1"
         DataSource      =   "Adodc1"
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
         Left            =   240
         TabIndex        =   7
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox txtsconto2 
         Alignment       =   1  'Right Justify
         DataField       =   "sconto2"
         DataSource      =   "Adodc1"
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
         Left            =   1200
         TabIndex        =   8
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox txtsconto3 
         Alignment       =   1  'Right Justify
         DataField       =   "sconto3"
         DataSource      =   "Adodc1"
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
         Left            =   2160
         TabIndex        =   9
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox txtcap 
         DataField       =   "cap"
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
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   360
         MaxLength       =   50
         TabIndex        =   3
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox txtlistino 
         DataField       =   "listino"
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
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   360
         MaxLength       =   50
         TabIndex        =   6
         Top             =   4440
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox TXTCODE 
         DataField       =   "CodiceFornitore"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtCittà 
         BackColor       =   &H8000000B&
         DataField       =   "città"
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
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   360
         MaxLength       =   100
         TabIndex        =   2
         Top             =   2040
         Width           =   3495
      End
      Begin VB.TextBox txtiva 
         BackColor       =   &H00FFFFFF&
         DataField       =   "PartitaIva"
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
         Left            =   360
         TabIndex        =   4
         Top             =   3240
         Width           =   2895
      End
      Begin VB.TextBox txtNumber 
         DataField       =   "NumeroTelefonico"
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
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   360
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         Top             =   3840
         Width           =   2055
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H8000000B&
         DataField       =   "Indirizzo"
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
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   360
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   1
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H8000000B&
         DataField       =   "Fornitore"
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
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   360
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   0
         Top             =   840
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
         RecordSource    =   "SELECT * FROM FORNITORI ORDER BY FORNITORE"
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
         Bindings        =   "frmFornitori.frx":4D66
         Height          =   2415
         Left            =   240
         TabIndex        =   11
         Top             =   5760
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   4260
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   0   'False
         HeadLines       =   1
         RowHeight       =   19
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmFornitori.frx":4D7B
         DataField       =   "tipopagamento"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   3840
         TabIndex        =   23
         Top             =   5280
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "tipoPagamento"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   495
         Left            =   5520
         Top             =   4560
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
         RecordSource    =   "select * from archiviopagamentiIT"
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
      Begin VB.Label lblpagamento 
         Alignment       =   2  'Center
         Caption         =   "Tipo Pagamento"
         Height          =   255
         Left            =   3840
         TabIndex        =   24
         Top             =   5040
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblsconto1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Discount 1"
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
         Left            =   240
         TabIndex        =   22
         Top             =   5040
         Width           =   975
      End
      Begin VB.Label lblsconto2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Discount 2"
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
         Left            =   1200
         TabIndex        =   21
         Top             =   5040
         Width           =   975
      End
      Begin VB.Label lblsconto3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Discount 3"
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
         Left            =   2160
         TabIndex        =   20
         Top             =   5040
         Width           =   975
      End
      Begin VB.Label lblcap 
         AutoSize        =   -1  'True
         Caption         =   "ZIP"
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
         Left            =   360
         TabIndex        =   19
         Top             =   2400
         Width           =   330
      End
      Begin VB.Label lbllistino 
         AutoSize        =   -1  'True
         Caption         =   "Fax"
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
         Left            =   360
         TabIndex        =   18
         Top             =   4200
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Ville"
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
         Left            =   360
         TabIndex        =   16
         Top             =   1800
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contact"
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
         Left            =   360
         TabIndex        =   15
         Top             =   3000
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Phone"
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
         Left            =   360
         TabIndex        =   14
         Top             =   3600
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Address"
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
         Left            =   360
         TabIndex        =   13
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label Label2 
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
         Left            =   360
         TabIndex        =   12
         Top             =   600
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmFornitori"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mb

Private Sub cmdAdd_Click()
Adodc1.Recordset.AddNew

addupdate

Randomize
TXTCODE = "FOR" & Round(Rnd() * 999999) & TXTCODE + Chr(Round(Rnd() * 25) + 65)
End Sub

Private Sub cmdBack_Click()

Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
frmFornitori.Show
End Sub

Private Sub cmdDelete_Click()
Adodc1.Recordset.delete

delete
End Sub

Private Sub cmdReport_Click()
'Set DataReport4.DataSource = Adodc1
'DataReport4.Show
frmFornitori.Enabled = False

End Sub

Private Sub cmdSave_Click()

If Len(txtName) > 0 And Len(txtAddress) > 0 And Len(txtCittà) > 0 Then
    Adodc1.Recordset.UpdateBatch
    savecancel
    delete
Else
    mb = MsgBox("Compilare campi obbligatori", vbCritical, messaggi(8))
    txtName.SetFocus
End If

End Sub

Private Sub cmdUpdate_Click()
addupdate
End Sub

Private Sub Command1_Click()
Form7.Show vbModal
On Error Resume Next
              Adodc1.Recordset.MoveFirst
  Adodc1.Recordset.Find "codicefornitore = '" & variabile1 & "'", 0, adSearchForward

End Sub

Private Sub Form_Activate()
lingue
delete
End Sub

Private Function addupdate()
cmdAdd.Enabled = False
cmdUpdate.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True
cmddelete.Enabled = False
'cmdReport.Enabled = False

txtName.Locked = False
txtAddress.Locked = False
txtNumber.Locked = False
txtName.SetFocus


Adodc1.Enabled = False
DataGrid1.Enabled = False
End Function

Private Function savecancel()
DataGrid1.Refresh

cmdAdd.Enabled = True
cmdSave.Enabled = False
cmdCancel.Enabled = False

txtName.Locked = True
txtAddress.Locked = True
txtNumber.Locked = True
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



Private Sub txtiva_LostFocus()
'ControllaPIVA (txtiva)
End Sub
Function ControllaPIVA(txtiva)
If txtiva = "" Then
ControllaPIVA = ""
Else
'---------------------------------------------------------------------
If Len(txtiva) <> 11 Then
 mb = MsgBox("lunghezza partita iva errata", vbCritical, "Attenzione")
Else
'-----------------------------------------------------------

   
'------------------------------------------------
Dim s, s1, s2, c, I, char
s1 = 0
    For I = 0 To 9
    I = I + 1
    char = Mid(txtiva, I, 1)
        s1 = s1 + Asc(char) - Asc("0")
'''''''''''''''''''''''''''''''
'controllo dell incremento della variabile
'    response.write(("valore = ")& (asc(char)- asc("0")) & (" s1 = ") & s1 &("<br>") )
''''''''''''''''''''''''''''''''
    Next

    For I = 1 To 9
    I = I + 1
    char = Mid(txtiva, I, 1)
        c = 2 * (Asc(char) - Asc("0"))
            If c > 9 Then
            c = c - 9
            s2 = s2 + c
            Else
            s2 = s2 + c
            End If
'''''''''''''''''''''''''''''''
'controllo dell incremento della variabile
'    response.write(("valore = ")& (asc(char)- asc("0")) & (" c = ") & c & (" s2 = ") & s2 &("<br>") )
''''''''''''''''''''''''''''''''
            Next
            s = s1 + s2
'''''''''''''''''''''''''
'verifica della variabile
'response.Write(s &  ("<br>"))
'''''''''''''''''''''''''
    If ((10 - s Mod 10) Mod 10 <> Asc(Mid(txtiva, 11, 1)) - Asc("0")) Then
        
       mb = MsgBox("Partita Iva non valida", vbCritical, "Attenzione")
    Else
     
    End If

'------------------------------------------------
End If
'------------------------------------------------------------
End If
'---------------------------------------------------------------------

End Function

Public Sub lingue()

Select Case lingua
  
 Case Is = "1F"
   Exit Sub
   
   Case Is = "2I"
 Label2 = "Fornitori"
Label3 = "Indirizzo"
Label6 = "Città"
lblcap = "CAP"
Label1 = "Contatto"
Label4 = "Telefono"
lbllistino = "Fax"
lblsconto1 = "Sconto1"
lblsconto2 = "Sconto2"
lblsconto3 = "Sconto3"
lblpagamento = "Tipo Pagamento"
cmdAdd.Caption = "Nuovo"
cmdSave.Caption = "Salva"
cmdUpdate.Caption = "Aggiorna"
cmddelete.Caption = "Elimina"
cmdback.Caption = "Esci"
cmdCancel.Caption = "Cancella"
frmFornitori.Caption = "Fornitori"


 Case Is = "3G"
 Label2 = "Vendor"
Label3 = "Address"
Label6 = "City"
lblcap = "ZIP"
Label1 = "Contact"
Label4 = "Phone"
lbllistino = "Fax"
lblsconto1 = "Discount 1"
lblsconto2 = "Discount 2"
lblsconto3 = "Discount 3"
lblpagamento = "Payment Type"
cmdAdd.Caption = "New"
cmdSave.Caption = "Save"
cmdUpdate.Caption = "Update"
cmddelete.Caption = "Delete"
cmdback.Caption = "Exit"
cmdCancel.Caption = "Cancel"
frmFornitori.Caption = "Vendors"

Case Is = "4S"
 Label2 = "Proveedor"
Label3 = "Direccion"
Label6 = "Ciudad"
lblcap = "Code"
Label1 = "Contact"
Label4 = "Telefono"
lbllistino = "Fax"
lblsconto1 = "Descuento 1"
lblsconto2 = "Descuento 2"
lblsconto3 = "Descuento 3"
lblpagamento = "Terminos de Pago"
cmdAdd.Caption = "Nuevo"
cmdSave.Caption = "Guarda"
cmdUpdate.Caption = "Actualiza"
cmddelete.Caption = "Borra"
cmdback.Caption = "Salir"
cmdCancel.Caption = "Cancela"
frmFornitori.Caption = "Proveedors"

Case Else
End Select
 
End Sub






