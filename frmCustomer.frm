VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCustomer 
   BackColor       =   &H00808080&
   Caption         =   "Clients"
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
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10575
      Begin VB.TextBox Text1 
         DataField       =   "cellulare"
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
         Left            =   4200
         MaxLength       =   50
         TabIndex        =   4
         Top             =   4440
         Width           =   2055
      End
      Begin VB.TextBox txtlistino 
         BackColor       =   &H8000000E&
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
         Left            =   4320
         MaxLength       =   50
         TabIndex        =   27
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Sur le net"
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
         Left            =   5640
         Picture         =   "frmCustomer.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   7320
         Width           =   1455
      End
      Begin VB.TextBox txtfiscale 
         DataField       =   "CodiceFiscale"
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
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   7
         Top             =   5160
         Width           =   2055
      End
      Begin VB.TextBox txtiva 
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
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   50
         TabIndex        =   6
         Top             =   5160
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Rechercher"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5640
         Picture         =   "frmCustomer.frx":1D42
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   6480
         Width           =   1455
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Nouveau"
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
         Left            =   240
         Picture         =   "frmCustomer.frx":3A54
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   6480
         Width           =   1575
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
         Height          =   735
         Left            =   240
         Picture         =   "frmCustomer.frx":5E46
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   7320
         Width           =   1575
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
         Height          =   735
         Left            =   1920
         Picture         =   "frmCustomer.frx":5F90
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6480
         Width           =   1695
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
         Height          =   735
         Left            =   3840
         Picture         =   "frmCustomer.frx":60DA
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   6480
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Supprimér"
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
         Left            =   1920
         Picture         =   "frmCustomer.frx":651C
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   7320
         Width           =   1695
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
         Height          =   735
         Left            =   3840
         Picture         =   "frmCustomer.frx":6666
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   7320
         Width           =   1575
      End
      Begin VB.TextBox TXTCODE 
         DataField       =   "Customer Code"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   7440
         TabIndex        =   15
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtCittà 
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
         Left            =   120
         MaxLength       =   100
         TabIndex        =   2
         Top             =   3600
         Width           =   3495
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
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         Top             =   4320
         Width           =   2055
      End
      Begin VB.TextBox txtAddress 
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
         Left            =   120
         MaxLength       =   100
         TabIndex        =   1
         Top             =   2760
         Width           =   6135
      End
      Begin VB.TextBox txtName 
         DataField       =   "Cliente"
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
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   1920
         Width           =   6255
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
         RecordSource    =   "SELECT * FROM CUSTOMERS ORDER BY Cliente"
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
         Bindings        =   "frmCustomer.frx":6AA8
         Height          =   1455
         Left            =   360
         TabIndex        =   10
         Top             =   8400
         Visible         =   0   'False
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   2566
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
         Bindings        =   "frmCustomer.frx":6ABD
         DataField       =   "Pagamento"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   6000
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "tipoPagamento"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   495
         Left            =   3360
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
         RecordSource    =   "ArchivioPagamentiIT"
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
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Code"
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
         Index           =   9
         Left            =   6840
         TabIndex        =   31
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Portable"
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
         Index           =   5
         Left            =   4200
         TabIndex        =   30
         Top             =   4200
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Prix de Vente"
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
         Index           =   3
         Left            =   4320
         TabIndex        =   29
         Top             =   3360
         Width           =   1320
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "Conditions de Règlements"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   28
         Top             =   5760
         Width           =   2535
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Codice Fiscale"
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
         Index           =   8
         Left            =   2400
         TabIndex        =   25
         Top             =   4920
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Partita Iva"
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
         Index           =   7
         Left            =   120
         TabIndex        =   24
         Top             =   4920
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   120
         Picture         =   "frmCustomer.frx":6AD2
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliquez Nouveau, Insérer le Client et Cliquez Valider"
         Height          =   375
         Left            =   1080
         TabIndex        =   22
         Top             =   840
         Width           =   4335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Clients"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   1080
         TabIndex        =   21
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
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
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   3360
         Width           =   390
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Telephone"
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
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   4080
         Width           =   1020
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Adresse"
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
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   780
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Client"
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
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   555
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   1335
         Left            =   0
         Top             =   120
         Width           =   9615
      End
   End
End
Attribute VB_Name = "frmCustomer"
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
TXTCODE = "CUS" & Round(Rnd() * 999999) & TXTCODE + Chr(Round(Rnd() * 25) + 65)
End Sub

Private Sub cmdBack_Click()

Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
frmCustomer.Show
End Sub

Private Sub cmdDelete_Click()
Dim response As String
On Error Resume Next
response = MsgBox("Cancello questo cliente?", vbOKCancel + vbCancel, "Attenzione")
Select Case response
 Case 6
 Adodc1.Recordset.delete


 Case Else
End Select





End Sub

Private Sub cmdReport_Click()
'Set DataReport4.DataSource = Adodc1
'DataReport4.Show
frmCustomer.Enabled = False

End Sub

Private Sub cmdSave_Click()

On Error GoTo UpdateErr

If Len(txtName) > 0 Then
    Adodc1.Recordset.UpdateBatch
    savecancel
    delete
Else
    mb = MsgBox("Error ", vbCritical, "Attention")
    txtName.SetFocus
End If

 Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
addupdate
End Sub

Private Sub Command1_Click()
Form4.Show vbModal
On Error Resume Next
              Adodc1.Recordset.MoveFirst
  Adodc1.Recordset.Find "[Customer Code] = '" & variabile1 & "'", 0, adSearchForward

End Sub

Private Sub Command3_Click()

campo(0) = TXTCODE
campo(1) = txtName
campo(2) = txtAddress & " " & txtCittà
 
  art11 = txtNumber
  
If azienda <> "" Then
frm_client_net.Show vbModal

Else
MsgBox messaggi(1), vbOKOnly + vbExclamation, messaggi(2)

End If

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
cmdcancel.Enabled = False

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

Private Sub txtfiscale_LostFocus()
controllaCF (txtfiscale)
End Sub



Private Sub txtiva_LostFocus()
ControllaPIVA (txtiva)
End Sub

Function controllaCF(txtfiscale)
    If txtfiscale = "" Then
    
Else
'------------------------------------------
    If Len(txtfiscale) <> 16 Then
      mb = MsgBox("Codice fiscale errato", vbCritical, "Attenzione")
    Else
'--------------------------------------------------
 txtfiscale = UCase(txtfiscale)
'-----------------------------------------------------------

    Dim s, c, s1, s2, I
    s1 = 0
    I = 0
''''''''''''''''''''''''''''''''''''''
'prima versione - meno elegante
'    for i = 1 to 14
'    i = i + 1
''''''''''''''''''''''''''''''''''''''
For I = 2 To 14 Step 2
        c = Mid(txtfiscale, I, 1)
        If ("0" <= c And c <= "9") Then
            s1 = s1 + Asc(c) - Asc("0")
''''''''''''''''''''''''
'controlla il loop
'            response.write("c= "& c & " s1= "& s1 &"<br>")
''''''''''''''''''''''''
        Else
            s1 = s1 + Asc(c) - Asc("A")
''''''''''''''''''''''''
'controlla il loop
'            response.write("c= "& c & " s1= "& s1 &"<br>")
''''''''''''''''''''''''
        End If
    Next
'''''''''''''''''''''
'controlla la somma delle cifre pari
'    response.write("s1="&s1&"<br>")
''''''''''''''''''''''
    s2 = 0
'''''''''''''''''''''''''''''''
'prima versione - meno elegante
'    for i = 0 to 14
'    i = i + 1
'''''''''''''''''''''''''''''
For I = 1 To 15 Step 2
        c = Mid(txtfiscale, I, 1)
        Select Case (c)
        Case "0"
          s2 = s2 + 1
        Case "1"
          s2 = s2 + 0
        Case "2"
          s2 = s2 + 5
        Case "3"
          s2 = s2 + 7
        Case "4"
          s2 = s2 + 9
        Case "5"
          s2 = s2 + 13
        Case "6"
          s2 = s2 + 15
        Case "7"
          s2 = s2 + 17
        Case "8"
          s2 = s2 + 19
        Case "9"
          s2 = s2 + 21
        Case "A"
          s2 = s2 + 1
        Case "B"
          s2 = s2 + 0
        Case "C"
          s2 = s2 + 5
        Case "D"
          s2 = s2 + 7
        Case "E"
          s2 = s2 + 9
        Case "F"
          s2 = s2 + 13
        Case "G"
          s2 = s2 + 15
        Case "H"
          s2 = s2 + 17
        Case "I"
          s2 = s2 + 19
        Case "J"
          s2 = s2 + 21
        Case "K"
          s2 = s2 + 2
        Case "L"
          s2 = s2 + 4
        Case "M"
          s2 = s2 + 18
        Case "N"
          s2 = s2 + 20
        Case "O"
          s2 = s2 + 11
        Case "P"
          s2 = s2 + 3
        Case "Q"
          s2 = s2 + 6
        Case "R"
          s2 = s2 + 8
        Case "S"
          s2 = s2 + 12
        Case "T"
          s2 = s2 + 14
        Case "U"
          s2 = s2 + 16
        Case "V"
          s2 = s2 + 10
        Case "W"
          s2 = s2 + 22
        Case "X"
          s2 = s2 + 25
        Case "Y"
          s2 = s2 + 24
        Case "Z"
          s2 = s2 + 23
        End Select
''''''''''''''''''''''''
'controlla il loop
'        response.write("c= "& c & " s2= "& s2 &"<br>")
''''''''''''''''''''''''
        Next
        s = s1 + s2
''''''''''''''''''''''''
'controlla la somma dispari
'            response.write("s2="&s2&"<br>")
'controlla il totale
'        response.write(s&"<br>")
''''''''''''''''''''''''
        If Chr((s Mod 26) + Asc("A")) <> Mid(txtfiscale, 16, 1) Then
       mb = MsgBox("Codice fiscale errato", vbCritical, "Attenzione")
      
        Else
        mb = ""
        End If
'-----------------------------------------------------------
        End If
'--------------------------------------------------
 End If
'------------------------------------------
'End If
Print "s1"; s1; "s2"; s2; "s"; s
End Function
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

Select Case LINGUA
  
 Case Is = "1F"
messaggi(1) = "Connexion au server n'établi pas "
messaggi(2) = "Attention"

   
   Case Is = "2I"
 frmCustomer.Caption = "Clienti"
 Label5 = "Clienti"
 Label7 = "Cliccare su Nuovo, inserire i dati e Premere Salva"
 
lbl(0) = "Nome"

lbl(1) = "Indirizzo"
lbl(2) = "Città"
lbl(3) = "Prezzo di Vendita"
lbl(4) = "Telefono"
lbl(5) = "Cellulare"
lbl(6) = "Pagamenti"
lbl(7) = "Partita Iva"
lbl(8) = "Codice Fiscale"
'Command2.Caption = "Storico"
cmdAdd.Caption = "Nuovo"
cmdSave.Caption = "Salva"
cmdUpdate.Caption = "Aggiorna"
cmddelete.Caption = "Elimina"
cmdBack.Caption = "Esci"
cmdcancel.Caption = "Cancella"
Command1.Caption = "Ricerca"

messaggi(1) = "Connexion au server n'établi pas "
messaggi(2) = "Attention"


 Case Is = "3G"


 Case Is = "4S"


Case Else
End Select
 
End Sub





