VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCategory 
   BackColor       =   &H00808080&
   Caption         =   "Archivio Pagamenti"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8190
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
   ScaleHeight     =   5325
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCat 
      Height          =   495
      Left            =   4080
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   0
      TabIndex        =   6
      Top             =   -120
      Width           =   9255
      Begin VB.CommandButton cmdBack 
         Caption         =   "<<Ritorna"
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
         Left            =   5400
         TabIndex        =   10
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtCategory2 
         DataField       =   "Category"
         DataSource      =   "Adodc2"
         Height          =   495
         Left            =   5760
         TabIndex        =   9
         Top             =   1200
         Visible         =   0   'False
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
         Left            =   3960
         TabIndex        =   1
         Top             =   360
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
         Left            =   3960
         TabIndex        =   2
         Top             =   960
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
         Left            =   3960
         TabIndex        =   3
         Top             =   1560
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
         Left            =   5400
         TabIndex        =   4
         Top             =   240
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
         Left            =   5400
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   5640
         Top             =   2400
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   582
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
         RecordSource    =   "ArchivioPagamenti"
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
         Bindings        =   "frmCategory.frx":0000
         Height          =   1455
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   9015
         _ExtentX        =   15901
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
         ColumnCount     =   9
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
            DataField       =   "tipoPagamento"
            Caption         =   "tipoPagamento"
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
            DataField       =   "scadenza30"
            Caption         =   "scadenza30"
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
            DataField       =   "scadenza60"
            Caption         =   "scadenza60"
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
            DataField       =   "scadenza90"
            Caption         =   "scadenza90"
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
            DataField       =   "scadenza120"
            Caption         =   "scadenza120"
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
            DataField       =   "scadenza180"
            Caption         =   "scadenza180"
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
            DataField       =   "scadenza210"
            Caption         =   "scadenza210"
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
               ColumnWidth     =   1275,024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2429,858
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2429,858
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1244,976
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1244,976
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1244,976
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1365,165
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1365,165
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1365,165
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   5520
         Top             =   600
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
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
         RecordSource    =   "ArchivioPagamenti"
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
      Begin VB.TextBox txtCategory 
         DataField       =   "Category"
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
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   0
         Top             =   1680
         Width           =   3495
      End
      Begin VB.Image Image1 
         Height          =   915
         Left            =   240
         Picture         =   "frmCategory.frx":0015
         Top             =   360
         Width           =   2460
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Pagamenti"
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
         TabIndex        =   7
         Top             =   1440
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmCategory"
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
frmCategory.Show
End Sub

Private Sub cmdDelete_Click()
Adodc1.Recordset.delete

delete
End Sub

Private Sub cmdSave_Click()
If txtCat = txtCategory Then
    Adodc1.Recordset.UpdateBatch
    savecancel
    delete
    txtCat = ""
    GoTo skip
End If

Adodc2.Refresh
Adodc2.Recordset.Find "Category=" & "'" & txtCategory & "'"

If Len(txtCategory) > 0 Then
    If Len(txtCategory2) = 0 Then
        Adodc1.Recordset.UpdateBatch
        savecancel
        delete
    Else
        mb = MsgBox("Category already in used.", vbCritical, "Yokohama")
        txtCategory.SetFocus
       
    End If
Else
    mb = MsgBox("Invalid data.", vbCritical, "Yokohama")
    txtCategory.SetFocus
End If

skip:
End Sub

Private Sub cmdUpdate_Click()
addupdate

txtCat = txtCategory
End Sub

Private Sub Form_Activate()

delete
End Sub

Private Function addupdate()
cmdAdd.Enabled = False
cmdUpdate.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True
cmdDelete.Enabled = False

txtCategory.Locked = False
txtCategory.SetFocus


DataGrid1.Enabled = False
End Function

Private Function savecancel()
DataGrid1.Refresh

cmdAdd.Enabled = True
cmdUpdate.Enabled = True
cmdSave.Enabled = False
cmdCancel.Enabled = False

txtCategory.Locked = True
txtCategory.SetFocus
End Function

Private Function delete()
DataGrid1.Refresh

If Adodc1.Recordset.RecordCount = 0 Then
    DataGrid1.Enabled = False
    cmdDelete.Enabled = False
    cmdUpdate.Enabled = False
Else
    DataGrid1.Enabled = True
    cmdDelete.Enabled = True
    cmdUpdate.Enabled = True
End If

txtCategory.SetFocus
End Function

Private Sub Form_Unload(Cancel As Integer)
Adodc1.Recordset.CancelBatch
End Sub
