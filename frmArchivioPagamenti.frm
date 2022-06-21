VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmArchivioPagamenti 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ArchivioPagamenti"
   ClientHeight    =   3840
   ClientLeft      =   1095
   ClientTop       =   435
   ClientWidth     =   5775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5775
   Begin VB.CheckBox chkFields 
      DataField       =   "scadenza90"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   25
      Top             =   1965
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "scadenza60"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   24
      Top             =   1635
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "scadenza30"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   23
      Top             =   1320
      Width           =   3375
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   5775
      TabIndex        =   17
      Top             =   3210
      Width           =   5775
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Chiudi"
         Height          =   300
         Left            =   4675
         TabIndex        =   22
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Rivisualizza"
         Height          =   300
         Left            =   3521
         TabIndex        =   21
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Elimina"
         Height          =   300
         Left            =   2367
         TabIndex        =   20
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Aggiorna"
         Height          =   300
         Left            =   1213
         TabIndex        =   19
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Aggi&ungi"
         Height          =   300
         Left            =   59
         TabIndex        =   18
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "scadenza180"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   9
      Left            =   2040
      TabIndex        =   16
      Top             =   2940
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "scadenza150"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   8
      Left            =   2040
      TabIndex        =   14
      Top             =   2620
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "scadenza120"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   12
      Top             =   2300
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "numerorate"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1020
      Width           =   735
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "finemese"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   700
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "tipoPagamento"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   380
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "codice"
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
      Top             =   3510
      Width           =   5775
      _ExtentX        =   10186
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
      RecordSource    =   $"frmArchivioPagamenti.frx":0000
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
   Begin VB.Label lblLabels 
      Caption         =   "scadenza180:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   15
      Top             =   2940
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "scadenza150:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   13
      Top             =   2620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "scadenza120:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   11
      Top             =   2300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "scadenza90:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   10
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "scadenza60:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   1660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "scadenza30:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "numerorate:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "finemese:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "tipoPagamento:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "codice:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmArchivioPagamenti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

