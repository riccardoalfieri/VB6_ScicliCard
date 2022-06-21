VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmArchivioReparti 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Categorie"
   ClientHeight    =   4575
   ClientLeft      =   1095
   ClientTop       =   435
   ClientWidth     =   5775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5775
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1260
      Left            =   0
      ScaleHeight     =   1260
      ScaleWidth      =   5775
      TabIndex        =   2
      Top             =   3315
      Width           =   5775
      Begin VB.CommandButton cmdClose 
         Caption         =   "Quitter"
         Height          =   900
         Left            =   4320
         Picture         =   "frmArchivioReparti.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Ajourner"
         Height          =   900
         Left            =   3240
         Picture         =   "frmArchivioReparti.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Supprimer"
         Height          =   900
         Left            =   2160
         Picture         =   "frmArchivioReparti.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Valider"
         Height          =   900
         Left            =   1080
         Picture         =   "frmArchivioReparti.frx":06D6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Nouvelle"
         Height          =   900
         Left            =   0
         Picture         =   "frmArchivioReparti.frx":0820
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc datPrimaryRS 
         Height          =   330
         Left            =   -120
         Top             =   960
         Width           =   6015
         _ExtentX        =   10610
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
         RecordSource    =   "select * from reparti"
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
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Categoria"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Categorie"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   380
      Width           =   1815
   End
End
Attribute VB_Name = "frmArchivioReparti"
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

