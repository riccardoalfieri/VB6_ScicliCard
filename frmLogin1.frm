VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmLogin1 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autenticazione al server ScicliCard"
   ClientHeight    =   4770
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin1.frx":0000
   ScaleHeight     =   2818.272
   ScaleMode       =   0  'User
   ScaleWidth      =   5140.729
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "Save Password"
      DataField       =   "server"
      DataSource      =   "datPrimaryRS"
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmLogin1.frx":18BF3B
      Height          =   855
      Left            =   4680
      TabIndex        =   12
      Top             =   1560
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1508
      _Version        =   393216
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
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "id"
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
      BeginProperty Column02 
         DataField       =   "codice_azienda"
         Caption         =   "codice_azienda"
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
         DataField       =   "cliente"
         Caption         =   "cliente"
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
         DataField       =   "codice_cliente"
         Caption         =   "codice_cliente"
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
         DataField       =   "intestazione"
         Caption         =   "intestazione"
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
         DataField       =   "indirizzo"
         Caption         =   "indirizzo"
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
         DataField       =   "telefono"
         Caption         =   "telefono"
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
      BeginProperty Column09 
         DataField       =   "libero"
         Caption         =   "libero"
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
            ColumnWidth     =   859,158
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1633,677
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1126,913
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1633,677
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1070,487
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1633,677
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1633,677
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1633,677
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1070,487
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1633,677
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmLogin1.frx":18BF50
      Height          =   315
      Left            =   360
      TabIndex        =   9
      ToolTipText     =   "Choisir le salon de coiffure"
      Top             =   840
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "azienda"
      Text            =   ""
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H0080FF80&
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
      Height          =   615
      Left            =   2880
      Picture         =   "frmLogin1.frx":18BF65
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FF8080&
      Caption         =   "Accedi"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      Picture         =   "frmLogin1.frx":18DCA7
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox txtcodice 
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      ToolTipText     =   "Insérer Votre pseudo. Si vous n'avez pseudo, demandez-le au salon."
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Versione Demo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Insérer Votre mot de passe. Si vous n'avez mot de passe, demandez-le au salon."
      Top             =   1920
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   5475
      _ExtentX        =   9657
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
      Connect         =   $"frmLogin1.frx":190099
      OLEDBString     =   $"frmLogin1.frx":190126
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select distinct [azienda]  from accesso where indirizzo='eat'"
      Caption         =   " adodc1"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   4110
      Visible         =   0   'False
      Width           =   5475
      _ExtentX        =   9657
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
      Connect         =   $"frmLogin1.frx":1901B3
      OLEDBString     =   $"frmLogin1.frx":190240
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *  from accesso"
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
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   3780
      Visible         =   0   'False
      Width           =   5475
      _ExtentX        =   9657
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
      RecordSource    =   "select * from azienda"
      Caption         =   " dati locali"
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
   Begin VB.Label Label4 
      Caption         =   "Label4"
      DataField       =   "libero"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   1800
      TabIndex        =   14
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      DataField       =   "codice_azienda"
      DataSource      =   "Adodc2"
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Restaurant Pizzeria"
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   3960
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Inserire User Id e Password per l'accesso ai Vostri dati online"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   3600
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   15
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   135
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Id Esercizio"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1800
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Password"
      Height          =   390
      Index           =   1
      Left            =   225
      TabIndex        =   3
      Top             =   1920
      Width           =   1800
   End
End
Attribute VB_Name = "frmLogin1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()

Unload Me
    
End Sub

Private Sub cmdOK_Click()
    'Verifica la validità della password
   On Error Resume Next
              Adodc2.Recordset.MoveFirst
  Adodc2.Recordset.Find "[azienda] = '" & txtcodice & "'", 0, adSearchForward
       
        If Left(Label3, Len(txtPassword)) = txtPassword And Label3 <> "" Then
        
     
      
 
        txtPassword.SetFocus
      ' id_cliente = txtcodice
      ' pw_cliente = txtPassword
       azienda = txtcodice
        codiceazienda = Label3
         cittadina = Label4
         
Select Case Check2.Value
  Case Is = 1
    datPrimaryRS.Recordset.Update
     Case Else
     End Select
     
       Me.Hide
    
    Else


        MsgBox messaggi(3), , messaggi(1)
        
        txtPassword.SetFocus
        
       End If
  
    
End Sub

Private Sub Command1_Click()
txtPassword = "XX"
cmdOK_Click
End Sub

Private Sub DataCombo1_lostfocus()
Dim stringa As String
stringa = "select distinct [azienda]  from accesso  where azienda='" & DataCombo1 & "' order by azienda "

With Adodc1
 .RecordSource = stringa
    .Refresh
 End With
 With DataCombo1
  .ReFill
  End With
  
 azienda = DataCombo1 'salone
End Sub

Private Sub Form_Load()
On Error GoTo sotto
Dim a As String
Dim COSTANTE As Integer
 Select Case Check2.Value
   Case Is = 0
   txtPassword = ""
   Case Else
   End Select

LINGUA
Exit Sub

sotto:
'Form2.Show vbModal
  
End Sub

Public Sub LINGUA()
messaggi(1) = "Attenzione"
messaggi(3) = "User Id o Password invalidi"
End Sub

Private Sub Form_Unload(Cancel As Integer)

End

    
End Sub


Private Sub txtcodice_KeyDown( _
           KeyCode As Integer, Shift As Integer)

     Select Case KeyCode
     Case vbKeyReturn:
     txtPassword.SetFocus
     End Select
     

End Sub
