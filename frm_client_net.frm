VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_client_net 
   BackColor       =   &H00FF0000&
   Caption         =   "Client sur le net"
   ClientHeight    =   9345
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   12060
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frm_client_net.frx":0000
      Height          =   7695
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   13573
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
            Object.Visible         =   0   'False
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00FFFF80&
      Caption         =   "Sauver sur le net"
      Height          =   900
      Left            =   2880
      Picture         =   "frm_client_net.frx":0015
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7680
      Width           =   1935
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Quitter"
      Height          =   900
      Left            =   8280
      Picture         =   "frm_client_net.frx":015F
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FF80FF&
      Caption         =   "Supprimer"
      Height          =   900
      Left            =   4800
      Picture         =   "frm_client_net.frx":05A1
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Liste Complete"
      Height          =   855
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7680
      Width           =   2415
   End
   Begin VB.TextBox Textclient 
      DataField       =   "Customer Code"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   4080
      Width           =   4935
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
      Left            =   360
      MaxLength       =   100
      TabIndex        =   5
      Top             =   4920
      Width           =   6135
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
      TabIndex        =   4
      Top             =   6480
      Width           =   2055
   End
   Begin VB.TextBox txtmot 
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
      TabIndex        =   3
      Text            =   "password"
      Top             =   5760
      Width           =   3495
   End
   Begin VB.TextBox TXTCODE 
      DataField       =   "Customer Code"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   4080
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      DataField       =   "cliente"
      DataSource      =   "Adodc2"
      Height          =   495
      Left            =   9000
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm_client_net.frx":06EB
      Height          =   735
      Left            =   -120
      TabIndex        =   0
      Top             =   120
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   1296
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
            Object.Visible         =   0   'False
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739,906
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
      Connect         =   $"frm_client_net.frx":0700
      OLEDBString     =   $"frm_client_net.frx":078D
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from accesso"
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
      Index           =   3
      Left            =   3240
      TabIndex        =   11
      Top             =   3840
      Width           =   555
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      Caption         =   "User Id"
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
      Left            =   360
      TabIndex        =   9
      Top             =   3840
      Width           =   705
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
      Left            =   360
      TabIndex        =   8
      Top             =   4680
      Width           =   780
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
      Left            =   360
      TabIndex        =   7
      Top             =   6240
      Width           =   1020
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      Caption         =   "Mot de passe"
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
      Left            =   360
      TabIndex        =   6
      Top             =   5520
      Width           =   1335
   End
End
Attribute VB_Name = "frm_client_net"
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

response = MsgBox("Supprimer ce client?", vbOKCancel + vbCancel, "Attention")
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

DataGrid1.Columns(1) = azienda
DataGrid1.Columns(2) = codiceazienda
DataGrid1.Columns(3) = TXTCODE
DataGrid1.Columns(4) = txtmot
DataGrid1.Columns(5) = Textclient
DataGrid1.Columns(6) = txtAddress
DataGrid1.Columns(7) = txtNumber
DataGrid1.Columns(8) = "eat"
DataGrid1.Columns(9) = cittadina
Adodc2.Recordset.Update


End Sub

Private Sub Command1_Click()
Dim stringa As String
DataGrid2.Visible = True
cmdUpdate.Visible = False


stringa = "select * from accesso where azienda='" & azienda & "'"

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


stringa = "select * from accesso where azienda='" & azienda & "' and cliente='" & campo(0) & "'"

With Adodc2
 .RecordSource = stringa
 .Refresh
 End With
 
 With DataGrid1
  .ReBind
  End With
  

  
 TXTCODE = campo(0)
 Textclient = campo(1)
 txtAddress = campo(2)
  txtNumber = art11
 

sotto:

End Sub
