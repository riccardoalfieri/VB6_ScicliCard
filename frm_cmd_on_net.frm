VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_cmd_on_net 
   Caption         =   "Commandes sur le net"
   ClientHeight    =   9345
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   12060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Menu complet"
      Height          =   855
      Left            =   4200
      TabIndex        =   4
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Supprimer"
      Height          =   900
      Left            =   2640
      Picture         =   "frm_cmd_on_net.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Quitter"
      Height          =   900
      Left            =   1560
      Picture         =   "frm_cmd_on_net.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Sauver sur le net"
      Height          =   900
      Left            =   480
      Picture         =   "frm_cmd_on_net.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7920
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm_cmd_on_net.frx":06D6
      Height          =   2655
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   4683
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
      ColumnCount     =   22
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
         DataField       =   "Client"
         Caption         =   "Client"
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
         DataField       =   "Clientcode"
         Caption         =   "Clientcode"
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
         DataField       =   "date"
         Caption         =   "date"
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
         DataField       =   "employee"
         Caption         =   "employee"
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
         DataField       =   "comment"
         Caption         =   "comment"
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
         DataField       =   "service"
         Caption         =   "service"
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
         DataField       =   "comment1"
         Caption         =   "comment1"
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
         DataField       =   "Formula"
         Caption         =   "Formula"
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
         DataField       =   "price"
         Caption         =   "price"
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
      BeginProperty Column11 
         DataField       =   "flag"
         Caption         =   "flag"
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
         DataField       =   "flag1"
         Caption         =   "flag1"
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
         DataField       =   "flag2"
         Caption         =   "flag2"
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
         DataField       =   "importopagato"
         Caption         =   "importopagato"
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
      BeginProperty Column15 
         DataField       =   "tipopagamento"
         Caption         =   "tipopagamento"
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
      BeginProperty Column16 
         DataField       =   "sconto"
         Caption         =   "sconto"
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
      BeginProperty Column17 
         DataField       =   "importototale"
         Caption         =   "importototale"
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
      BeginProperty Column18 
         DataField       =   "numric"
         Caption         =   "numric"
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
      BeginProperty Column19 
         DataField       =   "pezzi"
         Caption         =   "pezzi"
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
      BeginProperty Column20 
         DataField       =   "codart"
         Caption         =   "codart"
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
      BeginProperty Column21 
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
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739,906
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
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column21 
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=class_1;Data Source=PC-MEMMO\SQLEXPRESS"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=class_1;Data Source=PC-MEMMO\SQLEXPRESS"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from commandes"
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
End
Attribute VB_Name = "frm_cmd_on_net"
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



Private Sub Form_Load()
Text13 = variabile1

txtclient = variabile2

  stringa = "select * from commande where clientcode='" & variabile1 & "'"
    
    
   With Adodc2
    .RecordSource = stringa
    .Refresh
    End With
    
    With DataGrid1
     .ClearFields
     .HoldFields
     .ReBind
     End With

 'calcola
End Sub
