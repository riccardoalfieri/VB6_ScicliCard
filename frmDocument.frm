VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDocument 
   Caption         =   "Desktop"
   ClientHeight    =   9330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14010
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmDocument.frx":5A6C0
   ScaleHeight     =   12500
   ScaleMode       =   0  'User
   ScaleWidth      =   18390.21
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      DataField       =   "client"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   14010
      _ExtentX        =   24712
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
   Begin VB.Label lbllingua 
      Caption         =   "SCONTO"
      DataField       =   "Azienda"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   4080
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub rtfText_SelChange()
  
End Sub


Private Sub Form_Load()
SCONTO = lbllingua
End Sub
