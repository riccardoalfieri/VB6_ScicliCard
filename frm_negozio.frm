VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_negozio 
   Caption         =   "Avviso esito transazione"
   ClientHeight    =   4260
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6510
   LinkTopic       =   "Form2"
   ScaleHeight     =   4260
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H80000010&
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
      Height          =   855
      Left            =   0
      Picture         =   "frm_negozio.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
      Width           =   6495
   End
   Begin MSDataGridLib.DataGrid DataGrid6 
      Bindings        =   "frm_negozio.frx":0442
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   3240
      Visible         =   0   'False
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   1931
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
      ColumnCount     =   8
      BeginProperty Column00 
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
      BeginProperty Column01 
         DataField       =   "pin"
         Caption         =   "pin"
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
         DataField       =   "credito"
         Caption         =   "credito"
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
      BeginProperty Column03 
         DataField       =   "non_disp"
         Caption         =   "non_disp"
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
      BeginProperty Column04 
         DataField       =   "data_att"
         Caption         =   "data_att"
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
         DataField       =   "data_scad"
         Caption         =   "data_scad"
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
         DataField       =   "data_ult"
         Caption         =   "data_ult"
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
         DataField       =   "vuota"
         Caption         =   "vuota"
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
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1214,929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1214,929
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
            ColumnWidth     =   1140,095
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   6510
      _ExtentX        =   11483
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
      Connect         =   $"frm_negozio.frx":0457
      OLEDBString     =   $"frm_negozio.frx":04E4
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from carte"
      Caption         =   "carte"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   3270
      Visible         =   0   'False
      Width           =   6510
      _ExtentX        =   11483
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
      Connect         =   $"frm_negozio.frx":0571
      OLEDBString     =   $"frm_negozio.frx":05FE
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from utenti_datasistema"
      Caption         =   "datasistema"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm_negozio.frx":068B
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   1931
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "data"
         Caption         =   "data"
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
      BeginProperty Column02 
         DataField       =   "perso"
         Caption         =   "perso"
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
         DataField       =   "max_sconto"
         Caption         =   "max_sconto"
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
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frm_negozio.frx":06A0
      Height          =   735
      Left            =   -360
      TabIndex        =   8
      Top             =   4920
      Visible         =   0   'False
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   1296
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "data"
         Caption         =   "data"
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
         DataField       =   "negozio"
         Caption         =   "negozio"
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
      BeginProperty Column03 
         DataField       =   "sconto"
         Caption         =   "sconto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "importo"
         Caption         =   "importo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1214,929
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1214,929
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   3930
      Visible         =   0   'False
      Width           =   6510
      _ExtentX        =   11483
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
      Connect         =   $"frm_negozio.frx":06B5
      OLEDBString     =   $"frm_negozio.frx":0742
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from carte_st"
      Caption         =   "carte storico"
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   $"frm_negozio.frx":07CF
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3975
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   $"frm_negozio.frx":0875
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   3975
      Left            =   -120
      TabIndex        =   5
      Top             =   0
      Width           =   6495
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   855
      Index           =   2
      Left            =   3120
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   855
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   855
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frm_negozio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim calc1, calc2 As Double
Label3.Visible = False

On Error GoTo dopo


 Adodc6.Recordset.MoveFirst  'AGGIORNA CONTO CLIENTE
    Adodc6.Recordset.Find "[codice] = '" & art1 & "'", 0, adSearchForward


If DataGrid6.Columns(0) = art1 Then
  If DataGrid6.Columns(1) = art2 Then   'controlla il pin

       
                           If DataGrid6.Columns(2) >= valore_sconto Then ' controlla se c'è ancora credito sulla card
       
       DataGrid6.Columns(2) = DataGrid6.Columns(2) - valore_sconto
        DataGrid6.Columns(6) = DataGrid1.Columns(0)  ' data ultima transazione
         
         
       Adodc6.Recordset.MoveFirst  ' AGGIORNA CONTO AZIENDA
          Adodc6.Recordset.Find "[codice] = '" & azienda & "'", 0, adSearchForward
          
          calc1 = valore_sconto / DataGrid1.Columns(2) 'calcolo sconto non utilizzabile
          calc2 = valore_sconto - calc1
          Label1(0) = calc2
          Label1(1) = calc1
          Label1(2) = valore_sconto
          DataGrid6.Columns(2) = DataGrid6.Columns(2) + calc2  ' somma lo sconto utilizzabile
           DataGrid6.Columns(3) = DataGrid6.Columns(3) + calc1  ' somma lo sconto NON UTILIZZABILE
         
          DataGrid6.Columns(6) = DataGrid1.Columns(0)  ' data ultima transazione
          
           ' AGGIORNA LO STORICO
          Adodc2.Recordset.AddNew
          DataGrid2.Columns(0) = Date   ' data della transazione
          DataGrid2.Columns(1) = azienda  ' codice azienda
          DataGrid2.Columns(2) = art1    ' codice cliente
          DataGrid2.Columns(3) = valore_sconto
          Adodc2.Recordset.Update
          
          
          
          
          frmpos.txtAmtTendered.Text = ""
          frmpos.txtChange.Text = ""
          frmpos.txtTotal.Text = ""
          
                             Else
                             MsgBox "Sconto residuo: " & DataGrid6.Columns(2) & " non sufficiente", vbInformation, "Transazione NON ESEGUITA"
                              Label3.Visible = True
                            End If
   
          
        
   
   Else
    MsgBox "Codice PIN errato", vbInformation, "Transazione NON ESEGUITA"
     Label3.Visible = True
      End If
      
Else
 MsgBox "Numero Card Errato", vbInformation, "Transazione NON ESEGUITA"
    Label3.Visible = True
   
End If

 
 
Exit Sub


dopo:
 MsgBox "Errore nella procedura", vbInformation, "Transazione NON ESEGUITA"
 Label3.Visible = True
 

End Sub

