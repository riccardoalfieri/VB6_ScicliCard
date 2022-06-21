VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmcommande 
   BackColor       =   &H00FFC0C0&
   Caption         =   "La Caisse"
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000013&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9405
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Cmddel 
      BackColor       =   &H000000FF&
      Caption         =   "Supprimer ticket"
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
      Left            =   12000
      Picture         =   "frmcommande.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton cmdstampa 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Imprimer Ticket"
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
      Left            =   5760
      Picture         =   "frmcommande.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   8040
      Width           =   1695
   End
   Begin VB.TextBox Text12 
      DataField       =   "riga3"
      DataSource      =   "Adoticket"
      Height          =   360
      Left            =   3360
      TabIndex        =   96
      Text            =   "riga3"
      Top             =   8040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text11 
      DataField       =   "riga2"
      DataSource      =   "Adoticket"
      Height          =   360
      Left            =   2640
      TabIndex        =   95
      Text            =   "riga2"
      Top             =   8040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text10 
      DataField       =   "riga1"
      DataSource      =   "Adoticket"
      Height          =   360
      Left            =   1920
      TabIndex        =   94
      Text            =   "riga1"
      Top             =   8040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txttestata 
      DataField       =   "Azienda"
      DataSource      =   "Adoticket"
      Height          =   360
      Left            =   1200
      TabIndex        =   93
      Text            =   "testata"
      Top             =   8040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame frame11 
      Caption         =   "Frame1"
      Height          =   6375
      Left            =   4080
      TabIndex        =   41
      Top             =   840
      Width           =   7935
      Begin VB.TextBox txtAmtTendered 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3081
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   1920
         TabIndex        =   72
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtTotal 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3081
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   240
         TabIndex        =   71
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtPmt 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   3480
         TabIndex        =   70
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtChange 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3081
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   1920
         TabIndex        =   69
         Top             =   2160
         Width           =   3495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000013&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   0
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   5280
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000013&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000013&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   2
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000013&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   3
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000013&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   4
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000013&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   5
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000013&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   6
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000013&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   7
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000013&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   8
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000013&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   9
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   5280
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000013&
         Caption         =   ","
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   11
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   5280
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FF00&
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5520
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   5280
         Width           =   2295
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Espèce"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   4320
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Carte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   1200
         Width           =   2295
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cheque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   10
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   5280
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "€5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   5
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "€50"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   50
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "€20"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   20
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "€10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   10
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Cheque Details"
         Height          =   1215
         Left            =   120
         TabIndex        =   42
         Top             =   -960
         Visible         =   0   'False
         Width           =   4695
         Begin VB.TextBox Text2 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   1440
            TabIndex        =   45
            Top             =   720
            Width           =   3495
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   1440
            TabIndex        =   44
            Top             =   1200
            Width           =   3495
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   1440
            TabIndex        =   43
            Top             =   240
            Width           =   3495
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Account No"
            Height          =   375
            Left            =   120
            TabIndex        =   46
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   375
            Left            =   240
            TabIndex        =   48
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cheque No"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   47
            Top             =   1320
            Width           =   1215
         End
      End
      Begin MSAdodcLib.Adodc Adodc9 
         Height          =   330
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         RecordSource    =   "cassa"
         Caption         =   "Adodc9"
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
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmcommande.frx":1E5C
         Height          =   255
         Left            =   0
         TabIndex        =   73
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   18
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
            DataField       =   "libero1"
            Caption         =   "libero1"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1454,74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2775,118
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2775,118
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2775,118
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2775,118
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2775,118
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   2775,118
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   2775,118
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   2775,118
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   2775,118
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   2775,118
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1214,929
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1214,929
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1214,929
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   2775,118
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   2775,118
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   2775,118
            EndProperty
            BeginProperty Column17 
               ColumnWidth     =   2775,118
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Paiement"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   74
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "A rendre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   240
         TabIndex        =   75
         Top             =   2280
         Width           =   1575
      End
   End
   Begin VB.TextBox txtcatego 
      DataField       =   "Submenudesc"
      DataSource      =   "Adocatego"
      Height          =   360
      Left            =   6120
      TabIndex        =   92
      Text            =   "Txtcatego"
      Top             =   -120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdclavier 
      BackColor       =   &H00FFFF00&
      Caption         =   "Hotel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   8
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   9120
      Width           =   1695
   End
   Begin VB.CommandButton cmdclavier 
      BackColor       =   &H00FFFF80&
      Caption         =   "Market"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   7
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   9120
      Width           =   1695
   End
   Begin VB.CommandButton cmdclavier 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Magazin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   6
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   9120
      Width           =   1695
   End
   Begin VB.CommandButton cmdclavier 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Boutique"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   5
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   9120
      Width           =   1695
   End
   Begin VB.CommandButton cmdclavier 
      BackColor       =   &H00FF0000&
      Caption         =   "Tables"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   4
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   9120
      Width           =   1695
   End
   Begin VB.CommandButton cmdclavier 
      BackColor       =   &H00FF8080&
      Caption         =   "Salon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   3
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   9120
      Width           =   1695
   End
   Begin VB.CommandButton cmdclavier 
      BackColor       =   &H00C000C0&
      Caption         =   "Pizza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   2
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   9120
      Width           =   1695
   End
   Begin VB.CommandButton cmdclavier 
      BackColor       =   &H00FF00FF&
      Caption         =   "Restaurant"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   1
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   9120
      Width           =   1695
   End
   Begin VB.CommandButton cmdclavier 
      BackColor       =   &H00FF80FF&
      Caption         =   "Bar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   0
      Left            =   -120
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   9120
      Width           =   1695
   End
   Begin VB.CheckBox chkvisible 
      Caption         =   "Visible"
      DataField       =   "Visibile"
      DataSource      =   "Adodc6"
      Height          =   255
      Left            =   3720
      TabIndex        =   82
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtchiave 
      DataField       =   "chiave"
      DataSource      =   "Adodc6"
      Height          =   360
      Left            =   5040
      TabIndex        =   81
      Text            =   "txtchiave"
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtpic 
      DataField       =   "Immagine"
      DataSource      =   "Adodc6"
      Height          =   360
      Left            =   4320
      TabIndex        =   80
      Text            =   "txtpic"
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text9 
      DataField       =   "listino1"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   0
      EndProperty
      DataSource      =   "Adodc6"
      Height          =   360
      Left            =   3600
      TabIndex        =   79
      Text            =   "Text9"
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text8 
      DataField       =   "BColor"
      DataSource      =   "Adodc6"
      Height          =   360
      Left            =   5400
      TabIndex        =   78
      Text            =   "Text8"
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      DataField       =   "Itemdesc"
      DataSource      =   "Adodc6"
      Height          =   360
      Left            =   3600
      TabIndex        =   77
      Text            =   "Text7"
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0C0FF&
      Caption         =   """a"" & chr$(13) & ""b"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   -120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      DataField       =   "menu"
      DataSource      =   "Adoazienda"
      Height          =   360
      Left            =   4080
      TabIndex        =   40
      Text            =   "Text4"
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text5 
      DataField       =   "righe"
      DataSource      =   "Adoazienda"
      Height          =   375
      Left            =   4920
      TabIndex        =   39
      Text            =   "Text5"
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text4 
      DataField       =   "righe"
      DataSource      =   "Adoazienda"
      Height          =   360
      Left            =   4080
      TabIndex        =   38
      Text            =   "Text4"
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtaliquota 
      DataField       =   "Iva"
      DataSource      =   "Adodc6"
      Height          =   360
      Left            =   3960
      TabIndex        =   37
      Text            =   "Text4"
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text3 
      DataField       =   "listino1"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   1
      EndProperty
      DataSource      =   "Adodc6"
      Height          =   375
      Index           =   0
      Left            =   5400
      TabIndex        =   36
      Text            =   "Text3"
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text2 
      DataField       =   "descrizione"
      DataSource      =   "Adodc6"
      Height          =   375
      Index           =   0
      Left            =   4680
      TabIndex        =   35
      Text            =   "Text2"
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txttotale 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   14160
      TabIndex        =   33
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtpezzi 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   12360
      TabIndex        =   31
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txtean 
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
      Left            =   6960
      TabIndex        =   30
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtdesart 
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
      Left            =   8040
      TabIndex        =   28
      Top             =   2880
      Width           =   4335
   End
   Begin VB.TextBox txtordine 
      Alignment       =   1  'Right Justify
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
      Left            =   12360
      TabIndex        =   26
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtfidelity 
      Alignment       =   1  'Right Justify
      CausesValidation=   0   'False
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
      Left            =   6840
      TabIndex        =   24
      Top             =   7440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdcassa 
      BackColor       =   &H0080C0FF&
      Caption         =   "Caisse"
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
      Left            =   14760
      Picture         =   "frmcommande.frx":1E71
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8040
      Width           =   1335
   End
   Begin VB.TextBox txttotiva 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12240
      TabIndex        =   19
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox txttotordine 
      Alignment       =   1  'Right Justify
      DataSource      =   "Adodc4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13440
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   7440
      Width           =   1335
   End
   Begin VB.TextBox txttotimponibile 
      Alignment       =   1  'Right Justify
      DataSource      =   "Adodc4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00FFC0FF&
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
      Left            =   7440
      Picture         =   "frmcommande.frx":3BB3
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Supprimer ligne"
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
      Left            =   10560
      Picture         =   "frmcommande.frx":3CFD
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton cmdesci 
      BackColor       =   &H0080FF80&
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
      Left            =   13440
      Picture         =   "frmcommande.frx":3E47
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdmemo 
      BackColor       =   &H00FF8080&
      Caption         =   "Valider"
      Height          =   735
      Left            =   8880
      Picture         =   "frmcommande.frx":5B89
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8040
      Width           =   1695
   End
   Begin VB.TextBox txtprice 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   13080
      TabIndex        =   11
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdcerca 
      Caption         =   "...."
      Height          =   615
      Left            =   6240
      Picture         =   "frmcommande.frx":7F7B
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2640
      Width           =   735
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmcommande.frx":9C8D
      Height          =   360
      Left            =   7320
      TabIndex        =   8
      Top             =   2280
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "Descrizione"
      BoundColumn     =   ""
      Text            =   ""
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmcommande.frx":9CA2
      Height          =   375
      Left            =   7800
      TabIndex        =   7
      Top             =   7680
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   39
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
         DataField       =   "Articolo"
         Caption         =   "Articolo"
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
         DataField       =   "Cliente"
         Caption         =   "Cliente"
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
         DataField       =   "Numero Ordine"
         Caption         =   "Numero Ordine"
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
         DataField       =   "Data Ordine"
         Caption         =   "Data Ordine"
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
         DataField       =   "Data Consegna"
         Caption         =   "Data Consegna"
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
         DataField       =   "Numero Pezzi"
         Caption         =   "Numero Pezzi"
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
         DataField       =   "Prezzo"
         Caption         =   "Prezzo"
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
         DataField       =   "Sabbiatura"
         Caption         =   "Sabbiatura"
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
         DataField       =   "Zincatura"
         Caption         =   "Zincatura"
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
         DataField       =   "Verniciatura"
         Caption         =   "Verniciatura"
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
         DataField       =   "UM"
         Caption         =   "UM"
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
         DataField       =   "Aliquota"
         Caption         =   "Aliquota"
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
         DataField       =   "Data1"
         Caption         =   "Data1"
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
         DataField       =   "Destinazione"
         Caption         =   "Destinazione"
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
         DataField       =   "Riferimento"
         Caption         =   "Riferimento"
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
         DataField       =   "Importo"
         Caption         =   "Importo"
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
      BeginProperty Column18 
         DataField       =   "CodiceCliente"
         Caption         =   "CodiceCliente"
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
         DataField       =   "vuoto"
         Caption         =   "vuoto"
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
         DataField       =   "Codice prodotto"
         Caption         =   "Codice prodotto"
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
         DataField       =   "Ubicazione"
         Caption         =   "Ubicazione"
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
      BeginProperty Column22 
         DataField       =   "flagstampato"
         Caption         =   "flagstampato"
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
      BeginProperty Column23 
         DataField       =   "flagcancellato"
         Caption         =   "flagcancellato"
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
      BeginProperty Column24 
         DataField       =   "flagfatturato"
         Caption         =   "flagfatturato"
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
      BeginProperty Column25 
         DataField       =   "colore"
         Caption         =   "colore"
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
      BeginProperty Column26 
         DataField       =   "base"
         Caption         =   "base"
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
      BeginProperty Column27 
         DataField       =   "altezza"
         Caption         =   "altezza"
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
      BeginProperty Column28 
         DataField       =   "quantità"
         Caption         =   "quantità"
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
      BeginProperty Column29 
         DataField       =   "totaledocumento"
         Caption         =   "totaledocumento"
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
      BeginProperty Column30 
         DataField       =   "Numero Fattura"
         Caption         =   "Numero Fattura"
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
      BeginProperty Column31 
         DataField       =   "Data Fattura"
         Caption         =   "Data Fattura"
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
      BeginProperty Column32 
         DataField       =   "progressivo"
         Caption         =   "progressivo"
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
      BeginProperty Column33 
         DataField       =   "Corpo"
         Caption         =   "Corpo"
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
      BeginProperty Column34 
         DataField       =   "Numero DDT"
         Caption         =   "Numero DDT"
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
      BeginProperty Column35 
         DataField       =   "Data DDT"
         Caption         =   "Data DDT"
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
      BeginProperty Column36 
         DataField       =   "sconto1"
         Caption         =   "sconto1"
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
      BeginProperty Column37 
         DataField       =   "sconto2"
         Caption         =   "sconto2"
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
      BeginProperty Column38 
         DataField       =   "sconto3"
         Caption         =   "sconto3"
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
            ColumnWidth     =   1454,74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   929,764
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   1425,26
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   1544,882
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   1305,071
         EndProperty
         BeginProperty Column25 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column26 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column27 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column28 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column29 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column30 
            ColumnWidth     =   1665,071
         EndProperty
         BeginProperty Column31 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column32 
            ColumnWidth     =   1454,74
         EndProperty
         BeginProperty Column33 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column34 
            ColumnWidth     =   1454,74
         EndProperty
         BeginProperty Column35 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column36 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column37 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column38 
            ColumnWidth     =   2775,118
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      DataField       =   "NumeroMassimo"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   1
      EndProperty
      DataSource      =   "Adodc3"
      Height          =   360
      Index           =   0
      Left            =   5880
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtdestino 
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
      Left            =   11520
      TabIndex        =   1
      Top             =   2280
      Width           =   3735
   End
   Begin VB.TextBox txtdtaordine 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   3
      EndProperty
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
      Left            =   11040
      TabIndex        =   0
      Top             =   1680
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   9840
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      RecordSource    =   "listapatologie"
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   495
      Left            =   12840
      Top             =   5880
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
      RecordSource    =   "select * from anamnesi where numric=9999999"
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
   Begin MSDataGridLib.DataGrid DataGrid5 
      Bindings        =   "frmcommande.frx":9CB7
      Height          =   3975
      Left            =   6600
      TabIndex        =   4
      Top             =   3240
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7011
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16777215
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   20
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
         Caption         =   "Produit"
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
         Caption         =   "Prix"
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
         Caption         =   "Montant"
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
         Caption         =   "N."
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
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column14 
            Object.Visible         =   0   'False
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column15 
            Object.Visible         =   0   'False
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column16 
            Object.Visible         =   0   'False
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column17 
            Alignment       =   1
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column18 
            Object.Visible         =   0   'False
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column19 
            Alignment       =   1
            ColumnWidth     =   2085,166
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   495
      Left            =   6960
      Top             =   0
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
      RecordSource    =   "select max(numric) as NumeroMassimo FROM anamnesi"
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   495
      Left            =   8880
      Top             =   7680
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      RecordSource    =   "SELECT * FROM Dettaglio WHERE [numero ordine]=''"
      Caption         =   "Adodc4"
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   11040
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from tires"
      Caption         =   "Adodc6"
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
   Begin MSAdodcLib.Adodc Adoazienda 
      Height          =   330
      Left            =   8640
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      RecordSource    =   "Azienda"
      Caption         =   "Adoazienda"
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
   Begin MSAdodcLib.Adodc Adocatego 
      Height          =   330
      Left            =   5160
      Top             =   -240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      RecordSource    =   "Submenu"
      Caption         =   "Adocatego"
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
   Begin MSAdodcLib.Adodc Adoticket 
      Height          =   330
      Left            =   120
      Top             =   8040
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      RecordSource    =   "ticket"
      Caption         =   "Adoticket"
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
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "la Caisse"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   5760
      TabIndex        =   98
      Top             =   0
      Width           =   9495
   End
   Begin VB.Label lblmontant 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Montant"
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
      Left            =   14160
      TabIndex        =   34
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblpezzi 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No"
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
      Left            =   12360
      TabIndex        =   32
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lblcode 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CodeBarre"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6960
      TabIndex        =   29
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblordine 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No."
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
      Left            =   12360
      TabIndex        =   27
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Carte de Fidélité"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   25
      Top             =   7200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lbltotiva 
      Alignment       =   2  'Center
      Caption         =   "Taxes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12240
      TabIndex        =   22
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label lbltotordine 
      Alignment       =   2  'Center
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13440
      TabIndex        =   21
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Montant"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10800
      TabIndex        =   20
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Prix"
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
      Index           =   0
      Left            =   13080
      TabIndex        =   12
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Collaborateur"
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
      Index           =   0
      Left            =   7320
      TabIndex        =   9
      Top             =   2040
      Width           =   4215
   End
   Begin VB.Label lbldestino 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Commentaire"
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
      Left            =   11520
      TabIndex        =   5
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Label lbldtaordine 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date"
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
      Left            =   11040
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Produit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8040
      TabIndex        =   2
      Top             =   2640
      Width           =   4410
   End
End
Attribute VB_Name = "frmcommande"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub btDialog_Click()
'On Error GoTo 100
    cdl.CancelError = True
    cdl.Flags = cdlCCRGBInit
    cdl.Color = lbcolor.BackColor
    cdl.ShowColor
    Col = cdl.Color
    lbcolor.BackColor = cdl.Color
    lbcolor.Caption = cdl.Color
   
10
    Exit Sub
100
    Resume 10
End Sub

Private Sub cmd1_Click(index As Integer)
On Error Resume Next

    
    txtean = index
    
        Adodc6.Recordset.MoveFirst
  Adodc6.Recordset.Find "chiave = '" & chiave(index) & "'", 0, adSearchForward
 txtpezzi = 1
 
txtdesart = Text2(0)
txtprice = Text3(0)
txtpezzi = 1
        
If txtdesart = "" Then GoTo dopo

calcola

cmdmemo_Click



     
dopo:


End Sub

Private Sub cmdclavier_Click(index As Integer)
Dim stringa As String
stringa = "select * from tires where categorie = '" & cmdclavier(index).Caption & "'"

With Adodc6
    .RecordSource = stringa
    .Refresh
    End With
    
    For i = 0 To 53
    cmd1(i).Visible = False
    Next i
    
    Form_Load
    
End Sub

Private Sub Cmddel_Click()
Dim stringa As String


On Error GoTo dopo

response = MsgBox("Supprimèr ticket?", vbOKCancel + vbCancel, "Attenzione")
Select Case response
 Case 6
stringa = "delete * from anamnesi WHERE numric=" & txtordine

rs.Open stringa, cn

calcola


 Case Else
 
 End Select
 
 
dopo:
End Sub

Private Sub cmdDelete_Click()
On Error GoTo dopo

response = MsgBox("Supprimèr?", vbOKCancel + vbCancel, "Attenzione")
Select Case response
 Case 6
 Adodc5.Recordset.delete


 Case Else
 
 End Select
 
 stringa = "select * from anamnesi where numric=" & txtordine


   With Adodc5
    .RecordSource = stringa
    .Refresh
    End With
    
    With DataGrid5
     .ClearFields
     .HoldFields
     .ReBind
     End With
 
 
 calcola
  
dopo:
  
End Sub

Private Sub cmdstampa_Click()
Set DataReport21.DataSource = Adodc5


With DataReport21
  With .Sections("IntestazionePagina").Controls
              .Item("label10").Caption = txttestata
               .Item("label2").Caption = txtordine
                .Item("label8").Caption = txtdtaordine
         End With
         
     With .Sections("PièdiPaginaReport").Controls
      .Item("label1").Caption = txttotordine
              .Item("label3").Caption = Text10
               .Item("label5").Caption = Text11
                .Item("label6").Caption = Text12
     End With
    
              
    .Show
 End With
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo dopo





DataGrid5.Columns(3) = txtdtaordine
DataGrid5.Columns(4) = DataCombo1
DataGrid5.Columns(5) = txtdestino
DataGrid5.Columns(6) = txtdesart
DataGrid5.Columns(7) = txtean
DataGrid5.Columns(9) = txtprice
DataGrid5.Columns(10) = txtaliquota
DataGrid5.Columns(17) = txttotale
DataGrid5.Columns(19) = txtpezzi
DataGrid5.Columns(18) = txtordine


Adodc5.Recordset.Update



dopo:
calcola

End Sub

Private Sub cmdesci_Click()
Unload Me
End Sub

Private Sub Command7_Click()
Text7 = txtdesc
Text8 = cdl.Color
Adodc6.Recordset!itemdesc = txtdesc
Adodc6.Recordset!bcolor = cdl.Color
Adodc6.Recordset.Update
Frame2.Visible = False

End Sub

Private Sub Command8_Click()
Frame2.Visible = False
End Sub

Private Sub DataCombo1_Change()

'calcola
End Sub





Private Sub cmdBack_Click()
'cn.Close
Unload Me
End Sub

Private Sub cmdmemo_Click()
On Error GoTo dopo


Adodc5.Recordset.AddNew

DataGrid5.Columns(3) = txtdtaordine
DataGrid5.Columns(4) = DataCombo1
DataGrid5.Columns(5) = txtdestino
DataGrid5.Columns(6) = txtdesart
DataGrid5.Columns(7) = txtean
DataGrid5.Columns(9) = txtprice
DataGrid5.Columns(10) = txtaliquota
DataGrid5.Columns(17) = txttotale
DataGrid5.Columns(19) = txtpezzi
DataGrid5.Columns(18) = txtordine
Adodc5.Recordset.Update



DataGrid5.Visible = True


dopo:

calcola



End Sub



Private Sub cmdcerca_Click()
Dim stringa As String
variabile3 = 1 ' prezzo 1

Form6.Show vbModal

On Error Resume Next

txtcodart = art1
txtdesart = art2
txtpezzi = 1
txtprice = Format(art4, "###,###,##0.00")


txtaliquota = art8


     
     calcola
     

'DataGrid2.Enabled = True
DataCombo1.Enabled = True
txtpezzi.Enabled = True
cmdUpdate.Enabled = False
cmddelete.Enabled = False
cmdmemo.Enabled = True

 


End Sub




 



Private Sub cmdcassa_Click()
frame11.Visible = True

txtTotal = txttotordine
variabile2 = txtordine
'frmTotal.Show vbModal

End Sub

Private Sub DataCombo2_Change()
 txtprice = Format(DataCombo2.BoundText, "###,##0.00")
 cmdmemo.Enabled = True

End Sub

Private Sub DataGrid5_Click()
On Error GoTo dopo




 txtdtaordine = DataGrid5.Columns(3)
DataCombo1 = DataGrid5.Columns(4)
txtdestino = DataGrid5.Columns(5)
txtdesart = DataGrid5.Columns(6)
txtean = DataGrid5.Columns(7)
txttotale = DataGrid5.Columns(17)
txtprice = DataGrid5.Columns(9)
txtaliquota = DataGrid5.Columns(10)
txtpezzi = DataGrid5.Columns(19)
txtordine = DataGrid5.Columns(18)

cmdmemo.Enabled = False
cmdUpdate.Enabled = True
cmddelete.Enabled = True


dopo:
End Sub

Private Sub Form_Load()
On Error Resume Next
  cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False;"
   cn.Open
   
 Dim stringa As String
 
 
Adocatego.Recordset.MoveFirst
For i = 0 To 8
 cmdclavier(i).Caption = txtcatego
 If Not Adocatego.Recordset.EOF Then
  Adocatego.Recordset.MoveNext
   End If
   Next i
  
   
   

    
Dim l, h As Integer
l = 0: h = 0
cmd1(0).Visible = False

Adodc6.Recordset.MoveFirst

For i = 0 To 53

 
 If Adodc6.Recordset.EOF Then Exit For
 
  txtpezzi = 1
Load cmd1(i)

' Sposta il controllo e lo ridimensiona.
cmd1(i).Move 0 + (l * 1000), 0 + (h * 1000), 1000, 1000
' Imposta altre proprietà se necessario
cmd1(i).Caption = Text7 & Chr$(13) & "€ " & Text9
 cmd1(i).BackColor = Text8
  cmd1(i).Picture = txtpic
     
    If chkvisible.Value > 0 Then cmd1(i).Visible = True
   
     
   
    
  chiave(i) = txtchiave
' Infine rende visibile il controllo.

l = l + 1
If l = 6 Then l = 0: h = h + 1


'Cmd1(i).BackColor = &HFFFF&


Adodc6.Recordset.MoveNext


Next i


'DataGrid2.Enabled = False
cmdmemo.Enabled = False
'DataGrid5.Enabled = False


frame11.Visible = False


txtdtaordine = Date
txtordine = Val(Text1(0)) + 1



a = 0: b = 0: c = 0

txtean.SetFocus

End Sub





Private Sub txtaltezza_Validate(Cancel As Boolean)
 If Not IsNumeric(txtaltezza.Text) Then
        Cancel = True
   
    End If
    If Cancel Then
        MsgBox "Introdurre un valore numerico ", vbExclamation
    End If
End Sub

Private Sub txtbase_Validate(Cancel As Boolean)

    If Not IsNumeric(txtbase.Text) Then
        Cancel = True
   
    End If
    If Cancel Then
        MsgBox "Introdurre un valore numerico ", vbExclamation
    End If
End Sub

Private Sub txtconsegna_Validate(Cancel As Boolean)
    ' Prepare to edit in short-date format.
    On Error Resume Next
    txtconsegna.Text = Format$(CDate(txtconsegna.Text), "short date")
End Sub




Private Sub Touch21_Click()

End Sub

Private Sub txtAmount_Change()

End Sub



Private Sub txtpezzi_Change()
cmdmemo.Enabled = True
End Sub



Public Sub calcola()
On Error Resume Next

txttotale = Format(txtpezzi * txtprice, "###,###,##0.00")


  stringa = "select * from anamnesi where numric=" & txtordine


   With Adodc5
    .RecordSource = stringa
    .Refresh
    End With
    
    With DataGrid5
     .ClearFields
     .HoldFields
     .ReBind
     End With


a = 0:

Adodc5.Recordset.MoveFirst
 For i = 0 To Adodc5.Recordset.RecordCount - 1
  a = a + DataGrid5.Columns(17)
  
     Adodc5.Recordset.MoveNext
   Next i
   
    
   txttotimponibile = Format(a, "###,##0.00")
   txttotordine = Format(a, "###,##0.00")
   
 
DataGrid5.Visible = True
   

End Sub



Private Sub txtcodart_Change()

End Sub

Private Sub txtean_KeyDown( _
           KeyCode As Integer, Shift As Integer)

On Error Resume Next

     Select Case KeyCode
     Case vbKeyReturn:
     
        Adodc6.Recordset.MoveFirst
  Adodc6.Recordset.Find "codiceEAN = '" & txtean.Text & "'", 0, adSearchForward
 txtpezzi = 1
 
txtdesart = Text2(0)
txtprice = Text3(0)
txtpezzi = 1
        
If txtdesart = "" Then GoTo dopo

calcola

cmdmemo_Click


dopo:
 
txtean = ""


     End Select

End Sub

Private Sub txtpezzi_Validate(Cancel As Boolean)
calcola
End Sub

Private Sub txttotiva_Validate(Cancel As Boolean)
On Error Resume Next

txttotordine = CDbl(txttotimponibile) + CDbl(txttotiva)

txttotiva = Format(txttotiva, "###,##0.00")

txttotordine = Format(txttotordine, "###,##0.00")


End Sub

Private Sub VScroll1_Change()

End Sub
Private Sub Command1_Click(index As Integer)
Select Case index
    Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 500, 1000, 2000, 5000
        Me.txtAmtTendered.Text = Me.txtAmtTendered.Text & index
    Case 10
        Me.txtAmtTendered.Text = ""
        Me.txtChange = ""
        Me.txtPmt.Text = ""
        Frame1.Visible = False
    Case 11
        Me.txtAmtTendered.Text = Me.txtAmtTendered.Text & ","
End Select
End Sub

Private Sub Command2_Click()
Dim i As Integer
Dim stringa, stringa1 As String



If txtAmtTendered = "" Then
frame11.Visible = False
Exit Sub
End If

'On Error GoTo dopo



Adodc9.Recordset.AddNew


DataGrid2.Columns(3) = txtdtaordine
DataGrid2.Columns(14) = txtTotal
DataGrid2.Columns(15) = txtPmt
'DataGrid2.Columns(16) = txtDiscount
DataGrid2.Columns(17) = txtTotal
DataGrid2.Columns(4) = DataCombo1



Adodc9.Recordset.Update

 stringa1 = "update tires, anamnesi set giacenza = giacenza - anamnesi.pezzi,tires.dataultimavendita = date  WHERE numric=" & txtordine & " and date=#" & txtdtaordine & "# and flag=false and tires.descrizione=anamnesi.service"
  
 rs.Open stringa1, cn

stringa = "update anamnesi set flag=true, numric=" & txtordine & " WHERE clientcode='" & variabile1 & "' and date=#" & txtdtaordine & "# and flag=false"

rs.Open stringa, cn





Set f = New frmpos
f.Show
Unload Me

dopo:


DataGrid5.Visible = False

End Sub

Private Sub Command3_Click()
If Me.txtAmtTendered = "" Then
MsgBox ("Paiement!!")
Else
txtPmt = "Espèce"
txtChange.Text = txtAmtTendered.Text - txtTotal.Text
Frame1.Visible = False
End If
End Sub

Private Sub Command4_Click()
If Me.txtAmtTendered = "" Then
MsgBox ("Paiement!!")
Else
txtPmt = "Carte"
txtChange.Text = txtAmtTendered.Text - txtTotal.Text
Frame1.Visible = False
End If
End Sub

Private Sub Command5_Click()
On Error Resume Next
If Me.txtAmtTendered = "" Then
MsgBox ("Paiement!!")
Else
'Frame1.Visible = True
'Me.Caption = "Please Enter Cheque Details"
txtPmt = "Cheque"
txtChange.Text = txtAmtTendered.Text - txtTotal.Text
End If
End Sub

Private Sub Command6_Click(index As Integer)

    Select Case index
        Case 5, 10, 20, 50
            If Not Me.txtAmtTendered.Text = "" Then
                Me.txtAmtTendered.Text = index + Me.txtAmtTendered.Text
            Else
                Me.txtAmtTendered.Text = index
            End If
    End Select

End Sub

