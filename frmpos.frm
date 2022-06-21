VERSION 5.00
Begin VB.Form frmpos 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Calcolo sconto"
   ClientHeight    =   9645
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   14790
   Icon            =   "frmpos.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   14790
   WindowState     =   2  'Maximized
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
      Height          =   7455
      Left            =   8640
      Picture         =   "frmpos.frx":49BEE
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   360
      Width           =   855
   End
   Begin VB.Frame frame11 
      Caption         =   "Cassa"
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   8535
      Begin VB.TextBox txtPIN 
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "*"
         TabIndex        =   31
         Top             =   6480
         Width           =   2415
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
         TabIndex        =   26
         Top             =   2160
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
         TabIndex        =   25
         Top             =   3240
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
         TabIndex        =   24
         Top             =   3240
         Width           =   1095
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
         TabIndex        =   23
         Top             =   2160
         Width           =   1095
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
         TabIndex        =   22
         Top             =   5280
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Assegno"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2040
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Carta Credito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sconto "
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
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FF00&
         Caption         =   "ESEGUI"
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
         TabIndex        =   18
         Top             =   5280
         Width           =   2295
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
         TabIndex        =   17
         Top             =   5280
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
         TabIndex        =   16
         Top             =   5280
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   4200
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
         TabIndex        =   12
         Top             =   4200
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   3120
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
         TabIndex        =   8
         Top             =   3120
         Width           =   1215
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
         TabIndex        =   7
         Top             =   5280
         Width           =   1215
      End
      Begin VB.TextBox txtChange 
         Alignment       =   1  'Right Justify
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
         Left            =   3000
         TabIndex        =   6
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H000000C0&
         Height          =   855
         Left            =   1920
         TabIndex        =   4
         Top             =   2160
         Width           =   3495
      End
      Begin VB.TextBox txtAmtTendered 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   3
         Top             =   240
         Width           =   3495
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Altro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF8080&
         Caption         =   "Invia dati"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   6360
         Width           =   1455
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
         IMEMode         =   3  'DISABLE
         Left            =   2520
         TabIndex        =   5
         Top             =   6360
         Width           =   5535
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Inserire Card"
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
         Left            =   240
         TabIndex        =   30
         Top             =   6480
         Width           =   2175
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Da Pagare"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   0
         TabIndex        =   29
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sconto"
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
         TabIndex        =   28
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Totale"
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
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmpos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)
Select Case Index
    Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 500, 1000, 2000, 5000
        Me.txtAmtTendered.Text = Me.txtAmtTendered.Text & Index
    Case 10
        Me.txtAmtTendered.Text = ""
        Me.txtChange = ""
        Me.txtPmt.Text = ""
        Me.txtTotal.Text = ""
        'Frame1.Visible = False
    Case 11
        Me.txtAmtTendered.Text = Me.txtAmtTendered.Text & ","
End Select
End Sub

Private Sub Command2_Click()

txtPmt.Visible = True
Label5.Visible = True
txtPmt.SetFocus

End Sub

Private Sub Command8_Click()
On Error Resume Next
If Me.txtAmtTendered = "" Then
MsgBox (messaggi(3))
Else
'Frame1.Visible = True
'Me.Caption = "Please Enter Cheque Details"
txtPmt = pagamenti(4)
txtChange.Text = txtAmtTendered.Text - txtTotal.Text
End If
End Sub

Private Sub Command9_Click()
Command9.Visible = False
txtPIN.Visible = False
Label5.Visible = False

Dim a, b, c, d As Integer
  a = Val(Left(txtPmt, 4))
  b = Val(Mid(txtPmt, 5, 4))
  c = Val(Right(txtPmt, 4))
  If c <> a + b Or Len(txtPmt) <> 12 Then
   MsgBox "Numero Card Errato", vbInformation, "Transazione NON ESEGUITA"
  Exit Sub
  End If
  
  If Len(txtPIN) <> 4 Then
   MsgBox "Numero PIN errato", vbInformation, "Transazione NON ESEGUITA"
  Exit Sub
  End If

 MsgBox "Transazione in corso. Attendere!"
 art1 = txtPmt
 art2 = txtPIN
 txtPIN = ""
 
Select Case Mid(txtPmt, 4, 1) ' la tessera negozianti contiene uno 0 in quarta posizione
  Case "0"
   frm_negozio.Show vbModal
    
     Case Else

      frm_cliente.Show vbModal
 
      End Select
 
End Sub

Private Sub DataCombo1_Change()

'calcola
End Sub





Private Sub cmdBack_Click()
'cn.Close
Unload Me
End Sub

Private Sub Command3_Click()
If Me.txtAmtTendered = "" Then
MsgBox (messaggi(3))
Else
txtPmt = pagamenti(0)
txtChange.Text = Format(txtAmtTendered.Text - txtTotal.Text, "###,##0.00")
Frame1.Visible = False
End If
End Sub

Private Sub Command4_Click()
If Me.txtAmtTendered = "" Then
MsgBox (messaggi(3))
Else
txtPmt = pagamenti(2)
txtChange.Text = Format(txtAmtTendered.Text - txtTotal.Text, "###,##0.00")
Frame1.Visible = False
End If
End Sub

Private Sub Command5_Click()
On Error Resume Next
If Me.txtAmtTendered = "" Then
MsgBox (messaggi(3))
Else
'Frame1.Visible = True
'Me.Caption = "Please Enter Cheque Details"
txtPmt = pagamenti(1)
txtChange.Text = txtAmtTendered.Text - txtTotal.Text
End If
End Sub

Private Sub Command6_Click(Index As Integer)
On Error Resume Next
    Select Case Index
        Case 5, 10, 20, 50
            If Not Me.txtAmtTendered.Text = "" Then
                Me.txtAmtTendered.Text = Format(Index + Me.txtAmtTendered.Text, "###,##0.00")
            Else
                Me.txtAmtTendered.Text = Format(Index, "###,##0.00")
            End If
    End Select

End Sub

Private Sub Form_Load()
txtPmt.Visible = False
Label5.Visible = False
txtPIN.Visible = False
Command9.Visible = False

If azienda = "" Then
Form1.Show vbModal
End If
Command3.Caption = "Sconto: " & SCONTO & "%"
End Sub

Private Sub txtAmtTendered_Change()
Dim valore As Double
On Error Resume Next
valore = SCONTO / 100

 Me.txtChange.Text = Format(txtAmtTendered * valore, "###,##0.00")
 Me.txtTotal.Text = Format(txtAmtTendered - txtChange, "###,##0.00")
 
End Sub

Private Sub txtChange_Change()
On Error Resume Next
  Me.txtTotal.Text = Format(txtAmtTendered - txtChange, "###,##0.00")

valore_sconto = txtChange
End Sub

Private Sub txtPIN_KeyDown( _
           KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
     Case vbKeyReturn:
   
      
art1 = txtPmt
art2 = txtPIN

Command9_Click

End Select

End Sub

Private Sub txtpmt_KeyDown( _
           KeyCode As Integer, Shift As Integer)

     Select Case KeyCode
     Case vbKeyReturn:
     
      txtPIN.Visible = True
      txtPIN.SetFocus
      Label5.Caption = "Inserire PIN"
      txtPmt.Visible = False
       Command9.Visible = True

    End Select
     End Sub


Private Sub txtTotal_KeyDown( _
           KeyCode As Integer, Shift As Integer)

     Select Case KeyCode
     Case vbKeyReturn:
     art1 = txtTotal
    
             txtPmt.SetFocus
               
     End Select
     End Sub

