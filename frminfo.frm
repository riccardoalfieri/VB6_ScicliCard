VERSION 5.00
Begin VB.Form frminfo 
   BackColor       =   &H008080FF&
   Caption         =   "Richiesta dati carta"
   ClientHeight    =   2085
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8235
   Icon            =   "frminfo.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2085
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
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
      Left            =   2280
      TabIndex        =   2
      Top             =   720
      Width           =   5535
   End
   Begin VB.TextBox txtPIN 
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2415
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1455
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
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
End
Attribute VB_Name = "frminfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
 
frmtitolare.Show vbModal
 Unload Me
End Sub

Private Sub Form_Load()
If azienda = "" Then
Form1.Show vbModal
End If

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
