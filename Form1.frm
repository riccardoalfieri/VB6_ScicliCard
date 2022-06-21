VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Internet"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   1935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   960
      Top             =   240
   End
   Begin VB.CommandButton cmdHangup 
      Caption         =   "Hang up connection"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   360
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Internet_Autodial_Force_Unattended As Long = 2
Dim lResult As Long

Private Sub Command2()
 temp = GetIPAddress
 If temp <> Text1.Text Then
  Text1.Text = temp
  If Text1.Text <> "127.0.0.1" Then
   cmdHangup.Enabled = True
   Unload Me
   MsgBox "Waiting for connection...."
   
  frmLogin1.Show vbModal

  Else
   cmdHangup.Enabled = False
   Unload Me
   MsgBox "You are not connected to the internet"
  End If
 End If
End Sub

Private Sub cmdHangup_Click()
 lResult = InternetAutodialHangup(0&)
 cmdHangup.Enabled = False
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
 Text1.Text = 0
 Text1.Enabled = False
 Timer1.Enabled = True
 Timer1.Interval = 1
End Sub

Private Sub Timer1_Timer()
 Call Command2
End Sub
