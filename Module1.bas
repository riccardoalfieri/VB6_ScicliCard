Attribute VB_Name = "Module1"
Public fMainForm As frmMain

Public cn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public rs1 As New ADODB.Recordset
Global variabile1, variabile2, variabile3, variabile4, variabile5, variabile6 As String
Global art1, art2, art3, art4, art5, art6, art7, art8, art9, art10, art11 As String
Global a, b, c As Single
Global valore1, valore2, valore3 As Single
Global varia1, varia2, varia3, varia4, varia5 As Single
Global righe As Integer
 
Global somma1, somma2, somm3 As Single






Sub Main()

   
    Set fMainForm = New frmMain
    fMainForm.Show
End Sub

