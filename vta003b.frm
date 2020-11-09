VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form vta_facturacion2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Composicion de IVA"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   3930
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox t_modulo 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton btnsale 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3360
      Picture         =   "vta003b.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir sin Modificar"
      Top             =   3120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   5106
      _Version        =   393216
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "vta_facturacion2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnsale_Click()
Me.Hide
End Sub

Private Sub Form_Load()
Call armagrid
End Sub

Sub armagrid()
  msf1.clear
  msf1.Rows = 10
  msf1.Cols = 3
  msf1.ColWidth(0) = 1000
  msf1.ColWidth(1) = 1300
  msf1.ColWidth(2) = 1300
  msf1.TextMatrix(0, 0) = "Tasa"
  msf1.TextMatrix(0, 1) = "Neto"
  msf1.TextMatrix(0, 2) = "Iva"
  Call cargatasa
  'en fila 9 tiene los totales
  msf1.TextMatrix(9, 0) = "Iva Total"
  msf1.TextMatrix(8, 1) = "--------------------------------"
  msf1.TextMatrix(8, 2) = "--------------------------------"
End Sub
Sub cargatasa()
Set rs = New ADODB.Recordset
q = "select * from g4 "
rs.Open q, cn1
c = 1
While Not rs.EOF
  msf1.TextMatrix(c, 0) = Format$(rs("tasa"), "#0.00")
  c = c + 1
  rs.MoveNext
Wend
Set rs = Nothing
End Sub
Sub sacatotales()
x = 0
nt = 0
IT = 0
For i = 1 To 7
    msf1.TextMatrix(i, 1) = Format$(msf1.TextMatrix(i, 1), "######0.00")
    msf1.TextMatrix(i, 2) = Format$(msf1.TextMatrix(i, 2), "######0.00")
    nt = nt + Val(msf1.TextMatrix(i, 1))
    IT = IT + Val(msf1.TextMatrix(i, 2))
Next i
msf1.TextMatrix(9, 1) = Format$(nt, "######0.00")
msf1.TextMatrix(9, 2) = Format$(IT, "######0.00")

End Sub

Private Sub msf1_DblClick()
  If msf1.col = 1 Then
    If Val(msf1.TextMatrix(msf1.Row, 0)) > 0 Then
      d = InputBox$("Ingrese Neto gravado para la tasa", "MODIFICACION DE VALORES DE IVA ")
      If Val(d) >= 0 Then
        msf1.TextMatrix(msf1.Row, msf1.col) = Format(Val(d), "#####0.00")
        msf1.TextMatrix(msf1.Row, msf1.col + 1) = Format((Val(d) * Val(msf1.TextMatrix(msf1.Row, 0))) / 100, "#####0.00")
        Call sacatotales
      End If
    End If
  End If

End Sub
