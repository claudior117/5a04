VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form vta_listaprecios4 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OPERACIONES POR PRODUCTO"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   8520
      TabIndex        =   10
      Top             =   120
      Width           =   3255
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ventas"
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Compras"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   8295
      Begin VB.TextBox t_prod 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2640
         TabIndex        =   15
         Top             =   240
         Width           =   5415
      End
      Begin VB.TextBox t_idprod 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Producto:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   4
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "VTA030.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "VTA030.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   8460
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   8820
            MinWidth        =   8820
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "15/06/2012"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:26 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5535
      Left            =   0
      TabIndex        =   9
      Top             =   1560
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9763
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "vta_listaprecios4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub carga()
 espere.Show
 espere.Label1 = "Espere...... Buscando Informacion de Compras"
 espere.Refresh
 Call armagrid
 q = "select * from A5, A6, a1 where a5.[num_int] = a6.[num_int]  and a5.[id_proveedor] = a1.[id_proveedor] "
 c = " and "
 p = 1
 If t_idprod <> "" Then
  q = q & c & " [id_producto] = " & Val(t_idprod)
 Else
   If t_prod <> "" Then
      q = q & c & " [detalle] like '%" & t_prod & "%'"
   Else
     MsgBox ("Se debe ingresar cod. producto o descripcion")
     p = 0
   End If
 
 End If
 
 
 
  If IsDate(t_fecha) Then
     q = q & c & " datevalue(a5.[fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue(a5.[fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  q = q & " order by a5.[fecha] desc, [letra], [num_comprobante]"
  
  If p = 1 Then
  
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  co = 0
  While Not rs.EOF
     f = Format$(rs("a5.fecha"), "dd/mm/yy")
     nc = rs("letra") & " " & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comprobante"), "00000000")
     p = rs("denominacion")
     m = rs("moneda")
     pu = Format$(rs("pusindto"), "#####0.00")
     d = Format$(rs("descuento"), "###0.00")
     ti = Format$(rs("tasa_iva"), "##0.0")
     c = Format$(rs("cantidad"), "#####0.00")
     pd = Format$(rs("pu"), "#####0.00")
     CUIT = rs("cuit05")
     msf1.AddItem p & Chr(9) & CUIT & Chr$(9) & f & Chr(9) & pu & Chr(9) & d & Chr(9) & ti & Chr(9) & pd & Chr(9) & c & Chr(9) & m & Chr(9) & nc & Chr(9) & rs("a5.num_int")
     co = co + 1
    rs.MoveNext
  Wend
  msf1.AddItem " "
  msf1.AddItem "Cant. Operaciones :" & co
 End If
   Unload espere
     
End Sub

Sub carga2()
 espere.Show
 espere.Label1 = "Espere...... Buscando Informacion de Ventas"
 espere.Refresh
 Call armagrid
 q = "select * from vta_02, vta_03 where vta_02.[num_int] = vta_03.[num_int]  and  [id_producto] = " & Val(t_idprod)
 c = " and "
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  q = q & " order by [fecha] desc, [letra], [num_comp]"
  Set rs = New ADODB.Recordset
   rs.Open q, cn1
  co = 0
  While Not rs.EOF
     f = Format$(rs("fecha"), "dd/mm/yy")
     nc = rs("letra") & " " & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
     p = rs("cliente02")
     m = rs("moneda")
     pu = Format$(rs("pu"), "#####0.00")
     d = Format$(0, "###0.00")
     ti = Format$(rs("tasaiva"), "##0.0")
     c = Format$(rs("cantidad"), "#####0.00")
     pd = Format$(rs("pu_final"), "#####0.00")
     CUIT = rs("cuit02")
     msf1.AddItem p & Chr(9) & CUIT & Chr$(9) & f & Chr(9) & pu & Chr(9) & d & Chr(9) & ti & Chr(9) & pd & Chr(9) & c & Chr(9) & m & Chr(9) & nc & Chr(9) & rs("vta_02.num_int")
     co = co + 1
    rs.MoveNext
  Wend
  msf1.AddItem " "
  msf1.AddItem "Cant. Operaciones :" & co
   Unload espere

End Sub

Private Sub btnacepta_Click()
If Option1 = True Then
  Call nivel_acceso(2) 'compras
  If para.id_grupo_modulo_actual > 5 Then
    Call carga
  Else
    Call sinpermisos
  End If
Else
  Call nivel_acceso(1) 'ventas
  If para.id_grupo_modulo_actual > 5 Then
    Call carga2
  Else
    Call sinpermisos
  End If
End If
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub










Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     gen_tools.Show
End Select
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 11
msf1.ColWidth(0) = 3000
msf1.ColWidth(1) = 1200
msf1.ColWidth(2) = 1000
msf1.ColWidth(3) = 1100
msf1.ColWidth(4) = 900
msf1.ColWidth(5) = 900
msf1.ColWidth(6) = 1100
msf1.ColWidth(7) = 1100
msf1.ColWidth(8) = 500
msf1.ColWidth(9) = 2000
msf1.ColWidth(10) = 800
msf1.TextMatrix(0, 1) = "CUIT"
msf1.TextMatrix(0, 2) = "Fecha"
msf1.TextMatrix(0, 3) = "P.U "
msf1.TextMatrix(0, 4) = "% Dto"
msf1.TextMatrix(0, 5) = "% Iva "
If Option1 = True Then
  msf1.TextMatrix(0, 6) = "P.U. c/ dto"
  msf1.TextMatrix(0, 0) = "Proveedor"
Else
  msf1.TextMatrix(0, 6) = "P.Final"
  msf1.TextMatrix(0, 0) = "Cliente"
End If
msf1.TextMatrix(0, 7) = "Cantidad"
msf1.TextMatrix(0, 8) = "Moneda"
msf1.TextMatrix(0, 9) = "Comprobante"
msf1.TextMatrix(0, 10) = "Num.int"

For i = 1 To 7
  msf1.ColAlignment(i) = 9 'der
Next i
msf1.ColAlignment(0) = 1 'izq
msf1.ColAlignment(9) = 1 'izq


End Sub

Private Sub Form_Load()

Call armagrid
Option1 = True

End Sub



Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F7] Imprime - [F11] Excel -[ENTER] Ver Comp."

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
  Dim c(15) As Double
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    c(0) = 0
    c(1) = 1
    c(2) = 2
    c(3) = 3
    c(4) = 4
    c(5) = 5
    c(6) = 6
    c(7) = 7
    c(8) = 8

    For i = 9 To 14
      c(i) = -1
    Next i
    If Option1 = True Then
      modulo = "COMPRAS"
    Else
      modulo = "VENTAS"
    End If
    Call imprimegrid(msf1, c(), "Operaciones por Producto", "Producto: (" & t_idprod & ") " & t_prod, "Periodo: " & t_fecha & " : " & t_fecha2, "Modulo: " & modulo, 95, 6, True, False)
  End If

End If


If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If
End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Option1 = True Then 'comptras
   If msf1.Row > 0 Then
      Load cc_detalle
      cc_detalle.t_NUMINT = msf1.TextMatrix(msf1.Row, 10)
      cc_detalle.Show
    End If
  Else
   If msf1.Row > 0 Then
      Load vta_cc_detalle
      vta_cc_detalle.t_NUMINT = msf1.TextMatrix(msf1.Row, 10)
      vta_cc_detalle.Show
    End If
  
  End If
End If

End Sub

Private Sub t_fecha_GotFocus()
t_fecha = ""
End Sub

Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
    t_fecha = ""
  End If
End If
End Sub

Private Sub t_fecha2_GotFocus()
t_fecha2 = ""
End Sub

Private Sub t_fecha2_LostFocus()
If t_fecha2 <> "" Then
  If Not IsDate(t_fecha2) Then
    t_fecha2 = ""
  End If
End If

End Sub

Private Sub t_idprod_GotFocus()
t_idprod = ""
End Sub
