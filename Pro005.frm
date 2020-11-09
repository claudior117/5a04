VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form prod_detalle_pedidos 
   BackColor       =   &H00E0E0E0&
   Caption         =   "SEGUIMIENTO DE PEDIDOS"
   ClientHeight    =   8775
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12180
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   12180
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9960
      TabIndex        =   8
      Top             =   7200
      Width           =   1575
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "Pro005.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "Pro005.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   1215
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   6495
      Begin VB.TextBox t_producto 
         Height          =   285
         Left            =   2280
         TabIndex        =   7
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox t_idproducto 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox t_referencia 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         Caption         =   "Producto:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "Nro. referencia:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9975
      _Version        =   393216
      HighLight       =   2
      AllowUserResizing=   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   8520
      Width           =   12180
      _ExtentX        =   21484
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   11465
            MinWidth        =   11465
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "05/10/2012"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "9:50"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "prod_detalle_pedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim EXISTE As String
Dim rs2 As Recordset
Sub limpia()
   Call armagrid
   t_subtotal = ""
   t_nograbado = ""
   t_perc = ""
   t_iva = ""
   t_total = ""
  
End Sub



Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 11
msf1.ColWidth(0) = 1500
msf1.ColWidth(1) = 3000
msf1.ColWidth(2) = 1400
msf1.ColWidth(3) = 700
msf1.ColWidth(4) = 1400
msf1.ColWidth(5) = 2500
msf1.ColWidth(6) = 2000
msf1.ColWidth(7) = 2000
msf1.ColWidth(8) = 1000
msf1.ColWidth(9) = 2500
msf1.ColWidth(10) = 800


msf1.TextMatrix(0, 0) = "Fecha"
msf1.TextMatrix(0, 1) = "Comprobante"
msf1.TextMatrix(0, 2) = "Cantidad"
msf1.TextMatrix(0, 3) = "Unid."
msf1.TextMatrix(0, 4) = "P.U."
msf1.TextMatrix(0, 5) = "Proveedor/Obra"
msf1.TextMatrix(0, 6) = "Usuario"
msf1.TextMatrix(0, 7) = "Observaciones"
msf1.TextMatrix(0, 8) = "Num.Int."
msf1.TextMatrix(0, 9) = "Detalle"
msf1.TextMatrix(0, 10) = "Modulo"

End Sub


 

Private Sub btnsale_Click()
Me.Hide
End Sub

Private Sub Form_Activate()
Call busca
End Sub
Sub carga_pedido()
     Set cl_usuarios = New usuarios
     cl_usuarios.cargar (rs2("id_usuario"))
     If cl_usuarios.idusuario > 0 Then
       u = cl_usuarios.denominACION
     Else
       u = "Inexistente"
     End If
     Set cl_usuarios = Nothing
     
     f = rs2("fecha")
     tc = rs2("abreviatura")
     nc = tc & " " & Format$(rs2("sucursal"), "0000") & "-" & Format$(rs2("num_comprobante"), "00000000")
     cp = Format$(rs2("pro_01.id_obra"), "0000")
     p = rs2("a4.descripcion")
     ni = rs2("pro_01.num_int")
     o = rs2("pro_01.observaciones")
     c = rs2("cantidad")
     un = rs2("unidad")
     pu = "0.00"
     msf1.AddItem f & Chr(9) & nc & Chr(9) & c & Chr(9) & un & Chr(9) & pu & Chr(9) & p & Chr(9) & u & Chr(9) & o & Chr(9) & ni & Chr(9) & rs("obs") & Chr(9) & "P"

End Sub
Sub carga_compras()
     Set cl_usuarios = New usuarios
     cl_usuarios.cargar (rs2("a5.id_usuario"))
     If cl_usuarios.idusuario > 0 Then
       u2 = cl_usuarios.denominACION
     Else
       u2 = "Inexistente"
     End If
     Set cl_usuarios = Nothing
     
     
     f = rs2("a5.fecha")
     tc = rs2("abreviatura")
     nc = tc & " " & Format$(rs2("sucursal"), "0000") & "-" & Format$(rs2("num_comprobante"), "00000000")
     cp = Format$(rs2("a1.id_proveedor"), "0000")
     p = rs2("denominacion")
     ni = rs2("a5.num_int")
     o = rs2("observaciones")
     c = rs2("cantidad")
     pu = rs2("pu")
     u = rs2("unidad")
     msf1.AddItem f & Chr(9) & nc & Chr(9) & c & Chr(9) & u & Chr(9) & pu & Chr(9) & p & Chr(9) & u2 & Chr(9) & o & Chr(9) & ni & Chr(9) & rs("obs") & Chr(9) & "C"

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
  
End Select
End Sub
Sub busca()
Call armagrid
If Val(t_referencia) > 0 Then
   q = "select * from pro_05 where [num_referencia] = " & Val(t_referencia) & " order by [fecha]"
   Set rs = New ADODB.Recordset
   rs.Open q, cn1
   While Not rs.EOF
     Select Case rs("modulo")
     Case Is = "P"
         'pedido
         Set rs2 = New ADODB.Recordset
         q = "select * from pro_02, pro_01, pro_03, a4 where   pro_01.[id_tipocomp] = " & rs("tipo_comprobante") & " and pro_01.[num_int] = " & rs("num_int") & " and pro_02.[num_referencia] = " & Val(t_referencia) & " and pro_02.[num_int] = pro_01.[num_int] and pro_01.[id_tipocomp] = pro_03.[id_tipocomp] and pro_01.[id_obra] = a4.[id_obra] "
         rs2.Open q, cn1
         If Not rs2.EOF And Not rs2.BOF Then
           Call carga_pedido
         End If
         Set rs2 = Nothing
     Case Else
         'oc o recepcion
         Set rs2 = New ADODB.Recordset
         q = "select * from a6, a5, g2, a1 where [num_int_item] = " & Val(t_referencia) & " and  [a5.num_int] = " & rs("num_int") & " and a5.[num_int]= a6.[num_int] and a5.[id_tipocomp] = g2.[id_tipo_comp] and a5.[id_proveedor] = a1.[id_proveedor] "
         rs2.Open q, cn1
         If Not rs2.EOF And Not rs2.BOF Then
           Call carga_compras
         End If
         Set rs2 = Nothing
     End Select
     rs.MoveNext
   Wend
End If



End Sub
Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = ""
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If

End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    If msf1.TextMatrix(msf1.Row, 10) = "P" Then
      'produccion
       If para.id_grupo_modulo_actual >= 3 Then
        Load prod_cc_detalle
        prod_cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 8)
        prod_cc_detalle.Show
       End If
    Else
      'compras
      If para.id_grupo_modulo_compras >= 3 Then
       Load cc_detalle
       cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 8)
       cc_detalle.Show
      End If
    End If
  End If
End If
End Sub
