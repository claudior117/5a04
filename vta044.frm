VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_compraventa 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "INFORME DE COMPRA VENTAS DE PRODUCTOS"
   ClientHeight    =   8595
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   975
      Left            =   120
      TabIndex        =   30
      Top             =   1440
      Width           =   11655
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   36
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4560
         MaxLength       =   10
         TabIndex        =   35
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox c_cli 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7680
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "Combo1"
         Top             =   600
         Width           =   3855
      End
      Begin VB.ComboBox c_prov 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7680
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "Combo1"
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   37
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H00800000&
         Caption         =   "Cliente:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6480
         TabIndex        =   34
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H00800000&
         Caption         =   "Proveedor"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6480
         TabIndex        =   32
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Vigentes"
      Height          =   615
      Left            =   2760
      TabIndex        =   25
      Top             =   7440
      Width           =   3375
      Begin VB.OptionButton Option7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Caducados"
         Height          =   195
         Left            =   2160
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Vigentes"
         Height          =   195
         Left            =   1080
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Buscar"
      Height          =   855
      Left            =   10080
      TabIndex        =   23
      Top             =   7320
      Width           =   1695
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   960
         Picture         =   "vta044.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   495
         Left            =   120
         Picture         =   "vta044.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtros Varios"
      Height          =   615
      Left            =   6240
      TabIndex        =   21
      Top             =   7440
      Width           =   2535
      Begin VB.ComboBox c_tipo 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "vta044.frx":1104
         Left            =   120
         List            =   "vta044.frx":111D
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "Combo1"
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Height          =   4815
      Left            =   120
      TabIndex        =   19
      Top             =   2400
      Width           =   12255
      Begin MSFlexGridLib.MSFlexGrid msf1 
         Height          =   4695
         Left            =   0
         TabIndex        =   20
         Top             =   120
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   8281
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   0
         HighLight       =   2
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ordenado por"
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   7440
      Width           =   2535
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   1200
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Basico"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Registros"
      Height          =   735
      Left            =   8760
      TabIndex        =   14
      Top             =   7440
      Width           =   1215
      Begin VB.TextBox t_encontrados 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   13
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   11655
      Begin VB.ComboBox c_marca 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7680
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "Combo1"
         Top             =   960
         Width           =   3855
      End
      Begin VB.ComboBox c_depto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7680
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Combo1"
         Top             =   600
         Width           =   3855
      End
      Begin VB.ComboBox c_grupo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7680
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "Combo1"
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox t_detalle 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   0
         Top             =   600
         Width           =   4695
      End
      Begin VB.TextBox t_codbarra 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   13
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox t_basico 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackColor       =   &H00800080&
         Caption         =   "Marca"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6480
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00800080&
         Caption         =   "Departamento"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6480
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800080&
         Caption         =   "Grupo"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6480
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800080&
         Caption         =   "Detalle prod."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800080&
         Caption         =   "Cod. Barra"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800080&
         Caption         =   "Basico"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   8235
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   26458
            MinWidth        =   26458
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_compraventa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Sub carga()
 Call armagrid
 ct = Space$(10)
 Set rs = New ADODB.Recordset
 q = "select [id_producto], a2.[descripcion] from a2"
 c = " where "
 filtro = 0
 If t_basico <> "" Then
   q = q & c & "[id_producto] = " & Val(t_basico)
   c = " and "
   filtro = 1
 End If
 
 If t_codbarra <> "" Then
   s = " = "
   q = q & c & "[cod_barra] " & s & Val(t_codbarra)
   c = " and "
   filtro = 1
 End If
 
 If t_detalle <> "" Then
   q = q & c & "a2.[descripcion] like  '%" & t_detalle & "%'"
   c = " and "
   filtro = 1
 End If
 
 If c_grupo.ListIndex > 0 Then
   q = q & c & "[id_grupo] = " & c_grupo.ItemData(c_grupo.ListIndex)
   c = " and "
   filtro = 1
 End If
 
 If c_depto.ListIndex > 0 Then
   q = q & c & "[id_departamento] = " & c_depto.ItemData(c_depto.ListIndex)
   c = " and "
   filtro = 1
 End If
 
  If c_marca.ListIndex > 0 Then
   q = q & c & "[id_marca] = " & c_marca.ItemData(c_marca.ListIndex)
   c = " and "
   filtro = 1
 End If
 
  
  
 
 
If Option5 = False Then
 If Option6 = True Then
    q = q & c & " [vigente] = true"
    c = " and "
 Else
    q = q & c & " [vigente] = false"
    c = " and "
 End If
 'filtro = 1
End If
 
 
 If Option1 = True Then
   q = q & " order by [id_producto]"
 Else
   q = q & " order by a2.[descripcion]"
 End If

If filtro = 0 Then
  J = MsgBox("Mostrar lista de precios completa? (S/N)", 4)
  If J = 6 Then
    muestra = 1
  Else
    muestra = 0
  End If
Else
 muestra = 1
End If
    

If muestra = 1 Then
espere.Show
espere.Label1 = "Geenerando Informe"
espere.Refresh

 rs.Open q, cn1
 t_encontrados = 0
 While Not rs.EOF
    b = Format$(rs("id_producto"), "00000")
    d = rs("descripcion")
    
    'busco ventas
    
    
    'busco compras
    
    
    
    msf1.AddItem b & Chr$(9) & d & Chr$(9) & p & Chr$(9) & c & Chr$(9) & m & Chr$(9) & ti & "%" & Chr$(9) & ee & Chr$(9) & F & Chr$(9) & rs("pedidos")
    t_encontrados = Val(t_encontrados) + 1
    rs.MoveNext
 Wend
 msf1.SetFocus
Set rs = Nothing
Unload espere
End If


End Sub


Private Sub btnacepta_Click()

Call carga

End Sub

Private Sub btnsale_Click()
Me.Hide
End Sub

Private Sub c_depto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub

Private Sub c_grupo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub

Private Sub c_marca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub

Private Sub c_prov_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub

Private Sub c_tipo_LostFocus()
If c_tipo.ListIndex < 0 Then
  c_tipo.ListIndex = 0
End If
End Sub


Private Sub Form_Load()
  Call carga_grupos(c_grupo)
  c_grupo.AddItem "<Todos>", 0
  c_grupo.ListIndex = 0
  Call carga_deptos_venta(c_depto)
  c_depto.AddItem "<Todos>", 0
  c_depto.ListIndex = 0
  Call carga_marcas(c_marca)
  c_marca.AddItem "<Todas>", 0
  c_marca.ListIndex = 0
  Call carga_proveedores(c_prov)
  c_prov.AddItem "<Todos>", 0
  c_prov.ListIndex = 0
  Option2 = True
  
  Call armagrid
  Check1 = 0
  Option6 = True
  Call carga_clientes(c_cli)
  c_cli.AddItem "<Todos>", 0
  c_cli.ListIndex = 0
  
End Sub
Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 9
msf1.ColWidth(0) = 800
msf1.ColWidth(1) = 4000
msf1.ColWidth(2) = 2000
msf1.ColWidth(3) = 900
msf1.ColWidth(4) = 1000
msf1.ColWidth(5) = 2000
msf1.ColWidth(6) = 900
msf1.ColWidth(7) = 1000
msf1.ColWidth(8) = 200

msf1.TextMatrix(0, 0) = "Basico"
msf1.TextMatrix(0, 1) = "Producto"
msf1.TextMatrix(0, 2) = "Cliente"
msf1.TextMatrix(0, 3) = "Cuit"
msf1.TextMatrix(0, 4) = "Venta"
msf1.TextMatrix(0, 5) = "Proveedor"
msf1.TextMatrix(0, 6) = "Cuit"
msf1.TextMatrix(0, 7) = "Compra"
msf1.TextMatrix(0, 8) = ""

For i = 0 To 1
  msf1.ColAlignment(i) = 1 'izq
Next i
For i = 2 To 6
  msf1.ColAlignment(i) = 9 'der
Next i

End Sub

  




Private Sub Form_Unload(Cancel As Integer)
Unload vta_listaprecios2
Unload vta_listaprecios3
End Sub

Sub muestra2()
 r = msf1.Row
 c = Val(msf1.TextMatrix(r, 0))
 If c > 1 Then
   Set rs = New ADODB.Recordset
   q = " select * from a2 where [id_producto] = " & c
   rs.Open q, cn1
   If Not rs.BOF And Not rs.EOF Then
     vta_listaprecios2.t_linea = r
     
     vta_listaprecios2.t_basico = rs("id_producto")
     vta_listaprecios2.t_codbarra = rs("cod_barra")
     vta_listaprecios2.t_detalle = rs("descripcion")
     vta_listaprecios2.c_grupo.ListIndex = buscaindice(vta_listaprecios2.c_grupo, rs("id_grupo"))
     vta_listaprecios2.c_depto.ListIndex = buscaindice(vta_listaprecios2.c_depto, rs("id_departamento"))
     vta_listaprecios2.c_marca.ListIndex = buscaindice(vta_listaprecios2.c_marca, rs("id_marca"))
     vta_listaprecios2.c_prov.ListIndex = buscaindice(vta_listaprecios2.c_prov, rs("id_proveedor"))
     vta_listaprecios2.c_unidad.ListIndex = buscaindice(vta_listaprecios2!c_unidad, rs("id_unidad"))
     vta_listaprecios2.t_envase = rs("envase")
     vta_listaprecios2.t_pu = rs("pu")
     vta_listaprecios2.c_iva.ListIndex = buscaindice(vta_listaprecios2!c_iva, rs("cod_tasaiva"))
     vta_listaprecios2.t_stockminimo = rs("stock_minimo")
     vta_listaprecios2.t_utilidad = rs("porc_utilidad")
     vta_listaprecios2.t_costo = rs("costoreal")
     vta_listaprecios2.t_fletecompra = rs("flete_compra")
     vta_listaprecios2.t_dtocompra = rs("dto_compra")
     vta_listaprecios2.t_final = rs("precio_final")
     vta_listaprecios2.t_tasaimpint = rs("tasa_imp_interno")
     vta_listaprecios2.t_tipo = rs("tipo_producto")
     vta_listaprecios2.t_moneda = rs("moneda")
     vta_listaprecios2.t_impuesto = rs("impuesto")
     vta_listaprecios2.t_observaciones = rs("observaciones")
     vta_listaprecios2.t_preciocompra = rs("precio_ult_compra")
     vta_listaprecios2.t_ultvta = rs("ultima_venta")
     vta_listaprecios2.t_ultimacompra = rs("ultima_compra")
     vta_listaprecios2.t_fechaactu = rs("fecha_actu_precio_venta")
     vta_listaprecios2.t_stock = rs("stock")
     vta_listaprecios2.t_oc = rs("pedidos")
     vta_listaprecios2.t_pedidos = rs("requeridos")
     vta_listaprecios2.t_fechaactuc = rs("fecha_ult_compra")
     vta_listaprecios2.t_textocentral = rs("texto_central")
     vta_listaprecios2.t_tipocarga = rs("tipo_carga_tique")
     If rs("vigente") = True Then
      vta_listaprecios2.Check1 = 1
     Else
      vta_listaprecios2.Check1 = 0
     End If
     
     vta_listaprecios2.Show
   End If
   Set rs = Nothing
   
 End If
 
End Sub


Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(1) = "[F1] P.F - [F2] Sel. - [F3] A Faltantes -  [F4] Saca - [F5] Grupal - [F6] Op.  - [F7] Imprime - [F10] Imp. Etiq. - [F11] Marca Etiq. - [ENTER] Detalle  - [Esc] Cancela"

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)



If KeyCode = vbKeyF4 Then
  r = msf1.Row
  p = Val(msf1.TextMatrix(r, 0))
  If p > 1 Then
    msf1.RemoveItem r
    t_encontrados = Val(t_encontrados) - 1
  End If
End If


If KeyCode = vbKeyF7 Then
  Dim c(15) As Double
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    c(0) = 0
    c(1) = 1
    c(2) = 2
      
    For i = 3 To 14
      c(i) = -1
    Next i
    
    If t_detalle <> "" Then
      t = "Detalle: " & t_detalle
    Else
      t = ""
    End If
    
    If c_grupo.ListIndex > 0 Then
       t1 = "Grupo: " & c_grupo
    End If
    Call imprimegrid(msf1, c(), "LISTA DE PRECIOS", "", t, t1, 80, 8, True, False, "V")
  End If
End If





End Sub
Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call muestra2
Else
 If KeyAscii <> 27 Then
  vta_listaprecios5.t_texto = Chr$(KeyAscii)
  vta_listaprecios5.Show
 End If
End If


End Sub

Private Sub msf1_LostFocus()
'Call barra(Me)
End Sub

Private Sub t_basico_GotFocus()
t_basico = ""
End Sub

Private Sub t_basico_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub

Private Sub t_codbarra_GotFocus()
t_codbarra = ""
End Sub

Private Sub t_codbarra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub

Private Sub T_detalle_GotFocus()
t_detalle = ""
End Sub

Private Sub t_detalle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub

