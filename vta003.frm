VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_facturacion1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INGRESO DE ARTICULOS"
   ClientHeight    =   2175
   ClientLeft      =   135
   ClientTop       =   4815
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2175
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   1695
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   11775
      Begin VB.TextBox t_tasaib 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   10920
         MaxLength       =   8
         TabIndex        =   20
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_costo 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   9960
         MaxLength       =   10
         TabIndex        =   19
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_unidad 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   7440
         MaxLength       =   5
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   120
         Picture         =   "vta003.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox t_ip 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   8880
         MaxLength       =   5
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_detalle 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   1680
         MaxLength       =   49
         TabIndex        =   1
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox t_basico 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   720
         MaxLength       =   20
         TabIndex        =   0
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox t_importe 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   10680
         MaxLength       =   11
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox c_tasa 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   9360
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox t_pu 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   8280
         MaxLength       =   10
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox t_cantidad 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   6480
         MaxLength       =   8
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox t_renglon 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   7680
         MaxLength       =   8
         TabIndex        =   9
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Unidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7320
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Basico"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   720
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Importe"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   10560
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Tasa Iva"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9240
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Pu"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8280
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Cantidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6360
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Detalle"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   4815
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1920
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   12347
            MinWidth        =   12347
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "08/11/2023"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:27 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_facturacion1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim grecargocc As Single


Private Sub c_tasa_LostFocus()
If c_tasa.ListIndex < 0 Then
  c_tasa.ListIndex = 0
End If
End Sub

Private Sub Command1_Click()
ABM_PROD.Show
End Sub

Private Sub Form_Activate()

't_basico.SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyUp
     Call tabup(Me)
   
     
         
End Select
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 6)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)

For i = 0 To 9
  c_tasa.AddItem para.tasaiva(i)
Next i
c_tasa.ListIndex = 0

Set rs = New ADODB.Recordset
q = "select [recargo_cc] from g0 where [sucursal] = 0"
rs.Open q, cn1
If Not rs.BOF And Not rs.EOF Then
  grecargocc = rs("recargo_cc")
Else
  grecargocc = 0
End If
Set rs = Nothing
End Sub


Private Sub t_basico_GotFocus()
Call nivel_acceso(1)


Me.StatusBar1.Panels.item(2) = "[ENTER] Acepta - [ESC] Sale - [F6] Dto1 - [F7] Dto2 - [F8]Lista Precios  "

t_detalle.Enabled = False
If para.producto_sel > 0 Then
  t_basico = para.producto_sel
End If
End Sub

Private Sub t_basico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF8 Then
  vta_listaprecios.Show
End If

If KeyCode = vbKeyF6 And t_renglon = "" Then
  Set rs = New ADODB.Recordset
  q = "select * from g0 where [sucursal] = 0"
  rs.Open q, cn1
  d1 = rs("descuento1")
  Set rs = Nothing
  t_basico = 1
  t_detalle = "Descuento " & Format$(d1, "##0.00") & "%"
  t_pu = "0.00"
  c_tasa.ListIndex = 1
  t_ip = 1
  t_cantidad = "1.00"
  t_pu = Format$(-Val(vta_facturacion.t_subtotal) * d1 / 100, "######0.00")
  t_importe = t_pu
  
End If

If KeyCode = vbKeyF7 And t_renglon = "" Then
  Set rs = New ADODB.Recordset
  q = "select * from g0 where [sucursal] = 0"
  rs.Open q, cn1
  d2 = rs("descuento2")
  Set rs = Nothing
  t_basico = 1
  t_detalle = "Descuento " & Format$(d2, "##0.00") & "%"
  t_pu = "0.00"
  c_tasa.ListIndex = 1
  t_ip = 1
  t_cantidad = "1.00"
  t_pu = Format$(-Val(vta_facturacion.t_subtotal) * d2 / 100, "######0.00")
  t_importe = t_pu

End If

End Sub

Private Sub t_basico_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call carga
End If
End Sub

Sub carga()
If t_basico = "" Then
  t_basico = 1
End If
If IsNumeric(t_basico) Then

 If Val(t_basico) <= 1 Then
    c_tasa.Locked = False
    t_basico = 1
    t_ip = 1
    t_tasaib = para.tasaib
    t_detalle.Enabled = True
    c_tasa.ListIndex = buscaindice2(c_tasa, para.tasageneral)
    t_detalle.SetFocus
    t_pu.Locked = False
  Else
    'c_tasa.Locked = True
    If Len(t_basico) <= 5 Then
       Call busca("I") 'busca por id. producto
    Else
       Call busca("B") 'busca por cod. barra
    End If
    
     If para.id_grupo_modulo_actual >= 7 Then
         t_pu.Locked = False
     Else
         t_pu.Locked = True
     End If
 End If
Else
   
  Call busca("B")
End If
End Sub
Sub busca(tipo As String)
'tipo = I por id_producto tipo = B por cod_barra
Set rs = New ADODB.Recordset
q = "select id_producto, cod_barra, descripcion, precio_final, pu, moneda, costoreal, cod_tasaiva, unidad, tasaib  from a2, g5, g12 where a2.[id_unidad] = g5.[id_unidad] and a2.[id_tasaib] = g12.[id_tasaib] "
If tipo = "I" Then
  q = q & " and [id_producto] = " & Val(t_basico)
Else
  q = q & " and [cod_barra] = '" & RTrim$(t_basico) & "'"
End If
rs.MaxRecords = 1
rs.Open q, cn1
precio21 = 0

If Not rs.BOF And Not rs.EOF Then
  t_detalle = rs("descripcion")
  
  
  If para.tipoprecioventa = 1 Then
    precio21 = rs("precio_final")
  Else
    precio21 = rs("pu")
  End If
  If precio21 = 0 Then
    t_pu = ""
  Else
  
   If vta_facturacion.Option4 = True Then
     If rs("moneda") = "P" Then
        'factura en $ y producto en $
        t_pu = precio21
        t_costo = rs("costoreal")
        
     Else
        'factura en $ y producto en U$s
        t_pu = precio21 * Val(vta_facturacion.t_cotizacion)
        t_costo = rs("costoreal") * Val(vta_facturacion.t_cotizacion)
     End If
  Else
        If rs("moneda") = "D" Then
        'factura en U$s y producto en U$s
           t_pu = precio21
           t_costo = rs("costoreal")
     Else
        'factura en u$s y producto en $
        t_pu = precio21 / Val(vta_facturacion.t_cotizacion)
        t_costo = rs("costoreal") / Val(vta_facturacion.t_cotizacion)
     End If
  End If
   
   If vta_facturacion.Option1 = True And grecargocc > 0 Then
       r = (Val(t_pu) * grecargocc) / 100
       t_pu = Format(Val(t_pu) + r, "#####0.00")
   End If
  End If
  
  t_pu = Format$(Val(t_pu), "#####0.00")
  c_tasa.ListIndex = rs("cod_tasaiva")
  t_ip = rs("id_producto")
  t_unidad = rs("unidad")
  t_tasaib = rs("tasaib")
  
Else
  MsgBox ("Producto no Ingresado")
  t_basico.SetFocus
  t_costo = 0
End If
Set rs = Nothing
End Sub

Private Sub t_basico_LostFocus()
Call barraesag(Me)
End Sub

Private Sub t_cantidad_KeyPress(KeyAscii As Integer)
   Call solonum(KeyAscii, 1)
End Sub

Sub cargarenglon(t As String)
  
  ip = Val(t_ip)
  d = t_detalle
  cu = Format$(Val(t_cantidad), "######0.00")
  ti = Format$(c_tasa, "####0.00")
  u = RTrim$(t_unidad)
  'If vta_facturacion.t_letra = "A" Then
   'se cargan todos los comprobantes sin iva, en las que no se descrimina solo se imprimen con iva
  
  If para.tipo_iva_empresa = 1 Then 'inscriptos
        If para.tipoprecioventa = 1 Then
         pu = Format$(Val(t_pu) / (1 + Val(c_tasa) / 100), "#####0.0000")
         im = Format$(Val(pu) * Val(cu), "#####0.00")
         puf = Format$(Val(t_pu), "#####0.0000")
        Else
         pu = Format$(Val(t_pu), "#####0.0000")
         im = Format$(Val(pu) * Val(cu), "#####0.00")
         puf = Format$(Val(t_pu) * (1 + Val(c_tasa) / 100), "#####0.0000")
        End If
  Else ' monotributos
        pu = Format$(Val(t_pu), "#####0.0000")
        im = Format$(Val(pu) * Val(cu), "#####0.00")
       puf = Format$(Val(t_pu), "#####0.0000")
  End If
  'End If
  If u = "" Then
    u = " "
  End If
  cr = Format$(Val(t_costo) * Val(cu), "#####0.00")
  If t = "A" Then
    'nueva linea
    r = vta_facturacion.msf1.Rows
    If r <= Val(vta_facturacion.t_cantlineas) Then
       vta_facturacion.msf1.AddItem r & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & cu & Chr(9) & u & Chr$(9) & pu & Chr(9) & ti & Chr(9) & im & Chr(9) & puf & Chr(9) & Chr(9) & cr & Chr(9) & Format$(Val(t_tasaib), "####0.00")
    Else
       MsgBox ("Se ha superado el limite maximo de renglones para este comprobante")
    End If
  
  
  Else
    r = t_renglon
    vta_facturacion.msf1.AddItem r & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & cu & Chr$(9) & u & Chr$(9) & pu & Chr(9) & ti & Chr(9) & im & Chr(9) & puf & Chr(9) & Chr(9) & cr & Chr(9) & Format$(Val(t_tasaib), "####0.00"), r
    vta_facturacion.msf1.RemoveItem r + 1
  End If
   
 ' s = 0
 ' v = 0
 ' For i = 1 To vta_facturacion.msf1.Rows - 1
 '   If r > 0 Then
 '     k = Val(vta_facturacion.msf1.TextMatrix(i, 7))
 '     s = s + k
 '     v = v + (k * Val(vta_facturacion.msf1.TextMatrix(i, 6)) / 100)
 '   End If
 ' Next i
 'vta_facturacion.t_subtotal = s
 ' vta_facturacion.t_iva = v
  
  vta_facturacion.CALCULATOTALES
  vta_facturacion.sacatotales
  vta_facturacion.sacaperc
  vta_facturacion.sacatotales
  para.producto_sel = 0
    
  If vta_facturacion.c_tipocomp.ListIndex = 0 And vta_facturacion.Option1 Then
    If vta_facturacion.Option4 Then
     'pesos
     tpl = Val(vta_facturacion.t_total)
    Else
     tpl = Val(vta_facturacion.T_total2)
    End If
     
    If Val(vta_clientes.t_saldo1) + tpl > Val(vta_clientes.t_limite) Then
      MsgBox ("El comprobante actual ha superado el LIMITE de CREDITO establecido para el cliente")
    
    End If
  End If
     



End Sub
 
  
Sub limpia()
t_cantidad = ""
t_basico = ""
t_detalle = ""
t_pu = ""
t_importe = ""
t_ip = ""
t_unidad = ""
End Sub

Private Sub T_detalle_GotFocus()
Me.StatusBar1.Panels.item(2) = "[ENTER] Acepta - [ESC] Sale - [F3] Descripcion extra   "
End Sub

Private Sub t_detalle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
  Form1.Show
End If
End Sub

Private Sub t_detalle_LostFocus()
Me.StatusBar1.Panels.item(2) = "[ENTER] Acepta - [ESC] Sale "
End Sub

Private Sub t_importe_GotFocus()
t_importe = Format$(Val(t_cantidad) * Val(t_pu), "#####0.00")
End Sub

Private Sub t_importe_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If t_renglon = "" Then
   Call cargarenglon("A")
   t_basico.SetFocus
   
  Else
   Call cargarenglon("M")
   t_basico.SetFocus
   Me.Hide
  End If
  Call limpia
  
Else
  Call solonum(KeyAscii, 1)
End If
End Sub

Private Sub t_pu_GotFocus()
Me.StatusBar1.Panels.item(2) = "[ENTER] Acepta - [F6] Dto % - [F7] Dto $ - "
End Sub

Private Sub t_pu_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF6 Then
  d = InputBox("Ingrese % descuento", "Descuento")
  If Val(d) > 0 Then
     pd = Format(Val(t_pu) * Val(d) / 100, "######0.00")
     t_pu = Val(t_pu) - pd
  End If
End If
  



If KeyCode = vbKeyF7 Then
  d = InputBox("Ingrese descuento en pesos", "Descuento $")
  If Val(d) > 0 Then
     pd = Format(Val(d), "######0.00")
     t_pu = Val(t_pu) - pd
  End If
End If
End Sub
