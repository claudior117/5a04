VERSION 5.00
Begin VB.Form gen_parametros 
   Caption         =   "Definir Cuentas"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12330
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   12330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Define Cuentas Contables"
      Height          =   375
      Left            =   6240
      TabIndex        =   25
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Frame Frame8 
      Caption         =   "IMPORTANTE"
      Height          =   615
      Left            =   360
      TabIndex        =   22
      Top             =   8040
      Width           =   8535
      Begin VB.Label Label23 
         Caption         =   "Si realiza cambios en la configuración ,  deberá salir del Sistema y volver a Ingresar para un correcto funcionamiento. "
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   8295
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Parametros de COMPRAS"
      Height          =   1335
      Left            =   360
      TabIndex        =   17
      Top             =   1320
      Width           =   5175
      Begin VB.TextBox t_impintgo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2640
         MaxLength       =   14
         TabIndex        =   28
         Top             =   960
         Width           =   1215
      End
      Begin VB.ComboBox c_tipopuactu 
         Height          =   315
         ItemData        =   "cyb002.frx":0000
         Left            =   1440
         List            =   "cyb002.frx":000A
         TabIndex        =   24
         Top             =   600
         Width           =   3135
      End
      Begin VB.ComboBox c_tipopreciocompra 
         Height          =   315
         ItemData        =   "cyb002.frx":0059
         Left            =   1440
         List            =   "cyb002.frx":0066
         TabIndex        =   18
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label29 
         BackColor       =   &H00C00000&
         Caption         =   "Impuesto Int. Gasoil $/lt"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         Caption         =   "Actualizacion Precios Compra y Est. Costos"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Parametros de VENTAS"
      Height          =   5175
      Left            =   360
      TabIndex        =   8
      Top             =   2760
      Width           =   11175
      Begin VB.TextBox t_webservice 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   7680
         MaxLength       =   1
         TabIndex        =   61
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox t_muestrasaldo 
         Height          =   285
         Left            =   7680
         MaxLength       =   1
         TabIndex        =   57
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox t_ncenrbo 
         Height          =   285
         Left            =   7680
         MaxLength       =   1
         TabIndex        =   54
         Top             =   720
         Width           =   495
      End
      Begin VB.ComboBox c_redondeo 
         Height          =   315
         ItemData        =   "cyb002.frx":00C1
         Left            =   1920
         List            =   "cyb002.frx":00D4
         TabIndex        =   52
         Top             =   4680
         Width           =   2535
      End
      Begin VB.CheckBox Check4 
         Height          =   315
         Left            =   3240
         TabIndex        =   50
         Top             =   4320
         Width           =   495
      End
      Begin VB.TextBox t_pie2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   49
         Top             =   3960
         Width           =   3135
      End
      Begin VB.TextBox t_pie1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   47
         ToolTipText     =   "Ingrese el texto a impirir o * (asterisco) sino desea imprimir"
         Top             =   3600
         Width           =   3135
      End
      Begin VB.TextBox t_tasafinanciera 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7680
         MaxLength       =   14
         TabIndex        =   44
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox Check3 
         Height          =   315
         Left            =   3240
         TabIndex        =   42
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox t_numfact 
         Height          =   285
         Left            =   1920
         MaxLength       =   1
         TabIndex        =   37
         Top             =   1920
         Width           =   495
      End
      Begin VB.ComboBox c_tipoprecio 
         Height          =   315
         ItemData        =   "cyb002.frx":0106
         Left            =   1920
         List            =   "cyb002.frx":0110
         TabIndex        =   36
         Top             =   2400
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Cambia a todos los usuarios"
         Height          =   675
         Left            =   3840
         TabIndex        =   35
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox t_recargocc 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1920
         MaxLength       =   14
         TabIndex        =   34
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox t_tasaiva 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         MaxLength       =   14
         TabIndex        =   15
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox t_dto2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         MaxLength       =   14
         TabIndex        =   13
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox t_dto1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         MaxLength       =   14
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox t_minretib 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         MaxLength       =   14
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "Versión WEBSERVICE ARCA"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   6360
         TabIndex        =   60
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C00000&
         Caption         =   "Calcula Saldo en Pantalla Factura"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   6360
         TabIndex        =   59
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "[S] Si   [N] No"
         Height          =   375
         Left            =   8280
         TabIndex        =   58
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "[S] Si   [N] No"
         Height          =   375
         Left            =   8280
         TabIndex        =   56
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Emite Nc en Recibos"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   6360
         TabIndex        =   55
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label38 
         BackColor       =   &H00C00000&
         Caption         =   "                                Tipo Redondeo"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   53
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label37 
         BackColor       =   &H00C00000&
         Caption         =   "Utiliza precio lista en Facturas por Rtos"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   4320
         Width           =   2895
      End
      Begin VB.Label Label36 
         BackColor       =   &H00C00000&
         Caption         =   "Texto 2 pie resumen"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label Label35 
         BackColor       =   &H00C00000&
         Caption         =   "Texto 1 pie resumen"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label34 
         BackColor       =   &H00C00000&
         Caption         =   "Tasa Financiera (mensual)"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   6360
         TabIndex        =   45
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label33 
         BackColor       =   &H00C00000&
         Caption         =   "Impide Facturar Exceso Limite Credito"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   3240
         Width           =   2895
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C00000&
         Caption         =   "Numeracion de Fact. y Nc correlativas:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   41
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "[S] Si   [N] No (Numeracion Independiente)"
         Height          =   375
         Left            =   2520
         TabIndex        =   40
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C00000&
         Caption         =   "Precio Utilizado para Facturar"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   39
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label32 
         BackColor       =   &H00C00000&
         Caption         =   "% Recargo Cta.Cte."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C00000&
         Caption         =   "Tasa General Iva"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C00000&
         Caption         =   "Descuento Rapido 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C00000&
         Caption         =   "Descuento Rapido 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C00000&
         Caption         =   "Importe MInimo para Retener IB"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos de Trabajo"
      Height          =   1095
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   11175
      Begin VB.CheckBox Check2 
         Height          =   195
         Left            =   9360
         TabIndex        =   33
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox t_fechacorte 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5400
         MaxLength       =   14
         TabIndex        =   30
         ToolTipText     =   "El sistema impedira que se ingresen movimientos anteriores a la fecha de corte"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox t_puntovpf 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   9360
         MaxLength       =   14
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox t_puntovf 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5400
         MaxLength       =   14
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox t_puntov 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1920
         MaxLength       =   14
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox t_cotizacion 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1920
         MaxLength       =   14
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label31 
         BackColor       =   &H00C00000&
         Caption         =   "Genera Asientos Automat."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7200
         TabIndex        =   32
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label30 
         BackColor       =   &H00C00000&
         Caption         =   "Fecha Corte General"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3600
         TabIndex        =   31
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label28 
         BackColor       =   &H00C00000&
         Caption         =   "Punto Venta Prueba Fiscal"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7200
         TabIndex        =   27
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C00000&
         Caption         =   "Punto Venta Fiscal"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3600
         TabIndex        =   21
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "Punto Venta Manual"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Cotizacion Dolar"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   9120
      TabIndex        =   0
      Top             =   8040
      Width           =   2655
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "gen_parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private Sub c_acreedores_LostFocus()
If c_acreedores.ListIndex < 0 Then
     c_acreedores.ListIndex = 0
  End If
End Sub

Private Sub c_caja_LostFocus()
If c_caja.ListIndex < 0 Then
     c_caja.ListIndex = 0
  End If

End Sub

Private Sub c_costo_LostFocus()
If c_costo.ListIndex < 0 Then
     c_costo.ListIndex = 0
  End If
End Sub

Private Sub c_deudores_LostFocus()
If c_deudores.ListIndex < 0 Then
     c_deudores.ListIndex = 0
  End If

End Sub

Private Sub c_inventario_Change()
If c_inventario.ListIndex < 0 Then
     c_inventario.ListIndex = 0
End If
End Sub

Private Sub c_redondeo_LostFocus()
If c_redondeo.ListIndex < 0 Then
  c_redondeo.ListIndex = 0
End If
End Sub

Private Sub c_tipoprecio_LostFocus()
If c_tipoprecio.ListIndex < 0 Then
  c_tipoprecio.ListIndex = 0
End If
End Sub

Private Sub c_tipopreciocompra_LostFocus()
If c_tipopreciocompra.ListIndex < 0 Then
  c_tipopreciocompra.ListIndex = 0
End If
End Sub

Private Sub c_tipopuactu_LostFocus()
If c_tipopuactu.ListIndex < 0 Then
  c_tipoactupu.ListIndex = 0
End If
End Sub

Private Sub c_ventas_LostFocus()
If c_ventas.ListIndex < 0 Then
     c_ventas.ListIndex = 0
  End If

End Sub

Private Sub Command1_Click()
J = MsgBox("Confirma Actualizar Parametros generales del sistema", 4)
If J = 6 Then
  Call graba
End If

End Sub
Sub graba()

Set rs = New ADODB.Recordset
q = "select * from g0 where [sucursal] = 0"
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
If Not rs.EOF And Not rs.BOF Then
  If Val(t_cotizacion) <= 0 Then
    t_cotizacion = "1.00"
  End If
  rs("id_cuenta_deudores") = gen_parametros1.c_deudores.ItemData(gen_parametros1.c_deudores.ListIndex)
  rs("id_cuenta_acreedores") = gen_parametros1.c_acreedores.ItemData(gen_parametros1.c_acreedores.ListIndex)
  rs("id_cuenta_ventas") = gen_parametros1.c_ventas.ItemData(gen_parametros1.c_ventas.ListIndex)
  rs("cotizacion") = Val(t_cotizacion)
  rs("sucursal_actual") = Val(t_puntov)
  rs("id_cuenta_iva_compras") = gen_parametros1.c_ivac.ItemData(gen_parametros1.c_ivac.ListIndex)
  rs("id_cuenta_ret_gan") = gen_parametros1.c_retgan.ItemData(gen_parametros1.c_retgan.ListIndex)
  rs("id_cuenta_ret_ib") = gen_parametros1.c_retib.ItemData(gen_parametros1.c_retib.ListIndex)
  rs("id_cuenta_nograbados") = gen_parametros1.c_nograb.ItemData(gen_parametros1.c_nograb.ListIndex)
  rs("id_cuenta_iva_ventas") = gen_parametros1.c_ivav.ItemData(gen_parametros1.c_ivav.ListIndex)
  rs("id_cuenta_perc_ib") = gen_parametros1.C_percib.ItemData(gen_parametros1.C_percib.ListIndex)
  rs("id_cuenta_perc_iva") = gen_parametros1.c_perciva.ItemData(gen_parametros1.c_perciva.ListIndex)
  rs("id_cuenta_compras_varias") = gen_parametros1.c_compvarias.ItemData(gen_parametros1.c_compvarias.ListIndex)
  rs("id_cuenta_inventario") = gen_parametros1.c_inventario.ItemData(gen_parametros1.c_inventario.ListIndex)
  rs("id_cuenta_costo_merc") = gen_parametros1.c_costo.ItemData(gen_parametros1.c_costo.ListIndex)
  rs("id_cuenta_retibba_ventas") = gen_parametros1.c_retibbav.ItemData(gen_parametros1.c_retibbav.ListIndex)
  rs("id_cuenta_retiva_ventas") = gen_parametros1.c_retivav.ItemData(gen_parametros1.c_retivav.ListIndex)
  rs("id_cuenta_retgan_ventas") = gen_parametros1.c_retganv.ItemData(gen_parametros1.c_retganv.ListIndex)
  rs("id_cuenta_retsuss_ventas") = gen_parametros1.c_retsussv.ItemData(gen_parametros1.c_retsussv.ListIndex)
  rs("numeracion_comun_Fact_nc") = t_numfact
  rs("minimo_retib") = Val(t_minretib)
  rs("descuento1") = Val(t_dto1)
  rs("descuento2") = Val(t_dto2)
  rs("tasa_general_iva") = Val(t_tasaiva)
  rs("tipo_actu_precio_comp_compra") = c_tipopreciocompra.ListIndex + 1
  rs("tipo_precio_venta") = c_tipoprecio.ListIndex
  rs("id_cuenta_resultado_tenencia") = gen_parametros1.c_resultado.ItemData(gen_parametros1.c_resultado.ListIndex)
  rs("sucursal_prueba") = Val(t_puntovpf)
  rs("impuesto_int_gasoil") = Format(Val(t_impintgo), "###0.0000")
  rs("fecha_corte") = Format$(t_fechacorte, "dd/mm/yyyy")
  rs("recargo_cc") = Val(t_recargocc)
  rs("tasa_financiera") = Val(t_tasafinanciera)
  rs("texto_resumen1") = t_pie1
  rs("texto_resumen2") = t_pie2
  rs("precio_remito_factura") = Check4
  rs("tipo_redondeo") = c_redondeo.ListIndex
  rs("nc_en_recibo") = t_ncenrbo
  If Val(t_webservice) >= 3 And Val(t_webservice) <= 4 Then
     rs("version_webservice") = Val(t_webservice)
     para.version_webservice = Val(t_webservice)
  Else
     MsgBox ("La versión WebService solo puede ser 3 o 4. No se actualizará, verifíquelo!!!")
  End If
  'rs ("muestra_saldo_fact_venta")
  rs("muestra_saldo_fact_venta") = t_muestrasaldo
  
  
  If c_tipopuactu.ListIndex = 0 Then
    tpu = "C"
  Else
    tpu = "S"
  End If
  rs("tipo_actu_pu_compra") = tpu
  rs("graba_asientos_auto") = Check2
  rs("tipo_control_limite_credito") = Check3
  
  rs.Update
  
  Set rs1 = New ADODB.Recordset
  q = "select * from cyb_01 where [id_forma_pago] = 1"
  rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
  If Not rs1.EOF And Not rs1.BOF Then
    rs1("id_cuenta_cont") = gen_parametros1.c_caja.ItemData(gen_parametros1.c_caja.ListIndex)
    rs1.Update
    para.cuenta_caja = gen_parametros1.c_caja.ItemData(gen_parametros1.c_caja.ListIndex)
  
  End If
  Set rs1 = Nothing
  
  para.cuenta_acreedores = gen_parametros1.c_acreedores.ItemData(gen_parametros1.c_acreedores.ListIndex)
  para.cuenta_deudores = gen_parametros1.c_deudores.ItemData(gen_parametros1.c_deudores.ListIndex)
  para.cuenta_ventas = gen_parametros1.c_ventas.ItemData(gen_parametros1.c_ventas.ListIndex)
  para.cotizacion = Val(t_cotizacion)
  para.cuenta_iva_compras = gen_parametros1.c_ivac.ItemData(gen_parametros1.c_ivac.ListIndex)
  para.cuenta_retgan = gen_parametros1.c_retgan.ItemData(gen_parametros1.c_retgan.ListIndex)
  para.cuenta_retib = gen_parametros1.c_retib.ItemData(gen_parametros1.c_retib.ListIndex)
  para.cuenta_conceptos_nograbados = gen_parametros1.c_nograb.ItemData(gen_parametros1.c_nograb.ListIndex)
  para.cuenta_iva_ventas = gen_parametros1.c_ivav.ItemData(gen_parametros1.c_ivav.ListIndex)
  para.cuenta_perc_IB = gen_parametros1.C_percib.ItemData(gen_parametros1.C_percib.ListIndex)
  para.cuenta_compras_varias = gen_parametros1.c_compvarias.ItemData(gen_parametros1.c_compvarias.ListIndex)
  para.cuenta_inventario = gen_parametros1.c_inventario.ItemData(gen_parametros1.c_inventario.ListIndex)
  para.cuenta_costo = gen_parametros1.c_costo.ItemData(gen_parametros1.c_costo.ListIndex)
  para.cuenta_retibbav = gen_parametros1.c_retibbav.ItemData(gen_parametros1.c_retibbav.ListIndex)
  para.cuenta_retivav = gen_parametros1.c_retivav.ItemData(gen_parametros1.c_retivav.ListIndex)
  para.cuenta_retganv = gen_parametros1.c_retganv.ItemData(gen_parametros1.c_retganv.ListIndex)
  para.cuenta_retsussv = gen_parametros1.c_retsussv.ItemData(gen_parametros1.c_retsussv.ListIndex)
  para.numeracion_comun_Fact_nc = t_numfact
  'glo.sucursal = Val(t_puntov)
  glo.sucursalf = Val(t_puntovf)
  para.minimo_retib = Val(t_minretib)
  para.tasageneral = Val(t_tasaiva)
  para.tipoactupreciocompcompra = c_tipopreciocompra.ListIndex + 1
  para.tipoprecioventa = c_tipoprecio.ListIndex
  para.ncenrecibo = t_ncenrbo
  para.muestrasaldofactventa = t_muestrasaldo
  If Check1 = 1 Then
    q = "select * from g1"
    Set rs1 = New ADODB.Recordset
    rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
    While Not rs1.EOF
      rs1("tipo_precio_venta") = c_tipoprecio.ListIndex
      rs1.Update
      rs1.MoveNext
    Wend
    Set rs1 = Nothing
  End If
  Unload Me
Else
  MsgBox ("Error. El sistema no fue actualizado!")
End If
Set rs = Nothing

End Sub
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
gen_parametros1.Show
End Sub

Private Sub Form_Load()
Load gen_parametros1


t_cotizacion = Format$(para.cotizacion, "######0.00")
t_puntov = Format$(glo.sucursal, "0000")
t_puntovf = Format$(glo.sucursalf, "0000")
t_numfact = para.numeracion_comun_Fact_nc
t_minretib = para.minimo_retib
t_tasaiva = para.tasageneral
t_fechacorte = para.fechacorte
Set rs = New ADODB.Recordset
q = "select * from g0 where [sucursal] = 0"
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
  t_impintgo = Format$(rs("impuesto_int_gasoil"), "###0.0000")
  t_dto1 = Format$(rs("descuento1"), "####0.00")
  t_dto2 = Format$(rs("descuento2"), "####0.00")
  c_tipopreciocompra.ListIndex = rs("tipo_actu_precio_comp_compra") - 1
  c_tipoprecio.ListIndex = rs("tipo_precio_venta")
  If rs("tipo_actu_pu_compra") = "C" Then
    c_tipopuactu.ListIndex = 0
  Else
     c_tipopuactu.ListIndex = 1
  End If
  
  t_puntovpf = Format$(rs("sucursal_prueba"), "0000")
  If rs("graba_asientos_auto") Then
    Check2 = 1
  Else
    Check2 = 0
  End If
  t_recargocc = Format$(rs("recargo_cc"), "####0.00")
  If rs("tipo_control_limite_credito") Then
    Check3 = 1
  Else
   Check3 = 0
  End If
  t_tasafinanciera = Format$(rs("tasa_financiera"), "##0.00")
  t_pie1 = rs("texto_resumen1")
  t_pie2 = rs("texto_resumen2")
  Check4 = rs("precio_remito_factura")
  c_redondeo.ListIndex = rs("tipo_redondeo")
  gen_parametros1.c_resultado.ListIndex = buscaindice(gen_parametros1.c_resultado, rs("id_cuenta_resultado_tenencia"))
 t_ncenrbo = rs("nc_en_recibo")
 t_muestrasaldo = rs("muestra_saldo_fact_venta")
 t_webservice = rs("version_webservice")
 
End If
Set rs = Nothing
End Sub

Private Sub m_muestrasaldo_LostFocus()
t_muestrasaldo = Format$(t_muestrasaldo, ">@")
If t_muestrasaldo <> "N" And t_muestrasaldo <> "S" Then
  t_muestrasaldo = "N"
End If

End Sub

Private Sub t_fechacorte_LostFocus()
Call solofecha(t_fechacorte)
End Sub

Private Sub t_minretib_LostFocus()
If Val(t_minretib) < 0 Then
   t_minretib = "0.00"
End If
End Sub

Private Sub t_ncenrbo_Change()
t_ncenrbo = Format$(t_ncenrbo, ">@")
If t_ncenrbo <> "N" And t_ncenrbo <> "S" Then
  t_ncenrbo = "N"
End If
End Sub

Private Sub t_numfact_LostFocus()
t_numfact = Format$(t_numfact, ">@")
If t_numfact <> "S" And t_numfact <> "N" Then
   t_numfact = "S"
End If
End Sub

Private Sub t_pie1_LostFocus()
If Len(t_pie1) < 5 Then
  t_pie1 = "*"
End If
End Sub

Private Sub t_pie2_LostFocus()
If Len(t_pie2) < 5 Then
  t_pie2 = "*"
End If

End Sub

Private Sub t_puntov_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)
End Sub

Private Sub t_puntov_LostFocus()
If Val(t_puntov) <= 0 Then
  t_puntov = Format$(glo.sucursal, "0000")
End If
End Sub

Private Sub t_tipopuactu_LostFocus()

End Sub

Private Sub t_puntovpf_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)
End Sub

Private Sub t_recargocc_LostFocus()
If Not IsNumeric(t_recargocc) Then
  t_recargocc = "0.00"
Else
  t_recargocc = Format$(Val(t_recargocc), "####0.00")
End If
End Sub
