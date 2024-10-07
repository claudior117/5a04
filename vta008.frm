VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form vta_cc_detalle 
   BackColor       =   &H00E0E0E0&
   Caption         =   "COMPROBANTE DE VENTA(DETALLE)"
   ClientHeight    =   8760
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12075
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   12075
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Archivo Electronico"
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   6240
      Width           =   11775
      Begin VB.TextBox t_path 
         Height          =   405
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   9975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ver o Modificar"
         Height          =   375
         Left            =   10560
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4785
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   11895
   End
   Begin VB.Frame CUIT 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   1215
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11655
      Begin VB.TextBox t_numint 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   10200
         MaxLength       =   10
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox t_tipocomp 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   9000
         MaxLength       =   6
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox t_numcomp 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6960
         MaxLength       =   8
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox t_letra 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5520
         MaxLength       =   6
         TabIndex        =   3
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox t_sucursal 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6120
         MaxLength       =   6
         TabIndex        =   4
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox t_prov 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         MaxLength       =   10
         TabIndex        =   2
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox t_idprov 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800080&
         Caption         =   "Numero Interno"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   10200
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800080&
         Caption         =   "Tipo"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9000
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800080&
         Caption         =   "Comprobante:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5520
         TabIndex        =   11
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800080&
         Caption         =   "Cliente"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10200
      TabIndex        =   6
      Top             =   7440
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta008.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "vta008.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   7
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
      TabIndex        =   5
      Top             =   8505
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   21167
            MinWidth        =   21167
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_cc_detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim l1 As String


Private Sub btnsale_Click()
Unload Me
End Sub

Private Sub Command1_Click()
gen_path.t_id = t_numint
gen_path.t_modulo = "Ventas"
gen_path.t_origen = "V"
gen_path.t_path = t_path
gen_path.Show

End Sub

Private Sub Form_Activate()
Call carga
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     gen_tools.Show
End Select
End Sub

Sub electronico()
Set rs2 = New ADODB.Recordset
q = "select path from vta_014 where [num_int] = " & Val(t_numint)
rs2.Open q, cn1
If Not rs2.EOF And Not rs2.BOF Then
  t_path = rs2("path")
Else
 t_path = ""
End If
Set rs2 = Nothing

End Sub


Sub carga()
List1.clear
l1 = "-------------------------------------------------------------------------------------------------------"
l2 = "*************************************"

If t_numint <> "" Then
  'q = "select * from vta_02, vta_01, vta_06, g1 where vta_02.[id_cliente] = vta_01.[id_cliente] and vta_02.[id_tipocomp] = vta_06.[id_tipocomp] and vta_02.[num_int] = " & Val(t_numint) & " and vta_02.[sucursal_ingreso] = vta_06.[sucursal] and vta_02.[id_usuario] = g1.[id_usuario]"
  
q = "select num_int, abreviatura, fecha, sucursal_INGRESO, vta_02.id_cliente, cliente02, " _
& "direccion02, localidad02, cuit02, vta_02.id_tipocomp, SUBTOTAL, DESCUENTO, transporte, chofer02, dominio02, " _
& "dominio_acoplado02, vta_02.iva, impuestos, perc_ib, total, vta_02.observaciones, vta_02.id_vendedor, contado, vta_02.moneda, " _
& "total_otra_moneda, cae, cae_vence, estado, ctacte, vta_02.stock, grabado, id_cuenta, fecha_vto, " _
& "fecha_pago, recibo_pago, cotizacion_dolar, vta_02.id_usuario,  vta_02.iva, estado_pago, " _
& " vta_06.id_tipocomp, vta_06.sucursal, letra, vta_02.sucursal, num_comp, cp, te, email, inscripto_operador_granos, usuario, numint_asociado, valor_declarado, total_bultos " _
& " from vta_02, vta_06, vta_01, g1 where vta_02.[num_int] = " & Val(t_numint) & " and  vta_02.[id_tipocomp] = vta_06.[id_tipocomp]  and vta_02.[sucursal_ingreso] = vta_06.[sucursal]  and vta_02.[id_cliente] = vta_01.[id_cliente]" _
& " and vta_02.[id_usuario] =  g1.[id_usuario]"

  Set rs = New ADODB.Recordset
'MsgBox (q)
  
  rs.Open q, cn1
  If Not rs.EOF And Not rs.BOF Then
     List1.AddItem Space$(60) & "Numero.......:" & rs("abreviatura") & " " & rs("letra") & " " & Format$(rs("vta_02.sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
     List1.AddItem Space$(60) & "Fecha........:" & rs("fecha")
     List1.AddItem Space$(60) & "Punto Venta..:" & Format$(rs("sucursal_ingreso"), "0000")
     List1.AddItem "Cliente......:(" & Format$(rs("id_cliente"), "00000") & ") " & rs("cliente02")
     List1.AddItem "Direccion....: " & rs("direccion02") & " (" & rs("cp") & ") " & rs("localidad02")
     List1.AddItem "TE...........: " & rs("TE")
     List1.AddItem "Email........: " & rs("email")
     List1.AddItem "Cuit.........: " & rs("cuit02")
    
     Select Case rs("vta_02.id_tipocomp")
      Case Is = 45, Is = 46
        Call remitos
      Case Is = 50
        Call PAGOS
      Case Is = 401
        Call pagare
      
      Case Else
        Call COMPROBANTES
        If rs("VTA_02.ID_TIPOCOMP") = 1 Then
          Call MUESTRaREMITOS
        End If
        
     End Select
     
     Call electronico
            
     e = Space$(12)
     List1.AddItem ""
     List1.AddItem ""
     List1.AddItem "************************"
     List1.AddItem "TOTALES"
     List1.AddItem "************************"
     RSet e = Format$(rs("subtotal") + rs("descuento"), "########0.00")
     List1.AddItem "Subtotal  : " & e
     RSet e = Format$(rs("descuento"), "########0.00")
     List1.AddItem "Descuento : " & e & "-"
        
     RSet e = Format$(rs("subtotal"), "########0.00")
     List1.AddItem "Subtotal2 : " & e
     RSet e = Format$(rs("iva"), "######0.00")
     List1.AddItem "Iva       : " & e & " (Ver detalle)"
     RSet e = Format$(rs("impuestos"), "######0.00")
     RSet e = Format$(rs("perc_ib"), "######0.00")
     List1.AddItem "Percepcion: " & e & "  " & "(Ver detalle)"
     List1.AddItem "           ----------------- "
     RSet e = Format$(rs("total"), "######0.00")
     List1.AddItem "Total     : " & e
     List1.AddItem " "
     If rs("vta_02.id_tipocomp") >= 205 And rs("vta_02.id_tipocomp") <= 207 Then
       'muestro retenciones
       Set rs1 = New ADODB.Recordset
       q = "select descripcion, importe from vta_012, a12 where [num_int] = " & rs("num_int") & " and [id_retencion] = [id_percepcion]"
       rs1.Open q, cn1
       List1.AddItem "Retenciones"
       List1.AddItem "Tipo                      Importe"
       List1.AddItem "-----------------------------------"
       tr = 0
       While Not rs1.EOF
          List1.AddItem Format$(Left$(rs1("descripcion"), 20), "@@@@@@@@@@@@@@@@@@@@@@@@!") & "  " & Format$(rs1("importe"), "#######0.00")
          tr = tr + rs1("importe")
          rs1.MoveNext
       Wend
       If tr > 0 Then
          List1.AddItem "                 ________________________"
          List1.AddItem "         Total retenido" & "  " & Format$(tr, "#######0.00")
       End If
       Set rs1 = Nothing
     End If
     List1.AddItem " "
     
     List1.AddItem "Observaciones....: " & rs("observaciones")
     If rs("id_vendedor") > 0 Then
        Set rs2 = New ADODB.Recordset
        q = "select denominacion from vta_05 where [id_vendedor] = " & rs("id_vendedor")
        rs2.Open q, cn1
        If Not rs2.EOF And Not rs2.BOF Then
           vendedor = rs2("denominacion")
        Else
           vendedor = "No identificado"
        End If
        List1.AddItem "Vendedor.........: " & vendedor
        
        Set rs2 = Nothing
     End If
     List1.AddItem ""
     List1.AddItem ""
     
     Set rs2 = New ADODB.Recordset
     q = "select * from vta_09 where [num_int] = " & Val(t_numint)
     rs2.Open q, cn1
     p = 0
     While Not rs2.EOF
       If p = 0 Then
          List1.AddItem "Detalle por Tasa de Iva"
          List1.AddItem ""
          List1.AddItem "Tasa      Neto      Iva    "
          List1.AddItem "---------------------------"
          p = 1
        End If
        List1.AddItem Format$(rs2("tasa_iva"), "@@@@@") & "%" & "  " & Format$(rs2("neto"), "@@@@@@@@@@") & Format$(rs2("iva"), "@@@@@@@@@@")
        rs2.MoveNext
      Wend
     Set rs2 = Nothing
     
     
    
     
     'percvepciones
     List1.AddItem ""
     List1.AddItem ""
     Set rs2 = New ADODB.Recordset
     q = "select * from vta_016, I_01 where [num_int] = " & Val(t_numint) & " and id_percepcion = id_impuesto"
     rs2.Open q, cn1
     p = 0
     While Not rs2.EOF
       If p = 0 Then
          List1.AddItem "Detalle de percepciones"
          List1.AddItem ""
          List1.AddItem "Percepcion                 Importe   "
          List1.AddItem "-------------------------------------"
          p = 1
        End If
        List1.AddItem Format$(rs2("detalle"), "@@@@@@@@@@@@@@@@@@@@@@@@@!") & "  " & Format$(rs2("importe"), "#######0.00")
        rs2.MoveNext
      Wend
     Set rs2 = Nothing
     
     
     
     
     List1.AddItem ""
     If rs("vta_02.id_tipocomp") >= 1 And rs("vta_02.id_tipocomp") < 10 Then
       If rs("contado") = "N" Then
         
         q = "select * from vta_010 where [num_int_comp] = " & rs("num_int")
         Set rs3 = New ADODB.Recordset
         rs3.Open q, cn1
             If rs("moneda") = "P" Then
               t = rs("total")
             Else
               t = rs("total_otra_moneda")
             End If
             RSet e = Format$(t, "#####0.00")
             List1.AddItem "Cancelacion en $"
             List1.AddItem "-------------------------------------------------------------"
             List1.AddItem "Fecha       Recibo                 Importe         Saldo    "
             List1.AddItem "                                   Cancelado     Pendiente"
             List1.AddItem "-------------------------------------------------------------"
             List1.AddItem "            Deuda Original                    " & e
             s = Space$(10)
         
         While Not rs3.EOF
           q = "select sucursal, num_comp, fecha from vta_02 where [num_int] = " & rs3("num_int_rbo")
           Set rs4 = New ADODB.Recordset
           rs4.Open q, cn1
           If Not rs4.EOF And Not rs4.BOF Then
              r = Format$(rs4("sucursal"), "0000") & "-" & Format$(rs4("num_comp"), "00000000")
              F = rs4("fecha")
           Else
              r = "0000-00000000"
              F = "          "
           End If
           Set rs4 = Nothing
           t = t - rs3("importe_pagado")
           RSet e = Format$(rs3("importe_pagado"), "######0.00")
           RSet s = Format$(t, "######0.00")
           List1.AddItem F & "  " & r & "    " & e & "       " & s
           rs3.MoveNext
         Wend
         Set rs3 = Nothing
       Else
          List1.AddItem "Cancelacion en $"
          List1.AddItem "-------------------------------------------------------------"
          List1.AddItem "*** Comprobante Contado  *** "
          List1.AddItem "-------------------------------------------------------------"
         
          
          List1.AddItem ""
          List1.AddItem "FORMA PAGO "
          List1.AddItem l1
          List1.AddItem "Forma Pago   Num.Ch.     Fecha dif. Banco/Detalle                                           Importe "
          List1.AddItem l1
     
          Set rs3 = New ADODB.Recordset
          q = "select * from vta_04 where [num_int] = " & Val(t_numint)
          rs3.Open q, cn1
          i = Space$(10)
          While Not rs3.EOF
           F = Format$(rs3("fecha_dif"), "dd/mm/yy")
           FP = Format$(rs3("id_formapago"), "000") & " " & Format$(Left$(rs3("formapago"), 7), "@@@@@@@!")
           nch = Format$(rs3("num_ch"), "0000000000")
            b = Format$(Left$(rs3("detalle_banco"), 50), ">@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
            t = "" 'Format$(Left$(rs1("titular"), 25), ">@@@@@@@@@@@@@@@@@@@@@@@@@!")
            RSet i = Format$(rs3("importe"), "########0.00")
            'nix = Format$(rs3("num_int_fp"), "000000")
            List1.AddItem FP & "  " & nch & "  " & F & "  " & b & "  " & i
            rs3.MoveNext
          Wend
          Set rs3 = Nothing
     

         
         
       End If
     
     End If
     
     
          
     If rs("vta_02.id_tipocomp") = 36 Then
        Call planepago(rs("num_int"))
     End If
     
     
     If rs("vta_02.id_tipocomp") = 251 Then
        Call planepago(rs("numint_asociado"))
     End If
     
     
     List1.AddItem " "
     List1.AddItem " "
     List1.AddItem Space$(60) & "***** DATOS DE AUDITORIA ******"
     List1.AddItem Space$(60) & "CAE..........:" & rs("cae")
     List1.AddItem Space$(60) & "Venc. CAE....:" & rs("cae_vence")
     List1.AddItem Space$(60) & "Estado.......:" & rs("estado")
     List1.AddItem Space$(60) & "Cta.Cte......:" & rs("ctacte")
     List1.AddItem Space$(60) & "Stock........:" & rs("stock")
     List1.AddItem Space$(60) & "Iva..........:" & rs("grabado")
     List1.AddItem Space$(60) & "Num.Int......:" & rs("num_int")
     If rs("id_cuenta") > 0 Then
       Set rs2 = New ADODB.Recordset
       q = "select [descripcion] from c_01 where [id_cuenta] = " & rs("id_cuenta")
       rs2.MaxRecords = 1
       rs2.Open q, cn1
       If Not rs2.EOF And Not rs2.BOF Then
         cuenta = rs2("descripcion")
       Else
         cuenta = "Inexistente"
       End If
       Set rs2 = Nothing
       List1.AddItem Space$(60) & "Cuenta.......:" & rs("id_cuenta") & " - " & cuenta
     End If
     List1.AddItem Space$(60) & "Vencimiento..:" & rs("fecha_vto")
     List1.AddItem Space$(60) & "Op. Granos...: " & rs("inscripto_operador_granos")
     List1.AddItem Space$(60) & "Estado Cob...: " & rs("estado_pago")
     List1.AddItem Space$(60) & "Recibo.......: " & rs("recibo_pago")
     If rs("moneda") = "P" Then
       List1.AddItem Space$(60) & "Moneda.......: $"
     Else
       List1.AddItem Space$(60) & "Moneda.......: U$s"
     End If
     List1.AddItem Space$(60) & "Cotizacion...:" & Format$(rs("cotizacion_dolar"), "#####0.00")
     List1.AddItem Space$(60) & "Contado......:" & Format$(rs("contado"), "#####0.00")
     List1.AddItem Space$(60) & "Usuario......:" & Format$(rs("usuario"), "#####0.00")
     
 
 End If
 Set rs = Nothing
End If
End Sub


      


Sub planepago(ni)
          List1.AddItem "PLAN DE PAGO"
          List1.AddItem "-------------------------------------------------------------"
          List1.AddItem "Cuota    Fecha       Importe       Estado    Fecha   Recibo  "
          List1.AddItem "          Vto                                Pago"
          List1.AddItem "-------------------------------------------------------------"
          
          co = Space$(9)
          cf = Space$(9)
          cp = Space$(9)
          p = Space$(10)
          i = Space$(10)
          v = Space$(4)
         q = "select num_comp, fecha, total, estado_pago, fecha_pago, recibo_pago from vta_02 where [numint_asociado] = " & ni
         Set rs3 = New ADODB.Recordset
         rs3.Open q, cn1
         While Not rs3.EOF
          nc = Format$(rs3("num_comp"), "00000000")
          F = Format$(rs3("fecha"), "dd/mm/yyyy")
          RSet i = Format$(rs3("total"), "######0.00")
          If rs3("estado_pago") = "P" Then
             e = "Cancelada"
             FP = Format$(rs3("fecha_pago"), "dd/mm/yyyy")
             rp = Format$(rs3("recibo_pago"), "@@@@@@@@@@@@@@!")
         
          Else
             If DateValue(F) < DateValue(Now) Then
                e = "Vencida  "
             Else
                e = "         "
             End If
             FP = "          "
             rp = "             "
          End If
          List1.AddItem nc & " " & F & "  " & i & " " & e & " " & FP & "  " & rp
          
           rs3.MoveNext
           
         Wend
         Set rs3 = Nothing
End Sub
Sub remitos()
     List1.AddItem ""
     List1.AddItem "Transporte.........:" & rs("transporte")
     List1.AddItem "Chofer.............:" & rs("chofer02")
     List1.AddItem "Dominio Chasis.....:" & rs("dominio02") & "                           Dominio Acoplado...:" & rs("dominio_acoplado02")
     List1.AddItem ""
     List1.AddItem l1
     List1.AddItem "Codigo   Detalle                                                  Cantidad                  PU     "
     List1.AddItem "                                                        Original  Facturada  Pendiente                         "
     
     List1.AddItem l1
     
     Set rs1 = New ADODB.Recordset
     q = "select cantidad_original, cantidad, pu, id_producto, descripcion, importe, tasaiva from vta_03 where [num_int] = " & Val(t_numint) & " order by [renglon]"
     rs1.Open q, cn1
     co = Space$(9)
     cf = Space$(9)
     cp = Space$(9)
     p = Space$(10)
     i = Space$(10)
     v = Space$(4)

     While Not rs1.EOF
         RSet co = Format$(rs1("cantidad_original"), "#####0.00")
         RSet cp = Format$(rs1("cantidad"), "#####0.00")
         RSet cf = Format$(Val(co) - Val(cp), "#####0.00")
         RSet p = Format$(rs1("pu"), "######0.00")
         b = Format$(rs1("id_producto"), "00000")
         d = Format$(Left$(rs1("descripcion"), 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
         RSet i = Format$(rs1("importe"), "######0.00")
         RSet v = Format$(rs1("tasaiva"), "#0.0")
         If Val(co) > 0 Then
             List1.AddItem b & " " & d & "  " & co & " " & cf & " " & cp & "  " & p
         Else
             List1.AddItem b & " " & d
         End If
         rs1.MoveNext
     Wend
     
     
     Set rs1 = Nothing

End Sub
Sub COMPROBANTES()
     Dim a(5) As String
     List1.AddItem ""
     List1.AddItem l1
     List1.AddItem "Codigo   Detalle                                      Cantidad      PU      %Iva       Importe   %IB"
     List1.AddItem l1
     
     Set rs1 = New ADODB.Recordset
     q = "select cantidad, pu_final, pu, importe,id_producto, descripcion, tasaiva, tasaib, renglon from vta_03 where [num_int] = " & Val(t_numint) & " order by [renglon]"
     rs1.Open q, cn1
     c = Space$(7)
     p = Space$(10)
     i = Space$(10)
     v = Space$(4)
     ib = Space$(4)

     While Not rs1.EOF
         RSet c = Format$(rs1("cantidad"), "###0.00")
         If Val(t_tipocomp) = 40 Then 'presupuestos muestra precio final
            RSet p = Format$(rs1("pu_final"), "######0.00")
            RSet i = Format$(Val(c) * Val(p), "######0.00")
         Else

           If para.tipoprecioventa = 0 Then
            RSet p = Format$(rs1("pu"), "######0.00")
            RSet i = Format$(rs1("importe"), "######0.00")
          Else
            RSet p = Format$(rs1("pu_final"), "######0.00")
            RSet i = Format$(Val(c) * Val(p), "######0.00")
          End If
         End If
         
          b = Format$(rs1("id_producto"), "000")
         d = Format$(Left$(rs1("descripcion"), 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
         
         RSet v = Format$(rs1("tasaiva"), "#0.0")
         RSet ib = Format$(rs1("tasaib"), "#0.0")
         
         If Val(c) > 0 Then
             List1.AddItem b & " " & d & "  " & c & "  " & p & " (" & v & ") " & i & "   " & ib
         Else
             List1.AddItem b & " " & d
         End If
         
         'muestra descripcion extendida
          Set rs2 = New ADODB.Recordset
          q = "select * from vta_015 where [num_int] = " & Val(t_numint) & " and [renglon] = " & rs1("renglon")
          rs2.Open q, cn1
          If Not rs2.EOF And Not rs2.BOF Then
             'imprimo lineas
             Call lee_desc_extra(a, rs2("desc_ext"))
             For k = 0 To 4
              If a(k) <> "%%" Then
               d = Format$(Left$(a(k), 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
               List1.AddItem "   " & " " & d
             End If
             Next k
          End If
          Set rs2 = Nothing
         
         
         rs1.MoveNext
     Wend
     Set rs1 = Nothing
   
     
     
End Sub

Sub MUESTRaREMITOS()
     List1.AddItem ""
     List1.AddItem ""
     List1.AddItem "=========================="
     List1.AddItem "Remitos Aplicados"
     List1.AddItem "=========================="
     
     Set rst = New ADODB.Recordset
     q = "select id_tipocomp, sucursal, num_comp from vta_08, vta_02 where [id_factura] = " & Val(t_numint) & " and [id_remito] = [num_int]"
     rst.Open q, cn1
     While Not rst.EOF
         If rst("id_tipocomp") = 45 Then
            c = "Rto"
         Else
            c = "Dev"
         End If
         List1.AddItem c & " " & Format$(rst("sucursal"), "0000") & "-" & Format$(rst("num_comp"), "00000000")
         rst.MoveNext
     Wend
     Set rst = Nothing
     List1.AddItem ""
     List1.AddItem ""
End Sub


Sub PAGOS()
     i = Space$(12)
     List1.FontSize = 9
     l1 = "-------------------------------------------------------------------------------------------------------------"
     l2 = "--------------------------------------------------------------------"
     List1.AddItem "COMPROBANTES APLICADOS "
     List1.AddItem l2
     List1.AddItem "Fecha       Comprobante         Saldo $     Importe $"
     List1.AddItem "                               Pendiente    Ingresado "
     List1.AddItem l2
     
     Set rs1 = New ADODB.Recordset
     q = "select * from vta_010 where [num_int_rbo] = " & Val(t_numint) & ""
     rs1.Open q, cn1
     i = Space$(10)
     ia = Space$(10)
     While Not rs1.EOF
         Set rs2 = New ADODB.Recordset
         q = "select num_int, fecha, letra, sucursal, num_comp  from vta_02 where [num_int] = " & rs1("num_int_comp")
         
         rs2.Open q, cn1
         If Not rs2.EOF And Not rs2.BOF Then
          F = Format$(rs2("fecha"), "dd/mm/yyyy")
          nc = Format$(rs2("letra"), ">@") & " " & Format$(rs2("sucursal"), "0000") & "-" & Format$(rs2("num_comp"), "00000000")
          RSet ia = Format$(rs1("importe_pagado"), "######0.00")
          RSet i = Format$(rs1("saldo_comprobante"), "######0.00")
          List1.AddItem F & "  " & nc & "  " & i & "  " & ia
         End If
         Set rs2 = Nothing
         rs1.MoveNext
     Wend
     Set rs1 = Nothing

     List1.AddItem ""
     List1.AddItem ""
     List1.AddItem ""
     List1.AddItem "FORMA PAGO "
     List1.AddItem l1
     List1.AddItem "Forma Pago   Num.Ch.     Fecha dif. Banco/Detalle                                           Importe   Num.Int."
     List1.AddItem l1
     
     Set rs1 = New ADODB.Recordset
     q = "select * from vta_04 where [num_int] = " & Val(t_numint)
     rs1.Open q, cn1
     While Not rs1.EOF
         F = Format$(rs1("fecha_dif"), "dd/mm/yy")
         FP = Format$(rs1("id_formapago"), "000") & " " & Format$(Left$(rs1("formapago"), 7), "@@@@@@@!")
         nch = Format$(rs1("num_ch"), "0000000000")
         b = Format$(Left$(rs1("detalle_banco"), 50), ">@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
         t = "" 'Format$(Left$(rs1("titular"), 25), ">@@@@@@@@@@@@@@@@@@@@@@@@@!")
         RSet i = Format$(rs1("importe"), "########0.00")
         nix = Format$(rs1("num_int_fp"), "000000")
         List1.AddItem FP & "  " & nch & "  " & F & "  " & b & "  " & i & "   " & nix
         rs1.MoveNext
     Wend
     Set rs1 = Nothing
     
 


End Sub

Sub pagare()
List1.AddItem ""
 List1.AddItem ""
 List1.AddItem "DATOS DEL PLAN DE PAGO DEL PAGARÉ"
 List1.AddItem ""
 List1.AddItem ""
 List1.AddItem "Neto sin financiacion        : " & rs("total")

 List1.AddItem "Cantidad de cuotas           : " & rs("total_bultos")
 List1.AddItem "Importe por cuota            : " & rs("valor_declarado")
 List1.AddItem "Interés compensatorio mensual: " & rs("descuento") & "%"
 List1.AddItem "Interés compensatorio anual  : " & rs("descuento") * 12 & "%"
 List1.AddItem "Interés moratorio mensual    : " & rs("descuento") & "%"
 List1.AddItem "Interés moratorio anual      : " & rs("descuento") * 12 & "%"
 List1.AddItem "Costo financiero total       :$" & Format$((rs("valor_declarado") * rs("total_bultos")) - rs("total"), "#######0.00")
 List1.AddItem ""

 
End Sub

Sub FLETES()
     i = Space$(12)
     List1.FontSize = 9
     l1 = "-------------------------------------------------------------------------------------------------------------"
     l2 = "--------------------------------------------------------------------"
     List1.AddItem "FLETES "
     List1.AddItem l1
     List1.AddItem "Fecha    Producto       Chofer         Origen        Destino          C.P.     Ton.   kmts  Tarifa  Importe"
     List1.AddItem l1
     
     Set rs1 = New ADODB.Recordset
     q = "select * from vta_011 where [num_int] = " & Val(t_numint) & ""
     rs1.Open q, cn1
     t = Space$(6)
     k = Space$(6)
     tf = Space$(6)
     i = Space$(9)
     While Not rs1.EOF
          F = Format$(rs1("fecha"), "dd/mm/yy")
          RSet t = Format$(rs1("toneladas"), "##0.00")
          RSet k = Format$(rs1("kmts"), "##0.00")
          RSet i = Format$(rs1("importe"), "#####0.00")
          RSet tf = Format$(rs1("tarifa"), "##0.00")
          List1.AddItem F & " " & Format$(Left$(rs1("detalle"), 14), "@@@@@@@@@@@@@@!") & " " & Format$(Left$(rs1("chofer"), 14), "@@@@@@@@@@@@@@!") & " " & Format$(Left$(rs1("origen"), 12), "@@@@@@@@@@@@!") & " " & Format$(Left$(rs1("destino"), 14), "@@@@@@@@@@@@@@!") & " " & Format$(Left$(rs1("carta_porte"), 10), "@@@@@@@@@@!") & " " & t & " " & k & " " & tf & " " & i
         rs1.MoveNext
     Wend
     Set rs1 = Nothing

     
End Sub
Private Sub Form_Load()

Call barraesag(Me)


End Sub



Private Sub List1_GotFocus()
Me.StatusBar1.Panels.item(1) = "[F4]Historial Prod - [F5] Imprime Comp. - [F8] Borra Comp. - [F3] Cambia Estado - [F2]Consulta AFIP"

End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then 'historial producto
  If glo.sucursale = Val(t_sucursal) Then
    Call fe_consulta_comp
  
  Else
  
    MsgBox ("Imposible consultar comprobante en el AFIP")
  
  
  End If

End If



If KeyCode = vbKeyF4 Then 'historial producto
 Call nivel_acceso(2)
 item = Val(Mid$(List1.List(List1.ListIndex), 1, 5))
 
 If para.id_grupo_modulo_actual >= 5 Then
  If item > 1 Then
     
     vta_listaprecios4.t_idprod = item
     vta_listaprecios4.Option2 = True
     vta_listaprecios4.Show
   End If
 Else
   Call sinpermisos
 End If
End If


If KeyCode = vbKeyF8 Then
  Call nivel_acceso(1)
  If para.id_grupo_modulo_actual >= 8 Then
       J = MsgBox("Confirma Eliminar Comprobante", 4)
       If J = 6 Then
        Set cl_compvta = New comprobantes_venta
        cl_compvta.cargar2 (Val(t_numint))
        If cl_compvta.numint > 0 Then
            cl_compvta.borrar
        End If
         Set cl_compvta = Nothing
       End If
     
  Else
    Call sinpermisos
  End If
End If

If KeyCode = vbKeyF5 Then
   Call nivel_acceso(1)
   If para.id_grupo_modulo_actual >= 6 Then
     'If glo.sucursalf = 0   Then
       J = MsgBox("Imprime Comprobante", 4)
       If J = 6 Then
        Set cl_compvta = New comprobantes_venta
         cl_compvta.cargar2 (Val(t_numint))
         If cl_compvta.numint > 0 Then
            cl_compvta.imprimir
         End If
         Set cl_compvta = Nothing
       
                
       End If
     'Else
     '  MsgBox ("Por disposicion del AFIP teniendo una impresora fiscal definida no se permite imprimir otro tipo de comprobantes. Gracias")
     'End If
   Else
     Call sinpermisos
   End If
  
End If

If KeyCode = vbKeyF7 Then
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    k = 0
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 9
    While k <= List1.ListCount - 1
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
     Printer.Print List1.List(k)
     k = k + 1
    Wend
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.EndDoc
  End If
End If


If KeyCode = vbKeyF3 Then
     Call nivel_acceso(2)
     If para.id_grupo_modulo_actual >= 8 Then
       Set cl_compvta = New comprobantes_venta
       cl_compvta.cargar2 (Val(t_numint))
       If cl_compvta.numint > 0 Then
          Load vta_cambia_estado_pago
          vta_cambia_estado_pago.t_id = cl_compvta.numint
          vta_cambia_estado_pago.t_idtipocomp = cl_compvta.idtipocomp
          vta_cambia_estado_pago.t_descripcion = cl_compvta.abreviatura
          vta_cambia_estado_pago.t_estado = cl_compvta.estadopago
          vta_cambia_estado_pago.T_newestado = cl_compvta.estadopago
          vta_cambia_estado_pago.t_newestadoc = cl_compvta.estadopago
          vta_cambia_estado_pago.t_numcomp = Mid$(cl_compvta.recibopago, 6, 8)
          vta_cambia_estado_pago.t_sucursal = Mid$(cl_compvta.recibopago, 1, 4)
          vta_cambia_estado_pago.t_idprov = cl_compvta.idcliente
          vta_cambia_estado_pago.t_obs = cl_compvta.observaciones
          vta_cambia_estado_pago.t_estado2 = cl_compvta.estado
          vta_cambia_estado_pago.t_newestado2 = cl_compvta.estado
          vta_cambia_estado_pago.t_moneda = cl_compvta.moneda
          vta_cambia_estado_pago.t_cotizacion = cl_compvta.cotizaciondolar
          vta_cambia_estado_pago.t_subtotal = cl_compvta.subtotal
          vta_cambia_estado_pago.t_nograv = cl_compvta.impuestos
          vta_cambia_estado_pago.t_iva = cl_compvta.iva
          vta_cambia_estado_pago.T_TOTAL = cl_compvta.total
          vta_cambia_estado_pago.T_total2 = cl_compvta.totalotramoneda
 
          
          vta_cambia_estado_pago.Show
       End If
      Set cl_compvta = Nothing
     End If
End If


End Sub



Sub borracomp()
J = MsgBox("Confirma borrar comprobante", 4)
If J = 6 Then
     On Error GoTo errborra
     'busco el comprobante
           
     Set cl_compvta = New comprobantes_venta
     cl_compvta.cargar2 (Val(t_numint))
          
     If cl_compvta.STOCK <> "N" Then
        'modifica stock
        Set rs1 = New ADODB.Recordset
        q = "select * from vta_03, a2 where [num_int] = " & cl_compvta.numint & " and vta_03.[id_producto] = a2.[id_producto]"
        rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
        While Not rs1.EOF
           If cl_compvta.STOCK = "E" Then
              rs1("stock") = rs1("stock") - rs1("cantidad")
           Else
              rs1("stock") = rs1("stock") + rs1("cantidad")
           End If
           rs1.Update
           rs1.MoveNext
        Wend
        Set rs1 = Nothing
     End If
     
     
     cn1.BeginTrans
     'borro detalle de productos
       q = "delete * from vta_03 where [num_int] = " & cl_compvta.numint
       cn1.Execute q
     'borro stock
       q = "delete * from stk_01 where [num_mov_int] = " & cl_compvta.numint & " and [modulo] = 'V'"
       cn1.Execute q
     'borro caja
       q = "delete * from cyb_05 where [num_mov_int] = " & cl_compvta.numint & " and [modulo] = 'V'"
       cn1.Execute q
     
     
     If cl_compvta.idtipocomp = 50 Then  'recibo
        
        q = "delete * from vta_04 where [num_int] = " & cl_compvta.numint
        cn1.Execute q
        
        q = "delete * from cyb_03 where [num_int_rbo] = " & cl_compvta.numint
        cn1.Execute q
        
       
        'actualizo comp.aplicados
        q = "update vta_02 set [estado_pago] = 'N' where [recibo_pago]= '" & Format$(cl_compvta.sucursal, "0000") & "-" & Format$(cl_compvta.numcomp, "00000000") & "'"
        cn1.Execute q
        
     
     End If
     
     
      
      
     'borro comp
     q = "delete * from vta_02 where [num_int] = " & cl_compvta.numint
     cn1.Execute q
     
     cn1.CommitTrans
      
     Set cl_compvta = Nothing
    
     Unload Me
     

End If

Exit Sub

errborra:
MsgBox ("Error al Borrar Comprobante")
cn1.RollbackTrans
Exit Sub
End Sub

Sub fe_consulta_comp()
 Dim seguir As Boolean
 Set cl_compvta = New comprobantes_venta
 cl_compvta.sucursal = Val(t_sucursal)
 cl_compvta.actual (Val(t_tipocomp))
 
 
 
 List1.clear
 List1.AddItem "**********************************************************************"
 List1.AddItem " Consultas de comprobantes en el Web Service del afip"
 List1.AddItem "**********************************************************************"
 List1.AddItem ""
 
 seguir = True
    
    'On Error GoTo ManejoError
    
 If Not fe_valida_tique() Then
        'el tique esta vencido y tenemos que generarlo de nuevo
        If Not fe_genera_wsaa() Then
          MsgBox ("Error al generar tique WSAA, verificar conexion y regisar log")
          seguir = False
        End If
 End If
    
    
 If seguir Then
 
 Set WSFEv1 = CreateObject("WSFEv1")
 WSFEv1.Token = para.facte_token
 WSFEv1.Sign = para.facte_sign
 WSFEv1.CUIT = Mid$(glo.CUIT, 1, 2) & Mid$(glo.CUIT, 4, 8) & Mid$(glo.CUIT, 13, 1)
 WSFEv1.LanzarExcepciones = False
 proxy = "" ' "usuario:clave@localhost:8000"
 wsdl = para.facte_servidor_wsfe
 cache = "" 'Path
 wrapper = "" ' libreria http (httplib2, urllib2, pycurl)
 cacert = ""
 ok = WSFEv1.Conectar(cache, wsdl, proxy, wrapper, cacert) ' homologación
 ControlarExcepcion WSFEv1
 WSFEv1.Dummy
 ControlarExcepcion WSFEv1
 If (WSFEv1.AppServerStatus = "OK" And WSFEv1.DbServerStatus = "OK" And WSFEv1.AuthServerStatus = "OK") Then
    ' Buscar la factura
    If t_letra = "A" Then
        tipo_cbte = cl_compvta.cod_afip_a
    Else
        tipo_cbte = cl_compvta.cod_afip_b
    End If
    punto_vta = t_sucursal
    cbte_nro = t_numcomp
    Debug.Print tipo_cbte
    Debug.Print punto_vta
    Debug.Print cbte_nro
    
    
    cae2 = WSFEv1.CompConsultar(tipo_cbte, punto_vta, cbte_nro) 'cae garbado en el afip
    ControlarExcepcion WSFEv1

    List1.AddItem "Fecha Comprobante:" & WSFEv1.FechaCbte
     List1.AddItem "CAE:" & WSFEv1.cae
    List1.AddItem "Fecha Vencimiento CAE" & WSFEv1.Vencimiento
    List1.AddItem "Resultado:" & WSFEv1.Resultado
    List1.AddItem ""
    List1.AddItem "########################################"
    List1.AddItem "Abalisis XML Response"
    List1.AddItem "########################################"
    
    
        ok = WSFEv1.AnalizarXml("XmlResponse")
        If ok Then
            
            List1.AddItem "CbteFch:" & WSFEv1.ObtenerTagXml("CbteFch")
            List1.AddItem "Moneda:" & WSFEv1.ObtenerTagXml("MonId")
            List1.AddItem "Cotizacion:" & WSFEv1.ObtenerTagXml("MonCotiz")
            List1.AddItem "DocTIpo:" & WSFEv1.ObtenerTagXml("DocTipo")
            List1.AddItem "DocNro:" & WSFEv1.ObtenerTagXml("DocNro")
            
            ' ejemplos con arreglos (primer elemento = 0):
             List1.AddItem "Importe Total:" & WSFEv1.ImpTotal
            List1.AddItem "Primer IVA (alci id):" & WSFEv1.ObtenerTagXml("Iva", "AlicIva", 0, "Id")
            List1.AddItem "Primer IVA (importe):" & WSFEv1.ObtenerTagXml("Iva", "AlicIva", 0, "Importe")
            List1.AddItem "Segundo IVA (alic id):" & WSFEv1.ObtenerTagXml("Iva", "AlicIva", 1, "Id")
            List1.AddItem "Segundo IVA (importe):" & WSFEv1.ObtenerTagXml("Iva", "AlicIva", 1, "Importe")
            List1.AddItem "Percepcion IB (ds):" & WSFEv1.ObtenerTagXml("Tributos", "Tributo", 0, "Desc")
            List1.AddItem "Percepcion Ib (importe):" & WSFEv1.ObtenerTagXml("Tributos", "Tributo", 0, "Importe")
            List1.AddItem "Percepcion Iva (ds):" & WSFEv1.ObtenerTagXml("Tributos", "Tributo", 1, "Desc")
            List1.AddItem "Percepcion Iva (importe):" & WSFEv1.ObtenerTagXml("Tributos", "Tributo", 1, "Importe")
        Else
            ' hubo error, muestro mensaje
            Debug.Print WSFEv1.Excepcion
        End If
    
    List1.AddItem ""
    List1.AddItem "Analisis del CAE"
    List1.AddItem ""
    If cae = "" Then
        List1.AddItem "Error en el CAE"
        
        ' hubo error, no comparo
    Else
    If cae <> cae2 Then
        List1.AddItem "El CAE del comprobante guardafo localmentedifiere del guardado en el AFIP: " & cae & " vs " & cae2
    Else
        List1.AddItem "El CAE de la factura concuerdan con el recuperado de la AFIP"
    End If
    End If
End If
End If
End Sub

