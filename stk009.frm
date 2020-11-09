VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form stk_cc_detalle 
   BackColor       =   &H00E0E0E0&
   Caption         =   "MOVIMIENTOS STOCK(DETALLE)"
   ClientHeight    =   8700
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12075
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   12075
   StartUpPosition =   3  'Windows Default
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
      Height          =   5235
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   11895
   End
   Begin VB.Frame CUIT 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   11655
      Begin VB.TextBox t_numint 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox t_tipocomp 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   240
         MaxLength       =   6
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800080&
         Caption         =   "Numero Interno"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800080&
         Caption         =   "Tipo"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   2
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "stk009.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "stk009.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   3
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
      TabIndex        =   1
      Top             =   8445
      Width           =   12075
      _ExtentX        =   21299
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
            TextSave        =   "21/01/2010"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "06:09 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "stk_cc_detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim l1 As String


Private Sub btnsale_Click()
Unload Me
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

Sub carga()
List1.clear
l1 = "--------------------------------------------------------------------------------------------------------------------------------"
l2 = "*************************************"

If t_numint <> "" Then
  q = "select * from STK_02, STK_03 where  stk_02.[num_int] = stk_03.[num_int] and stk_02.[num_int] = " & Val(t_numint)
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  If Not rs.EOF And Not rs.BOF Then
     
     Select Case rs("tipo_comprobante")
     Case Is = 1
       t = "Ajuste"
       List1.AddItem "================================="
       List1.AddItem t
       List1.AddItem "================================="
       List1.AddItem ""
       List1.AddItem Space$(60) & "Numero Int...:" & t & "  " & Format$(rs("stk_02.num_int"), "00000000")
       List1.AddItem Space$(60) & "Fecha........:" & rs("fecha")
     
     Case Is = 20
       t = "Entrada"
       List1.AddItem "================================="
       List1.AddItem t
       List1.AddItem "================================="
       List1.AddItem ""
       List1.AddItem Space$(60) & "Numero Int...:" & t & "  " & Format$(rs("stk_02.num_int"), "00000000")
       List1.AddItem Space$(60) & "Fecha........:" & rs("fecha")
       Set rs1 = New ADODB.Recordset
       q = "select * from a1 where [id_proveedor] = " & rs("id_proveedor")
       rs1.Open q, cn1
       If Not rs1.EOF And Not rs1.BOF Then
         List1.AddItem "Proveedor......: " & rs1("denominacion")
       Else
        List1.AddItem "Proveedor......: "
       End If
       Set rs1 = Nothing
   
       
       List1.AddItem "Comprobante....: " & rs("letra") & " " & Format$(rs("sucursal"), "0000") & " - " & Format$(rs("num_comprobante"), "00000000")
       Set cl_prov = Nothing
     Case Is = 30
       t = "Salida"
       List1.AddItem "================================="
       List1.AddItem t
       List1.AddItem "================================="
       List1.AddItem ""
       List1.AddItem Space$(60) & "Numero Int...:" & t & "  " & Format$(rs("stk_02.num_int"), "00000000")
       List1.AddItem Space$(60) & "Fecha........:" & rs("fecha")
 
       Set rs1 = New ADODB.Recordset
       q = "select * from a4 where [id_obra] = " & rs("id_obra")
       rs1.Open q, cn1
       If Not rs1.EOF And Not rs1.BOF Then
         List1.AddItem "Obra......: " & rs1("descripcion")
       Else
        List1.AddItem "Obra......: "
       End If
       Set rs1 = Nothing
     End Select
 
     List1.AddItem l1
     List1.AddItem "Codigo  Producto                                            Cantidad  Unidad   Detalle                   Tipo"
     List1.AddItem l1
     
     co = Space$(9)
     
     While Not rs.EOF
         RSet co = Format$(rs("cantidad"), "#####0.00")
         b = Format$(rs("id_producto"), "00000")
         d = Format$(Left$(rs("descripcion"), 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
         u = Format$(Left$(rs("unidad"), 6), "@@@@@@!")
         de = Format$(Left$(rs("stk_03.detalle") & " ", 30), "@@@@@@@@@@@@@@@@@@@@@@@@@@!")
         List1.AddItem b & "    " & d & " " & co & " " & u & "   " & de & "  " & rs("ubicacion")
         rs.MoveNext
     Wend
     
 End If
 Set rs = Nothing
End If
End Sub
Sub remitos()
     List1.AddItem ""
     List1.AddItem "Transporte..:" & rs("transporte")
     List1.AddItem ""
     List1.AddItem l1
     List1.AddItem "Codigo   Detalle                                                  Cantidad                  PU     "
     List1.AddItem "                                                        Original  Facturada  Pendiente                         "
     
     List1.AddItem l1
     
     Set rs1 = New ADODB.Recordset
     q = "select * from vta_03 where [num_int] = " & Val(t_numint)
     rs1.Open q, cn1
     co = Space$(9)
     cf = Space$(9)
     cp = Space$(9)
     p = Space$(10)
     i = Space$(10)
     V = Space$(4)

     While Not rs1.EOF
         RSet co = Format$(rs1("cantidad_original"), "#####0.00")
         RSet cp = Format$(rs1("cantidad"), "#####0.00")
         RSet cf = Format$(Val(co) - Val(cp), "#####0.00")
         RSet p = Format$(rs1("pu"), "######0.00")
         b = Format$(rs1("id_producto"), "000")
         d = Format$(Left$(rs1("descripcion"), 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
         RSet i = Format$(rs1("importe"), "######0.00")
         RSet V = Format$(rs1("tasaiva"), "#0.0")
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
     List1.AddItem ""
     List1.AddItem l1
     List1.AddItem "Codigo   Detalle                                      Cantidad      PU      %Iva       Importe"
     List1.AddItem l1
     
     Set rs1 = New ADODB.Recordset
     q = "select * from vta_03 where [num_int] = " & Val(t_numint)
     rs1.Open q, cn1
     c = Space$(7)
     p = Space$(10)
     i = Space$(10)
     V = Space$(4)

     While Not rs1.EOF
         RSet c = Format$(rs1("cantidad"), "###0.00")
         RSet p = Format$(rs1("pu"), "######0.00")
         b = Format$(rs1("id_producto"), "000")
         d = Format$(Left$(rs1("descripcion"), 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
         RSet i = Format$(rs1("importe"), "######0.00")
         RSet V = Format$(rs1("tasaiva"), "#0.0")
         If Val(c) > 0 Then
             List1.AddItem b & " " & d & "  " & c & "  " & p & " (" & V & ") " & i
         Else
             List1.AddItem b & " " & d
         End If
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
     q = "select * from vta_08, vta_02 where [id_factura] = " & Val(t_numint) & " and [id_remito] = [num_int]"
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
     l2 = "------------------------------------------------"
     List1.AddItem ""
     List1.AddItem "COMPROBANTES APLICADOS "
     List1.AddItem l2
     List1.AddItem "Fecha       Comprobante      Importe   Tipo"
     List1.AddItem l2
     
     nop = Format$(t_sucursal, "0000") & "-" & Format$(t_numcomp, "00000000")
     Set rs1 = New ADODB.Recordset
     q = "select * from vta_02 where [recibo_pago] = '" & nop & "'"
     rs1.Open q, cn1
     While Not rs1.EOF
         f = Format$(rs1("fecha"), "dd/mm/yyyy")
         nc = Format$(rs1("letra"), ">@") & " " & Format$(rs1("sucursal"), "0000") & "-" & Format$(rs1("num_comp"), "00000000")
         i = Format$(rs1("total"), "######0.00")
         List1.AddItem f & "  " & nc & "  " & i & "  " & rs1("id_tipocomp")
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
         f = Format$(rs1("fecha_dif"), "dd/mm/yy")
         fp = Format$(rs1("id_formapago"), "000") & " " & Format$(Left$(rs1("formapago"), 7), "@@@@@@@!")
         nch = Format$(rs1("num_ch"), "0000000000")
         b = Format$(Left$(rs1("detalle_banco"), 50), ">@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
         t = "" 'Format$(Left$(rs1("titular"), 25), ">@@@@@@@@@@@@@@@@@@@@@@@@@!")
         RSet i = Format$(rs1("importe"), "########0.00")
         nix = Format$(rs1("num_int_fp"), "000000")
         List1.AddItem fp & "  " & nch & "  " & f & "  " & b & "  " & i & "   " & nix
         rs1.MoveNext
     Wend
     Set rs1 = Nothing
     
 


End Sub
Private Sub Form_Load()

Call barraesag(Me)


End Sub



Private Sub List1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F5] Imprime Mov. - [F8] Borra Mov. "

End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF8 Then
  Call nivel_acceso(8)
  If para.id_grupo_modulo_actual >= 8 Then
       J = MsgBox("Confirma Eliminar Comprobante", 4)
       If J = 6 Then
        Set cl_stock = New STOCK
             
        Call cl_stock.borra_mov_stk(Val(t_numint), "S")
        
        Set cl_stock = Nothing
       End If
     
  Else
    Call sinpermisos
  End If
End If

If KeyCode = vbKeyF5 Then
   Call nivel_acceso(1)
   If para.id_grupo_modulo_actual >= 6 Then
     
       J = MsgBox("Imprime Comprobante", 4)
       If J = 6 Then
        Set cl_compvta = New comprobantes_venta
         cl_compvta.cargar2 (Val(t_numint))
         If cl_compvta.numint > 0 Then
            cl_compvta.imprimir
         End If
         Set cl_compvta = Nothing
       End If
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
     If para.id_grupo_modulo_actual >= 9 Then
       Set cl_compvta = New comprobantes_venta
       cl_compvta.cargar2 (Val(t_numint))
       If cl_compvta.numint > 0 Then
          Load vta_cambia_estado_pago
          vta_cambia_estado_pago.t_id = cl_compvta.numint
          vta_cambia_estado_pago.t_descripcion = cl_compvta.abreviatura
          vta_cambia_estado_pago.t_estado = cl_compvta.estadopago
          vta_cambia_estado_pago.T_newestado = cl_compvta.estadopago
          vta_cambia_estado_pago.t_numcomp = Mid$(cl_compvta.recibopago, 6, 8)
          vta_cambia_estado_pago.t_sucursal = Mid$(cl_compvta.recibopago, 1, 4)
          vta_cambia_estado_pago.T_IDPROV = cl_compvta.idcliente
          vta_cambia_estado_pago.t_obs = cl_compvta.observaciones
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

