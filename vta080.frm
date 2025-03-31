VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form vta_cc_detalle_ws 
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
         Picture         =   "vta080.frx":0000
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
         Picture         =   "vta080.frx":0882
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
Attribute VB_Name = "vta_cc_detalle_ws"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim l1 As String


Private Sub btnacepta_Click()
'Call verificar

End Sub

Sub verificar()
' Buscar la factura
cae2 = WSFEv1.CompConsultar(tipo_cbte, punto_vta, cbte_nro)

Debug.Print "Fecha Comprobante:", WSFEv1.FechaCbte
Debug.Print "Fecha Vencimiento CAE", WSFEv1.Vencimiento
Debug.Print "Importe Total:", WSFEv1.ImpTotal

' comparar con los datos del ejemplo anterior:
If cae <> cae2 Then
    MsgBox "El CAE de la factura no concuerdan con el recuperado en la AFIP!: " & cae & " vs " & cae2
Else
    MsgBox "El CAE de la factura concuerdan con el recuperado de la AFIP"
End If

' obtener datos del encabezado (a partir de actualización 1.17a)
cae = WSFEv1.ObtenerCampoFactura("cae")
tipo_doc = WSFEv1.ObtenerCampoFactura("tipo_doc")
nro_doc = WSFEv1.ObtenerCampoFactura("nro_doc")
imp_total = WSFEv1.ObtenerCampoFactura("imp_total")
' obtener primer alicuota de IVA
imp_iva1 = WSFEv1.ObtenerCampoFactura("iva", 0, "importe")
' obtener primer tributo
imp_trib1 = WSFEv1.ObtenerCampoFactura("tributos", 0, "importe")
' obtener primer opcional
valor_opcional1 = WSFEv1.ObtenerCampoFactura("opcionales", 0, "valor")
' obtener primer código de observacion de AFIP
obs_code1 = WSFEv1.ObtenerCampoFactura("obs", 0, "code")
' pruebo obtener el segundo mensaje de observacion inexistente
obs_code2 = WSFEv1.ObtenerCampoFactura("obs", 1, "msg")
Debug.Print WSFEv1.Excepcion  ' "El campo 1 solicitado no existe"
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub


Private Sub Form_Activate()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
End Sub

Private Sub Form_Load()

Call barraesag(Me)


End Sub



Private Sub List1_GotFocus()
Me.StatusBar1.Panels.item(1) = ""

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

