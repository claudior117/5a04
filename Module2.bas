Attribute VB_Name = "Module2"
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Public Type varudt 'se crea una estructura para los parametros
   id_usuario As Integer 'usuario actual
   id_grupo As String
   grupo As String
   usuario As String
   id_obraactual As Integer
   tasaiva(10) As Double 'carga un array con las tasas con id como indice
   cuenta_deudores As Long 'clientes ctacte
   cuenta_acreedores As Long 'proveedores ctacte
   cuenta_ventas As Long
   cuenta_caja As Long
   cuenta_iva_compras As Long
   cuenta_iva_ventas As Long
   cuenta_retgan As Long
   cuenta_retib As Long
   cuenta_conceptos_nograbados As Long
   cuenta_perc_IB As Long
   cuenta_perc_iva As Long
   cuenta_retibbav As Long
   cuenta_retivav As Long
   cuenta_retsussv As Long
   cuenta_retganv As Long
   cuenta_valores_terceros As Long
   numeracion_comun_Fact_nc As String
   cotizacion As Double
   moneda As String
   tasageneral As Single
   producto_sel As Double 'producto seleccionado
   cuenta_sel As Double 'cuenta seleccionada
   sincroniza_bancos As String
   id_grupo_modulo_actual As Integer 'asigna cada vez que ingresa  un modulo
   id_grupo_modulo_compras As Integer 'nivel de usuario en compras
   calcula_ret_ib As String
   calcula_perc_ib As String
   usuario_inicio As Integer ' lleva el listindex del usuario habitual de la maquina
   archivo_exportacion As String
   minimo_retib As Double 'importe minimo de la op para retener ib
   empresa As String 'tiene el nombre de carpeta del sistema que se pasa en el ejecutable
   id_periodo_contable As Long
   ancho As Long
   alto As Long
   tipoactupreciocompcompra As Integer 'define el tipo de actualizacion en la lista de precio segun comp. compra
   muestraagenda As String
   tipoprecioventa As Integer 'define el tipo de precio venta a utilizar 0 pu  1 pf
   password_adm  As String
   fiscal As Integer ' 0 NO  1 SI
   IMPRESORA_PREDETERMINADA As String
   impresora_actual As String
   cuenta_compras_varias As Long
   cuenta_inventario As Long
   cuenta_costo As Long
   imprime_pie_reportes As Boolean
   HABILITACION As Integer
   exporta_sel As Long
   numint_regfaltante As Long
   z_actual As Long 'el numero de z actual si tiene fiscal sino 0
   tipolistaprecios As Integer
   fechacorte As Date
   tasaib As Single
   tiporedondeo As Integer
   ncenrecibo As String
   muestrasaldofactventa As String
   imprime_cabecera_reportes As String
   facte_token As String
   facte_sign As String
   facte_expira As Date
   facte_claveprivada As String
   facte_certificado As String
   facte_servidor_wsaa As String
   facte_servidor_wsfe As String
   nombre_fantasia As String
   fecha_inicio_actividades As String
   numero_ib As String
   idsistema As Integer
   punto_venta_usuario As Integer
   tipo_iva_empresa As Integer
 End Type


Public para As varudt 'define la variable para con estruct. varudt



Public cl_prod As productos
Public cl_comp As COMPROBANTES
Public cl_prov As proveedores
Public cl_cli As Clientes
Public cl_compvta As comprobantes_venta
Public cl_usuarios As usuarios
Public cl_compprod As comprobantes_produccion
Public cl_chterc As chterceros
Public cl_banco As bancos
Public cl_padronib As padron_ib
Public cl_stock As STOCK
Public cl_fiscal As fiscal



Public a_nr() As String 'array de numeros de requesicion
Public cnib As ADODB.Connection
Public cnrep As ADODB.Connection

'factura electronica
Dim WSAA As Object, WSFEv1 As Object

'controlador fiscal nuevo protocolo (2020)
Public fiscal As Driver
Public cMODELO As Integer
Public cPUERTO As Integer
Public cBAUDIOS As Long


Public Function busca_saldos_prov(ByVal cp As Long, ByVal m As String, ByVal F As Date) As Double
    'busca saldos  cp a= cod. proveedor   m = moneda p pesos d dolares f = fecha hasta
    q = "select * from a5 where [id_proveedor] = " & cp & " and [ctacte] <> 'N' "
    q = q & " and datevalue([fecha]) < datevalue('" & F & "')"
    Set rs = New ADODB.Recordset
    rs.Open q, cn1
    da = 0
    ha = 0
    sa = 0
    While Not rs.EOF
     If rs("ctacte") = "D" Then
        da = da + rs("total")
     Else
        ha = ha + rs("total")
     End If
     rs.MoveNext
    Wend
    sa = da - ha
    busca_saldos_prov = sa
End Function


Public Function sacaactividadsucursal(ByVal s As Integer) As Integer
 'devuelve la actividad comercial predefinida para la sucursal s
  Set rs = New ADODB.Recordset
  q = "select * from g8 where [sucursal_predefinida] = " & s
  rs.Open q, cn1
  If Not rs.EOF And Not rs.BOF Then
    a = rs("id_actividad")
  Else
    a = 1
  End If
  Set rs = Nothing
  sacaactividadsucursal = a
End Function
 


Public Function abrirconexion(u As String, p As String) As Boolean   'proc. que abre la conexion con la base de datos
  'u usuario
  ' password
  
 
  abrirconexion = False
  On Error GoTo manerr
  Set cn1 = New ADODB.Connection
  gconexion = "Provider=Microsoft.Jet.oledb.4.0;Data Source=" & App.Path & "\dat\5a04.mdb;User id=" & u & ";password=" & p & ";" & "Jet OLEDB:System database=" & App.Path & "\SEG\system1.mdw;"
 
 ' (sql) gconexion = "Provider=SQLOLEDB; Initial Catalog=5a04sql; Data Source=(local)\SQL5A04; integrated security=SSPI; persist security info=True;"
 
  
  cn1.Open gconexion
  
  abrirconexion = True
  glo.conexion = gconexion
  
  
  
  Exit Function


manerr:
   MsgBox ("Error al Abrir Base de Datos, Verifique su Usuario y Password")
   abrirconexion = False
   End
End Function


Public Function abrirconexionib() As Boolean   'proc. que abre la conexion con la base de datos del padron ib
  'u usuario
  ' password
  u = "claudio"
  p = "0969"
  abrirconexionib = False
  On Error GoTo manerr
  Set cnib = New ADODB.Connection
  gconexion = "Provider=Microsoft.Jet.oledb.4.0;Data Source=" & App.Path & "\dat\pib.mdb;User id=" & u & ";password=" & p & ";" & "Jet OLEDB:System database=" & App.Path & "\SEG\system2.mdw;"
 
  cnib.Open gconexion
  abrirconexionib = True
   
  Exit Function


manerr:
   MsgBox ("Error al Abrir Base de Datos del Padron de IB, Verifique su Usuario y Password")
   abrirconexionib = False
   End
End Function
Public Function codproddesdebarras(ByVal codbarra As Double)
'devuelve el id del producto teniendo el cod. de barras
cp = 0
Set rscb = New ADODB.Recordset
q = "select * from a2 where [cod_barras] = " & codbarra
rscb.Open q, cn1
If Not rscb.EOF And Not rscb.BOF Then
      cp = rs("id_producto")
End If
codproddesdebarras = cp
Set recb = Nothing
End Function
Public Function abrirconexionrep() As Boolean 'proc. que abre la conexion con la base de datos de reportes para emitir etiquetas
 'u usuario
  ' password
  u = "claudio"
  p = "0969"
  abrirconexionrep = False
  'On Error GoTo manerr
  Set cnrep = New ADODB.Connection
  gconexion = "Provider=Microsoft.Jet.oledb.4.0;Data Source=" & "c:\5a04\rep\dat\rep2.mdb;"
  cnrep.Open gconexion
  abrirconexionrep = True
   
  Exit Function


manerr:
   MsgBox ("Error al Abrir Base de Datos de Reportes, Verifique su Usuario y Password")
   abrirconexionrep = False
   End


End Function

Sub barraesag(p_form As Form)
'BARRA ESTADO
p_form.StatusBar1.Panels.item(1) = "[ENTER] Avanza - [Up] Regresa - [ESC] Sale - [F4] Abre Combo -  [F12] Tools"

End Sub

Sub barracgr(p_form As Form)
'BARRA ESTADO
Set rs = New ADODB.Recordset
q = "select * from c_10 where [id_periodo] = " & para.id_periodo_contable
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
  t = rs("descripcion")
Else
  t = " "
End If
Set rs = Nothing
'FIXIT: 'StatusBar1.Panels.Item(1' no es una propiedad del objeto genérico 'Form' en Visual Basic .NET. Para obtener acceso a 'StatusBar1.Panels.Item(1', declare 'p_form' utilizando su tipo real en lugar de 'Form'     FixIT90210ae-R1460-RCFE85
p_form.StatusBar1.Panels.item(1) = "Periodo Trabajo: " & t & " (" & para.id_periodo_contable & ")"
'FIXIT: 'StatusBar1.Panels.Item(2' no es una propiedad del objeto genérico 'Form' en Visual Basic .NET. Para obtener acceso a 'StatusBar1.Panels.Item(2', declare 'p_form' utilizando su tipo real en lugar de 'Form'     FixIT90210ae-R1460-RCFE85
p_form.StatusBar1.Panels.item(2) = "[ENTER] Avanza - [Up] Regresa - [ESC] Regresa - [F4] Despliega"

End Sub
Sub ejecutareporte(a As Adodc, r As Report)
Load reportes
Set rs = a.Recordset
r.DiscardSavedData
r.Database.SetDataSource rs
r.ReadRecords
r.Texto1.SetText (glo.nombrecli)
r.Texto2.SetText (glo.direccioncli)
reportes!CRViewer1.ReportSource = r
reportes!CRViewer1.ViewReport
reportes.Show
Set rs = Nothing


End Sub

Sub ejecutareporte2(ByVal a As Recordset, ByVal r As Report)
Load reportes
r.DiscardSavedData
r.Database.SetDataSource a
r.ReadRecords
r.Texto1.SetText (glo.nombrecli)
r.Texto2.SetText (glo.direccioncli)
reportes!CRViewer1.ReportSource = r
reportes!CRViewer1.ViewReport
reportes.Show
Set a = Nothing


End Sub

Function saca_ultnumero_comp(ByVal tc As Integer) As Long
'devuelve el ultimo numero utilizado del comprobante Y ACTUALIZA BASE

Set rs = New ADODB.Recordset
q = "select [ult_num] from g2 where [id_tipo_comp] = " & tc
rs.MaxRecords = 1
rs.Open q, cn1, adOpenStatic, adLockOptimistic
If Not rs.EOF And Not rs.BOF Then
     p = rs("ult_num") + 1
     rs("ult_num") = p
     rs.Update
     saca_ultnumero_comp = p
Else
  MsgBox ("Error al Inicializar Comprobante. Funcion Saca_ultnumero_comp")
  saca_ultnumero_comp = 0
End If
Set rs = Nothing
End Function

Function saca_ultnumero_comp2(ByVal tc As Integer) As Long
'devuelve el ultimo numero utilizado del comprobante SIN ACTUALIZA R BASE

Set rs = New ADODB.Recordset
q = "select [ult_num] from g2 where [id_tipo_comp] = " & tc
rs.MaxRecords = 1
rs.Open q, cn1, adOpenStatic, adLockOptimistic
If Not rs.EOF And Not rs.BOF Then
     p = rs("ult_num") + 1
     saca_ultnumero_comp2 = p
Else
  MsgBox ("Error al Inicializar Comprobante. Funcion Saca_ultnumero_comp")
  saca_ultnumero_comp2 = 0
End If
Set rs = Nothing
End Function

Function saca_ultnumero_int_comp(modulo As String) As Long
'devuelve el ultimo numero interno modulo = C Compras // modulo = V ventas

Set rs = New ADODB.Recordset
q = "select [ult_num_int_comp], [ult_num_int_vta], [ult_num_int_cgr], [ult_num_int_prod]  from g0 where [sucursal] = " & 0
rs.MaxRecords = 1
rs.Open q, cn1, adOpenStatic, adLockOptimistic
If Not rs.EOF And Not rs.BOF Then
   
   Select Case modulo
     Case Is = "C" 'COMPRAS
     p = rs("ult_num_int_comp") + 1
     rs("ult_num_int_comp") = p
     rs.Update
     Case Is = "V" 'VENTAS
     p = rs("ult_num_int_vta") + 1
     rs("ult_num_int_vta") = p
     rs.Update
     Case Is = "G" 'CONTABILIDAD
     p = rs("ult_num_int_CGR") + 1
     rs("ult_num_int_CGR") = p
     rs.Update
     Case Is = "P" 'PRODUCCION
     p = rs("ult_num_int_PROD") + 1
     rs("ult_num_int_PROD") = p
     rs.Update
  
  End Select
  saca_ultnumero_int_comp = p
Else
  MsgBox ("Error al Inicializar Comprobante. Funcion Saca_ultnumero_Int_comp. La sucuesal no esta definida")
  saca_ultnumero_int_comp = 0
End If
Set rs = Nothing
End Function

Sub keyform(F As Form, ByVal k As String)
  'k = A  Activa keyprevie en el formulariuo
  'k = D desactiva
  If k = "A" Then
     F.KeyPreview = True
  Else
     F.KeyPreview = False
  End If
  
End Sub


Function HEXABIN(NUM As String) As String
Const CD = 4 'cantidad de digitos hexadecimal
Static POSH(CD) As String * 1
b = 16
NUM = Format$(NUM, "0000")
For i = 0 To CD - 1
 POSH(i) = Mid$(NUM, CD - i, 1)
Next i

NUMDEC = 0
For i = 0 To CD - 1
  Select Case POSH(i)
    Case Is = "0"
       v = 0
    Case Is = "1"
       v = 1
    Case Is = "2"
       v = 2
    Case Is = "3"
       v = 3
    Case Is = "4"
       v = 4
    Case Is = "5"
       v = 5
    Case Is = "6"
       v = 6
    Case Is = "7"
       v = 7
    Case Is = "8"
       v = 8
    Case Is = "9"
       v = 9
    Case Is = "A"
       v = 10
    Case Is = "B"
       v = 11
    Case Is = "C"
       v = 12
    Case Is = "D"
       v = 13
    Case Is = "E"
       v = 14
    Case Is = "F"
       v = 15
    Case Else
      HEXABIN = "0000"
      Exit Function
 End Select
 NUMDEC = NUMDEC + ((b ^ i) * v)
Next i

'convertir binario
v = NUMDEC
nb = ""
While v >= 2
   v = v / 2
   r = v - Fix(v)
   v = Fix(v)
   db = r * 2
   nb = Format$(db, "@") & nb
Wend
nb = Format$(v, "@") & nb
HEXABIN = Format$(nb, "0000000000000000")



End Function

Function abrir_archivo_digital(ByVal p As String) As Boolean
'p es el path del archivo
'devuelve true si el archivo exist y se pude abrir
'On Error GoTo erra12
s = False
If p = "" Then
  MsgBox ("No hay archivo seleccionado")
Else
  ShellExecute hWnd, "open", p, "", "", 4
 
  s = True
End If
abrir_archivo_digital = s
Exit Function

erra12:
MsgBox ("Error al abrir el archivo")
abrir_archivo_digital = False
Exit Function

End Function

Public Sub lee_desc_extra(a() As String, ByVal t As String)
'el array a tiene las 5 lineas de 50 caracters como maximo a imprimir
'las lineas sin datos se marcan con un doble porcentaje (%%)
'la funcion tom a el texto t que viene deun multilinea con finalñes de carro incluidos
' los devide en 5 lineas imprimibles.
fin = Len(t)
i = 1
la = 0 'linea del array actual
pa = 0 'posicion del texto actual (no puede ser myor a 50)
l = ""
For k = 0 To 4
  a(k) = ""
Next k

While i <= fin
  'observo carater por caracter
    
  If Mid$(t, i, 1) = Chr$(13) Then 'encuentro un retorno de carro
     'termino unalinea porque encontre un final de carro
     
     If la <= 4 Then
         a(la) = l
         l = ""
         pa = 0
         la = la + 1
      Else
         MsgBox ("ERROR!!! El texto tiene mas de 5 lineas")
      End If
     i = i + 1
  Else
    If pa < 50 Then
      l = l & Mid$(t, i, 1) 'voy creando la linea
      pa = pa + 1
    Else
      'llegue al final de una linea
      If la <= 4 Then
         a(la) = l
         l = ""
         pa = 0
         la = la + 1
      Else
         MsgBox ("ERROR!!! El texto tiene mas de 5 lineas")
      End If
    End If
  End If
  i = i + 1
Wend
If l <> "" Then
   a(la) = l
   la = la + 1
End If
If la <= 4 Then
   For k = la To 4
     a(k) = "%%"
   Next k
End If

'el sistema devuelve el array a completo



End Sub

