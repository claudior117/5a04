Attribute VB_Name = "Module3"
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private Declare Function GetVolumeInformation Lib "Kernel32" _
    Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
                                    ByVal lpVolumeNameBuffer As String, _
                                    ByVal nVolumeNameSize As Long, _
                                    lpVolumeSerialNumber As Long, _
                                    lpMaximumComponentLength As Long, _
                                    lpFileSystemFlags As Long, _
                                    ByVal lpFileSystemNameBuffer As String, _
                                    ByVal nFileSystemNameSize As Long) As Long


Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public a_estado_m(5) As String 'estado materiales
Public a_estado_o(4) As String 'estado obras
Public Const lineacompleta = "--------------------------------------------------------------------------------------"
Sub carga_marcas(c As ComboBox)
 'marcas
Set rs = New ADODB.Recordset
q = "select * from a10 order by descripcion"
rs.Open q, cn1
Call llena_combo(rs, "descripcion", "id_marca", c, True)
c.ListIndex = 0
Set rs = Nothing

End Sub
Sub carga_camiones(c As ComboBox, ByVal t As Long)
't es el codigo de transporte 0 todos
Set rs = New ADODB.Recordset
q = "select [id_camion], [camion], [chofer], [dominio] from a17"
If t > 0 Then
  q = q & " where [id_transporte] = " & t
End If
q = q & " order by [camion], [id_camion]"
rs.Open q, cn1
c.clear
If Not rs.EOF And Not rs.BOF Then
 While Not rs.EOF
  c.AddItem rs("camion") & " [" & rs("dominio") & "]" & rs("chofer")
  c.ItemData(c.NewIndex) = rs("id_camion")
  rs.MoveNext
 Wend
Else
  c.AddItem "Sin Datos"
  c.ItemData(c.NewIndex) = 1
End If
c.ListIndex = 0
Set rs = Nothing
End Sub
Function verificaperiodo(ByVal f As String) As String
'f es una fecha
'devuelve A Abierto  C Cerrado
p = Val(Mid$(Format$(f, "dd/mm/yyyy"), 7, 4) & Mid$(Format$(f, "dd/mm/yyyy"), 4, 2))
q = "select [estado] from a14 where [id_periodo] = " & p
Set rs = New ADODB.Recordset
rs.MaxRecords = 1
rs.Open q, cn1
If Not rs.BOF And Not rs.EOF Then
   'existe
   verificaperiodo = rs("estado")
Else
   verificaperiodo = "A"
End If
Set rs = Nothing
End Function

Function verificaperiodog(ByVal f As String) As String
'f es una fecha
'devuelve A Abierto  C Cerrado
p = Val(Mid$(Format$(f, "dd/mm/yyyy"), 7, 4) & Mid$(Format$(f, "dd/mm/yyyy"), 4, 2))
q = "select [estado] from g10 where [periodo] = " & p
Set rsx = New ADODB.Recordset
rsx.MaxRecords = 1
rsx.Open q, cn1
If Not rsx.BOF And Not rsx.EOF Then
   'existe
   verificaperiodog = rsx("estado")
Else
   verificaperiodog = "A"
End If
Set rsx = Nothing
End Function

Sub borracontabilidad(ByVal nro As Long, ByVal m As String)
'borra los mov. contables
'se le debe pasar el num. interno del ciomprobnante y el modulo que lo genera
' el llamado debe estar entre un cn1.BeginTrans y un cn1.comminstrans

      'contabilidad
      
      Set rsm = New ADODB.Recordset
      q = "select * from c_02 where [num_mov_int] = " & nro & " and [modulo] = '" & m & "'"
      rsm.Open q, cn1
      If Not rsm.EOF And Not rsm.BOF Then
        nicgr = rsm("num_interno")
      Else
        nicgr = 0
      End If
      Set rsm = Nothing
      
      QUERY = "DELETE FROM c_02 WHERE [num_interno] = " & nicgr
      cn1.Execute QUERY
      
      QUERY = "DELETE FROM c_03 WHERE [num_interno] = " & nicgr
      cn1.Execute QUERY
      
      
      
End Sub

Sub borracontabilidad2(ByVal nro As Long, ByVal m As String)
'borra los mov. contables
'se le debe pasar el num. interno del ciomprobnante y el modulo que lo genera

      'contabilidad
      
      Set rsm = New ADODB.Recordset
      q = "select * from c_02 where [num_mov_int] = " & nro & " and [modulo] = '" & m & "'"
      rsm.Open q, cn1, adOpenDynamic, adLockOptimistic
      If Not rsm.EOF And Not rsm.BOF Then
        q = "select * from c_03 where [num_interno] = " & rsm("num_interno")
        Set rsm2 = New ADODB.Recordset
        rsm2.Open q, cn1, adOpenDynamic, adLockOptimistic
        While Not rsm2.EOF
          rsm2.Delete
          rsm2.MoveNext
        Wend
        Set rsm2 = Nothing
        rsm.Delete
        
      End If
      Set rsm = Nothing
      
      
End Sub
Sub carga_percepciones(c As ComboBox)
 'marcas
Set rs = New ADODB.Recordset
q = "select * from a12 order by descripcion"
rs.Open q, cn1
Call llena_combo(rs, "descripcion", "id_percepcion", c, True)
c.ListIndex = 0
Set rs = Nothing

End Sub

Sub carga_deptos_venta(c As ComboBox)
'departamentos
Set rs = New ADODB.Recordset
q = "select * from a9 order by descripcion"
rs.Open q, cn1
Call llena_combo(rs, "descripcion", "id_departamento", c, True)
c.ListIndex = 0
Set rs = Nothing

End Sub

Sub carga_actividades(c As ComboBox)
'departamentos
Set rs = New ADODB.Recordset
q = "select [id_actividad], [descripcion] from g8 "
rs.Open q, cn1
Call llena_combo(rs, "descripcion", "id_actividad", c, True)
c.ListIndex = 0
Set rs = Nothing

End Sub


Sub carga_grupos(c As ComboBox)
'grupos
Set rs = New ADODB.Recordset
q = "select * from a8 order by descripcion"
rs.Open q, cn1
Call llena_combo(rs, "descripcion", "id_grupo", c, True)
c.ListIndex = 0
Set rs = Nothing

End Sub

Sub SACARSERIAL()
    'programa que obtiene  el nro. de serie del hard disk y lo guarda en un archivo llamado c:\temp.txt
    Dim lVSN As Long, n As Long, s1 As String, s2 As String
    Dim unidad As String
    Dim sTmp As String
    Dim t As String
    On Local Error Resume Next
    'Se debe especificar el directorio raiz
    unidad = "C:\"
    'Reservar espacio para las cadenas que se pasarán al API
    s1 = String$(255, Chr$(0))
    s2 = String$(255, Chr$(0))
    n = GetVolumeInformation(unidad, s1, Len(s1), lVSN, 0, 0, s2, Len(s2))
    's1 será la etiqueta del volumen
    'lVSN tendrá el valor del Volume Serial Number (número de serie del volumen)
    's2 el tipo de archivos: FAT, etc.
    'Convertirlo a hexadecimal para mostrarlo como en el Dir.
    sTmp = Hex$(lVSN)
    Open "c:\windows\system\temp.txt" For Output As #1
    t = (Left$(sTmp, 4) & "-" & Right$(sTmp, 4))
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7593-R67265
    Print #1, t
    Close #1
    
End Sub

Sub nivel_acceso(i As Integer)
'i es numero de modulo  1  Ventas
'                       2  Compras
'                       3 Caja
'                       4 Bancos
'                       5 Productos
'                       6 Produccion
'                       7 contabilidad
'                       8 stock
n = Mid$(para.id_grupo, i, 1)
c = Mid$(para.id_grupo, 2, 1)
para.id_grupo_modulo_actual = Val(n)
para.id_grupo_modulo_compras = Val(c)

End Sub
Sub carga_usuarios_ini(c As ComboBox)
  c.clear
  On Error GoTo manerr
  Set cn1 = New ADODB.Connection
  gconexion = "Provider=Microsoft.Jet.oledb.4.0;Data Source=" & App.Path & "\dat\5a04.mdb;User id=" & "claudio" & ";password=0969" & ";" & "Jet OLEDB:System database=" & App.Path & "\SEG\system2.mdw;"
  cn1.Open gconexion
  
Set rs = New ADODB.Recordset
q = "select * from g1 "
rs.Open q, cn1
While Not rs.EOF
  c.AddItem rs("usuario")
'FIXIT: c.ItemData(c.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
  c.ItemData(c.NewIndex) = rs("id_usuario")
  rs.MoveNext
Wend
c.ListIndex = 0
Set rs = Nothing
cn1.Close

  
Exit Sub
manerr:
   MsgBox ("Error Tabla Usuarios")
   End


End Sub

Sub carga_usuarios(c As ComboBox)
c.clear

Set rs = New ADODB.Recordset
q = "select * from g1 "
rs.Open q, cn1
While Not rs.EOF
  c.AddItem rs("usuario")
'FIXIT: c.ItemData(c.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
  c.ItemData(c.NewIndex) = rs("id_usuario")
  
  rs.MoveNext
Wend
c.ListIndex = 0
Set rs = Nothing

End Sub

Sub carga_estadosM(c As ComboBox)
c.AddItem "[R] Req.Pendiente", 0
c.AddItem "[P] Pedido Realizado", 1
c.AddItem "[S] Req.Cumplida ", 2
c.AddItem "[C] Cancelada", 3
c.AddItem "[O] Otros", 4
c.ListIndex = 0

End Sub
Function activaobra(ByVal i As Long) As Integer
'devuelve 1 si la actrivacion fue correcta o 0 si no se pudo
On Error GoTo e1
Set rs = New ADODB.Recordset
q = "select * from g0 where [sucursal] = 0"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
rs("id_obraactual") = i
rs.Update
para.id_obraactual = i
activaobra = 1
Set rs = Nothing

Exit Function
e1:
  activaobra = 0
End Function
Sub carga_tipoiva(c As ComboBox)
Set rs = New ADODB.Recordset
q = "select * from g3"
rs.Open q, cn1
c.clear
While Not rs.EOF
  c.AddItem rs("descripcion")
'FIXIT: c.ItemData(c.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
  c.ItemData(c.NewIndex) = rs("cod_tipoiva")
  rs.MoveNext
Wend

c.ListIndex = 0
Set rs = Nothing
End Sub

Sub carga_tasaiva(c As ComboBox)
c.clear
For i = 0 To 9
  c.AddItem para.tasaiva(i)
'FIXIT: c.ItemData(c.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
  c.ItemData(c.NewIndex) = i
Next i

c.ListIndex = 0

End Sub
Sub carga_vendedores(c As ComboBox)
Set rs = New ADODB.Recordset
q = "select [id_vendedor], [denominacion] from vta_05 order by [denominacion]"
rs.Open q, cn1
c.clear
While Not rs.EOF
  c.AddItem rs("denominacion")
'FIXIT: c.ItemData(c.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
  c.ItemData(c.NewIndex) = rs("id_vendedor")
  rs.MoveNext
Wend

c.ListIndex = 0
Set rs = Nothing
End Sub

Sub carga_impuestos(c As ComboBox, ByVal i As Integer)
'i = cod. de impuesto

Set rs = New ADODB.Recordset
q = "select * from i_02 where [id_impuesto] = " & i
rs.Open q, cn1
c.clear
While Not rs.EOF
  c.AddItem rs("concepto")
'FIXIT: c.ItemData(c.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
  c.ItemData(c.NewIndex) = rs("id_concepto")
  rs.MoveNext
Wend

c.ListIndex = 0
Set rs = Nothing
End Sub

Sub carga_impuesto(c As ComboBox)
'impuesto

Set rs = New ADODB.Recordset
q = "select * from i_01  "
rs.Open q, cn1
c.clear
While Not rs.EOF
  c.AddItem rs("detalle")
'FIXIT: c.ItemData(c.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
  c.ItemData(c.NewIndex) = rs("id_impuesto")
  rs.MoveNext
Wend

'c.ListIndex = 0
Set rs = Nothing
End Sub

Sub carga_cuentas_cont(c As ComboBox, ByVal tipo As String, ByVal orden As String, Optional caja As String)
'tipo = "C" Cuenta   "T" Titulos    S "Todos"
'orden = "C" Id. Cuenta   "D" descripcion"

'Caja: Tipo cuenta caja    T Todas, I Ingreso, E Egreso

If IsMissing(caja) Then
  caja = "T"
End If

Set rs = New ADODB.Recordset
q = "select [id_cuenta], [descripcion], [tipo_cuentacaja] from c_01"
co = " where "
If tipo <> "S" Then
     q = q & co & " [tipo] = '" & tipo & "'"
     co = " and "
End If

If caja <> "T" Then
  If caja = "I" Then
    q = q & co & "([tipo_cuentacaja] = 'A' or [tipo_cuentacaja] = 'I')"
  Else
    q = q & co & "([tipo_cuentacaja] = 'A' or [tipo_cuentacaja] = 'E')"
  End If
End If



If orden = "C" Then
  q = q & " ORDER BY [id_cuenta]"
Else
  q = q & " ORDER BY [DESCRIPCION]"
End If

rs.Open q, cn1
c.clear
While Not rs.EOF
  c.AddItem rs("descripcion")
'FIXIT: c.ItemData(c.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
  c.ItemData(c.NewIndex) = rs("id_cuenta")
  rs.MoveNext
Wend

c.ListIndex = 0
Set rs = Nothing
End Sub




Sub carga_formas_pago(c As ComboBox, ByVal tipo As String)
'tipo = "B" Bancos(<=50)   "O" Otras (<20) "T" Otras con ch. terc.(<=20)   S "Todas"  y = otras sin las 1234
'C solo mueven caja
Set rs = New ADODB.Recordset
q = "select [id_forma_pago], [descripcion] from cyb_01"
Select Case tipo
Case Is = "B"
     q = q & " where [id_forma_pago] >= 50"
Case Is = "O"
     q = q & " where [id_forma_pago] <= 20 and [id_forma_pago] <> 3 and [id_forma_pago] <> 4 "
Case Is = "T"
     q = q & " where [id_forma_pago] <= 20"
Case Is = "Y"
     q = q & " where [id_forma_pago] = 1 or ([id_forma_pago] > 4 and  [id_forma_pago] < 20)"
Case Is = "C"
     q = q & " where [caja] = 'S'"


End Select
rs.Open q, cn1
c.clear
While Not rs.EOF
  c.AddItem rs("descripcion")
'FIXIT: c.ItemData(c.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
  c.ItemData(c.NewIndex) = rs("id_forma_pago")
  rs.MoveNext
Wend

c.ListIndex = 0
Set rs = Nothing
End Sub

Sub carga_iva()
'carga un array con las tasas con id como indice
For i = 0 To 9
  para.tasaiva(i) = 0
Next i


Set rs = New ADODB.Recordset
q = "select * from g4"
rs.Open q, cn1
While Not rs.EOF
  para.tasaiva(rs("id_tasaiva")) = Format$(rs("tasa"), "00.00")
  rs.MoveNext
Wend
Set rs = Nothing
End Sub

Sub carga_tipoib(c As ComboBox)
Set rs = New ADODB.Recordset
q = "select * from g6"
rs.Open q, cn1
c.clear
While Not rs.EOF
  c.AddItem rs("descripcion")
'FIXIT: c.ItemData(c.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
  c.ItemData(c.NewIndex) = rs("id_tipoib")
  rs.MoveNext
Wend

c.ListIndex = 0
Set rs = Nothing

End Sub
Sub carga_tipocomp(c As ComboBox)
Set rs = New ADODB.Recordset
q = "select [id_tipo_comp], [descripcion] from g2"
rs.Open q, cn1
c.clear
While Not rs.EOF
  c.AddItem rs("descripcion")
'FIXIT: c.ItemData(c.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
  c.ItemData(c.NewIndex) = rs("id_tipo_comp")
  rs.MoveNext
Wend

c.ListIndex = 0
Set rs = Nothing
End Sub

Sub carga_provincias(c As ComboBox)
Set rs = New ADODB.Recordset
q = "select * from g9 order by [provincia]"
rs.Open q, cn1
c.clear
While Not rs.EOF
  c.AddItem rs("provincia")
'FIXIT: c.ItemData(c.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
  c.ItemData(c.NewIndex) = rs("id_prov")
  rs.MoveNext
Wend

c.ListIndex = 0
Set rs = Nothing
End Sub

Sub carga_tipocompprod(c As ComboBox)
Set rs = New ADODB.Recordset
q = "select [id_tipocomp], [descripcion] from pro_03"
rs.Open q, cn1
c.clear
While Not rs.EOF
  c.AddItem rs("descripcion")
'FIXIT: c.ItemData(c.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
  c.ItemData(c.NewIndex) = rs("id_tipocomp")
  rs.MoveNext
Wend

c.ListIndex = 0
Set rs = Nothing
End Sub

Sub carga_unidad(c As ComboBox)
Set rs = New ADODB.Recordset
q = "select * from g5"
rs.Open q, cn1
c.clear
While Not rs.EOF
  c.AddItem rs("unidad")
'FIXIT: c.ItemData(c.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
  c.ItemData(c.NewIndex) = rs("id_unidad")
  rs.MoveNext
Wend

c.ListIndex = 0
Set rs = Nothing
End Sub

Sub carga_dbcrbanco(c As ComboBox)
Set rs = New ADODB.Recordset
q = "select * from cyb_07 order by [descripcion]"
rs.Open q, cn1
c.clear
While Not rs.EOF
  c.AddItem rs("Descripcion")
'FIXIT: c.ItemData(c.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
  c.ItemData(c.NewIndex) = rs("id_tipomov")
  rs.MoveNext
Wend

c.ListIndex = 0
Set rs = Nothing
End Sub

Sub carga_clientes(c As ComboBox)
Set rs = New ADODB.Recordset
q = "select [id_cliente], [denominacion] from vta_01 order by denominacion"
rs.Open q, cn1
c.clear
While Not rs.EOF
  c.AddItem rs("Denominacion")
'FIXIT: c.ItemData(c.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
  c.ItemData(c.NewIndex) = rs("id_cliente")
  rs.MoveNext
Wend

c.ListIndex = 0
Set rs = Nothing
End Sub

Sub carga_piezas(c As ComboBox)
Set rs = New ADODB.Recordset
q = "select [id_pieza], [descripcion] from pro_06"
q = q & " order by [descripcion]"
rs.Open q, cn1
c.clear
If Not rs.EOF And Not rs.BOF Then
 While Not rs.EOF
  c.AddItem rs("descripcion")
  c.ItemData(c.NewIndex) = rs("id_pieza")
  rs.MoveNext
 Wend
Else
  c.AddItem "Sin Datos"
  c.ItemData(c.NewIndex) = 1
End If
c.ListIndex = 0
Set rs = Nothing
End Sub

Sub carga_SUCURSALES(c As ComboBox)
Set rs = New ADODB.Recordset
q = "select * from vta_06 order by [SUCURSAL]"
rs.Open q, cn1
c.clear
p = 0
's = 0
While Not rs.EOF
  If p = 0 Then
     c.AddItem Format$(rs("SUCURSAL"), "0000")
     c.ItemData(c.NewIndex) = rs("SUCURSAL")
     p = rs("SUCURSAL")
  End If
  
  If p <> rs("SUCURSAL") Then
     c.AddItem Format$(rs("SUCURSAL"), "0000")
     c.ItemData(c.NewIndex) = rs("SUCURSAL")
     p = rs("SUCURSAL")
  End If
  rs.MoveNext
Wend
c.ListIndex = buscaindice(c, glo.sucursal)
Set rs = Nothing
End Sub

Sub carga_empleados(c As ComboBox)
Set rs = New ADODB.Recordset
q = "select [id_legajo], [denominacion] from emp_01 where [estado] = 'A' order by denominacion"
rs.Open q, cn1
c.clear
While Not rs.EOF
  c.AddItem rs("Denominacion")
'FIXIT: c.ItemData(c.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
  c.ItemData(c.NewIndex) = rs("id_legajo")
  rs.MoveNext
Wend

c.ListIndex = 0
Set rs = Nothing
End Sub

Sub carga_periodos(c As ComboBox)
Set rs = New ADODB.Recordset
q = "select * from c_10 where [id_periodo] > 0 "
rs.Open q, cn1
c.clear
While Not rs.EOF
  c.AddItem rs("descripcion")
'FIXIT: c.ItemData(c.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
  c.ItemData(c.NewIndex) = rs("id_periodo")
  rs.MoveNext
Wend

c.ListIndex = 0
Set rs = Nothing
End Sub

Sub carga_obras(c As ComboBox, estado As String)
'T TERMINADAS  E EN EJECUCION  S SUSPENDIDAS  O OTRAS  A TODAS
Set rs = New ADODB.Recordset
q = "select * from a4"
If estado <> "A" Then
  q = q & " where [estado] = '" & estado & "'"
End If

q = q & " order by descripcion"

rs.Open q, cn1
c.clear
While Not rs.EOF
  c.AddItem rs("Descripcion")
'FIXIT: c.ItemData(c.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
  c.ItemData(c.NewIndex) = rs("id_obra")
  rs.MoveNext
Wend
If c.ListIndex >= 0 Then
  c.ListIndex = 0
End If
Set rs = Nothing
End Sub

Sub carga_productos(c As ComboBox)
Set rs = New ADODB.Recordset
q = "select [id_producto], [descripcion] from a2 order by descripcion"
rs.Open q, cn1
c.clear
While Not rs.EOF
  c.AddItem rs("Descripcion")
'FIXIT: c.ItemData(c.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
  c.ItemData(c.NewIndex) = rs("id_producto")
  rs.MoveNext
Wend

c.ListIndex = 0
Set rs = Nothing
End Sub
Sub carga_proveedores(c As ComboBox)
Set rs = New ADODB.Recordset
q = "select [id_proveedor], [denominacion] from a1 order by denominacion"
rs.Open q, cn1
c.clear
While Not rs.EOF
  c.AddItem rs("Denominacion")
'FIXIT: c.ItemData(c.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
  c.ItemData(c.NewIndex) = rs("id_proveedor")
  rs.MoveNext
Wend

c.ListIndex = 0
Set rs = Nothing
End Sub

Sub carga_transporte(c As ComboBox)
Set rs = New ADODB.Recordset
q = "select [id_proveedor], [denominacion] from a1 where [transporte] = 'S' order by denominacion"
rs.Open q, cn1
c.clear

While Not rs.EOF
  c.AddItem rs("Denominacion")
'FIXIT: c.ItemData(c.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
  c.ItemData(c.NewIndex) = rs("id_proveedor")
  rs.MoveNext
Wend

Set rs = Nothing

End Sub



Sub sinpermisos()
J = MsgBox("No tiene los permisos necesarios para esta operacion", vbCritical, "Error de Seguridad")

End Sub


