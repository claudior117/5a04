VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "productos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"STOCK"
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
'variables locales para almacenar los valores de las propiedades
Private mvaridproducto As Long 'copia local
Private mvardetalle As String 'copia local
Private mvarcodtasaiva As Integer 'copia local
Private mvarenvase As Double 'copia local
Private mvarSTOCK As Double 'copia local
Private mvarrequerido As Double 'copia local
Private mvarpedido As Double 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarprecio_ult_compra As Double 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarpreciounitario As Double 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarid_unidad As Integer 'copia local
Private mvarunidad As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvaridproveedor As Long 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvartasaiva As Single 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarcostoreal As Double 'copia local
Public Sub borrar(ByVal idproducto As Long)
'verifico en ventas
b = 1
q = "select [id_producto] from vta_03 where [id_producto] = " & idproducto
Set rs = New adodb.Recordset
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
  b = 0
End If
Set rs = Nothing

'verifico en comprs
q = "select [id_producto] from a6 where [id_producto] = " & idproducto
Set rs = New adodb.Recordset
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
  b = 0
End If
Set rs = Nothing

If b = 1 Then
  Set rs = New adodb.Recordset
  q = "select * from a2 where [id_producto] = " & idproducto
  rs.Open q, cn1, adOpenDynamic, adLockOptimistic
  If Not rs.EOF And Not rs.BOF Then
    rs.Delete
    rs.Update
    MsgBox ("Producto borrado satisfactoriamente")
  Else
    MsgBox ("El producto ya no existe en la Base de datos")
  End If
  Set rs = Nothing
Else
  MsgBox ("Imposible eliminar el producto porque tiene movimientos asociados")
End If
End Sub

Public Property Let costoreal(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.costoreal = 5
    mvarcostoreal = vData
End Property


Public Property Get costoreal() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.costoreal
    costoreal = mvarcostoreal
End Property



Public Property Let tasaiva(ByVal vData As Single)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.tasaiva = 5
    mvartasaiva = vData
End Property


Public Property Get tasaiva() As Single
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.tasaiva
    tasaiva = mvartasaiva
End Property



Public Property Let idproveedor(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.idproveedor = 5
    mvaridproveedor = vData
End Property


Public Property Get idproveedor() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.idproveedor
    idproveedor = mvaridproveedor
End Property



Public Property Let unidad(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.unidad = 5
    mvarunidad = vData
End Property


Public Property Get unidad() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.unidad
    unidad = mvarunidad
End Property



Public Property Let id_unidad(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.id_unidad = 5
    mvarid_unidad = vData
End Property


Public Property Get id_unidad() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.id_unidad
    id_unidad = mvarid_unidad
End Property



Public Property Let preciounitario(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.preciounitario = 5
    mvarpreciounitario = vData
End Property


Public Property Get preciounitario() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.preciounitario
    preciounitario = mvarpreciounitario
End Property



Public Function stock_anterior(ByVal idprod As Long, ByVal fecha As Date) As Double
'devuelve el stock anterior a la fecha del producto
Set rs = New adodb.Recordset
q = "select * from stk_01 where [id_producto] = " & idprod & " and datevalue([fecha]) < datevalue('" & fecha & "')"
rs.Open q, cn1
s = 0
While Not rs.EOF
  If rs("ubicacion") = "E" Then
    s = s + rs("cantidad")
  Else
    s = s - rs("cantidad")
  End If
  rs.MoveNext
Wend
Set rs = Nothing
stock_anterior = s
End Function

Public Function stock_anterior_por_cliente(ByVal idprod As Long, ByVal fecha As Date, ByVal idcli As Long) As Double
'devuelve el stock anterior a la fecha del producto
Set rs = New adodb.Recordset
q = "select * from stk_01 where [id_producto] = " & idprod & " and [id_cliente]= " & idcli & "  and datevalue([fecha]) < datevalue('" & fecha & "')"
rs.Open q, cn1
s = 0
While Not rs.EOF
  If rs("ubicacion") = "E" Then
    s = s + rs("cantidad")
  Else
    s = s - rs("cantidad")
  End If
  rs.MoveNext
Wend
Set rs = Nothing
stock_anterior_por_cliente = s
End Function
Public Property Let precio_ult_compra(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.precio_ult_compra = 5
    mvarprecio_ult_compra = vData
End Property


Public Property Get precio_ult_compra() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.precio_ult_compra
    precio_ult_compra = mvarprecio_ult_compra
End Property




Public Sub actualizar(ByVal ip As Long, ByVal STOCK As Double, ByVal requerido As Double, ByVal pedido As Double)
'actualiza campos stock, requeridos o pediddos.
' si un campo no se tiene que actualizar se pasa 0
Set rs = New adodb.Recordset
q = "select * from a2 where [id_producto] = " & ip
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
If Not rs.EOF And Not rs.BOF Then
    rs("stock") = rs("stock") + STOCK
    rs("requeridos") = rs("requeridos") + requerido
    rs("pedidos") = rs("pedidos") + pedido
    rs.Update
End If
Set rs = Nothing

End Sub

Public Sub cargar(ByVal ip As Long)
Set rsc = New adodb.Recordset
q = "select * from a2, g5, g4 where [id_producto] = " & ip & " and a2.[id_unidad] = g5.[id_unidad] and [id_tasaiva] = [cod_tasaiva]"
rsc.Open q, cn1
If Not rsc.EOF And Not rsc.BOF Then
  Me.idproducto = rsc("id_producto")
  Me.Detalle = rsc("a2.descripcion")
  Me.codtasaiva = rsc("cod_tasaiva")
  Me.envase = rsc("envase")
  Me.pedido = rsc("pedidos")
  Me.requerido = rsc("requeridos")
  Me.STOCK = rsc("stock")
  Me.precio_ult_compra = rsc("precio_ult_compra")
  Me.preciounitario = rsc("pu")
  Me.id_unidad = rsc("a2.id_unidad")
  Me.unidad = rsc("unidad")
  Me.idproveedor = rsc("id_proveedor")
  Me.tasaiva = rsc("tasa")
  Me.costoreal = rsc("costoreal")
Else
  Me.idproducto = 0
End If
Set rsc = Nothing
End Sub
Public Sub cargafaltante(ByVal ip As Long, ByVal cant As Double, ByVal idprov As Long)
          Me.cargar (ip)
          If Me.idproducto > 0 Then
            Set rs1 = New adodb.Recordset
            q = "select * from a6 where [num_int] = " & para.numint_regfaltante
            rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
            If Not rs1.EOF And Not rs1.BOF Then
              rs1.MoveLast
              nr = rs1("renglon") + 1
            
            Else
              nr = 1
            
            End If
            rs1.AddNew
            rs1("num_int") = para.numint_regfaltante
            rs1("renglon") = nr
            rs1("id_producto") = Me.idproducto
            rs1("detalle") = Me.Detalle
            rs1("cantidad") = cant
            If idprov = 0 Then
               rs1("envase") = Me.idproveedor  'contiene el id. del proveedor
            Else
               rs1("envase") = idprov
            End If
            rs1("fecha") = Format$(Now, "dd/mm/yyyy")
            rs1("unidad") = Me.unidad
            rs1("pu") = Me.precio_ult_compra 'precio sugerido de ultima compra
            rs1("id_usuario") = para.id_usuario
            rs1.Update
            Set rs1 = Nothing
            
            Set rs1 = New adodb.Recordset
            q = "select * from a2 where [id_producto] = " & ip
            rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
            If Not rs1.EOF And Not rs1.BOF Then
               rs1("Reg_faltante") = rs1("reg_faltante") + cant
               rs1.Update
            End If
            Set rs1 = Nothing
         End If
      
End Sub

Public Sub modificafaltante(ByVal r As Integer, ByVal cant As Double, ByVal idprov As Long)
      'modifica un articulo ingresado en el registro de faltantes en el renglo r y la nueva canticac cant
      
            Set rs1 = New adodb.Recordset
            q = "select * from a6 where [num_int] = " & para.numint_regfaltante & " and [renglon] = " & r
            rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
            If Not rs1.EOF And Not rs1.BOF Then
               cantant = rs1("cantidad")
               ip = rs1("id_producto")
               rs1("cantidad") = cant
               rs1("envase") = idprov
               rs1("fecha") = Format$(Now, "dd/mm/yyyy")
               rs1("id_usuario") = para.id_usuario
                rs1.Update
            End If
            Set rs1 = Nothing
            
           Set rs1 = New adodb.Recordset
           q = "select * from a2 where [id_producto] = " & ip
           rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
           If Not rs1.EOF And Not rs1.BOF Then
               rs1("Reg_faltante") = rs1("reg_faltante") - cantant + cant
               rs1.Update
            End If
            Set rs1 = Nothing
      
End Sub


Public Sub sacafaltante(ByVal r As Double)
            Set rs1 = New adodb.Recordset
            q = "select * from a6 where [num_int] = " & para.numint_regfaltante & " and [renglon] = " & r
            rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
            c = 0
            If Not rs1.EOF And Not rs1.BOF Then
               ip = rs1("id_producto")
               c = rs1("cantidad")
               
               rs1.Delete
            End If
            rs1.Update
            Set rs1 = Nothing
            
            Set rs1 = New adodb.Recordset
            q = "select * from a2 where [id_producto] = " & ip
            rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
            If Not rs1.EOF And Not rs1.BOF Then
               rs1("Reg_faltante") = rs1("reg_faltante") - c
               rs1.Update
            End If
            Set rs1 = Nothing
      End Sub

Public Sub borraprodfaltante(ByVal idproducto As Long)

q = "select * from a6 where [num_int] = " & para.numint_regfaltante & " and [id_producto] = " & idproducto
Set rsx = New adodb.Recordset
rsx.Open q, cn1, adOpenDynamic, adLockOptimistic
While Not rsx.EOF
  rsx.Delete
  rsx.MoveNext
Wend
Set rsx = Nothing

Set rsx = New adodb.Recordset
q = "select * from a2 where [id_producto] = " & idproducto
rsx.Open q, cn1, adOpenDynamic, adLockOptimistic
If Not rsx.EOF And Not rsx.BOF Then
    rsx("Reg_faltante") = 0
    rsx.Update
End If
Set rsx = Nothing

End Sub
Public Property Let pedido(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.pedido = 5
    mvarpedido = vData
End Property


Public Property Get pedido() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.pedido
    pedido = mvarpedido
End Property



Public Property Let requerido(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.requerido = 5
    mvarrequerido = vData
End Property


Public Property Get requerido() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.requerido
    requerido = mvarrequerido
End Property



Public Property Let STOCK(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.stock = 5
    mvarSTOCK = vData
End Property


Public Property Get STOCK() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.stock
    STOCK = mvarSTOCK
End Property



Public Property Let envase(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.envase = 5
    mvarenvase = vData
End Property


Public Property Get envase() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.envase
    envase = mvarenvase
End Property



Public Property Let codtasaiva(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.codtasaiva = 5
    mvarcodtasaiva = vData
End Property


Public Property Get codtasaiva() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codtasaiva
    codtasaiva = mvarcodtasaiva
End Property



Public Property Let Detalle(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.detalle = 5
    mvardetalle = vData
End Property


Public Property Get Detalle() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.detalle
    Detalle = mvardetalle
End Property



Public Property Let idproducto(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.idproducto = 5
    mvaridproducto = vData
End Property


Public Property Get idproducto() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.idproducto
    idproducto = mvaridproducto
End Property



