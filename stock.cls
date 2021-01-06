VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'variables locales para almacenar los valores de las propiedades
Private mvaridproducto As Long 'copia local
Private mvarstock_instantaneo As Double 'copia local
Private mvarstock_movimientos As Double 'copia local
Public Sub sacastock(ByVal idproducto As Long, Optional ByVal F As Date)
Set rsc = New adodb.Recordset
q = "select [stock] from a2 where [id_producto] = " & idproducto
rsc.Open q, cn1
If Not rsc.EOF And Not rsc.BOF Then
  Me.stock_instantaneo = rsc("stock")
Else
  Me.stock_instantaneo = 0
End If

Set rsc = Nothing

Set rsc = New adodb.Recordset
q = "select [ubicacion], [cantidad] from stk_01 where [id_producto] = " & idproducto

 If IsMissing(F) Then
      q = q & " and datevalue([fecha]) <= datevalue('" & F & "') "
 End If

rsc.Open q, cn1
s = 0
While Not rsc.EOF
 If rsc("ubicacion") = "E" Then
   s = s + rsc("cantidad")
 Else
    s = s - rsc("cantidad")
 End If
 rsc.MoveNext
Wend
Me.stock_movimientos = s
Set rsc = Nothing


End Sub





Public Function sacastockvalorizado(ByVal F As String, ByVal t As String)
'f fecha corte
't tipo:   C Precio Compra  S Costo   V Precio Final Venta
Set rsc = New adodb.Recordset
q = "select * from a2 where [id_producto] > 1 "
rsc.Open q, cn1
st = 0

While Not rsc.EOF
  Select Case t
   Case Is = "C"
     precio = rsc("precio_ult_compra")
   Case Is = "S"
     precio = rsc("costoreal")
   Case Is = "V"
     precio = rsc("precio_final")
   Case Else
     precio = rsc("costoreal")
  End Select
     
  Set rsc2 = New adodb.Recordset
  q = "select * from stk_01 where [id_producto] = " & rsc("id_producto") & " and datevalue([fecha]) <= datevalue('" & F & "')"
  rsc2.Open q, cn1
  s = 0
  While Not rsc2.EOF
   If rsc2("ubicacion") = "E" Then
      s = s + (rsc2("cantidad"))
   Else
      s = s - (rsc2("cantidad"))
   End If
   rsc2.MoveNext
  Wend
  If IsNull(precio) Then
    precio = 0
  End If
  st = st + (Format(s, "######0.00") * precio)
  rsc.MoveNext
Wend
Set rsc = Nothing
sacastockvalorizado = Format(st, "######0.00")

End Function

Public Function sacafaltante(ByVal ip As Long)
q = "select * from a6 where [num_int] = " & para.numint_regfaltante & " and [id_producto] = " & ip
Set rsx = New adodb.Recordset
rsx.Open q, cn1
trf = 0
While Not rsx.EOF
  trf = trf + rsx("cantidad")
  rsx.MoveNext
Wend
Set rsx = Nothing
sacafaltante = trf

End Function



Sub actualizastock()
  J = MsgBox("Este proceso actualiza stock Instantaneo desde Stock de Movimientos  y puede demorar. Es necesario salir del sistema de todas las terminales para ejecutarlo. ¿Confirma?", 4)
  If J = 6 Then
   Load espere
   q = "select * from a2 order by [id_producto]"
   Set rs = New adodb.Recordset
   rs.Open q, cn1, adOpenDynamic, adLockOptimistic
   rs.MoveLast
   ri = 1
   rf = rs("id_producto")
   rs.MoveFirst
   espere.ProgressBar1.Min = 1
   espere.ProgressBar1.Max = rf
   espere.ProgressBar1.Value = 1
   espere.Show
   espere.Refresh
   While Not rs.EOF
     ip = rs("id_producto")
     Call sacastock(ip)
     rs("stock") = Me.stock_movimientos
     rs("reg_faltante") = sacafaltante(ip)
     
     rs.Update
     espere.ProgressBar1.Value = ip
     espere.ProgressBar1.Refresh
     rs.MoveNext
   Wend
   Set rs = Nothing
   
   
   
   
   
   
   
   Unload espere
   MsgBox ("Proceso Terminado")
  End If
  
End Sub



Public Sub borra_mov_stk(ByVal numint As Long, ByVal modulo As String)
'actualizo stock instantaneo y borro stock movimientos
Set rsc = New adodb.Recordset
q = "select * from stk_01 where [num_mov_int] = " & numint & " and [modulo] = '" & modulo & "'"
rsc.Open q, cn1, adOpenDynamic, adLockOptimistic
While Not rsc.EOF
 Set rsc2 = New adodb.Recordset
 q = "select * from a2 where [id_producto] = " & rsc("id_producto")
 rsc2.Open q, cn1, adOpenDynamic, adLockOptimistic
 If Not rsc2.EOF And Not rsc2.BOF Then
   If rsc("ubicacion") = "E" Then
      rsc2("stock") = rsc2("stock") - rsc("cantidad")
   Else
      rsc2("stock") = rsc2("stock") + rsc("cantidad")
   End If
   rsc2.Update
 End If
 Set rsc2 = Nothing
 rsc.Delete
 rsc.MoveNext
Wend
Set rsc = Nothing

If modulo = "S" Then ' si el movimientos es del modulo stock lo borro
  Set rsc = New adodb.Recordset
  q = "select * from stk_02 where [num_int] = " & numint
  rsc.Open q, cn1, adOpenDynamic, adLockOptimistic
  While Not rsc.EOF
     rsc.Delete
     rsc.MoveNext
  Wend
  Set rsc = Nothing
  
  Set rsc = New adodb.Recordset
  q = "select * from stk_03 where [num_int] = " & numint
  rsc.Open q, cn1, adOpenDynamic, adLockOptimistic
  While Not rsc.EOF
     rsc.Delete
     rsc.MoveNext
  Wend
  Set rsc = Nothing
  
  Call borracontabilidad2(numint, "S")
  
End If



End Sub



Public Property Let stock_movimientos(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.stock_movimientos = 5
    mvarstock_movimientos = vData
End Property


Public Property Get stock_movimientos() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.stock_movimientos
    stock_movimientos = mvarstock_movimientos
End Property



Public Property Let stock_instantaneo(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.stock_instantaneo = 5
    mvarstock_instantaneo = vData
End Property


Public Property Get stock_instantaneo() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.stock_instantaneo
    stock_instantaneo = mvarstock_instantaneo
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


'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Public Sub imprimir(ByVal id As Long)
Set rs22 = New adodb.Recordset
q = "SELECT * FROM STK_02, STK_03, g1 WHERE STK_02.[NUM_INT] = " & id & " AND STK_02.[NUM_INT] = STK_03.[NUM_INT] and stk_02.[Id_usuario] = g1.[id_usuario]"
rs22.Open q, cn1
If Not rs22.EOF And Not rs22.BOF Then
 c = InputBox$("Cantidad de Copias", "Impresion de Movimientos Stock", 1)
 If Val(c) > 0 Then
  For Y = 1 To c
   Call imprimeempresa(14)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Times New Roman"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "MOVIMIENTO INTERNO STOCK"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(80); "Nro. Comprobante:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FillStyle = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & Format$(id, "00000000")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(93); "Fecha:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & Format$(rs22("FECHA"), "dd/mm/yyyy")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(3); "Detalle....................:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & rs22("stk_02.detalle"); Tab(92);
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "Emitida por:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & rs22("usuario")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Courier new"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 8
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "____________________________________________________________________________________________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "        Basico  Detalle                                        Cantidad    Observaciones                     Tipo   "
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "____________________________________________________________________________________________________________________"
   r = 1
   cant = Space$(10)
   pu = Space$(10)
   While Not rs22.EOF
       renglon = Format$(r, "000")
       basico = Format$(rs22("id_producto"), "00000")
       Detalle = Format$(Left$(rs22("descripcion"), 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
       obsp = Format$(Left$(rs22("stk_03.detalle"), 22), "@@@@@@@@@@@@@@@@@@@@@@!")
       RSet cant = Format$(rs22("cantidad"), "#####0.000")
       u = Format$(rs22("ubicacion") & " ", "@@!")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "   " & renglon & " " & basico & " " & Detalle & " " & cant & " " & obsp & " " & u
       r = r + 1
       rs22.MoveNext
   Wend
   If Y < c Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.NewPage
      rs22.MoveFirst
   End If
 Next Y
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.EndDoc
End If
Else
 MsgBox ("COMPROBANTE INEXISTENTE")
End If

End Sub


