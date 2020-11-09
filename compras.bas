Attribute VB_Name = "mo_compras"
Public Function sacafechaultimopago(ByVal ni As Long) As String
'saca el ultimo pago que se le hizo a una factura a partir del num. int. de la factura
Set rs = New ADODB.Recordset
q = "select * from a15, a5 where [num_int_comp] = " & ni & " and [num_int_op] = [num_int]"
rs.Open q, cn1
ultni_op = 0
f = "01/01/2000"
While Not rs.EOF
  If DateValue(f) < rs("fecha") Then
     f = rs("fecha")
  End If
  rs.MoveNext
Wend
Set rs = Nothing
sacafechaultimopago = f
End Function


Public Function buscafacturaapocrifa(ByVal c As Double) As Boolean
'busca un cuit si se encuetra en el listado e facturas apocrifas del afip  PIB(FA)
'si lo encuentra devuelve true

Set rs = New ADODB.Recordset
q = "select * from fa where [cuit] = " & c
rs.Open q, cnib
If Not rs.EOF And Not rs.BOF Then
   buscafacturaapocrifa = True
Else
     buscafacturaapocrifa = False
End If
Set rs = Nothing



End Function
