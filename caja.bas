Attribute VB_Name = "mo_caja"
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Function saldoanterior(ByVal F As Date, ByVal iu As Integer, ByVal cuenta As Long, Optional u As Long) As Double
 Set rs = New ADODB.Recordset
 q = "select * from cyb_05 where datevalue([fecha]) < datevalue('" & F & "')"
 If iu > 0 Then
   q = q & " and [id_usuario] = " & iu
 End If
 
 
 If cuenta > 0 Then
   q = q & " and [id_forma_pago] = " & cuenta
 End If
 
 'If IsMissing(u) = True Then
 '     q = q & " and [id_usuario] = " & u
 'End If
 
 rs.Open q, cn1
 sa = 0
 While Not rs.EOF
   If rs("ubicaCION") = "D" Then
     sa = sa + rs("importe")
   Else
     sa = sa - rs("importe")
   End If
   rs.MoveNext
 Wend
 saldoanterior = sa
 
End Function
Function estadocaja(ByVal F As String) As String
  Set rsx = New ADODB.Recordset
  q = "select * from cyb_09 where datevalue([fecha]) = datevalue('" & F & "')"
  rsx.MaxRecords = 1
  rsx.Open q, cn1
  
  If Not rsx.BOF And Not rsx.EOF Then
     estadocaja = rsx("estado")
  Else
     estadocaja = "A"
  End If
  Set rsx = Nothing
  
End Function

Function saldoalafecha(ByVal F As Date, ByVal iu As Integer, ByVal cuenta As Long, Optional u As Long) As Double
 'IU NUMRO USUARIO IU = 0 TODOS
 'parametros de busqueda en 0 no busca
 'rubro
 'cuenta
 'persona
 Set rs = New ADODB.Recordset
 q = "select * from cyb_05 where datevalue([fecha]) <= datevalue('" & F & "')"
 If iu > 0 Then
   q = q & " and [id_usuario] = " & iu
 End If
 
 
 If cuenta > 0 Then
   q = q & " and [id_forma_pago] = " & cuenta
 End If
 

'If IsMissing(u) Then
'      q = q & " and [id_usuario] = " & u
' End If

 
 rs.Open q, cn1
 sa = 0
 While Not rs.EOF
   If rs("ubicaCION") = "D" Then
     sa = sa + rs("importe")
   Else
     sa = sa - rs("importe")
   End If
   rs.MoveNext
 Wend
 saldoalafecha = sa
 
End Function

Sub borramovcaja(ByVal ni As Long)
'ni = id_movimiento
J = MsgBox("Confirma eliminar movimiento nro. " & Format$(ni, "000000"), 4)
If J = 6 Then
      q = "select * from cyb_05 where [num_mov_caja] = " & ni
      Set rs = New ADODB.Recordset
      rs.Open q, cn1, adOpenDynamic, adLockOptimistic
      If Not rs.EOF And Not rs.BOF Then
        If rs("modulo") = "J" Then
            If estadocaja(rs("fecha")) = "A" Then
              If verificaperiodog(rs("fecha")) = "A" Then
                 Call borracontabilidad2(rs("num_mov_caja"), "J")
                 
                 Set rs1 = New ADODB.Recordset
                 q = "select * from cyb_04 where [modulo] = 'J' and [num_mov_int] = " & rs("num_mov_caja")
                 rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
                 While Not rs1.EOF
                    rs1.Delete
                    rs1.MoveNext
                 Wend
                 Set rs1 = Nothing
                 
                 rs.Delete
                 rs.Update
                 
               Else
                 MsgBox ("Periodo Cerrado. Imposible realizar operacion")
               End If
            Else
               MsgBox ("Caja Cerrada. Imposible realizar operacion")
            End If
         Else
           MsgBox ("Movimiento no ingresado por caja. Para eliminarlo necesita Borrar el comprobante original")
         End If
      End If
End If




End Sub
