Attribute VB_Name = "mo_vta"
Sub verifica_tasa_iva(ByVal ni)
'verifica y corrige los acumulados de tasa de iva para cada comprobante
Set rs1 = New ADODB.Recordset
q = "select * from vta_09 where [num_int] = " & ni
rs1.Open q, cn1, adOpenStatic, adLockOptimistic
While Not rs1.EOF
    espere.Label1 = "Espere... Borrando Totales" & a
    espere.Label1.Refresh
    a = a + 1
    rs1.Delete
    rs1.MoveNext
Wend
Set rs1 = Nothing
 
 
qm = "select * from vta_02 where  [num_int] = " & ni
Set rs = New ADODB.Recordset
rs.Open qm, cn1
If Not rs.EOF And Not rs.BOF Then
  'verifico tasa iguales
  Set rs1 = New ADODB.Recordset
  q = "select * from vta_03 where [num_int] = " & rs("num_int")
  rs1.Open q, cn1
  ti = 1
  p = 0
  tasa = 21
  While Not rs1.EOF
     If p = 0 Then
       tasa = rs1("tasaiva")
       p = 1
     End If
     If tasa <> rs1("tasaiva") Then
        ti = 0
        rs1.MoveNext
     Else
       rs1.MoveNext
     End If
  Wend
  Set rs1 = Nothing
  If ti = 1 Then
    'tasunica
    cn1.BeginTrans
    QUERY = "INSERT INTO vta_09([num_int], [tasa_iva], [iva], [neto], [tipo_iva], [id_cuenta09])"
    QUERY = QUERY & " VALUES (" & rs("num_int") & ", " & tasa & ", " & rs("iva") & ", " & rs("subtotal") & ", " & rs("id_tipo_iva02") & ", " & rs("id_cuenta") & ")"
    cn1.Execute QUERY
    cn1.CommitTrans
  Else
   Set rs1 = New ADODB.Recordset
   q = "select * from vta_03 where [num_int] = " & rs("num_int") & " order by [tasaiva]"
   rs1.Open q, cn1
   p = 0
   su = 0
   iv = 0
   While Not rs1.EOF
     If p = 0 Then
       tasa = rs1("tasaiva")
       p = 1
     End If
     If tasa <> rs1("tasaiva") Then
        cn1.BeginTrans
        QUERY = "INSERT INTO vta_09([num_int], [tasa_iva], [iva], [neto], [tipo_iva], [id_cuenta09])"
        QUERY = QUERY & " VALUES (" & rs("num_int") & ", " & tasa & ", " & rs("iva") & ", " & rs("subtotal") & ", " & rs("id_tipo_iva02") & ", " & rs("id_cuenta") & ")"
        cn1.Execute QUERY
        cn1.CommitTrans
        tasa = rs1("tasaiva")
        su = 0
        iv = 0
     End If
     s = rs1("cantidad") * rs1("pu")
     su = su + s
     i = s * (rs1("tasaiva") / 100)
     iv = iv + Format(i, "#####0.00")
     rs1.MoveNext
   Wend
   cn1.BeginTrans
   QUERY = "INSERT INTO vta_09([num_int], [tasa_iva], [iva], [neto], [tipo_iva], [id_cuenta09])"
   QUERY = QUERY & " VALUES (" & rs("num_int") & ", " & tasa & ", " & rs("iva") & ", " & rs("subtotal") & ", " & rs("id_tipo_iva02") & ", " & rs("id_cuenta") & ")"
   cn1.Execute QUERY
   cn1.CommitTrans
   Set rs1 = Nothing
  End If
End If
End Sub


Sub verificaerrfiscal(ByVal ef As String, ByVal ei As String)
'trata de identificar el err fiscal y el de la impresora
Select Case ef
 Case Is = "8E20", Is = "8E00"
   MsgBox ("Error FISCAL(" & ef & "): es necesario realizar un nuevo CIERE Z")
 Case Is = "B610"
     MsgBox ("Error FISCAL(B610): Datos Incorrectos. Verifique el CUIT y la Razon Social, en caso de ser NC/ND verifique adicionalmente la Direccion. Estos campos no pueden estar vacios ni contener caracteres especiales(solo letras y numeros)")
 Case Is = "B620"
     MsgBox ("Error FISCAL(B620): Documento Fiscal Abierto. Apague y vuelva a encender la impresora")
 Case Is = "8610"
     MsgBox ("Error FISCAL(8610): Datos Incorrectos. Verifique el CUIT y la Razon Social, en caso de ser NC/ND verifique adicionalmente la Direccion. Estos campos no pueden estar vacios ni contener caracteres especiales(solo letras y numeros)")
 
 Case Is = "8620"
     MsgBox ("Error FISCAL(8620): Comando no valido, puede ser un problema en la memoria fiscal.  Apague y vuelva a encender la impresora, si el prblema persigue llame al servicio tecnico")
 Case Is = "8E00"
     'todo normal
 Case Is = "0600"
     MsgBox ("Estado Fiscal OK")
 Case Else
     MsgBox ("Error no identificado (" & ef & ")")
 End Select
 
 Select Case ei
 Case Is = "8610"
   MsgBox ("Error IMPRESOR(8610): Verifique el papel")
 Case Is = "80A0"
     MsgBox ("Error IMPRESOR(80A0): Sin papel")
  Case Is = "0080"
     'todo normal
     MsgBox ("Estado Impresor OK")
 Case Else
     MsgBox ("Error no identificado (" & ei & ")")
 End Select

End Sub

Public Function textofiscal(ByVal tf As String) As String


 t = ""
 For i = 1 To Len(tf)
    l = Mid$(tf, i, 1)
    If Asc(l) < 32 Or Asc(l) > 127 Then
       l = "#"
    End If
    t = t & l
 Next i

If Len(t) <= 0 Then
  t = " "
End If
textofiscal = t
End Function


Public Function sacacbu() As String
q = "select cbu from g0 where  [sucursal] = 0"
Set rs = New ADODB.Recordset
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
  sacacbu = rs("cbu")
Else
  sacacbu = " "
End If
Set rs = Nothing
End Function


