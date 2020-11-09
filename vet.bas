Attribute VB_Name = "mo_vet"
Sub carga_tipo_mascota(c As ComboBox)
Set rs = New ADODB.Recordset
q = "select * from vet_03 order by [tipo]"
rs.Open q, cn1
c.clear
While Not rs.EOF
  c.AddItem rs("tipo")
  c.ItemData(c.NewIndex) = rs("id_tipo")
  rs.MoveNext
Wend

c.ListIndex = 0
Set rs = Nothing
End Sub


Sub carga_raza(c As ComboBox, ByVal t As Long)
Set rs = New ADODB.Recordset
q = "select * from vet_04 where [id_tipo] = " & t & " order by [raza]"
rs.Open q, cn1
c.clear
While Not rs.EOF
  c.AddItem rs("raza")
  c.ItemData(c.NewIndex) = rs("id_raza")
  rs.MoveNext
Wend

c.ListIndex = 0
Set rs = Nothing
End Sub

Sub carga_mascotas(c As ComboBox, ByVal t As Long)
'c es id cliente
On Error GoTo err2
Set rs = New ADODB.Recordset
q = "select * from vet_02 where [id_cliente] = " & t & " order by [nombre]"
rs.Open q, cn1
c.clear
While Not rs.EOF
  c.AddItem rs("nombre")
  c.ItemData(c.NewIndex) = rs("id_animal")
  rs.MoveNext
Wend

c.ListIndex = 0
Set rs = Nothing

Exit Sub
err2:
 Exit Sub
End Sub

