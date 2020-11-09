Attribute VB_Name = "mo_acp"
Sub carga_cereales(c As ComboBox)
Set rs = New ADODB.Recordset
q = "select [id_cereal], [cereal] from acp_02  order by [cereal]"
rs.Open q, cn1
c.clear

While Not rs.EOF
  c.AddItem rs("cereal")
  c.ItemData(c.NewIndex) = rs("id_cereal")
  rs.MoveNext
Wend

Set rs = Nothing

End Sub
