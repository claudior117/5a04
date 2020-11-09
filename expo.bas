Attribute VB_Name = "mo_expo"
Public Function sacareintegro(ByVal ni As Long)
'saca el total ingresado como reintegro de una exportacion
Set rse = New ADODB.Recordset
q = "select * from exp02 where [num_exp] = " & ni
rse.Open q, cn1
tr = 0
While Not rse.EOF
 tr = tr + (rse("cantidad") * rse("pusiva"))
 rse.MoveNext
Wend
Set rse = Nothing
sacareintegro = tr

End Function
