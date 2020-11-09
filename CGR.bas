Attribute VB_Name = "mo_cgr"
Public Function saldocuentacgr(ByVal c As Long, f As String) As Double
'saca el saldo de una cuenta a una fecha
Set rsg = New ADODB.Recordset
q = "select * from c_02, c_03 where c_02.[num_interno] = c_03.[num_interno] and datevalue([fecha]) <= datevalue('" & f & "') and [id_cuenta] = " & c
rsg.Open q, cn1
sc = 0
While Not rsg.EOF
  If rsg("ubicacion") = "D" Then
    sc = sc + rsg("importe")
  Else
     sc = sc - rsg("importe")
  End If
   rsg.MoveNext
Wend
Set rsg = Nothing
saldocuentacgr = sc
End Function


Public Function Generaasientosauto() As Boolean
'devuelve verdadero si graba asentos auto o falso si no
Dim v As Boolean
v = True
Set rsg = New ADODB.Recordset
q = "select * from G0 "
rsg.Open q, cn1
If Not rsg.EOF And Not rsg.BOF Then
  v = rsg("graba_asientos_auto")
End If
Set rsg = Nothing
Generaasientosauto = v

End Function
