Attribute VB_Name = "Module_ACTU"
Public cn1 As ADODB.Connection 'conexion a base de datos
Public rs As Recordset

Public Function abrirconexion(ByVal o As Integer) As Boolean    'proc. que abre la conexion con la base de datos
  'u usuario
  ' password
  abrirconexion = False
  On Error GoTo manerr
  Set cn1 = New ADODB.Connection
  If o = 1 Then
     gconexion = "Provider=Microsoft.Jet.oledb.4.0;Data Source=" & App.Path & "\dat\5a04.mdb;User id=claudio;password=0969" & ";" & "Jet OLEDB:System database=" & App.Path & "\SEG\system1.mdw;"
  Else
        gconexion = "Provider=Microsoft.Jet.oledb.4.0;Data Source=" & App.Path & "\dat\5a04.mdb;User id=claudio;password=0969" & ";" & "Jet OLEDB:System database=" & App.Path & "\SEG\system2.mdw;"
  End If
 ' (sql) gconexion = "Provider=SQLOLEDB; Initial Catalog=5a04sql; Data Source=(local)\SQL5A04; integrated security=SSPI; persist security info=True;"
 
  
  cn1.Open gconexion
  
  abrirconexion = True
 
  
  
  Exit Function


manerr:
   MsgBox ("Error al Abrir Base de Datos, Verifique su Usuario y Password")
   abrirconexion = False
   End
End Function

