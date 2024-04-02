Attribute VB_Name = "mo_facte"
'factura electronica
Public WSAA As Object, WSFEv1 As Object


Public Function fe_genera_wsaa() As Boolean
'genera un tiquet wsaa
'devuelve true la renovacion es correcta, si la renovacion da error devuelve false

Dim respuesta As Boolean

respuesta = True

Set WSAA = CreateObject("WSAA")
WSAA.LanzarExcepciones = False
        
 
 
' Generar un Ticket de Requerimiento de Acceso (TRA) para WSFEv1
ttl = 36000 ' tiempo de vida = 10hs hasta expiración
tra = WSAA.CreateTRA("wsfe", ttl)
    
' Especificar la ubicacion de los archivos certificado y clave privada

If para.empresa <> "" Then
   Path = "c:\" & para.empresa & "\5a04\seg\"
Else
   Path = "c:\5a04\seg\"
End If
  
' Generar el mensaje firmado (CMS)
cms = WSAA.SignTRA(tra, Path + para.facte_certificado, Path + para.facte_claveprivada)
ControlarExcepcion WSAA
Debug.Print cms
    
    ' Conectarse con el webservice de autenticación:
    cache = ""
    proxy = "" '"usuario:clave@localhost:8000"
    wrapper = "" ' libreria http (httplib2, urllib2, pycurl)
    cacert = "" 'WSAA.InstallDir & "\conf\afip_ca_info.crt" ' certificado de la autoridad de certificante
    wsdl = para.facte_servidor_wsaa
    ok = WSAA.Conectar(cache, wsdl, proxy, wrapper, cacert)
    ControlarExcepcion WSAA
    
    ' Llamar al web service para autenticar:
    ta = WSAA.LoginCMS(cms)
    ControlarExcepcion WSAA

    ' Imprimir el ticket de acceso, ToKen y Sign de autorización
    Debug.Print ta
    Debug.Print "Token:", WSAA.Token
    Debug.Print "Sign:", WSAA.Sign
    
    If ta <> "" Then
        'wsaa existoso
        para.facte_sign = WSAA.Sign
        para.facte_token = WSAA.Token
        para.facte_expira = Now
        Set rs = New ADODB.Recordset
        q = "select * from fe_01  where id = 1"
        Set rs = New ADODB.Recordset
        rs.Open q, cn1, adOpenDynamic, adLockOptimistic
        If Not rs.EOF And Not rs.BOF Then
            rs("token") = para.facte_token
            rs("sign") = para.facte_sign
            rs("fecha_expira") = para.facte_expira
            rs.Update
        End If
        Set rs = Nothing
    Else
      respuesta = False
    End If
 

End Function


Public Function fe_valida_tique() As Boolean
'si el tique es valido(no expiró) devuelve True, en caso contarrio devuelve false
Dim d, r As Date
Dim d2 As Double
Dim respuesta As Boolean

d = CDate((Now))
r = CDate(para.facte_expira)
d2 = (d - r) * 24
If d2 > 11 Then
 'un tiquet dura 12 hs, si supera 11hs
 respuesta = False
Else
 respuesta = True
End If
fe_valida_tique = respuesta

End Function

Public Function fe_valida_wsfe() As Boolean

Dim respuesta As Boolean
respuesta = True

   ' Crear objeto interface Web Service de Factura Electrónica de Mercado Interno
    Set WSFEv1 = CreateObject("WSFEv1")
    
    ' Setear tocken y sing de autorización (pasos previos)
    WSFEv1.Token = para.facte_token
    WSFEv1.Sign = para.facte_sign
    
    ' CUIT del emisor (debe estar registrado en la AFIP)
    WSFEv1.CUIT = glo.CUIT
    
    ' deshabilito errores no manejados
    WSFEv1.LanzarExcepciones = False
    
    ' Conectar al Servicio Web de Facturación
    proxy = "" ' "usuario:clave@localhost:8000"
    wsdl = para.facte_servidor_wsfe
    cache = "" 'Path
    wrapper = "" ' libreria http (httplib2, urllib2, pycurl)
    cacert = "" 'WSAA.InstallDir & "\conf\afip_ca_info.crt" ' certificado de la autoridad de certificante (solo pycurl)
    
    ok = WSFEv1.Conectar(cache, wsdl, proxy, wrapper, cacert)
    'ControlarExcepcion WSFEv1
    
    
    ' Llamo a un servicio nulo, para obtener el estado del servidor (opcional)
    WSFEv1.Dummy
    ControlarExcepcion WSFEv1
    Debug.Print "appserver status", WSFEv1.AppServerStatus
    Debug.Print "dbserver status", WSFEv1.DbServerStatus
    Debug.Print "authserver status", WSFEv1.AuthServerStatus

    If (WSFEv1.AppServerStatus = "OK" And WSFEv1.DbServerStatus = "OK" And WSFEv1.AuthServerStatus = "OK") Then
         respuesta = True
    Else
         respuesta = False
    End If
    
    fe_valida_wsfe = respuesta

End Function





Sub ControlarExcepcion(obj As Object)
    ' Nueva funcion para verificar que no haya habido errores:
    On Error GoTo 0
    If obj.Excepcion <> "" Then
        ' Depuración (grabar a un archivo los detalles del error)
        fd = FreeFile
        Open "c:\5a04\log\excepcion.txt" For Append As fd
        Print #fd, obj.Excepcion
        Print #fd, obj.Traceback
        Print #fd, obj.XmlRequest
        Print #fd, obj.XmlResponse
        Close fd
        MsgBox obj.Excepcion, vbExclamation, "Excepción"
        End
    End If
End Sub
