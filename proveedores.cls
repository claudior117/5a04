VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "proveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
'variables locales para almacenar los valores de las propiedades
Private mvaridprov As Long 'copia local
Private mvarrazonsocial As String 'copia local
Private mvardireccion As String 'copia local
Private mvarte As String 'copia local
Private mvarlocalidad As String 'copia local
Private mvarcp As String 'copia local
Private mvarprovincia As String 'copia local
Private mvaremail As String 'copia local
Private mvarcuit As String 'copia local
Private mvarcodtipoiva As Integer 'copia local
Private mvaridcodretgan As Integer 'copia local
Private mvarinscriptogan As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarletra As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarretieneib As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarnumib As String 'copia local
Private mvaridtipoib As Single 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarfechavtoexepcionib As Date 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarSaldoprov As Double 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarDEBE As Double 'copia local
Private mvarHABER As Double 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarcontacto As String 'copia local
Private mvartecontacto As String 'copia local
Private mvaremailcontacto As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvaridcodretib As Integer 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvaralicuotaretib As Single 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarcalcularetib As String 'copia local
Private mvarcalcularetgan As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvaridcuenta As Variant 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarsaldoprovv As Double 'copia local
Public Property Let saldoprovv(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.saldoprovv = 5
    mvarsaldoprovv = vData
End Property


Public Property Get saldoprovv() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.saldoprovv
    saldoprovv = mvarsaldoprovv
End Property



Public Property Let idcuenta(ByVal vData As Variant)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.idcuenta = 5
    mvaridcuenta = vData
End Property


Public Property Set idcuenta(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.idcuenta = Form1
    Set mvaridcuenta = vData
End Property


Public Property Get idcuenta() As Variant
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.idcuenta
    If IsObject(mvaridcuenta) Then
        Set idcuenta = mvaridcuenta
    Else
        idcuenta = mvaridcuenta
    End If
End Property



Public Property Let calcularetgan(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.calcularetgan = 5
    mvarcalcularetgan = vData
End Property


Public Property Get calcularetgan() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.calcularetgan
    calcularetgan = mvarcalcularetgan
End Property



Public Property Let calcularetib(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.calcularetib = 5
    mvarcalcularetib = vData
End Property


Public Property Get calcularetib() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.calcularetib
    calcularetib = mvarcalcularetib
End Property



Public Property Let alicuotaretib(ByVal vData As Single)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.alicuotaretib = 5
    mvaralicuotaretib = vData
End Property


Public Property Get alicuotaretib() As Single
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.alicuotaretib
    alicuotaretib = mvaralicuotaretib
End Property







Public Property Let idcodretib(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.idcodretib = 5
    mvaridcodretib = vData
End Property


Public Property Get idcodretib() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.idcodretib
    idcodretib = mvaridcodretib
End Property



Public Property Let emailcontacto(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.emailcontacto = 5
    mvaremailcontacto = vData
End Property


Public Property Get emailcontacto() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.emailcontacto
    emailcontacto = mvaremailcontacto
End Property



Public Property Let tecontacto(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.tecontacto = 5
    mvartecontacto = vData
End Property


Public Property Get tecontacto() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.tecontacto
    tecontacto = mvartecontacto
End Property



Public Property Let contacto(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.contacto = 5
    mvarcontacto = vData
End Property


Public Property Get contacto() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.contacto
    contacto = mvarcontacto
End Property



Public Property Let HABER(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.haber = 5
    mvarHABER = vData
End Property


Public Property Get HABER() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.haber
    HABER = mvarHABER
End Property



Public Property Let DEBE(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.debe = 5
    mvarDEBE = vData
End Property


Public Property Get DEBE() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.debe
    DEBE = mvarDEBE
End Property



Public Property Let Saldoprov(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Saldoprov = 5
    mvarSaldoprov = vData
End Property


Public Property Get Saldoprov() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Saldoprov
    Saldoprov = mvarSaldoprov
End Property



Public Property Let fechavtoexepcionib(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.fechavtoexepcionib = 5
    mvarfechavtoexepcionib = vData
End Property


Public Property Get fechavtoexepcionib() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.fechavtoexepcionib
    fechavtoexepcionib = mvarfechavtoexepcionib
End Property



Public Property Let idtipoib(ByVal vData As Single)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.idtipoib = 5
    mvaridtipoib = vData
End Property


Public Property Get idtipoib() As Single
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.idtipoib
    idtipoib = mvaridtipoib
End Property



Public Property Let numib(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.numib = 5
    mvarnumib = vData
End Property


Public Property Get numib() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.numib
    numib = mvarnumib
End Property



Public Property Let retieneib(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.retieneib = 5
    mvarretieneib = vData
End Property


Public Property Get retieneib() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.retieneib
    retieneib = mvarretieneib
End Property



Public Property Let letra(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.letra = 5
    mvarletra = vData
End Property


Public Property Get letra() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.letra
    letra = mvarletra
End Property



Public Sub carga(ByVal ip As Long)
'On Error GoTo er_carga
q = "select [id_proveedor], [denominacion], [direccion], [localidad], [provincia], [cp], [cuit], [email], [te], [id_codretgan], [inscripto_gan], [fecha_vto_exepcion_ib], [num_ib], [contacto], [te_contacto], [email_contacto], [id_codretib], a1.[cod_tipoiva], a1.[id_tipoib], [alicuota_retib], [id_cuenta_a1], [calcula_ret_ib], [calcula_ret_gan], [letra_prov], [retiene]    from a1, g3, g6 where a1.[id_proveedor] = " & ip & " and a1.[cod_tipoiva] = g3.[cod_tipoiva]  and a1.[id_tipoib] = g6.[id_tipoib] "
Set rs = New adodb.Recordset
rs.MaxRecords = 1
rs.Open q, cn1
If Not rs.BOF And Not rs.EOF Then
  Me.idprov = rs("id_proveedor")
  Me.razonsocial = rs("denominacion")
  Me.direccion = rs("direccion")
  Me.localidad = rs("localidad")
  Me.provincia = rs("provincia")
  Me.cp = rs("cp")
  Me.CUIT = rs("cuit")
  Me.email = rs("email")
  Me.te = rs("te")
  Me.codtipoiva = rs("cod_tipoiva")
  Me.idcodretgan = rs("id_codretgan")
  Me.inscriptogan = rs("inscripto_gan")
  Me.letra = rs("letra_prov")
  Me.retieneib = rs("retiene")
  Me.idtipoib = rs("id_tipoib")
  Me.fechavtoexepcionib = rs("fecha_vto_exepcion_ib")
  Me.numib = rs("num_ib")
  Me.contacto = rs("contacto")
  Me.tecontacto = rs("te_contacto")
  Me.emailcontacto = rs("email_contacto")
  Me.idcodretib = rs("id_codretib")
  Me.alicuotaretib = rs("alicuota_retib")
  If rs("calcula_ret_ib") = "S" And rs("id_codretib") > 0 Then
      Me.calcularetib = "S"
  Else
      Me.calcularetib = "N"
  End If
  Me.calcularetgan = rs("calcula_ret_gan")
  Me.idcuenta = rs("id_cuenta_a1")
Else
  Me.idprov = 0
End If
Set rs = Nothing
Exit Sub
 
er_carga:
  MsgBox ("Error. Clase:Proveedores  Modulo:Carga")
  Me.idprov = 0
  Exit Sub
End Sub

Public Property Let inscriptogan(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.inscriptogan = 5
    mvarinscriptogan = vData
End Property


Public Property Get inscriptogan() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.inscriptogan
    inscriptogan = mvarinscriptogan
End Property



Public Property Let idcodretgan(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.idcodretgan = 5
    mvaridcodretgan = vData
End Property


Public Property Get idcodretgan() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.idcodretgan
    idcodretgan = mvaridcodretgan
End Property



Public Property Let codtipoiva(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.codtipoiva = 5
    mvarcodtipoiva = vData
End Property


Public Property Get codtipoiva() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codtipoiva
    codtipoiva = mvarcodtipoiva
End Property



Public Property Let CUIT(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.cuit = 5
    mvarcuit = vData
End Property


Public Property Get CUIT() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.cuit
    CUIT = mvarcuit
End Property



Public Property Let email(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.email = 5
    mvaremail = vData
End Property


Public Property Get email() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.email
    email = mvaremail
End Property



Public Property Let provincia(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.provincia = 5
    mvarprovincia = vData
End Property


Public Property Get provincia() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.provincia
    provincia = mvarprovincia
End Property



Public Property Let cp(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.cp = 5
    mvarcp = vData
End Property


Public Property Get cp() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.cp
    cp = mvarcp
End Property



Public Property Let localidad(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.localidad = 5
    mvarlocalidad = vData
End Property


Public Property Get localidad() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.localidad
    localidad = mvarlocalidad
End Property



Public Property Let te(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.te = 5
    mvarte = vData
End Property


Public Property Get te() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.te
    te = mvarte
End Property



Public Property Let direccion(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.direccion = 5
    mvardireccion = vData
End Property


Public Property Get direccion() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.direccion
    direccion = mvardireccion
End Property



Public Property Let razonsocial(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.razonsocial = 5
    mvarrazonsocial = vData
End Property


Public Property Get razonsocial() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.razonsocial
    razonsocial = mvarrazonsocial
End Property



Public Property Let idprov(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.idprov = 5
    mvaridprov = vData
End Property


Public Property Get idprov() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.idprov
    idprov = mvaridprov
End Property

Public Function saldo(menoreigual As Boolean, F As String, p As Boolean, zona As Integer) As Double
    'busca el saldo de un prov cc a una fecha f en una moneda p(true pesos)
    'agrega a la propiedad debe y haber los valores correspondientes
    'zona  Todas = 0  Zona 1 = 1 Zona 2 = 2
    On Error GoTo err2
    Dim QUERY As String
    'Dim rs As ADODB.Recordset
    Dim saldoant As Double
    Dim tp As Double
    Dim td As Double
 
    Set rs = New adodb.Recordset
    If menoreigual Then
       QUERY = "SELECT * FROM A5 where [id_proveedor] = " & idprov
       QUERY = QUERY & " and datevalue(fecha) <= " & "DateValue('" & F & "') "
       QUERY = QUERY & " and [ctacte] <> " & "'N' AND [CONTADO] = 'N'"
    Else
       QUERY = "SELECT * FROM A5 where [id_proveedor] = " & idprov
       QUERY = QUERY & " and datevalue(fecha) < " & "DateValue('" & F & "') "
       QUERY = QUERY & " and [ctacte] <> " & "'N' AND [CONTADO] = 'N'"
    End If
    
    If zona > 0 Then
      QUERY = QUERY & " and zona = " & zona
    End If
    
    rs.Open QUERY, cn1, , , 1
    saldoant = 0
    While Not rs.EOF
       
    If rs("ctacte") = "D" Then
      If rs("moneda") = "P" Then
         tp = rs("total")
         td = rs("total_d")
      Else
         tp = rs("total_d")
         td = rs("total")
      End If
    Else
     If rs("moneda") = "P" Then
       tp = -rs("total")
       td = -rs("total_d")
     Else
       tp = -rs("total_d")
       td = -rs("total")
     End If
    End If
   
   If p Then
         saldoant = saldoant + tp
   Else
       saldoant = saldoant + td
   End If
   rs.MoveNext
Wend
Set rs = Nothing
saldo = saldoant
Saldoprov = saldoant

Exit Function
err2:

MsgBox ("Error al generar los saldos. -Class Clientes Module:Saldo-")
saldo = 0
saldocli = 0
Exit Function
End Function

Public Function saldov(menoreigual As Boolean, F As String, p As Boolean, zona As Integer) As Double
    'busca el saldo de un prov cc a una fecha f en una moneda p(true pesos)
    'agrega a la propiedad debe y haber los valores correspondientes
    'zona  Todas = 0  Zona 1 = 1 Zona 2 = 2
    On Error GoTo err2
    Dim QUERY As String
    Dim saldoant As Double
    Dim tp As Double
    Dim td As Double
 
    Set rs = New adodb.Recordset
    If menoreigual Then
       QUERY = "SELECT * FROM A5 where [id_proveedor] = " & idprov
       QUERY = QUERY & " and datevalue([fecha_vto]) <= " & "DateValue('" & F & "') "
       QUERY = QUERY & " and [ctacte] <> " & "'N' AND [CONTADO] = 'N'"
    Else
       QUERY = "SELECT * FROM A5 where [id_proveedor] = " & idprov
       QUERY = QUERY & " and datevalue([fecha_vto]) < " & "DateValue('" & F & "') "
       QUERY = QUERY & " and [ctacte] <> " & "'N' AND [CONTADO] = 'N'"
    End If
    
    If zona > 0 Then
      QUERY = QUERY & " and zona = " & zona
    End If
    
    rs.Open QUERY, cn1
    
    saldoant = 0
    While Not rs.EOF
       
    If rs("ctacte") = "D" Then
      If rs("moneda") = "P" Then
         tp = rs("total")
         td = rs("total_d")
      Else
         tp = rs("total_d")
         td = rs("total")
      End If
    Else
     If rs("moneda") = "P" Then
       tp = -rs("total")
       td = -rs("total_d")
     Else
       tp = -rs("total_d")
       td = -rs("total")
     End If
    End If
   
   If p Then
         saldoant = saldoant + tp
   Else
       saldoant = saldoant + td
   End If
   
   rs.MoveNext
Wend
Set rs = Nothing
saldov = saldoant
saldoprovv = saldoant

Exit Function
err2:

MsgBox ("Error al generar los saldos. -Class Clientes Module:Saldo-")
saldov = 0
saldoprovv = 0
Exit Function
End Function


Public Function saldoentrefechas(f1 As String, f2 As String, p As Boolean, zona As Integer) As Double
    'busca el saldo de un cliente cc a entre f1 y f2 en una moneda p(true pesos)
    'agrega a la propiedad debe y haber los valores correspondientes
    'zona = 0  todas o numero de zona
    
    On Error GoTo err2
    Dim QUERY As String
    'Dim rs As ADODB.Recordset
    Dim saldoant As Double
    Dim saldoact As Double
    Dim tp As Double
    Dim td As Double
    Dim hp As Double 'haber
    Dim hd As Double
    Dim dp As Double 'debe
    Dim dd As Double
    Dim totaldebe As Double 'debe
    Dim totalhaber As Double

    saldoant = saldo(False, f1, p, zona)
    
    Set rs = New adodb.Recordset
    QUERY = "SELECT * FROM A5 where [id_proveedor] = " & idprov
    QUERY = QUERY & " and datevalue(fecha) >= " & "DateValue('" & f1 & "') "
    QUERY = QUERY & " and datevalue(fecha) <= " & "DateValue('" & f2 & "') "
    QUERY = QUERY & " and  [ctacte] <> " & "'N' and [contado] = 'N'"
    If zona > 0 Then
      QUERY = QUERY & " and zona = " & zona
    End If
    rs.Open QUERY, cn1
    totaldebe = 0
    totalhaber = 0
    saldoact = saldoant
    While Not rs.EOF
      dp = 0
      dd = 0
      hp = 0
      hd = 0
      If rs("ctacte") = "D" Then
       If rs("moneda") = "P" Then
         tp = rs("total")
         td = rs("total_d")
         dp = rs("total")
         dd = rs("total_d")
       Else
         tp = rs("total_d")
         td = rs("total")
         dp = rs("total_d")
         dd = rs("total")
      End If
     Else
      If rs("moneda") = "P" Then
        tp = -rs("total")
        td = -rs("total_d")
        hp = -rs("total")
        hd = -rs("total_d")
      
      Else
        tp = -rs("total_d")
        td = -rs("total")
         hp = -rs("total_d")
         hd = -rs("total")
    
      End If
     End If

      
   
      If p Then
         saldoact = saldoact + tp
         totaldebe = totaldebe + dp
         totalhaber = totalhaber + hp
      Else
         saldoact = saldoact + td
         totaldebe = totaldebe + dd
         totalhaber = totalhaber + hd
      
      End If
   rs.MoveNext
Wend
Set rs = Nothing
Saldoprov = saldoact
HABER = totalhaber
DEBE = totaldebe
saldoentrefechas = saldoact
Exit Function

err2:

MsgBox ("Error al generar los saldos. -Class Proveedores Module:Saldoentrefechas-")
Saldoprov = 0
HABER = 0
DEBE = 0
saldoentrefechas = 0
Exit Function

End Function

Public Function saldoentrefechasv(f1 As String, f2 As String, p As Boolean, zona As Integer) As Double
    'busca el saldo de un cliente cc a entre f1 y f2 en una moneda p(true pesos)
    'agrega a la propiedad debe y haber los valores correspondientes
    'zona = 0  todas o numero de zona
    
    On Error GoTo err2
    Dim QUERY As String
    'Dim rs As ADODB.Recordset
    Dim saldoant As Double
    Dim saldoact As Double
    Dim tp As Double
    Dim td As Double
    Dim hp As Double 'haber
    Dim hd As Double
    Dim dp As Double 'debe
    Dim dd As Double
    Dim totaldebe As Double 'debe
    Dim totalhaber As Double

    saldoant = saldo(False, f1, p, zona)
    
    Set rs = New adodb.Recordset
    QUERY = "SELECT * FROM A5 where [id_proveedor] = " & idprov
    QUERY = QUERY & " and datevalue([fecha_vto]) >= " & "DateValue('" & f1 & "') "
    QUERY = QUERY & " and datevalue([fecha_vto]) <= " & "DateValue('" & f2 & "') "
    QUERY = QUERY & " and  [ctacte] <> " & "'N' and [contado] = 'N'"
    If zona > 0 Then
      QUERY = QUERY & " and zona = " & zona
    End If
    rs.Open QUERY, cn1, , , 1
    totaldebe = 0
    totalhaber = 0
    saldoact = saldoant
    While Not rs.EOF
      dp = 0
      dd = 0
      hp = 0
      hd = 0
      If rs("ctacte") = "D" Then
       If rs("moneda") = "P" Then
         tp = rs("total")
         td = rs("total_d")
         dp = rs("total")
         dd = rs("total_d")
       Else
         tp = rs("total_d")
         td = rs("total")
         dp = rs("total_d")
         dd = rs("total")
      End If
     Else
      If rs("moneda") = "P" Then
        tp = -rs("total")
        td = -rs("total_d")
        hp = -rs("total")
        hd = -rs("total_d")
      
      Else
        tp = -rs("total_d")
        td = -rs("total")
         hp = -rs("total_d")
         hd = -rs("total")
    
      End If
     End If

      
   
      If p Then
         saldoact = saldoact + tp
         totaldebe = totaldebe + dp
         totalhaber = totalhaber + hp
      Else
         saldoact = saldoact + td
         totaldebe = totaldebe + dd
         totalhaber = totalhaber + hd
      
      End If
   rs.MoveNext
Wend
Set rs = Nothing
Me.saldoprovv = saldoact
Me.HABER = totalhaber
Me.DEBE = totaldebe
saldoentrefechasv = saldoact
Exit Function

err2:

MsgBox ("Error al generar los saldos por vencimientos. -Class Proveedores Module:Saldoentrefechasv-")
Me.Saldoprov = 0
Me.HABER = 0
Me.DEBE = 0
saldoentrefechasv = 0
Exit Function

End Function

