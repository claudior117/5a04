VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
'variables locales para almacenar los valores de las propiedades
Private mvarid As Long 'copia local
Private mvaridtipoiva As Integer 'copia local
Private mvarcp As String 'copia local
Private mvarcuit As String 'copia local
Private mvardireccion As String 'copia local
Private mvaremail As String 'copia local
Private mvarlocalidad As String 'copia local
Private mvarprovincia As String 'copia local
Private mvarrazonsocial As String 'copia local
Private mvarte As String 'copia local
Private mvarletra As String 'copia local
Private mvarabreviatura_tipoiva As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvardescrimina As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvaridvendedor As Long 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarDEBE As Double 'copia local
Private mvarHABER As Double 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarsaldocli As Double 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvaroperador_granos As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarperciveib As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvaridproveedor As Long 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarid_provincia As Integer 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarcodfiscal As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarsaldocliv As Double 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarlimitecredito As Double 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvardireccion_local As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarcodfiscal2 As Integer 'copia local
Public Property Let codfiscal2(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.codfiscal2 = 5
    mvarcodfiscal2 = vData
End Property


Public Property Get codfiscal2() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codfiscal2
    codfiscal2 = mvarcodfiscal2
End Property



Public Property Let direccion_local(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.direccion_local = 5
    mvardireccion_local = vData
End Property


Public Property Get direccion_local() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.direccion_local
    direccion_local = mvardireccion_local
End Property



Public Property Let limitecredito(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.limitecredito = 5
    mvarlimitecredito = vData
End Property


Public Property Get limitecredito() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.limitecredito
    limitecredito = mvarlimitecredito
End Property



Public Property Let saldocliv(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.saldocliv = 5
    mvarsaldocliv = vData
End Property


Public Property Get saldocliv() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.saldocliv
    saldocliv = mvarsaldocliv
End Property



Public Property Let codfiscal(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.codfiscal = 5
    mvarcodfiscal = vData
End Property


Public Property Get codfiscal() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codfiscal
    codfiscal = mvarcodfiscal
End Property



Public Property Let id_provincia(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.id_provincia = 5
    mvarid_provincia = vData
End Property


Public Property Get id_provincia() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.id_provincia
    id_provincia = mvarid_provincia
End Property



Public Property Let idproveedor(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.idproveedor = 5
    mvaridproveedor = vData
End Property


Public Property Get idproveedor() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.idproveedor
    idproveedor = mvaridproveedor
End Property



Public Property Let perciveib(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.perciveib = 5
    mvarperciveib = vData
End Property


Public Property Get perciveib() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.perciveib
    perciveib = mvarperciveib
End Property



Public Property Let operador_granos(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.operador_granos = 5
    mvaroperador_granos = vData
End Property


Public Property Get operador_granos() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.operador_granos
    operador_granos = mvaroperador_granos
End Property



Public Property Let saldocli(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.saldocli = 5
    mvarsaldocli = vData
End Property


Public Property Get saldocli() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.saldocli
    saldocli = mvarsaldocli
End Property



Public Property Let HABER(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.HABER = 5
    mvarHABER = vData
End Property


Public Property Get HABER() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.HABER
    HABER = mvarHABER
End Property



Public Property Let DEBE(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.DEBE = 5
    mvarDEBE = vData
End Property


Public Property Get DEBE() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.DEBE
    DEBE = mvarDEBE
End Property



Public Function busca_saldos(c As Long, m As String, F As Date) As Double
    'busca saldos  c = cod. cliente   m = moneda p pesos d dolares f = fecha hasta
    q = "select cta_cte, total from vta_02 where [id_cliente] = " & c & " and [cta_cte] <> 'N' " & " and [contado] = " & "'N'"
    q = q & " and datevalue([fecha]) < datevalue('" & F & "')"
    Set rs = New ADODB.Recordset
    
    rs.Open q, cn1
    da = 0
    ha = 0
    sa = 0
    While Not rs.EOF
     If rs("cta_cte") = "D" Then
        da = da + rs("total")
     Else
        ha = ha + rs("total")
     End If
     rs.MoveNext
    Wend
    sa = da - ha
    busca_saldos = sa

End Function
Public Function saldo(menoreigual As Boolean, F As String, p As Boolean) As Double
    'busca el saldo de un CLIENTE cc a una fecha f en una moneda p(true pesos)
    'agrega a la propiedad debe y haber los valores correspondientes
    'On Error GoTo err2
    Dim QUERY As String
    'Dim rs As ADODB.Recordset
    Dim saldoant As Double
    Dim tp As Double
    Dim td As Double
 
    Set rs = New ADODB.Recordset
    If menoreigual Then
       QUERY = "SELECT cta_cte, moneda, total, total_otra_moneda, id_tipocomp, estado_pago FROM VTA_02 where [id_CLIENTE] = " & Me.id
       QUERY = QUERY & " and datevalue(fecha) <= " & "DateValue('" & F & "') "
       QUERY = QUERY & " and [cta_cte] <> " & "'N'" & " and [contado] = " & "'N'"
    Else
       QUERY = "SELECT cta_cte, moneda, total, total_otra_moneda, id_tipocomp, estado_pago FROM VTA_02 where [id_CLIENTE] = " & Me.id
       QUERY = QUERY & " and datevalue(fecha) < " & "DateValue('" & F & "') "
       QUERY = QUERY & " and [cta_cte] <> " & "'N'" & " and [contado] = " & "'N'"
    End If
    
    rs.Open QUERY, cn1, , , 1
    saldoant = 0
    While Not rs.EOF
       
      If rs("cta_cte") = "D" Then
       If rs("moneda") = "P" Then
          tp = rs("total")
          td = rs("total_otra_moneda")
       Else
          td = rs("total")
          tp = rs("total_otra_moneda")
       End If
      Else
       If rs("moneda") = "P" Then
          tp = -rs("total")
          td = -rs("total_otra_moneda")
       Else
          td = -rs("total")
          tp = -rs("total_otra_moneda")
       End If
     End If
       If Not rs("id_tipocomp") = 25 Then
          If p Then
            saldoant = saldoant + tp
          Else
            saldoant = saldoant + td
          End If
       Else
            If rs("estado_pago") = "N" Then
               If p Then
                    saldoant = saldoant + tp
               Else
                    saldoant = saldoant + td
                End If
            End If
       End If
      
   rs.MoveNext
Wend
Set rs = Nothing
Me.saldocli = saldoant
saldo = saldoant
Exit Function
err2:

MsgBox ("Error al generar los saldos. -Class Clientes Module:Saldo-")
Me.saldocli = 0
saldo = 0
Exit Function
End Function

Public Function saldov(menoreigual As Boolean, F As String, p As Boolean) As Double
    'busca el saldo de un CLIENTE cc a una fecha f en una moneda p(true pesos)
    'agrega a la propiedad debe y haber los valores correspondientes
    'On Error GoTo err2
    Dim QUERY As String
    'Dim rs As ADODB.Recordset
    Dim saldoant As Double
    Dim tp As Double
    Dim td As Double
 
    Set rs = New ADODB.Recordset
    If menoreigual Then
       QUERY = "SELECT cta_cte, moneda, total, total_otra_moneda, id_tipocomp, estado_pago  FROM VTA_02 where [id_CLIENTE] = " & Me.id
       QUERY = QUERY & " and datevalue([fecha_vto]) <= " & "DateValue('" & F & "') "
       QUERY = QUERY & " and [cta_cte] <> " & "'N'" & " and [contado] = " & "'N'"
    Else
       QUERY = "SELECT cta_cte, moneda, total, total_otra_moneda, id_tipocomp, estado_pago  FROM VTA_02 where [id_CLIENTE] = " & Me.id
       QUERY = QUERY & " and datevalue([fecha_vto]) < " & "DateValue('" & F & "') "
       QUERY = QUERY & " and [cta_cte] <> " & "'N'" & " and [contado] = " & "'N'"
    End If
    
    rs.Open QUERY, cn1, , , 1
    saldoant = 0
    While Not rs.EOF
       
      If rs("cta_cte") = "D" Then
       If rs("moneda") = "P" Then
          tp = rs("total")
          td = rs("total_otra_moneda")
       Else
          td = rs("total")
          tp = rs("total_otra_moneda")
       End If
      Else
       If rs("moneda") = "P" Then
          tp = -rs("total")
          td = -rs("total_otra_moneda")
       Else
          td = -rs("total")
          tp = -rs("total_otra_moneda")
       End If
     End If
   
   If Not rs("id_tipocomp") = 25 Then
          If p Then
            saldoant = saldoant + tp
          Else
            saldoant = saldoant + td
          End If
       Else
            If rs("estado_pago") = "N" Then
               If p Then
                    saldoant = saldoant + tp
               Else
                    saldoant = saldoant + td
                End If
            End If
       End If
      
   rs.MoveNext
Wend
Set rs = Nothing
Me.saldocliv = saldoant
saldov = saldoant
Exit Function
err2:

MsgBox ("Error al generar los saldos por vencimientos. -Class Clientes Module:Saldov-")
Me.saldocliv = 0
saldov = 0
Exit Function
End Function


Public Function saldoentrefechas(f1 As String, f2 As String, p As Boolean) As Double
    'busca el saldo de un cliente cc a entre f1 y f2 en una moneda p(true pesos)
    'agrega a la propiedad debe y haber los valores correspondientes
    'On Error GoTo err2
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

    saldoant = saldo(False, f1, p)
    
    Set rs = New ADODB.Recordset
    QUERY = "SELECT * FROM VTA_02 where [id_CLIENTE] = " & Me.id
    QUERY = QUERY & " and datevalue(fecha) >= " & "DateValue('" & f1 & "') "
    QUERY = QUERY & " and datevalue(fecha) <= " & "DateValue('" & f2 & "') "
    QUERY = QUERY & " and  [cta_cte] <> " & "'N'" & " and [contado] = " & "'N'"
    rs.Open QUERY, cn1, , , 1
    totaldebe = 0
    totalhaber = 0
    saldoact = saldoant
    While Not rs.EOF
      dp = 0
      dd = 0
      hp = 0
      hd = 0
      If rs("cta_cte") = "D" Then
       If rs("moneda") = "P" Then
          tp = rs("total")
          td = rs("total_otra_moneda")
          dp = rs("total")
          dd = rs("total_otra_moneda")
       Else
          td = rs("total")
          tp = rs("total_otra_moneda")
          dd = rs("total")
          dp = rs("total_otra_moneda")
       End If
      Else
        If rs("moneda") = "P" Then
          tp = -rs("total")
          td = -rs("total_otra_moneda")
          hp = -rs("total")
          hd = -rs("total_otra_moneda")
       Else
          td = -rs("total")
          tp = -rs("total_otra_moneda")
          hd = -rs("total")
          hp = -rs("total_otra_moneda")
       End If
      End If
   
      If Not rs("id_tipocomp") = 25 Then
         If p Then
         saldoact = saldoact + tp
         totaldebe = totaldebe + dp
         totalhaber = totalhaber + hp
      Else
         saldoact = saldoact + td
         totaldebe = totaldebe + dd
         totalhaber = totalhaber + hd
      End If
    
       Else
            If rs("estado_pago") = "N" Then
                If p Then
                 saldoact = saldoact + tp
                 totaldebe = totaldebe + dp
                totalhaber = totalhaber + hp
            Else
                saldoact = saldoact + td
                totaldebe = totaldebe + dd
                totalhaber = totalhaber + hd
            End If
        End If
       End If
         rs.MoveNext
Wend
Set rs = Nothing
saldocli = saldoact
Me.HABER = totalhaber
Me.DEBE = totaldebe
saldoentrefechas = saldoact
Exit Function

err2:

MsgBox ("Error al generar los saldos. -Class Clientes Module:Saldoentrefechas-")
Me.saldocli = 0
Me.HABER = 0
Me.DEBE = 0
saldoentrefechas = 0
Exit Function

End Function


Public Function saldoentrefechasv(f1 As String, f2 As String, p As Boolean) As Double
    'busca el saldo de un cliente cc a entre f1 y f2 en una moneda p(true pesos)
    'agrega a la propiedad debe y haber los valores correspondientes
    'On Error GoTo err2
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

    saldoant = saldov(False, f1, p)
    
    Set rs = New ADODB.Recordset
    QUERY = "SELECT cta_cte, moneda, total, total_otra_moneda, id_tipocomp, estado_pago  FROM VTA_02 where [id_CLIENTE] = " & Me.id
    QUERY = QUERY & " and datevalue([fecha_vto]) >= " & "DateValue('" & f1 & "') "
    QUERY = QUERY & " and datevalue([fecha_vto]) <= " & "DateValue('" & f2 & "') "
    QUERY = QUERY & " and  [cta_cte] <> " & "'N'" & " and [contado] = " & "'N'"
    rs.Open QUERY, cn1, , , 1
    totaldebe = 0
    totalhaber = 0
    saldoact = saldoant
    While Not rs.EOF
      dp = 0
      dd = 0
      hp = 0
      hd = 0
      If rs("cta_cte") = "D" Then
       If rs("moneda") = "P" Then
          tp = rs("total")
          td = rs("total_otra_moneda")
          dp = rs("total")
          dd = rs("total_otra_moneda")
       Else
          td = rs("total")
          tp = rs("total_otra_moneda")
          dd = rs("total")
          dp = rs("total_otra_moneda")
       End If
      Else
        If rs("moneda") = "P" Then
          tp = -rs("total")
          td = -rs("total_otra_moneda")
          hp = -rs("total")
          hd = -rs("total_otra_moneda")
       Else
          td = -rs("total")
          tp = -rs("total_otra_moneda")
          hd = -rs("total")
          hp = -rs("total_otra_moneda")
       End If
      End If
     
     If Not rs("id_tipocomp") = 25 Then
         If p Then
         saldoact = saldoact + tp
         totaldebe = totaldebe + dp
         totalhaber = totalhaber + hp
      Else
         saldoact = saldoact + td
         totaldebe = totaldebe + dd
         totalhaber = totalhaber + hd
      End If
    
       Else
            If rs("estado_pago") = "N" Then
                If p Then
                 saldoact = saldoact + tp
                 totaldebe = totaldebe + dp
                totalhaber = totalhaber + hp
            Else
                saldoact = saldoact + td
                totaldebe = totaldebe + dd
                totalhaber = totalhaber + hd
            End If
        End If
       End If
      
   rs.MoveNext
Wend
Set rs = Nothing
Me.saldocliv = saldoact
Me.HABER = totalhaber
Me.DEBE = totaldebe
saldoentrefechasv = saldoact
Exit Function

err2:

MsgBox ("Error al generar los saldos por vencimiento. -Class Clientes Module:Saldoentrefechasv-")
Me.saldocliv = 0
Me.HABER = 0
Me.DEBE = 0
saldoentrefechasv = 0
Exit Function

End Function

Public Property Let idvendedor(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.idvendedor = 5
    mvaridvendedor = vData
End Property


Public Property Get idvendedor() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.idvendedor
    idvendedor = mvaridvendedor
End Property



Public Property Let descrimina(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.descrimina = 5
    mvardescrimina = vData
End Property


Public Property Get descrimina() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.descrimina
    descrimina = mvardescrimina
End Property



Public Sub carga(ByVal ip As Long)
'On Error GoTo er_carga
q = "select id_cliente, denominacion, direccion, localidad, provincia, cp, cuit, email, te, id_tipoiva, letra_cliente, descrimina_cliente, abreviatura, id_vendedor, inscripto_operador_granos, percive_ib, id_proveedor, id_prov, cod_fiscal, limite_credito, direccion_local from vta_01, g3 where [id_cliente] = " & ip & " and [id_tipoiva] = [cod_tipoiva]"
'MsgBox (q)
Set rs = New ADODB.Recordset
rs.Open q, cn1
If Not rs.BOF And Not rs.EOF Then
  Me.id = rs("id_cliente")
  Me.razonsocial = rs("denominacion")
  Me.direccion = rs("direccion")
  Me.localidad = rs("localidad")
  Me.provincia = rs("provincia")
  Me.cp = rs("cp")
  Me.CUIT = rs("cuit")
  Me.email = rs("email")
  Me.te = rs("te")
  Me.idtipoiva = rs("id_tipoiva")
  
  If para.tipo_iva_empresa <> 1 Then
    Me.letra = "C"
    Me.descrimina = "N"
  Else
    Me.letra = rs("letra_cliente")
    Me.descrimina = rs("descrimina_cliente")
  End If
    
  Me.abreviatura_tipoiva = rs("abreviatura")
  Me.idvendedor = rs("id_vendedor")
  Me.operador_granos = rs("inscripto_operador_granos")
  Me.perciveib = rs("percive_ib")
  Me.idproveedor = rs("id_proveedor")
  Me.id_provincia = rs("id_prov")
  Me.codfiscal = rs("cod_fiscal")
  Me.limitecredito = rs("limite_credito")
  Me.direccion_local = rs("direccion_local")
  'Me.codfiscal2 = rs("cod_fiscal2")
Else
  Me.id = 0
End If
Set rs = Nothing
Exit Sub
 
er_carga:
  MsgBox ("Error. Clase:Clientes  Modulo:Carga  Cliente Id: " & ip)
  Me.id = 0
  Exit Sub

End Sub

Public Property Let abreviatura_tipoiva(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.abreviatura_tipoiva = 5
    mvarabreviatura_tipoiva = vData
End Property


Public Property Get abreviatura_tipoiva() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.abreviatura_tipoiva
    abreviatura_tipoiva = mvarabreviatura_tipoiva
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



Public Property Let direccion(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Direccion = 5
    mvardireccion = vData
End Property


Public Property Get direccion() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Direccion
    direccion = mvardireccion
End Property



Public Property Let CUIT(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.CUIT = 5
    mvarcuit = vData
End Property


Public Property Get CUIT() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.CUIT
    CUIT = mvarcuit
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



Public Property Let idtipoiva(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.idtipoiva = 5
    mvaridtipoiva = vData
End Property


Public Property Get idtipoiva() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.idtipoiva
    idtipoiva = mvaridtipoiva
End Property



Public Property Let id(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.id = 5
    mvarid = vData
End Property


Public Property Get id() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.id
    id = mvarid
End Property



