VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "padron_ib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
'variables locales para almacenar los valores de las propiedades
Private mvarcuit As Double 'copia local
Private mvartasa_percib As Single 'copia local
Private mvartasa_retib As Single 'copia local
Private mvartipo_contribuyente As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarcuit_texto As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarestado_consulta As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarfecha_desde As Date 'copia local
Private mvarfecha_hasta As Date 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarembargado As Boolean 'copia local
Private mvarfecha_embargo As Date 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarestado_embargo As String 'copia local
Public Property Let estado_embargo(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.estado_embargo = 5
    mvarestado_embargo = vData
End Property


Public Property Get estado_embargo() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.estado_embargo
    estado_embargo = mvarestado_embargo
End Property



Public Property Let fecha_embargo(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.fecha_embargo = 5
    mvarfecha_embargo = vData
End Property


Public Property Get fecha_embargo() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.fecha_embargo
    fecha_embargo = mvarfecha_embargo
End Property



Public Property Let embargado(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.embargado = 5
    mvarembargado = vData
End Property


Public Property Get embargado() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.embargado
    embargado = mvarembargado
End Property



Public Sub inicia()
 Set rsp1 = New ADODB.Recordset
 q = "select [fecha_desde], [fecha_hasta] from i1 where [id_padron] = 1"
 rsp1.MaxRecords = 1
 rsp1.Open q, cnib
 If Not rsp1.EOF And Not rsp1.BOF Then
   Me.fecha_desde = rsp1("fecha_desde")
   Me.fecha_hasta = rsp1("fecha_hasta")
 End If
 Set rsp1 = Nothing
    

End Sub

Public Property Let fecha_hasta(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.fecha_hasta = 5
    mvarfecha_hasta = vData
End Property


Public Property Get fecha_hasta() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.fecha_hasta
    fecha_hasta = mvarfecha_hasta
End Property



Public Property Let fecha_desde(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.fecha_desde = 5
    mvarfecha_desde = vData
End Property


Public Property Get fecha_desde() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.fecha_desde
    fecha_desde = mvarfecha_desde
End Property



Public Property Let estado_consulta(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.estado_consulta = 5
    mvarestado_consulta = vData
End Property


Public Property Get estado_consulta() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.estado_consulta
    estado_consulta = mvarestado_consulta
End Property



Public Property Let cuit_texto(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.cuit_texto = 5
    mvarcuit_texto = vData
End Property


Public Property Get cuit_texto() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.cuit_texto
    cuit_texto = mvarcuit_texto
End Property



Public Sub buscar()

'en estado de consulta devuelve OK si lo encontro en el padron, NO si no lo encontro en el padron , ER si esta mal el numero de cuit
'en estado_embargo devuelve OK si lo encontro en el padron, NO si no lo encontro en el padron , ER si esta mal el numero de cuit

Me.cuit_texto = RTrim$(Me.cuit_texto)
If Len(Me.cuit_texto) = 11 Or Len(Me.cuit_texto) = 13 Then
 If Len(Me.cuit_texto) = 13 Then
  Me.CUIT = Val(Mid$(Me.cuit_texto, 1, 2) & Mid$(Me.cuit_texto, 4, 8) & Mid$(Me.cuit_texto, 13, 1))
 Else
  Me.CUIT = Val(Me.cuit_texto)
 End If
  If Me.CUIT > 0 Then
   Set rscl1 = New ADODB.Recordset
   q = "select [tasa_perc], [tasa_ret] from i2 where [id_padron] = 1 and [cuit] = " & Me.CUIT
   rscl1.MaxRecords = 1
   rscl1.Open q, cnib
   If Not rscl1.EOF And Not rscl1.BOF Then
     'encontrado
     Me.tasa_retib = rscl1("tasa_ret")
     Me.estado_consulta = "OK"
   Else
     'no encontrado carga tasa maxcima
     Me.estado_consulta = "NO"
     
     Set rs7 = New ADODB.Recordset
     q = "select * from i_01 where [id_impuesto] = 50 "
     rs7.Open q, cn1
     If Not rs7.EOF And Not rs7.BOF Then
          Me.tasa_retib = rs7("retencion-minima")
     Else
          Me.tasa_retib = 0
     End If
     Set rs7 = Nothing
   End If
   Set rscl1 = Nothing
  
   Set rscl1 = New ADODB.Recordset
   q = "select [fecha] from e1 where [cuit] = " & Me.CUIT
   rscl1.MaxRecords = 1
   rscl1.Open q, cnib
   If Not rscl1.EOF And Not rscl1.BOF Then
     Me.fecha_embargo = rscl1("fecha")
     Me.estado_embargo = "OK"
   Else
     Me.estado_embargo = "NO"
   End If
   Set rscl1 = Nothing
  
  End If
Else
  MsgBox ("Error en el numero de cuit")
  Me.estado_consulta = "ER"
  Me.estado_embargo = "ER"
  Me.tasa_percib = 0
  Me.tasa_retib = 0

End If

End Sub

Public Sub buscar_perc()

'en estado de consulta devuelve OK si lo encontro en el padron, NO si no lo encontro en el padron , ER si esta mal el numero de cuit
'en estado_embargo devuelve OK si lo encontro en el padron, NO si no lo encontro en el padron , ER si esta mal el numero de cuit

Me.cuit_texto = RTrim$(Me.cuit_texto)
If Len(Me.cuit_texto) = 11 Or Len(Me.cuit_texto) = 13 Then
 If Len(Me.cuit_texto) = 13 Then
  Me.CUIT = Val(Mid$(Me.cuit_texto, 1, 2) & Mid$(Me.cuit_texto, 4, 8) & Mid$(Me.cuit_texto, 13, 1))
 Else
  Me.CUIT = Val(Me.cuit_texto)
 End If
  If Me.CUIT > 0 Then
   Set rscl1 = New ADODB.Recordset
   q = "select [tasa_perc] from i3 where [id_padron] = 1 and [cuit] = " & Me.CUIT
   rscl1.MaxRecords = 1
   rscl1.Open q, cnib
   If Not rscl1.EOF And Not rscl1.BOF Then
     Me.tasa_percib = rscl1("tasa_perc")
     Me.estado_consulta = "OK"
   Else
     Me.estado_consulta = "NO"
     Set rs7 = New ADODB.Recordset
     q = "select * from i_01 where [id_impuesto] = 1 "
     rs7.Open q, cn1
     If Not rs7.EOF And Not rs7.BOF Then
          Me.tasa_percib = rs7("retencion-minima")
     Else
         Me.tasa_percib = 0
     End If
     Set rs7 = Nothing
     
     
     'Me.tasa_retib = 0
   End If
   Set rscl1 = Nothing
  
  
  End If
Else
  MsgBox ("Error en el numero de cuit")
  Me.estado_consulta = "ER"
  Me.estado_embargo = "ER"
  Me.tasa_percib = 0
  Me.tasa_retib = 0

End If

End Sub
Public Property Let tipo_contribuyente(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.tipo_contribuyente = 5
    mvartipo_contribuyente = vData
End Property


Public Property Get tipo_contribuyente() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.tipo_contribuyente
    tipo_contribuyente = mvartipo_contribuyente
End Property



Public Property Let tasa_retib(ByVal vData As Single)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.tasa_retib = 5
    mvartasa_retib = vData
End Property


Public Property Get tasa_retib() As Single
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.tasa_retib
    tasa_retib = mvartasa_retib
End Property



Public Property Let tasa_percib(ByVal vData As Single)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.tasa_percib = 5
    mvartasa_percib = vData
End Property


Public Property Get tasa_percib() As Single
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.tasa_percib
    tasa_percib = mvartasa_percib
End Property



Public Property Let CUIT(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.cuit = 5
    mvarcuit = vData
End Property


Public Property Get CUIT() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.cuit
    CUIT = mvarcuit
End Property



