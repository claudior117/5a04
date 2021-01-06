VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "chterceros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
'variables locales para almacenar los valores de las propiedades
Private mvarnuminterno As Long 'copia local
Private mvarfecha_emision As Date 'copia local
Private mvarnum_cheque As Long 'copia local
Private mvarbanco As String 'copia local
Private mvarsucursal As String 'copia local
Private mvartitular As String 'copia local
Private mvarimporte As Double 'copia local
Private mvarestado As String 'copia local
Private mvarfechadif As Date 'copia local
Private mvarorigen As String 'copia local
Private mvardestino As String 'copia local
Private mvarnummovbancoi As Long 'copia local
Private mvarnummovbancoe As Long 'copia local
Private mvarnumintop As Long 'copia local
Private mvarnumintrbo As Long 'copia local
Private mvarfechasalida As Date 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvartiposalida As String 'copia local
Public Property Let tiposalida(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.tiposalida = 5
    mvartiposalida = vData
End Property


Public Property Get tiposalida() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.tiposalida
    tiposalida = mvartiposalida
End Property



Public Sub cargar(ByVal numint As Long)
'On Error GoTo err1
Set rs = New adodb.Recordset
q = "select * from cyb_03 where [num_interno] = " & numint
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
   Me.numinterno = rs("num_interno")
   Me.fecha_emision = rs("fecha_emision")
   Me.num_cheque = rs("num_cheque")
   Me.banco = rs("banco")
   Me.sucursal = rs("sucursal")
   Me.titular = rs("titular")
   Me.importe = rs("importe")
   Me.estado = rs("estado")
   Me.fechadif = rs("fecha_dif")
   Me.origen = rs("origen")
   Me.destino = rs("destino")
   Me.nummovbancoi = rs("num_mov_banco_i")
   Me.nummovbancoe = rs("num_mov_banco_e")
   Me.numintop = rs("num_int_op")
   Me.numintrbo = rs("num_int_rbo")
   Me.fechasalida = rs("fecha_salida")
   Me.tiposalida = rs("tipo_salida")
Else
  Me.numinterno = 0
End If
Set rs = Nothing

Exit Sub
err1:
  MsgBox ("Error. Clase:chterceros  Modulo:Carga")
  Me.numinterno = 0
  Exit Sub
End Sub

Public Property Let fechasalida(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.fechasalida = 5
    mvarfechasalida = vData
End Property


Public Property Get fechasalida() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.fechasalida
    fechasalida = mvarfechasalida
End Property



Public Property Let numintrbo(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.numintrbo = 5
    mvarnumintrbo = vData
End Property


Public Property Get numintrbo() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.numintrbo
    numintrbo = mvarnumintrbo
End Property



Public Property Let numintop(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.numintop = 5
    mvarnumintop = vData
End Property


Public Property Get numintop() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.numintop
    numintop = mvarnumintop
End Property



Public Property Let nummovbancoe(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.nummovbancoe = 5
    mvarnummovbancoe = vData
End Property


Public Property Get nummovbancoe() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.nummovbancoe
    nummovbancoe = mvarnummovbancoe
End Property



Public Property Let nummovbancoi(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.nummovbancoi = 5
    mvarnummovbancoi = vData
End Property


Public Property Get nummovbancoi() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.nummovbancoi
    nummovbancoi = mvarnummovbancoi
End Property



Public Property Let destino(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.destino = 5
    mvardestino = vData
End Property


Public Property Get destino() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.destino
    destino = mvardestino
End Property



Public Property Let origen(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.origen = 5
    mvarorigen = vData
End Property


Public Property Get origen() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.origen
    origen = mvarorigen
End Property



Public Property Let fechadif(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.fechadif = 5
    mvarfechadif = vData
End Property


Public Property Get fechadif() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.fechadif
    fechadif = mvarfechadif
End Property



Public Property Let estado(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.estado = 5
    mvarestado = vData
End Property


Public Property Get estado() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.estado
    estado = mvarestado
End Property



Public Property Let importe(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.importe = 5
    mvarimporte = vData
End Property


Public Property Get importe() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.importe
    importe = mvarimporte
End Property



Public Property Let titular(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.titular = 5
    mvartitular = vData
End Property


Public Property Get titular() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.titular
    titular = mvartitular
End Property



Public Property Let sucursal(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.sucursal = 5
    mvarsucursal = vData
End Property


Public Property Get sucursal() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.sucursal
    sucursal = mvarsucursal
End Property



Public Property Let banco(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.banco = 5
    mvarbanco = vData
End Property


Public Property Get banco() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.banco
    banco = mvarbanco
End Property



Public Property Let num_cheque(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.num_cheque = 5
    mvarnum_cheque = vData
End Property


Public Property Get num_cheque() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.num_cheque
    num_cheque = mvarnum_cheque
End Property



Public Property Let fecha_emision(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.fecha_emision = 5
    mvarfecha_emision = vData
End Property


Public Property Get fecha_emision() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.fecha_emision
    fecha_emision = mvarfecha_emision
End Property



Public Property Let numinterno(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.numinterno = 5
    mvarnuminterno = vData
End Property


Public Property Get numinterno() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.numinterno
    numinterno = mvarnuminterno
End Property



