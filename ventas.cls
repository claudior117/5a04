VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"comprobantes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarcomprobantes As comprobantes
'variables locales para almacenar los valores de las propiedades
Private mvarabreviatura As String 'copia local
Private mvarCTACTE As String 'copia local
Private mvarestado As String 'copia local
Private mvarestado_pago As String 'copia local
Private mvarfecha As Date 'copia local
Private mvargrabado As String 'copia local
Private mvaridcuenta As Long 'copia local
Private mvaridproveedor As Long 'copia local
Private mvaridtipocomp As Integer 'copia local
Private mvaridusuario As Long 'copia local
Private mvariva As Double 'copia local
Private mvarletra As String 'copia local
Private mvarnograbado As Double 'copia local
Private mvarnum_op As String 'copia local
Private mvarnumcomp As Long 'copia local
Private mvarnumint As Long 'copia local
Private mvarpercep As Double 'copia local
Private mvarSTOCK As String 'copia local
Private mvarsubtotal As Double 'copia local
Private mvarsucursal As Integer 'copia local
Private mvartotal As Double 'copia local
Public Property Let total(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.total = 5
    mvartotal = vData
End Property


Public Property Get total() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.total
    total = mvartotal
End Property



Public Property Let sucursal(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.sucursal = 5
    mvarsucursal = vData
End Property


Public Property Get sucursal() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.sucursal
    sucursal = mvarsucursal
End Property



Public Property Let subtotal(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.subtotal = 5
    mvarsubtotal = vData
End Property


Public Property Get subtotal() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.subtotal
    subtotal = mvarsubtotal
End Property



Public Property Let STOCK(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.STOCK = 5
    mvarSTOCK = vData
End Property


Public Property Get STOCK() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.STOCK
    STOCK = mvarSTOCK
End Property



Public Property Let percep(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.percep = 5
    mvarpercep = vData
End Property


Public Property Get percep() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.percep
    percep = mvarpercep
End Property



Public Property Let numint(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.numint = 5
    mvarnumint = vData
End Property


Public Property Get numint() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.numint
    numint = mvarnumint
End Property



Public Property Let numcomp(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.numcomp = 5
    mvarnumcomp = vData
End Property


Public Property Get numcomp() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.numcomp
    numcomp = mvarnumcomp
End Property



Public Property Let num_op(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.num_op = 5
    mvarnum_op = vData
End Property


Public Property Get num_op() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.num_op
    num_op = mvarnum_op
End Property



Public Property Let nograbado(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.nograbado = 5
    mvarnograbado = vData
End Property


Public Property Get nograbado() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.nograbado
    nograbado = mvarnograbado
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



Public Property Let iva(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.iva = 5
    mvariva = vData
End Property


Public Property Get iva() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.iva
    iva = mvariva
End Property



Public Property Let idusuario(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.idusuario = 5
    mvaridusuario = vData
End Property


Public Property Get idusuario() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.idusuario
    idusuario = mvaridusuario
End Property



Public Property Let idtipocomp(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.idtipocomp = 5
    mvaridtipocomp = vData
End Property


Public Property Get idtipocomp() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.idtipocomp
    idtipocomp = mvaridtipocomp
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



Public Property Let idcuenta(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.idcuenta = 5
    mvaridcuenta = vData
End Property


Public Property Get idcuenta() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.idcuenta
    idcuenta = mvaridcuenta
End Property



Public Property Let grabado(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.grabado = 5
    mvargrabado = vData
End Property


Public Property Get grabado() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.grabado
    grabado = mvargrabado
End Property



Public Property Let fecha(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.fecha = 5
    mvarfecha = vData
End Property


Public Property Get fecha() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.fecha
    fecha = mvarfecha
End Property



Public Property Let estado_pago(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.estado_pago = 5
    mvarestado_pago = vData
End Property


Public Property Get estado_pago() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.estado_pago
    estado_pago = mvarestado_pago
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



Public Property Let CTACTE(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.CTACTE = 5
    mvarCTACTE = vData
End Property


Public Property Get CTACTE() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.CTACTE
    CTACTE = mvarCTACTE
End Property



Public Property Let abreviatura(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.abreviatura = 5
    mvarabreviatura = vData
End Property


Public Property Get abreviatura() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.abreviatura
    abreviatura = mvarabreviatura
End Property






Private Sub Class_Terminate()
    Set mvarcomprobantes = Nothing
End Sub



