VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "fiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'variables locales para almacenar los valores de las propiedades
Private mvarid As Long 'copia local
Private mvarimpresora As String 'copia local
Private mvarsucursalfiscal As Integer 'copia local
Private mvarimprimetique As String 'copia local
Private mvarimprimetf As String 'copia local
Private mvarimprimefact As String 'copia local
Private mvarabrecajon As String 'copia local
Private mvarlimitetique As Double 'copia local
Private mvarlimitetf As Double 'copia local
Private mvarlimitefact As Double 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarimprimerto As String 'copia local
Private mvarimprimend As String 'copia local
Private mvarimprimenc As String 'copia local
Private mvarimprimerbo As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarpuerto As Integer 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarCODFACT As String 'copia local
Private mvarCODRTO As String 'copia local
Private mvarCODRBO As String 'copia local
Private mvarCODNC As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarcaracteresmax As Integer 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarultimo_z As Long 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvaridmodelo As Integer 'copia local
Private mvarbaudios As Long 'copia local
Public Property Let baudios(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.baudios = 5
    mvarbaudios = vData
End Property


Public Property Get baudios() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.baudios
    baudios = mvarbaudios
End Property



Public Property Let idmodelo(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.idmodelo = 5
    mvaridmodelo = vData
End Property


Public Property Get idmodelo() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.idmodelo
    idmodelo = mvaridmodelo
End Property



Public Property Let ultimo_z(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.ultimo_z = 5
    mvarultimo_z = vData
End Property


Public Property Get ultimo_z() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.ultimo_z
    ultimo_z = mvarultimo_z
End Property



Public Property Let caracteresmax(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.caracteresmax = 5
    mvarcaracteresmax = vData
End Property


Public Property Get caracteresmax() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.caracteresmax
    caracteresmax = mvarcaracteresmax
End Property



Public Property Let CODNC(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.CODNC = 5
    mvarCODNC = vData
End Property


Public Property Get CODNC() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.CODNC
    CODNC = mvarCODNC
End Property



Public Property Let CODRBO(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.CODRBO = 5
    mvarCODRBO = vData
End Property


Public Property Get CODRBO() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.CODRBO
    CODRBO = mvarCODRBO
End Property



Public Property Let CODRTO(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.CODRTO = 5
    mvarCODRTO = vData
End Property


Public Property Get CODRTO() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.CODRTO
    CODRTO = mvarCODRTO
End Property



Public Property Let CODFACT(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.CODFACT = 5
    mvarCODFACT = vData
End Property


Public Property Get CODFACT() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.CODFACT
    CODFACT = mvarCODFACT
End Property



Public Property Let puerto(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.puerto = 5
    mvarpuerto = vData
End Property


Public Property Get puerto() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.puerto
    puerto = mvarpuerto
End Property



Public Property Let imprimerbo(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.imprimerbo = 5
    mvarimprimerbo = vData
End Property


Public Property Get imprimerbo() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.imprimerbo
    imprimerbo = mvarimprimerbo
End Property



Public Property Let imprimenc(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.imprimenc = 5
    mvarimprimenc = vData
End Property


Public Property Get imprimenc() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.imprimenc
    imprimenc = mvarimprimenc
End Property



Public Property Let imprimend(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.imprimend = 5
    mvarimprimend = vData
End Property


Public Property Get imprimend() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.imprimend
    imprimend = mvarimprimend
End Property



Public Property Let imprimerto(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.imprimerto = 5
    mvarimprimerto = vData
End Property


Public Property Get imprimerto() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.imprimerto
    imprimerto = mvarimprimerto
End Property



Public Sub carga(ByVal suc As Integer)
Set rsF = New ADODB.Recordset
q = "select * from fsc_001 where [sucursal_fiscal] = " & suc
rsF.Open q, cn1
If Not rsF.EOF And Not rsF.BOF Then
   Me.id = rsF("id")
   Me.abrecajon = rsF("abre_cajon")
   Me.impresora = rsF("impresora_fiscal")
   Me.imprimefact = rsF("imprime_factura")
   Me.imprimetf = rsF("imprime_tiquefactura")
   Me.imprimetique = rsF("imprime_tique")
   Me.limitefact = rsF("limite_factura")
   Me.limitetf = rsF("limite_tiquefactura")
   Me.limitetique = rsF("limite_tique")
   Me.imprimerto = rsF("imprime_remito")
   Me.imprimerbo = rsF("imprime_recibo")
   Me.imprimend = rsF("imprime_nd")
   Me.imprimenc = rsF("imprime_nc")
   Me.puerto = rsF("puerto_serie")
   Me.CODFACT = rsF("cod_fact")
   Me.CODNC = rsF("cod_nc")
   Me.CODRBO = rsF("cod_rbo")
   Me.CODRTO = rsF("cod_rto")
   Me.caracteresmax = rsF("caracteres_max")
   Me.ultimo_z = rsF("ult_z")
   Me.idmodelo = rsF("id_modelo")
   Me.baudios = rsF("baudios")
 Else
   Me.id = 0
   Me.sucursalfiscal = 0
   Me.ultimo_z = 0
 End If
 Set rsF = Nothing
   
End Sub

Public Property Let limitefact(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.limitefact = 5
    mvarlimitefact = vData
End Property


Public Property Get limitefact() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.limitefact
    limitefact = mvarlimitefact
End Property



Public Property Let limitetf(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.limitetf = 5
    mvarlimitetf = vData
End Property


Public Property Get limitetf() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.limitetf
    limitetf = mvarlimitetf
End Property



Public Property Let limitetique(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.limitetique = 5
    mvarlimitetique = vData
End Property


Public Property Get limitetique() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.limitetique
    limitetique = mvarlimitetique
End Property



Public Property Let abrecajon(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.abrecajon = 5
    mvarabrecajon = vData
End Property


Public Property Get abrecajon() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.abrecajon
    abrecajon = mvarabrecajon
End Property



Public Property Let imprimefact(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.imprimefact = 5
    mvarimprimefact = vData
End Property


Public Property Get imprimefact() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.imprimefact
    imprimefact = mvarimprimefact
End Property



Public Property Let imprimetf(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.imprimetf = 5
    mvarimprimetf = vData
End Property


Public Property Get imprimetf() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.imprimetf
    imprimetf = mvarimprimetf
End Property



Public Property Let imprimetique(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.imprimetique = 5
    mvarimprimetique = vData
End Property


Public Property Get imprimetique() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.imprimetique
    imprimetique = mvarimprimetique
End Property



Public Property Let sucursalfiscal(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.sucursalfiscal = 5
    mvarsucursalfiscal = vData
End Property


Public Property Get sucursalfiscal() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.sucursalfiscal
    sucursalfiscal = mvarsucursalfiscal
End Property



Public Property Let impresora(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.impresora = 5
    mvarimpresora = vData
End Property


Public Property Get impresora() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.impresora
    impresora = mvarimpresora
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



