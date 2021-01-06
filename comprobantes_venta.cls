VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "comprobantes_venta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"STOCK"
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
'variables locales para almacenar los valores de las propiedades
Private mvarabreviatura As String 'copia local
Private mvarCTACTE As String 'copia local
Private mvargrabado As String 'copia local
Private mvarSTOCK As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarfecha As Date 'copia local
Private mvarsucursal As Integer 'copia local
Private mvarnumcomp As Long 'copia local
Private mvarletra As String 'copia local
Private mvaridcliente As Long 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvaridtipocomp As Integer 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarventa As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarnumint As Long 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvaridusuario As Long 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarsubtotal As Double 'copia local
Private mvarimpuestos As Double 'copia local
Private mvariva As Double 'copia local
Private mvartotal As Double 'copia local
Private mvardescuento As Double 'copia local
Private mvarobservaciones As String 'copia local
Private mvaridcuenta As Long 'copia local
Private mvarcotizaciondolar As Double 'copia local
Private mvartotalotramoneda As Double 'copia local
Private mvarmoneda As String 'copia local
Private mvaridvendedor As Long 'copia local
Private mvarrecibopago As String 'copia local
Private mvaridtransporte As Double 'copia local
Private mvarperciva As Double 'copia local
Private mvarpercgan As Double 'copia local
Private mvarpercib As Double 'copia local
Private mvarfechapago As Date 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarestadopago As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarPROPIO As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarcontabilidad As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarfecha_vto As Date 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvartransporte As String 'copia local
Private mvardir_transporte As String 'copia local
Private mvarcuit_transporte As String 'copia local
Private mvarloc_transporte As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarbultos_totales As Integer 'copia local
Private mvarvalor_declarado As Double 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvartipoimpresora As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarcontado As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarrazonsocial As String 'copia local
Private mvardireccion As String 'copia local
Private mvarlocalidad As String 'copia local
Private mvarcuit As String 'copia local
Private mvaridtipoiva As Variant 'copia local
Private mvarabreviatura_tipoiva As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarcant_lineas As Integer 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarestado As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarCHOFER As String 'copia local
Private mvarPATENTE As String 'copia local
Private mvarPATENTE_ACOPLADO As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvaridcamion As Long 'copia local
Private mvardnichofer As Long 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarimprimedescextra As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarcod_afip_a As Integer 'copia local
Private mvarcod_afip_b As Integer 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarcae As Variant 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarfecha_vto_cae As Date 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvartipoop As Integer 'copia local
Public Property Let tipoop(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.tipoop = 5
    mvartipoop = vData
End Property


Public Property Get tipoop() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.tipoop
    tipoop = mvartipoop
End Property



Public Property Let fecha_vto_cae(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.fecha_vto_cae = 5
    mvarfecha_vto_cae = vData
End Property


Public Property Get fecha_vto_cae() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.fecha_vto_cae
    fecha_vto_cae = mvarfecha_vto_cae
End Property



Public Property Let cae(ByVal vData As Variant)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.cae = 5
    mvarcae = vData
End Property


Public Property Set cae(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.cae = Form1
    Set mvarcae = vData
End Property


Public Property Get cae() As Variant
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.cae
    If IsObject(mvarcae) Then
        Set cae = mvarcae
    Else
        cae = mvarcae
    End If
End Property



Public Property Let cod_afip_b(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.cod_afip_b = 5
    mvarcod_afip_b = vData
End Property


Public Property Get cod_afip_b() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.cod_afip_b
    cod_afip_b = mvarcod_afip_b
End Property



Public Property Let cod_afip_a(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.cod_afip_a = 5
    mvarcod_afip_a = vData
End Property


Public Property Get cod_afip_a() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.cod_afip_a
    cod_afip_a = mvarcod_afip_a
End Property



Public Property Let imprimedescextra(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.imprimedescextra = 5
    mvarimprimedescextra = vData
End Property
Sub imprimirrto4(ByVal c As Integer)
'tulio
 dd = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 1, 2)
 mm = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 4, 2)
 aa = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 7, 4)
 'lugar = Left$(RTrim$(cl_cli.direccion) & " - (" & RTrim$(cl_cli.cp) & ")" & RTrim$(cl_cli.localidad), 120)

 If Me.idtipocomp = 45 Then
  tacha = " "
 Else
   tacha = Me.abreviatura & "     ###################"
 End If
 
 For Y = 1 To c
  Printer.FontName = "Times New Roman"
  Printer.FontSize = 10
  Printer.FontBold = False
  For i = 1 To 7
    Printer.Print
  Next i
  Printer.FontName = "Times New Roman"
  Printer.FontUnderline = False
  Printer.FontBold = False
  Printer.FontSize = 12
  Printer.Print Tab(71); tacha
  Printer.Print Tab(80); Spc(4); dd & "/" & mm & "/" & aa
  For i = 1 To 3
     Printer.Print
  Next i
 Printer.FontSize = 8
 Printer.Print
 Printer.Print
 Printer.FontName = "Times New Roman"
 Printer.FontSize = 14
 Printer.FontBold = True
 Printer.Print Tab(19); "["; Format$(Me.idcliente, "0000"); "] "; Me.razonsocial
 Printer.FontName = "Courier New"
 Printer.FontSize = 10
 Printer.FontBold = False
 Printer.Print
 Printer.Print Tab(20); Me.direccion; Tab(68); Me.CUIT
 Printer.FontSize = 8
 Printer.Print Tab(25); Me.localidad
 Printer.FontSize = 10
 Printer.Print
 Printer.Print Tab(20); Me.abreviatura_tipoiva
 Printer.FontName = "Courier New"
 Printer.FontSize = 9
 Printer.FontBold = False
 Printer.Print
 Printer.Print
 Printer.Print
 Printer.Print
 Set rs = New adodb.Recordset
 q = "Select * from vta_03 where [num_int] = " & Me.numint
 rs.Open q, cn1
 l = 0
 While Not rs.EOF
    i = Space$(10)
    t = Space$(4)
    pu = Space$(10)
    d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    dt = Space$(9)
    cant = Space$(10)
    cp2 = Space$(9)
   If rs("cantidad") > 0 Then
      RSet cant = Format$(rs("cantidad"), "######0.00")
    End If
    Printer.Print Tab(8); cant; Spc(4);
    Printer.FontBold = True
    Printer.Print d
 
    rs.MoveNext
    l = l + 1
   Wend
   Set rs = Nothing
  
   
   For i = l To 33
     Printer.Print
   Next i
   
   Set cl_usuarios = New usuarios
   cl_usuarios.cargar (Me.idusuario)
   If cl_usuarios.idusuario > 0 Then
        Printer.Print Tab(10); "Emitida por: " & cl_usuarios.denominACION
   Else
        Printer.Print
   End If
   Set cl_usuarios = Nothing
   If Y < c Then
      Printer.NewPage
   End If
  Next Y
 
 Printer.EndDoc

End Sub


Public Property Get imprimedescextra() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.imprimedescextra
    imprimedescextra = mvarimprimedescextra
End Property

















Public Property Let dnichofer(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.dnichofer = 5
    mvardnichofer = vData
End Property


Public Property Get dnichofer() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.dnichofer
    dnichofer = mvardnichofer
End Property



Public Property Let idcamion(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.idcamion = 5
    mvaridcamion = vData
End Property


Public Property Get idcamion() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.idcamion
    idcamion = mvaridcamion
End Property



Public Property Let PATENTE_ACOPLADO(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.PATENTE_ACOPLADO = 5
    mvarPATENTE_ACOPLADO = vData
End Property


Public Property Get PATENTE_ACOPLADO() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.PATENTE_ACOPLADO
    PATENTE_ACOPLADO = mvarPATENTE_ACOPLADO
End Property



Public Property Let PATENTE(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.PATENTE = 5
    mvarPATENTE = vData
End Property


Public Property Get PATENTE() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.PATENTE
    PATENTE = mvarPATENTE
End Property



Public Property Let CHOFER(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.CHOFER = 5
    mvarCHOFER = vData
End Property


Public Property Get CHOFER() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.CHOFER
    CHOFER = mvarCHOFER
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



Public Property Let cant_lineas(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.cant_lineas = 5
    mvarcant_lineas = vData
End Property


Public Property Get cant_lineas() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.cant_lineas
    cant_lineas = mvarcant_lineas
End Property



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



Public Property Let idtipoiva(ByVal vData As Variant)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.idtipoiva = 5
    mvaridtipoiva = vData
End Property


Public Property Set idtipoiva(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.idtipoiva = Form1
    Set mvaridtipoiva = vData
End Property


Public Property Get idtipoiva() As Variant
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.idtipoiva
    If IsObject(mvaridtipoiva) Then
        Set idtipoiva = mvaridtipoiva
    Else
        idtipoiva = mvaridtipoiva
    End If
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



Public Property Let contado(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.contado = 5
    mvarcontado = vData
End Property


Public Property Get contado() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.contado
    contado = mvarcontado
End Property



Public Sub carga_datos_cot(ByVal id_comp As Double)



End Sub

Public Property Let tipoImpresora(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.tipoimpresora = 5
    mvartipoimpresora = vData
End Property


Public Property Get tipoImpresora() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.tipoimpresora
    tipoImpresora = mvartipoimpresora
End Property



Public Property Let valor_declarado(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.valor_declarado = 5
    mvarvalor_declarado = vData
End Property


Public Property Get valor_declarado() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.valor_declarado
    valor_declarado = mvarvalor_declarado
End Property



Public Property Let bultos_totales(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.bultos_totales = 5
    mvarbultos_totales = vData
End Property


Public Property Get bultos_totales() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.bultos_totales
    bultos_totales = mvarbultos_totales
End Property



Public Property Let loc_transporte(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.loc_transporte = 5
    mvarloc_transporte = vData
End Property


Public Property Get loc_transporte() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.loc_transporte
    loc_transporte = mvarloc_transporte
End Property



Public Property Let cuit_transporte(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.cuit_transporte = 5
    mvarcuit_transporte = vData
End Property


Public Property Get cuit_transporte() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.cuit_transporte
    cuit_transporte = mvarcuit_transporte
End Property



Public Property Let dir_transporte(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.dir_transporte = 5
    mvardir_transporte = vData
End Property


Public Property Get dir_transporte() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.dir_transporte
    dir_transporte = mvardir_transporte
End Property



Public Property Let transporte(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.transporte = 5
    mvartransporte = vData
End Property


Public Property Get transporte() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.transporte
    transporte = mvartransporte
End Property



Public Property Let fecha_vto(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.fecha_vto = 5
    mvarfecha_vto = vData
End Property


Public Property Get fecha_vto() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.fecha_vto
    fecha_vto = mvarfecha_vto
End Property



Public Property Let contabilidad(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.contabilidad = 5
    mvarcontabilidad = vData
End Property


Public Property Get contabilidad() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.contabilidad
    contabilidad = mvarcontabilidad
End Property



Public Property Let PROPIO(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.PROPIO = 5
    mvarPROPIO = vData
End Property


Public Property Get PROPIO() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.PROPIO
    PROPIO = mvarPROPIO
End Property



Public Property Let estadopago(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.estadopago = 5
    mvarestadopago = vData
End Property


Public Property Get estadopago() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.estadopago
    estadopago = mvarestadopago
End Property



Public Property Let fechapago(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.fechapago = 5
    mvarfechapago = vData
End Property


Public Property Get fechapago() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.fechapago
    fechapago = mvarfechapago
End Property



Public Property Let percib(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.percib = 5
    mvarpercib = vData
End Property


Public Property Get percib() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.percib
    percib = mvarpercib
End Property



Public Property Let percgan(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.percgan = 5
    mvarpercgan = vData
End Property


Public Property Get percgan() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.percgan
    percgan = mvarpercgan
End Property



Public Property Let perciva(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.perciva = 5
    mvarperciva = vData
End Property


Public Property Get perciva() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.perciva
    perciva = mvarperciva
End Property



Public Property Let idtransporte(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.idtransporte = 5
    mvaridtransporte = vData
End Property


Public Property Get idtransporte() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.idtransporte
    idtransporte = mvaridtransporte
End Property



Public Property Let recibopago(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.recibopago = 5
    mvarrecibopago = vData
End Property


Public Property Get recibopago() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.recibopago
    recibopago = mvarrecibopago
End Property






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



Public Property Let moneda(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.moneda = 5
    mvarmoneda = vData
End Property


Public Property Get moneda() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.moneda
    moneda = mvarmoneda
End Property



Public Property Let totalotramoneda(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.totalotramoneda = 5
    mvartotalotramoneda = vData
End Property


Public Property Get totalotramoneda() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.totalotramoneda
    totalotramoneda = mvartotalotramoneda
End Property



Public Property Let cotizaciondolar(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.cotizaciondolar = 5
    mvarcotizaciondolar = vData
End Property


Public Property Get cotizaciondolar() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.cotizaciondolar
    cotizaciondolar = mvarcotizaciondolar
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



Public Property Let observaciones(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.observaciones = 5
    mvarobservaciones = vData
End Property


Public Property Get observaciones() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.observaciones
    observaciones = mvarobservaciones
End Property



Public Property Let descuento(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.descuento = 5
    mvardescuento = vData
End Property


Public Property Get descuento() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.descuento
    descuento = mvardescuento
End Property



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



Public Property Let impuestos(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.impuestos = 5
    mvarimpuestos = vData
End Property


Public Property Get impuestos() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.impuestos
    impuestos = mvarimpuestos
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



Public Property Let venta(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.venta = 5
    mvarventa = vData
End Property


Public Property Get venta() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.venta
    venta = mvarventa
End Property



Public Sub ACTUALIZA_NUMERADOR()
   Set rs = New adodb.Recordset
   q = "select * from vta_06 where [sucursal] = " & Me.sucursal & " and  [id_tipocomp] = " & Me.idtipocomp
   rs.Open q, cn1, adOpenStatic, adLockOptimistic
   If Not rs.BOF And Not rs.EOF Then
      Select Case Me.letra
       Case Is = "A"
               rs("ULT_NUM_a") = Me.numcomp
               rs.Update
       Case Is = "B"
               rs("ULT_NUM_b") = Me.numcomp
               rs.Update
       
       Case Is = "C"
               rs("ULT_NUM_c") = Me.numcomp
               rs.Update
       Case Is = "E"
               rs("ULT_NUM_e") = Me.numcomp
               rs.Update

       Case Else
               rs("ULT_NUM_a") = Me.numcomp
               rs.Update
      End Select
                    
      
   
   Else
     MsgBox ("Error. No se puede Inicializar el Comprobante")
   End If
       
   Set rs = Nothing
   
   If Me.idtipocomp >= 41 And Me.idtipocomp < 50 Then
       Set rs = New adodb.Recordset
       q = "select * from  vta_06 where [sucursal] = " & Me.sucursal & " and  [id_tipocomp] >= 41 and [id_tipocomp] < 50"
       rs.Open q, cn1, adOpenDynamic, adLockOptimistic
       While Not rs.EOF
          rs("ULT_NUM_a") = Me.numcomp
          rs.Update
          rs.MoveNext
       Wend
       Set rs = Nothing
   End If
   
   If para.numeracion_comun_Fact_nc = "S" And Me.idtipocomp <= 10 Then
     Set rs = New adodb.Recordset
     q = "select * from vta_06 where [sucursal] = " & Me.sucursal & " and  [id_tipocomp] <= 10"
     rs.Open q, cn1, adOpenStatic, adLockOptimistic
     While Not rs.EOF
        Select Case Me.letra
       Case Is = "A"
               rs("ULT_NUM_a") = Me.numcomp
               rs.Update
       Case Is = "B"
               rs("ULT_NUM_b") = Me.numcomp
               rs.Update
       Case Is = "C"
               rs("ULT_NUM_c") = Me.numcomp
               rs.Update
      End Select
      rs.MoveNext
    Wend
    Set rs = Nothing
  End If



End Sub

Public Sub web(ni)
     
     
     'actualizacion web
     ' If para.idsistema > 0 Then
     '    cl_compvta.cargar2 (ni)
             
         
     '    Set cnn_nube = New ADODB.Connection
     '    cnn_nube.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=159.65.106.144;uid=aero;pwd=s7VXRGBZ27E=;database=5a04;connection="
     '    On Error GoTo errw
     '    cnn_nube.Open
     '    cnn_nube.BeginTrans
     '    q = ""
     '    F = Year(cl_compvta.fecha) & "-" & Month(cl_compvta.fecha) & "-" & Day(cl_compvta.fecha)
          
     '   If cl_compvta.contado = "S" Then
     '     u = "N"
     '   Else
     '     u = cl_compvta.ctacte
     '   End If
          
     '   a = cl_compvta.abreviatura
           
           
     '   c = a & " " & cl_compvta.letra & Format$(cl_compvta.sucursal, "0000") & "-" & Format$(cl_compvta.numcomp, "00000000")
              
     '  q = ""
     '  q = "INSERT INTO ventas(id_sistema, fecha, importe, ubicacion, comprobante, id_cliente , cliente, num_int)"
     '  q = q & " VALUES (" & para.idsistema & ",'" & F & "', " & cl_compvta.total & ", '" & u & "', '" & c & "', " & cl_compvta.idcliente & ", '" & cl_compvta.razonsocial & "', " & cl_compvta.numint & ")"
     '  cnn_nube.Execute q
        
     '  cnn_nube.CommitTrans
      
      
       
      
     ' End If
      
   ' Exit Sub
    
'errw:
'    MsgBox ("Error al actualizar comrpobante en la WEB, ejecute proceso de sicncronizacion")
'    Exit Sub
  



End Sub


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



Public Property Let idcliente(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.idcliente = 5
    mvaridcliente = vData
End Property


Public Property Get idcliente() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.idcliente
    idcliente = mvaridcliente
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


'FIXIT: Declare 'numint' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Sub cargar2(ByVal numint)
Set rs = New adodb.Recordset
q = "select * from vta_02, vta_06 where [num_int] = " & numint & " and vta_02.[id_tipocomp] = vta_06.[id_tipocomp] and  vta_02.[sucursal_ingreso] = vta_06.[sucursal]"
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
   Me.numint = rs("num_int")
   Me.ctacte = rs("cta_cte")
   Me.fecha = rs("fecha")
   Me.grabado = rs("grabado")
   Me.idcliente = rs("id_cliente")
   Me.idtipocomp = rs("vta_02.id_tipocomp")
   Me.letra = rs("letra")
   Me.numcomp = rs("num_comp")
   Me.STOCK = rs("vta_02.stock")
   Me.sucursal = rs("vta_02.sucursal")
   Me.venta = rs("vta_02.venta")
   Me.idusuario = rs("id_usuario")
   Me.subtotal = rs("subtotal")
   Me.impuestos = rs("impuestos")
   Me.iva = rs("vta_02.iva")
   Me.total = rs("total")
   Me.descuento = rs("descuento")
   Me.observaciones = rs("observaciones")
   Me.idcuenta = rs("id_cuenta")
   Me.cotizaciondolar = rs("cotizacion_dolar")
   Me.totalotramoneda = rs("total_otra_moneda")
   Me.moneda = rs("vta_02.moneda")
   Me.idvendedor = rs("id_vendedor")
   Me.estadopago = rs("estado_pago")
   Me.recibopago = rs("recibo_pago")
   'Me.fechapago = rs("fecha_pago")
   Me.idtransporte = rs("id_transporte")
   Me.perciva = rs("perc_iva")
   Me.percgan = rs("perc_gan")
   Me.percib = rs("perc_ib")
   If Not IsNull(rs("fecha_vto")) Then
     Me.fecha_vto = rs("fecha_vto")
   Else
     Me.fecha_vto = Format$(Now, "dd/mm/yyyy")
   End If
   Me.abreviatura = rs("abreviatura")
   If Not IsNull(rs("transporte")) Then
      Me.transporte = rs("transporte")
   Else
      Me.transporte = " "
   End If
   If Not IsNull(rs("direccion_transp")) Then
      Me.dir_transporte = rs("direccion_transp")
   Else
      Me.dir_transporte = " "
   End If
   If Not IsNull(rs("cuit_transp")) Then
      Me.cuit_transporte = rs("cuit_transp")
   Else
      Me.cuit_transporte = " "
   End If
   If Not IsNull(rs("CHOFER02")) Then
      Me.CHOFER = rs("chofer02")
   Else
      Me.CHOFER = " "
   End If
   If Not IsNull(rs("dominio02")) Then
      Me.PATENTE = rs("dominio02")
   Else
      Me.PATENTE = " "
   End If
   If Not IsNull(rs("dominio_acoplado02")) Then
      Me.PATENTE_ACOPLADO = rs("dominio_acoplado02")
   Else
      Me.PATENTE_ACOPLADO = " "
   End If
   Me.idcamion = rs("id_camion02")
   Me.dnichofer = rs("dni_chofer02")
   
   
   
   If Not IsNull(rs("total_bultos")) Then
      Me.bultos_totales = rs("total_bultos")
   Else
      Me.bultos_totales = 0
   End If
   If Not IsNull(rs("valor_declarado")) Then
      Me.valor_declarado = rs("valor_declarado")
   Else
      Me.valor_declarado = 0
   End If
   Me.tipoImpresora = rs("tipo_impresora")
   If Not IsNull(rs("contado")) Then
      Me.contado = rs("contado")
   Else
      Me.contado = "N"
   End If
   Me.razonsocial = rs("cliente02")
   Me.direccion = rs("direccion02")
   Me.localidad = rs("localidad02")
   Me.CUIT = rs("cuit02")
   Me.idtipoiva = rs("id_tipo_iva02")
   Me.cant_lineas = rs("cant_lineas")
   Me.estado = rs("estado")
   Me.cae = rs("cae")
   Me.fecha_vto_cae = rs("cae_vence")
   Me.tipoop = rs("tipo_op")
   
   
   Set rs2 = New adodb.Recordset
   q = "select * from g3 where [cod_tipoiva] = " & Me.idtipoiva
   rs2.Open q, cn1
   If Not rs2.EOF And Not rs2.BOF Then
     Me.abreviatura_tipoiva = rs2("abreviatura")
   Else
     Me.abreviatura_tipoiva = "ERr."
   End If
   Set rs2 = Nothing
     
   'tipocompafip
   Set rs2 = New adodb.Recordset
   q = "select * from vta_06 where [sucursal] = " & Me.sucursal & " and [id_tipocomp]=" & Me.idtipocomp
   rs2.Open q, cn1
   If Not rs2.EOF And Not rs2.BOF Then
     Me.cod_afip_a = rs2("cod_afip_a")
     Me.cod_afip_b = rs2("cod_afip_b")
   Else
     Me.cod_afip_a = 1
     Me.cod_afip_b = 6
   
   End If
   Set rs2 = Nothing
   
     
     
     
Else
   Me.numint = 0
End If
Set rs = Nothing
End Sub

Sub borrocomp()
'borrar comprbates de venta
b = 1
If Me.contado = "S" Then
  q = "select * from cyb_03 where [num_int_rbo] = " & Me.numint & " and [estado] <> 'C'"
  Set rs = New adodb.Recordset
  rs.Open q, cn1
  If Not rs.BOF And Not rs.EOF Then
    MsgBox ("Para poder eliminar el Comprobante de Contado todos los cheques ingresados deben estar en Cartera")
    b = 0
  End If
  Set rs = Nothing
End If

If b = 1 Then
    
    'borro comprobantes
      cn1.BeginTrans
      If Me.STOCK <> "N" Then
        q = "select * from vta_03 where [num_int] = " & Me.numint & " and [id_producto] > 1 "
        Set rs = New adodb.Recordset
        rs.Open q, cn1
        While Not rs.EOF
          Set rs2 = New adodb.Recordset
          q = "select * from a2 where [id_producto] = " & rs("id_producto")
          rs2.Open q, cn1, adOpenDynamic, adLockOptimistic, 1
          If Not rs2.BOF And Not rs2.EOF Then
             If Me.STOCK = "E" Then
              s = -Val(rs("cantidad"))
             Else
              s = Val(rs("cantidad"))
             End If
             
             QUERY = "update a2 set [stock] = [stock] + " & s & " where [id_producto] = " & rs2("id_producto")
             cn1.Execute QUERY
            
          End If
          Set rs2 = Nothing
          rs.MoveNext
        Wend
        Set rs = Nothing
      End If
      
      QUERY = "DELETE FROM vta_02 WHERE [num_int] = " & Me.numint
      cn1.Execute QUERY
      
      QUERY = "DELETE FROM vta_03 WHERE [num_int] = " & Me.numint
      cn1.Execute QUERY
        
      QUERY = "DELETE FROM vta_09 WHERE [num_int] = " & Me.numint
      cn1.Execute QUERY
   
      QUERY = "DELETE FROM vta_011 WHERE [num_int] = " & Me.numint
      cn1.Execute QUERY
   
      QUERY = "DELETE FROM vta_012 WHERE [num_int] = " & Me.numint
      cn1.Execute QUERY
   
      QUERY = "DELETE FROM vta_015 WHERE [num_int] = " & Me.numint
      cn1.Execute QUERY
      
      
      QUERY = "DELETE FROM stk_01 WHERE [num_mov_int] = " & Me.numint & " and [modulo] = 'V'"
      cn1.Execute QUERY
      
      QUERY = "DELETE FROM cyb_05 WHERE [num_mov_int] = " & Me.numint & " and [modulo] = 'V'"
      cn1.Execute QUERY
      
      
      
      'formas de pago
            'ch terc
      QUERY = "delete from cyb_03 where [num_int_RBO]= " & Me.numint
      cn1.Execute QUERY
      
        
      
      'FORMAS DE PAGO
      QUERY = "DELETE FROM VTA_04 WHERE [num_int] = " & Me.numint
      cn1.Execute QUERY
      
      
      'mov banco por transferencias
      QUERY = "DELETE FROM cyb_04 WHERE [modulo] = 'V' and [num_mov_int] = " & Me.numint
      cn1.Execute QUERY

      
      
      
      q = "select * from c_02 where [modulo] = 'V' and [num_mov_int] = " & Me.numint
      Set rs = New adodb.Recordset
      rs.Open q, cn1
      If Not rs.EOF And Not rs.BOF Then
        nic = rs("num_interno")
      Else
        nic = 0
      End If
      Set rs = Nothing
      
      If nic <> 0 Then
         QUERY = "DELETE FROM c_02 WHERE [num_interno] = " & nic
         cn1.Execute QUERY
        
         QUERY = "DELETE FROM c_03 WHERE [num_interno] = " & nic
         cn1.Execute QUERY
      End If
      
      
     'busco remitos en facturas
     q = "select * from vta_08 where [id_factura] = " & Me.numint
     Set rs = New adodb.Recordset
     rs.Open q, cn1
     If Not rs.EOF And Not rs.BOF Then
       J = MsgBox("Este comprobante tiene remitos asociados, una vez terminada la operacíon de borrado deberá verificar el estado de los mismos")
     End If
     Set rs = Nothing
      
      
     
     QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
     QUERY = QUERY & " VALUES ('Borra Comp. venta NI " & Me.numint & "' , " & para.id_usuario & ", 'V'," & Me.numint & ", '" & Now & "', '" & Me.abreviatura & " " & Me.letra & " " & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000") & "', 5, " & Me.idcliente & ")"
     cn1.Execute QUERY
      
     
      cn1.CommitTrans
     
     ' MsgBox ("Comprobante Eliminado!!!!")
End If
Exit Sub
borraerr:
MsgBox ("Imposible Eliminar Comprobante. Verifique sus permisos e intente mas tarde. Si aun asi no puede llame al serv. Tecnico(Clase: Comprobantes Vta. Modulo: Borra Comp.")

End Sub
Public Sub borrar()
'borra el comprobante previamente cargado en el objeto
Select Case Me.idtipocomp
 Case Is = 50
    'borro rbo
     q = "select * from cyb_09 where datevalue([fecha]) = datevalue('" & Me.fecha & "')"
     Set rs = New adodb.Recordset
     rs.Open q, cn1
     If Not rs.EOF And Not rs.BOF Then
       If rs("estado") = "A" Then
         Call BORRoRBO
       Else
         MsgBox ("Operacion Imposible de realizar. La caja del dia esta cerrada")
       End If
     Else
       Call BORRoRBO
     End If
     Set rs = Nothing
 Case Is = 400
    MsgBox ("Recuerde que si la Liquidacion tiene Deducciones que fueron cargadas en compras, deberá eliminar manualmente esta liquidacion tambien del modulo de COMPRAS")
    Call borrocomp
 Case Else
      Call borrocomp
End Select
End Sub

Sub borroliq()
 
End Sub
Sub imprimerbo1(ByVal c As Integer)
l = InputBox("Indique cantidad de Copias", , c)
If Val(l) > 0 And Val(l) <= 5 Then
 
 For Y = 1 To Val(l)
   Call imprimeempresa(14)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Times New Roman"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "RECIBO OFICIAL"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(80); "Nro. Comprobante:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FillStyle = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(93); "Fecha:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & Format$(Me.fecha, "dd/mm/yyyy")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(3); "Cliente:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
   Set cl_cli = New Clientes
   cl_cli.carga (Me.idcliente)
   If cl_cli.id > 0 Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print "  " & cl_cli.razonsocial; Tab(92);
   Else
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print "  "; Tab(92);
   End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "Emitido por:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
   Set cl_usuarios = New usuarios
   cl_usuarios.cargar (Me.idusuario)
   If cl_usuarios.idusuario > 0 Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
        Printer.Print "  " & cl_usuarios.denominACION
   Else
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
        Printer.Print
   End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(3); "Direccion:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & cl_cli.direccion
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(3); "Localidad:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & cl_cli.localidad
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(9); "Cuit:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & cl_cli.CUIT

'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "___________________________________________________________________________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "Recibí/mos de " & cl_cli.razonsocial
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "La suma de pesos: *****" & convierte(Me.total) & "*****"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "___________________________________________________________________________________________________"

'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
   Printer.FontSize = 16
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(25); "_______________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(30); "Son $ "; asteriscos(Format$(Me.total, "######0.00"), 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(25); "_______________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "Forma de Pago"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Courier new"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 8
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   linea = 1
   Printer.Print Tab(10); "----------------------------------------------------------------------------------------------"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "Forma de Pago       Num.Ch.    Detalle/Banco               Fecha Dif.    Importe  Sucursal"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "----------------------------------------------------------------------------------------------"
   Set rs = New adodb.Recordset
   q = "Select * from vta_04 where [num_int] = " & Me.numint
   rs.Open q, cn1
   i = Space$(10)
   While Not rs.EOF
    ifp = Format$(rs("id_formapago"), "000")
    FP = Format$(Left$(rs("formapago"), 12), "@@@@@@@@@@@@!")
    nch = Format$(rs("num_ch"), "0000000000")
    db = Format$(Left$(rs("detalle_Banco") & " ", 27), "@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    s = Format$(Left$(rs("sucursal") & " ", 10), "@@@@@@@@@@!")
    fd = Format$(rs("fecha_dif"), "dd/mm/yyyy")
    RSet i = Format$(rs("importe"), "######0.00")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(10); "(" & ifp & ") " & FP & "  " & nch & " " & db & " " & fd & " " & i & " " & s
    rs.MoveNext
    linea = linea + 1
   Wend
   Set rs = Nothing
   
   Printer.FontSize = 8
   Printer.Print
   Printer.Print
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "Comprobantes Aplicados"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Courier new"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "-------------------------------------------------------"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "Fecha        Comprobante            Importe   Obs."
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "-------------------------------------------------------"
   Set rs = New adodb.Recordset
   q = "Select * from vta_02 where [estado_pago] = 'P' and [recibo_pago] = '" & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000") & "' AND [ID_TIPOCOMP] <= 10 "
   rs.Open q, cn1
   TCOMP = 0
   While Not rs.EOF
    Set rs1 = New adodb.Recordset
    q = "select * from vta_06 where [id_tipocomp] = " & rs("id_tipocomp") & " and [sucursal] = " & rs("sucursal_INGRESO")
    rs1.Open q, cn1
    If Not rs1.EOF And Not rs1.BOF Then
      d = rs1("abreviatura")
    Else
      d = " "
    End If
    Set rs1 = Nothing
    nc = "[" & Format$(rs("id_tipocomp"), "000") & "] " & rs("letra") & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
    fd = Format$(rs("fecha"), "dd/mm/yyyy")
    RSet i = Format$(rs("total"), "######0.00")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(10); fd & "  " & nc & "  " & i & "  " & d
    
    If rs("cta_cte") = "D" Then
      TCOMP = TCOMP + rs("total")
    Else
       TCOMP = TCOMP - rs("total")
    End If
       
    rs.MoveNext
    linea = linea + 1
   Wend
   Set rs = Nothing
   Printer.Print Tab(10); "  TOTAL COMPROBANTES APLICADOS:     " & Format$(TCOMP, "#######0.00")
   
   
   Printer.Print
   Printer.Print
   Printer.FontBold = True
   Printer.Print Tab(10); "Retenciones Ingresadas"
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Courier new"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 8
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "-------------------------------------------------------"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "Fecha        Retencion              Importe   Obs."
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "-------------------------------------------------------"
   Set rs = New adodb.Recordset
   q = "Select * from vta_02 where [estado_pago] = 'P' and [recibo_pago] = '" & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000") & "' AND [ID_TIPOCOMP] >= 100 "
   rs.Open q, cn1
   TRET = 0
   While Not rs.EOF
    Set rs1 = New adodb.Recordset
    q = "select * from vta_06 where [id_tipocomp] = " & rs("id_tipocomp") & " and [sucursal] = " & rs("sucursal_INGRESO")
    rs1.Open q, cn1
    If Not rs1.EOF And Not rs1.BOF Then
      d = rs1("abreviatura")
    Else
      d = " "
    End If
    Set rs1 = Nothing
    nc = "[" & Format$(rs("id_tipocomp"), "000") & "] " & rs("letra") & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
    fd = Format$(rs("fecha"), "dd/mm/yyyy")
    RSet i = Format$(rs("total"), "######0.00")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(10); fd & "  " & nc & "  " & i & "  " & d
    TRET = TRET + rs("TOTAL")
    rs.MoveNext
    linea = linea + 1
   Wend
   Set rs = Nothing
   Printer.Print Tab(10); "  TOTAL RETENCIONES INGRESADAS:       " & Format$(TRET, "#######0.00")
   
   
   Printer.FontName = "Times New Roman"
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.FontSize = 10
   Printer.Print
   Printer.Print
   Printer.Print Tab(3); "Observaciones:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  "; Me.observaciones
   Printer.Print
   Printer.PaintPicture gen_logo.Picture2.Picture, 6100, 10100 + (linea * 100)
   
   Set cl_cli = Nothing
   If Y < Val(l) Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.NewPage
   End If
  Next Y
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
  Printer.EndDoc
End If
End Sub


Sub imprimerboW(ByVal c As Integer)
l = InputBox("Indique cantidad de Copias", , c)
If Val(l) > 0 And Val(l) <= 5 Then
 
 For Y = 1 To Val(l)
   Printer.Print
   Printer.FontName = "Times New Roman"
   Printer.FontSize = 12
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print "RECIBO"
   Printer.FontSize = 10
   Printer.Print Tab(80); "Nro. Comprobante:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.FillStyle = False
   Printer.Print "  " & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000")
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print Tab(93); "Fecha:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.Print "  " & Format$(Me.fecha, "dd/mm/yyyy")
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print
   Printer.Print Tab(3); "Cliente:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Set cl_cli = New Clientes
   cl_cli.carga (Me.idcliente)
   If cl_cli.id > 0 Then
      Printer.Print "  " & cl_cli.razonsocial; Tab(92);
   Else
      Printer.Print "  "; Tab(92);
   End If
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print Tab(3); "Direccion:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.Print "  " & cl_cli.direccion
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print Tab(3); "Localidad:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.Print "  " & cl_cli.localidad
   Printer.Print "___________________________________________________________________________________________________"
   Printer.Print
   Printer.FontSize = 12
   Printer.Print "Recibí/mos de " & cl_cli.razonsocial
   Printer.Print "La suma de pesos: *****" & convierte(Me.total) & "*****"
   Printer.FontSize = 10
   Printer.Print "___________________________________________________________________________________________________"

   Printer.FontSize = 16
   Printer.Print Tab(25); "_______________________________"
   Printer.Print Tab(30); "Son $ "; asteriscos(Format$(Me.total, "######0.00"), 12)
   Printer.Print Tab(25); "_______________________________"
   Printer.FontSize = 10
   Printer.Print
   Printer.FontBold = True
   Printer.Print Tab(10); "Forma de Pago"
   Printer.FontBold = False
   Printer.FontName = "Courier new"
   Printer.FontSize = 8
   linea = 1
   Printer.Print Tab(10); "----------------------------------------------------------------------------------------------"
   Printer.Print Tab(10); "Forma de Pago       Num.Ch.    Detalle/Banco               Fecha Dif.    Importe  Sucursal"
   Printer.Print Tab(10); "----------------------------------------------------------------------------------------------"
   Set rs = New adodb.Recordset
   q = "Select * from vta_04 where [num_int] = " & Me.numint
   rs.Open q, cn1
   i = Space$(10)
   While Not rs.EOF
    ifp = Format$(rs("id_formapago"), "000")
    FP = Format$(Left$(rs("formapago"), 12), "@@@@@@@@@@@@!")
    nch = Format$(rs("num_ch"), "0000000000")
    db = Format$(Left$(rs("detalle_Banco") & " ", 27), "@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    s = Format$(Left$(rs("sucursal") & " ", 10), "@@@@@@@@@@!")
    fd = Format$(rs("fecha_dif"), "dd/mm/yyyy")
    RSet i = Format$(rs("importe"), "######0.00")
    Printer.Print Tab(10); "(" & ifp & ") " & FP & "  " & nch & " " & db & " " & fd & " " & i & " " & s
    rs.MoveNext
    linea = linea + 1
   Wend
   Set rs = Nothing
   
   Printer.FontSize = 8
   Printer.Print
   Printer.FontBold = True
   Printer.Print Tab(10); "Comprobantes Aplicados"
   Printer.FontBold = False
   Printer.FontName = "Courier new"
   
   Printer.Print Tab(10); "-------------------------------------------------------"
   Printer.Print Tab(10); "Fecha        Comprobante            Importe   Obs."
   Printer.Print Tab(10); "-------------------------------------------------------"
   Set rs = New adodb.Recordset
   q = "Select * from vta_02 where [estado_pago] = 'P' and [recibo_pago] = '" & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000") & "' AND [ID_TIPOCOMP] <= 10 "
   rs.Open q, cn1
   TCOMP = 0
   While Not rs.EOF
    Set rs1 = New adodb.Recordset
    q = "select * from vta_06 where [id_tipocomp] = " & rs("id_tipocomp") & " and [sucursal] = " & rs("sucursal_INGRESO")
    rs1.Open q, cn1
    If Not rs1.EOF And Not rs1.BOF Then
      d = rs1("abreviatura")
    Else
      d = " "
    End If
    Set rs1 = Nothing
    nc = "[" & Format$(rs("id_tipocomp"), "000") & "] " & rs("letra") & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
    fd = Format$(rs("fecha"), "dd/mm/yyyy")
    RSet i = Format$(rs("total"), "######0.00")
    Printer.Print Tab(10); fd & "  " & nc & "  " & i & "  " & d
    
    If rs("cta_cte") = "D" Then
      TCOMP = TCOMP + rs("total")
    Else
       TCOMP = TCOMP - rs("total")
    End If
       
    rs.MoveNext
    linea = linea + 1
   Wend
   Set rs = Nothing
   
   'saldo = cl_cli.saldo(True, Now, True)
   'Printer.FontName = "Times New Roman"
   'Printer.FontSize = 10
   Printer.Print
   'Printer.Print " **** Su saldo a la fecha a quedado de --> "; Format$(saldo, "#####0.00"); " ****"
   Printer.Print "                                                                                                "; "                                       ________________________"
   Set cl_cli = Nothing
   If Y < Val(l) Then
      Printer.NewPage
   End If
  Next Y
  Printer.EndDoc
End If
End Sub

Sub imprimeminutafiscal()
l = InputBox("Indique cantidad de Copias", , 1)
If Val(l) > 0 And Val(l) <= 5 Then
 
 For Y = 1 To Val(l)
   Call imprimeempresa(14)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Times New Roman"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "MINUTA por Recibo: ";
      
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FI
   Printer.Print "  " & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000")
   Printer.FontSize = 9
   Printer.Print Tab(10); "Documento no válido como Factura ni como Recibo de pago ";
   Printer.Print
   Printer.Print
   Printer.FontBold = True
   Printer.Print Tab(10); "Forma de Pago"
   Printer.FontBold = False
   Printer.FontName = "Courier new"
   Printer.FontSize = 8
   Printer.Print Tab(10); "----------------------------------------------------------------------------------------------"
   Printer.Print Tab(10); "Forma de Pago       Num.Ch.    Detalle/Banco               Fecha Dif.    Importe  Sucursal"
   Printer.Print Tab(10); "----------------------------------------------------------------------------------------------"
   Set rs = New adodb.Recordset
   q = "Select * from vta_04 where [num_int] = " & Me.numint
   rs.Open q, cn1
   i = Space$(10)
   While Not rs.EOF
    ifp = Format$(rs("id_formapago"), "000")
    FP = Format$(Left$(rs("formapago"), 12), "@@@@@@@@@@@@!")
    nch = Format$(rs("num_ch"), "0000000000")
    db = Format$(Left$(rs("detalle_Banco") & " ", 27), "@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    s = Format$(Left$(rs("sucursal") & " ", 10), "@@@@@@@@@@!")
    fd = Format$(rs("fecha_dif"), "dd/mm/yyyy")
    RSet i = Format$(rs("importe"), "######0.00")
    Printer.Print Tab(10); "(" & ifp & ") " & FP & "  " & nch & " " & db & " " & fd & " " & i & " " & s
    rs.MoveNext
   Wend
   Set rs = Nothing
   
   Printer.FontSize = 8
   Printer.Print
   Printer.Print
   Printer.FontBold = True
   Printer.Print Tab(10); "Comprobantes Aplicados"
   Printer.FontBold = False
   Printer.FontName = "Courier new"
   Printer.Print Tab(10); "-------------------------------------------------------"
   Printer.Print Tab(10); "Fecha        Comprobante            Importe   Obs."
   Printer.Print Tab(10); "-------------------------------------------------------"
   Set rs = New adodb.Recordset
   q = "Select * from vta_02 where [estado_pago] = 'P' and [recibo_pago] = '" & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000") & "' AND [ID_TIPOCOMP] <= 10 "
   rs.Open q, cn1
   TCOMP = 0
   While Not rs.EOF
    Set rs1 = New adodb.Recordset
    q = "select * from vta_06 where [id_tipocomp] = " & rs("id_tipocomp") & " and [sucursal] = " & rs("sucursal_INGRESO")
    rs1.Open q, cn1
    If Not rs1.EOF And Not rs1.BOF Then
      d = rs1("abreviatura")
    Else
      d = " "
    End If
    Set rs1 = Nothing
    nc = "[" & Format$(rs("id_tipocomp"), "000") & "] " & rs("letra") & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
    fd = Format$(rs("fecha"), "dd/mm/yyyy")
    RSet i = Format$(rs("total"), "######0.00")
    Printer.Print Tab(10); fd & "  " & nc & "  " & i & "  " & d
    
    If rs("cta_cte") = "D" Then
      TCOMP = TCOMP + rs("total")
    Else
       TCOMP = TCOMP - rs("total")
    End If
       
    rs.MoveNext
   Wend
   Set rs = Nothing
   Printer.Print Tab(10); "  TOTAL COMPROBANTES APLICADOS:     " & Format$(TCOMP, "#######0.00")
   
   Printer.Print
   Printer.Print
   Printer.FontBold = True
   Printer.Print Tab(10); "Retenciones Ingresadas"
   Printer.FontBold = False
   Printer.FontBold = False
   Printer.FontName = "Courier new"
   Printer.FontSize = 8
   Printer.Print Tab(10); "-------------------------------------------------------"
   Printer.Print Tab(10); "Fecha        Retencion              Importe   Obs."
   Printer.Print Tab(10); "-------------------------------------------------------"
   Set rs = New adodb.Recordset
   q = "Select * from vta_02 where [estado_pago] = 'P' and [recibo_pago] = '" & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000") & "' AND [ID_TIPOCOMP] >= 100 "
   rs.Open q, cn1
   TRET = 0
   While Not rs.EOF
    Set rs1 = New adodb.Recordset
    q = "select * from vta_06 where [id_tipocomp] = " & rs("id_tipocomp") & " and [sucursal] = " & rs("sucursal_INGRESO")
    rs1.Open q, cn1
    If Not rs1.EOF And Not rs1.BOF Then
      d = rs1("abreviatura")
    Else
      d = " "
    End If
    Set rs1 = Nothing
    nc = "[" & Format$(rs("id_tipocomp"), "000") & "] " & rs("letra") & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
    fd = Format$(rs("fecha"), "dd/mm/yyyy")
    RSet i = Format$(rs("total"), "######0.00")
    Printer.Print Tab(10); fd & "  " & nc & "  " & i & "  " & d
    TRET = TRET + rs("TOTAL")
    rs.MoveNext
   Wend
   Set rs = Nothing
   Printer.Print Tab(10); "  TOTAL RETENCIONES INGRESADAS:       " & Format$(TRET, "#######0.00")
   
   
   
   Set cl_cli = Nothing
   If Y < Val(l) Then
      Printer.NewPage
   End If
  Next Y
  Printer.EndDoc
End If
End Sub

Sub imprimerbo4(ByVal c As Integer)
l = InputBox("Indique cantidad de Copias", , c)
If Val(l) > 0 And Val(l) <= 5 Then
 
 For Y = 1 To Val(l)
   Printer.Print
   Printer.Print
   Printer.Print
   Printer.Print
   Printer.FontName = "Times New Roman"
   Printer.FontSize = 12
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print Tab(65); "RECIBO OFICIAL"
   Printer.FontSize = 10
   Printer.Print
   Printer.Print
   
   Printer.Print Tab(90); "Nro. Comprobante:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.FillStyle = False
   Printer.Print "  " & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000")
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print Tab(103); "Fecha:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.Print "  " & Format$(Me.fecha, "dd/mm/yyyy")
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print
   Printer.Print
   Printer.Print
   Printer.Print
   Printer.Print
   Printer.Print Tab(13); "Cliente:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Set cl_cli = New Clientes
   cl_cli.carga (Me.idcliente)
   If cl_cli.id > 0 Then
      Printer.Print "  " & cl_cli.razonsocial; Tab(102);
   Else
      Printer.Print "  "; Tab(102);
   End If
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print "Emitido por:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Set cl_usuarios = New usuarios
   cl_usuarios.cargar (Me.idusuario)
   If cl_usuarios.idusuario > 0 Then
        Printer.Print "  " & cl_usuarios.denominACION
   Else
        Printer.Print
   End If
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print Tab(13); "Direccion:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.Print "  " & cl_cli.direccion
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print Tab(13); "Localidad:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.Print "  " & cl_cli.localidad
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print Tab(19); "Cuit:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.Print "  " & cl_cli.CUIT
   Printer.Print
   Printer.Print Tab(13); "___________________________________________________________________________________________________"
   Printer.Print
   Printer.FontSize = 12
   Printer.Print Tab(13); "Recibí/mos de " & cl_cli.razonsocial
   Printer.Print
   Printer.Print Tab(13); "La suma de pesos: *****" & convierte(Me.total) & "*****"
   Printer.FontSize = 10
   Printer.Print Tab(13); "___________________________________________________________________________________________________"

   Printer.FontSize = 16
   Printer.Print
   Printer.Print Tab(25); "_______________________________"
   Printer.Print Tab(30); "Son $ "; asteriscos(Format$(Me.total, "######0.00"), 12)
   Printer.Print Tab(25); "_______________________________"
   Printer.FontSize = 10
   Printer.Print
   Printer.Print
   Printer.FontBold = True
   Printer.Print Tab(13); "Forma de Pago"
   Printer.FontBold = False
   Printer.FontName = "Courier new"
   Printer.FontSize = 8
   Printer.Print Tab(13); "----------------------------------------------------------------------------------------------"
   Printer.Print Tab(13); "Forma de Pago       Num.Ch.    Detalle/Banco               Fecha Dif.    Importe  Sucursal"
   Printer.Print Tab(13); "----------------------------------------------------------------------------------------------"
   Set rs = New adodb.Recordset
   q = "Select * from vta_04 where [num_int] = " & Me.numint
   rs.Open q, cn1
   i = Space$(10)
   While Not rs.EOF
    ifp = Format$(rs("id_formapago"), "000")
    FP = Format$(Left$(rs("formapago"), 12), "@@@@@@@@@@@@!")
    nch = Format$(rs("num_ch"), "0000000000")
    db = Format$(Left$(rs("detalle_Banco") & " ", 27), "@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    s = Format$(Left$(rs("sucursal") & " ", 10), "@@@@@@@@@@!")
    fd = Format$(rs("fecha_dif"), "dd/mm/yyyy")
    RSet i = Format$(rs("importe"), "######0.00")
    Printer.Print Tab(13); "(" & ifp & ") " & FP & "  " & nch & " " & db & " " & fd & " " & i & " " & s
    rs.MoveNext
   Wend
   Set rs = Nothing
   
   Printer.FontSize = 10
   Printer.Print
   Printer.Print
   Printer.FontBold = True
   Printer.Print Tab(13); "Comprobantes Aplicados"
   Printer.FontBold = False
   Printer.FontName = "Courier new"
   Printer.FontSize = 8
   Printer.Print Tab(13); "-------------------------------------------------------"
   Printer.Print Tab(13); "Fecha        Comprobante               Importe     Obs."
   Printer.Print Tab(13); "-------------------------------------------------------"
   Set rs = New adodb.Recordset
   q = "Select * from vta_02 where [estado_pago] = 'P' and [recibo_pago] = '" & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000") & "'"
   rs.Open q, cn1
   While Not rs.EOF
    Set rs1 = New adodb.Recordset
    q = "select * from vta_06 where [sucursal] = " & rs("sucursal_ingreso") & " and [id_tipocomp] = " & rs("id_tipocomp")
    rs1.Open q, cn1
    If Not rs1.EOF And Not rs1.BOF Then
      d = rs1("abreviatura")
    Else
      d = " "
    End If
    Set rs1 = Nothing
    nc = "[" & Format$(rs("id_tipocomp"), "000") & "] " & rs("letra") & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
    fd = Format$(rs("fecha"), "dd/mm/yyyy")
    If rs("cta_cte") = "D" Then
      RSet i = Format$(rs("total"), "######0.00")
    Else
      RSet i = Format$(-rs("total"), "######0.00")
    End If
    Printer.Print Tab(13); fd & "  " & nc & "  " & i & "  " & d
    rs.MoveNext
   Wend
   Set rs = Nothing
   
   Printer.FontName = "Times New Roman"
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.FontSize = 10
   Printer.Print
   Printer.Print Tab(13); "Observaciones:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.Print "  "; Me.observaciones
   Printer.Print
   Printer.Print
   Printer.Print Tab(75); "______________________________"
   Printer.Print Tab(75); "   por "; glo.nombrecli
   
   Set cl_cli = Nothing
   If Y < Val(l) Then
      Printer.NewPage
   End If
  Next Y
  Printer.EndDoc
End If
End Sub


Sub imprimerbo2(ByVal c As Integer)
'degser
l = InputBox("Indique cantidad de Copias", , c)
If Val(l) > 0 And Val(l) <= 5 Then
 
 For Y = 1 To Val(l)
   'Call imprimeempresa(14)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Line Step(1000, 1)-Step(4000, 1)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Times New Roman"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(11); "RECIBO OFICIAL"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Line Step(1000, 1)-Step(4000, 1)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
   Printer.Print
   Printer.Print Tab(87); "Nro. Comprobante:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FillStyle = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(100); "Fecha:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & Format$(Me.fecha, "dd/mm/yyyy")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "Cliente.......:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 14
   Set cl_cli = New Clientes
   cl_cli.carga (Me.idcliente)
   If cl_cli.id > 0 Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print "  " & cl_cli.razonsocial
   Else
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print "  "; Tab(70)
   End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "Direccion.....:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & cl_cli.direccion & "   CP. " & cl_cli.cp
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "Localidad.....:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & cl_cli.localidad
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "Cuit..........:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & cl_cli.CUIT
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Times New Roman"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "___________________________________________________________________________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "Recibí/mos de " & cl_cli.razonsocial
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "La suma de pesos: *****" & convierte(Me.total) & "*****"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "___________________________________________________________________________________________________"

'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 16
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(25); "_______________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(30); "Son $ "; asteriscos(Format$(Me.total, "######0.00"), 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(25); "_______________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "  Forma de Pago"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Courier new"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 8
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "----------------------------------------------------------------------------------------------"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "Forma de Pago       Num.Ch.    Detalle/Banco               Fecha Dif.    Importe  Sucursal    "
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "----------------------------------------------------------------------------------------------"
   Set rs = New adodb.Recordset
   q = "Select * from vta_04 where [num_int] = " & Me.numint
   rs.Open q, cn1
   i = Space$(10)
   While Not rs.EOF
    ifp = Format$(rs("id_formapago"), "000")
    FP = Format$(Left$(rs("formapago"), 12), "@@@@@@@@@@@@!")
    nch = Format$(rs("num_ch"), "0000000000")
    db = Format$(Left$(rs("detalle_Banco") & " ", 27), "@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    s = Format$(Left$(rs("sucursal") & " ", 11), "@@@@@@@@@@@!")
    fd = Format$(rs("fecha_dif"), "dd/mm/yyyy")
    nix = Format$(rs("num_int_fp"), "000000")
    RSet i = Format$(rs("importe"), "######0.00")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(12); "(" & ifp & ") " & FP & "  " & nch & " " & db & " " & fd & " " & i & " " & s & " " & nix
    rs.MoveNext
   Wend
   Set rs = Nothing
   
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "Comprobantes Aplicados"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Courier new"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 8
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "-------------------------------------------------------"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "Fecha        Comprobante               Importe     Obs."
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "-------------------------------------------------------"
   Set rs = New adodb.Recordset
   q = "Select * from vta_010 where [num_int_rbo] = " & Me.numint
   rs.Open q, cn1
   While Not rs.EOF
    Set rs1 = New adodb.Recordset
    q = "select * from vta_02, vta_06 where [num_int] = " & rs("num_int_comp") & " and vta_02.[id_tipocomp] = vta_06.[id_tipocomp] and  vta_02.[sucursal_ingreso] = vta_06.[sucursal]"
    rs1.Open q, cn1
    If Not rs1.EOF And Not rs1.BOF Then
      abr = Format$(rs1.Fields("abreviatura"), "@@@@@!")
      nc = "[" & abr & "] " & rs1("letra") & Format$(rs1("vta_02.sucursal"), "0000") & "-" & Format$(rs1("num_comp"), "00000000")
      fd = Format$(rs1("fecha"), "dd/mm/yyyy")
      tot = rs("importe_pagado")
      If rs1("cta_cte") = "D" Then
        RSet i = Format$(tot, "######0.00")
      Else
        RSet i = Format$(-tot, "######0.00")
      End If
      
      If rs("saldo_comprobante") > 0 Then
        msgp = " (Parcial) "
      Else
        msgp = "           "
      End If
      
      Printer.Print Tab(12); fd & "  " & nc & "  " & i & "  " & d & msgp
    End If
    Set rs1 = Nothing
    rs.MoveNext
   Wend
   Set rs = Nothing
   
   
   
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Times New Roman"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(14); "Observaciones:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(14); "  "; Me.observaciones
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(75); "______________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(76); "por "; glo.nombrecli
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 6
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(15); "___________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(15); "  Emitido por: "; para.usuario
   
   Set cl_cli = Nothing
   If Y < Val(l) Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.NewPage
   End If
  Next Y
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
  Printer.EndDoc
End If
End Sub

Sub imprimerbo3(ByVal c As Integer)
l = InputBox("Indique cantidad de Copias", , c)
If Val(l) > 0 And Val(l) <= 5 Then
 
 For Y = 1 To Val(l)
   'Call imprimeempresa(14)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Times New Roman"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
Printer.PaintPicture gen_logo.Picture1.Picture, 0, 0
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Line Step(1000, 1)-Step(4000, 1)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Times New Roman"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(11); "RECIBO OFICIAL"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Line Step(1000, 1)-Step(4000, 1)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(87); "Nro. Comprobante:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FillStyle = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(100); "Fecha:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & Format$(Me.fecha, "dd/mm/yyyy")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "Cliente.......:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 14
   Set cl_cli = New Clientes
   cl_cli.carga (Me.idcliente)
   If cl_cli.id > 0 Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print "  " & cl_cli.razonsocial
   Else
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print "  "; Tab(70)
   End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "Direccion.....:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & cl_cli.direccion & "   CP. " & cl_cli.cp
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "Localidad.....:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & cl_cli.localidad
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "Cuit..........:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & cl_cli.CUIT
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Times New Roman"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "___________________________________________________________________________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "Recibímos de " & cl_cli.razonsocial
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "La suma de pesos: *****" & convierte(Me.total) & "*****"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "___________________________________________________________________________________________________"

'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 16
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(25); "_______________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(30); "Son $ "; asteriscos(Format$(Me.total, "######0.00"), 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(25); "_______________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "  Forma de Pago"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Courier new"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 8
   linea = 1
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "-----------------------------------------------------------------------------------------------------"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "Forma de Pago       Num.Ch.    Detalle/Banco                                  Fecha Dif.    Importe  "
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "-----------------------------------------------------------------------------------------------------"
   Set rs = New adodb.Recordset
   q = "Select * from vta_04 where [num_int] = " & Me.numint
   rs.Open q, cn1
   i = Space$(10)
   While Not rs.EOF
    
    ifp = Format$(rs("id_formapago"), "000")
    FP = Format$(Left$(rs("formapago"), 12), "@@@@@@@@@@@@!")
    nch = Format$(rs("num_ch"), "0000000000")
    db = Format$(Left$(rs("detalle_Banco") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    s = Format$(Left$(rs("sucursal") & " ", 11), "@@@@@@@@@@@!")
    fd = Format$(rs("fecha_dif"), "dd/mm/yyyy")
    nix = Format$(rs("num_int_fp"), "000000")
    
    RSet i = Format$(rs("importe"), "######0.00")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(12); "(" & ifp & ") " & FP & "  " & nch & " " & db & " " & fd & " " & i
    linea = linea + 1
    rs.MoveNext
   Wend
   Set rs = Nothing
   
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "Comprobantes Aplicados"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Courier new"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 8
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "-------------------------------------------------------"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "Fecha        Comprobante               Imp.Aplic.    Obs."
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "-------------------------------------------------------"
   Set rs = New adodb.Recordset
   q = "Select * from vta_010 where [num_int_rbo] = " & Me.numint
   rs.Open q, cn1
   While Not rs.EOF
    Set rs1 = New adodb.Recordset
    q = "select * from vta_02, vta_06 where [num_int] = " & rs("num_int_comp") & " and vta_02.[id_tipocomp] = vta_06.[id_tipocomp] and  vta_02.[sucursal_ingreso] = vta_06.[sucursal]"
    rs1.Open q, cn1
    If Not rs1.EOF And Not rs1.BOF Then
      abr = Format$(rs1.Fields("abreviatura"), "@@@@@!")
      nc = "[" & abr & "] " & rs1("letra") & Format$(rs1("vta_02.sucursal"), "0000") & "-" & Format$(rs1("num_comp"), "00000000")
      fd = Format$(rs1("fecha"), "dd/mm/yyyy")
      tot = rs("importe_pagado")
      If rs1("cta_cte") = "D" Then
        RSet i = Format$(tot, "######0.00")
      Else
        RSet i = Format$(-tot, "######0.00")
      End If
      
      If rs("saldo_comprobante") > 0 Then
        msgp = " (Parcial) "
      Else
        msgp = "           "
      End If
      
      Printer.Print Tab(12); fd & "  " & nc & "  " & i & "  " & d & msgp
      linea = linea + 1
    End If
    Set rs1 = Nothing
    rs.MoveNext
   Wend
   Set rs = Nothing
   
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Times New Roman"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(14); "Observaciones:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(14); "  "; Me.observaciones
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
  
   Printer.Print Tab(75); "______________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(76); "por "; glo.nombrecli
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
   Printer.PaintPicture gen_logo.Picture2.Picture, 6100, 10100 + (linea * 100)
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 6
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(15); "___________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(15); "  Emitido por: "; para.usuario
   
   Set cl_cli = Nothing
   If Y < Val(l) Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.NewPage
   End If
  Next Y
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
  Printer.EndDoc
  
End If
End Sub

Public Sub generapdf(ByVal c As Integer)
    
 If c > 0 And c < 4 Then
    
   Set rsr = New adodb.Recordset
   q = "select * from vta_08, vta_02 where [id_factura] = " & Me.numint & " and [id_remito] = [num_int]"
   rsr.Open q, cn1
   tr = ""
   While Not rsr.EOF
     If rsr("id_tipocomp") = 45 Then
      tc = "R"
     Else
      tc = "D"
     End If
     tr = tr & Format$(rsr("sucursal"), "0000") & "-" & Format$(rsr("num_comp"), "00000000") & tc & " * "
     rsr.MoveNext
   Wend
   Set rsr = Nothing
  
    
    
    
    
    Set rs = New adodb.Recordset
    q = "select dias_pago_cc, path_facturas from fe_01 where id = 1"
    rs.Open q, cn1
    If Not rs.EOF And Not rs.BOF Then
       diaspagocc = rs("dias_pago_cc")
       pathfacturas = rs("path_facturas")
    Else
       MsgBox ("Advertencia: Error en los parametros de factura electronica tabla fe_01")
       diaspagocc = 0
       pathfacturas = "C:\"
   End If
   Set rs = Nothing
    
    
    
    
    
    Dim PyFEPDF As Object
    
    'On Error GoTo ManejoError
    
    
    ' Crear objeto interface para generación de F.E. en PDF
    Set PyFEPDF = CreateObject("PyFEPDF")
        
    
    ' CUIT del emisor
    PyFEPDF.CUIT = glo.CUIT
    
    ' Formato numérico (cantidad decimales)
    PyFEPDF.FmtCantidad = "0.2"
    PyFEPDF.FmtPrecio = "0.4"
    
    If Me.letra = "A" Then
        tipo_cbte = Me.cod_afip_a
    Else
        tipo_cbte = Me.cod_afip_b
    End If
    punto_vta = Me.sucursal    ' prefijo
    cbte_nro = Me.numcomp ' número de factura
    fecha = Me.fecha
    concepto = 3 'Me.tipoop
    
    ' datos del cliente:
    If Me.idtipoiva = 3 Then
      tipo_doc = 96
    Else
      tipo_doc = 80
    End If
    nro_doc = Me.CUIT
      
    '--------------------------
    
    nombre_cliente = Me.razonsocial
    domicilio_cliente = Me.direccion
    pais_dst_cmp = 200 ' código para exportación
    id_impositivo = Me.abreviatura_tipoiva
    
    ' totales del comprobante:
    imp_total = Format$(Me.total, "#####0.00")
    imp_tot_conc = "0.00"
    imp_neto = Format$(Me.subtotal, "#####0.00")
    imp_iva = Format$(Me.iva, "#####0.00")
    imp_trib = ""
    imp_op_ex = ""
    imp_subtotal = Format$(Me.subtotal, "#####0.00")
    descuento = 0
    fecha_cbte = Format$(fecha, "dd/mm/yyyy")
    fecha_venc_pago = Format$(DateValue(Me.fecha) + diaspagocc, "dd/mm/yyyy")
    ' Fechas del período del servicio facturado
    fecha_serv_desde = Format$(fecha, "dd/mm/yyyy")
    fecha_serv_hasta = Format$(fecha, "dd/mm/yyyy")
      
    obs_generales = ""
    obs_comerciales = ""
    
    If Me.moneda = "P" Then
       moneda_id = "PES"
       moneda_ctz = Format$(Me.cotizaciondolar, "####0.000")
      
    Else
       moneda_id = "DOL"
       moneda_ctz = Format$(Me.cotizaciondolar, "####0.000")
       'obs_comerciales = "Comprobante en dolares estadounidenses Cotización: " & moneda_ctz & " $ por  dolar"

    End If
    
    
    
    If Me.contado = "S" Then
       forma_pago = "Contado"
    Else
       forma_pago = "Cuenta Corriente"
    End If
    incoterms = "" ' termino de comercio exterior para exportación
    idioma_cbte = 1   ' idioma para exportación (no usado por el momento)
    ' motivo de observación (F136 y otros - RG2485/08 Art. 30 inc. c):
    motivo_obs = ""

    ' Código de Autorización Electrónica y fecha de vencimiento:
    ' (para facturas tradicionales, no imprimir el CAE ni código de barras)
    cae = Format$(Me.cae)
    fch_vto_cae = Mid$(Format$(Me.fecha_vto_cae, "dd/mm/yyyy"), 7, 4) & Mid$(Format$(Me.fecha_vto_cae, "dd/mm/yyyy"), 4, 2) & Mid$(Format$(Me.fecha_vto_cae, "dd/mm/yyyy"), 1, 2)
   
    
    ' Creo la factura (internamente en la interfaz)
    ok = PyFEPDF.CrearFactura( _
        concepto, tipo_doc, nro_doc, tipo_cbte, punto_vta, _
        cbte_nro, imp_total, imp_tot_conc, imp_neto, _
        imp_iva, imp_trib, imp_op_ex, fecha_cbte, fecha_venc_pago, _
        fecha_serv_desde, fecha_serv_hasta, _
        moneda_id, moneda_ctz, cae, fch_vto_cae, id_impositivo, _
        nombre_cliente, domicilio_cliente, pais_dst_cmp, _
        obs_comerciales, obs_generales, forma_pago, incoterms, _
        idioma_cbte, motivo_obs, descuento)
    
    
    
    ok = PyFEPDF.EstablecerParametro("localidad_cliente", Me.localidad)
    ok = PyFEPDF.EstablecerParametro("provincia_cliente", " ")

    ' Logotipo AFIP Comprobante Autorizado (cambiar resultado="A")
    ok = PyFEPDF.EstablecerParametro("resultado", "A")

    ' Agregar comprobantes asociados (si es una NC/ND):
    'tipo = 19
    'pto_vta = 2
    'nro = 1234
    'pyfepdf.AgregarCmpAsoc(tipo, pto_vta, nro)
    
    
    ' Agrego subtotales de IVA (uno por alicuota)
   Set rs = New adodb.Recordset
   q = "Select * from vta_09, g4 where [num_int] = " & Me.numint & " and tasa_iva = tasa"
   rs.Open q, cn1
    While Not rs.EOF
      iva_id = rs("id_afip")      ' código para alícuota del 21%
      base_imp = Format(rs("neto"), "######0.00")  ' importe neto sujeto a esta alícuota
      importe = Format(rs("iva"), "######0.00")    ' importe liquidado de iva
      ok = PyFEPDF.AgregarIva(iva_id, base_imp, importe)
      
      rs.MoveNext
    Wend
    
    
        
    
    ' Agregar cada impuesto (por ej. IIBB, retenciones, percepciones, etc.):
    'tributo_id = 99         ' codigo para 99-otros tributos
    'Desc = "Impuesto Municipal Matanza"
    'base_imp = "100.00"     ' importe sujeto a este tributo
    'alic = "1.00"           ' alicuota (porcentaje) de este tributo
    'importe = "1.00"        ' importe liquidado de este tributo
    'ok = PyFEPDF.AgregarTributo(tributo_id, Desc, base_imp, alic, importe)

    ' Agrego detalles de cada item de la factura:
Set rs = New adodb.Recordset
q = "Select * from vta_03, g4 where [num_int] = " & Me.numint & " and tasaiva = tasa"
rs.Open q, cn1
While Not rs.EOF
    u_mtx = 123456              ' unidades
    cod_mtx = 1234567890123#    ' código de barras
    CODIGO = "[" & Format$(rs("id_producto"), "00000") & "]"            ' codigo interno a imprimir (ej. "articulo")
    ds = rs("vta_03.descripcion")
    qty = rs("cantidad")                     ' cantidad
    umed = 7                    ' código de unidad de medida (ej. 7 para "unidades")
    If Me.letra = "A" Then
           precio = rs("pu")
           importe = rs("pu") * rs("cantidad")
           imp_iva = 0 '(rs("pu_final") - rs("pu"))
           
    Else
           precio = Format$(rs("pu_final"), "######0.00")
           importe = rs("pu_final") * rs("cantidad")
           
    End If
    bonif = 0                   ' importe de descuentos
    iva_id = rs("id_afip")                  ' código para alícuota del 21%
                  ' importe liquidado de iva
   
    despacho = ""      ' numero de despacho de importación
    dato_a = ""           ' primer dato adicional del item
    dato_b = ""
    dato_c = ""
    dato_d = ""
    dato_e = ""           ' ultimo dato adicional del item
    ok = PyFEPDF.AgregarDetalleItem(u_mtx, cod_mtx, CODIGO, ds, qty, umed, _
            precio, bonif, iva_id, imp_iva, importe, despacho, _
            dato_a, dato_b, dato_c, dato_d, dato_e)
 
    
    rs.MoveNext
 Wend
 Set rs = Nothing
  
    
    
    ' Agrego datos adicionales fijos:
    ok = PyFEPDF.AgregarDato("logo", "c:\5a04\tools\logo.png")
    ok = PyFEPDF.AgregarDato("afip", "c:\5a04\tools\afip.png")
    ok = PyFEPDF.AgregarDato("EMPRESA", glo.nombrecli)
    ok = PyFEPDF.AgregarDato("MEMBRETE1", glo.direccioncli)
    ok = PyFEPDF.AgregarDato("MEMBRETE2", "Rojas")
    ok = PyFEPDF.AgregarDato("CUIT", "CUIT " & glo.CUIT)
    ok = PyFEPDF.AgregarDato("IIBB", "IIBB " & para.numero_ib)
    ok = PyFEPDF.AgregarDato("IVA", "IVA Responsable Inscripto")
    ok = PyFEPDF.AgregarDato("INICIO", "Inicio de Actividad: " & para.fecha_inicio_actividades)
    ok = PyFEPDF.AgregarDato("ObservacionesGenerales1", "Nota al pie1")
    ok = PyFEPDF.AgregarDato("ObservacionesGenerales2", "CBU para Transferencias ")
    ok = PyFEPDF.AgregarDato("ObservacionesGenerales3", "")


    ' Cargo el formato desde el archivo CSV (opcional)
    ' (carga todos los campos a utilizar desde la planilla)
    ok = PyFEPDF.CargarFormato("C:\5A04\tools\factura4.csv")
    
        ' Agrego campos manualmente (opcional):
        nombre = "remitos": tipo = "T"           ' "T" texto, "L" lineas, "I" imagen, etc.
        X1 = 8.7: Y1 = 247: X2 = 300: Y2 = 251   ' coordenadas (en milimetros)
        Font = "Arial": Size = 7: Bold = 0: Italic = 0: Underline = 0 ' tipo de letra
        foreground = "000000": background = "FFFFFF"    ' colores de frente y fondo
        Align = "I" ' Alineación: Centrado, Izquierda, Derecha
        prioridad = 2 ' Orden Z, menor prioridad se dibuja primero (para superposiciones)
        Text = "REMITOS:" & tr
        ok = PyFEPDF.AgregarCampo(nombre, tipo, X1, Y1, X2, Y2, _
                        Font, Size, Bold, Italic, Underline, _
                        foreground, background, _
                        Align, Text, Priority)
 
    
    
    
    If Me.moneda = "D" Then
        ' Agrego campos manualmente (opcional):
        nombre = "moneda": tipo = "T"           ' "T" texto, "L" lineas, "I" imagen, etc.
        X1 = 8.7: Y1 = 254: X2 = 300: Y2 = 258   ' coordenadas (en milimetros)
        Font = "Arial": Size = 7: Bold = 1: Italic = 0: Underline = 0 ' tipo de letra
        foreground = "000000": background = "FFFFFF"    ' colores de frente y fondo
        Align = "I" ' Alineación: Centrado, Izquierda, Derecha
        prioridad = 2 ' Orden Z, menor prioridad se dibuja primero (para superposiciones)
        Text = "Comprobante emitido en dolares estadounidenses. A fines impositivos se deberá utilizar un tipo de cambio de: " & moneda_ctz & " pesos por dolar."
        ok = PyFEPDF.AgregarCampo(nombre, tipo, X1, Y1, X2, Y2, _
                        Font, Size, Bold, Italic, Underline, _
                        foreground, background, _
                        Align, Text, Priority)
 
   End If
    
    ' Agrego campos manualmente (opcional):
    nombre = "cbu": tipo = "T"           ' "T" texto, "L" lineas, "I" imagen, etc.
    X1 = 100: Y1 = 261: X2 = 250: Y2 = 266   ' coordenadas (en milimetros)
    Font = "Arial": Size = 8: Bold = 1: Italic = 0: Underline = 0 ' tipo de letra
    foreground = "000000": background = "FFFFFF"    ' colores de frente y fondo
    Align = "I" ' Alineación: Centrado, Izquierda, Derecha
    prioridad = 2 ' Orden Z, menor prioridad se dibuja primero (para superposiciones)
    Text = "CBU para Transferencias: " & sacacbu()
    ok = PyFEPDF.AgregarCampo(nombre, tipo, X1, Y1, X2, Y2, _
                        Font, Size, Bold, Italic, Underline, _
                        foreground, background, _
                        Align, Text, Priority)

  
    ' Creo plantilla para esta factura (papel A4 vertical):
    papel = "A4" ' o "letter" para carta, "legal" para oficio
    orientacion = "portrait" ' o landscape (apaisado)
    ok = PyFEPDF.CrearPlantilla(papel, orientacion)
    
   
    num_copias = 3  ' original, duplicado y triplicado
    lineas_max = 24 ' cantidad de linas de items por página
    qty_pos = "der" ' (cantidad a la izquierda de la descripción del artículo)
    ' Proceso la plantilla
    ok = PyFEPDF.ProcesarPlantilla(num_copias, lineas_max, qty_pos)
    
    ' Genero el PDF de salida según la plantilla procesada
    nf = Me.letra & Format$(Me.sucursal, "0000") & "_" & Format$(Me.numcomp, "0000000000") & "_" & Me.abreviatura & ".pdf"
    
 
      
    salida = pathfacturas & "\" & nf
    ok = PyFEPDF.GenerarPDF(salida)
    MsgBox ("E7")
    ' Abro el visor de PDF y muestro lo generado
    ' (es necesario tener instalado Acrobat Reader o similar)
    gi = False ' cambiar a True para que lo envie directo a la impresora
    ok = PyFEPDF.MostrarPDF(salida, gi)
    
    
   
    
    
    


End If


    Exit Sub
ManejoError:
    ' Si hubo error:
    MsgBox ("Error al generar PDF. Verifique que el comrpobante no este abierto, la unidad de disco este disponible y vuelva a generar pdf del comprobante")
    
   ' Debug.Print Err.Description            ' descripción error afip
   ' Debug.Print Err.Number - vbObjectError ' codigo error afip
   ' Select Case MsgBox(Err.Description, vbCritical + vbRetryCancel, "Error:" & Err.Number - vbObjectError & " en " & Err.Source)
   '     Case vbRetry
   '         Debug.Print PyFEPDF.Excepcion
   '         Debug.Print PyFEPDF.Traceback
   '         Debug.Assert False
   '         Resume
   '     Case vbCancel
   '         Debug.Print Err.Description
   ' End Select
'Debug.Assert False

   
End Sub

Public Sub imprimir()
'previamente es necesario cargar el comprobante

'los fortmatos de los comp. venta son:
'1 Fabrimac
'2 Degser
'3 framahi
'4 el nuevo
'5 veterinaria aguer
'6 Agrosumma
'A andreoli
'W la calera (fotos)

Set rs = New adodb.Recordset
q = "select * from vta_06 where [sucursal] = " & Me.sucursal & " and  [id_tipocomp] = " & Me.idtipocomp
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
 Me.imprimedescextra = rs("imprime_desc_extra")
  
  
  Select Case Me.letra
  Case Is = "B"
   copias = rs("cant_copias_b")
  Case Is = "C"
     copias = rs("cant_copias_c")
  Case Is = "E"
     copias = rs("cant_copias_e")
  Case Else
     copias = rs("cant_copias_a")
  End Select
  formato = rs("formato")
Else
  copias = 0
  formato = "1"
  Me.imprimedescextra = "N"
End If
Set rs = Nothing

Select Case Me.idtipocomp
  Case Is = 50
    Select Case formato
      Case Is = "1"
         Load gen_logo
         gen_logo.Picture2.Picture = LoadPicture(App.Path & "\tools\firma1.jpg")
         Call imprimerbo1(copias)
         Unload gen_logo
      
      Case Is = "2"
         Call imprimerbo2(copias)
      Case Is = "3"
         Load gen_logo
         gen_logo.Picture2.Picture = LoadPicture(App.Path & "\tools\firma1.jpg")
         Call imprimerbo3(copias) 'framahi/el nuevo
         Unload gen_logo
      Case Is = "4"
         Call imprimerbo4(copias)
      Case Is = "6"
         Call imprimerbo6(copias) 'agrosumma
      Case Is = "7"
         Call imprimerbo7(copias) 'hugo diaz
      
      Case Is = "9"
         Load gen_logo
         gen_logo.Picture1.Picture = LoadPicture(App.Path & "\tools\logo2.jpg")
         Call imprimerbo3(copias) 'radio
         Unload gen_logo
         
    Case Is = "W"
         Call imprimerboW(copias) 'La calera 2
         
    End Select
  Case Is <= 10
     Call imprimefact(copias, formato)
  
  Case Is = 45, Is = 46
    Select Case formato
      Case Is = "1"
            'varios hoja en blanco sin valorizar
            Call imprimirrto1(copias)
      Case Is = "2"
         Call imprimirrto2(copias) 'degser
      Case Is = "3"
         Call imprimirrto3(copias) 'en blanco hoja a4 valorizado
       
      Case Is = "4"
         Call imprimirrto4(copias) 'Tulio
      Case Is = "5"
         Call imprimirrto5(copias) 'Grasticcini
      Case Is = "6"
         Call imprimirrto6(copias) 'Agrosumma
      Case Is = "7"
         Call imprimirrto7(copias) 'Hugo Diaz
      Case Is = "8"
         Call imprimirrto8(copias) 'Los vasquitos
      Case Is = "9"
         Call imprimirrto9(copias) 'Fabrimac
      Case Is = "Y"
         Call imprimirrtoY(copias) 'hoja en balnco 2 x hoja
      Case Is = "A"
         Call imprimirrto_A(copias) 'andreoli
      Case Is = "B"
         Call imprimirrto_B(copias) 'Diaz Hugo laser
      Case Is = "W"
         Call imprimirrtoW(copias) 'la calera
      Case Is = "E"
         Call imprimirrtoE(copias)   'electronico
      
      End Select
  
  Case Is = 40
    Select Case formato
      Case Is = "1"
            'varios hoja en blanco
            Call imprimirpresup1(copias)
      Case Is = "2"
            'lA CALERA
            Call imprimirpresup2(copias)
      Case Is = "A"
            'grasticcini
            Call imprimirpresup_A(copias)
        Case Is = "Z"
            'superkarr
            Call formatoZ(copias)
     
     End Select
  
  Case Is = 80, Is = 81
      Select Case formato
      Case Is = "1"
               'general
      
      Case Is = "2"
               ' el nuevo
               
      Case Is = "3"
               ' radio
         
    
    End Select
  
  Case Is = 150
       Call imprimirempaque(copias)
  
  
End Select

End Sub

Sub imprimefact(ByVal c As Integer, ByVal F As String)
Select Case F
Case Is = "1"
   'fabrimac
    If Me.letra = "A" Then
       Call formato1A(c)
    Else
       
    End If

Case Is = "2"
    'degser
       Call formato2A(c)
Case Is = "3"
    'Framahi
  Select Case Me.idtipocomp
   Case Is = 1 'facturas
    If Me.letra = "A" Then
       Call formato3A(c)
    Else
       
    End If
   Case Is = 3 'NC
    If Me.letra = "A" Then
       Call formato3ncA(c)
    Else
       
    End If
  End Select
Case Is = "4"
    'Tulio
    
       Call formato4A(c) 'fact a y b juntas
    
Case Is = "5"
    'veterinaria Aguer
    Call formato5(c) 'fact a y b juntas
    
Case Is = "6"
    'Agrosumma
    Call formato6(c) 'fact a y b juntas
    
Case Is = "7"
    'Hugo Diaz (texto)
    Call formato7(c) 'fact a y b juntas
  
Case Is = "8"
    'Los vasquitos
    Call formato8(c) 'fact a y b juntas
  

Case Is = "E"
    'Electronica
    
       'Call formatoE(c) 'fact a y b juntas
       Call generapdf(c)

Case Is = "X"
    'El negro
    Call formatoX(c) 'fact a y b juntas
Case Is = "Y"
    'El negro  2 por hojas
    Call formatoY(c) 'fact a y b juntas
Case Is = "W"
    'Idem Y pero sin cabecera (la calera 2)
    Call formatoW(c) 'fact a y b juntas

Case Is = "Z"
    'Idem a Y pero sin membrete(superkarr)
    Call formatoZ(c) 'fact a y b juntas
  
Case Is = "A"
    'Walter Andreoli
    Call formato_A(c) 'fact a y b juntas
  
  
  End Select
  
  
End Sub
Sub formato7(c As Integer)
  'FORMATO HUGO DIAZ TEXTO
Dim puerto As String
Set cl_cli = New Clientes
cl_cli.carga (Me.idcliente)
If cl_cli.id > 0 Then
 On Error GoTo IMP2
 puerto = "LPT1"
 Open puerto For Output As #1
   sonpesos = convierte(Me.total)
   dd = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 1, 2)
   mm = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 4, 2)
   aa = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 7, 4)
   For Y = 1 To c
    cc = Space$(5)
    cp = Space$(9)
    cn = Space$(8)
    pu = Space$(10)
    dtop = Space$(7)
    pp = Space$(10)
    im = Space$(10)
    s = Space$(10)
    s2 = Space$(10)
    dto = Space$(7)
    i1 = Space$(9)
    I2 = Space$(9)
    tot = Space$(10)
    RSet cc = Format$(Me.idcliente, "0000")
    n = Left$(Format$(Me.razonsocial, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!"), 30)
    d = Left$(Format$(Me.direccion, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!"), 30)
    l = Left$(Format$(Me.localidad, "@@@@@@@@@@@@@@@@@@@!"), 15)
    di = Left$(Format$(Me.abreviatura_tipoiva, "@@@@@@@@@@@@@@@@@@@@@@@@@@!"), 25)
    nc = Left$(Format$(Me.CUIT, "@@@@@@@@@@@@@@@@!"), 16)
    Print #1, Chr$(18)
    Print #1, Spc(58); Chr$(14); Me.abreviatura
    Print #1,
    Print #1,
    Print #1, Spc(59); dd & "/" & mm & "/" & aa; Spc(2); Format$(Me.letra, ">@"); Spc(1); Format$(Me.sucursal, "0000"); "-"; Format$(Me.numcomp, "00000000")
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1, Spc(9); n; Spc(17); "("; cc; ")"
    Print #1,
    Print #1, Spc(9); d; " "; Spc(17); nc
    Print #1, Spc(9); "(" & l & ")"
    Print #1, Spc(9); di
    Print #1,
    'Print #1,
    q = "select * from vta_08, vta_02 where [id_factura] = " & Me.numint & " and [id_remito] = [num_int]"
    Set rs = New adodb.Recordset
    rs.Open q, cn1
    tr = " "
    While Not rs.EOF
     nc = Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
     tr = tr & nc & " "
     rs.MoveNext
   Wend
   Set rs = Nothing
  
   If Me.moneda = "P" Then
     mda = " ($) "
   Else
     mda = "(U$s)"
   End If
 
    If Me.contado = "S" Then
         Print #1, Spc(16); "Cuenta Corriente "; mda; Spc(11); tr
    Else
         Print #1, Spc(16); "         Contado "; mda; Spc(11); tr
     End If
    Print #1,
    Print #1,
    Print #1,
    
    Set rs = New adodb.Recordset
    q = "Select * from vta_03 where [num_int] = " & Me.numint
    rs.Open q, cn1
    k = 0
    While Not rs.EOF
        i = Space$(10)
        t = Space$(4)
        pu = Space$(10)
        cant = Space$(8)
        d = Format$(Left$(rs("descripcion") & " ", 40), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
        dt = Space$(9)
        u = Space$(5)
        cp = Space$(7)
        If rs("cantidad") > 0 Then
            RSet pu = Format$(rs("pu"), "######0.00")
            RSet t = Format$(rs("tasaiva"), "#0.0")
            RSet i = Format$(rs("importe"), "######0.00")
            RSet cant = Format$(rs("cantidad"), "####0.00")
            dt = "  (" & t & "%)"
            u = Format$(rs("tunidad"), "@@@@@!")
            cp = "(" & Format$(rs("id_producto"), "00000") & ")"
        End If
        Print #1, cant; u; Spc(1); cp; Spc(2); d; Spc(2); pu; Spc(2); i
        rs.MoveNext
        k = k + 1
   Wend
   Set rs = Nothing
   
    For z = k To Me.cant_lineas
        Print #1,
    Next z
    
    PVTO = Format$(Me.fecha_vto, "dd/mm/yyyy")
    Print #1, "Vencimiento:"; Left$(Format$(PVTO, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!"), 30)
    Print #1, Chr$(15);
    If Me.moneda = "P" Then
      Print #1, "Son pesos:"; Left$(Format$(sonpesos, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!"), 100); Chr$(18);
      Print #1, Spc(15); Format$(Format$(totii, "######0.00"), "@@@@@@@@@@")
      d = asteriscos(Format$(Me.totalotramoneda, "0.00"), 10)
      Print #1, "El monto de la presente factura equivale a U$S "; Format$(d, "@@@@@@@@@@"); Spc(24);
      If Me.letra = "A" Then
        Print #1, Format$(Format$(Me.subtotal, "######0.00"), "@@@@@@@@@@")
      Else
        Print #1,
      End If
      Print #1, "debiendose cancelar al vencimiento estipulado, en Pesos   "; Spc(23);
      If Me.letra = "A" Then
        Print #1, Format$(Format$(Me.iva, "######0.00"), "@@@@@@@@@@")
      Else
        Print #1,
      End If
      Print #1, "segun Dolar Transferencia Tipo Vendedor del dia anterior "
      Print #1, "al efectivo Pago."
      Print #1, "Cambio al dia de Facturacion a Pesos "; Format$(Me.cotizaciondolar, "##0.000"); " por Dolar.";
      d = asteriscos(Format$(Me.total, "0.00"), 10)
      Print #1, Spc(27); "$"; d
    Else
      Print #1, " Son Dolares Estadounidenses:"; Left$(Format$(sonpesos, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!"), 80); Chr$(18);
      Print #1, Spc(15); Format$(Format$(totii, "######0.00"), "@@@@@@@@@@")
      d = asteriscos(Format$(Val(total) / Me.cotizaciondolar, "0.00"), 10)
      Print #1, " Importes expresados en Dolares Estadounidenses.          "; Spc(23);
      If Me.letra = "A" Then
        Print #1, Format$(Format$(Me.subtotal, "######0.00"), "@@@@@@@@@@")
      Else
        Print #1,
      End If
      Print #1, " A fines Impositivos el tipo de cambio es "; Format$(Me.cotizaciondolar, "##0.000"); " Iva $";
      ip = asteriscos(Format$(Me.iva * Me.cotizaciondolar, "0.00"), 8)
      Print #1, ip; Spc(20);
      If Me.letra = "A" Then
        Print #1, Format$(Format$(Me.iva, "######0.00"), "@@@@@@@@@@")
      Else
        Print #1,
      End If
      Print #1, " El tipo de cambio aplicable para la cancelacion de esta"
      Print #1, " factura sera el correspondiente al dolar billete libre "
      Print #1, " vendedor vigente al momento de la acreditacion del pago ";
      d = asteriscos(Format$(Me.total, "0.00"), 10)
      Print #1, Spc(21); "u$s"; d
    End If
 
    Print #1, 'Chr$(18)
    Print #1,
    Print #1,
    Print #1,
    Print #1,

    Next Y
    ok = 1
    Close #1

End If
Exit Sub


IMP2:
i = MsgBox("Error de Impresion. Verifique", 4)
If i = 6 Then
    Resume
Else
    ok = 0
    Close #1
    Exit Sub
End If




End Sub

Sub formato2A(c As Integer)
'Degser
Set cl_cli = New Clientes
cl_cli.carga (Me.idcliente)


If cl_cli.id > 0 Then
 Set rsr = New adodb.Recordset
 q = "select * from vta_08, vta_02 where [id_factura] = " & Me.numint & " and [id_remito] = [num_int]"
 rsr.Open q, cn1
 tr = ""
 While Not rsr.EOF
  If rsr("id_tipocomp") = 45 Then
    tc = "R"
  Else
    tc = "D"
  End If
  tr = tr & Format$(rsr("sucursal"), "0000") & "-" & Format$(rsr("num_comp"), "00000000") & tc & " * "
  rsr.MoveNext
 Wend
 Set rsr = Nothing
 
 dd = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 1, 2)
 mm = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 4, 2)
 aa = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 7, 4)
 lugar = Left$(RTrim$(cl_cli.direccion) & " - (" & RTrim$(cl_cli.cp) & ")" & RTrim$(cl_cli.localidad), 120)

 If Me.moneda = "P" Then
   sondolares = convierte(Me.totalotramoneda)
 Else
   sondolares = convierte(Me.total)
 End If

lw = InputBox("Indique cantidad de Copias", , c)
If Val(lw) > 0 And Val(lw) <= 5 Then

 For Y = 1 To Val(lw)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Times New Roman"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
    cc = Space$(5)
    cp2 = Space$(6)
    cant = Space$(8)
    pu = Space$(10)
    dtop = Space$(7)
    pp = Space$(10)
    im = Space$(10)
    s = Space$(9)
    s2 = Space$(9)
    dto = Space$(9)
    i1 = Space$(9)
    I2 = Space$(9)
    tot = Space$(9)
    totd = Space$(9)
    fd = Space$(6)
    piva = Space$(10)
    pib = Space$(10)
    RSet cc = Format$(Me.idcliente, "00000")
    RSet s = Format$(Me.subtotal, "0.00")
    RSet dto = Format$(Me.descuento, "0.00")
    RSet i1 = Format$(Me.iva, "0.00")
    RSet tot = Format$(Me.total, "0.00")
    RSet totd = Format$(Me.totalotramoneda, "0.00")
    RSet fd = Format$(Me.cotizaciondolar, "#0.000")
    RSet pib = Format$(Me.percib, "0.00")
    RSet piva = Format$(Me.perciva, "0.00")
    n = Format$(Left$(cl_cli.razonsocial, 40), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    d = Format$(Left$(cl_cli.direccion, 40), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    l = Format$(Left$(cl_cli.localidad, 30), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    di = Format$(Left$(cl_cli.abreviatura_tipoiva, 21), "@@@@@@@@@@@@@@@@@@@@@!")
    nc = Format$(Left$(cl_cli.CUIT, 13), "@@@@@@@@@@@@@!")
    cp = Left$(cl_cli.cp, 10)
    For i = 1 To 7
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print
    Next i
 
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Times New Roman"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(79); Me.abreviatura; Spc(2); tacha; Spc(2); Me.letra; Spc(1); Format$(Me.sucursal, "0000"); "-"; Format$(Me.numcomp, "00000000")      '; Spc(4); dD & "/" & mM & "/" & aA
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 10
    If Me.idtipocomp <> 1 Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
        Printer.Print Tab(89); "#############"
    Else
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
        Printer.Print
    End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 12
    For i = 1 To 4
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print
    Next i

'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 9
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(90); "[" & Format$(Me.idvendedor, "000") & "]"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(13); dd & "/" & mm & "/" & aa;
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Spc(6); "[" & cc & "]"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(92); nc
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Times New Roman"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 14
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(11); n;
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 8
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(102); di
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(10); d; Tab(70);
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 6
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 9
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = True
    If Me.moneda = "P" Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print Tab(90); "Pesos"
    Else
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
     Printer.Print Tab(90); "U$S"
    End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(10); l; Tab(43); "(" & cp & ")"; Tab(68);
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 9
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
    If Me.ctacte <> "N" Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print "Cta.Cte."
    Else
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print "Contado "
    End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 9
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
    Set rs = New adodb.Recordset
    q = "Select * from vta_03 where [num_int] = " & Me.numint
    rs.Open q, cn1
    cp2 = Space$(6)
    cant = Space$(8)
    pu = Space$(10)
    dtop = Space$(6)
    u = Space$(5)
    t = Space$(4)
    l = 0
    While Not rs.EOF
       d = Format$(Left$(rs("descripcion") & " ", 40), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
       If rs("cantidad") > 0 Then
         RSet cp2 = Format$(rs("id_producto"), "000000")
         RSet cant = Format$(rs("cantidad"), "####0.00")
         If Me.letra = "A" Then
          RSet pu = Format$(rs("pu"), "#####0.00")
          RSet im = Format$(rs("importe"), "######0.00")
         Else
          RSet pu = Format$(rs("pu_final"), "#####0.00")
          RSet im = Format$(Val(pu) * Val(cant), "######0.00")
         End If
         
         RSet t = Format$(rs("tasaiva"), "#0.0")
         dt = "(" & t & "%)"
         u = Format$(Left$(rs("tunidad"), 5), "@@@@@!")
       End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.FontSize = 9
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print Tab(10); "[" & cp2 & "]"; cant; Spc(5); u; Spc(1);
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print d;
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print pu;
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.FontSize = 6
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print dt;
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.FontSize = 9
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print Spc(1); im
       rs.MoveNext
       l = l + 1
   Wend
   Set rs = Nothing
     
   For i = l To 25
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
     Printer.Print
   Next i
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 6
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "           Remitos: ";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print tr
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Line Step(30, 1)-Step(9000, 1)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 8
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "      Son dolares: ";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print sondolares
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "      Tipo Cambio: ";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  U$s 1 = $ " & Format$(Me.cotizaciondolar, "0.000")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "      Total de Iva en $: ";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
   If Me.moneda = "P" Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Format$(Me.iva, "0.00");
   Else
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Format$(Me.iva * Me.cotizaciondolar, "0.00");
   End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(60); "Total Perc. I.B en $: ";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
   If Me.moneda = "P" Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Format$(Me.percib, "0.00")
   Else
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Format$(Me.percib * Me.cotizaciondolar, "0.00")
   End If
   PVTO = Me.fecha_vto
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "      Vencimiento: "; Left$(Format$(PVTO, "@@@@@@@@@@@@!"), 12);
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(60); "Total Perc. Iva en $: ";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
   If Me.moneda = "P" Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Format$(Me.perciva, "0.00")
   Else
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Format$(Me.perciva * Me.cotizaciondolar, "0.00")
   End If

'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Line Step(30, 1)-Step(9000, 1)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontSize = 9
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print Tab(8); "Condiciones de Pago:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print " Para fines impositivos los importes deben convertirse  al tipo  de"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print Tab(8); "cambio vendedor del B.N.A. correspondiente al cierre del día anterior a la acreditacion"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print Tab(8); "efectiva de los valores recibidos en pago."
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print Tab(8); "Cualquier  pago  efectuado  antes  de  conocerse  el tipo  de  cambio  aplicable,  será"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print Tab(8); "considerado pago a cuenta del importe que en definitiva corresponda."

 
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontBold = False
 
 For h = 1 To 2
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
 Next h
 
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontSize = 8
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print Tab(55); "Percepciones"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 If Me.letra = "A" Then
   Printer.Print Tab(11); s; Tab(28); i1;
Else
    Printer.Print Tab(11); " "; Tab(28); " ";
End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontSize = 8
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print Tab(55); "I.B.."; pib;
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print Tab(60); porcdt; Tab(72);
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontBold = True
 If Me.moneda = "P" Then
   TX = "  $  "
 Else
   TX = " U$S "
 End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print TX & " " & tot
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontSize = 8
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print Tab(55); "Iva.."; piva

'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontSize = 6
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print Tab(20); "ORIGINAL: Blanco  -  DUPLICADO: Celeste  -  TRIPLICADO: Amarillo"
  If Y < c Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.NewPage
  End If
Next Y
 ok = 1
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.EndDoc
 End If
End If
Exit Sub

imp1:
i = MsgBox("La impresora no esta lista.Verifique", 4)
If i = 6 Then
    Resume
Else
    ok = 0
    Exit Sub
End If


IMP2:
MsgBox ("Error durante la impresion")
ok = 0
Close #1
Exit Sub

End Sub

Sub formato3A(c As Integer)
'Framahi
Set cl_cli = New Clientes
cl_cli.carga (Me.idcliente)


If cl_cli.id > 0 Then
 Set rsr = New adodb.Recordset
 q = "select * from vta_08, vta_02 where [id_factura] = " & Me.numint & " and [id_remito] = [num_int]"
 rsr.Open q, cn1
 tr = ""
 While Not rsr.EOF
  If rsr("id_tipocomp") = 45 Then
    tc = "R"
  Else
    tc = "D"
  End If
  tr = tr & Format$(rsr("sucursal"), "0000") & "-" & Format$(rsr("num_comp"), "00000000") & tc & " * "
  rsr.MoveNext
 Wend
 Set rsr = Nothing
 
 dd = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 1, 2)
 mm = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 4, 2)
 aa = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 7, 4)
 lugar = Left$(RTrim$(cl_cli.direccion) & " - (" & RTrim$(cl_cli.cp) & ")" & RTrim$(cl_cli.localidad), 120)

 If Me.moneda = "P" Then
   sondolares = convierte(Me.totalotramoneda)
 Else
   sondolares = convierte(Me.total)
 End If

lw = InputBox("Indique cantidad de Copias", , c)
If Val(lw) > 0 And Val(lw) <= 5 Then

 For Y = 1 To Val(lw)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Times New Roman"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
    cc = Space$(5)
    cp2 = Space$(6)
    cant = Space$(8)
    pu = Space$(10)
    dtop = Space$(7)
    pp = Space$(10)
    im = Space$(10)
    s = Space$(9)
    s2 = Space$(9)
    dto = Space$(9)
    i1 = Space$(9)
    I2 = Space$(9)
    tot = Space$(9)
    totd = Space$(9)
    fd = Space$(6)
    piva = Space$(10)
    pib = Space$(10)
    RSet cc = Format$(Me.idcliente, "00000")
    RSet s = Format$(Me.subtotal, "0.00")
    RSet dto = Format$(Me.descuento, "0.00")
    RSet i1 = Format$(Me.iva, "0.00")
    RSet tot = Format$(Me.total, "0.00")
    RSet totd = Format$(Me.totalotramoneda, "0.00")
    RSet fd = Format$(Me.cotizaciondolar, "#0.000")
    RSet pib = Format$(Me.percib, "0.00")
    RSet piva = Format$(Me.perciva, "0.00")
    n = Format$(Left$(cl_cli.razonsocial, 40), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    d = Format$(Left$(cl_cli.direccion, 40), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    l = Format$(Left$(cl_cli.localidad, 30), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    di = Format$(Left$(cl_cli.abreviatura_tipoiva, 21), "@@@@@@@@@@@@@@@@@@@@@!")
    nc = Format$(Left$(cl_cli.CUIT, 13), "@@@@@@@@@@@@@!")
    cp = Left$(cl_cli.cp, 10)
    For i = 1 To 5
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print
    Next i
 
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Times New Roman"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(85); "Comprobante..: "; Me.abreviatura; Spc(2); tacha; Spc(2); Me.letra; Spc(1); Format$(Me.sucursal, "0000"); "-"; Format$(Me.numcomp, "00000000")       '; Spc(4); dD & "/" & mM & "/" & aA
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(85); "Fecha..............: "; Me.fecha
    For i = 1 To 5
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print
    Next i
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(6); "RAZON SOCIAL:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print " "; n; Tab(92); "[" & cc & "]"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(6); "DIRECCION:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print " "; d; Tab(90);
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print "IVA:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print " "; di
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
    
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(6); "LOCALIDAD:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print " "; l; "(" & cp & ")"; Tab(90);
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print "CUIT:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print " "; nc
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
    If Me.ctacte <> "N" Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print Tab(10); "Condicion de Venta: Cuenta Corriente"
    Else
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print Tab(10); "Condicion de Venta: Contado"
    End If
    
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 9
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 9
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(4); "Art."; Tab(12); "Cantidad"; Tab(24); "Detalle"; Tab(79); "P.U."; Tab(84); "%Iva"; Tab(95); "TOTAL"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Line Step(100, 1)-Step(11900, 1)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
    Set rs = New adodb.Recordset
    q = "Select * from vta_03 where [num_int] = " & Me.numint
    rs.Open q, cn1
    l = 0
    While Not rs.EOF
      cp2 = Space$(6)
      cant = Space$(8)
      pu = Space$(10)
      dtop = Space$(6)
      u = Space$(5)
      t = Space$(4)
      im = Space$(10)
       d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
       If rs("cantidad") > 0 Then
         RSet cp2 = Format$(rs("id_producto"), "000000")
         RSet cant = Format$(rs("cantidad"), "####0.00")
         RSet pu = Format$(rs("pu"), "#####0.00")
         RSet im = Format$(rs("importe"), "######0.00")
         RSet t = Format$(rs("tasaiva"), "#0.0")
         dt = "(" & t & "%)"
         u = Format$(Left$(rs("tunidad"), 5), "@@@@@!")
         CP3 = "[" & cp2 & "]"
       Else
        CP3 = " " & cp2 & " "
        dt = Space$(7)
       End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.FontSize = 9
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print Tab(4); CP3; cant; Spc(5);
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print d;
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print pu;
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.FontSize = 6
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print dt;
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.FontSize = 9
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print Spc(1); im
       rs.MoveNext
       l = l + 1
   Wend
   Set rs = Nothing
     
   For i = l To 31
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
     Printer.Print
   Next i
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 6
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "           Remitos: ";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print tr
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Line Step(100, 1)-Step(9000, 1)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 8
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print " ";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print " "
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print " ";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print " "
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "      Total de Iva en $: ";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
   If Me.moneda = "P" Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Format$(Me.iva, "0.00")
   Else
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Format$(Me.iva * Me.cotizaciondolar, "0.00")
   End If
   PVTO = Me.fecha_vto
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "      Vencimiento: "; Left$(Format$(PVTO, "@@@@@@@@@@@@!"), 12);
   
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Line Step(100, 1)-Step(9000, 1)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print
  
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontBold = False
 
 For h = 1 To 5
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
 Next h
 
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontSize = 9
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontBold = False
 
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print Tab(6); "SUBTOTAL"; Tab(23); "DESCUENTO"; Tab(41); "SUBTOTAL"; Tab(61); "IVA"; Tab(94); "TOTAL"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print Tab(6); s; Tab(53); i1;
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print Tab(81);
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontBold = True
 If Me.moneda = "P" Then
   TX = " $"
 Else
   TX = "U$s"
 End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print TX & " " & tot
  If Y < c Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.NewPage
  End If
Next Y
 ok = 1
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.EndDoc
 End If
End If
Exit Sub

imp1:
i = MsgBox("La impresora no esta lista.Verifique", 4)
If i = 6 Then
    Resume
Else
    ok = 0
    Exit Sub
End If


IMP2:
MsgBox ("Error durante la impresion")
ok = 0
Exit Sub

End Sub


Sub formato3ncA(c As Integer)
'Framahi
Set cl_cli = New Clientes
cl_cli.carga (Me.idcliente)


If cl_cli.id > 0 Then
 Set rsr = New adodb.Recordset
 q = "select * from vta_08, vta_02 where [id_factura] = " & Me.numint & " and [id_remito] = [num_int]"
 rsr.Open q, cn1
 tr = ""
 While Not rsr.EOF
  If rsr("id_tipocomp") = 45 Then
    tc = "R"
  Else
    tc = "D"
  End If
  tr = tr & Format$(rsr("sucursal"), "0000") & "-" & Format$(rsr("num_comp"), "00000000") & tc & " * "
  rsr.MoveNext
 Wend
 Set rsr = Nothing
 
 dd = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 1, 2)
 mm = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 4, 2)
 aa = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 7, 4)
 lugar = Left$(RTrim$(cl_cli.direccion) & " - (" & RTrim$(cl_cli.cp) & ")" & RTrim$(cl_cli.localidad), 120)

 If Me.moneda = "P" Then
   sondolares = convierte(Me.totalotramoneda)
 Else
   sondolares = convierte(Me.total)
 End If

lw = InputBox("Indique cantidad de Copias", , c)
If Val(lw) > 0 And Val(lw) <= 5 Then

 For Y = 1 To Val(lw)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Times New Roman"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
    cc = Space$(5)
    cp2 = Space$(6)
    cant = Space$(8)
    pu = Space$(10)
    dtop = Space$(7)
    pp = Space$(10)
    im = Space$(10)
    s = Space$(9)
    s2 = Space$(9)
    dto = Space$(9)
    i1 = Space$(9)
    I2 = Space$(9)
    tot = Space$(9)
    totd = Space$(9)
    fd = Space$(6)
    piva = Space$(10)
    pib = Space$(10)
    RSet cc = Format$(Me.idcliente, "00000")
    RSet s = Format$(Me.subtotal, "0.00")
    RSet dto = Format$(Me.descuento, "0.00")
    RSet i1 = Format$(Me.iva, "0.00")
    RSet tot = Format$(Me.total, "0.00")
    RSet totd = Format$(Me.totalotramoneda, "0.00")
    RSet fd = Format$(Me.cotizaciondolar, "#0.000")
    RSet pib = Format$(Me.percib, "0.00")
    RSet piva = Format$(Me.perciva, "0.00")
    n = Left$(Format$(cl_cli.razonsocial, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!"), 40)
    d = Left$(Format$(cl_cli.direccion, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!"), 40)
    l = Left$(Format$(cl_cli.localidad, "@@@@@@@@@@@@@@@@@@@@@@@@!"), 20)
    di = Left$(Format$(cl_cli.abreviatura_tipoiva, "@@@@@@@@@@@@@@@@@@@@@!"), 21)
    nc = Left$(Format$(cl_cli.CUIT, "@@@@@@@@@@@@@!"), 13)
    cp = Left$(cl_cli.cp, 10)
    For i = 1 To 4
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print
    Next i
 
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Times New Roman"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(120); ""; Me.fecha
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(85); "Comprobante..: "; Me.abreviatura; Spc(2); tacha; Spc(2); Me.letra; Spc(1); Format$(Me.sucursal, "0000"); "-"; Format$(Me.numcomp, "00000000")       '; Spc(4); dD & "/" & mM & "/" & aA
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 12
    For i = 1 To 2
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print
    Next i
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(14); n; Tab(110); "[" & cc & "]"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(14); d; Tab(82); l; Tab(118); "("; cp; ")"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(14); di; Tab(78); nc
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
   If Me.ctacte <> "N" Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print Tab(30); "Cuenta Corriente"
    Else
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print Tab(30); "Contado"
    End If
    
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 9
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 9
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
    Set rs = New adodb.Recordset
    q = "Select * from vta_03 where [num_int] = " & Me.numint
    rs.Open q, cn1
    l = 0
    While Not rs.EOF
      cp2 = Space$(6)
      cant = Space$(8)
      pu = Space$(10)
      dtop = Space$(6)
      u = Space$(5)
      t = Space$(4)
      im = Space$(10)
       d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
       If rs("cantidad") > 0 Then
         RSet cp2 = Format$(rs("id_producto"), "000000")
         RSet cant = Format$(rs("cantidad"), "####0.00")
         RSet pu = Format$(rs("pu"), "#####0.00")
         RSet im = Format$(rs("importe"), "######0.00")
         RSet t = Format$(rs("tasaiva"), "#0.0")
         dt = "(" & t & "%)"
         u = Format$(Left$(rs("tunidad"), 5), "@@@@@!")
         CP3 = "[" & cp2 & "]"
       Else
        CP3 = " " & cp2 & " "
        dt = Space$(7)
       End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.FontSize = 9
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print Tab(2); CP3; cant; Spc(5);
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print d;
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print pu;
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.FontSize = 6
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print dt;
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.FontSize = 9
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print Spc(1); im
       rs.MoveNext
       l = l + 1
   Wend
   Set rs = Nothing
     
   For i = l To 30
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
     Printer.Print
   Next i
     
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontBold = False
 
 For h = 1 To 4
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
 Next h
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print Tab(4); s; Tab(52); i1;
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print Tab(80);
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.FontBold = True
 If Me.moneda = "P" Then
   TX = " $"
 Else
   TX = "U$s"
 End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print TX & " " & tot
  If Y < c Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.NewPage
  End If
Next Y
 ok = 1
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.EndDoc
 End If
End If
Exit Sub

imp1:
i = MsgBox("La impresora no esta lista.Verifique", 4)
If i = 6 Then
    Resume
Else
    ok = 0
    Exit Sub
End If


IMP2:
MsgBox ("Error durante la impresion")
ok = 0
Exit Sub

End Sub
Sub imprimirrto1(ByVal c As Integer)
'formato en hoja A4 en blanco grafica sin valorizar
Dim a(5) As String
 
 
 If Me.idtipocomp = 45 Then
   comp = "REMITO  "
 Else
   comp = "DEVOLUCION  "
 End If
 
l = InputBox("Indique cantidad de Copias", , c)
If Val(l) > 0 And Val(l) <= 6 Then
 
For Y = 1 To Val(l)
 Printer.FontName = "Times New Roman"
 cc = Space$(5)
 cp = Space$(9)
 cant = Space$(8)
 pu = Space$(9)
 dtop = Space$(7)
 pp = Space$(10)
 im = Space$(10)
 s = Space$(10)
 s2 = Space$(10)
 dto = Space$(7)
 i1 = Space$(9)
 I2 = Space$(9)
 tot = Space$(10)
 RSet cc = Format$(Me.idcliente, "0000")
 n = Format$(Me.razonsocial, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 d = Format$(Me.direccion, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 l = Format$(Me.localidad, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 di = Format$(Me.abreviatura_tipoiva, "@@@@@@@@@@!")
 nc = Format$(Me.CUIT, "@@@@@@@@@@@@@@@@!")
 Printer.FontSize = 14
 Printer.Print glo.nombrecli
 Printer.FontSize = 10
 Printer.Print glo.direccioncli & " - Rojas(B) - CP.2705 - TE:" & glo.TECLI
 Printer.Print "Cuit:" & glo.CUIT
 Printer.Print "______________________________________________________________________________________"
 Printer.Print
 Printer.Print
 Printer.FontSize = 16
 Printer.Print Spc(30); comp;
 Printer.FontSize = 12
 Printer.Print "(Documento no valido como Factura)"
 Printer.Print
 Printer.Print
 Printer.Print Spc(80); " Fecha.....: " & Format$(Me.fecha)
 Printer.Print Spc(80); " Nro.......: "; Me.letra; Spc(1); Format$(Me.sucursal, "0000"); "-"; Format$(Me.numcomp, "00000000")
 Printer.Print
 Printer.Print
 Printer.FontSize = 12
 Printer.FontBold = False
 Printer.FontUnderline = True
 Printer.Print "Cliente.....:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "   ["; cc; "]  "; n
 Printer.Print
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Direccion...:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "   " & direccion & "  - " & localidad
 Printer.Print
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Iva.........:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "  " & di & "         ";
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Cuit.........: ";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print nc
 Printer.Print
 Printer.FontName = "Courier New"
 Printer.FontSize = "10"
 Printer.Print
 Printer.Print
 Printer.Print "      ==============================================================================="
 Printer.Print "      Cantidad   Unidad   Codigo    Articulo                                 "
 Printer.Print "      ==============================================================================="
 Set rs = New adodb.Recordset
 q = "Select * from vta_03 where [num_int] = " & Me.numint
 rs.Open q, cn1
 l = 0
 While Not rs.EOF
    d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    dt = Space$(9)
    cant = Space$(10)
    cp2 = Space$(5)
    tu = Space$(8)
    If rs("cantidad_original") > 0 Then
      RSet cant = Format$(rs("cantidad_original"), "######0.00")
      RSet cp2 = Format$(rs("id_producto"), "00000")
      tu = Format$(rs("tunidad"), "@@@@@@@@!")
    End If
    Printer.Print Tab(4); cant; "    "; tu; "  "; cp2; Spc(4);
    Printer.FontBold = True
    Printer.Print d
    Printer.FontBold = False
    
    'imprime extra
     If Me.imprimedescextra = "S" Then
       Set rs1 = New adodb.Recordset
       q = "select * from vta_015 where [num_int] = " & Me.numint & " and [renglon] = " & rs("renglon")
       rs1.Open q, cn1
       If Not rs1.EOF And Not rs1.BOF Then
         Call lee_desc_extra(a, rs1("desc_ext"))
         For k = 0 To 4
           If a(k) <> "%%" Then
             d = Format$(Left$(a(k), 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
             Printer.Print Tab(37); d
             l = l + 1
           End If
         Next k
       
       End If
       Set rs1 = Nothing
     End If
    
    
    
    rs.MoveNext
    l = l + 1
 Wend
 Set rs = Nothing
 
 
For ó = l To Me.cant_lineas
  Printer.Print
Next ó

For ó = 1 To 2
  Printer.Print
Next ó
Printer.Print Spc(55); "________________________________"
Printer.Print Spc(55); "          Recibi Conforme"
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.FontSize = 8
Printer.Print "________________________________"
 
  Set cl_usuarios = New usuarios
   cl_usuarios.cargar (Me.idusuario)
   If cl_usuarios.idusuario > 0 Then
        Printer.Print "Emitida por: " & cl_usuarios.denominACION
   Else
        Printer.Print
   End If
   Set cl_usuarios = Nothing
   If Y < c Then
      Printer.NewPage
   End If
  Next Y
 
Printer.EndDoc
End If
End Sub

Sub imprimirrto3(ByVal c As Integer)
'formato en hoja A4 en blanco grafica  valorizado
 If Me.idtipocomp = 45 Then
   comp = "REMITO  "
 Else
   comp = "DEVOLUCION  "
 End If
 
l = InputBox("Indique cantidad de Copias", , c)
If Val(l) > 0 And Val(l) <= 6 Then
 
For Y = 1 To Val(l)
 Printer.FontName = "Times New Roman"
 cc = Space$(5)
 cp = Space$(9)
 cant = Space$(8)
 pu = Space$(9)
 dtop = Space$(7)
 pp = Space$(10)
 im = Space$(10)
 s = Space$(10)
 s2 = Space$(10)
 dto = Space$(7)
 i1 = Space$(9)
 I2 = Space$(9)
 tot = Space$(10)
 RSet cc = Format$(Me.idcliente, "0000")
 n = Format$(Me.razonsocial, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 d = Format$(Me.direccion, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 l = Format$(Me.localidad, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 di = Format$(Me.abreviatura_tipoiva, "@@@@@@@@@@!")
 nc = Format$(Me.CUIT, "@@@@@@@@@@@@@@@@!")
 Printer.FontSize = 14
 Printer.Print Tab(4); glo.nombrecli
 Printer.FontSize = 10
 Printer.Print Tab(4); glo.direccioncli & " - Rojas(B) - CP.2705 - TE:" & glo.TECLI
 Printer.Print Tab(4); "Cuit:" & glo.CUIT
 Printer.Print Tab(4); "______________________________________________________________________________________"
 Printer.Print
 Printer.Print
 Printer.FontSize = 16
 Printer.Print Spc(30); comp;
 Printer.FontSize = 12
 Printer.Print "(Documento no valido como Factura)"
 Printer.Print
 Printer.Print
 Printer.Print Spc(80); " Fecha.....: " & Format$(Me.fecha)
 Printer.Print Spc(80); " Nro.......: "; Me.letra; Spc(1); Format$(Me.sucursal, "0000"); "-"; Format$(Me.numcomp, "00000000")
 Printer.Print
 Printer.Print
 Printer.FontSize = 12
 Printer.FontBold = False
 Printer.FontUnderline = True
 Printer.Print Tab(4); "Cliente.....:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "   ["; cc; "]  "; n
 Printer.Print
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print Tab(4); "Direccion...:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "   " & direccion & "  - " & localidad
 Printer.Print
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print Tab(4); "Iva.........:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "  " & di;
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print Tab(80); "Cuit.........: ";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print nc
 Printer.Print
 Printer.FontName = "Courier New"
 Printer.FontSize = "9"
 Printer.Print
 Printer.Print
 Printer.Print "      ==============================================================================================="
 Printer.Print "       Cant. Unidad   Codigo  Articulo                                             PU       IMPORTE  "
 Printer.Print "      =============================================================================================="
 Set rs = New adodb.Recordset
 q = "Select * from vta_03 where [num_int] = " & Me.numint
 rs.Open q, cn1
 l = 0
 While Not rs.EOF
    d = Format$(Left$(rs("descripcion") & " ", 45), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    dt = Space$(9)
    cant = Space$(10)
    cp2 = Space$(5)
    tu = Space$(8)
    im = Space$(10)
    pu = Space$(10)
    If rs("cantidad_original") > 0 Then
      RSet cant = Format$(rs("cantidad_original"), "######0.00")
      RSet cp2 = Format$(rs("id_producto"), "00000")
      tu = Format$(rs("tunidad"), "@@@@@@@@!")
      RSet pu = Format$(rs("PU_FINAL"), "#####0.00")
      RSet im = Format$(Val(pu) * Val(cant), "######0.00")
    End If
    Printer.Print Tab(2); cant; "  "; tu; "  "; cp2; Spc(2);
    Printer.FontBold = True
    Printer.Print d;
    Printer.FontBold = False
    Printer.Print Spc(2); pu; Spc(2); im
    rs.MoveNext
    l = l + 1
 Wend
 Set rs = Nothing
 
 
For ó = l To Me.cant_lineas
  Printer.Print
Next ó

For ó = 1 To 2
  Printer.Print
Next ó
Printer.FontSize = 11
Printer.Print Tab(8); "***************************************"
Printer.Print Tab(8); "Total a Facturar  $"; asteriscos(Format$(Me.total, "#####0.00"), 12)
Printer.Print Tab(8); "***************************************"
Printer.FontSize = 10
Printer.Print
Printer.Print
Printer.Print
Printer.Print Tab(9); "Nota: " & Me.observaciones
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print Spc(55); "________________________________"
Printer.Print Spc(55); "          Recibi Conforme"
Printer.Print
Printer.Print
Printer.Print
Printer.FontSize = 8
 
  Set cl_usuarios = New usuarios
   cl_usuarios.cargar (Me.idusuario)
   If cl_usuarios.idusuario > 0 Then
        Printer.Print Tab(2); "Emitido por: " & cl_usuarios.denominACION
   Else
        Printer.Print
   End If
   Set cl_usuarios = Nothing
   If Y < c Then
      Printer.NewPage
   End If
  Next Y
 
Printer.EndDoc
End If
End Sub
Sub imprimirpresup1(ByVal c As Integer)
'formato en hoja A4 en blanco grafica
 comp = "PRESUPUESTO  "
 
l = InputBox("Indique cantidad de Copias", , c)
If Val(l) > 0 And Val(l) <= 6 Then
 
For Y = 1 To Val(l)
 Printer.FontName = "Times New Roman"
 cc = Space$(5)
 cp = Space$(9)
 cant = Space$(8)
 pu = Space$(9)
 dtop = Space$(7)
 pp = Space$(10)
 im = Space$(10)
 s = Space$(10)
 s2 = Space$(10)
 dto = Space$(7)
 i1 = Space$(9)
 I2 = Space$(9)
 tot = Space$(10)
 RSet cc = Format$(Me.idcliente, "0000")
 n = Format$(Me.razonsocial, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 d = Format$(Me.direccion, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 l = Format$(Me.localidad, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 di = Format$(Me.abreviatura_tipoiva, "@@@@@@@@@@!")
 nc = Format$(Me.CUIT, "@@@@@@@@@@@@@@@@!")
 Printer.FontSize = 14
 Printer.Print glo.nombrecli
 Printer.FontSize = 10
 Printer.Print glo.direccioncli & " - Rojas(B) - CP.2705 - TE:" & glo.TECLI
 Printer.Print "Cuit:" & glo.CUIT
 Printer.Print "______________________________________________________________________________________"
 Printer.Print
 Printer.Print
 Printer.FontSize = 16
 Printer.Print Spc(30); comp;
 Printer.FontSize = 12
 Printer.Print "(Documento no valido como Factura)"
 Printer.Print
 Printer.Print
 Printer.Print Spc(80); " Fecha.....: " & Format$(Me.fecha)
 Printer.Print Spc(80); " Nro.......: "; Me.letra; Spc(1); Format$(Me.sucursal, "0000"); "-"; Format$(Me.numcomp, "00000000")
 Printer.Print
 Printer.Print
 Printer.FontSize = 12
 Printer.FontBold = False
 Printer.FontUnderline = True
 Printer.Print "Cliente.....:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "   ["; cc; "]  "; n
 Printer.Print
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Direccion...:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "   " & direccion & "  - " & localidad
 Printer.Print
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Iva.........:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "  " & di & "         ";
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Cuit.........: ";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print nc; "                   ";
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Condicion....: ";
 Printer.FontBold = False
 Printer.FontUnderline = False
 If Me.contado = "S" Then
   condicion = " CONTADO"
 Else
   condicion = " CUENTA CORRIENTE"
 End If
 Printer.Print condicion
 Printer.Print
 Printer.FontName = "Courier New"
 Printer.FontSize = 9
 Printer.Print
 Printer.Print
 Printer.Print "      ================================================================================================="
 Printer.Print "      Cant.  Unidad  Codigo  Articulo                                              P.U    IMPORTE"
 Printer.Print "      ================================================================================================="
 Set rs = New adodb.Recordset
 q = "Select * from vta_03 where [num_int] = " & Me.numint
 rs.Open q, cn1
 l = 0
 While Not rs.EOF
    d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    pu = Space$(8)
    i = Space$(9)
    cant = Space$(8)
    cp2 = Space$(5)
    tu = Space$(6)
    If rs("cantidad_original") > 0 Then
      RSet cant = Format$(rs("cantidad_original"), "####0.00")
      RSet cp2 = Format$(rs("id_producto"), "00000")
      tu = Format$(Left$(rs("tunidad"), 6), "@@@@@@!")
      RSet pu = Format$(rs("PU_FINAL"), "####0.00")
      RSet i = Format$(rs("PU_FINAL") * rs("cantidad_original"), "######0.00")
    
    End If
    Printer.Print Tab(4); cant; "  "; tu; "  "; cp2; Spc(2);
    Printer.FontBold = True
    Printer.Print d;
    Printer.FontBold = False
    Printer.Print " "; pu; " "; i
    
    rs.MoveNext
    l = l + 1
 Wend
 Set rs = Nothing
 
 
For ó = l To Me.cant_lineas
  Printer.Print
Next ó

For ó = 1 To 2
  Printer.Print
Next ó
Printer.Print Spc(35); "****************************"
Printer.Print Spc(25); "TOTAL      ";
Printer.FontSize = 12
Printer.FontBold = True
Printer.Print asteriscos(Format$(Me.total, "######0.00"), 15)
Printer.FontSize = 9
Printer.FontBold = False
Printer.Print Spc(35); "****************************"

Printer.Print
Printer.Print
Printer.Print
Printer.FontSize = 8
Printer.Print "________________________________"
 
  Set cl_usuarios = New usuarios
   cl_usuarios.cargar (Me.idusuario)
   If cl_usuarios.idusuario > 0 Then
        Printer.Print "Emitida por: " & cl_usuarios.denominACION
   Else
        Printer.Print
   End If
   Set cl_usuarios = Nothing
   If Y < Val(l) Then
      Printer.NewPage
   End If
Next Y
 
Printer.EndDoc
End If

End Sub

Sub imprimirpresup2(ByVal c As Integer)
'formato en hoja A4 lA CALERA
 comp = "PRESUPUESTO  "
 
l = InputBox("Indique cantidad de Copias", , c)
If Val(l) > 0 And Val(l) <= 6 Then
 
For Y = 1 To Val(l)
 Printer.FontName = "Times New Roman"
 cc = Space$(5)
 cp = Space$(9)
 cant = Space$(8)
 pu = Space$(9)
 dtop = Space$(7)
 pp = Space$(10)
 im = Space$(10)
 s = Space$(10)
 s2 = Space$(10)
 dto = Space$(7)
 i1 = Space$(9)
 I2 = Space$(9)
 tot = Space$(10)
 RSet cc = Format$(Me.idcliente, "0000")
 n = Format$(Me.razonsocial, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 d = Format$(Me.direccion, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 l = Format$(Me.localidad, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 di = Format$(Me.abreviatura_tipoiva, "@@@@@@@@@@!")
 nc = Format$(Me.CUIT, "@@@@@@@@@@@@@@@@!")
 Printer.FontSize = 14
 Printer.Print
 Printer.Print
 Printer.Print
 Printer.Print
 Printer.FontSize = 16
 Printer.Print Spc(35); comp;
 Printer.FontSize = 12
 Printer.Print "(Documento no valido como Factura)"
 Printer.Print
 Printer.Print
 Printer.Print Spc(80); " Fecha.....: " & Format$(Me.fecha)
 Printer.Print Spc(80); " Nro.......: "; Me.letra; Spc(1); Format$(Me.sucursal, "0000"); "-"; Format$(Me.numcomp, "00000000")
 Printer.Print
 Printer.Print
 Printer.FontSize = 12
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "Cliente.....:";
 Printer.FontBold = False
 Printer.Print "   ["; cc; "]  "; n
 Printer.Print
 Printer.FontSize = 12
 Printer.Print "Direccion...:";
 Printer.FontBold = False
 Printer.Print "   " & direccion & "  - " & localidad
 Printer.Print
 Printer.FontSize = 12
 Printer.Print "Iva.........:";
 Printer.FontBold = False
 Printer.Print "  " & di & "         ";
 Printer.FontSize = 12
 Printer.Print "Cuit.........: ";
 Printer.FontBold = False
 Printer.Print nc; "                   ";
 Printer.FontSize = 12
 Printer.Print "Condicion....: ";
 Printer.FontBold = False
 If Me.contado = "S" Then
   condicion = " CONTADO"
 Else
   condicion = " CUENTA CORRIENTE"
 End If
 Printer.Print condicion
 Printer.Print
 Printer.FontName = "Courier New"
 Printer.FontSize = 9
 Printer.Print
 Printer.Print
 Printer.Print "      ================================================================================================="
 Printer.Print "      Cant.  Unidad  Codigo  Articulo                                              P.U    IMPORTE"
 Printer.Print "      ================================================================================================="
 Set rs = New adodb.Recordset
 q = "Select * from vta_03 where [num_int] = " & Me.numint
 rs.Open q, cn1
 l = 0
 While Not rs.EOF
    d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    pu = Space$(8)
    i = Space$(9)
    cant = Space$(8)
    cp2 = Space$(5)
    tu = Space$(6)
    If rs("cantidad_original") > 0 Then
      RSet cant = Format$(rs("cantidad_original"), "####0.00")
      RSet cp2 = Format$(rs("id_producto"), "00000")
      tu = Format$(Left$(rs("tunidad"), 6), "@@@@@@!")
      RSet pu = Format$(rs("PU_FINAL"), "####0.00")
      RSet i = Format$(rs("PU_FINAL") * rs("cantidad_original"), "######0.00")
    
    End If
    Printer.Print Tab(4); cant; "  "; tu; "  "; cp2; Spc(2);
    Printer.FontBold = True
    Printer.Print d;
    Printer.FontBold = False
    Printer.Print " "; pu; " "; i
    
    rs.MoveNext
    l = l + 1
 Wend
 Set rs = Nothing
 
 
For ó = l To Me.cant_lineas
  Printer.Print
Next ó

Printer.Print
Printer.Print Spc(55); "****************************"
Printer.Print Spc(45); "TOTAL      ";
Printer.FontSize = 12
Printer.FontBold = True
Printer.Print asteriscos(Format$(Me.total, "######0.00"), 15)
Printer.FontSize = 9
Printer.FontBold = False
Printer.Print Spc(55); "****************************"
Printer.Print
Printer.Print
Printer.Print "OBSERVACIONES: " & Me.observaciones


Printer.FontSize = 8
 
 
   If Y < Val(l) Then
      Printer.NewPage
   End If
Next Y
 
Printer.EndDoc
End If

End Sub


Sub formatoX(ByVal c As Integer)
'formato en hoja A4 en blanco grafica
 comp = "PRESUPUESTO  "
 
 
l = InputBox("Indique cantidad de Copias", , c)
If Val(l) > 0 And Val(l) <= 6 Then
 
For Y = 1 To Val(l)
 Printer.FontName = "Times New Roman"
 cc = Space$(5)
 cp = Space$(9)
 cant = Space$(8)
 pu = Space$(9)
 dtop = Space$(7)
 pp = Space$(10)
 im = Space$(10)
 s = Space$(10)
 s2 = Space$(10)
 dto = Space$(7)
 i1 = Space$(9)
 I2 = Space$(9)
 tot = Space$(10)
 RSet cc = Format$(Me.idcliente, "0000")
 n = Format$(Me.razonsocial, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 d = Format$(Me.direccion, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 l = Format$(Me.localidad, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 di = Format$(Me.abreviatura_tipoiva, "@@@@@@@@@@!")
 nc = Format$(Me.CUIT, "@@@@@@@@@@@@@@@@!")
 Printer.FontSize = 14
 Printer.Print glo.nombrecli
 Printer.FontSize = 10
 Printer.Print glo.direccioncli & " - Rojas(B) - CP.2705 - TE:" & glo.TECLI
 Printer.Print "Cuit:" & glo.CUIT
 Printer.Print "______________________________________________________________________________________"
 Printer.Print
 Printer.FontSize = 16
 Printer.Print Spc(30); comp;
 Printer.FontSize = 12
 Printer.Print "(Documento no valido como Factura)"
 Printer.Print
 Printer.Print Spc(80); " Fecha.....: " & Format$(Me.fecha)
 Printer.Print Spc(80); " Nro.......: "; Me.letra; Spc(1); Format$(Me.sucursal, "0000"); "-"; Format$(Me.numcomp, "00000000")
 Printer.Print
 Printer.FontSize = 12
 Printer.FontBold = False
 Printer.FontUnderline = True
 Printer.Print "Cliente.....:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "   ["; cc; "]  "; n
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Direccion...:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "   " & direccion & "  - " & localidad
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Iva.........:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "  " & di & "         ";
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Cuit.........: ";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print nc; "                   ";
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Condicion....: ";
 Printer.FontBold = False
 Printer.FontUnderline = False
 If Me.contado = "S" Then
   condicion = " CONTADO"
 Else
   condicion = " CUENTA CORRIENTE"
 End If
 Printer.Print condicion
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Observaciones.: ";
 Printer.FontBold = False
 Printer.FontUnderline = False
 ref = "Ref. " & Me.abreviatura
 Printer.Print ref
 Printer.FontName = "Courier New"
 Printer.FontSize = 9
 Printer.Print
 Printer.Print
 Printer.Print "      ================================================================================================="
 Printer.Print "      Cant.  Unidad  Codigo  Articulo                                              P.U    IMPORTE"
 Printer.Print "      ================================================================================================="
 Set rs = New adodb.Recordset
 q = "Select * from vta_03 where [num_int] = " & Me.numint
 rs.Open q, cn1
 l = 0
 While Not rs.EOF
    d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    pu = Space$(8)
    i = Space$(9)
    cant = Space$(8)
    cp2 = Space$(5)
    tu = Space$(6)
    If rs("cantidad_original") > 0 Then
      RSet cant = Format$(rs("cantidad_original"), "####0.00")
      RSet cp2 = Format$(rs("id_producto"), "00000")
      tu = Format$(Left$(rs("tunidad"), 6), "@@@@@@!")
      RSet pu = Format$(rs("PU_FINAL"), "####0.00")
      RSet i = Format$(rs("PU_FINAL") * rs("cantidad_original"), "######0.00")
    
    End If
    Printer.Print Tab(4); cant; "  "; tu; "  "; cp2; Spc(2);
    Printer.FontBold = True
    Printer.Print d;
    Printer.FontBold = False
    Printer.Print " "; pu; " "; i
    
    rs.MoveNext
    l = l + 1
 Wend
 Set rs = Nothing
 
 
For ó = l To Me.cant_lineas
  Printer.Print
Next ó

For ó = 1 To 2
  Printer.Print
Next ó
Printer.Print Spc(35); "****************************"
Printer.Print Spc(25); "TOTAL      ";
Printer.FontSize = 12
Printer.FontBold = True
Printer.Print asteriscos(Format$(Me.total, "######0.00"), 15)
Printer.FontSize = 9
Printer.FontBold = False
Printer.Print Spc(35); "****************************"

Printer.Print
Printer.FontSize = 8
   If Y < Val(l) Then
      Printer.NewPage
   End If
Next Y
 
Printer.EndDoc
End If

End Sub


Sub formatoY(ByVal c As Integer)
'formato en hoja A4 en blanco grafica  2 x hoja
 comp = "PRESUPUESTO  "
 
 
l = InputBox("Indique cantidad de Copias", , c)
If Val(l) > 0 And Val(l) <= 6 Then
 
For Y = 1 To Val(l)
 For ki = 1 To 2
 
 Printer.FontName = "Times New Roman"
 cc = Space$(5)
 cp = Space$(9)
 cant = Space$(8)
 pu = Space$(9)
 dtop = Space$(7)
 pp = Space$(10)
 im = Space$(10)
 s = Space$(10)
 s2 = Space$(10)
 dto = Space$(7)
 i1 = Space$(9)
 I2 = Space$(9)
 tot = Space$(10)
 RSet cc = Format$(Me.idcliente, "0000")
 n = Format$(Me.razonsocial, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 d = Format$(Me.direccion, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 l = Format$(Me.localidad, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 di = Format$(Me.abreviatura_tipoiva, "@@@@@@@@@@!")
 nc = Format$(Me.CUIT, "@@@@@@@@@@@@@@@@!")
 Printer.FontSize = 10
 Printer.Print
 Printer.FontSize = 14
 Printer.Print glo.nombrecli
 Printer.FontSize = 10
 Printer.Print glo.direccioncli & " - Rojas(B) - CP.2705 - TE:" & glo.TECLI
 Printer.Print "Cuit:" & glo.CUIT
 Printer.Print "______________________________________________________________________________________"
 Printer.Print
 Printer.FontSize = 16
 Printer.Print Spc(30); comp;
 Printer.FontSize = 12
 Printer.Print "(Documento no valido como Factura)"
 Printer.Print
 Printer.Print Spc(80); " Fecha.....: " & Format$(Me.fecha)
 Printer.Print Spc(80); " Nro.......: "; Me.letra; Spc(1); Format$(Me.sucursal, "0000"); "-"; Format$(Me.numcomp, "00000000")
 Printer.Print
 Printer.FontSize = 12
 Printer.FontBold = False
 Printer.FontUnderline = True
 Printer.Print "Cliente.....:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "   ["; cc; "]  "; n
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Direccion...:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "   " & direccion & "  - " & localidad
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Iva.........:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "  " & di & "         ";
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Cuit.........: ";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print nc; "                   ";
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Condicion....: ";
 Printer.FontBold = False
 Printer.FontUnderline = False
 If Me.contado = "S" Then
   condicion = " CONTADO"
 Else
   condicion = " CUENTA CORRIENTE"
 End If
 Printer.Print condicion
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Observaciones.: ";
 Printer.FontBold = False
 Printer.FontUnderline = False
 ref = "Ref. " & Me.abreviatura
 Printer.Print ref
 Printer.FontName = "Courier New"
 Printer.FontSize = 9
 Printer.Print
 Printer.Print
 Printer.Print "      ================================================================================================="
 Printer.Print "      Cant.  Unidad  Codigo  Articulo                                              P.U    IMPORTE"
 Printer.Print "      ================================================================================================="
 Set rs = New adodb.Recordset
 q = "Select * from vta_03 where [num_int] = " & Me.numint
 rs.Open q, cn1
 l = 0
 While Not rs.EOF
    d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    pu = Space$(8)
    i = Space$(9)
    cant = Space$(8)
    cp2 = Space$(5)
    tu = Space$(6)
    If rs("cantidad_original") > 0 Then
      RSet cant = Format$(rs("cantidad_original"), "####0.00")
      RSet cp2 = Format$(rs("id_producto"), "00000")
      tu = Format$(Left$(rs("tunidad"), 6), "@@@@@@!")
      RSet pu = Format$(rs("PU_FINAL"), "####0.00")
      RSet i = Format$(rs("PU_FINAL") * rs("cantidad_original"), "######0.00")
    
    End If
    Printer.Print Tab(4); cant; "  "; tu; "  "; cp2; Spc(2);
    Printer.FontBold = True
    Printer.Print d;
    Printer.FontBold = False
    Printer.Print " "; pu; " "; i
    
    rs.MoveNext
    l = l + 1
 Wend
 Set rs = Nothing
 
 
For ó = l To Me.cant_lineas
  Printer.Print
Next ó

For ó = 1 To 2
  Printer.Print
Next ó
Printer.Print Spc(35); "****************************"
Printer.Print Spc(25); "TOTAL      ";
Printer.FontSize = 12
Printer.FontBold = True
Printer.Print asteriscos(Format$(Me.total, "######0.00"), 15)
Printer.FontSize = 9
Printer.FontBold = False
Printer.Print Spc(35); "****************************"

Printer.Print

Printer.FontSize = 8
Next ki
If Y < Val(l) Then
      Printer.NewPage
End If
Next Y
 
Printer.EndDoc
End If

End Sub

Sub formatoZ(ByVal c As Integer)
'formato en hoja A4 en blanco grafica  2 x hoja
'Idem a Y pero sin membrete
 comp = "PRESUPUESTO  "
 
 
l = InputBox("Indique cantidad de Copias", , c)
If Val(l) > 0 And Val(l) <= 6 Then
 
For Y = 1 To Val(l)
 
 Printer.FontName = "Times New Roman"
 cc = Space$(5)
 cp = Space$(9)
 cant = Space$(8)
 pu = Space$(9)
 dtop = Space$(7)
 pp = Space$(10)
 im = Space$(10)
 s = Space$(10)
 s2 = Space$(10)
 dto = Space$(7)
 i1 = Space$(9)
 I2 = Space$(9)
 tot = Space$(10)
 RSet cc = Format$(Me.idcliente, "0000")
 n = Format$(Me.razonsocial, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 d = Format$(Me.direccion, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 l = Format$(Me.localidad, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 di = Format$(Me.abreviatura_tipoiva, "@@@@@@@@@@!")
 nc = Format$(Me.CUIT, "@@@@@@@@@@@@@@@@!")
 Printer.FontSize = 10
 Printer.Print
 Printer.FontSize = 14
 Printer.Print " " 'glo.nombrecli
 Printer.FontSize = 10
 Printer.Print " " 'glo.direccioncli
 Printer.Print " " 'glo.CUIT
 Printer.Print "______________________________________________________________________________________"
 Printer.Print
 Printer.FontSize = 16
 Printer.Print Spc(30); comp;
 Printer.FontSize = 12
 Printer.Print "(Documento no valido como Factura)"
 Printer.Print
 Printer.Print Spc(80); " Fecha.....: " & Format$(Me.fecha)
 Printer.Print Spc(80); " Nro.......: "; Me.letra; Spc(1); Format$(Me.sucursal, "0000"); "-"; Format$(Me.numcomp, "00000000")
 Printer.Print
 Printer.FontSize = 12
 Printer.FontBold = False
 Printer.FontUnderline = True
 Printer.Print "Cliente.....:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "   ["; cc; "]  "; n
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Direccion...:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "   " & direccion & "  - " & localidad
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Iva.........:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "  " & di & "         ";
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Cuit.........: ";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print nc; "                   ";
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Condicion....: ";
 Printer.FontBold = False
 Printer.FontUnderline = False
 If Me.contado = "S" Then
   condicion = " CONTADO"
 Else
   condicion = " CUENTA CORRIENTE"
 End If
 Printer.Print condicion
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Observaciones.: ";
 Printer.FontBold = False
 Printer.FontUnderline = False
 ref = "Ref. " & Me.abreviatura
 Printer.Print ref
 Printer.FontName = "Courier New"
 Printer.FontSize = 9
 Printer.Print
 Printer.Print
 Printer.Print "      ================================================================================================="
 Printer.Print "      Cant.  Unidad  Codigo  Articulo                                              P.U    IMPORTE"
 Printer.Print "      ================================================================================================="
 Set rs = New adodb.Recordset
 q = "Select * from vta_03 where [num_int] = " & Me.numint
 rs.Open q, cn1
 l = 0
 While Not rs.EOF
    d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    pu = Space$(8)
    i = Space$(9)
    cant = Space$(8)
    cp2 = Space$(5)
    tu = Space$(6)
    If rs("cantidad_original") > 0 Then
      RSet cant = Format$(rs("cantidad_original"), "####0.00")
      RSet cp2 = Format$(rs("id_producto"), "00000")
      tu = Format$(Left$(rs("tunidad"), 6), "@@@@@@!")
      RSet pu = Format$(rs("PU_FINAL"), "####0.00")
      RSet i = Format$(rs("PU_FINAL") * rs("cantidad_original"), "######0.00")
    
    End If
    Printer.Print Tab(4); cant; "  "; tu; "  "; cp2; Spc(2);
    Printer.FontBold = True
    Printer.Print d;
    Printer.FontBold = False
    Printer.Print " "; pu; " "; i
    
    rs.MoveNext
    l = l + 1
 Wend
 Set rs = Nothing
 
 
For ó = l To Me.cant_lineas
  Printer.Print
Next ó

For ó = 1 To 2
  Printer.Print
Next ó
Printer.Print Spc(35); "****************************"
Printer.Print Spc(25); "TOTAL      ";
Printer.FontSize = 12
Printer.FontBold = True
Printer.Print asteriscos(Format$(Me.total, "######0.00"), 15)
Printer.FontSize = 9
Printer.FontBold = False
Printer.Print Spc(35); "****************************"

Printer.Print

Printer.FontSize = 8
If Y < Val(l) Then
      Printer.NewPage
End If
Next Y
 
Printer.EndDoc
End If

End Sub


Sub formatoW(ByVal c As Integer)
'formato en hoja A4 en blanco grafica  2 x hoja
'Idem a Y pero sin membrete(quitado todo el espacio)
 
 'arma lista de remitos
  rtos = ""
  Set rs2 = New adodb.Recordset
  q = "select * from vta_08, vta_02 where id_factura = " & Me.numint & " and id_remito = num_int"
  rs2.Open q, cn1
  While Not rs2.EOF
    rtos = rtos & Format$(rs2("num_comp"), "000000") & "-"
    rs2.MoveNext
  Wend
  Set rs2 = Nothing
 
 If Me.idtipocomp <> 3 Then
   comp = "PRESUPUESTO  "
 Else
   comp = "DEVOLUCION  "
 End If
 
lc = InputBox("Indique cantidad de Copias", , c)
If Val(lc) > 0 And Val(lc) <= 6 Then
 
For Y = 1 To Val(lc)
 
 Printer.FontName = "Times New Roman"
 cc = Space$(5)
 cp = Space$(9)
 cant = Space$(8)
 pu = Space$(9)
 dtop = Space$(7)
 pp = Space$(10)
 im = Space$(10)
 s = Space$(10)
 s2 = Space$(10)
 dto = Space$(7)
 i1 = Space$(9)
 I2 = Space$(9)
 tot = Space$(10)
 RSet cc = Format$(Me.idcliente, "0000")
 n = Format$(Me.razonsocial, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 d = Format$(Me.direccion, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
'd22 = Format$(Me.di, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 
 l = Format$(Me.localidad, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 di = Format$(Me.abreviatura_tipoiva, "@@@@@@@@@@!")
 nc = Format$(Me.CUIT, "@@@@@@@@@@@@@@@@!")
 Printer.FontSize = 16
 Printer.Print Spc(30); comp;
 Printer.FontSize = 12
 Printer.Print "(Documento no valido como Factura)"
 Printer.Print
 Printer.Print Spc(80); " Fecha.....: " & Format$(Me.fecha)
 Printer.Print Spc(80); " Nro.......: "; Me.letra; Spc(1); Format$(Me.sucursal, "0000"); "-"; Format$(Me.numcomp, "00000000")
 Printer.Print
 Printer.FontSize = 12
 Printer.FontBold = False
 Printer.FontUnderline = True
 Printer.Print "Cliente........:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "   ["; cc; "]  "; n
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Direccion....:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "   " & direccion & "  - " & localidad;
 If Me.contado = "S" Then
   condicion = " CONTADO"
 Else
   condicion = " CUENTA CORRIENTE"
 End If
 Printer.Print condicion
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Observaciones.:  ";
 Printer.FontBold = False
 Printer.FontUnderline = False
  Printer.Print rtos
 Printer.FontName = "Courier New"
 Printer.FontSize = 9
 Printer.Print "      ================================================================================================="
 Printer.Print "      Cant.  Unidad  Codigo  Articulo                                              P.U    IMPORTE"
 Printer.Print "      ================================================================================================="
 Set rs = New adodb.Recordset
 q = "Select * from vta_03 where [num_int] = " & Me.numint
 rs.Open q, cn1
 l = 0
 While Not rs.EOF
    d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    pu = Space$(8)
    i = Space$(9)
    cant = Space$(8)
    cp2 = Space$(5)
    tu = Space$(6)
    If rs("cantidad_original") > 0 Then
      RSet cant = Format$(rs("cantidad_original"), "####0.00")
      RSet cp2 = Format$(rs("id_producto"), "00000")
      tu = Format$(Left$(rs("tunidad"), 6), "@@@@@@!")
      RSet pu = Format$(rs("PU_FINAL"), "####0.00")
      RSet i = Format$(rs("PU_FINAL") * rs("cantidad_original"), "######0.00")
    
    End If
    Printer.Print Tab(4); cant; "  "; tu; "  "; cp2; Spc(2);
    Printer.FontBold = True
    Printer.Print d;
    Printer.FontBold = False
    Printer.Print " "; pu; " "; i
    
    rs.MoveNext
    l = l + 1
 Wend
 Set rs = Nothing
 
 
For ó = l To Me.cant_lineas
  Printer.Print
Next ó

'For ó = 1 To 2
  Printer.Print
'Next ó
Printer.Print Spc(35); "****************************"
Printer.Print Spc(25); "TOTAL      ";
Printer.FontSize = 12
Printer.FontBold = True
Printer.Print asteriscos(Format$(Me.total, "######0.00"), 15)
Printer.FontSize = 9
Printer.FontBold = False
Printer.Print Spc(35); "****************************"; Spc(35); Format$(Y, "0")


Printer.FontSize = 8
If Y < Val(lc) Then
      Printer.NewPage
End If
Next Y
 
Printer.EndDoc
End If

End Sub


Sub imprimirrtoY(ByVal c As Integer)
'formato en hoja A4 en blanco grafica  2 x hoja
 

 
 
l = InputBox("Indique cantidad de Copias", , c)
If Val(l) > 0 And Val(l) <= 6 Then
 If Me.idtipocomp = 45 Then
  comp = "REMITO  "
 Else
   comp = "DEVOLUCION"
 End If
For Y = 1 To Val(l)
 For ki = 1 To 2
 
 Printer.FontName = "Times New Roman"
 cc = Space$(5)
 cp = Space$(9)
 cant = Space$(8)
 pu = Space$(9)
 dtop = Space$(7)
 pp = Space$(10)
 im = Space$(10)
 s = Space$(10)
 s2 = Space$(10)
 dto = Space$(7)
 i1 = Space$(9)
 I2 = Space$(9)
 tot = Space$(10)
 RSet cc = Format$(Me.idcliente, "0000")
 n = Format$(Me.razonsocial, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 d = Format$(Me.direccion, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 l = Format$(Me.localidad, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 di = Format$(Me.abreviatura_tipoiva, "@@@@@@@@@@!")
 nc = Format$(Me.CUIT, "@@@@@@@@@@@@@@@@!")
 Printer.FontSize = 10
 Printer.Print
 Printer.FontSize = 14
 Printer.Print glo.nombrecli
 Printer.FontSize = 10
 Printer.Print glo.direccioncli & " - Rojas(B) - CP.2705 - TE:" & glo.TECLI
 Printer.Print "Cuit:" & glo.CUIT
 Printer.Print "______________________________________________________________________________________"
 Printer.Print
 Printer.FontSize = 16
 Printer.Print Spc(30); comp;
 Printer.FontSize = 12
 Printer.Print "(Documento no valido como Factura)"
 Printer.Print
 Printer.Print Spc(80); " Fecha.....: " & Format$(Me.fecha)
 Printer.Print Spc(80); " Nro.......: "; Me.letra; Spc(1); Format$(Me.sucursal, "0000"); "-"; Format$(Me.numcomp, "00000000")
 Printer.Print
 Printer.FontSize = 12
 Printer.FontBold = False
 Printer.FontUnderline = True
 Printer.Print "Cliente.....:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "   ["; cc; "]  "; n
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Direccion...:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "   " & direccion & "  - " & localidad
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Iva.........:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "  " & di & "         ";
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Cuit.........: ";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print nc; "                   ";
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Condicion....: ";
 Printer.FontBold = False
 Printer.FontUnderline = False
 If Me.contado = "S" Then
   condicion = " CONTADO"
 Else
   condicion = " CUENTA CORRIENTE"
 End If
 Printer.Print condicion
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Observaciones.: ";
 Printer.FontBold = False
 Printer.FontUnderline = False
 ref = "Ref. " & Me.abreviatura
 Printer.Print ref
 Printer.FontName = "Courier New"
 Printer.FontSize = 9
 Printer.Print
 Printer.Print
 Printer.Print "      ================================================================================================="
 Printer.Print "      Cant.  Unidad  Codigo  Articulo                                              "
 Printer.Print "      ================================================================================================="
 Set rs = New adodb.Recordset
 q = "Select * from vta_03 where [num_int] = " & Me.numint
 rs.Open q, cn1
 l = 0
 While Not rs.EOF
    d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    pu = Space$(8)
    i = Space$(9)
    cant = Space$(8)
    cp2 = Space$(5)
    tu = Space$(6)
    If rs("cantidad_original") > 0 Then
      RSet cant = Format$(rs("cantidad_original"), "####0.00")
      RSet cp2 = Format$(rs("id_producto"), "00000")
      tu = Format$(Left$(rs("tunidad"), 6), "@@@@@@!")
      'RSet pu = Format$(rs("PU_FINAL"), "####0.00")
      'RSet i = Format$(rs("PU_FINAL") * rs("cantidad_original"), "######0.00")
    
    End If
    Printer.Print Tab(4); cant; "  "; tu; "  "; cp2; Spc(2);
    Printer.FontBold = True
    Printer.Print d
    Printer.FontBold = False
    'Printer.Print " "; pu; " "; i
    
    rs.MoveNext
    l = l + 1
 Wend
 Set rs = Nothing
 
 
For ó = l To Me.cant_lineas
  Printer.Print
Next ó

For ó = 1 To 2
  Printer.Print
Next ó
'Printer.Print " "

Printer.FontSize = 12
Printer.FontBold = True
Printer.Print " "
Printer.FontSize = 9
Printer.FontBold = False
Printer.Print " "

Printer.Print

Printer.FontSize = 8
Next ki
If Y < Val(l) Then
      Printer.NewPage
End If
Next Y
 
Printer.EndDoc
End If

End Sub

Sub imprimirrtoW(ByVal c As Integer)
'formato en hoja A4 en blanco grafica  2 x hoja
 

 
 
lc = InputBox("Indique cantidad de Copias", , c)
If Val(lc) > 0 And Val(lc) <= 6 Then
 If Me.idtipocomp = 45 Then
  comp = "Nota de Pedido "
 Else
   comp = "Nota de Devolución"
 End If
For Y = 1 To Val(lc)
 Printer.FontName = "Times New Roman"
 cc = Space$(5)
 cp = Space$(9)
 cant = Space$(8)
 pu = Space$(9)
 dtop = Space$(7)
 pp = Space$(10)
 im = Space$(10)
 s = Space$(10)
 s2 = Space$(10)
 dto = Space$(7)
 i1 = Space$(9)
 I2 = Space$(9)
 tot = Space$(10)
 RSet cc = Format$(Me.idcliente, "0000")
 n = Format$(Me.razonsocial, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 d = Format$(Me.direccion, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 l = Format$(Me.localidad, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 di = Format$(Me.abreviatura_tipoiva, "@@@@@@@@@@!")
 nc = Format$(Me.CUIT, "@@@@@@@@@@@@@@@@!")
 Printer.FontSize = 16
 Printer.Print Spc(30); comp;
 Printer.FontSize = 12
 Printer.Print "(Documento no valido como Factura)"
 Printer.Print
 Printer.Print Spc(80); " Fecha.....: " & Format$(Me.fecha)
 Printer.Print Spc(80); " Nro.......: "; Me.letra; Spc(1); Format$(Me.sucursal, "0000"); "-"; Format$(Me.numcomp, "00000000")
 Printer.Print
 Printer.FontSize = 12
 Printer.FontBold = False
 Printer.FontUnderline = True
 Printer.Print "Cliente.....:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "   ["; cc; "]  "; n
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Direccion...:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "   " & direccion & "  - " & localidad
  Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Observaciones.: ";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print Me.observaciones
 Printer.FontName = "Courier New"
 Printer.FontSize = 9
 Printer.Print
 Printer.Print
 Printer.Print "      ================================================================================================="
 Printer.Print "      Cant.  Unidad  Codigo  Articulo                                              "
 Printer.Print "      ================================================================================================="
 Set rs = New adodb.Recordset
 q = "Select * from vta_03 where [num_int] = " & Me.numint
 rs.Open q, cn1
 l = 0
 While Not rs.EOF
    d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    pu = Space$(8)
    i = Space$(9)
    cant = Space$(8)
    cp2 = Space$(5)
    tu = Space$(6)
    If rs("cantidad_original") > 0 Then
      RSet cant = Format$(rs("cantidad_original"), "####0.00")
      RSet cp2 = Format$(rs("id_producto"), "00000")
      tu = Format$(Left$(rs("tunidad"), 6), "@@@@@@!")
      'RSet pu = Format$(rs("PU_FINAL"), "####0.00")
      'RSet i = Format$(rs("PU_FINAL") * rs("cantidad_original"), "######0.00")
    
    End If
    Printer.Print Tab(4); cant; "  "; tu; "  "; cp2; Spc(2);
    Printer.FontBold = True
    Printer.Print d
    Printer.FontBold = False
    'Printer.Print " "; pu; " "; i
    
    rs.MoveNext
    l = l + 1
 Wend
 Set rs = Nothing
 
 For o = l To Me.cant_lineas
  Printer.Print
  Next o
Printer.FontSize = 9
Printer.FontBold = False
Printer.Print Spc(100); Format$(Y, "0")


 If Y < Val(lc) Then
      Printer.NewPage
 End If
Next Y
 
Printer.EndDoc
End If

End Sub


Sub imprimerbo7(ByVal c As Integer)
'formato texto hugo diaz sin mostrar el saldo en el recibo
puerto = "Lpt1"

 On Error GoTo err7
 Open puerto For Output As #1
 
Set cl_cli = New Clientes
cl_cli.carga (Me.idcliente)
If cl_cli.id > 0 Then
 dd = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 1, 2)
 mm = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 4, 2)
 aa = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 7, 4)
For J = 1 To c
 Print #1, Chr$(18)
 Print #1,
 Print #1,
 Print #1,
 Print #1,
 Print #1, Spc(59); dd; "/"; mm; "/"; aa; "   "; Format$(Me.sucursal, "0000"); "-"; Format$(Me.numcomp, "00000000")
 Print #1,
 Print #1,
 Print #1,
 Print #1,
 Print #1,
 Print #1, Chr$(27); Chr$(69); "Recibimos de :"; Chr$(27); Chr$(70);
 Print #1, Spc(3); Left$(Me.razonsocial, 40); "  ("; Format$(Me.idcliente, "00000"); ")"
 Print #1, Chr$(27); Chr$(69); "Direccion    :"; Chr$(27); Chr$(70);
 Print #1, Spc(3); Left$(Me.direccion, 35); "      "; Chr$(27); Chr$(69); "Localidad    :"; Chr$(27); Chr$(70);
 Print #1, Spc(3); Me.localidad
 Print #1, Chr$(27); Chr$(69); "C.U.I.T.     :"; Chr$(27); Chr$(70);
 Print #1, Spc(3); Me.CUIT; "   "; Me.abreviatura_tipoiva
 Print #1,
 Print #1, Chr$(27); Chr$(69); "Son $: "; Chr$(27); Chr$(70);
 Print #1, convierte(Me.total) & "----"
 Print #1,
 Print #1, Chr$(27); Chr$(69); "En Concepto de:"; Chr$(27); Chr$(70);
 Print #1, Spc(3); Me.observaciones
 Print #1,
 Print #1, Chr$(15);
 Print #1, "Forma de Pago                                                        Comprobantes Cancelados      "
 Print #1, "-----------------------------------------------------------------    --------------------------------------------"
 Print #1, " Nro.Ch        Fecha           Banco               Importe            Fecha      Nro. Comprobante       Importe"
 Print #1, "-----------------------------------------------------------------    --------------------------------------------"
 sch = 0
 p1 = 0
 p2 = 0
 'forma de pago
 Set rs = New adodb.Recordset
 q = "Select * from vta_04 where [num_int] = " & Me.numint
 rs.Open q, cn1
 i = Space$(10)
 
 Set rs2 = New adodb.Recordset
 q = "Select * from vta_02 where [estado_pago] = 'P' and [recibo_pago] = '" & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000") & "' AND [ID_TIPOCOMP] <= 10 "
 rs2.Open q, cn1

 
 While Not rs.EOF
    nch = Format$(rs("num_ch"), "0000000000")
    db = Format$(Left$(rs("detalle_Banco") & " ", 25), "@@@@@@@@@@@@@@@@@@@@@@@@@!")
    fd = Format$(rs("fecha_dif"), "dd/mm/yyyy")
    RSet i = Format$(rs("importe"), "######0.00")
    renglon1 = nch & "   " & fd & "   " & db & "   " & i
    
    If Not rs2.EOF And Not rs2.BOF Then
        fc = Format$(rs2.Fields("fecha"), "dd/mm/yyyy")
        nc = rs2.Fields("letra") & Format$(rs2.Fields("sucursal"), "0000") & "-" & Format$(rs2.Fields("num_comp"), "00000000")
        RSet i = Format$(rs2("total"), "######0.00")
        renglon1 = renglon1 & "          " & fc & "  " & nc & " " & i
        rs2.MoveNext
    End If
   Print #1, renglon1
   p1 = p1 + 1
   rs.MoveNext
 Wend
 If Not rs2.BOF And Not rs2.BOF Then
   While Not rs2.EOF
          renglon1 = Space$(63)
          fca = rs2.Fields("fecha")
          nc = rs2.Fields("letra") & Format$(rs2.Fields("sucursal"), "0000") & "-" & Format$(rs2.Fields("num_comprobante"), "00000000")
          RSet i = Format$(rs2("total"), "######0.00")
          renglon1 = renglon1 & "          " & fca & "  " & nc & " " & i
          rs2.MoveNext
          Print #1, renglon1
          p1 = p1 + 1
          rs2.MoveNext
    Wend
 End If
 Set rs = Nothing
 Set rs2 = Nothing
 Print #1,
 Print #1, Chr$(18);
 Print #1, Chr$(27); Chr$(69); "Importe Recibido:"; Chr$(27); Chr$(70);
 Print #1, asteriscos(Format$(Me.total, "#####0.00"), 10) & "***"
 For z = p1 To Me.cant_lineas
   Print #1,
 Next z
Next J
Close #1
End If
Exit Sub

err7:
i = MsgBox("Error de Impresion.Verifique la Impresora", 4)
If i = 6 Then
   Resume
Else
   Exit Sub
End If

End Sub
Sub imprimerbo6(ByVal c As Integer)
l = InputBox("Indique cantidad de Copias", , c)
If Val(l) > 0 And Val(l) <= 5 Then
 
 For Y = 1 To Val(l)
   Printer.FontName = "Times New Roman"
   Printer.FontSize = 10
   Printer.FontBold = False
   For X = 1 To 7
     Printer.Print
   Next X
   Printer.Print
   Printer.FontName = "Times New Roman"
   Printer.FontSize = 12
   Printer.Print Tab(80); Me.fecha
   For X = 1 To 4
     Printer.Print
   Next X

   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print Tab(15); "RECIBO OFICIAL"
   Printer.FontSize = 10
   Printer.Print Tab(80); "Nro. Comprobante:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.FillStyle = False
   Printer.Print "  " & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000")
   Printer.FontUnderline = True
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(93); "Fecha:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & Format$(Me.fecha, "dd/mm/yyyy")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(15); "Cliente:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
   Set cl_cli = New Clientes
   cl_cli.carga (Me.idcliente)
   If cl_cli.id > 0 Then
      Printer.Print "  " & cl_cli.razonsocial; " ("; cl_cli.id; ") "; Tab(92);
   Else
      Printer.Print "  "; Tab(92);
   End If
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print "Emitido por:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Set cl_usuarios = New usuarios
   cl_usuarios.cargar (Me.idusuario)
   If cl_usuarios.idusuario > 0 Then
        Printer.Print "  " & cl_usuarios.denominACION
   Else
        Printer.Print
   End If
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print Tab(15); "Direccion:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.Print "  " & cl_cli.direccion
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print Tab(15); "Localidad:";
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & cl_cli.localidad & " " & cl_cli.cp
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(21); "Cuit:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & cl_cli.CUIT

'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(15); "___________________________________________________________________________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(15); "Recibí/mos de " & cl_cli.razonsocial
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(15); "La suma de pesos: *****" & convierte(Me.total) & "*****"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(15); "___________________________________________________________________________________________________"

'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 6
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
   Printer.FontSize = 16
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(25); "_______________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(30); "Son $ "; asteriscos(Format$(Me.total, "######0.00"), 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(25); "_______________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 6
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
   Printer.Print Tab(12); "Forma de Pago"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Courier new"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 8
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "----------------------------------------------------------------------------------------------"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "Forma de Pago       Num.Ch.    Detalle/Banco               Fecha Dif.    Importe  Sucursal"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "----------------------------------------------------------------------------------------------"
   Set rs = New adodb.Recordset
   q = "Select * from vta_04 where [num_int] = " & Me.numint
   rs.Open q, cn1
   i = Space$(10)
   While Not rs.EOF
    ifp = Format$(rs("id_formapago"), "000")
    FP = Format$(Left$(rs("formapago"), 12), "@@@@@@@@@@@@!")
    nch = Format$(rs("num_ch"), "0000000000")
    db = Format$(Left$(rs("detalle_Banco") & " ", 27), "@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    s = Format$(Left$(rs("sucursal") & " ", 10), "@@@@@@@@@@!")
    fd = Format$(rs("fecha_dif"), "dd/mm/yyyy")
    RSet i = Format$(rs("importe"), "######0.00")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(12); "(" & ifp & ") " & FP & "  " & nch & " " & db & " " & fd & " " & i & " " & s
    rs.MoveNext
   Wend
   Set rs = Nothing
   
   Printer.FontSize = 8
   Printer.Print
   Printer.Print
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "Comprobantes Aplicados"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Courier new"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "-------------------------------------------------------"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "Fecha        Comprobante            Importe   Obs."
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "-------------------------------------------------------"
   Set rs = New adodb.Recordset
   q = "Select * from vta_02 where [estado_pago] = 'P' and [recibo_pago] = '" & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000") & "' AND [ID_TIPOCOMP] <= 10 "
   rs.Open q, cn1
   TCOMP = 0
   While Not rs.EOF
    Set rs1 = New adodb.Recordset
    q = "select * from vta_06 where [id_tipocomp] = " & rs("id_tipocomp") & " and [sucursal] = " & rs("sucursal_INGRESO")
    rs1.Open q, cn1
    If Not rs1.EOF And Not rs1.BOF Then
      d = rs1("abreviatura")
    Else
      d = " "
    End If
    Set rs1 = Nothing
    nc = "[" & Format$(rs("id_tipocomp"), "000") & "] " & rs("letra") & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
    fd = Format$(rs("fecha"), "dd/mm/yyyy")
    RSet i = Format$(rs("total"), "######0.00")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(12); fd & "  " & nc & "  " & i & "  " & d
    
    If rs("cta_cte") = "D" Then
      TCOMP = TCOMP + rs("total")
    Else
       TCOMP = TCOMP - rs("total")
    End If
       
    rs.MoveNext
   Wend
   Set rs = Nothing
   Printer.Print Tab(12); "  TOTAL COMPROBANTES APLICADOS:     " & Format$(TCOMP, "#######0.00")
   
   
   Printer.Print
   Printer.Print
   Printer.FontBold = True
   Printer.Print Tab(12); "Retenciones Ingresadas"
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Courier new"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 8
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "-------------------------------------------------------"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "Fecha        Retencion              Importe   Obs."
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "-------------------------------------------------------"
   Set rs = New adodb.Recordset
   q = "Select * from vta_02 where [estado_pago] = 'P' and [recibo_pago] = '" & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000") & "' AND [ID_TIPOCOMP] >= 100 "
   rs.Open q, cn1
   TRET = 0
   While Not rs.EOF
    Set rs1 = New adodb.Recordset
    q = "select * from vta_06 where [id_tipocomp] = " & rs("id_tipocomp") & " and [sucursal] = " & rs("sucursal_INGRESO")
    rs1.Open q, cn1
    If Not rs1.EOF And Not rs1.BOF Then
      d = rs1("abreviatura")
    Else
      d = " "
    End If
    Set rs1 = Nothing
    nc = "[" & Format$(rs("id_tipocomp"), "000") & "] " & rs("letra") & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
    fd = Format$(rs("fecha"), "dd/mm/yyyy")
    RSet i = Format$(rs("total"), "######0.00")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(12); fd & "  " & nc & "  " & i & "  " & d
    TRET = TRET + rs("TOTAL")
    rs.MoveNext
   Wend
   Set rs = Nothing
   Printer.Print Tab(12); "  TOTAL RETENCIONES INGRESADAS:       " & Format$(TRET, "#######0.00")
   
   
   Printer.FontName = "Times New Roman"
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.FontSize = 10
   Printer.Print
   Printer.Print Tab(17); "Cotizacion U$S...:";
   Printer.FontUnderline = False
   Printer.Print "  "; Format$(Me.cotizaciondolar, "####0.000")
   Printer.Print
   Printer.FontUnderline = True
   Printer.Print Tab(15); "Observaciones....:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  "; Me.observaciones
   Printer.Print Tab(95); "______________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(95); "   por "; glo.nombrecli
   
   Set cl_cli = Nothing
   If Y < Val(l) Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.NewPage
   End If
  Next Y
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
  Printer.EndDoc
End If

End Sub

Sub imprimirrto8(ByVal c As Integer)
 Dim a(5) As String
 
 
 
 dd = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 1, 2)
 mm = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 4, 2)
 aa = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 7, 4)
 
 tacha = Me.abreviatura & " "
 tachan = Me.letra & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000")
 For Y = 1 To c
  Printer.FontName = "Times New Roman"
  Printer.FontSize = 10
  Printer.FontBold = False
  For i = 1 To 8
    Printer.Print
  Next i
   Printer.FontName = "Times New Roman"
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.FontSize = 12
   Printer.Print Tab(72); tacha;
   Printer.FontSize = 8
   Printer.Print tachan;
   Printer.FontSize = 12
   Printer.Print Tab(94); Spc(4); dd & "/" & mm & "/" & aa
   For i = 1 To 3
     Printer.Print
   Next i
 Printer.FontSize = 8
 Printer.Print
 Printer.Print
 Printer.FontName = "Times New Roman"
 Printer.FontSize = 14
 Printer.FontBold = True
 Printer.Print Tab(16); Me.razonsocial
 Printer.FontName = "Courier New"
 Printer.FontSize = 10
 Printer.FontBold = False
 Printer.Print Tab(65); "["; Format$(Me.idcliente, "0000"); "] ";
 Printer.Print
 Printer.Print Tab(16); Me.direccion
 Printer.Print Tab(16); Me.localidad; Tab(65); Me.CUIT
 Printer.Print
 Printer.Print Tab(16); "Condicion Iva: "; Me.abreviatura_tipoiva
 Printer.FontName = "Courier New"
 Printer.FontSize = 9
 Printer.FontBold = False
 Printer.Print
 Printer.Print
 Printer.Print
 Printer.Print
 Printer.Print
 Printer.Print
' Printer.Print
   Set rs = New adodb.Recordset
   q = "Select * from vta_03 where [num_int] = " & Me.numint
   rs.Open q, cn1
   cant = Space$(10)
   l = 0
   While Not rs.EOF
    i = Space$(10)
    t = Space$(4)
    pu = Space$(10)
    d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    dt = Space$(9)
    cant = Space$(8)
    cp2 = Space$(9)
    RSet cant = Format$(rs("cantidad"), "####0.00")
    cp2 = " [" & Format$(rs("id_producto"), "00000") & "] "
    uni = Format$(Left$(rs("tunidad"), 6), "@@@@@@!")
    If Val(cant) > 0 Then
     Printer.Print Tab(6); cant; Spc(1); uni; Spc(1); cp2; Spc(1);
     Printer.FontBold = True
     Printer.Print d; Spc(1);
     Printer.FontBold = False
    Else
     cant = "        "
     uni = "      "
     cp2 = "         "
     Printer.Print Tab(6); cant; Spc(1); uni; Spc(1); cp2; Spc(1);
     Printer.FontBold = True
     Printer.Print d; Spc(1);
     Printer.FontBold = False
     
     End If
    'imprime extra
     If Me.imprimedescextra = "S" Then
       Set rs1 = New adodb.Recordset
       q = "select * from vta_015 where [num_int] = " & Me.numint & " and [renglon] = " & rs("renglon")
       rs1.Open q, cn1
       If Not rs1.EOF And Not rs1.BOF Then
         Call lee_desc_extra(a, rs1("desc_ext"))
         For k = 0 To 4
           If a(k) <> "%%" Then
             d = Format$(Left$(a(k), 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
             Printer.Print Tab(25); d
             l = l + 1
           End If
         Next k
       
       End If
       Set rs1 = Nothing
     End If
 
 
 
    rs.MoveNext
    l = l + 1
   Wend
   Set rs = Nothing
  
   For X = l To Me.cant_lineas + 19
     Printer.Print
   Next X
   
  ' If Me.letra = "B" Then
  '   For i = 1 To 3
  '    Printer.Print
  '   Next i
  ' Else
  ' End If
   Set cl_usuarios = New usuarios
   cl_usuarios.cargar (Me.idusuario)
   If cl_usuarios.idusuario > 0 Then
        Printer.Print Tab(10); "Emitida por  : " & cl_usuarios.denominACION;
   Else
        'Printer.Print
   End If
   xs = Space$(10)
   RSet xs = Format$(Me.subtotal, "######0.00")
   If Me.letra = "A" Then
     Printer.Print Tab(93); xs
   Else
     Printer.Print
   End If

   Set cl_usuarios = Nothing
   
   If Y < c Then
      Printer.NewPage
   End If
  Next Y
 
 Printer.EndDoc

End Sub

Sub imprimirrto7(ByVal c As Integer)
  'FORMATO HUGO DIAZ TEXTO
Set cl_cli = New Clientes
cl_cli.carga (Me.idcliente)
puerto = "lpt1"
If cl_cli.id > 0 Then
 'On Error GoTo IMP2
 Open puerto For Output As #1
   sonpesos = convierte(Me.total)
   dd = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 1, 2)
   mm = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 4, 2)
   aa = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 7, 4)
   For Y = 1 To c
    cc = Space$(5)
    cp = Space$(9)
    cn = Space$(8)
    pu = Space$(10)
    dtop = Space$(7)
    pp = Space$(10)
    im = Space$(10)
    s = Space$(10)
    s2 = Space$(10)
    dto = Space$(7)
    i1 = Space$(9)
    I2 = Space$(9)
    tot = Space$(10)
    RSet cc = Format$(Me.idcliente, "0000")
    n = Left$(Format$(Me.razonsocial, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!"), 30)
    d = Left$(Format$(Me.direccion, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!"), 30)
    l = Left$(Format$(Me.localidad, "@@@@@@@@@@@@@@@@@@@!"), 15)
    di = Left$(Format$(Me.abreviatura_tipoiva, "@@@@@@@@@@@@@@@@@@@@@@@@@@!"), 25)
    nc = Left$(Format$(Me.CUIT, "@@@@@@@@@@@@@@@@!"), 16)
    Print #1, Chr$(18)
    Print #1, Spc(68); Chr$(14); Me.abreviatura
    Print #1,
    Print #1,
    Print #1, Spc(58); dd & "/" & mm & "/" & aa; Spc(2); Format$(Me.letra, ">@"); Spc(1); Format$(Me.sucursal, "0000"); "-"; Format$(Me.numcomp, "00000000")
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1, Spc(8); n; Spc(17); "("; cc; ")"
    Print #1,
    Print #1, Spc(9); d; " "; Spc(17); nc
    Print #1, Spc(9); "(" & l & ")"
    Print #1, Spc(9); di
    Print #1,
    Print #1,
    Print #1, Spc(10); Left$(Format$(Me.transporte, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!"), 30); Spc(15); Me.cuit_transporte
    Print #1, Spc(3); "Chofer:"; Left$(Format$(Me.CHOFER, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!"), 33); Spc(5); "D.N.I"; Me.dnichofer
    Print #1, Chr$(18); Spc(10); Me.dir_transporte; ; Chr$(15); "       Dom.Chasis/Acop.: "; Me.PATENTE; " / "; Me.PATENTE_ACOPLADO
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Set rs = New adodb.Recordset
    q = "Select * from vta_03 where [num_int] = " & Me.numint
    rs.Open q, cn1
    k = 0
    While Not rs.EOF
        i = Space$(10)
        t = Space$(4)
        pu = Space$(10)
        cant = Space$(8)
        cp = Space$(7)
        d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
        dt = Space$(9)
        u = Space$(5)
        If rs("cantidad") > 0 Then
            RSet cant = Format$(rs("cantidad"), "####0.00")
            u = Format$(rs("tunidad"), "@@@@@!")
            cp = "(" & Format$(rs("id_producto"), "00000") & ")"
        End If
        Print #1, cant; Spc(1); u; Spc(2); cp; Spc(4); d
        rs.MoveNext
        k = k + 1
   Wend
   Set rs = Nothing
   
    For z = k To Me.cant_lineas
        Print #1,
    Next z
    
    
    Print #1, 'Chr$(18)
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    'Print #1,

    Next Y
    ok = 1
    Close #1

End If
Exit Sub


IMP2:
i = MsgBox("La impresora no esta lista.Verifique", 4)
If i = 6 Then
    Resume
Else
    ok = 0
    Close #1
    Exit Sub
End If




End Sub
Sub imprimirrto6(ByVal c As Integer)

'Agrosumma
 dd = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 1, 2)
 mm = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 4, 2)
 aa = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 7, 4)
 'lugar = Left$(RTrim$(cl_cli.direccion) & " - (" & RTrim$(cl_cli.cp) & ")" & RTrim$(cl_cli.localidad), 120)

 If Me.idtipocomp = 45 Then
  tacha = " "
 Else
   tacha = "#########                                           " & "DEVOL."
 End If
 
h = InputBox("Indique cantidad de Copias", , c)
If Val(h) > 0 And Val(h) <= 6 Then
 For Y = 1 To h
  Printer.FontName = "Times New Roman"
  Printer.FontSize = 10
  Printer.FontBold = False
  For i = 1 To 5
    Printer.Print
  Next i
  Printer.FontName = "Times New Roman"
  Printer.FontUnderline = False
  Printer.FontBold = False
  Printer.FontSize = 12
  Printer.Print Tab(65); tacha
  Printer.Print Tab(80); Spc(4); dd & "/" & mm & "/" & aa
  Printer.FontName = "Times New Roman"
  Printer.FontSize = 10
  Printer.FontBold = False
  Printer.Print
  Printer.Print Tab(98); Me.abreviatura; " "; Format$(Me.sucursal, "0000"); "-" & Format$(Me.numcomp, "00000000")
  Printer.Print
  Printer.Print
  Printer.FontSize = 14
 Printer.Print
 Printer.FontName = "Times New Roman"
 Printer.FontBold = True
 Printer.Print Tab(17); "["; Format$(Me.idcliente, "0000"); "] "; Me.razonsocial
 Printer.FontName = "Courier New"
 Printer.FontSize = 8
 Printer.FontBold = False
 Printer.Print
 Printer.FontSize = 10
 Printer.Print Tab(17); Me.direccion
 Printer.FontSize = 8
 Printer.Print Tab(22); Me.localidad
 Printer.FontSize = 8
 Printer.Print
 Printer.FontSize = 10
 Printer.Print Tab(20); Me.abreviatura_tipoiva; Tab(58); Me.CUIT
 Printer.Print
 Printer.Print
 If Me.moneda = "P" Then
   mda = " Pesos ($)"
 Else
   mda = "Dolares (U$s)"
 End If
 
 If Me.contado = "N" Then
   Printer.Print Tab(20); "Cuenta Corriente en "; mda
 Else
   Printer.Print Tab(20); "Contado en "; mda
 End If
 Printer.Print
 Printer.Print
 Printer.Print Tab(15); Me.transporte
 Printer.FontSize = 8
 Printer.Print
 Printer.FontSize = 10
 Printer.Print Tab(15); Me.dir_transporte; Tab(58); Me.PATENTE
 Printer.FontSize = 8
 Printer.Print
 Printer.FontSize = 10
 Printer.Print Tab(15); Me.cuit_transporte; Tab(58); Me.PATENTE_ACOPLADO

 Printer.FontName = "Courier New"
 Printer.FontSize = 9
 Printer.FontBold = False
 Printer.Print
 Printer.Print
 Printer.Print
 Set rs = New adodb.Recordset
 q = "Select * from vta_03 where [num_int] = " & Me.numint
 rs.Open q, cn1
 l = 0
 While Not rs.EOF
    i = Space$(10)
    t = Space$(4)
    pu = Space$(10)
    d = Format$(Left$(rs("descripcion") & " ", 45), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    dt = Space$(9)
    cant = Space$(10)
    cp2 = Space$(9)
   If rs("cantidad") > 0 Then
      RSet cant = Format$(rs("cantidad"), "######0.00")
    End If
    Printer.Print Tab(6); Format$(rs("id_producto"), "00000"); Tab(15);
    Printer.FontBold = True
    Printer.Print d;
    Printer.FontBold = False
    Printer.Print "   "; Format$(rs("tUnidad"), "@@@@@@@@!"); "                  "; cant
    rs.MoveNext
    l = l + 1
   Wend
   Set rs = Nothing
  
   
   For i = l To 16
        Printer.Print
   Next i
   Printer.FontName = "Times new Roman"
   Printer.FontSize = 10
   Printer.Print Tab(120); Format$(Me.valor_declarado, "######0.00")
   Printer.FontSize = 8
   Printer.Print
   Printer.FontSize = 10
   Printer.Print Tab(120); Format$(Me.bultos_totales, "#####0")
   Printer.Print
   Printer.Print
   Printer.Print Tab(10); Me.observaciones
   Printer.Print
   Printer.Print Tab(10); "Comprobante.......: " & Me.abreviatura & "  " & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000")
   Set cl_usuarios = New usuarios
   cl_usuarios.cargar (Me.idusuario)
   If cl_usuarios.idusuario > 0 Then
        Printer.Print Tab(10); "Emitida por.......: " & cl_usuarios.denominACION
   Else
        Printer.Print
   End If
   Set cl_usuarios = Nothing
   If Y < c Then
      Printer.NewPage
   End If
  Next Y
 
 Printer.EndDoc
End If
End Sub

Sub imprimirrto2(ByVal c As Integer)
Set cl_cli = New Clientes
cl_cli.carga (Me.idcliente)


If cl_cli.id > 0 Then
 dd = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 1, 2)
 mm = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 4, 2)
 aa = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 7, 4)
 lugar = Left$(RTrim$(cl_cli.direccion) & " - (" & RTrim$(cl_cli.cp) & ")" & RTrim$(cl_cli.localidad), 120)

 If Me.moneda = "P" Then
   sondolares = convierte(Me.totalotramoneda)
   
 Else
   sondolares = convierte(Me.total)
   
 End If
 
lw = InputBox("Indique cantidad de Copias", , c)
If Val(lw) > 0 And Val(lw) <= 5 Then

 For Y = 1 To Val(lw)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Times New Roman"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
    cc = Space$(5)
    cp = Space$(10)
    cp2 = Space$(6)
    cant = Space$(8)
    pu = Space$(10)
    dtop = Space$(7)
    pp = Space$(10)
    im = Space$(10)
    s = Space$(9)
    s2 = Space$(9)
    dto = Space$(9)
    i1 = Space$(9)
    I2 = Space$(9)
    tot = Space$(9)
    totd = Space$(9)
    fd = Space$(6)
    piva = Space$(10)
    pib = Space$(10)
    RSet cc = Format$(Me.idcliente, "00000")
    RSet s = Format$(Me.subtotal, "0.00")
    RSet dto = Format$(Me.descuento, "0.00")
    RSet i1 = Format$(Me.iva, "0.00")
    RSet tot = Format$(Me.total, "0.00")
    RSet totd = Format$(Me.totalotramoneda, "0.00")
    RSet fd = Format$(Me.cotizaciondolar, "#0.000")
    RSet pib = Format$(Me.percib, "0.00")
    RSet piva = Format$(Me.perciva, "0.00")
    n = Format$(Left$(cl_cli.razonsocial, 40), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    d = Format$(Left$(cl_cli.direccion, 40), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    l = Format$(Left$("(" & cl_cli.cp & ") " & cl_cli.localidad, 35), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    di = Format$(Left$(cl_cli.abreviatura_tipoiva, 21), "@@@@@@@@@@@@@@@@@@@@@!")
    nc = Format$(Left$(cl_cli.CUIT, 13), "@@@@@@@@@@@@@!")
    
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
     Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
     Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
     Printer.FontName = "Times New Roman"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
     Printer.FontSize = 12
     If Me.idtipocomp = 45 Then
        t = "          "
     Else
        t = "Devolucion"
     End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.FontBold = False
 
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
     Printer.Print Tab(95); t
    For i = 1 To 4
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
     Printer.Print
    Next i
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(87); Spc(4); dd & "/" & mm & "/" & aa
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(80); Spc(2); Me.abreviatura; Spc(1); Format$(Me.sucursal, "0000"); "-"; Format$(Me.numcomp, "00000000")
    For i = 1 To 5
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
    Next i
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Times New Roman"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 14
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(17); "["; cc; "] "; n
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(17); d; Tab(66);
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 9
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print l
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(20); di; Tab(68); nc
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 9
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
    For i = 1 To 5
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
        Printer.Print
    Next i
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 11
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
    
    Set rs = New adodb.Recordset
    q = "Select * from vta_03 where [num_int] = " & Me.numint
    rs.Open q, cn1
    cp2 = Space$(3)
    cant = Space$(8)
    pu = Space$(10)
    dtop = Space$(6)
    u = Space$(5)
    b = Space$(6)
    l = 0
    While Not rs.EOF
       d = Format$(Left$(rs("descripcion") & " ", 55), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
       If rs("cantidad_original") > 0 Then
         RSet cp2 = Format$(rs("id_producto"), "000")
         RSet cant = Format$(rs("cantidad"), "####0.00")
         RSet pu = Format$(rs("pu"), "#####0.00")
         RSet im = Format$(rs("importe"), "######0.00")
         RSet b = Format$(rs("bultos"), "####0")
         u = Format$(Left$(rs("tunidad"), 5), "@@@@@!")
         Printer.Print cant; Spc(1); u; Spc(1); "[" & cp2 & "]";
         Printer.Print d; Spc(1); b
       End If
              
       rs.MoveNext
       l = l + 1
   Wend
   Set rs = Nothing
     
   
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 9
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Line Step(700, 1)-Step(9950, 1)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 11
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
    If Len(Me.observaciones) > 5 Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print Tab(15); Me.observaciones
    Else
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
    End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print

    For i = l To 14
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
        Printer.Print
    Next i
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 9
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontBold = False
    'Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(22); "("; Format$(Me.idtransporte, "000"); ") "; Me.transporte
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(22); Me.dir_transporte; Tab(80); Me.cuit_transporte
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(20); Format$(Me.bultos_totales, "#####0"); " Bultos"; Tab(80); "Pesos "; asteriscos(Format$(Me.valor_declarado, "0.00"), 10)
   
   
 If Y < c Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.NewPage
 End If
Next Y
End If

ok = 1
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
Printer.EndDoc
End If
Exit Sub

imp1:
i = MsgBox("La impresora no esta lista.Verifique", 4)
If i = 6 Then
    Resume
Else
    ok = 0
    Exit Sub
End If


Exit Sub


End Sub


Sub imprimirrto9(ByVal c As Integer)
Set cl_cli = New Clientes
cl_cli.carga (Me.idcliente)


If cl_cli.id > 0 Then
 dd = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 1, 2)
 mm = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 4, 2)
 aa = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 7, 4)
 lugar = Left$(RTrim$(cl_cli.direccion) & " - (" & RTrim$(cl_cli.cp) & ")" & RTrim$(cl_cli.localidad), 120)

 If Me.moneda = "P" Then
   sondolares = convierte(Me.totalotramoneda)
   
 Else
   sondolares = convierte(Me.total)
   
 End If
 
lw = InputBox("Indique cantidad de Copias", , c)
If Val(lw) > 0 And Val(lw) <= 5 Then

 For Y = 1 To Val(lw)
    Printer.FontName = "Times New Roman"
    Printer.FontSize = 10
    Printer.FontBold = False
    cc = Space$(5)
    cp = Space$(10)
    cp2 = Space$(6)
    cant = Space$(8)
    pu = Space$(10)
    dtop = Space$(7)
    pp = Space$(10)
    im = Space$(10)
    s = Space$(9)
    s2 = Space$(9)
    dto = Space$(9)
    i1 = Space$(9)
    I2 = Space$(9)
    tot = Space$(9)
    totd = Space$(9)
    fd = Space$(6)
    piva = Space$(10)
    pib = Space$(10)
    RSet cc = Format$(Me.idcliente, "00000")
    RSet s = Format$(Me.subtotal, "0.00")
    RSet dto = Format$(Me.descuento, "0.00")
    RSet i1 = Format$(Me.iva, "0.00")
    RSet tot = Format$(Me.total, "0.00")
    RSet totd = Format$(Me.totalotramoneda, "0.00")
    RSet fd = Format$(Me.cotizaciondolar, "#0.000")
    RSet pib = Format$(Me.percib, "0.00")
    RSet piva = Format$(Me.perciva, "0.00")
    n = Format$(Left$(cl_cli.razonsocial, 40), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    d = Format$(Left$(cl_cli.direccion, 40), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    l = Format$(Left$("(" & cl_cli.cp & ") " & cl_cli.localidad, 35), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    di = Format$(Left$(cl_cli.abreviatura_tipoiva, 21), "@@@@@@@@@@@@@@@@@@@@@!")
    nc = Format$(Left$(cl_cli.CUIT, 13), "@@@@@@@@@@@@@!")
     Printer.Print
     Printer.Print
     Printer.FontName = "Times New Roman"
     Printer.FontSize = 12
     If Me.idtipocomp = 45 Then
        t = "          "
     Else
        t = "Devolucion"
     End If
      Printer.FontBold = False
     Printer.Print Tab(95); t
    For i = 1 To 5
     Printer.Print
    Next i
    Printer.FontSize = 14
    Printer.Print Tab(80); Spc(4); dd & "     " & mm & "   " & aa
     Printer.FontSize = 13
    For i = 1 To 2
        Printer.Print
    Next i
    Printer.FontName = "Times New Roman"
    Printer.FontSize = 14
    Printer.FontBold = True
    Printer.Print Tab(19); "["; cc; "] "; n
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.Print
    Printer.FontSize = 12
    Printer.Print Tab(17); d & "  " & l
    Printer.FontSize = 10
    Printer.Print
    Printer.FontSize = 12
    Printer.Print Tab(20); di; Tab(85); nc
    Printer.FontName = "Courier New"
    Printer.FontSize = 9
    Printer.FontBold = False
    Printer.Print
    Printer.Print Tab(80); Spc(2); Me.abreviatura; Spc(1); Format$(Me.sucursal, "0000"); "-"; Format$(Me.numcomp, "00000000")
    
    For i = 1 To 5
        Printer.Print
    Next i
    Printer.FontBold = True
    'Printer.Print
    'Printer.Print
    Printer.FontName = "Courier New"
    Printer.FontSize = 11
    Printer.FontBold = False
    
    Set rs = New adodb.Recordset
    q = "Select * from vta_03 where [num_int] = " & Me.numint
    rs.Open q, cn1
    cp2 = Space$(3)
    cant = Space$(8)
    pu = Space$(10)
    dtop = Space$(6)
    u = Space$(5)
    b = Space$(6)
    l = 0
    While Not rs.EOF
       d = Format$(Left$(rs("descripcion") & " ", 55), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
       If rs("cantidad_original") > 0 Then
         RSet cp2 = Format$(rs("id_producto"), "000")
         RSet cant = Format$(rs("cantidad"), "####0.00")
         RSet pu = Format$(rs("pu"), "#####0.00")
         RSet im = Format$(rs("importe"), "######0.00")
         'RSet b = Format$(rs("bultos"), "####0")
         u = Format$(Left$(rs("tunidad"), 5), "@@@@@!")
         Printer.Print Tab(5); cant; Spc(1); u; Spc(1);
         Printer.Print d
       End If
              
       rs.MoveNext
       l = l + 1
   Wend
   Set rs = Nothing
     
   
    Printer.FontSize = 9
    Printer.Print
    Printer.Line Step(700, 1)-Step(9950, 1)
    Printer.FontName = "Courier New"
    Printer.FontSize = 11
    Printer.FontBold = True
    Printer.Print
    If Len(Me.observaciones) > 5 Then
       Printer.Print Tab(15); Me.observaciones
    Else
       Printer.Print
    End If
   Printer.Print

    For i = l To 23
        Printer.Print
    Next i
    Printer.FontBold = False
    Printer.FontName = "Courier New"
    Printer.FontSize = 12
    Printer.FontBold = False
    Printer.Print Tab(22); "("; Format$(Me.idtransporte, "000"); ") "; Me.transporte
    Printer.Print
    Printer.Print Tab(22); Me.dir_transporte
    Printer.Print
    Printer.Print Tab(22); Me.cuit_transporte
   
 If Y < c Then
   Printer.NewPage
 End If
Next Y
End If

ok = 1
Printer.EndDoc
End If
Exit Sub

imp1:
i = MsgBox("La impresora no esta lista.Verifique", 4)
If i = 6 Then
    Resume
Else
    ok = 0
    Exit Sub
End If


Exit Sub


End Sub



Sub imprimirempaque(ByVal c As Integer)
'formato en hoja A4 en blanco grafica sin valorizar
Dim a(5) As String
  
l = InputBox("Indique cantidad de Copias", , c)
If Val(l) > 0 And Val(l) <= 6 Then
 
For Y = 1 To Val(l)
 Call cabeceraempaque
 
 Printer.FontName = "Courier New"
 Printer.FontSize = "10"
 Printer.Print "      ==============================================================================="
 Printer.Print "      Cantidad   Unidad   Nro.Bulto Articulo/Pieza                           "
 Printer.Print "      ==============================================================================="
 Set rs = New adodb.Recordset
 q = "Select * from vta_03 where [num_int] = " & Me.numint
 rs.Open q, cn1
 l = 0
 While Not rs.EOF
    d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    dt = Space$(9)
    cant = Space$(10)
    cp2 = Space$(5)
    tu = Space$(8)
    If rs("cantidad_original") > 0 Then
      RSet cant = Format$(rs("cantidad_original"), "######0.00")
      RSet cp2 = Format$(rs("bultos"), "00000")
      tu = Format$(rs("tunidad"), "@@@@@@@@!")
    End If
    Printer.Print Tab(4); cant; "    "; tu; "  "; cp2; Spc(4);
    Printer.FontBold = True
    Printer.Print d
    Printer.FontBold = False
    
    'imprime extra
     If Me.imprimedescextra = "S" Then
       Set rs1 = New adodb.Recordset
       q = "select * from vta_015 where [num_int] = " & Me.numint & " and [renglon] = " & rs("renglon")
       rs1.Open q, cn1
       If Not rs1.EOF And Not rs1.BOF Then
         Call lee_desc_extra(a, rs1("desc_ext"))
         For k = 0 To 4
           If a(k) <> "%%" Then
             d = Format$(Left$(a(k), 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
             Printer.Print Tab(37); d
             l = l + 1
           End If
         Next k
       
       End If
       Set rs1 = Nothing
     End If
    
    
    
    rs.MoveNext
    l = l + 1
    If l > Me.cant_lineas Then
       pieempaque (Y)
       Printer.NewPage
       cabeceraempaque
 
    End If
 Wend
 Set rs = Nothing
 
Printer.Print
Printer.FontSize = 12
Printer.FontBold = True
Printer.Print "Observaciones: ";
Printer.FontBold = False
Printer.FontSize = 10
Printer.Print Me.observaciones


For ó = l To Me.cant_lineas
  Printer.Print
Next ó


Call pieempaque(Y)

If Y < c Then
      Printer.NewPage
End If

Next Y
 
Printer.EndDoc
End If

End Sub
Sub pieempaque(ByVal p As Integer)
Printer.Print
Printer.Print
Printer.FontSize = 8
Printer.Print "___________________________________________________________________________________________________"
 
  Set cl_usuarios = New usuarios
   cl_usuarios.cargar (Me.idusuario)
   If cl_usuarios.idusuario > 0 Then
        Printer.Print "Emitida por: " & cl_usuarios.denominACION & "    Hoja Nro: " & p & "  Fecha Imp: " & Format$(Now, "dd/mm/yy") & "   Orden de Empaque nro: " & Format$(Me.numcomp, "00000000")
   Else
        Printer.Print
   End If
   Set cl_usuarios = Nothing
End Sub

Sub cabeceraempaque()
Printer.FontName = "Times New Roman"
 cc = Space$(5)
 cp = Space$(9)
 cant = Space$(8)
 pu = Space$(9)
 dtop = Space$(7)
 pp = Space$(10)
 im = Space$(10)
 s = Space$(10)
 s2 = Space$(10)
 dto = Space$(7)
 i1 = Space$(9)
 I2 = Space$(9)
 tot = Space$(10)
 RSet cc = Format$(Me.idcliente, "0000")
 n = Format$(Me.razonsocial, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 d = Format$(Me.direccion, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 l = Format$(Me.localidad, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 comp = "ORDEN DE EMPAQUE "
 
 Printer.FontSize = 14
 Printer.Print glo.nombrecli
 Printer.FontSize = 10
 Printer.Print glo.direccioncli & " - Rojas(B) - CP.2705 - TE:" & glo.TECLI
 Printer.Print "Cuit:" & glo.CUIT
 Printer.Print "______________________________________________________________________________________"
 Printer.Print
 Printer.Print
 Printer.FontSize = 16
 Printer.Print Spc(30); comp;
 Printer.FontSize = 12
 Printer.Print "(Documento Interno)"
 Printer.Print
 Printer.Print Spc(80); " Fecha.....: " & Format$(Me.fecha)
 Printer.Print Spc(80); " Nro.......: "; Me.letra; Spc(1); Format$(Me.sucursal, "0000"); "-"; Format$(Me.numcomp, "00000000")
 Printer.Print
 Printer.FontSize = 12
 Printer.FontBold = False
 Printer.FontUnderline = True
 Printer.Print "Cliente.....:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "   ["; cc; "]  "; n
 Printer.Print
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Direccion...:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "   " & direccion & "  - " & localidad
 Printer.Print
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print "Obra........:";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print "  " & di & "         ";
 Printer.FontSize = 12
 Printer.FontUnderline = True
 Printer.Print " ";
 Printer.FontBold = False
 Printer.FontUnderline = False
 Printer.Print nc
 Printer.Print

End Sub
Sub formato1A(ByVal c As Integer)
Set cl_cli = New Clientes
cl_cli.carga (Me.idcliente)


If cl_cli.id > 0 Then
 dd = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 1, 2)
 mm = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 4, 2)
 aa = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 7, 4)
 lugar = Left$(RTrim$(cl_cli.direccion) & " - (" & RTrim$(cl_cli.cp) & ")" & RTrim$(cl_cli.localidad), 120)

 For Y = 1 To c
   
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Times New Roman"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(98); dd; "      "; mm; "       "; aa
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 6
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 14
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(20); cl_cli.razonsocial
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(22); lugar
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(20); cl_cli.abreviatura_tipoiva; Tab(92); cl_cli.CUIT
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(45);
   If Me.ctacte <> "N" Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print "Cuenta Corriente"
   Else
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print "Contado"
   End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
   Set rs = New adodb.Recordset
   q = "Select * from vta_03 where [num_int] = " & Me.numint
   rs.Open q, cn1
   i = Space$(10)
   t = Space$(4)
   pu = Space$(10)
   l = 0
   While Not rs.EOF
    i = Space$(10)
    t = Space$(4)
    pu = Space$(10)
    d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    dt = Space$(9)
    If rs("cantidad") > 0 Then
      RSet pu = Format$(rs("pu"), "######0.00")
      RSet t = Format$(rs("tasaiva"), "#0.0")
      RSet i = Format$(rs("importe"), "######0.00")
      dt = "  (" & t & "%)"
    End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(9); d; Tab(61); pu; Tab(75); i; dt
    rs.MoveNext
    l = l + 1
   Wend
   Set rs = Nothing
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
   For X = 1 To 29 - l
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
     Printer.Print
   Next X
   xs = Space$(10)
   RSet xs = Format$(Me.subtotal, "######0.00")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(78); xs
   Set cl_usuarios = New usuarios
   cl_usuarios.cargar (Me.idusuario)
   If cl_usuarios.idusuario > 0 Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
        Printer.Print Tab(10); "Emitida por: " & cl_usuarios.denominACION
   Else
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
        Printer.Print
   End If
   
   Set cl_usuarios = Nothing
   RSet xs = Format$(Me.percib, "######0.00")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(63); "Perc.IBBA"; Tab(78); xs
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
   RSet xs = Format$(Me.iva, "######0.00")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(78); xs
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(67); Format$(Me.total, "######0.00")
   If Y < c Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.NewPage
   End If
  Next Y
 End If
 Set cl_cli = Nothing
 
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.EndDoc

End Sub

Sub formato6(ByVal c As Integer)
Set cl_cli = New Clientes
cl_cli.carga (Me.idcliente)

Set rsr = New adodb.Recordset
 q = "SELECT [tasa_financiera] from g0 where [sucursal] = 0"
 rsr.Open q, cn1
 cf = Format$(rsr("tasa_financiera"), "#0.00")
 Set rsr = Nothing
 

If cl_cli.id > 0 Then
 dd = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 1, 2)
 mm = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 4, 2)
 aa = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 7, 4)
 lugar = Left$(RTrim$(cl_cli.direccion) & "    -   " & RTrim$(cl_cli.cp) & "    " & RTrim$(cl_cli.localidad), 120)

 For Y = 1 To c
   
   Printer.FontName = "Times New Roman"
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.FontSize = 12
   Printer.Print
   Printer.Print
   Printer.Print
   Printer.Print
   Printer.Print Tab(65); Me.abreviatura
   Printer.FontSize = 6
   Printer.Print
   Printer.FontSize = 12
   Printer.Print Tab(80); dd; " / "; mm; " / "; aa
   Printer.Print
   Printer.Print Tab(88); Me.abreviatura & " " & Me.letra & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000")
   Printer.Print
   Printer.Print
   'Printer.Print
   Printer.FontSize = 14
   Printer.Print Tab(19); cl_cli.razonsocial; "  ("; cl_cli.id; ")"
   Printer.FontSize = 6
   Printer.Print
   Printer.FontSize = 10
   Printer.Print Tab(26); lugar
   Printer.FontSize = 10
   Printer.Print
   Printer.FontSize = 10
   Printer.Print Tab(20); cl_cli.abreviatura_tipoiva; Tab(92); cl_cli.CUIT
   Printer.Print
   Printer.Print
   Printer.Print Tab(20);
   If Me.moneda = "P" Then
     mda = "Pesos ($)"
   Else
     mda = "Dolares (U$s)"
   End If
   If Me.ctacte <> "N" Then
      Printer.Print "Cuenta Corriente en " & mda
   Else
      Printer.Print "Contado " & mda
   End If
   
   Printer.Print
   Printer.Print Tab(24); mda; Tab(58); Format$(Me.cotizaciondolar, "###0.000"); Tab(70); "Vto: "; Me.fecha_vto
   Printer.Print
   Printer.Print
   Printer.Print
   Printer.Print
   Printer.FontName = "Courier New"
   Printer.FontSize = 10
   Set rs = New adodb.Recordset
   q = "Select * from vta_03 where [num_int] = " & Me.numint
   rs.Open q, cn1
   i = Space$(10)
   t = Space$(4)
   pu = Space$(10)
   l = 0
   While Not rs.EOF
    i = Space$(10)
    t = Space$(4)
    pu = Space$(10)
    cant = Space$(8)
    d = Format$(Left$(rs("descripcion") & " ", 40), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    dt = Space$(9)
    u = Space$(5)
    If rs("cantidad") > 0 Then
      RSet pu = Format$(rs("pu"), "######0.00")
      RSet t = Format$(rs("tasaiva"), "#0.0")
      RSet i = Format$(rs("importe"), "######0.00")
      RSet cant = Format$(rs("cantidad"), "####0.00")
      dt = "  (" & t & "%)"
      u = Format$(rs("tunidad"), "@@@@@!")
    End If
    Printer.FontSize = 8
    Printer.Print Tab(10); Format$(rs("id_producto"), "00000"); "  ";
    Printer.FontBold = True
    Printer.Print d;
    Printer.FontBold = False
    Printer.Print "  "; cant; " "; u; "  "; pu; " "; dt;
    Printer.FontSize = 10
    Printer.Print "     "; i
    rs.MoveNext
    l = l + 1
   Wend
   Set rs = Nothing
   Printer.FontSize = 10
   For X = 1 To 18 - l
     Printer.Print
   Next X
   
   q = "select * from vta_08, vta_02 where [id_factura] = " & Me.numint & " and [id_remito] = [num_int]"
   Set rs = New adodb.Recordset
   rs.Open q, cn1
   tr = "Remitos: "
   While Not rs.EOF
     nc = Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
     tr = tr & nc & " "
     rs.MoveNext
   Wend
   Set rs = Nothing
   Printer.Print Tab(8); tr
   Printer.Print
   Printer.Print
   Printer.FontSize = 8
   If Me.moneda = "P" Then
     If Me.letra = "A" Then
      sonpesos = convierte(Me.total)
      Printer.Print Tab(20); "Son Pesos:"; Left$(Format$(sonpesos, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!"), 69); Spc(2);
      Printer.FontSize = 11
      Printer.Print Format$(Me.subtotal, "#####0.00")
      Printer.FontSize = 8
      d = asteriscos(Format$(Me.total / Me.cotizaciondolar, "0.00"), 10)
      Printer.Print Tab(20); "El monto de la presente factura equivale a U$S "; Format$(d, "@@@@@@@@@@")
      Printer.Print Tab(20); "debiendose cancelar al vencimiento estipulado, en Pesos   "
      Printer.Print Tab(20); "segun Dolar Transferencia Tipo Vendedor del dia anterior "; Spc(29);
      Printer.Print Tab(20); "al efectivo Pago."
      Printer.Print Tab(20); "Cambio al dia de Facturacion a Pesos "; Format$(Me.cotizaciondolar, "##0.000"); " por Dolar."
      d = asteriscos(Format$(Me.total, "0.00"), 10)
      Printer.Print Tab(20); "La mora por atraso en la cancelación será del 2% mensual, "
      Printer.Print Tab(20); "produciéndose la misma en forma automática.-"
      Printer.Print
      Printer.Print
      Printer.Print Spc(94);
      Printer.FontSize = 11
      Printer.Print Format$(Format$(Me.iva, "######0.00"), "@@@@@@@@@@")
      Printer.FontSize = 8
      Printer.Print
      Printer.Print
      Printer.Print Spc(98);
      Printer.FontSize = 11
      Printer.FontBold = True
      Printer.Print "$"; Format$(Me.total, "#####0.00")
      Printer.FontBold = False
      Printer.FontSize = 9
      Printer.Print Tab(8); "Observaciones: " & Me.observaciones

      
  Else
      sonpesos = convierte(Me.total)
      Printer.Print Tab(20); "Son Pesos:"; Left$(Format$(sonpesos, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!"), 69); Spc(2);
      Printer.FontSize = 11
      Printer.Print " " 'Format$(Me.subtotal, "#####0.00")
      Printer.FontSize = 8
      d = asteriscos(Format$(Me.total / Me.cotizaciondolar, "0.00"), 10)
      Printer.Print Tab(20); "El monto de la presente factura equivale a U$S "; Format$(d, "@@@@@@@@@@")
      Printer.Print Tab(20); "debiendose cancelar al vencimiento estipulado, en Pesos   "
      Printer.Print Tab(20); "segun Dolar Transferencia Tipo Vendedor del dia anterior "; Spc(29);
      Printer.Print Tab(20); "al efectivo Pago."
      Printer.Print Tab(20); "Cambio al dia de Facturacion a Pesos "; Format$(Me.cotizaciondolar, "##0.000"); " por Dolar."
      d = asteriscos(Format$(Me.total, "0.00"), 10)
      Printer.Print Tab(20); "La mora por atraso en la cancelación será del 2% mensual, "
      Printer.Print Tab(20); "produciéndose la misma en forma automática.-"
      Printer.Print Tab(20); "Orientacion al Consumidor Prov. BA 08002229042"
      Printer.Print
      Printer.Print Spc(94);
      Printer.FontSize = 11
      Printer.Print " " 'Format$(Format$(Me.iva, "######0.00"), "@@@@@@@@@@")
      Printer.FontSize = 8
      Printer.Print
      Printer.Print
      Printer.Print Spc(98);
      Printer.FontSize = 11
      Printer.FontBold = True
      Printer.Print "$"; Format$(Me.total, "#####0.00")
      Printer.FontBold = False
      Printer.FontSize = 9
      Printer.Print Tab(8); "Observaciones: " & Me.observaciones

  End If


Else
  If Me.letra = "A" Then
      sondolares = convierte(Me.total)
      Printer.Print Tab(20); "Son U$s:  "; Left$(Format$(sondolares, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!"), 69); Spc(1);
      Printer.FontSize = 11
      Printer.Print Format$(Me.subtotal, "######0.00")
      Printer.FontSize = 8
      d = asteriscos(Format$(Me.total / Me.cotizaciondolar, "0.00"), 10)
      Printer.Print Tab(20); "Importes expresados en Dolares Estadounidenses.          "
      Printer.Print Tab(20); "A fines Impositivos y para la cancelacion de esta factura el "
      Printer.Print Tab(20); "tipo de cambio  será el del día  de  acreditación del pago de"
      Printer.Print Tab(20); "acuerdo al tipo de cambio vendedor del  Bco. Nacion cotizador"
      Printer.Print Tab(20); "de Billetes.-"
      Printer.Print Tab(20); "Pasado el vencimiento indicado la mora será del "; Format$(cf, "#0.0"); "% mensual"
      Printer.Print Tab(20); "produciendóse la misma en forma automática.-"
      Printer.FontSize = 8
      d = asteriscos(Format$(Me.total, "0.00"), 10)
      ip = asteriscos(Format$(Me.iva * Me.cotizaciondolar, "0.00"), 8)
      Printer.Print Tab(20); "Tipo cambio Banco Nació Vendedor: "; Format$(Me.cotizaciondolar, "##0.000"); " Iva: $"; ip
      Printer.Print
      Printer.Print Spc(94);
      Printer.FontSize = 11
      Printer.Print Format$(Format$(Me.iva, "######0.00"), "@@@@@@@@@@")
      Printer.FontSize = 8
      Printer.Print
      Printer.Print
      Printer.Print Spc(94);
      Printer.FontSize = 11
      Printer.FontBold = True
      Printer.Print "u$s"; d
      Printer.FontBold = False
      Printer.FontSize = 9
      Printer.Print Tab(7); "Observaciones: " & Me.observaciones
   
   Else
      sondolares = convierte(Me.total)
      Printer.Print Tab(18); "Son U$s:  "; Left$(Format$(sondolares, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!"), 69); Spc(1);
      Printer.FontSize = 11
      Printer.Print " " 'Format$(Me.subtotal, "######0.00")
      Printer.FontSize = 8
      d = asteriscos(Format$(Me.total / Me.cotizaciondolar, "0.00"), 10)
      Printer.Print Tab(20); "Importes expresados en Dolares Estadounidenses.          "
      Printer.Print Tab(20); "A fines Impositivos y para la cancelacion de esta factura el "
      Printer.Print Tab(20); "tipo de cambio  será el del día  de  acreditación del pago de"
      Printer.Print Tab(20); "acuerdo al tipo de cambio vendedor del  Bco. Nacion cotizador"
      Printer.Print Tab(20); "de Billetes.-"
      Printer.Print Tab(20); "Pasado el vencimiento indicado la mora será del "; Format$(cf, "#0.0"); "% mensual"
      Printer.Print Tab(20); "produciendóse la misma en forma automática.-"
      Printer.Print Tab(20); "Tipo cambio Banco Nació Vendedor: "; Format$(Me.cotizaciondolar, "##0.000")
      Printer.Print Tab(20); "Orientacion al Consumidor Prov. BA 08002229042"
      Printer.Print Spc(94);
      Printer.FontSize = 11
      Printer.Print " " 'Format$(Format$(Me.iva, "######0.00"), "@@@@@@@@@@")
      Printer.FontSize = 8
      Printer.Print
      Printer.Print
      Printer.Print Spc(94);
      Printer.FontSize = 11
      Printer.FontBold = True
      Printer.Print "u$s"; d
      Printer.FontBold = False
      Printer.FontSize = 9
      Printer.Print Tab(8); "Observaciones: " & Me.observaciones

   End If
      
  End If

   
   
   
   
   
   
      If Y < c Then
      Printer.NewPage
   End If
  Next Y
 End If
 Set cl_cli = Nothing
 
 Printer.EndDoc

End Sub

Sub formato4A(ByVal c As Integer)
 dd = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 1, 2)
 mm = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 4, 2)
 aa = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 7, 4)
 'lugar = Left$(RTrim$(cl_cli.direccion) & " - (" & RTrim$(cl_cli.cp) & ")" & RTrim$(cl_cli.localidad), 120)

 If Me.idtipocomp = 1 Then
  tacha = "               "
 Else
  tacha = Me.abreviatura & "     ###############"
 End If

 For Y = 1 To c
  Printer.FontName = "Times New Roman"
  Printer.FontSize = 10
  Printer.FontBold = False
  For i = 1 To 7
    Printer.Print
  Next i
   Printer.FontName = "Times New Roman"
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.FontSize = 12
   Printer.Print Tab(68); tacha
   Printer.Print Tab(80); Spc(4); dd & "/" & mm & "/" & aa
   For i = 1 To 3
     Printer.Print
   Next i
 Printer.FontSize = 8
 Printer.Print
 Printer.Print
 Printer.FontName = "Times New Roman"
 Printer.FontSize = 14
 Printer.FontBold = True
 Printer.Print Tab(19); "["; Format$(Me.idcliente, "0000"); "] "; Me.razonsocial
 Printer.FontName = "Courier New"
 Printer.FontSize = 10
 Printer.FontBold = False
 Printer.Print
 Printer.Print Tab(20); Me.direccion; Tab(68); Me.CUIT
 Printer.FontSize = 8
 Printer.Print Tab(25); Me.localidad
 Printer.FontSize = 10
 Printer.Print
 Printer.Print Tab(20); Me.abreviatura_tipoiva
 Printer.FontName = "Courier New"
 Printer.FontSize = 9
 Printer.FontBold = False
 Printer.Print
 Printer.Print
 If Me.contado = "N" Then
    Printer.Print Tab(38); "Cuenta Corriente"
 Else
    Printer.Print Tab(38); "Contado"
 End If
 Printer.Print
 Printer.Print
 Printer.Print
 Printer.Print
  Set rs = New adodb.Recordset
   q = "Select * from vta_03 where [num_int] = " & Me.numint
   rs.Open q, cn1
   cant = Space$(10)
   l = 0
   While Not rs.EOF
    i = Space$(10)
    t = Space$(4)
    pu = Space$(10)
    d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    dt = Space$(9)
    cant = Space$(10)
    cp2 = Space$(9)
   If rs("cantidad") > 0 Then
      If Me.letra = "A" Then
           RSet pu = Format$(rs("pu"), "######0.00")
           RSet i = Format$(rs("importe"), "######0.00")
      Else
           RSet pu = Format$(rs("pu_final"), "######0.00")
           RSet i = Format$(rs("pu_final") * rs("cantidad"), "######0.00")
      End If
      RSet t = Format$(rs("tasaiva"), "#0.0")
      RSet cant = Format$(rs("cantidad"), "######0.00")
      dt = "  (" & t & "%)"
      cp2 = " [" & Format$(rs("id_producto"), "00000") & "] "
    End If
    Printer.Print Tab(8); cant; Spc(4);
    Printer.FontBold = True
    Printer.Print d; ; Spc(3);
    Printer.FontBold = False
    Printer.Print pu;
    Printer.FontSize = 6
    Printer.Print dt;
    Printer.FontSize = 9
    Printer.Print Spc(1); i
 
    rs.MoveNext
    l = l + 1
   Wend
   Set rs = Nothing
  
   For X = l To 23
     Printer.Print
   Next X
   xs = Space$(10)
   RSet xs = Format$(Me.subtotal, "######0.00")
   If Me.letra = "A" Then
     Printer.Print Tab(93); xs
   Else
     Printer.Print
   End If
   Set cl_usuarios = New usuarios
   cl_usuarios.cargar (Me.idusuario)
   If cl_usuarios.idusuario > 0 Then
        Printer.Print Tab(10); "Emitida por: " & cl_usuarios.denominACION
   Else
        Printer.Print
   End If
   
   Set cl_usuarios = Nothing
   'RSet xs = Format$(Me.percib, "######0.00")
   Printer.Print 'Tab(63); "Perc.IBBA"; Tab(78); xs
   Printer.Print Tab(10); "Son Pesos: " & convierte(Me.total)
   Printer.Print
   Printer.Print
   RSet xs = Format$(Me.iva, "######0.00")
   If Me.letra = "A" Then
     Printer.Print Tab(93); xs
   Else
     Printer.Print
   End If
   Printer.Print
   Printer.Print
   Printer.Print
   Printer.Print
   Printer.Print
   Printer.FontSize = 12
   Printer.FontBold = True
   Printer.Print Tab(73); Format$(Me.total, "######0.00")
   If Y < c Then
      Printer.NewPage
   End If
  Next Y
 
 Printer.EndDoc

End Sub

Sub formatoE(ByVal c As Integer)
 Load gen_logo
'recuadro hoja

h = InputBox("Indique cantidad de Copias", , c)
If Val(h) > 0 And Val(h) <= 6 Then
 
 
 
 
 
 For Y = 1 To h
Printer.ScaleMode = vbCentimeters
Printer.Line (1.5, 1)-(20, 20), vbRed, B 'recuadro hoja
Printer.Line (1.5, 5)-(20, 5), vbRed     'encabezado
Printer.Line (10, 1.2)-(11.5, 3), vbRed, B 'letra
'Printer.Line (1.5, 7)-(20, 7), vbRed     'datos cliente
Printer.Line (1.52, 7)-(19.98, 7.6), RGB(233, 236, 232), BF 'cabecera productos
Printer.Line (1.5, 7)-(20, 7), vbBlack
Printer.Line (1.5, 7.6)-(20, 7.6), vbBlack
Printer.Line (1.5, 20.2)-(20, 24), vbRed, B 'recuadro totales
Printer.Line (1.54, 20.3)-(7, 20.7), RGB(233, 236, 232), BF 'cabecera iva
Printer.Line (13, 23.2)-(19, 23.9), RGB(233, 236, 232), BF  'recuadro total


'letra factura
Printer.CurrentX = 10.5
Printer.CurrentY = 1.5
Printer.FontName = "Arial"
Printer.FontSize = 21
Printer.FontBold = False
Printer.Print Me.letra

If Me.letra = "A" Then
 codcomp = Me.cod_afip_a
Else
 codcomp = Me.cod_afip_b
End If
 
'codigo factura
Printer.CurrentX = 10.3
Printer.CurrentY = 2.5
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.Print "Cod." & Format$(codcomp, "00")

'fantasia
Printer.CurrentX = 2
Printer.CurrentY = 2
Printer.FontName = "Arial"
Printer.FontSize = 21
Printer.Print para.nombre_fantasia

'Datos Empresa
Printer.CurrentX = 2
Printer.CurrentY = 3.5
Printer.FontName = "Sans Serif"
Printer.FontSize = 7
Printer.Print glo.nombrecli
Printer.CurrentX = 2
Printer.CurrentY = 4
Printer.Print glo.direccioncli
Printer.CurrentX = 2
Printer.CurrentY = 4.5
Printer.Print glo.TECLI

Printer.CurrentX = 12
Printer.CurrentY = 3.5
Printer.Print "CUIT:" & glo.CUIT
Printer.CurrentX = 12
Printer.CurrentY = 4
Printer.Print "Ingresos Brutos:" & para.numero_ib
Printer.CurrentX = 12
Printer.CurrentY = 4.5
Printer.Print "Inicio de Actividades: " & para.fecha_inicio_actividades

'Numero Factura
Printer.CurrentX = 12
Printer.CurrentY = 1.5
Printer.FontSize = 10
Printer.FontName = "Arial"
Printer.FontBold = True
Printer.Print Me.abreviatura & " Nro: " & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000")

'fecha
Printer.Line (12, 2.3)-(17, 3), vbRed, B
Printer.CurrentX = 12.5
Printer.CurrentY = 2.5
Printer.FontSize = 10
Printer.Print "   Fecha: " & Me.fecha

'cliente
Printer.CurrentX = 2
Printer.CurrentY = 5.5
Printer.FontSize = 12
Printer.Print Me.razonsocial
Printer.CurrentX = 2
Printer.CurrentY = 6
Printer.FontSize = 8
Printer.FontBold = False
Printer.Print Me.direccion
Printer.CurrentX = 2
Printer.CurrentY = 6.5
Printer.Print Me.localidad
Printer.CurrentX = 12
Printer.CurrentY = 5.5
If Me.contado = "S" Then
  Printer.Print "Condicion de Venta: CONTADO"
Else
   Printer.Print "Condicion de Venta: CUENTA CORRIENTE"
End If
Printer.CurrentX = 12
Printer.CurrentY = 6
Printer.Print "CUIT: " & Me.CUIT & "   -   " & Me.abreviatura_tipoiva

Select Case Me.tipoop
 Case Is = 1
    ttipoop = "Productos"
 Case Is = 2
    ttipoop = "Servicios"
 Case Is = 3
    ttipoop = "Servicios y Productos"
 Case Else
    ttipoop = "Productos"
End Select

Printer.CurrentX = 12
Printer.CurrentY = 6.5
Printer.Print "Tipo Op.: " & ttipoop

'cabecera productos
Printer.FontName = "Arial"
Printer.CurrentX = 1.7
Printer.CurrentY = 7.2
Printer.FontSize = 6
Printer.FontBold = True
Printer.Print "Codigo"
Printer.CurrentX = 3
Printer.CurrentY = 7.2
Printer.Print "Producto"
Printer.CurrentX = 13
Printer.CurrentY = 7.2
Printer.Print "Cant./Unid"
Printer.CurrentX = 15.7
Printer.CurrentY = 7.2
Printer.Print "P.U"
Printer.CurrentX = 17
Printer.CurrentY = 7.2
Printer.Print "%Iva"
Printer.CurrentX = 18.5
Printer.CurrentY = 7.2
Printer.Print "Importe"
Printer.FontBold = False

'productos
ly = 7.8
Printer.FontName = "Sans Serif"
Printer.FontBold = False
Printer.FontSize = 6

Set rs = New adodb.Recordset
q = "Select * from vta_03 where [num_int] = " & Me.numint
rs.Open q, cn1
cant = Space$(10)
l = 0
While Not rs.EOF
    i = Space$(10)
    t = Space$(4)
    pu = Space$(10)
    d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    dt = Space$(9)
    cant = Space$(8)
    cp2 = Space$(9)
    If rs("cantidad") > 0 Then
      'If Me.letra = "A" Then
           RSet pu = Format$(rs("pu"), "######0.00")
           RSet i = Format$(rs("importe"), "######0.00")
      'Else
      '     RSet pu = Format$(rs("pu_final"), "######0.00")
      '     RSet i = Format$(rs("pu_final") * rs("cantidad"), "######0.00")
      'End If
      RSet t = Format$(rs("tasaiva"), "#0.0")
      RSet cant = Format$(rs("cantidad"), "####0.00")
      dt = "  (" & t & "%)"
      cp2 = " [" & Format$(rs("id_producto"), "00000") & "] "
    End If
    
    Printer.CurrentX = 1.7
    Printer.CurrentY = ly
    Printer.Print cp2
    Printer.CurrentX = 2.7
    Printer.CurrentY = ly
    Printer.Print d
    Printer.CurrentX = 13
    Printer.CurrentY = ly
    Printer.Print cant & " " & Left$(rs("tunidad"), 5)
    Printer.CurrentX = 15.5
    Printer.CurrentY = ly
    Printer.Print pu
    Printer.CurrentX = 17
    Printer.CurrentY = ly
    Printer.Print dt
    Printer.CurrentX = 18.5
    Printer.CurrentY = ly
    Printer.Print i
    ly = ly + 0.35
    l = l + 1
    rs.MoveNext
 Wend
 Set rs = Nothing


'Cabecera ivas
Printer.CurrentX = 1.7
Printer.CurrentY = 20.3
Printer.Print "Impuesto"
Printer.CurrentX = 3
Printer.CurrentY = 20.3
Printer.Print "Alicuota"
Printer.CurrentX = 4
Printer.CurrentY = 20.3
Printer.Print "Base Imp."
Printer.CurrentX = 5.5
Printer.CurrentY = 20.3
Printer.Print "Importe"

'ivas
Set rs = New adodb.Recordset
q = "Select * from vta_09 where [num_int] = " & Me.numint
rs.Open q, cn1
i = Space$(10)
t = Space$(4)
bi = Space$(10)
dt = Space$(9)
ly = 20.8
While Not rs.EOF
      RSet t = Format$(rs("tasa_iva"), "#0.0")
      dt = "  (" & t & "%)"
      RSet bi = Format$(rs("neto"), "######0.00")
      RSet i = Format$(rs("iva"), "######0.00")
      Printer.CurrentX = 1.7
      Printer.CurrentY = ly
      Printer.Print "Iva"
      Printer.CurrentX = 3
      Printer.CurrentY = ly
      Printer.Print dt
      Printer.CurrentX = 4
      Printer.CurrentY = ly
      Printer.Print bi
      Printer.CurrentX = 5.5
      Printer.CurrentY = ly
      Printer.Print i
      rs.MoveNext
      ly = ly + 0.35
Wend

'poner no grabado y percepciones


If Me.impuestos > 0 Then
     RSet i = Format$(Me.impuestos, "######0.00")
     Printer.CurrentX = 1.7
     Printer.CurrentY = ly
     Printer.Print "No grabado"
     Printer.CurrentX = 5.5
     Printer.CurrentY = ly
     Printer.Print i
     ly = ly + 0.35
End If


If Me.perciva > 0 Then
      RSet i = Format$(Me.perciva, "######0.00")
      Printer.CurrentX = 1.7
      Printer.CurrentY = ly
      Printer.Print "Percep.Iva"
      Printer.CurrentX = 3
      Printer.CurrentY = ly
      Printer.CurrentX = 5.5
      Printer.CurrentY = ly
      Printer.Print i
      ly = ly + 0.35
End If

If Me.percib > 0 Then
      RSet i = Format$(Me.percib, "######0.00")
      Printer.CurrentX = 1.7
      Printer.CurrentY = ly
      Printer.Print "Percep.IB"
      Printer.CurrentX = 3
      Printer.CurrentY = ly
      Printer.CurrentX = 5.5
      Printer.CurrentY = ly
      Printer.Print i
      ly = ly + 0.35
End If


'totales
Printer.FontName = "Sans Serif"
Printer.FontBold = True
Printer.FontSize = 10
Printer.CurrentX = 13
Printer.CurrentY = 20.3
Printer.Print "Subtotal: "
Printer.CurrentX = 13
Printer.CurrentY = 21.3
Printer.Print "Iva:"
Printer.CurrentX = 13
Printer.CurrentY = 22.3
Printer.Print "Otros:"
Printer.CurrentX = 13
Printer.CurrentY = 23.3
Printer.Print "TOTAL:"

i = Space$(10)


Printer.FontName = "Sans Serif"
Printer.FontBold = False
Printer.FontSize = 10
Printer.CurrentX = 15
Printer.CurrentY = 20.3
RSet i = Format$(Me.subtotal, "######0.00")
Printer.Print i
Printer.CurrentX = 15
Printer.CurrentY = 21.3
RSet i = Format$(Me.iva, "######0.00")
Printer.Print i
Printer.CurrentX = 15
Printer.CurrentY = 22.3
RSet i = Format$(Me.percgan + Me.percib + Me.perciva + Me.impuestos, "######0.00")
Printer.Print i

'cbu
Printer.FontName = "Sans Serif"
Printer.FontBold = False
Printer.FontSize = 6
Printer.CurrentX = 1.7
Printer.CurrentY = 23.3
Printer.FontBold = True
Printer.Print "CBU para Transferencias:  " & sacacbu()

Printer.FontName = "Sans Serif"
Printer.FontBold = False
Printer.FontSize = 10
Printer.FontBold = True
Printer.CurrentX = 15
Printer.CurrentY = 23.3
Printer.FontBold = True
RSet i = Format$(Me.total, "######0.00")
Printer.Print i
Printer.FontBold = False
'recuadro total



'final cae y otors
Printer.FontName = "Sans Serif"
Printer.FontBold = False
Printer.FontSize = 10
Printer.CurrentX = 1.7
Printer.CurrentY = 24.3
Printer.Print "CAE Nro. " & Me.cae & "  Fecha de Vto: " & Me.fecha_vto_cae
'armar el codigo de barra
codbarra = Mid$(glo.CUIT, 1, 2) & Mid$(glo.CUIT, 4, 8) & Mid$(glo.CUIT, 13, 1) & Format$(codcomp, "00") & Format$(Me.sucursal, "0000") & Format$(Me.cae, "00000000000000") & Mid$(Format$(Me.fecha_vto_cae, "dd/mm/yyyy"), 7, 4) & Mid$(Format$(Me.fecha_vto_cae, "dd/mm/yyyy"), 4, 2) & Mid$(Format$(Me.fecha_vto_cae, "dd/mm/yyyy"), 1, 2)

'genero imagen codbarra
Dim PyI25 As Object
Set PyI25 = CreateObject("PyI25")
codbarra = codbarra & PyI25.DigitoVerificadorModulo10(codbarra)
' formato en jpg, aspecto 3x más ancho para imprimir o incrustar:
ok = PyI25.GenerarImagen(codbarra, "c:\5a04\barras.jpg", 9, 0, 90, "JPEG")
Set Py125 = Nothing
gen_logo.Picture3.Picture = LoadPicture("c:\5a04\barras.jpg")
Printer.PaintPicture gen_logo.Picture3.Picture, 1.7, 25, 10, 1.5
Printer.FontSize = 8
Printer.CurrentX = 3.41
Printer.CurrentY = 26.7
Printer.Print codbarra
gen_logo.Picture3.Picture = LoadPicture("c:\5a04\tools\afip.jpg")
Printer.PaintPicture gen_logo.Picture3.Picture, 12.1, 25, 2.1, 1






'mio
'codbarra = codbarra & digitoverificador(codbarra)
'Printer.FontName = "Interleaved 2of5 NT"
'Printer.FontBold = False
'Printer.FontSize = 9
'Printer.CurrentX = 1.7
'Printer.CurrentY = 25
'Printer.Print codbarra
'Printer.FontName = "Sans Serif"
'Printer.FontBold = False
'Printer.FontSize = 8
'Printer.CurrentX = 1.7
'Printer.CurrentY = 25.5
'Printer.Print codbarra

If Y < Val(h) Then
      Printer.NewPage
End If


Next Y
Printer.EndDoc
End If

 
 

End Sub

Sub imprimirrtoE(ByVal c As Integer)
Load gen_logo
'recuadro hoja
h = InputBox("Indique cantidad de Copias", , c)
If Val(h) > 0 And Val(h) <= 6 Then
 For Y = 1 To h

Printer.ScaleMode = vbCentimeters
'Printer.Line (1.5, 1)-(20, 20), vbRed, B 'recuadro hoja
Printer.Line (1.5, 5)-(20, 5), vbRed     'encabezado
'Printer.Line (1.5, 7)-(20, 7), vbRed     'datos cliente
''Printer.Line (1.5, 7)-(20, 7), vbBlack
''Printer.Line (1.5, 7.6)-(20, 7.6), vbBlack


'Printer.Line (10, 1.2)-(11.5, 3), vbRed, B 'letra
'letra factura
Printer.CurrentX = 10.5
Printer.CurrentY = 1.5
Printer.FontName = "Arial"
Printer.FontSize = 21
Printer.FontBold = False
Printer.Print Me.letra

codcomp = "91"
 
'codigo factura
Printer.CurrentX = 10.3
Printer.CurrentY = 2.5
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.Print "Cod." & Format$(codcomp, "00")

'fantasia
Printer.CurrentX = 2
Printer.CurrentY = 2
Printer.FontName = "Arial"
Printer.FontSize = 21
Printer.Print para.nombre_fantasia

'Datos Empresa
Printer.CurrentX = 2
Printer.CurrentY = 3.5
Printer.FontName = "Sans Serif"
Printer.FontSize = 7
Printer.Print glo.nombrecli
Printer.CurrentX = 2
Printer.CurrentY = 4
Printer.Print glo.direccioncli
Printer.CurrentX = 2
Printer.CurrentY = 4.5
Printer.Print glo.TECLI

Printer.CurrentX = 12
Printer.CurrentY = 3.5
Printer.Print "CUIT:" & glo.CUIT
Printer.CurrentX = 12
Printer.CurrentY = 4
Printer.Print "Ingresos Brutos:" & para.numero_ib
Printer.CurrentX = 12
Printer.CurrentY = 4.5
Printer.Print "Inicio de Actividades: " & para.fecha_inicio_actividades

'Numero Factura
Printer.CurrentX = 12
Printer.CurrentY = 1.5
Printer.FontSize = 10
Printer.FontName = "Arial"
Printer.FontBold = True
Printer.Print Me.abreviatura & " Nro: " & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000")

'fecha
'Printer.Line (12, 2.3)-(17, 3), vbRed, B
Printer.CurrentX = 12.5
Printer.CurrentY = 2.5
Printer.FontSize = 10
Printer.Print "   Fecha: " & Me.fecha

'cliente
Printer.CurrentX = 2
Printer.CurrentY = 5.5
Printer.FontSize = 12
Printer.Print Me.razonsocial
Printer.CurrentX = 2
Printer.CurrentY = 6
Printer.FontSize = 8
Printer.FontBold = False
Printer.Print Me.direccion
Printer.CurrentX = 2
Printer.CurrentY = 6.5
Printer.Print Me.localidad
Printer.CurrentX = 12
Printer.CurrentY = 5.5
If Me.contado = "S" Then
  Printer.Print "Condicion de Venta: CONTADO"
Else
   Printer.Print "Condicion de Venta: CUENTA CORRIENTE"
End If
Printer.CurrentX = 12
Printer.CurrentY = 6
Printer.Print "CUIT: " & Me.CUIT & "   -   " & Me.abreviatura_tipoiva


'cabecera productos
Printer.Line (1.52, 7)-(19.98, 7.6), RGB(233, 236, 232), BF 'cabecera productos
Printer.FontName = "Arial"
Printer.CurrentX = 1.7
Printer.CurrentY = 7.2
Printer.FontSize = 6
Printer.FontBold = True
Printer.Print "Codigo"
Printer.CurrentX = 3
Printer.CurrentY = 7.2
Printer.Print "Producto"
Printer.CurrentX = 13
Printer.CurrentY = 7.2
Printer.Print "Cantidad"
Printer.CurrentX = 15
Printer.CurrentY = 7.2
Printer.Print "Unidad"


Printer.FontBold = False

'productos
ly = 7.8
Printer.FontName = "Sans Serif"
Printer.FontBold = False
Printer.FontSize = 6

Set rs = New adodb.Recordset
q = "Select * from vta_03 where [num_int] = " & Me.numint
rs.Open q, cn1
cant = Space$(10)
l = 0
While Not rs.EOF
    i = Space$(10)
    t = Space$(4)
    pu = Space$(10)
    d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    dt = Space$(9)
    cant = Space$(8)
    cp2 = Space$(9)
    If rs("cantidad") > 0 Then
      RSet cant = Format$(rs("cantidad"), "####0.00")
      cp2 = " [" & Format$(rs("id_producto"), "00000") & "] "
    End If
    
    Printer.CurrentX = 1.7
    Printer.CurrentY = ly
    Printer.Print cp2
    Printer.CurrentX = 2.7
    Printer.CurrentY = ly
    Printer.Print d
    Printer.CurrentX = 13
    Printer.CurrentY = ly
    Printer.Print cant
    Printer.CurrentX = 15
    Printer.CurrentY = ly
    Printer.Print rs("tunidad")
    l = l + 1
    ly = ly + 0.35
    rs.MoveNext
 Wend
 Set rs = Nothing

Printer.Line (1.5, 20.2)-(20, 20.4), vbRed, B 'recuadro totales


 Printer.CurrentX = 1.7
 Printer.CurrentY = 21
 Printer.FontSize = 8
 Printer.FontBold = True
 Printer.Print "Observaciones: ";
 Printer.FontBold = False
 Printer.Print Me.observaciones

 Printer.Line (15, 23)-(19, 23), vbBlack
 Printer.CurrentX = 16
 Printer.CurrentY = 23.2
 Printer.FontSize = 8
 Printer.FontBold = False
 Printer.Print "Recibí Conforme ";
 


'final cae y otors
'Printer.FontName = "Sans Serif"
'Printer.FontBold = False
'Printer.FontSize = 10
'Printer.CurrentX = 1.7
'Printer.CurrentY = 24.3
'Printer.Print "CAE Nro. " & Me.cae & "  Fecha de Vto: " & Me.fecha_vto_cae
'armar el codigo de barra
'codbarra = Mid$(glo.CUIT, 1, 2) & Mid$(glo.CUIT, 4, 8) & Mid$(glo.CUIT, 13, 1) & Format$(codcomp, "00") & Format$(Me.sucursal, "0000") & Format$(Me.cae, "00000000000000") & Mid$(Format$(Me.fecha_vto_cae, "dd/mm/yyyy"), 7, 4) & Mid$(Format$(Me.fecha_vto_cae, "dd/mm/yyyy"), 4, 2) & Mid$(Format$(Me.fecha_vto_cae, "dd/mm/yyyy"), 1, 2)
If Y < Val(h) Then
      Printer.NewPage
End If


Next Y
Printer.EndDoc
End If
End Sub

Function digitoverificador(ByVal cb As String) As Integer
'calcula el digitoverificador del codigo de barra a imprimir
Dim totalpares As Integer
Dim totalimpares As Integer
Dim X As Double
Dim digito As Integer

totalimpares = 0
For X = 1 To Len(cb) Step 2
    totalimpares = totalimpares + Val(Mid$(cb, X, 1))
Next X
totalimpares = totalimpares * 3


totalpares = 0
For X = 2 To Len(cb) Step 2
    totalpares = totalpares + Val(Mid$(cb, X, 1))
Next X

X = totalimpares + totalpares
digito = 0

Do While Int(X / 10) * 10 <> X
  X = X + 1
  digito = digito + 1
Loop

digitoverificador = digito



End Function


Sub formato5(ByVal c As Integer)
 Dim a(5) As String
 
 Set rsr = New adodb.Recordset
 q = "select * from vta_08, vta_02 where [id_factura] = " & Me.numint & " and [id_remito] = [num_int]"
 rsr.Open q, cn1
 tr = ""
 tr1 = ""
 tr2 = ""
 xr = 1
 While Not rsr.EOF
  If rsr("id_tipocomp") = 45 Then
    tc = "R"
  Else
    tc = "D"
  End If
  Select Case xr
  Case Is <= 4
    tr = tr & Format$(rsr("sucursal"), "0000") & "-" & Format$(rsr("num_comp"), "00000000") & tc & " * "
  Case Is > 4
    tr1 = tr1 & Format$(rsr("sucursal"), "0000") & "-" & Format$(rsr("num_comp"), "00000000") & tc & " * "
  End Select
  xr = xr + 1
  rsr.MoveNext
 Wend
 Set rsr = Nothing
 
 
 
 dd = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 1, 2)
 mm = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 4, 2)
 aa = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 7, 4)
 'lugar = Left$(RTrim$(cl_cli.direccion) & " - (" & RTrim$(cl_cli.cp) & ")" & RTrim$(cl_cli.localidad), 120)

 tacha = Me.abreviatura & "   "
 tachan = Me.letra & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000")
 For Y = 1 To c
  Printer.FontName = "Times New Roman"
  Printer.FontSize = 10
  Printer.FontBold = False
  For i = 1 To 6
    Printer.Print
  Next i
  Printer.FontSize = 12
  Printer.Print
  
  If Me.letra = "A" Then
   Printer.FontSize = 6
   Printer.Print
  End If
  Printer.FontSize = 10
   Printer.FontName = "Times New Roman"
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.FontSize = 12
   Printer.Print Tab(72); tacha;
   Printer.FontSize = 6
   Printer.Print tachan;
   Printer.FontSize = 12
   Printer.Print Tab(94); Spc(4); dd & "/" & mm & "/" & aa
   For i = 1 To 3
     Printer.Print
   Next i
 Printer.FontSize = 8
 Printer.Print
 Printer.Print
 Printer.FontName = "Times New Roman"
 Printer.FontSize = 14
 Printer.FontBold = True
 Printer.Print Tab(16); Me.razonsocial
 Printer.FontName = "Courier New"
 Printer.FontSize = 10
 Printer.FontBold = False
 Printer.Print Tab(65); "["; Format$(Me.idcliente, "0000"); "] ";
 Printer.Print
 Printer.Print Tab(16); Me.direccion
 Printer.Print Tab(16); Me.localidad; Tab(65); Me.CUIT
 Printer.Print
 Printer.Print Tab(16); "Condicion Iva: "; Me.abreviatura_tipoiva
 Printer.FontName = "Courier New"
 Printer.FontSize = 9
 Printer.FontBold = False
 Printer.Print
  If Me.contado = "N" Then
    Printer.Print Tab(34); "Cuenta Corriente"
 Else
    Printer.Print Tab(34); "Contado"
 End If
 Printer.Print
 Printer.Print
 Printer.Print
 Printer.Print
' Printer.Print
   Set rs = New adodb.Recordset
   q = "Select * from vta_03 where [num_int] = " & Me.numint
   rs.Open q, cn1
   cant = Space$(10)
   l = 0
   While Not rs.EOF
    i = Space$(10)
    t = Space$(4)
    pu = Space$(10)
    d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    dt = Space$(9)
    cant = Space$(8)
    cp2 = Space$(9)
   If rs("cantidad") > 0 Then
      If Me.letra = "A" Then
           RSet pu = Format$(rs("pu"), "######0.00")
           RSet i = Format$(rs("importe"), "######0.00")
      Else
           RSet pu = Format$(rs("pu_final"), "######0.00")
           RSet i = Format$(rs("pu_final") * rs("cantidad"), "######0.00")
      End If
      RSet t = Format$(rs("tasaiva"), "#0.0")
      RSet cant = Format$(rs("cantidad"), "####0.00")
      dt = "  (" & t & "%)"
      cp2 = " [" & Format$(rs("id_producto"), "00000") & "] "
    End If
    
    Printer.Print Tab(6); cant; Spc(1); cp2; Spc(1);
    Printer.FontBold = True
    Printer.Print d; Spc(1);
    Printer.FontBold = False
    Printer.Print pu;
    Printer.FontSize = 6
    Printer.Print dt;
    Printer.FontSize = 9
    Printer.Print Spc(1); i
 
    'imprime extra
     If Me.imprimedescextra = "S" Then
       Set rs1 = New adodb.Recordset
       q = "select * from vta_015 where [num_int] = " & Me.numint & " and [renglon] = " & rs("renglon")
       rs1.Open q, cn1
       If Not rs1.EOF And Not rs1.BOF Then
         Call lee_desc_extra(a, rs1("desc_ext"))
         For k = 0 To 4
           If a(k) <> "%%" Then
             d = Format$(Left$(a(k), 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
             Printer.Print Tab(25); d
             l = l + 1
           End If
         Next k
       
       End If
       Set rs1 = Nothing
     End If
 
 
 
    rs.MoveNext
    l = l + 1
   Wend
   Set rs = Nothing
  
   For X = l To Me.cant_lineas
     Printer.Print
   Next X
   
  ' If Me.letra = "B" Then
  '   For i = 1 To 3
  '    Printer.Print
  '   Next i
  ' Else
  ' End If
   Set cl_usuarios = New usuarios
   cl_usuarios.cargar (Me.idusuario)
   If cl_usuarios.idusuario > 0 Then
        Printer.Print Tab(10); "Emitida por  : " & cl_usuarios.denominACION;
   Else
        'Printer.Print
   End If
   xs = Space$(10)
   RSet xs = Format$(Me.subtotal, "######0.00")
   If Me.letra = "A" Then
     Printer.Print Tab(93); xs
   Else
     Printer.Print
   End If

   Set cl_usuarios = Nothing
   'RSet xs = Format$(Me.percib, "######0.00")
   Printer.Print Tab(10); "Remitos      : "; tr
   If tr1 <> "" Then
     Printer.Print Tab(10); "               "; tr1
   Else
     Printer.Print
   End If
   Printer.Print Tab(10); "Vencimiento  : " & Me.fecha_vto
   Printer.Print Tab(10); "Observaciones: " & Me.observaciones
   RSet xs = Format$(Me.iva, "######0.00")
   If Me.letra = "A" Then
     Printer.Print Tab(93); xs
   'Else
   '  Printer.Print
   End If
   Printer.Print
   If Me.letra = "B" Then
     Printer.Print Tab(10); "Orientacion al Consumidor Prov. BA 08002229042"
      ' Printer.Print
      ' Printer.Print
   Else
     'Printer.Print
   End If
   
   Printer.FontSize = 12
   Printer.FontBold = True
   Printer.Print Tab(70); Format$(Me.total, "######0.00")
   If Y < c Then
      Printer.NewPage
   End If
  Next Y
 
 Printer.EndDoc

End Sub

Sub formato_A(ByVal c As Integer)
 Dim a(5) As String
 
 Set rsr = New adodb.Recordset
 q = "select * from vta_08, vta_02 where [id_factura] = " & Me.numint & " and [id_remito] = [num_int]"
 rsr.Open q, cn1
 tr = ""
 tr1 = ""
 tr2 = ""
 xr = 1
 While Not rsr.EOF
  If rsr("id_tipocomp") = 45 Then
    tc = "R"
  Else
    tc = "D"
  End If
  Select Case xr
  Case Is <= 8
    tr = tr & Format$(rsr("sucursal"), "0000") & "-" & Format$(rsr("num_comp"), "00000000") & tc & " * "
  Case Is > 8
    tr1 = tr1 & Format$(rsr("sucursal"), "0000") & "-" & Format$(rsr("num_comp"), "00000000") & tc & " * "
  End Select
  xr = xr + 1
  rsr.MoveNext
 Wend
 Set rsr = Nothing
 
 
 
 dd = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 1, 2)
 mm = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 4, 2)
 aa = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 7, 4)
 'lugar = Left$(RTrim$(cl_cli.direccion) & " - (" & RTrim$(cl_cli.cp) & ")" & RTrim$(cl_cli.localidad), 120)

 tacha = Me.abreviatura & "   "
 tachan = Me.letra & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000")
 For Y = 1 To c
  Printer.FontName = "Times New Roman"
  Printer.FontSize = 10
  Printer.FontBold = False
  For i = 1 To 1
    Printer.Print
  Next i
  Printer.FontSize = 12
  Printer.Print
  
  'If Me.letra = "A" Then
   Printer.FontSize = 6
   Printer.Print
   Printer.Print
  'End If
  Printer.FontSize = 10
   Printer.FontName = "Times New Roman"
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.FontSize = 12
   Printer.Print Tab(97); tacha;
   Printer.FontSize = 6
   Printer.Print tachan;
   Printer.FontSize = 6
   Printer.Print
   Printer.FontSize = 12
   Printer.Print
   Printer.Print
   If Me.idtipocomp <> 1 Then
     Printer.FontSize = 16
     Printer.Print Tab(51); "XXXXXXX    "; Me.abreviatura;
     Printer.FontSize = 12
   End If
   Printer.Print Tab(96); Spc(4); dd & "/" & mm & "/" & aa
   For i = 1 To 2
     Printer.Print
   Next i
 Printer.FontSize = 8
 Printer.Print
 Printer.Print
 Printer.Print
 Printer.FontName = "Times New Roman"
 Printer.FontSize = 14
 Printer.FontBold = True
 Printer.Print Tab(13); Me.razonsocial
 Printer.FontName = "Courier New"
 Printer.FontSize = 10
 Printer.FontBold = False
 Printer.Print Tab(13); "["; Format$(Me.idcliente, "0000"); "] ";
 Printer.Print Tab(57); Me.direccion
 Printer.Print Tab(57); Me.localidad
 Printer.Print Tab(13); Me.abreviatura_tipoiva; Tab(70); Me.CUIT
 Printer.FontName = "Courier New"
 Printer.FontSize = 9
 Printer.FontBold = False
 Printer.Print
  If Me.contado = "N" Then
    Printer.Print Tab(17); "Cuenta Corriente"
 Else
    Printer.Print Tab(17); "Contado"
 End If
 Printer.Print
 Printer.Print
 Printer.Print
 Printer.Print
' Printer.Print
   Set rs = New adodb.Recordset
   q = "Select * from vta_03 where [num_int] = " & Me.numint & " order by [renglon]"
   rs.Open q, cn1
   cant = Space$(10)
   l = 0
   While Not rs.EOF
    i = Space$(10)
    t = Space$(4)
    pu = Space$(10)
    d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    dt = Space$(9)
    cant = Space$(8)
    cp2 = Space$(9)
   If rs("cantidad") > 0 Then
      If Me.letra = "A" Then
           RSet pu = Format$(rs("pu"), "######0.00")
           RSet i = Format$(rs("importe"), "######0.00")
      Else
           RSet pu = Format$(rs("pu_final"), "######0.00")
           RSet i = Format$(rs("pu_final") * rs("cantidad"), "######0.00")
      End If
      RSet t = Format$(rs("tasaiva"), "#0.0")
      RSet cant = Format$(rs("cantidad"), "####0.00")
      dt = "  (" & t & "%)"
      cp2 = " [" & Format$(rs("id_producto"), "00000") & "] "
    End If
    
    Printer.Print Tab(3); cant; Spc(1); cp2; Spc(1);
    Printer.FontBold = True
    Printer.Print d; Spc(1);
    Printer.FontBold = False
    Printer.Print pu;
    Printer.FontSize = 6
    Printer.Print dt;
    Printer.FontSize = 9
    Printer.Print Spc(1); i
 
    'imprime extra
     If Me.imprimedescextra = "S" Then
       Set rs1 = New adodb.Recordset
       q = "select * from vta_015 where [num_int] = " & Me.numint & " and [renglon] = " & rs("renglon")
       rs1.Open q, cn1
       If Not rs1.EOF And Not rs1.BOF Then
         Call lee_desc_extra(a, rs1("desc_ext"))
         For k = 0 To 4
           If a(k) <> "%%" Then
             d = Format$(Left$(a(k), 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
             Printer.Print Tab(25); d
             l = l + 1
           End If
         Next k
       
       End If
       Set rs1 = Nothing
     End If
 
 
 
    rs.MoveNext
    l = l + 1
   Wend
   Set rs = Nothing
  
   For X = l To Me.cant_lineas
     Printer.Print
   Next X
  'Ultimas lineas
  Printer.Print "----------------------------------------------------------------"
  Printer.Print Tab(5); "Vencimiento  : " & Me.fecha_vto
  Printer.Print Tab(5); "Observaciones: " & Me.observaciones
  Printer.Print Tab(5); "Remitos      : "; tr
  If tr1 <> "" Then
     Printer.Print Tab(5); "               "; tr1
  Else
     Printer.Print
  End If
  
  If Me.letra = "B" Then
     Printer.Print Tab(5); "Orientacion al Consumidor Prov. BA 08002229042"
      ' Printer.Print
      ' Printer.Print
   Else
     Printer.Print
   End If
  
  Printer.Print
  Printer.Print
  Printer.Print
  Printer.Print
  Printer.Print
   xs = Space$(10)
   RSet xs = Format$(Me.subtotal, "######0.00")
   Printer.FontSize = 12
   If Me.letra = "A" Then
     Printer.Print Tab(5); xs;
     Printer.Print Tab(34); xs;
   End If

   RSet xs = Format$(Me.iva, "######0.00")
   If Me.letra = "A" Then
     Printer.Print Tab(51); xs;
   End If
  
      
   
   Printer.FontBold = True
   
   Printer.Print Tab(66); asteriscos(Format$(Me.total, "0.00"), 10)
   Printer.FontBold = False
   If Y < c Then
      Printer.NewPage
   End If
  Next Y
 
 Printer.EndDoc

End Sub

Sub imprimirpresup_A(ByVal c As Integer)

 Dim a(5) As String
 

 
 
 dd = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 1, 2)
 mm = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 4, 2)
 aa = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 7, 4)
 'lugar = Left$(RTrim$(cl_cli.direccion) & " - (" & RTrim$(cl_cli.cp) & ")" & RTrim$(cl_cli.localidad), 120)

 tacha = Me.abreviatura & "   "
 tachan = Me.letra & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000")
 For Y = 1 To c
  Printer.FontName = "Times New Roman"
  Printer.FontSize = 10
  Printer.FontBold = False
  For i = 1 To 1
    Printer.Print
  Next i
  Printer.FontSize = 12
  Printer.Print
  
  'If Me.letra = "A" Then
   Printer.FontSize = 6
   Printer.Print
   Printer.Print
  'End If
  Printer.FontSize = 10
   Printer.FontName = "Times New Roman"
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.FontSize = 12
   Printer.Print Tab(97); tacha;
   Printer.FontSize = 6
   Printer.Print tachan;
   Printer.FontSize = 6
   Printer.Print
   Printer.FontSize = 12
   Printer.Print
   Printer.Print
      Printer.Print Tab(96); Spc(4); dd & "/" & mm & "/" & aa
   For i = 1 To 2
     Printer.Print
   Next i
 Printer.FontSize = 8
 Printer.Print
 Printer.Print
 Printer.Print
 Printer.FontName = "Times New Roman"
 Printer.FontSize = 14
 Printer.FontBold = True
 Printer.Print Tab(13); Me.razonsocial
 Printer.FontName = "Courier New"
 Printer.FontSize = 10
 Printer.FontBold = False
 Printer.Print Tab(13); "["; Format$(Me.idcliente, "0000"); "] ";
 Printer.Print Tab(57); Me.direccion
 Printer.Print Tab(57); Me.localidad
 Printer.Print Tab(13); Me.abreviatura_tipoiva; Tab(70); Me.CUIT
 Printer.FontName = "Courier New"
 Printer.FontSize = 9
 Printer.FontBold = False
 Printer.Print
  If Me.contado = "N" Then
    Printer.Print Tab(17); "Cuenta Corriente"
 Else
    Printer.Print Tab(17); "Contado"
 End If
 Printer.Print
 Printer.Print
 Printer.Print
 Printer.Print
' Printer.Print
   Set rs = New adodb.Recordset
   q = "Select * from vta_03 where [num_int] = " & Me.numint & " order by [renglon]"
   rs.Open q, cn1
   cant = Space$(10)
   l = 0
   While Not rs.EOF
    i = Space$(10)
    t = Space$(4)
    pu = Space$(10)
    d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    dt = Space$(9)
    cant = Space$(8)
    cp2 = Space$(9)
   If rs("cantidad") > 0 Then
           RSet pu = Format$(rs("pu_final"), "######0.00")
           RSet i = Format$(rs("pu_final") * rs("cantidad"), "######0.00")
      RSet t = Format$(rs("tasaiva"), "#0.0")
      RSet cant = Format$(rs("cantidad"), "####0.00")
      dt = "  (" & t & "%)"
      cp2 = " [" & Format$(rs("id_producto"), "00000") & "] "
    End If
    
    Printer.Print Tab(3); cant; Spc(1); cp2; Spc(1);
    Printer.FontBold = True
    Printer.Print d; Spc(1);
    Printer.FontBold = False
    Printer.Print pu;
    Printer.FontSize = 6
    Printer.Print dt;
    Printer.FontSize = 9
    Printer.Print Spc(1); i
 
    'imprime extra
     If Me.imprimedescextra = "S" Then
       Set rs1 = New adodb.Recordset
       q = "select * from vta_015 where [num_int] = " & Me.numint & " and [renglon] = " & rs("renglon")
       rs1.Open q, cn1
       If Not rs1.EOF And Not rs1.BOF Then
         Call lee_desc_extra(a, rs1("desc_ext"))
         For k = 0 To 4
           If a(k) <> "%%" Then
             d = Format$(Left$(a(k), 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
             Printer.Print Tab(25); d
             l = l + 1
           End If
         Next k
       
       End If
       Set rs1 = Nothing
     End If
 
 
 
    rs.MoveNext
    l = l + 1
   Wend
   Set rs = Nothing
  
   For X = l To Me.cant_lineas
     Printer.Print
   Next X
  'Ultimas lineas
  Printer.Print "----------------------------------------------------------------"
  Printer.Print Tab(5); "Vencimiento  : " & Me.fecha_vto
  Printer.Print Tab(5); "Observaciones: " & Me.observaciones
     Printer.Print
  
     Printer.Print
  
  Printer.Print
  Printer.Print
  Printer.Print
  Printer.Print
  Printer.Print
   xs = Space$(10)
  
      
   Printer.FontSize = 12
   Printer.FontBold = True
   
   Printer.Print Tab(66); asteriscos(Format$(Me.total, "0.00"), 10)
   Printer.FontBold = False
   If Y < c Then
      Printer.NewPage
   End If
  Next Y
 
 Printer.EndDoc

End Sub

Sub imprimirrto_A(ByVal c As Integer)
 Dim a(5) As String
 
 
 
 dd = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 1, 2)
 mm = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 4, 2)
 aa = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 7, 4)
 'lugar = Left$(RTrim$(cl_cli.direccion) & " - (" & RTrim$(cl_cli.cp) & ")" & RTrim$(cl_cli.localidad), 120)

 tacha = Me.abreviatura & "   "
 tachan = Me.letra & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000")
 For Y = 1 To c
  Printer.FontName = "Times New Roman"
  Printer.FontSize = 10
  Printer.FontBold = False
  For i = 1 To 1
    Printer.Print
  Next i
  Printer.FontSize = 14
  Printer.Print
  
  If Me.letra = "A" Then
   Printer.FontSize = 6
   Printer.Print
  End If
  Printer.FontSize = 10
   Printer.FontName = "Times New Roman"
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.FontSize = 12
   Printer.Print Tab(97); tacha;
   Printer.FontSize = 6
   Printer.Print tachan;
   Printer.FontSize = 12
   Printer.Print
   Printer.Print
   Printer.Print
   If Me.idtipocomp <> 45 Then
     Printer.FontSize = 16
     Printer.Print Tab(64); Me.abreviatura;
     Printer.FontSize = 12
   End If
   Printer.Print Tab(96); Spc(4); dd & "/" & mm & "/" & aa
   For i = 1 To 2
     Printer.Print
   Next i
 Printer.FontSize = 8
 Printer.Print
 Printer.Print
 Printer.Print
 Printer.FontName = "Times New Roman"
 Printer.FontSize = 14
 Printer.FontBold = True
 Printer.Print Tab(13); Me.razonsocial
 Printer.FontName = "Courier New"
 Printer.FontSize = 10
 Printer.FontBold = False
 Printer.Print Tab(13); "["; Format$(Me.idcliente, "0000"); "] ";
 Printer.Print Tab(57); Me.direccion
 Printer.Print Tab(57); Me.localidad
 Printer.Print Tab(13); Me.abreviatura_tipoiva; Tab(70); Me.CUIT
 Printer.FontName = "Courier New"
 Printer.FontSize = 9
 Printer.FontBold = False
 Printer.Print
  If Me.contado = "N" Then
    Printer.Print Tab(17); "Cuenta Corriente"
 Else
    Printer.Print Tab(17); "Contado"
 End If
 Printer.Print
 Printer.Print
 Printer.Print
 Printer.Print
' Printer.Print
   Set rs = New adodb.Recordset
   q = "Select * from vta_03 where [num_int] = " & Me.numint
   rs.Open q, cn1
   cant = Space$(10)
   l = 0
   While Not rs.EOF
    i = Space$(10)
    t = Space$(4)
    pu = Space$(10)
    d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    dt = Space$(9)
    cant = Space$(8)
    cp2 = Space$(9)
   If rs("cantidad") > 0 Then
      RSet cant = Format$(rs("cantidad"), "####0.00")
      dt = "  (" & t & "%)"
      cp2 = " [" & Format$(rs("id_producto"), "00000") & "] "
   End If
    
    Printer.Print Tab(3); cant; Spc(1); cp2; Spc(1);
    Printer.FontBold = True
    Printer.Print d; Spc(1);
    Printer.FontBold = False
    
    'imprime extra
     If Me.imprimedescextra = "S" Then
       Set rs1 = New adodb.Recordset
       q = "select * from vta_015 where [num_int] = " & Me.numint & " and [renglon] = " & rs("renglon")
       rs1.Open q, cn1
       If Not rs1.EOF And Not rs1.BOF Then
         Call lee_desc_extra(a, rs1("desc_ext"))
         For k = 0 To 4
           If a(k) <> "%%" Then
             d = Format$(Left$(a(k), 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
             Printer.Print Tab(25); d
             l = l + 1
           End If
         Next k
       
       End If
       Set rs1 = Nothing
     End If
 
 
 
    rs.MoveNext
    l = l + 1
   Wend
   Set rs = Nothing
   Printer.Print
   Printer.Print Tab(22); "**************************************************"
   l = l + 1
   
   For X = l To Me.cant_lineas
     Printer.Print
   Next X
  'Ultimas lineas
  Printer.Print Tab(5); "Firma: _______________________________________"
  
  
  For i = 1 To 6
    Printer.Print
  Next
     
  Printer.Print Tab(5); Me.observaciones
   Printer.FontBold = False
   If Y < c Then
      Printer.NewPage
   End If
  Next Y
 
 Printer.EndDoc

End Sub


Sub imprimirrto_B(ByVal c As Integer)
'Hugo diaz laser a4 preimpreso
 Dim a(5) As String
 
 dd = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 1, 2)
 mm = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 4, 2)
 aa = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 7, 4)
 'lugar = Left$(RTrim$(cl_cli.direccion) & " - (" & RTrim$(cl_cli.cp) & ")" & RTrim$(cl_cli.localidad), 120)

 tacha = Me.abreviatura & "   "
 tachan = Me.letra & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000")
 For Y = 1 To c
  Printer.FontName = "Times New Roman"
  Printer.FontSize = 10
  Printer.FontBold = False
  For i = 1 To 6
    Printer.Print
  Next i
  Printer.FontSize = 10
  Printer.Print
   Printer.FontSize = 12
   Printer.Print Tab(96); Spc(4); dd & "/" & mm & "/" & aa
   Printer.FontSize = 6
   Printer.Print
   Printer.Print Tab(200); tacha; tachan
   For i = 1 To 12
     Printer.Print
   Next i
 Printer.FontName = "Times New Roman"
 Printer.FontSize = 14
 Printer.FontBold = True
 Printer.Print Tab(13); Me.razonsocial;
 Printer.FontName = "Courier New"
 Printer.FontSize = 10
 Printer.FontBold = False
 Printer.Print Tab(71); "["; Format$(Me.idcliente, "0000"); "] "
 Printer.FontSize = 14
 Printer.Print
 Printer.FontSize = 10
 Printer.Print Tab(13); Me.direccion; "   -    "; Me.localidad
 Printer.Print
 If Me.contado = "N" Then
   cond = "   Cuenta Corriente"
 Else
   cond = "   Contado"
 End If
 Printer.Print Tab(13); Me.abreviatura_tipoiva; cond; Tab(71); Me.CUIT
 Printer.FontName = "Courier New"
 Printer.Print
 Printer.Print
 Printer.Print
 Printer.Print Tab(16); Left$(Format$(Me.transporte, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!"), 40); Tab(71); Me.cuit_transporte
 Printer.FontSize = 14
 Printer.Print
 Printer.FontSize = 10
 Printer.Print Tab(13); Me.dir_transporte
 Printer.FontSize = 20
 Printer.Print
 Printer.FontSize = 10
 Printer.Print Tab(13); Left$(Format$(Me.CHOFER, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!"), 50); Spc(5); "D.N.I"; Me.dnichofer
 Printer.Print
 Printer.Print
 Printer.Print Tab(13); Me.PATENTE; Tab(62); Me.PATENTE_ACOPLADO
 Printer.Print
 Printer.FontSize = 20
 Printer.Print
 Printer.FontSize = 10
   Set rs = New adodb.Recordset
   q = "Select * from vta_03 where [num_int] = " & Me.numint
   rs.Open q, cn1
   cant = Space$(10)
   l = 0
   While Not rs.EOF
    i = Space$(10)
    t = Space$(4)
    pu = Space$(10)
    d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    dt = Space$(9)
    cant = Space$(8)
    cp2 = Space$(9)
    tu = Space$(8)
   If rs("cantidad") > 0 Then
      RSet cant = Format$(rs("cantidad"), "####0.00")
      dt = "  (" & t & "%)"
      cp2 = " [" & Format$(rs("id_producto"), "00000") & "] "
      tu = Format$(rs("tunidad"), "@@@@@@@@!")
   End If
    
    Printer.Print Tab(6); cant; Spc(1); tu; Spc(1); cp2; Spc(1);
    Printer.FontBold = True
    Printer.Print d; Spc(1);
    Printer.FontBold = False
    
    'imprime extra
     If Me.imprimedescextra = "S" Then
       Set rs1 = New adodb.Recordset
       q = "select * from vta_015 where [num_int] = " & Me.numint & " and [renglon] = " & rs("renglon")
       rs1.Open q, cn1
       If Not rs1.EOF And Not rs1.BOF Then
         Call lee_desc_extra(a, rs1("desc_ext"))
         For k = 0 To 4
           If a(k) <> "%%" Then
             d = Format$(Left$(a(k), 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
             Printer.Print Tab(28); d
             l = l + 1
           End If
         Next k
       
       End If
       Set rs1 = Nothing
     End If
 
 
 
    rs.MoveNext
    l = l + 1
   Wend
   Set rs = Nothing
   Printer.Print
   
   For X = l To Me.cant_lineas
     Printer.Print
   Next X
  
  For i = 1 To 6
    Printer.Print
  Next
     
  Printer.Print Tab(5); Me.observaciones
   Printer.FontBold = False
   If Y < c Then
      Printer.NewPage
   End If
  Next Y
 
 Printer.EndDoc

End Sub
Sub imprimirrto5(ByVal c As Integer)
 Dim a(5) As String
 
 
 
 dd = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 1, 2)
 mm = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 4, 2)
 aa = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 7, 4)
 'lugar = Left$(RTrim$(cl_cli.direccion) & " - (" & RTrim$(cl_cli.cp) & ")" & RTrim$(cl_cli.localidad), 120)

 tacha = Me.abreviatura & " "
 tachan = Me.letra & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000")
 For Y = 1 To c
  Printer.FontName = "Times New Roman"
  Printer.FontSize = 10
  Printer.FontBold = False
  For i = 1 To 1
    Printer.Print
  Next i
  Printer.FontSize = 14
  Printer.Print
  
  If Me.letra = "A" Then
   Printer.FontSize = 6
   Printer.Print
  End If
  Printer.FontSize = 10
   Printer.FontName = "Times New Roman"
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.FontSize = 12
   Printer.Print Tab(106); tacha;
   Printer.FontSize = 6
   Printer.Print tachan;
   Printer.FontSize = 12
   Printer.Print
   Printer.Print
   Printer.Print
   If Me.idtipocomp <> 45 Then
     Printer.FontSize = 16
     Printer.Print Tab(64); Me.abreviatura;
     Printer.FontSize = 12
   End If
   Printer.Print Tab(96); Spc(4); dd & "/" & mm & "/" & aa
   For i = 1 To 2
     Printer.Print
   Next i
 Printer.FontSize = 8
 Printer.Print
 Printer.Print
 Printer.Print
 Printer.FontName = "Times New Roman"
 Printer.FontSize = 14
 Printer.FontBold = True
 Printer.Print Tab(13); Me.razonsocial
 Printer.FontName = "Courier New"
 Printer.FontSize = 10
 Printer.FontBold = False
 Printer.Print Tab(13); "["; Format$(Me.idcliente, "0000"); "] ";
 Printer.Print Tab(57); Me.direccion
 Printer.Print Tab(57); Me.localidad
 Printer.Print Tab(13); Me.abreviatura_tipoiva; Tab(70); Me.CUIT
 Printer.FontName = "Courier New"
 Printer.FontSize = 9
 Printer.FontBold = False
 Printer.Print
  If Me.contado = "N" Then
    Printer.Print Tab(17); "Cuenta Corriente"
 Else
    Printer.Print Tab(17); "Contado"
 End If
 Printer.Print
 Printer.Print
 Printer.Print
 Printer.Print
' Printer.Print
   Set rs = New adodb.Recordset
   q = "Select * from vta_03 where [num_int] = " & Me.numint
   rs.Open q, cn1
   cant = Space$(10)
   l = 0
   While Not rs.EOF
    i = Space$(10)
    t = Space$(4)
    pu = Space$(10)
    d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    dt = Space$(9)
    cant = Space$(8)
    cp2 = Space$(9)
   If rs("cantidad") > 0 Then
      RSet cant = Format$(rs("cantidad"), "####0.00")
      dt = "  (" & t & "%)"
      cp2 = " [" & Format$(rs("id_producto"), "00000") & "] "
   End If
    
    Printer.Print Tab(3); cant; Spc(1); cp2; Spc(1);
    Printer.FontBold = True
    Printer.Print d; Spc(1);
    Printer.FontBold = False
    
    'imprime extra
     If Me.imprimedescextra = "S" Then
       Set rs1 = New adodb.Recordset
       q = "select * from vta_015 where [num_int] = " & Me.numint & " and [renglon] = " & rs("renglon")
       rs1.Open q, cn1
       If Not rs1.EOF And Not rs1.BOF Then
         Call lee_desc_extra(a, rs1("desc_ext"))
         For k = 0 To 4
           If a(k) <> "%%" Then
             d = Format$(Left$(a(k), 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
             Printer.Print Tab(25); d
             l = l + 1
           End If
         Next k
       
       End If
       Set rs1 = Nothing
     End If
 
 
 
    rs.MoveNext
    l = l + 1
   Wend
   Set rs = Nothing
   Printer.Print
   Printer.Print Tab(22); "**************************************************"
   l = l + 1
   
   For X = l To Me.cant_lineas
     Printer.Print
   Next X
  'Ultimas lineas
 
  
  For i = 1 To 2
    Printer.Print
  Next
     
  Printer.Print Tab(5); "Observaciones: " & Me.observaciones
   Printer.FontBold = False
   If Y < c Then
      Printer.NewPage
   End If
  Next Y
 
 Printer.EndDoc

End Sub


Sub formato8(ByVal c As Integer)
 'los vasquitos
 Dim a(5) As String
 
 Set rsr = New adodb.Recordset
 q = "SELECT [tasa_financiera] from g0 where [sucursal] = 0"
 rsr.Open q, cn1
 cf = Format$(rsr("tasa_financiera"), "#0.00")
 Set rsr = Nothing
 
 
 Set rsr = New adodb.Recordset
 q = "select * from vta_08, vta_02 where [id_factura] = " & Me.numint & " and [id_remito] = [num_int]"
 rsr.Open q, cn1
 tr = ""
 While Not rsr.EOF
  If rsr("id_tipocomp") = 45 Then
    tc = "R"
  Else
    tc = "D"
  End If
  tr = tr & Format$(rsr("num_comp"), "00000000") & tc & " * "
  rsr.MoveNext
 Wend
 Set rsr = Nothing
 
 
 
 dd = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 1, 2)
 mm = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 4, 2)
 aa = Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 7, 4)
 'lugar = Left$(RTrim$(cl_cli.direccion) & " - (" & RTrim$(cl_cli.cp) & ")" & RTrim$(cl_cli.localidad), 120)

 tacha = Me.abreviatura & "   "
 tachan = Me.letra & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000")
 For Y = 1 To c
  Printer.FontName = "Times New Roman"
  Printer.FontSize = 10
  Printer.FontBold = False
  For i = 1 To 7
    Printer.Print
  Next i
   Printer.FontName = "Times New Roman"
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.FontSize = 12
   Printer.Print Tab(71); tacha;
   Printer.FontSize = 8
   Printer.Print tachan;
   Printer.FontSize = 12
   Printer.Print Tab(94); Spc(4); dd & "/" & mm & "/" & aa
   For i = 1 To 3
     Printer.Print
   Next i
 Printer.FontSize = 8
 Printer.Print
 Printer.Print
 Printer.FontName = "Times New Roman"
 Printer.FontSize = 14
 Printer.FontBold = True
 Printer.Print Tab(16); Me.razonsocial
 Printer.FontName = "Courier New"
 Printer.FontSize = 10
 Printer.FontBold = False
 Printer.Print Tab(65); "["; Format$(Me.idcliente, "0000"); "] ";
 Printer.Print
 Printer.Print Tab(16); Me.direccion
 Printer.Print Tab(16); Me.localidad; Tab(65); Me.CUIT
 Printer.Print
 Printer.Print Tab(16); "Condicion Iva: "; Me.abreviatura_tipoiva
 Printer.FontName = "Courier New"
 Printer.FontSize = 9
 Printer.FontBold = False
 Printer.Print
  If Me.contado = "N" Then
    Printer.Print Tab(34); "Cuenta Corriente";
 Else
    Printer.Print Tab(34); "Contado";
 End If
 Printer.FontSize = 8
 Printer.Print Tab(84); Left$(tr, 30)
 Printer.FontSize = 9
 Printer.Print
 Printer.Print
 Printer.Print
 'Printer.Print
' Printer.Print
   Set rs = New adodb.Recordset
   q = "Select * from vta_03 where [num_int] = " & Me.numint
   rs.Open q, cn1
   cant = Space$(10)
   l = 0
   Printer.FontSize = 9
   While Not rs.EOF
    i = Space$(10)
    t = Space$(4)
    pu = Space$(10)
    d = Format$(Left$(rs("descripcion") & " ", 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    dt = Space$(9)
    cant = Space$(8)
    cp2 = Space$(9)
   If rs("cantidad") > 0 Then
      If Me.letra = "A" Then
           RSet pu = Format$(rs("pu"), "######0.00")
           RSet i = Format$(rs("importe"), "######0.00")
      Else
           RSet pu = Format$(rs("pu_final"), "######0.00")
           RSet i = Format$(rs("pu_final") * rs("cantidad"), "######0.00")
      End If
      RSet t = Format$(rs("tasaiva"), "#0.0")
      RSet cant = Format$(rs("cantidad"), "####0.00")
      dt = "  (" & t & "%)"
      cp2 = " [" & Format$(rs("id_producto"), "00000") & "] "
    End If
    
    Printer.Print Tab(4); cant; Spc(1); cp2; Spc(1);
    Printer.FontBold = True
    Printer.Print d; Spc(1);
    Printer.FontBold = False
    Printer.Print pu;
    Printer.FontSize = 7
    Printer.Print dt;
    Printer.FontSize = 9
    Printer.Print Spc(1); i
 
    'imprime extra
     If Me.imprimedescextra = "S" Then
       Set rs1 = New adodb.Recordset
       q = "select * from vta_015 where [num_int] = " & Me.numint & " and [renglon] = " & rs("renglon")
       rs1.Open q, cn1
       If Not rs1.EOF And Not rs1.BOF Then
         Call lee_desc_extra(a, rs1("desc_ext"))
         For k = 0 To 4
           If a(k) <> "%%" Then
             d = Format$(Left$(a(k), 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
             Printer.Print Tab(25); d
             l = l + 1
           End If
         Next k
       
       End If
       Set rs1 = Nothing
     End If
 
 
 
    rs.MoveNext
    l = l + 1
   Wend
   Set rs = Nothing
  
   For X = l To Me.cant_lineas
     Printer.Print
   Next X
   
  Printer.Print
  Printer.FontSize = 8
  Printer.Print Tab(9); "========================================================================================================="
  Printer.Print Tab(9); "La mora se producirá automáticamente por el solo vencimiento del plazo para el pago de la Factura."
  Printer.Print Tab(9); "Si el presente comprobante no se abona en término se le aplicará una carga financiera del " & cf & "% mensual"
  Printer.Print Tab(9); "========================================================================================================="
  Printer.Print
  Printer.Print
   
  xs = Space$(10)
   
  If Me.letra = "A" Then
    If Me.moneda = "P" Then
      'factura A en pesos
      Printer.FontSize = 8
      Printer.Print Tab(9); "Esta factura equivale a U$s " & asteriscos(Format$(Me.totalotramoneda, "#####0.00"), 9) & " Para su cancelación debe ";
      Printer.FontSize = 12
      Printer.Print Tab(67); "$"; asteriscos(Format$(Me.subtotal, "#####0.00"), 9)
      Printer.FontSize = 8
      Printer.Print Tab(9); "aplicase a  este importe la cotizacion vigente  al momento del  "
      Printer.Print Tab(9); "efectivo pago  o acreditación  en  nuestra  cta. cte.  El tipo  "
      Printer.Print Tab(9); "de cambio  es el  vendedor mayorista para el pago de Importac.  ";
      Printer.FontSize = 12
      Printer.Print Tab(67); "$"; asteriscos(Format$(Me.subtotal, "#####0.00"), 9)
      Printer.Print
      Printer.Print Tab(67); "$"; asteriscos(Format$(Me.iva, "#####0.00"), 9)
      Printer.FontSize = 8
      Printer.Print Tab(9); "Vencimiento  : " & Me.fecha_vto
      Printer.Print
      Printer.FontSize = 12
      Printer.FontBold = True
      Printer.Print Tab(67); "$"; asteriscos(Format$(Me.total, "#####0.00"), 9)
      
    Else
      'factura A en Dolars
      Printer.FontSize = 8
      Printer.Print Tab(9); "Esta factura deberá ser cancelada en Dolares estadounidenses. Que";
      Printer.FontSize = 12
      Printer.Print Tab(66); "u$s"; asteriscos(Format$(Me.subtotal, "#####0.00"), 9)
      Printer.FontSize = 8
      Printer.Print Tab(9); "Los Vasquitos Serv. Agrop. srl  haya aceptado,  o pueda aceptar,"
      Printer.Print Tab(9); "pago en pesos no  implica modificación ni renuncia  al cobro en "
      Printer.Print Tab(9); "Dolares estadounidenses. Al solo  efecto de establecer el monto ";
      Printer.FontSize = 12
      Printer.Print Tab(66); "u$s"; asteriscos(Format$(Me.subtotal, "#####0.00"), 9)
      Printer.FontSize = 8
      Printer.Print Tab(9); "del IVA,  a la  fecha de emisión de  este  comprobante el mismo "
      Printer.Print Tab(9); "asciende a la suma que se detalla a continuacion:";
      Printer.FontSize = 12
      Printer.Print Tab(66); "u$s"; asteriscos(Format$(Me.iva, "#####0.00"), 9)
      Printer.FontSize = 8
      Printer.Print Tab(9); "IVA: $"; Format$(Me.iva * Me.cotizaciondolar, "0.00"); " Al tipo de cambio de $"; Me.cotizaciondolar; " por U$s."
      Printer.Print
      Printer.FontSize = 10
      Printer.Print Tab(7); "Vencimiento  : " & Me.fecha_vto;
      Printer.FontSize = 12
      Printer.FontBold = True
      Printer.Print Tab(66); "u$s"; asteriscos(Format$(Me.total, "#####0.00"), 9)
    End If
 Else
    'factura B
   If Me.moneda = "P" Then
      'factura B en pesos
      Printer.FontSize = 8
      Printer.Print Tab(9); "Esta factura equivale a U$s " & asteriscos(Format$(Me.totalotramoneda, "#####0.00"), 9) & " Para su cancelación debe "
      Printer.Print Tab(9); "aplicase a  este importe la cotizacion vigente  al momento del  "
      Printer.Print Tab(9); "efectivo pago  o acreditación  en  nuestra  cta. cte.  El tipo  "
      Printer.Print Tab(9); "de cambio  es el  vendedor mayorista para el pago de Importac.  "
      Printer.Print
      Printer.Print
      Printer.Print Tab(9); "Vencimiento  : " & Me.fecha_vto
      Printer.Print
      Printer.FontSize = 12
      Printer.FontBold = True
      Printer.Print Tab(67); "$"; asteriscos(Format$(Me.total, "#####0.00"), 9)
      
    Else
      'factura A en Dolars
      Printer.FontSize = 8
      Printer.Print Tab(9); "Esta factura deberá ser cancelada en Dolares estadounidenses. Que"
      Printer.Print Tab(9); "Los Vasquitos Serv. Agrop. srl  haya aceptado,  o pueda aceptar,"
      Printer.Print Tab(9); "pago en pesos no  implica modificación ni renuncia  al cobro en "
      Printer.Print Tab(9); "Dolares estadounidenses. Al solo  efecto de establecer el monto "
      Printer.Print Tab(9); "del IVA,  a la  fecha de emisión de  este  comprobante el mismo "
      Printer.Print Tab(9); "asciende a la suma que se detalla a continuacion:"
      Printer.Print Tab(9); "IVA: $"; Format$(Me.iva * Me.cotizaciondolar, "0.00"); " Al tipo de cambio de $"; Me.cotizaciondolar; " por U$s."
      Printer.Print
      Printer.FontSize = 10
      Printer.Print Tab(7); "Vencimiento  : " & Me.fecha_vto;
      Printer.FontSize = 12
      Printer.FontBold = True
      Printer.Print Tab(66); "u$s"; asteriscos(Format$(Me.total, "#####0.00"), 9)
    End If
 End If

   
   
   If Me.letra = "B" Then
     Printer.Print Tab(10); "Orientacion al Consumidor Prov. BA 08002229042"
      ' Printer.Print
      ' Printer.Print
   Else
     'Printer.Print
   End If
   
   
      
   
   If Y < c Then
      Printer.NewPage
   End If
  Next Y
 
 Printer.EndDoc

End Sub

Sub BORRoRBO()
    'borro RBO
 On Error GoTo errbop
      
 'primero verifico que los ch. recibidos estan en cartera
 q = "select * from cyb_03 where [num_int_rbo] = " & Me.numint & " and [estado] <> 'C'"
 Set rs = New adodb.Recordset
 rs.Open q, cn1
 If Not rs.BOF And Not rs.EOF Then
    MsgBox ("Para poder eliminar el Recibo todos los cheques ingresados deben estar en Cartera")
    Set rs = Nothing
 Else
      Set rs = Nothing
      cn1.BeginTrans
      
      QUERY = "DELETE FROM VTA_02 WHERE [num_int] = " & Me.numint
      cn1.Execute QUERY
      
      QUERY = "DELETE FROM a7 WHERE [num_int] = " & Me.numint
      cn1.Execute QUERY
  
  
      'ch terc
      QUERY = "delete from cyb_03 where [num_int_RBO]= " & Me.numint
      cn1.Execute QUERY
      
        
      
      'FORMAS DE PAGO
      QUERY = "DELETE FROM VTA_04 WHERE [num_int] = " & Me.numint
      cn1.Execute QUERY
      
      'mov. caja
      QUERY = "DELETE FROM cyb_05 WHERE [modulo] = 'V' and [num_mov_int] = " & Me.numint
      cn1.Execute QUERY
      
      'mov banco por transferencias
      QUERY = "DELETE FROM cyb_04 WHERE [modulo] = 'V' and [num_mov_int] = " & Me.numint
      cn1.Execute QUERY
       
      
      nop = Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000")
       
      QUERY = "update VTA_02 set  [estado_pago]='N'"
      QUERY = QUERY & " where [RECIBO_PAGO]= '" & nop & "'"
      cn1.Execute QUERY
      
      
      
      q = "select * from c_02 where [modulo] = 'V' and [num_mov_int] = " & Me.numint
      Set rs = New adodb.Recordset
      rs.Open q, cn1
      If Not rs.EOF And Not rs.BOF Then
        nic = rs("num_interno")
      Else
        nic = 0
      End If
      Set rs = Nothing
      
      If nic <> 0 Then
         QUERY = "DELETE FROM c_02 WHERE [num_interno] = " & nic
         cn1.Execute QUERY
        
         QUERY = "DELETE FROM c_03 WHERE [num_interno] = " & nic
         cn1.Execute QUERY
      End If
       
      
      Set rs = New adodb.Recordset
      q = "select * from vta_010 where [num_int_rbo] = " & Me.numint
      rs.Open q, cn1, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
         nic = rs("num_int_comp")
         Set rs2 = New adodb.Recordset
         q = "select [estado_pago], [saldo_impago02], [recibo_pago] from vta_02 where [num_int] = " & nic
         rs2.Open q, cn1, adOpenDynamic, adLockOptimistic
         If Not rs2.EOF And Not rs2.BOF Then
            rs2("saldo_impago02") = rs2("saldo_impago02") + rs("importe_pagado")
            rs2("estado_pago") = "N"
            rs2("recibo_pago") = "0000-00000000"
            rs2.Update
         End If
         Set rs2 = Nothing
         rs.Delete
         rs.MoveNext
      Wend

      
      QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
      QUERY = QUERY & " VALUES ('Borra Recibo NI " & Me.numint & "' , " & para.id_usuario & ", 'V'," & Me.numint & ", '" & Now & "', '" & Me.abreviatura & " " & Me.letra & " " & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000") & "', 4, " & Me.idcliente & ")"
      cn1.Execute QUERY
      
      
      
      cn1.CommitTrans
  
     ' MsgBox ("Comprobante Eliminado!!!!")
  
  End If
  
      
Exit Sub
errbop:
MsgBox ("Imposible Eliminar Comprobante. Verifique sus permisos e intente mas tarde. Si aun asi no puede llame al serv. Tecnico(Clase: Comprobantes Vta. Modulo: Borra RBO")

End Sub
Public Sub actual(ByVal TIPOCOMP As Integer)
'SACA LOS DATOS DE UN COMPROBANTE QUE SE ESTA POR CREAR
'STOCK, CTACTE

 Set rs = New adodb.Recordset
 q = "select * from vta_06 where [sucursal] = " & Me.sucursal & " and  [id_tipocomp] = " & TIPOCOMP
 rs.MaxRecords = 1
 rs.Open q, cn1
 If Not rs.EOF And Not rs.BOF Then
     Me.STOCK = rs("STOCK")
     Me.ctacte = rs("CTACTE")
     Me.grabado = rs("iva")
     Me.abreviatura = rs("abreviatura")
     Me.venta = rs("venta")
     Me.PROPIO = rs("PROPIO")
     Me.contabilidad = rs("contabilidad")
     Me.tipoImpresora = rs("tipo_impresora")
     Me.cant_lineas = rs("cant_lineas")
     Me.imprimedescextra = rs("imprime_desc_extra")
     Me.cod_afip_a = rs("cod_afip_a")
     Me.cod_afip_b = rs("cod_afip_b")
 Else
  MsgBox ("Error al Inicializar Comprobante. Clase: Comprobantes Venta  Proc.Actual: Error en el punto de venta o en el tipo de comprobante")
  Me.STOCK = "N"
  Me.ctacte = "N"
  Me.grabado = "N"
  Me.abreviatura = "Err"
  Me.venta = "N"
  Me.PROPIO = "S"
  Me.contabilidad = "N"
  Me.tipoImpresora = "N"
  Me.cant_lineas = 0
  Me.imprimedescextra = "N"
  Me.cod_afip_a = 1
  Me.cod_afip_b = 6
 End If

Me.idtipocomp = TIPOCOMP

Set rs = Nothing


End Sub

Public Sub actual2(ByVal TIPOCOMP As Integer, ByVal suc As Integer)
'SACA LOS DATOS DE UN COMPROBANTE QUE SE ESTA POR CREAR
'STOCK, CTACTE

 Set rs = New adodb.Recordset
 q = "select * from vta_06 where [sucursal] = " & suc & " and  [id_tipocomp] = " & TIPOCOMP
 rs.MaxRecords = 1
 rs.Open q, cn1
 If Not rs.EOF And Not rs.BOF Then
     Me.STOCK = rs("STOCK")
     Me.ctacte = rs("CTACTE")
     Me.grabado = rs("iva")
     Me.abreviatura = rs("abreviatura")
     Me.venta = rs("venta")
     Me.PROPIO = rs("PROPIO")
     Me.contabilidad = rs("contabilidad")
     Me.tipoImpresora = rs("tipo_impresora")
     Me.cant_lineas = rs("cant_lineas")
     Me.imprimedescextra = rs("imprime_desc_extra")
     Me.cod_afip_a = rs("cod_afip_a")
     Me.cod_afip_b = rs("cod_afip_b")
    
     
 Else
  MsgBox ("Error al Inicializar Comprobante. Clase: Comprobantes Venta  Proc.Actual: Error en el punto de venta o en el tipo de comprobante")
  Me.STOCK = "N"
  Me.ctacte = "N"
  Me.grabado = "N"
  Me.abreviatura = "Err"
  Me.venta = "N"
  Me.PROPIO = "S"
  Me.contabilidad = "N"
  Me.tipoImpresora = "N"
  Me.cant_lineas = 0
  Me.imprimedescextra = "N"
  Me.cod_afip_a = 1
  Me.cod_afip_b = 6
    
 End If

Me.idtipocomp = TIPOCOMP

Set rs = Nothing


End Sub
Sub SACANUMCOMP()
   Set rs = New adodb.Recordset
   q = "select [ult_num_a], [ult_num_b], [ult_num_c], [ult_num_e] from vta_06 where [sucursal] = " & Me.sucursal & " and [id_tipocomp] = " & Me.idtipocomp
   rs.MaxRecords = 1
   rs.Open q, cn1, adOpenStatic, adLockOptimistic
   If Not rs.BOF And Not rs.EOF Then
     If Me.ctacte <> "N" Or Me.idtipocomp = 40 Then
      Select Case Me.letra
       Case Is = "A"
             Me.numcomp = Format$(rs("ULT_NUM_a") + 1, "00000000")
       Case Is = "B"
             Me.numcomp = Format$(rs("ULT_NUM_b") + 1, "00000000")
       Case Is = "C"
             Me.numcomp = Format$(rs("ULT_NUM_c") + 1, "00000000")
       Case Is = "E"
             Me.numcomp = Format$(rs("ULT_NUM_e") + 1, "00000000")
       
       Case Else
             Me.numcomp = Format$(rs("ULT_NUM_a") + 1, "00000000")
      End Select
     Else
       Me.numcomp = Format$(rs("ULT_NUM_a") + 1, "00000000")
     End If
                    
   Else
     MsgBox ("Error. No se puede Inicializar el Comprobante")
   End If
   Set rs = Nothing
End Sub
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



Public Property Let ctacte(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.CTACTE = 5
    mvarCTACTE = vData
End Property


Public Property Get ctacte() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.CTACTE
    ctacte = mvarCTACTE
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



