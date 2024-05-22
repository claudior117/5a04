VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "comprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"STOCK"
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
'variables locales para almacenar los valores de las propiedades
Private mvarnumint As Long 'copia local
Private mvarsucursal As Integer 'copia local
Private mvarnumcomp As Long 'copia local
Private mvarletra As String 'copia local
Private mvaridtipocomp As Integer 'copia local
Private mvaridproveedor As Long 'copia local
Private mvarfecha As Date 'copia local
Private mvarfechaprobentrega As Date 'copia local
Private mvarestado As String 'copia local
Private mvarsubtotal As Double 'copia local
Private mvariva As Double 'copia local
Private mvarnograbado As Double 'copia local
Private mvarpercep As Double 'copia local
Private mvartotal As Double 'copia local
Private mvarfecharecepcion As Date 'copia local
Private mvaridusuario As Long 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvaridcodretgan As Integer 'copia local
Private mvaridcuenta As Long 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarSTOCK As String 'copia local
Private mvarCTACTE As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvargrabado As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarestado_pago As String 'copia local
Private mvarnum_op As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarabreviatura As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarcompra As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarobservaciones As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarcontacto As String 'copia local
Private mvarcondiciones As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarcontabilidad As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarmoneda As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarfechavto As Date 'copia local
Private mvarcotizacion As Double 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarsaldoimpago As Double 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarcontado As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvartotal_om As Double 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarzona As Integer 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarimprimedescextra As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvaridcliente As Long 'copia local
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



Public Property Let imprimedescextra(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.imprimedescextra = 5
    mvarimprimedescextra = vData
End Property


Public Property Get imprimedescextra() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.imprimedescextra
    imprimedescextra = mvarimprimedescextra
End Property



Public Property Let zona(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.zona = 5
    mvarzona = vData
End Property


Public Property Get zona() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.zona
    zona = mvarzona
End Property



Public Property Let total_om(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.total_om = 5
    mvartotal_om = vData
End Property


Public Property Get total_om() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.total_om
    total_om = mvartotal_om
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



Public Property Let saldoimpago(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.saldoimpago = 5
    mvarsaldoimpago = vData
End Property


Public Property Get saldoimpago() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.saldoimpago
    saldoimpago = mvarsaldoimpago
End Property



Public Property Let cotizacion(ByVal vData As Double)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.cotizacion = 5
    mvarcotizacion = vData
End Property


Public Property Get cotizacion() As Double
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.cotizacion
    cotizacion = mvarcotizacion
End Property



Public Property Let fechavto(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.fechavto = 5
    mvarfechavto = vData
End Property


Public Property Get fechavto() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.fechavto
    fechavto = mvarfechavto
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



Public Sub imprimeretib(ByVal c As Integer)
'imprime retencion de iB
'n es el num-int
Dim gf1 As String

Set rs = New adodb.Recordset
q = "select [sucursal], [num_comprobante], [fecha], [denominacion], [localidad], [cp], [direccion], [cuit], [alicuota_ret], [monto_suj_ret], [total] from a5, a1, g1 where [num_int] = " & Me.numint & " and a5.[id_proveedor] = a1.[id_proveedor] and a5.[id_usuario] = g1.[id_usuario]"
rs.MaxRecords = 1
rs.Open q, cn1
If Not rs.BOF And Not rs.EOF Then
     gf1 = "Courier New"
     ip = Space$(9)
     ret = Space$(9)
     c2 = InputBox("Cantidad de Copias", "Imprimir Retencion IB", c)
     If Val(c2) > 0 Then
      For h = 1 To Val(c2)
       Call definefont(gf1, "S", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Comprobante Retencion Ingresos Brutos "
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Reg.Ret. de INGRESOS BRUTOS de la Prov. de Buenos Aires"
       Call definefont(gf1, "S", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print Spc(50); "Nro.: ";
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print Spc(3); Format$(rs("SUCURSAL"), "0000") & "-" & Format$(rs("num_comprobante"), "00000000")
       Call definefont(gf1, "S", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print Spc(50); "Fecha:";
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print Spc(3); rs("fecha")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
       
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "__________________________________________________________________________"
       Call definefont(gf1, "S", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Agente de Retencion:"; glo.nombrecli
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "                    "; glo.direccioncli
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "           c.u.i.t :"; glo.CUIT
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "__________________________________________________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "__________________________________________________________________________"
       Call definefont(gf1, "S", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Contribuyente: "; Spc(1);
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print rs("denominacion"); " -"; rs("localidad"); "- (" & rs("cp"); ")"
       Call definefont(gf1, "S", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Direccion....: ";
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print rs("direccion")
       Call definefont(gf1, "S", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "CUIT.........: ";
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print rs("cuit")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "__________________________________________________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
       Call definefont(gf1, "N", 8)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Datos de la Operacion"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "---------------------------------------------------"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Comp. Sujeto a Ret.        Importe Suj. Ret "
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "---------------------------------------------------"
       k = 1
       b = Space$(10)
       a = Space$(7)
       r = Space$(10)
       RSet a = Format$(rs("alicuota_ret"), "###0.00")
       'busco op.
       Set rs1 = New adodb.Recordset
       q = "select  [importe_imponible], [num_comprobante] from a16 where [num_int_ret] = " & Me.numint
       rs1.Open q, cn1
       ti = 0
       While Not rs1.EOF
        c2 = Format$(rs1("num_comprobante"), "@@@@@@@@@@@@@@@@@@@@@@@@@!")
        RSet b = Format$(rs1("importe_imponible"), "######0.00")
         Printer.Print c2 & "     " & b
         ti = ti + Val(b)
         rs1.MoveNext
       Wend
       Set rs1 = Nothing
       Printer.Print
       Call definefont(gf1, "S", 12)
       Printer.Print "       Total Imponible: " & Format$(rs("monto_suj_ret"), "#####0.00")
       Call definefont(gf1, "N", 8)
       Printer.Print

       Printer.Print "*** Alicuota de Retencion " & a & "% fijada según Padrón de Contribuyentes otorgado por Rentas de la Pcia. de Bs.AS. *** "
       Printer.Print
       Printer.Print
       Call definefont(gf1, "S", 12)
       Printer.Print "Total a retener segun O.P. " & Format$(sucop, "0000") & "-" & Format$(numop, "00000000") & " : ....... $" & Format$(rs("total"), "0.00")
       Printer.Print
       Printer.Print
       Printer.Print "Son pesos: ......";
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print convierte(rs("total"))
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
       Printer.PaintPicture gen_logo.Picture2.Picture, 7100, 10100

       Call definefont(gf1, "N", 9)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.NewPage
     Next h
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
     Printer.EndDoc
  End If
End If
Set rs = Nothing

End Sub

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



Public Property Let condiciones(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.condiciones = 5
    mvarcondiciones = vData
End Property


Public Property Get condiciones() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.condiciones
    condiciones = mvarcondiciones
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



Public Sub imprimir()
'previamente es necesario cargar el comprobante
Dim desc_comp As String
Set rs = New adodb.Recordset
q = "select * from g2 where [id_tipo_comp] = " & Me.idtipocomp
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
   copias = rs("copias")
   formato = rs("formato")
   Me.imprimedescextra = rs("imprime_desc_extra")
   
Else
  copias = 0
  formato = "1"
  Me.imprimedescextra = "N"
End If
Set rs = Nothing

Select Case Me.idtipocomp
  Case Is = 50
    Select Case formato
      Case Is = "1" 'generica
         Call imprimeop1(copias)
      Case Is = "2" 'degser
         Call imprimeop2(copias)
      Case Is = "3" 'framahi/el nuevo con logo.jpg
         Load gen_logo
         gen_logo.Picture2.Picture = LoadPicture(App.Path & "\tools\firma1.jpg")
         Call imprimeop3(copias)
         Unload gen_logo
      Case Is = "4" ' radio CON LOGO2.jpg
          Load gen_logo
          gen_logo.Picture1.Picture = LoadPicture(App.Path & "\tools\logo2.jpg")
          Call imprimeop3(copias)
          Unload gen_logo
      Case Is = "5" 'fabrimac
          Load gen_logo
          gen_logo.Picture2.Picture = LoadPicture(App.Path & "\tools\firma1.jpg")
          Call imprimeop5(copias)
          Unload gen_logo
     
         
         
                   
          
  End Select
  
  Case Is = 95
     Load gen_logo
     gen_logo.Picture2.Picture = LoadPicture(App.Path & "\tools\firma1.jpg")
     Call IMPRIMERETG(copias)
     Unload gen_logo
  
  Case Is = 97
     Load gen_logo
     gen_logo.Picture2.Picture = LoadPicture(App.Path & "\tools\firma1.jpg")
     Call IMPRIMERETiva(copias)
     Unload gen_logo
  
  
  Case Is = 65
     Load gen_logo
     gen_logo.Picture2.Picture = LoadPicture(App.Path & "\tools\firma1.jpg")
     Call imprimeoc(copias)
     Unload gen_logo
  Case Is = 70
     Call imprimecotiz(copias)
 
 Case Is = 96
     Load gen_logo
     gen_logo.Picture2.Picture = LoadPicture(App.Path & "\tools\firma1.jpg")
     Call imprimeretib(copias)
     Unload gen_logo

End Select

End Sub

Sub exporta_oc(ByVal nioc As Long)
  'exportacion orden de compra por numero interno
   t = "____________________________________________________________________________________________________"
  
   c = Space$(10)
   J = MsgBox("Confirma Exportacion de Reporte en C:\exporta" & a & ".txt", 4)
   If J = 6 Then
        'On Error GoTo errifg
        Open "c:\exporta.txt" For Output As #1
     
        Print #1, glo.nombrecli
        Print #1, glo.direccioncli
        Print #1, "Te: " & glo.TECLI & "   Cuit: " & glo.CUIT
        Print #1, t
        Print #1, ""
     
        Set rs = New adodb.Recordset
        q = "select [fecha], [sucursal], [num_comprobante], [denominacion], [direccion], [localidad] from a5, a1 where [num_int] = " & nioc & " and a5.[id_proveedor] = a1.[id_proveedor]"
        rs.MaxRecords = 1
        rs.Open q, cn1
        If Not rs.EOF And Not rs.BOF Then
            Print #1, "ORDEN DE COMPRA"
            Print #1,
            Print #1, Spc(45); "Fecha..... : "; rs("fecha")
            Print #1, Spc(45); "Numero.....: " & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comprobante"), "00000000")
            Print #1, "Proveedor........:"; rs("denominacion")
            Print #1, "Direccion........:"; rs("direccion")
            Print #1, "Localidad........:"; rs("localidad")
            Print #1, "---------------------------------------------------------------------------------------"
            Print #1, "Id.     Cod. Barra    Detalle                                              Cantidad   "
            Print #1, "---------------------------------------------------------------------------------------"
            Set rs2 = New adodb.Recordset
            q = "select a6.[id_producto], [cod_barra], [detalle], [cantidad] from a6, a2 where [num_int] = " & nioc & " and a6.[id_producto] = a2.[id_producto]"
            rs2.Open q, cn1
            nr = 0
            While Not rs2.EOF
              RSet c = Format$(rs2("cantidad"), "######0.00")
              Print #1, Format$(rs2("id_producto"), "00000") & " " & Format$(rs2("cod_barra"), "000000000000000") & " " & Format$(Left$(rs2("detalle"), 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!") & " " & c
              rs2.MoveNext
              nr = nr + 1
            Wend
            Set rs2 = Nothing
            Print #1,
            Print #1, "Cantidad de Registros:  " & nr
            
       End If
       Set rs = Nothing
       Close #1
       MsgBox ("Exportacion de archivo " & archivo & ".txt  CORRECTA!!")
End If

Exit Sub
errifg:
MsgBox ("Error al generar archivo " & archivo)
Exit Sub



End Sub


Sub imprimeoc(ByVal c As Integer)
Dim a(5) As String
Set rs3 = New adodb.Recordset
q = "select [descripcion] from g2 where [id_tipo_comp] = " & Me.idtipocomp
rs3.Open q, cn1
If Not rs3.EOF And Not rs3.BOF Then
   desc_comp = rs3("descripcion")
Else
  desc_comp = "Orden de Compra"
End If
Set rs3 = Nothing
l = InputBox("Indique cantidad de Copias", , c)
If Val(l) > 0 And Val(l) <= 5 Then
 For Y = 1 To l
   Printer.Orientation = 2
   Call imprimeempresa(14)
   Printer.Print
   Printer.FontName = "Times New Roman"
   Printer.FontSize = 12
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print Tab(10); desc_comp
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
   Printer.Print Tab(3); "Proveedor:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Set cl_prov = New proveedores
   cl_prov.carga (Me.idproveedor)
   If cl_prov.idprov > 0 Then
      Printer.Print "  " & cl_prov.razonsocial; Tab(92);
   Else
      Printer.Print "  "; Tab(92);
   End If
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print "Emitida por:";
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
   Printer.Print Tab(3); "Direccion:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.Print "  " & cl_prov.direccion
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print Tab(3); "Localidad:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.Print "  " & cl_prov.localidad; Tab(99);
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "Te: ";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & cl_prov.te
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(3); "Provincia: ";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & cl_prov.provincia
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(7); "Email: ";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & cl_prov.email; Tab(99);
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "Cuit:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & cl_prov.CUIT
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(3); "Contacto:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & Me.contacto; Tab(92);
Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "Moneda:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   If Me.moneda = "P" Then
     m = "Pesos"
   Else
     m = "Dolares"
   End If
   Printer.Print "  " & m
   
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Courier new"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 9
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "_____________________________________________________________________________________________________________________________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "        Basico  Detalle                                        Observaciones          Cant.   Unid.    P.U     % Iva     Importe"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "_____________________________________________________________________________________________________________________________________________________"
   Set rs = New adodb.Recordset
   q = "select * from a6 where [num_int] = " & Me.numint
   rs.Open q, cn1
   r = 1
   cant = Space$(10)
   pu = Space$(10)
   im = Space$(10)
   ti = Space$(7)
   
   l = 0
  
   While Not rs.EOF
       renglon = Format$(r, "000")
       basico = Format$(rs("id_producto"), "00000")
       Detalle = Format$(Left$(rs("detalle"), 45), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
       obsp = Format$(Left$(rs("observaciones"), 20), "@@@@@@@@@@@@@@@@@@@@!")
       RSet cant = Format$(rs("cantidad"), "#####0.00")
       RSet pu = Format$(rs("pu"), "######0.00")
       RSet im = Format$(rs("importe"), "######0.00")
       RSet ti = Format$(rs("tasa_iva"), "##0.00")
       u = rs("unidad")
       Printer.Print "   " & renglon & "  " & basico & "  " & Detalle & " " & obsp & " " & cant & " " & u & " " & pu & " " & ti & " " & im
      
        'busca descripcion extra
        'imprime extra
       If Me.imprimedescextra = "S" Then
       Set rs1 = New adodb.Recordset
       q = "select * from a21 where [num_int] = " & Me.numint & " and [renglon] = " & rs("renglon")
       rs1.Open q, cn1
       If Not rs1.EOF And Not rs1.BOF Then
         Call lee_desc_extra(a, rs1("descripcion"))
         For k = 0 To 4
           If a(k) <> "%%" Then
             d = Format$(Left$(a(k), 45), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
             Printer.Print Tab(15); d
             l = l + 1
             
           End If
         Next k
       
       End If
       Set rs1 = Nothing
     End If
               
               
               
       l = l + 1
       r = r + 1
       rs.MoveNext
   Wend
   Set rs = Nothing
   Printer.FontName = "Times New Roman"
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.FontSize = 10
   For X = 1 To 20 - l
      Printer.Print
   Next X
   s = Space$(10)
   i = Space$(10)
   t = Space$(10)
   RSet s = Format$(Me.subtotal, "######0.00")
   RSet i = Format$(Me.iva, "######0.00")
   RSet t = Format$(Me.total, "######0.00")
   Printer.Print "_____________________________________________________________________________"
   Printer.Print "  SUBTOTAL:  "; s; "          IVA: "; i; "             TOTAL:"; t
   Printer.Print "_____________________________________________________________________________"
   Printer.Print
   Printer.Print Tab(3); "Observaciones:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.Print "  "; Me.observaciones
   Printer.FontUnderline = True
   Printer.FontBold = True
'   Printer.Print
   Printer.Print Tab(3); "Fecha Entrega:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  "; Me.fechaprobentrega
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(3); "Condiciones:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  "; Me.condiciones
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.PaintPicture gen_logo.Picture2.Picture, 9100, 9100
   
   Set cl_prov = Nothing
   Set cl_comp = Nothing
   If Y < c Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.NewPage
   End If
 Next Y
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 'Printer.Orientation = 1
 Printer.EndDoc
 End If
End Sub




Sub imprimecotiz(ByVal c As Integer)
l = InputBox("Indique cantidad de Copias", , c)
If Val(l) > 0 And Val(l) <= 5 Then
 For Y = 1 To l
   'Printer.Orientation = 1
   Call imprimeempresa(14)
   Printer.Print
   Printer.FontName = "Times New Roman"
   Printer.FontSize = 12
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print "SOLICITUD DE COTIZACION"
   Printer.FontSize = 10
   Printer.Print Tab(78); "Nro. Comprobante:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.FillStyle = False
   Printer.Print "  " & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000")
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print Tab(91); "Fecha:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.Print "  " & Format$(Me.fecha, "dd/mm/yyyy")
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print
   Printer.Print Tab(3); "Proveedor:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Set cl_prov = New proveedores
   cl_prov.carga (Me.idproveedor)
   If cl_prov.idprov > 0 Then
      Printer.Print "  " & cl_prov.razonsocial; Tab(90);
   Else
      Printer.Print "  "; Tab(90);
   End If
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print "Emitida por:";
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
   Printer.Print Tab(3); "Direccion:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.Print "  " & cl_prov.direccion
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print Tab(3); "Localidad:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.Print "  " & cl_prov.localidad; Tab(99);
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "Te: ";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & cl_prov.te
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(3); "Provincia: ";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & cl_prov.provincia
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(7); "Email: ";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & cl_prov.email; Tab(97);
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "Cuit:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & cl_prov.CUIT
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(3); "Contacto:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & Me.contacto
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Courier new"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 9
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "__________________________________________________________________________________________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "       Basico  Detalle                                       Observaciones            Cant.   Unid.    "
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "__________________________________________________________________________________________________________________"
   Set rs = New adodb.Recordset
   q = "select * from a6 where [num_int] = " & Me.numint
   rs.Open q, cn1
   r = 1
   cant = Space$(10)
   pu = Space$(10)
   im = Space$(10)
   ti = Space$(7)
   
   While Not rs.EOF
       renglon = Format$(r, "000")
       basico = Format$(rs("id_producto"), "00000")
       Detalle = Format$(Left$(rs("detalle"), 45), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
       obsp = Format$(Left$(rs("observaciones"), 20), "@@@@@@@@@@@@@@@@@@@@!")
       RSet cant = Format$(rs("cantidad"), "######0.00")
       u = rs("unidad")
       Printer.Print "   " & renglon & "  " & basico & "  " & Detalle & " " & obsp & " " & cant & "    " & u
       r = r + 1
       rs.MoveNext
   Wend
   Set rs = Nothing
   Printer.FontName = "Times New Roman"
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.FontSize = 10
   For X = 1 To 40 - r
      Printer.Print
   Next X
   Printer.Print "_____________________________________________________________________________"
   Printer.Print "  "
   Printer.Print "_____________________________________________________________________________"
   Printer.Print
   Printer.Print Tab(3); "Observaciones:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.Print "  "; Me.observaciones
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(3); "Condiciones:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  "; Me.condiciones
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   
   Set cl_prov = Nothing
   Set cl_comp = Nothing
   If Y < c Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.NewPage
   End If
 Next Y
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 'Printer.Orientation = 1
 Printer.EndDoc
 End If
End Sub

Sub IMPRIMERETG(ByVal c As Integer)
'imprime retencion de ganancia
'c es copias

Dim gf1 As String

Set rs = New adodb.Recordset
q = "select * from a5, a1, g1 where [num_int] = " & Me.numint & " and a5.[id_proveedor] = a1.[id_proveedor] and a5.[id_usuario] = g1.[id_usuario]"
rs.MaxRecords = 1
rs.Open q, cn1
If Not rs.BOF And Not rs.EOF Then
     Set rs1 = New adodb.Recordset
     q = "select [concepto] from i_02 where [id_impuesto] = 217 and [id_concepto] = " & rs("a5.id_codretgan")
     rs1.Open q, cn1
     If Not rs1.EOF And Not rs1.BOF Then
        descconcepto = rs1("concepto")
     Else
        descconcepto = "Sin detallar"
     End If
     Set rs1 = Nothing
     
     gf1 = "Courier New"
     ip = Space$(19)
     ret = Space$(19)
     c2 = InputBox("Cantidad de Copias", "Imprimir Retencion Ganancias", c)
     If Val(c2) > 0 Then
      For h = 1 To Val(c2)
       Call definefont(gf1, "S", 12)
       Printer.Print
       Printer.Print "Comprobante Retencion Ganancia - R.G.(AFIP) 830"
       Call definefont(gf1, "N", 12)
       Printer.Print "Regimen de Retencion del IMPUESTO A LAS GANANCIAS"
       Call definefont(gf1, "S", 12)
       Printer.Print Spc(50); "Nro.: ";
       Call definefont(gf1, "N", 12)
       Printer.Print Spc(3); Format$(rs("SUCURSAL"), "0000") & "-" & Format$(rs("num_comprobante"), "00000000")
       Call definefont(gf1, "S", 12)
       Printer.Print Spc(50); "Fecha:";
       Call definefont(gf1, "N", 12)
       Printer.Print Spc(3); rs("fecha")
       Printer.Print
       Call definefont(gf1, "N", 12)
       Printer.Print "__________________________________________________________________________"
       Call definefont(gf1, "S", 12)
       Printer.Print "Agente de Retencion:"; glo.nombrecli
       Printer.Print "Direccion..........:"; glo.direccioncli; "   (2705)"
       Printer.Print "C.u.i.t............:"; glo.CUIT
       Call definefont(gf1, "N", 12)
       Printer.Print "__________________________________________________________________________"
       Printer.Print
       Printer.Print
       Printer.Print
       Call definefont(gf1, "N", 12)
       Printer.Print "__________________________________________________________________________"
 
       Call definefont(gf1, "S", 12)
       Printer.Print "Contribuyente: ";
       Call definefont(gf1, "N", 12)
       Printer.Print rs("denominacion"); "  ("; Format$(rs("a5.id_proveedor"), "00000"); ")    "
       Call definefont(gf1, "S", 12)
       Printer.Print "Direccion....: ";
       Call definefont(gf1, "N", 12)
       Printer.Print rs("direccion") & "  -" & rs("localidad") & "-   (" & rs("cp") & ")"
       Call definefont(gf1, "S", 12)
       Printer.Print "CUIT.........: ";
       Call definefont(gf1, "N", 12)
       Printer.Print rs("cuit")
       Printer.Print "__________________________________________________________________________"
       Printer.Print
       Printer.Print
       Call definefont(gf1, "N", 10)
       Printer.Print "Regimen: " & rs("a5.id_codretgan") & "  " & descconcepto
       Printer.Print
       RSet ip = Format$(rs("pagos_realizados"), para.formato_numerico)
       Printer.Print "Pagos a la fecha   : " & ip
       RSet ip = Format$(rs("pago_actual"), para.formato_numerico)
       Printer.Print "Pago a Efectuar    : " & ip
       RSet ip = Format$(rs("minimo_no_imp"), para.formato_numerico)
       Printer.Print "Minimo No Imponible: " & ip
       RSet ip = Format$(rs("monto_suj_ret"), para.formato_numerico)
       Printer.Print "Monto Suj.a Reten. : " & ip
       RSet ip = Format$(rs("alicuota_ret"), para.formato_numerico)
       Printer.Print "Alicuota           : " & ip
       RSet ip = Format$(rs("ret_mes"), para.formato_numerico)
       Printer.Print "Ret. Anteriores    : " & ip
       RSet ip = Format$(rs("total"), para.formato_numerico)
       Printer.Print "RETENCION          : " & ip
       Printer.Print
       Printer.Print
       Call definefont(gf1, "S", 12)
       Printer.Print "Son pesos: ";
       Call definefont(gf1, "N", 9)
       Printer.Print convierte(rs("total"))
       Printer.Print
       Printer.Print "Ingresa en DDJJ periodo: "; rs("condiciones")
       Printer.Print
       Printer.Print
       Call definefont(gf1, "N", 8)
       Printer.Print "Orden de Pago"
       sop = Val(Mid$(Me.num_op, 1, 4))
       nop = Val(Mid$(Me.num_op, 6, 8))
       Set rs2 = New adodb.Recordset
       q = "select [fecha], [total] from a5 where [id_tipocomp] = 50 and [sucursal] = " & sop & " and [num_comprobante] = " & nop
       rs2.Open q, cn1
       If Not rs2.BOF And Not rs2.EOF Then
         Printer.Print "------------------------------------------------"
         Printer.Print "Fecha         Numero                  Importe      "
         Printer.Print "------------------------------------------------"
         Printer.Print Me.fecha & "  " & Me.observaciones & "        " & Format$(rs2("total"), para.formato_numerico)
       Else
         Printer.Print "------------------------------------------------"
         Printer.Print "Fecha         Numero                  Importe      "
         Printer.Print "------------------------------------------------"
         Printer.Print Me.fecha & "  " & Me.observaciones
       End If
       Set rs2 = Nothing
       Printer.PaintPicture gen_logo.Picture2.Picture, 7100, 10100
       
         Call definefont(gf1, "N", 12)
       Printer.Print
       Printer.Print
       Printer.Print
       Printer.Print
       
      
       Call definefont(gf1, "N", 9)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       
        
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.NewPage
     Next h
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
     Printer.EndDoc
  End If
End If
Set rs = Nothing

End Sub



Sub IMPRIMERETiva(ByVal c As Integer)
'imprime retencion de iva
'c es copias

Dim gf1 As String

Set rs = New adodb.Recordset
q = "select * from a5, a1, g1 where [num_int] = " & Me.numint & " and a5.[id_proveedor] = a1.[id_proveedor] and a5.[id_usuario] = g1.[id_usuario]"
rs.MaxRecords = 1
rs.Open q, cn1
If Not rs.BOF And Not rs.EOF Then
        descconcepto = "Sin detallar"
     
     gf1 = "Courier New"
     ip = Space$(9)
     ret = Space$(9)
     c2 = InputBox("Cantidad de Copias", "Imprimir Retencion Iva", c)
     If Val(c2) > 0 Then
      For h = 1 To Val(c2)
       Call definefont(gf1, "S", 12)
       Printer.Print
       Printer.Print "Comprobante Retencion IVA - R.G.(AFIP) 1575"
       Call definefont(gf1, "N", 12)
       Printer.Print "Regimen de Retencion del IMPUESTO AL VALOR AGREGADO"
       Call definefont(gf1, "S", 12)
       Printer.Print Spc(50); "Nro.: ";
       Call definefont(gf1, "N", 12)
       Printer.Print Spc(3); Format$(rs("SUCURSAL"), "0000") & "-" & Format$(rs("num_comprobante"), "00000000")
       Call definefont(gf1, "S", 12)
       Printer.Print Spc(50); "Fecha:";
       Call definefont(gf1, "N", 12)
       Printer.Print Spc(3); rs("fecha")
       Printer.Print
       Call definefont(gf1, "N", 12)
       Printer.Print "__________________________________________________________________________"
       Call definefont(gf1, "S", 12)
       Printer.Print "Agente de Retencion:"; glo.nombrecli
       Printer.Print "Direccion..........:"; glo.direccioncli; "   (2705)"
       Printer.Print "C.u.i.t............:"; glo.CUIT
       Call definefont(gf1, "N", 12)
       Printer.Print "__________________________________________________________________________"
       Printer.Print
       Printer.Print
       Printer.Print
       Call definefont(gf1, "N", 12)
       Printer.Print "__________________________________________________________________________"
 
       Call definefont(gf1, "S", 12)
       Printer.Print "Contribuyente: ";
       Call definefont(gf1, "N", 12)
       Printer.Print rs("denominacion"); "  ("; Format$(rs("a5.id_proveedor"), "00000"); ")    "
       Call definefont(gf1, "S", 12)
       Printer.Print "Direccion....: ";
       Call definefont(gf1, "N", 12)
       Printer.Print rs("direccion") & "  -" & rs("localidad") & "-   (" & rs("cp") & ")"
       Call definefont(gf1, "S", 12)
       Printer.Print "CUIT.........: ";
       Call definefont(gf1, "N", 12)
       Printer.Print rs("cuit")
       Printer.Print "__________________________________________________________________________"
       Printer.Print
       Printer.Print
       Call definefont(gf1, "N", 10)
       Printer.Print "Regimen: 966"
       Printer.Print
       RSet ip = Format$(rs("monto_suj_ret"), "####0.00")
       Printer.Print "Monto Suj.a Reten. : " & ip
       RSet ip = Format$(rs("alicuota_ret"), "####0.00")
       Printer.Print "Alicuota           : " & ip
       RSet ip = Format$(rs("total"), "####0.00")
       Printer.Print "RETENCION          : " & ip
       Printer.Print
       Printer.Print
       Call definefont(gf1, "S", 12)
       Printer.Print "Son pesos: ";
       Call definefont(gf1, "N", 12)
       Printer.Print convierte(rs("total"))
       Printer.Print
       Printer.Print "Ingresa en DDJJ periodo: "; rs("condiciones")
       Printer.Print
       Printer.Print
       Call definefont(gf1, "N", 8)
       Printer.Print "Orden de Pago"
       sop = Val(Mid$(Me.num_op, 1, 4))
       nop = Val(Mid$(Me.num_op, 6, 8))
       Set rs2 = New adodb.Recordset
       q = "select [fecha], [total] from a5 where [id_tipocomp] = 50 and [sucursal] = " & sop & " and [num_comprobante] = " & nop
       rs2.Open q, cn1
       If Not rs2.BOF And Not rs2.EOF Then
         Printer.Print "------------------------------------------------"
         Printer.Print "Fecha         Numero                  Importe      "
         Printer.Print "------------------------------------------------"
         Printer.Print Me.fecha & "  " & Me.observaciones & "        " & rs2("total")
       Else
         Printer.Print "------------------------------------------------"
         Printer.Print "Fecha         Numero                  Importe      "
         Printer.Print "------------------------------------------------"
         Printer.Print Me.fecha & "  " & Me.observaciones
       End If
       Set rs2 = Nothing
       Printer.PaintPicture gen_logo.Picture2.Picture, 7100, 10100
       
         Call definefont(gf1, "N", 12)
       Printer.Print
       Printer.Print
       Printer.Print
       Printer.Print
       
      
       Call definefont(gf1, "N", 9)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       
        
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.NewPage
     Next h
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
     Printer.EndDoc
  End If
End If
Set rs = Nothing

End Sub




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




Public Sub borrar()
'borra el comprobante previamente cargado en el objeto
J = MsgBox("Confirma Eliminar " & Me.abreviatura & " " & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000"), 4)
If J = 6 Then
 Select Case Me.idtipocomp
  Case Is = 50
    'borro op
     Call borroop
 
  Case Else
      Call borrocomp
 End Select
End If
End Sub
Sub borrocomp()
  If verificaperiodo(Me.fecha) = "A" Then

    'borro comprobantes
      cn1.BeginTrans
      If Me.STOCK <> "N" Then
        q = "select * from a6 where [num_int] = " & Me.numint & " and [id_producto] > 1 "
        Set rs = New adodb.Recordset
        rs.Open q, cn1
        While Not rs.EOF
          Set rs2 = New adodb.Recordset
          q = "select * from a2 where [id_producto] = " & rs("id_producto")
          rs2.Open q, cn1, adOpenDynamic, adLockOptimistic, 1
          If Not rs2.BOF And Not rs2.EOF Then
             If Me.STOCK = "E" Then
              s = -rs("cantidad")
             Else
              s = rs("cantidad")
             End If
             
             QUERY = "update a2 set [stock] = [stock] + " & s & " where [id_producto] = " & rs2("id_producto")
             cn1.Execute QUERY
            
          End If
          Set rs2 = Nothing
          rs.MoveNext
        Wend
        Set rs = Nothing
      End If
      
      QUERY = "DELETE FROM a5 WHERE [num_int] = " & Me.numint
      cn1.Execute QUERY
      
      QUERY = "DELETE FROM a6 WHERE [num_int] = " & Me.numint
      cn1.Execute QUERY
  
      QUERY = "DELETE FROM a16 WHERE [num_int_ret] = " & Me.numint
      cn1.Execute QUERY
  
      QUERY = "DELETE FROM a19 WHERE [num_int] = " & Me.numint
      cn1.Execute QUERY
        
      QUERY = "DELETE FROM a20 WHERE [num_int] = " & Me.numint
      cn1.Execute QUERY
  
      QUERY = "DELETE FROM a21 WHERE [num_int] = " & Me.numint
      cn1.Execute QUERY
  
      QUERY = "DELETE FROM a23 WHERE [num_int] = " & Me.numint
      cn1.Execute QUERY
  
  
       
      QUERY = "DELETE FROM stk_01 WHERE [num_mov_int] = " & Me.numint & " and [modulo] = 'C'"
      cn1.Execute QUERY
      
      QUERY = "DELETE FROM cyb_05 WHERE [num_mov_int] = " & Me.numint & " and [modulo] = 'C'"
      cn1.Execute QUERY



      Set rs2 = New adodb.Recordset
      q = "select * from pro_05 where [modulo] = 'C' and [num_int] = " & Me.numint
      rs2.Open q, cn1
      While Not rs2.EOF
          If Me.idtipocomp = 65 Then 'o.c.
              'modifico si es oc por sol mat
              QUERY = "update pro_04 set [total_oc] = [total_oc] - " & rs2("cantidad") & " where [num_referencia] = " & rs2("num_referencia") & " and [tipo04] = 1"
              cn1.Execute QUERY
            
              'borro si es oc directa
              QUERY = "delete * from pro_04 where [num_referencia] = " & rs2("num_referencia") & " and [tipo04] = 2"
              cn1.Execute QUERY
            
            Else 'fact./remito
              QUERY = "update pro_04 set [total_recibido] = [total_recibido] - " & rs2("cantidad") & " where [num_referencia] = " & rs2("num_referencia")
              cn1.Execute QUERY
          End If
        rs2.MoveNext
      Wend
      Set rs2 = Nothing
      
      QUERY = "DELETE FROM pro_05 WHERE [modulo] = 'C' and [num_int] = " & Me.numint
      cn1.Execute QUERY

      Call borracontabilidad(Me.numint, "C")
      
      If Me.idtipocomp = 95 Then 'ret gan. actu acumulador
         'pagos en acumulador ret. gan
         QUERY = "update ret_01 set  [ret_mes]= [ret_mes] - " & Me.total
         QUERY = QUERY & " where [id_proveedor]= " & Me.idproveedor & " and [id_retgan] = " & Me.idcodretgan & " and [año] = " & Val(Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 7, 4)) & " and [mes] = " & Val(Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 4, 2))
         cn1.Execute QUERY
      End If
      
      cn1.CommitTrans
 
     MsgBox ("Operacion Terminada")
     
  Else
       MsgBox ("El periodo para el cual se deseas borrar este comprobante esta CERRADO!!!!!")
  End If

       
 Exit Sub
borraerr:
MsgBox ("Imposible Eliminar Comprobante. Verifique sus permisos e intente mas tarde. Si aun asi no puede llame al serv. Tecnico(Clase: Comprobantes Prov. Modulo: Borra Comp.")

End Sub

Sub borroop()
    'borro op
      'On Error GoTo errbop
      cn1.BeginTrans
      
      QUERY = "DELETE FROM a5 WHERE [num_int] = " & Me.numint
      cn1.Execute QUERY
      
      QUERY = "DELETE FROM a7 WHERE [num_int] = " & Me.numint
      cn1.Execute QUERY
  
  
        'ch terc
        QUERY = "update cyb_03 set  [estado]='C', [tipo_salida]='C'"
        QUERY = QUERY & " where [num_int_op]= " & Me.numint
        cn1.Execute QUERY
      
        'ch propio
        QUERY = "update cyb_02 set  [estado]='P'"
        QUERY = QUERY & " where [num_int_op]= " & Me.numint
        cn1.Execute QUERY
      
      
      'mov. banarios
      QUERY = "DELETE FROM cyb_04 WHERE [modulo] = 'C' and [num_mov_int] = " & Me.numint
      cn1.Execute QUERY
      
      'mov. caja
      QUERY = "DELETE FROM cyb_05 WHERE [modulo] = 'C' and [num_mov_int] = " & Me.numint
      cn1.Execute QUERY
      
      Call borracontabilidad(Me.numint, "C")
           
      nop = Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000")
      
            
      Set rs1 = New adodb.Recordset
      q = "select * from a15 where [num_int_op] = " & Me.numint
      rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
      While Not rs1.EOF
         QUERY = "update A5 set [estado_pago]='N',[saldo_impago]=[saldo_impago] + " & Abs(rs1("importe_pagado")) & " , [num_op]= '0000-00000000'"
         QUERY = QUERY & " where [num_int]= " & rs1("num_int_comp")
         cn1.Execute QUERY
         rs1.Delete
         rs1.MoveNext
      Wend
      Set rs1 = Nothing
      
      
      
      'pagos en acumulador ret. gan
      Set rs1 = New adodb.Recordset
      q = "select * from ret_01 where [id_proveedor]= " & Me.idproveedor & " and [id_retgan] = " & Me.idcodretgan & " and [año] = " & Val(Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 7, 4)) & " and [mes] = " & Val(Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 4, 2))
      rs1.Open q, cn1
      If Not rs1.EOF And Not rs1.BOF Then
         pm = rs1("pagos_mes") - Me.subtotal
         QUERY = "update ret_01 set  [pagos_mes]=" & pm
         QUERY = QUERY & " where [id_proveedor]= " & Me.idproveedor & " and [id_retgan] = " & Me.idcodretgan & " and [año] = " & Val(Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 7, 4)) & " and [mes] = " & Val(Mid$(Format$(Me.fecha, "dd/mm/yyyy"), 4, 2))
         cn1.Execute QUERY
      
      End If
      Set rs1 = Nothing
         
         
      cn1.CommitTrans
 
     MsgBox ("Operacion Terminada")
       
Exit Sub
errbop:
MsgBox ("Imposible Eliminar Comprobante. Verifique sus permisos e intente mas tarde. Si aun asi no puede llame al serv. Tecnico(Clase: Comprobantes Prov. Modulo: Borra OP")

End Sub
Public Property Let compra(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.compra = 5
    mvarcompra = vData
End Property


Public Property Get compra() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.compra
    compra = mvarcompra
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



Public Sub actual(ByVal TIPOCOMP As Integer)
'SACA LOS DATOS DE UN COMPROBANTE QUE SE ESTA POR CREAR
'STOCK, CTACTE
Set rs = New adodb.Recordset
q = "select [stock], [ctacte], [iva], [abreviatura], [compra], [contabilidad] from g2 where [id_tipo_comp] = " & TIPOCOMP
rs.MaxRecords = 1
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
     Me.STOCK = rs("STOCK")
     Me.ctacte = rs("CTACTE")
     Me.grabado = rs("iva")
     Me.abreviatura = rs("abreviatura")
     Me.compra = rs("compra")
     Me.contabilidad = rs("contabilidad")
     
Else
  MsgBox ("Error al Inicializar Comprobante. Clase: Comprobantes  Proc: Actual")
  Me.STOCK = "N"
  Me.ctacte = "N"
  Me.grabado = "N"
  Me.abreviatura = "Err"
  Me.compra = "N"
  Me.contabilidad = "N"
End If
Set rs = Nothing

End Sub

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



Public Property Let STOCK(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.STOCK = 5
    mvarSTOCK = vData
End Property

Sub imprimeop1(ByVal c As Integer)
'generica
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
   Printer.Print "ORDEN DE PAGO"
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
   Printer.Print Tab(3); "Proveedor:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
   Set cl_prov = New proveedores
   cl_prov.carga (Me.idproveedor)
   If cl_prov.idprov > 0 Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print "  " & cl_prov.razonsocial; Tab(92);
   Else
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print "  "; Tab(92);
   End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "Emitida por:";
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
   Printer.Print "  " & cl_prov.direccion
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
   Printer.Print "  " & cl_prov.localidad
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
   Printer.Print "  " & cl_prov.CUIT

'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "___________________________________________________________________________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "Recibí/mos de " & glo.nombrecli
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "La suma de pesos: *****" & convierte(Me.total) & "*****"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "___________________________________________________________________________________________________"

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
   Printer.Print Tab(10); "----------------------------------------------------------------------------------------------"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "Forma de Pago       Num.Ch.    Detalle/Banco               Fecha Dif.    Importe  Sucursal"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "----------------------------------------------------------------------------------------------"
   Set rs = New adodb.Recordset
   q = "Select * from a7 where [num_int] = " & Me.numint
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
   Printer.Print Tab(10); "--------------------------------------------------------------"
   Printer.Print Tab(10); "Fecha        Comprobante            Importe     Importe"
   Printer.Print Tab(10); "                                    a Cancelar  Cancelado"
   Printer.Print Tab(10); "--------------------------------------------------------------"
   ia = Space$(10)
   Set rs = New adodb.Recordset
   q = "Select * from a15 where [num_int_op] = " & Me.numint
   rs.Open q, cn1
   While Not rs.EOF
    Set rs2 = New adodb.Recordset
    q = "select [id_tipocomp], [fecha], [letra], [sucursal], [num_comprobante], [total] from a5 where [num_int] = " & rs("num_int_comp")
    rs2.Open q, cn1
    If Not rs2.EOF And Not rs2.BOF Then
        nc = "[" & Format$(rs2("id_tipocomp"), "000") & "] " & rs2("letra") & Format$(rs2("sucursal"), "0000") & "-" & Format$(rs2("num_comprobante"), "00000000")
        fd = Format$(rs2("fecha"), "dd/mm/yyyy")
        RSet i = Format$(rs("saldo_comprobante"), "######0.00")
        RSet ia = Format$(rs("importe_pagado"), "######0.00")
        Printer.Print Tab(10); fd & "  " & nc & "  " & i & "  " & ia
        totl = totl + 1
    End If
    Set rs2 = Nothing
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
   Printer.Print "Cotizacion U$s: ";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.Print Format$(Me.cotizacion, "###0.000")
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print
   Printer.Print Tab(3); "Observaciones:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.Print "  "; Me.observaciones
   Printer.Print
   Printer.Print
   Printer.Print
   Printer.Print Tab(75); "______________________________"
   Printer.Print Tab(75); "   por "; cl_prov.razonsocial
   
   Set cl_prov = Nothing
   Set cl_comp = Nothing
   If Y < Val(l) Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.NewPage
   End If
  Next Y
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
  Printer.EndDoc
End If
End Sub


Sub imprimeop5(ByVal c As Integer)
'fabrimac
'igual que la 1(generica) pero sin cotizacon dolar

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
   Printer.Print "ORDEN DE PAGO"
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
   Printer.Print Tab(3); "Proveedor:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
   Set cl_prov = New proveedores
   cl_prov.carga (Me.idproveedor)
   If cl_prov.idprov > 0 Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print "  " & cl_prov.razonsocial; Tab(92);
   Else
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print "  "; Tab(92);
   End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "Emitida por:";
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
   Printer.Print "  " & cl_prov.direccion
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
   Printer.Print "  " & cl_prov.localidad
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
   Printer.Print "  " & cl_prov.CUIT

'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "___________________________________________________________________________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "Recibí/mos de " & glo.nombrecli
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "La suma de pesos: *****" & convierte(Me.total) & "*****"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "___________________________________________________________________________________________________"

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
   Printer.Print Tab(10); "----------------------------------------------------------------------------------------------"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "Forma de Pago       Num.Ch.    Detalle/Banco               Fecha Dif.    Importe  Sucursal"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "----------------------------------------------------------------------------------------------"
   Set rs = New adodb.Recordset
   q = "Select * from a7 where [num_int] = " & Me.numint
   rs.Open q, cn1
   i = Space$(10)
   linea = 1
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
   Printer.Print Tab(10); "--------------------------------------------------------------"
   Printer.Print Tab(10); "Fecha        Comprobante            Importe     Importe"
   Printer.Print Tab(10); "                                    a Cancelar  Cancelado"
   Printer.Print Tab(10); "--------------------------------------------------------------"
   ia = Space$(10)
   Set rs = New adodb.Recordset
   q = "Select * from a15 where [num_int_op] = " & Me.numint
   rs.Open q, cn1
   While Not rs.EOF
    Set rs2 = New adodb.Recordset
    q = "select [id_tipocomp], [fecha], [letra], [sucursal], [num_comprobante], [total] from a5 where [num_int] = " & rs("num_int_comp")
    rs2.Open q, cn1
    If Not rs2.EOF And Not rs2.BOF Then
        nc = "[" & Format$(rs2("id_tipocomp"), "000") & "] " & rs2("letra") & Format$(rs2("sucursal"), "0000") & "-" & Format$(rs2("num_comprobante"), "00000000")
        fd = Format$(rs2("fecha"), "dd/mm/yyyy")
        RSet i = Format$(rs("saldo_comprobante"), "######0.00")
        RSet ia = Format$(rs("importe_pagado"), "######0.00")
        Printer.Print Tab(10); fd & "  " & nc & "  " & i & "  " & ia
        totl = totl + 1
        linea = linea + 1
    End If
    Set rs2 = Nothing
    rs.MoveNext
   Wend
   Set rs = Nothing
   
   
   
   
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Times New Roman"
   Printer.FontSize = 10
   Printer.Print
   Printer.FontUnderline = True
   Printer.FontBold = True
   Printer.Print
   Printer.Print Tab(3); "Observaciones:";
   Printer.FontUnderline = False
   Printer.FontBold = False
   Printer.Print "  "; Me.observaciones
   Printer.Print
   Printer.PaintPicture gen_logo.Picture2.Picture, 700, 10100 + (linea * 100)
   Printer.Print
   Printer.Print
   Printer.Print Tab(75); "______________________________"
   Printer.Print Tab(75); "   por "; cl_prov.razonsocial
   
   Set cl_prov = Nothing
   Set cl_comp = Nothing
   If Y < Val(l) Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.NewPage
   End If
  Next Y
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
  Printer.EndDoc
End If
End Sub




Sub cab1()
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Times New Roman"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 12
   'Printer.PaintPicture gen_logo.Picture1.Picture, 0, 0
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
   Printer.Line Step(800, 1)-Step(4000, 1)
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
   Printer.Print Tab(12); "ORDEN DE PAGO"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Line Step(800, 1)-Step(4000, 1)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
   Printer.Print
   Printer.Print Tab(80); "Nro. Comprobante:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FillStyle = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000")

End Sub
Sub CAB3()
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
   Printer.Line Step(800, 1)-Step(4000, 1)
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
   Printer.Print Tab(12); "ORDEN DE PAGO"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Line Step(800, 1)-Step(4000, 1)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
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

End Sub
Sub imprimeop2(ByVal c As Integer)
'degser
l = InputBox("Indique cantidad de Copias", , c)
If Val(l) > 0 And Val(l) <= 5 Then
  For Y = 1 To Val(l)
   Call cab1
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
   Printer.FontName = "courier new"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(8); "Proveedor.......:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
   Set cl_prov = New proveedores
   cl_prov.carga (Me.idproveedor)
   If cl_prov.idprov > 0 Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print "  " & cl_prov.razonsocial
   Else
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print "  "
   End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(8); "Direccion.......:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & cl_prov.direccion
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(8); "Localidad.......:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & cl_prov.localidad
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(8); "Cuit............:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & cl_prov.CUIT
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Times New Roman"
  
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "___________________________________________________________________________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "Recibí/mos de " & glo.nombrecli
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
   Printer.Print Tab(10); "----------------------------------------------------------------------------------------------"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "Forma de Pago       Num.Ch.    Detalle/Banco               Fecha Dif.    Importe  Sucursal"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "----------------------------------------------------------------------------------------------"

   Set rs = New adodb.Recordset
   q = "Select * from a7 where [num_int] = " & Me.numint
   rs.Open q, cn1
   i = Space$(10)
   totl = 0
   While Not rs.EOF
    ifp = Format$(rs("id_formapago"), "000")
    FP = Format$(Left$(rs("formapago"), 12), "@@@@@@@@@@@@!")
    nch = Format$(rs("num_ch"), "0000000000")
    db = Format$(Left$(rs("detalle_Banco") & " ", 27), "@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    s = Format$(Left$(rs("sucursal") & " ", 10), "@@@@@@@@@@!")
    fd = Format$(rs("fecha_dif"), "dd/mm/yyyy")
    RSet i = Format$(rs("importe"), "######0.00")
    If totl > 15 Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.NewPage
      Call cab1
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
      Printer.Print Tab(10); "----------------------------------------------------------------------------------------------"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print Tab(10); "Forma de Pago       Num.Ch.    Detalle/Banco               Fecha Dif.    Importe  Sucursal"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print Tab(10); "----------------------------------------------------------------------------------------------"
      totl = -35
    End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(10); "(" & ifp & ") " & FP & "  " & nch & " " & db & " " & fd & " " & i & " " & s
    totl = totl + 1
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
   Printer.Print Tab(10); "-------------------------------------------------------"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "Fecha        Comprobante               Importe"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "-------------------------------------------------------"
   Set rs = New adodb.Recordset
   q = "Select * from a5 where [estado_pago] = 'P' and [num_op] = '" & Format$(Me.sucursal, "0000") & "-" & Format$(Me.numcomp, "00000000") & "'"
   rs.Open q, cn1
   While Not rs.EOF
    nc = "[" & Format$(rs("id_tipocomp"), "000") & "] " & rs("letra") & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comprobante"), "00000000")
    fd = Format$(rs("fecha"), "dd/mm/yyyy")
    RSet i = Format$(rs("total"), "######0.00")
    If totl > 15 Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
     Printer.NewPage
     Call cab1
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
     Printer.Print Tab(10); "-------------------------------------------------------"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
     Printer.Print Tab(10); "Fecha        Comprobante               Importe"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
     Printer.Print Tab(10); "-------------------------------------------------------"
     totl = -35
    End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(10); fd & "  " & nc & "  " & i
    totl = totl + 1
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
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(9); "Observaciones:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  "; Me.observaciones
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(75); "______________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(75); "   por "; Left$(cl_prov.razonsocial, 30)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 6
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "___________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "  Emitido por: "; para.usuario

   Set cl_prov = Nothing
   Set cl_comp = Nothing
   If Y < Val(l) Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.NewPage
   End If
  Next Y
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
  Printer.EndDoc
End If
End Sub

Sub imprimeop3(ByVal c As Integer)
'framahi
l = InputBox("Indique cantidad de Copias", , c)
If Val(l) > 0 And Val(l) <= 5 Then
  
  For Y = 1 To Val(l)
   Call CAB3
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
   Printer.FontName = "courier new"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(8); "Proveedor.......:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
   Set cl_prov = New proveedores
   cl_prov.carga (Me.idproveedor)
   If cl_prov.idprov > 0 Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print "  " & cl_prov.razonsocial
   Else
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print "  "
   End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(8); "Direccion.......:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & cl_prov.direccion
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(8); "Localidad.......:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & cl_prov.localidad
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(8); "Cuit............:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & cl_prov.CUIT
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Times New Roman"
  
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "___________________________________________________________________________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 12
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "Recibí/mos de " & glo.nombrecli
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
   Printer.Print Tab(10); "----------------------------------------------------------------------------------------------"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "Forma de Pago       Num.Ch.    Detalle/Banco               Fecha Dif.    Importe  Sucursal"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(10); "----------------------------------------------------------------------------------------------"

   Set rs = New adodb.Recordset
   q = "Select * from a7 where [num_int] = " & Me.numint
   rs.Open q, cn1
   i = Space$(10)
   totl = 0
   While Not rs.EOF
    ifp = Format$(rs("id_formapago"), "000")
    FP = Format$(Left$(rs("formapago"), 12), "@@@@@@@@@@@@!")
    nch = Format$(rs("num_ch"), "0000000000")
    db = Format$(Left$(rs("detalle_Banco") & " ", 27), "@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    s = Format$(Left$(rs("sucursal") & " ", 10), "@@@@@@@@@@!")
    fd = Format$(rs("fecha_dif"), "dd/mm/yyyy")
    RSet i = Format$(rs("importe"), "######0.00")
    If totl > 15 Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.NewPage
      Call cab1
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
      Printer.Print Tab(10); "----------------------------------------------------------------------------------------------"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print Tab(10); "Forma de Pago       Num.Ch.    Detalle/Banco               Fecha Dif.    Importe  Sucursal"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print Tab(10); "----------------------------------------------------------------------------------------------"
      totl = -35
    End If
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.Print Tab(10); "(" & ifp & ") " & FP & "  " & nch & " " & db & " " & fd & " " & i & " " & s
    totl = totl + 1
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
   Printer.Print Tab(10); "--------------------------------------------------------------"
   Printer.Print Tab(10); "Fecha        Comprobante            Importe     Importe"
   Printer.Print Tab(10); "                                    Comprob.    Cancelado"
   Printer.Print Tab(10); "--------------------------------------------------------------"
   ia = Space$(10)
   Set rs = New adodb.Recordset
   q = "Select * from a15 where [num_int_op] = " & Me.numint
   rs.Open q, cn1
   While Not rs.EOF
    Set rs2 = New adodb.Recordset
    q = "select [id_tipocomp], [fecha], [letra], [sucursal], [num_comprobante], [total] from a5 where [num_int] = " & rs("num_int_comp")
    rs2.Open q, cn1
    If Not rs2.EOF And Not rs2.BOF Then
        nc = "[" & Format$(rs2("id_tipocomp"), "000") & "] " & rs2("letra") & Format$(rs2("sucursal"), "0000") & "-" & Format$(rs2("num_comprobante"), "00000000")
        fd = Format$(rs2("fecha"), "dd/mm/yyyy")
        RSet i = Format$(rs2("total"), "######0.00")
        RSet ia = Format$(rs("importe_pagado"), "######0.00")
        If totl > 15 Then
            Printer.NewPage
            Call CAB3
            Printer.FontSize = 10
            Printer.Print
            Printer.Print
            Printer.FontBold = True
            Printer.Print Tab(10); "Comprobantes Aplicados"
            Printer.FontBold = False
            Printer.FontName = "Courier new"
            Printer.FontSize = 8
            Printer.Print Tab(10); "--------------------------------------------------------------"
            Printer.Print Tab(10); "Fecha        Comprobante            Importe     Importe"
            Printer.Print Tab(10); "                                    Comprob.    Cancelado"
            Printer.Print Tab(10); "--------------------------------------------------------------"
            totl = -35
        End If
        Printer.Print Tab(10); fd & "  " & nc & "  " & i & "  " & ia
        totl = totl + 1
    End If
    Set rs2 = Nothing
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
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(9); "Observaciones:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  "; Me.observaciones
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
  
   
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(75); "______________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(75); "   por "; Left$(cl_prov.razonsocial, 30)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 6
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "___________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(12); "  Emitido por: "; para.usuario

   Set cl_prov = Nothing
   Set cl_comp = Nothing
   If Y < Val(l) Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.NewPage
   End If
  Next Y
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
  Printer.EndDoc
  
End If
End Sub



Public Property Get STOCK() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.STOCK
    STOCK = mvarSTOCK
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



Sub cargar(ByVal TIPOCOMP As Integer, ByVal letra As String, ByVal sucursal As Integer, ByVal numcomp As Long, ByVal idprov As Long)
Set rs = New adodb.Recordset
q = "select * from a5 where [sucursal] = " & sucursal & " and [num_comprobante] = " & numcomp & " and [letra] = '" & letra & "' and [id_tipocomp] = " & TIPOCOMP
Select Case TIPOCOMP
  Case Is = 65, Is = 100
  
  Case Else
    q = q & " and [id_proveedor] = " & idprov
End Select
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
   Me.estado = rs("estado")
   Me.fecha = rs("fecha")
   Me.fechaprobentrega = rs("fecha_prob_entrega")
   Me.fecharecepcion = rs("fecha_recepcion")
   Me.idproveedor = rs("id_proveedor")
   Me.idtipocomp = rs("id_tipocomp")
   Me.idusuario = rs("id_usuario")
   Me.iva = rs("iva")
   Me.letra = rs("letra")
   Me.nograbado = rs("no_grabado")
   Me.numcomp = rs("num_comprobante")
   Me.numint = rs("num_int")
   Me.percep = rs("percep_ret")
   Me.subtotal = rs("subtotal")
   Me.sucursal = rs("sucursal")
   Me.total = rs("total")
   Me.total_om = rs("total_d")
   Me.idcuenta = rs("id_cuenta")
   Me.idcodretgan = rs("id_codretgan")
   Me.moneda = rs("moneda")
   Me.cotizacion = rs("cotiz_dolar")
   Me.fechavto = rs("fecha_vto")
   Me.saldoimpago = rs("saldo_impago")
   Me.estado_pago = rs("estado_pago")
   Me.contado = rs("contado")
   Me.zona = rs("zona")
   Me.idcliente = rs("id_cliente")
Else
   Me.numint = 0
End If
Set rs = Nothing
End Sub

'FIXIT: Declare 'numint' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Sub cargar2(ByVal numint)
Set rs = New adodb.Recordset
q = "select * from a5 where [num_int] = " & numint
rs.MaxRecords = 1
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
   Me.estado = rs("estado")
   Me.fecha = rs("fecha")
   Me.fechaprobentrega = rs("fecha_prob_entrega")
   Me.fecharecepcion = rs("fecha_recepcion")
   Me.idproveedor = rs("id_proveedor")
   Me.idtipocomp = rs("id_tipocomp")
   Me.idusuario = rs("id_usuario")
   Me.iva = rs("iva")
   Me.letra = rs("letra")
   Me.nograbado = rs("no_grabado")
   Me.numcomp = rs("num_comprobante")
   Me.numint = rs("num_int")
   Me.percep = rs("percep_ret")
   Me.subtotal = rs("subtotal")
   Me.sucursal = rs("sucursal")
   Me.total = rs("total")
   Me.total_om = rs("total_d")
   Me.moneda = rs("moneda")
   If Not IsNull(rs("id_cuenta")) Then
     Me.idcuenta = rs("id_cuenta")
   Else
     Me.idcuenta = 0
   End If
   Me.idcodretgan = rs("id_codretgan")
   Me.STOCK = rs("stock")
   Me.ctacte = rs("ctacte")
   Me.grabado = rs("grabado")
   Me.sucursal = rs("sucursal")
   Me.numcomp = rs("num_comprobante")
   Me.observaciones = rs("obs")
   Me.contacto = rs("info_contacto")
   Me.condiciones = rs("condiciones")
   Me.cotizacion = rs("cotiz_dolar")
   Me.saldoimpago = rs("saldo_impago")
Me.idcliente = rs("id_cliente")
   If Not IsNull(rs("fecha_vto")) Then
     Me.fechavto = rs("fecha_vto")
   Else
     Me.fechavto = Format$(Now, "dd/mm/yyyy")
   End If
   If Not IsNull(rs("num_op")) Then
     Me.num_op = rs("num_op")
   Else
     Me.num_op = "0000-00000000"
   End If
   Me.estado_pago = rs("estado_pago")
   Me.contado = rs("contado")
   Me.zona = rs("zona")
Else
   Me.numint = 0
End If
Set rs = Nothing
End Sub
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



Public Property Let fecharecepcion(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.fecharecepcion = 5
    mvarfecharecepcion = vData
End Property


Public Property Get fecharecepcion() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.fecharecepcion
    fecharecepcion = mvarfecharecepcion
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



Public Property Let fechaprobentrega(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.fechaprobentrega = 5
    mvarfechaprobentrega = vData
End Property


Public Property Get fechaprobentrega() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.fechaprobentrega
    fechaprobentrega = mvarfechaprobentrega
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



