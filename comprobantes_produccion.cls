VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "comprobantes_produccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
'variables locales para almacenar los valores de las propiedades
Private mvarnumint As Long 'copia local
Private mvarsucursal As Integer 'copia local
Private mvarnumcomp As Long 'copia local
Private mvaridtipocomp As Integer 'copia local
Private mvarfecha As Date 'copia local
Private mvarestado As String 'copia local
Private mvaridusuario As Long 'copia local
Private mvarobservaciones As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarid_obra As Long 'copia local
Private mvarfechaesperado As Date 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarnumreferencia As Long 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarabreviatura As String 'copia local
Private mvarcopias As Integer 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarobra As String 'copia local
Private mvarusuario As String 'copia local
Private mvarobservacion_obra As String 'copia local
Private mvardesc_comprobante As String 'copia local
Public Property Let desc_comprobante(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.desc_comprobante = 5
    mvardesc_comprobante = vData
End Property


Public Property Get desc_comprobante() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.desc_comprobante
    desc_comprobante = mvardesc_comprobante
End Property



Public Property Let observacion_obra(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.observacion_obra = 5
    mvarobservacion_obra = vData
End Property


Public Property Get observacion_obra() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.observacion_obra
    observacion_obra = mvarobservacion_obra
End Property



Public Property Let usuario(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.usuario = 5
    mvarusuario = vData
End Property


Public Property Get usuario() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.usuario
    usuario = mvarusuario
End Property



Public Property Let obra(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.obra = 5
    mvarobra = vData
End Property


Public Property Get obra() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.obra
    obra = mvarobra
End Property



Public Property Let copias(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.copias = 5
    mvarcopias = vData
End Property


Public Property Get copias() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.copias
    copias = mvarcopias
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



Public Sub sacaultimonumero(ByVal tc As Integer)
'saca el ultimo numero de comprobante y lo pone
'en la propiedad numcomp.
'por parametro se le pasa el tipo de comprobante

On Error GoTo e1
Set rs = New ADODB.Recordset
q = "select * from pro_03 where [id_tipocomp] = " & tc
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
If Not rs.EOF And Not rs.BOF Then
  Me.numcomp = rs("ult_numero") + 1
  rs("ult_numero") = Me.numcomp
  rs.Update
Else
  Me.numcomp = 0
End If

Exit Sub
e1:
MsgBox ("Error Clase:Comprobantes Produccion  Metodo:Saca ultimo numero")
Me.numcomp = 0
Exit Sub
End Sub

Public Property Let numreferencia(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.numreferencia = 5
    mvarnumreferencia = vData
End Property


Public Property Get numreferencia() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.numreferencia
    numreferencia = mvarnumreferencia
End Property



Public Property Let fechaesperado(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.fechaesperado = 5
    mvarfechaesperado = vData
End Property


Public Property Get fechaesperado() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.fechaesperado
    fechaesperado = mvarfechaesperado
End Property



Public Property Let id_obra(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.id_obra = 5
    mvarid_obra = vData
End Property


Public Property Get id_obra() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.id_obra
    id_obra = mvarid_obra
End Property




Public Sub imprimir()
'previamente es necesario cargar el comprobante
Select Case Me.idtipocomp
  Case Is = 1
    J = InputBox$("Cantidad de Copias", "Imprimir", Me.copias)
    If Val(J) > 0 Then
     Call imprimesm(Val(J))
    End If
    
    J = InputBox$("Informe Recepcion Mercaderias. Cantidad de Copias", "Imprimir", Me.copias)
    If Val(J) > 0 Then
     Call imprimesmt(Val(J))
    End If
    
    
  Case Is = 2
   J = InputBox$("Cantidad de Copias", "Imprimir", Me.copias)
    If Val(J) > 0 Then
     Call imprimeminuta(Val(J))
    End If
      
      
End Select

End Sub
Sub imprimesm(ByVal c As Integer)
For Y = 1 To c
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
   Printer.FontSize = 16
   Printer.Print Tab(5); "CONTROL OFICINA TECNICA MATERIALES SOLICITADOS"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
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
   Printer.Print Tab(3); "Obra.......................:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & Me.obra; Tab(92);
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
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & Me.usuario
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(3); "Descripcion Obra....:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & Me.observacion_obra
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Courier new"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 8
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "____________________________________________________________________________________________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "        Ref.    Detalle                                        Observaciones              Cant.   Unid.   Para el:"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "____________________________________________________________________________________________________________________"
   Set rs = New ADODB.Recordset
   q = "select * from pro_02 where [num_int] = " & Me.numint
   rs.Open q, cn1
   r = 1
   cant = Space$(10)
   pu = Space$(10)
   While Not rs.EOF
       renglon = Format$(r, "000")
       basico = Format$(rs("num_referencia"), "00000")
       Detalle = Format$(Left$(rs("descripcion"), 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
       obsp = Format$(Left$(rs("observaciones"), 22), "@@@@@@@@@@@@@@@@@@@@@@!")
       RSet cant = Format$(rs("cantidad"), "######0.00")
       F = Format$(rs("fecha_esperado"), "dd/mm/yyyy")
       u = Format$(rs("unidad") & " ", "@@@@@@!")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "   " & renglon & " " & basico & " " & Detalle & " " & obsp & " " & cant & " " & u & F
       r = r + 1
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
   For X = 1 To 35 - r
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print
   Next X
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "___________________________________________________________________________________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(3); "Observaciones:";
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
   Printer.FontName = "Courier new"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 9
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(20); "_________________________"; Tab(75); "_________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(27); "Solicitó "; Tab(82); "Autorizó "
   
   Set cl_prov = Nothing
   Set cl_comp = Nothing
   If Y < c Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.NewPage
   End If
 Next Y
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.EndDoc

End Sub

Sub imprimesmt(ByVal c As Integer)
For Y = 1 To c
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
  Printer.FontSize = 16
   Printer.Print Tab(10); "CONTROL DE RECEPCION DE MERCADERIAS"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
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
   Printer.Print Tab(3); "Obra.......................:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & Me.obra; Tab(92);
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
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & Me.usuario
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(3); "Descripcion Obra....:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & Me.observacion_obra
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Courier new"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 8
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "____________________________________________________________________________________________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "        Ref.  Detalle                                        Observaciones              Cant.   Unid.   Para el:"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "____________________________________________________________________________________________________________________"
   Set rs = New ADODB.Recordset
   q = "select * from pro_02 where [num_int] = " & Me.numint
   rs.Open q, cn1
   r = 1
   cant = Space$(10)
   pu = Space$(10)
   While Not rs.EOF
       renglon = Format$(r, "000")
       basico = Format$(rs("num_referencia"), "00000")
       Detalle = Format$(Left$(rs("descripcion"), 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
       obsp = Format$(Left$(rs("observaciones"), 22), "@@@@@@@@@@@@@@@@@@@@@@!")
       RSet cant = Format$(rs("cantidad"), "######0.00")
       F = Format$(rs("fecha_esperado"), "dd/mm/yyyy")
       u = Format$(rs("unidad") & " ", "@@@@@@!")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "   " & renglon & " " & basico & " " & Detalle & " " & obsp & " " & cant & " " & u & F
       r = r + 1
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
   For X = 1 To 35 - r
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print
   Next X
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "___________________________________________________________________________________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(3); "Observaciones:";
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
   Printer.FontName = "Courier new"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 9
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(20); "_________________________"; Tab(75); "_________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(27); "Solicitó "; Tab(82); "Autorizó "
   
   Set cl_prov = Nothing
   Set cl_comp = Nothing
   If Y < c Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.NewPage
   End If
 Next Y
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.EndDoc

End Sub


Sub imprimeminuta(ByVal c As Integer)
Dim tipoop As String

Set rs4 = New ADODB.Recordset
q = "select [obs] from pro_05 where [num_int] = " & Me.numint & " and [modulo] = 'P'"
rs4.Open q, cn1
If Not rs4.EOF And Not rs4.BOF Then
  tipoop = rs4("obs")
Else
  tipoop = ""
End If
Set rs4 = Nothing


For Y = 1 To c
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
   Printer.Print Tab(10); "MINUTA para MODIFICAR REGISTRO de MATERIALES PEDIDOS"
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
   Printer.Print Tab(3); "Obra.......................:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & Me.obra; Tab(92);
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
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & Me.usuario
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(3); "Descripcion Obra....:";
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "  " & Me.observacion_obra
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Courier new"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 8
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "____________________________________________________________________________________________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "        Basico  Detalle                                        Observaciones              Cant.   "
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "____________________________________________________________________________________________________________________"
   Set rs = New ADODB.Recordset
   q = "select * from pro_02 where [num_int] = " & Me.numint
   rs.Open q, cn1
   r = 1
   cant = Space$(10)
   pu = Space$(10)
   While Not rs.EOF
       renglon = Format$(r, "000")
       basico = Format$(rs("id_producto"), "00000")
       Detalle = Format$(Left$(rs("descripcion"), 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
       obsp = Format$(Left$(rs("observaciones"), 22), "@@@@@@@@@@@@@@@@@@@@@@!")
       RSet cant = Format$(rs("cantidad"), "######0.00")
      ': El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "   " & renglon & " " & basico & " " & Detalle & " " & obsp & " " & cant
       r = r + 1
       rs.MoveNext
   Wend
   Set rs = Nothing
   Printer.Print
   Printer.Print Tab(20); "******************************************************************"
   Printer.Print Tab(20); "    ", tipoop
   Printer.Print Tab(20); "******************************************************************"
   
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontName = "Times New Roman"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontUnderline = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontBold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 10
   For X = 1 To 3
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.Print
   Next X
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print "___________________________________________________________________________________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(3); "Observaciones:";
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
   Printer.FontName = "Courier new"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.FontSize = 9
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(20); "_________________________"; Tab(75); "_________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print Tab(27); "Solicitó "; Tab(82); "Autorizó "
   Printer.Print
   
   
   Set cl_prov = Nothing
   Set cl_comp = Nothing
   If Y < c Then
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
      Printer.NewPage
   End If
 Next Y
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.EndDoc

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




Sub borrocomp()
 'borro COMPROBANTES
 borra = 0
 Set rs = New ADODB.Recordset
 q = "select * from pro_04, pro_05 where [modulo] = 'P' and [num_int] = " & Me.numint & " and pro_05.[num_referencia] = pro_04.[num_referencia]"
 rs.Open q, cn1
 While Not rs.EOF
   If rs("total_oc") > 0 Or rs("total_recibido") > 0 Then
       borra = 1
   End If
   rs.MoveNext
 Wend
 Set rs = Nothing
 If borra = 0 Then
    If Me.idusuario = para.id_usuario Then
      cn1.BeginTrans
      
      QUERY = "DELETE FROM PRO_01 WHERE [num_int] = " & Me.numint
      cn1.Execute QUERY
      
      QUERY = "DELETE FROM PRO_02 WHERE [num_int] = " & Me.numint
      cn1.Execute QUERY
  
      Set rs = New ADODB.Recordset
      q = "select * from pro_05 WHERE [modulo] = 'P' and [num_int] = " & Me.numint
      rs.Open q, cn1
      While Not rs.EOF
         QUERY = "DELETE FROM PRO_04 WHERE [num_referencia] = " & rs("num_referencia")
         cn1.Execute QUERY
         rs.MoveNext
      Wend
      Set rs = Nothing
      
      QUERY = "DELETE FROM PRO_05 WHERE [modulo] = 'P' and [num_int] = " & Me.numint
      cn1.Execute QUERY
  
      cn1.CommitTrans
    Else
      MsgBox ("Solo el usuario que ingreso el pedido puede eliminarlo")
    End If
  Else
    MsgBox ("El pedido tiene O.C. emitidas, no puede eliminarlo")
  End If
 
 Exit Sub
borraerr:
MsgBox ("Imposible Eliminar Comprobante. Verifique sus permisos e intente mas tarde. Si aun asi no puede llame al serv. Tecnico(Clase: Comprobantes Produccion. Modulo: Borrocomp")

End Sub


Sub borraproductorequisicion(ByVal numref As Long)
  q = "select * from pro_05 where [num_referencia] = " & numref & " and [modulo] <> 'P'"
  Set rsc = New ADODB.Recordset
  rsc.Open q, cn1
  borra = 0
  While Not rsc.EOF
    borra = borra + 1
    rsc.MoveNext
  Wend
  Set rsc = Nothing
 
  If borra <= 1 Then
    q = "select * from pro_04 where [num_referencia] = " & numref
    Set rsc = New ADODB.Recordset
    rsc.Open q, cn1, adOpenDynamic, adLockOptimistic
    While Not rsc.EOF
     If rsc("id_Usuario") = para.id_usuario Or para.id_grupo_modulo_compras >= 8 Then
        q = "select * from pro_05 where [num_referencia] = " & numref
        Set rsc2 = New ADODB.Recordset
        rsc2.Open q, cn1, adOpenDynamic, adLockOptimistic
        If Not rsc2.EOF And Not rsc2.BOF Then
            rsc2.Delete
        End If
        Set rsc2 = Nothing
        
       rsc.Delete
       rsc.MoveNext
     
     Else
       MsgBox ("Solo puede eliminar el producto de la requsicion el usuario que lo ingreso o ser Administrador de Compras")
       rsc.MoveNext
     End If
    Wend
    Set rsc = Nothing
 Else
   MsgBox ("Imposible Borrar producto de requisicion porque tiene movimientos asociados")
 End If
End Sub







'FIXIT: Declare 'numint' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Sub cargar2(ByVal numint)
Set rs = New ADODB.Recordset
q = "select * from pro_01, pro_03, g1, a4 where [num_int] = " & numint & " and pro_01.[id_tipocomp] = pro_03.[id_tipocomp] and pro_01.[id_usuario] = g1.[id_usuario] and pro_01.[id_obra] = a4.[id_obra]"
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
   Me.estado = rs("pro_01.estado")
   Me.fecha = rs("fecha")
   Me.fechaesperado = rs("fecha_esperado")
   Me.id_obra = rs("pro_01.id_obra")
   Me.idtipocomp = rs("pro_01.id_tipocomp")
   Me.idusuario = rs("pro_01.id_usuario")
   Me.numcomp = rs("num_comprobante")
   Me.numint = rs("num_int")
   Me.sucursal = rs("sucursal")
   Me.observaciones = rs("observaciones")
   Me.copias = rs("copias")
   Me.abreviatura = rs("abreviatura")
   Me.obra = rs("a4.descripcion")
   Me.observacion_obra = rs("obs")
   Me.usuario = rs("usuario")
   Me.desc_comprobante = rs("pro_03.descripcion")
   
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



