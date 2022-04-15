VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
'variables locales para almacenar los valores de las propiedades
Private mvaridusuario As Long 'copia local
Private mvardenominacion As String 'copia local
Private mvarnivel As Long 'copia local
Public Sub cargar(ByVal idusuario As Long)
On Error GoTo er_carga
q = "select * from  g1 where [id_usuario] = " & idusuario
Set rst = New ADODB.Recordset
rst.Open q, cn1
If Not rst.BOF And Not rst.EOF Then
  Me.idusuario = rst("id_usuario")
  Me.denominACION = rst("usuario")
  Me.nivel = rst("grupo")
Else
  Me.idusuario = 0
End If
Set rst = Nothing
Exit Sub
 
er_carga:
  MsgBox ("Error. Clase:Usuarios  Modulo:Carga")
  Me.idusuario = 0
  Exit Sub

End Sub

Public Property Let nivel(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.nivel = 5
    mvarnivel = vData
End Property


Public Property Get nivel() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.nivel
    nivel = mvarnivel
End Property



Public Property Let denominACION(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.denominacion = 5
    mvardenominacion = vData
End Property


Public Property Get denominACION() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.denominacion
    denominACION = mvardenominacion
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



