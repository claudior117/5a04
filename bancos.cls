VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
'variables locales para almacenar los valores de las propiedades
Private mvarnummov As Long 'copia local
Public Sub borrar(ByVal numint As Long)
  If numint > 0 Then
   cn1.BeginTrans
    
      
      
      
      Call borrarconciliacion(numint)
      
       Set rsb = New adodb.Recordset
       q = "select * from cyb_04 where [num_mov_banco] = " & numint
       rsb.Open q, cn1
       If Not rsb.EOF And Not rsb.BOF Then
          If rsb("num_mov_int_compras") > 0 Then
            ''borra compras
            Set cl_comp = New COMPROBANTES
            cl_comp.cargar2 (rsb("num_mov_int_compras"))
            cl_comp.borrar
            Set cl_comp = Nothing
          End If
       End If
       Set rsb = Nothing
       
      
      QUERY = "DELETE FROM cyb_04 WHERE [num_mov_banco] = " & numint
      cn1.Execute QUERY
      
      QUERY = "DELETE FROM cyb_05 WHERE [num_mov_int] = " & numint & " and [modulo] = 'B'"
      cn1.Execute QUERY
       
      QUERY = "update cyb_03 set  [estado]='C', [num_mov_banco_e]=0, [destino]='En Cartera ', [tipo_salida]='C'"
      QUERY = QUERY & " where [num_mov_banco_e]= " & numint
      cn1.Execute QUERY
           
      QUERY = "update cyb_02 set  [estado]='P', [num_mov_banco]=0, [destino]='Pendiente', [importe]=0"
      QUERY = QUERY & " where [num_mov_banco]= " & numint
      cn1.Execute QUERY
      
      
      
      Call borracontabilidad(numint, "B")

   cn1.CommitTrans
 End If
End Sub
Public Sub borrarconciliacion(ByVal ni)
   q = "select * from c_02 where [modulo] = 'B' and [num_mov_int] = " & ni & " and [descripcion] like '%Conciliacion%'"
   Set rsx = New adodb.Recordset
   rsx.MaxRecords = 1
   rsx.Open q, cn1
   If Not rsx.EOF And Not rsx.BOF Then
        nicgr = rsx("num_interno")
        QUERY = "DELETE FROM c_02 WHERE [num_interno] = " & nicgr
        cn1.Execute QUERY
    
        QUERY = "DELETE FROM c_03 WHERE [num_interno] = " & nicgr
        cn1.Execute QUERY
   
   Else
        nicgr = 0
   End If
   Set rsx = Nothing
      
   

End Sub


Public Property Let nummov(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.nummov = 5
    mvarnummov = vData
End Property


Public Property Get nummov() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.nummov
    nummov = mvarnummov
End Property



