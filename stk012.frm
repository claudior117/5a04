VERSION 5.00
Begin VB.Form stk_ajustedesdeinst 
   Caption         =   "Ajuste  de Stock de Movimientos desde stock instataneo"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   5760
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "stk012.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "stk012.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.TextBox t_detalle 
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Top             =   1080
      Width           =   4575
   End
   Begin VB.TextBox t_fecha 
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Detalle"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Fecha"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "stk_ajustedesdeinst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnacepta_Click()

  J = MsgBox("Este proceso realiza movimientos de ajuste en el stock segun stock instantaneo y puede demorar. Es necesario salir del sistema de todas las terminales para ejecutarlo. ¿Confirma?", 4)
  If J = 6 Then
   Load espere

   
   q = "select * from a2 order by [id_producto]"
   Set rs3 = New ADODB.Recordset
   rs3.Open q, cn1, adOpenDynamic, adLockOptimistic
   espere.Show
   
   espere.Refresh
   Set cl_stock = New STOCK

   While Not rs3.EOF
      
     ip = rs3("id_producto")
     dp = rs3("descripcion")
     cl_stock.sacastock (ip)
   
     sm = cl_stock.stock_movimientos
     si = cl_stock.stock_instantaneo
     If sm <> si Then
       If sm > si Then
         'realizar ajuste de salida de mercaderia
          u = "S"
          cant = sm - si
     
       Else
         'realizar ajuste de ingreso de mercaderia
          u = "E"
          cant = si - sm
       
       End If
     
      cn1.BeginTrans
      QUERY = "INSERT INTO stk_02([fecha], [letra], [num_comprobante], [id_usuario], [detalle], [sucursal], [tipo_comprobante], [id_proveedor], [id_obra])"
      QUERY = QUERY & " VALUES ('" & t_fecha & "', 'X', 0, " & para.id_usuario & ", '" & t_detalle & " ', 0, 1, 1,1)"
      cn1.Execute QUERY
      
      qr = "SELECT @@IDENTITY AS NewID"
      Set rs4 = cn1.Execute(qr)
      numint = rs4.Fields("NewID").Value

        
        QUERY = "INSERT INTO stk_03([num_int], [RENGLON], [id_producto], [descripcion], [unidad], [detalle], [cantidad], [ubicacion])"
        QUERY = QUERY & " VALUES (" & numint & ", 1, " & ip & ", '" & Left$(dp, 30) & "', ' ', '" & t_detalle & "', " & cant & ", '" & u & "')"
        cn1.Execute QUERY
      
        QUERY = "INSERT INTO stk_01([fecha], [id_producto], [cantidad], [ubicacion], [comprobante], [descripcion], [num_mov_int], [modulo], [id_cliente])"
        QUERY = QUERY & " VALUES ('" & t_fecha & "', " & ip & ", " & cant & ", '" & u & "', 'Mov.Int.Stk " & Format$(numint, "00000000") & _
        "', '" & t_detalle & " ', " & numint & ", 'S', 1)"
              
        cn1.Execute QUERY

     
         cn1.CommitTrans
      
     
     End If
     espere.Label1 = "Producto.." & ip
  
     espere.Label1.Refresh
     rs3.MoveNext
   

   Wend
   Set rs3 = Nothing
     Set cl_stock = Nothing
   
      
   Unload espere
   MsgBox ("Proceso Terminado")
  End If
  
End Sub










Private Sub btnsale_Click()
Unload Me
End Sub

