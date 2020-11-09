VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form prod_cc_detalle 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DETALLE DE COMPROBANTE DE PRODUCCION"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   12345
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5235
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   11895
   End
   Begin VB.Frame CUIT 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   11655
      Begin VB.TextBox t_numint 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   10320
         MaxLength       =   10
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800080&
         Caption         =   "Numero Interno Operacion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8040
         TabIndex        =   7
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   2
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "Pro003.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "Pro003.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   8370
      Width           =   12345
      _ExtentX        =   21775
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   8820
            MinWidth        =   8820
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "05/10/2012"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "17:00"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "prod_cc_detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim l1 As String


Private Sub btnsale_Click()
Unload Me
End Sub









Private Sub Form_Activate()
Call carga
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub

Sub carga()
List1.clear
l1 = "-------------------------------------------------------------------------------------------------------------"
l2 = "*************************************"

If t_NUMINT <> "" Then
  Set cl_compprod = New comprobantes_produccion
  cl_compprod.cargar2 (Val(t_NUMINT))
  If cl_compprod.numcomp > 0 Then
     List1.AddItem l2
     List1.AddItem cl_compprod.desc_comprobante
     List1.AddItem l2
     List1.AddItem Space$(60) & "Numero....:" & Format$(cl_compprod.sucursal, "0000") & "-" & Format$(cl_compprod.numcomp, "00000000")
     List1.AddItem Space$(60) & "Fecha.....:" & cl_compprod.fecha
     List1.AddItem Space$(60) & "Estado....:" & cl_compprod.estado
     List1.AddItem Space$(60) & "Esperado..:" & cl_compprod.fechaesperado
     List1.AddItem Space$(60) & "Num.Int...:" & cl_compprod.numint
     List1.AddItem Space$(60) & "Emitido...:" & cl_compprod.usuario
     
     List1.AddItem "Obra................: " & cl_compprod.obra
     List1.AddItem "Observaciones Obra..: " & cl_compprod.observacion_obra
     
     Call COMPROBANTES
     
     List1.AddItem l1
     List1.AddItem ""
     List1.AddItem ""
     List1.AddItem "Observaciones: " & cl_compprod.observaciones
     

 
 End If
 Set rs = Nothing
End If
End Sub
Sub COMPROBANTES()
     List1.AddItem ""
     List1.AddItem l1
     List1.AddItem "Codigo Detalle                                      Cantidad   Esperado    Nro. Oc.      Observaciones   "
     List1.AddItem l1
     Set rs1 = New ADODB.Recordset
     q = "select * from pro_02 where [num_int] = " & Val(t_NUMINT)
     rs1.Open q, cn1
     While Not rs1.EOF
         c = Format$(rs1("cantidad"), "@@@@@@.00")
         b = Format$(rs1("id_producto"), "00000")
         d = Format$(Left$(rs1("descripcion"), 45), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
         o = Format$(Left$(rs1("observaciones"), 45), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
         f = Format$(rs1("fecha_esperado"), "dd/mm/yyyy")
         noc = "0000-00000000"
         If rs1("num_int_oc") > 0 Then
           Set rs2 = New ADODB.Recordset
           q = "select * from a5 where [num_int] = " & rs1("num_int_oc") & " and [id_tipocomp] = 65"
           rs2.opem q, cn1
           If Not rs2.EOF And Not rs2.BOF Then
              noc = Format$(rs2("sucursal"), "0000") & "-" & Format$(rs2("num_comprobante"), "00000000")
           End If
           Set rs2 = Nothing
         End If
         List1.AddItem b & "  " & d & " " & c & " " & f & " " & noc & " " & o
         rs1.MoveNext
     Wend
     Set rs1 = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 3)
  'Case Is = 27
  '      Me.Hide
End Select

End Sub

Private Sub Form_Load()

Call barraesag(Me)


End Sub



Private Sub List1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F5] Imprime Comprobante -  [ESC] Termina "

End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
 If para.id_grupo_modulo_actual > 4 Then
  If cl_compprod.idtipocomp > 0 Then
   cl_compprod.imprimir
  End If
 End If
End If

If KeyCode = vbKeyF8 Then
   Call borracomp
End If


End Sub

Sub borracomp()
J = MsgBox("Confirma borrar comprobante", 4)
If J = 6 Then
     On Error GoTo errborra
     'busco el comprobante
     Set cl_comp = New COMPROBANTES
     cl_comp.cargar2 (Val(t_NUMINT))
          
     If cl_comp.STOCK <> "N" Then
        'modifica stock
        Set rs1 = New ADODB.Recordset
        q = "select * from a6, a2 where [num_int] = " & cl_comp.numint & " and a6.[id_producto] = a2.[id_producto]"
        rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
        While Not rs1.EOF
           If cl_comp.STOCK = "E" Then
              rs1("stock") = rs1("stock") - rs1("cantidad")
           Else
              rs1("stock") = rs1("stock") + rs1("cantidad")
           End If
           rs1.Update
           rs1.MoveNext
        Wend
        Set rs1 = Nothing
     End If
     
     
     cn1.BeginTrans
     'borro detalle de productos
       q = "delete * from a6 where [num_int] = " & Val(t_NUMINT)
       cn1.Execute q
     'borro stock
       q = "delete * from stk_01 where [num_mov_int] = " & Val(t_NUMINT) & " and [modulo] = 'C'"
       cn1.Execute q
     'borro caja
       q = "delete * from cyb_05 where [num_mov_int] = " & Val(t_NUMINT) & " and [modulo] = 'C'"
       cn1.Execute q
     
     
     If Val(t_tipocomp) = 50 Then 'ordenes de pago
        'actu ch. propios
        q = "select * from cyb_02 where [num_int_op] = " & Val(t_NUMINT)
        Set rs1 = New ADODB.Recordset
        rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
        While Not rs1.EOF
          rs1("estado") = "P"
          rs1("destino") = "Pendiente"
          rs1("importe") = 0
          rs1("num_int_op") = 0
          rs1.Update
          rs1.MoveNext
        Wend
        Set rs1 = Nothing
        
        q = "select * from a7, cyb_03 where [num_int] = " & Val(t_NUMINT) & " and [num_interno] = [num_int_fp] "
        Set rs1 = New ADODB.Recordset
        rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
        While Not rs1.EOF
          rs1("estado") = "C"
          rs1("destino") = " "
          rs1("num_int_op") = 0
          rs1.Update
          rs1.MoveNext
        Wend
        Set rs1 = Nothing
        
        q = "delete * from a7 where [num_int] = " & Val(t_NUMINT)
        cn1.Execute q
        
        q = "delete * from cyb_04 where [num_mov_int] = " & Val(t_NUMINT) & " and [modulo] = 'C'"
        cn1.Execute q
        
        
        'actualizo acumulador
        q = "update ret_01 set [pagos_mes] = [pagos_mes] - " & cl_comp.total & " where [id_proveedor] = " & cl_comp.idproveedor & " and [id_retgan] = " & cl_comp.idcodretgan & " and [mes] = " & Val(Mid$(cl_comp.fecha, 4, 2)) & " and [año] = " & Val(Mid$(cl_comp.fecha, 7, 4))
        cn1.Execute q
        
        'actualizo comp.aplicados
        q = "update a5 set [estado_pago] = 'N' where [num_op]= '" & Format$(cl_comp.sucursal, "0000") & "-" & Format$(cl_comp.numcomp, "00000000") & "'"
        cn1.Execute q
        
     
     End If
     
     
     If cl_comp.idtipocomp = 95 Then 'ret. ganancia
        'actualizo acumulador
        q = "update ret_01 set [ret_mes] = [ret_mes] - " & cl_comp.total & " where [id_proveedor] = " & cl_comp.idproveedor & " and [id_retgan] = " & cl_comp.idcodretgan & " and [mes] = " & Val(Mid$(cl_comp.fecha, 4, 2)) & " and [año] = " & Val(Mid$(cl_comp.fecha, 7, 4))
        cn1.Execute q
     End If
      
      
     'borro comp
     q = "delete * from a5 where [num_int] = " & Val(t_NUMINT)
     cn1.Execute q
     
     cn1.CommitTrans
      
     Set cl_comp = Nothing
    
     Unload Me
     

End If

Exit Sub

errborra:
MsgBox ("Error al Borrar Comprobante")
cn1.RollbackTrans
Exit Sub
End Sub

