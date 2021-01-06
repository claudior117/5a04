VERSION 5.00
Begin VB.Form fsc_cierrez 
   Caption         =   "CIERRE Z"
   ClientHeight    =   2565
   ClientLeft      =   3150
   ClientTop       =   2955
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   ScaleHeight     =   2565
   ScaleWidth      =   8730
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.CommandButton Command1 
         Caption         =   "CONFIRME EMISION CIERRE Z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   6615
      End
   End
End
Attribute VB_Name = "fsc_cierrez"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Fiscalz As Driver


Private Sub Command1_Click()
       espere.Show
       espere.Refresh
       espere.Label1 = "Espere.... Emitiendo Cierre Z"
       
       If Fiscalz.Inicializar Then
  
            Fiscalz.CancelarComprobante
            If Fiscalz.CierreZ Then
                Call grabaz2
            
            
                MsgBox ("Cierre realizado exitosamente")
            Else
                MsgBox (Fiscalz.ErrorDesc)
            End If
    
            Fiscalz.Finalizar
        Else
            MsgBox (Fiscalz.ErrorDesc)
        End If
        Unload espere
       
           
        
    
End Sub


Sub grabaz2()
  On Error GoTo ERRORGRABA
  numint = saca_ultnumero_int_comp("V")
  
  Set cl_compvta = New comprobantes_venta
  cl_compvta.sucursal = glo.sucursalf
  cl_compvta.actual (300)
  cl_compvta.letra = "Z"
  ep = "S"
  cp = "ctdo"
  contado = "S"
  cl_compvta.ctacte = "N"
  moneda = "P"
  numcomp = Fiscalz.CierreZTotales.NroCierre
  iva = Fiscalz.CierreZTotales.FNDTotalIVA
  total = Fiscalz.CierreZTotales.FNDTotalVentas
  subtotal = total - iva
  nograbado = Fiscalz.CierreZTotales.FNDTotalOtrosTributos
  If para.cotizacion > 0 Then
    total2 = total / para.cotizacion
  Else
    total2 = 0
  End If
      
  cn1.BeginTrans
  QUERY = "INSERT INTO vta_02([num_int], [sucursal], [num_comp], [letra], [id_tipocomp], [id_cliente], [fecha], [id_usuario], [subtotal], [impuestos], [iva], [total], [estado], [id_cuenta], [stock], [cta_cte], [grabado], [estado_pago], [recibo_Pago], [observaciones], [cotizacion_dolar], [total_otra_moneda], [moneda], [id_vendedor], [VENTA], " & _
  " [CONTADO], [perc_ib], [perc_gan], [perc_iva], [servicio], [fecha_pago], [fecha_vto], [cliente02], [direccion02], [cuit02], [localidad02], [id_tipo_iva02], [sucursal_ingreso], [id_actividad],[num_z])"
  QUERY = QUERY & " VALUES (" & numint & ", " & glo.sucursalf & ", " & numcomp & ", 'Z', 300, 1, '" & Format$(Now, "dd/mm/yyyy") & "', " & para.id_usuario & ", " & subtotal & ", " & nograbado & ", " & iva & ", " & total & ", 'A', 0, 'N', 'N','" & cl_compvta.grabado & "', '" & ep & "', '0000-00000000', 'Cierre Z', " & para.cotizacion & _
  ", " & total2 & ", 'P', 1, 'N', '" & contado & "', 0, 0, 0, 'N', '" & Format$(Now, "dd/mm/yyyy") & "', '" & Format$(Now, "dd/mm/yyyy") & "', '" & Left$(glo.nombrecli, 50) & "', '" & Left$(glo.direccioncli, 50) & "', '" & glo.CUIT & "', 'Rojas', 1, " & glo.sucursalf & ", 1, " & numcomp & ")"
  cn1.Execute QUERY
      
      
   QUERY = "update fsc_001 set  [ult_z]=" & Val(EPSON6.AnswerField_3)
   QUERY = QUERY & " where [sucursal_fiscal]= " & glo.sucursalf
   cn1.Execute QUERY
              
  cn1.CommitTrans
  Set rs = Nothing
  Set cl_compvta = Nothing
  
   
Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error al Grabar Cierre Z. Verifique los datos o sus permisos")
  

End Sub

Private Sub Form_Load()
Set Fiscalz = New Driver
Fiscalz.Modelo = cMODELO
Fiscalz.puerto = cPUERTO
Fiscalz.baudios = cBAUDIOS
End Sub
