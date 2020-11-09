VERSION 5.00
Object = "{0A6BE9FC-5039-11D5-98EC-0800460222F0}#1.0#0"; "IFEpson.ocx"
Begin VB.Form fsc_cierrez 
   Caption         =   "CIERRE Z"
   ClientHeight    =   1995
   ClientLeft      =   3150
   ClientTop       =   2955
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   ScaleHeight     =   1995
   ScaleWidth      =   8730
   Begin EPSON_Impresora_Fiscal.PrinterFiscal EPSON6 
      Left            =   120
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
   End
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

Private Sub Command1_Click()
       espere.Show
       espere.Refresh
       espere.Label1 = "Espere.... Emitiendo Cierre Z"
       'EPSON6.PortNumber = 1
       r = EPSON6.Status("N")
       If r Then
                   
      '   QUERY = "INSERT INTO fsc_002([ult_comp], [fecha_inicio_jornada], [hora_inicio_jornada], [ultimo_cierre], [auditoria_parcial], [auditoria_total], [texto_auditoria_impresor], [texto_auditoria])"
      '   QUERY = QUERY & " VALUES (" & Val(epson3.AnswerField_3) & ", '" & epson3.AnswerField_4 & "', '" & epson3.AnswerField_5 & "', " & Val(epson3.AnswerField_6) & ", " & Val(epson3.AnswerField_7) & ", " & Val(epson3.AnswerField_8) & ", '" & epson3.AnswerField_9 & "', '" & epson3.AnswerField_10 & "')"
      '   cn1.Execute QUERY
          
       Else
          MsgBox ("Error al generar datos para el cierre Z")
       End If
       
       
            
       r = EPSON6.CloseJournal("Z", "P")
       
       Unload espere
       If r Then
           Call grabaz
           MsgBox ("Cierre Z Emitido")
       Else
           MsgBox ("Error al Emitir el Cierre Z")
       End If
    
    
End Sub

Sub grabaz()
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
  numcomp = EPSON6.AnswerField_3
  iva = Val(Mid$(EPSON6.AnswerField_11, 1, Len(EPSON6.AnswerField_11) - 2) & "." & Mid$(EPSON6.AnswerField_11, Len(EPSON6.AnswerField_11) - 1, 2))
  total = Val(Mid$(EPSON6.AnswerField_10, 1, Len(EPSON6.AnswerField_10) - 2) & "." & Mid$(EPSON6.AnswerField_10, Len(EPSON6.AnswerField_10) - 1, 2))
  subtotal = total - iva
  nograbado = 0
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


