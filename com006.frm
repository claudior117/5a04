VERSION 5.00
Object = "{0A6BE9FC-5039-11D5-98EC-0800460222F0}#1.0#0"; "IFEpson.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form COM_ajustesint 
   BackColor       =   &H00E0E0E0&
   Caption         =   "AJUSTES INTERNOS EN CUENTA CORRIENTE"
   ClientHeight    =   4515
   ClientLeft      =   2175
   ClientTop       =   1485
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   11880
   Begin VB.Frame Frame13 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   9720
      TabIndex        =   28
      Top             =   840
      Width           =   1935
      Begin VB.CheckBox Check3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Moneda Unica"
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1575
      End
   End
   Begin EPSON_Impresora_Fiscal.PrinterFiscal epson1 
      Left            =   0
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   9720
      TabIndex        =   19
      Top             =   0
      Width           =   1935
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pesos"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "U$s"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   9840
      TabIndex        =   16
      Top             =   1680
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox t_funcion 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label12 
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Funcion"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos Operacion"
      Height          =   1455
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   9375
      Begin VB.TextBox T_total2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4920
         MaxLength       =   10
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox t_total 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox t_observaciones 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   4
         Top             =   360
         Width           =   6855
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Total U$s"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3720
         TabIndex        =   27
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "Total"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Detalle:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   2055
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   9375
      Begin VB.ComboBox c_zona 
         Height          =   315
         ItemData        =   "com006.frx":0000
         Left            =   7440
         List            =   "com006.frx":0002
         TabIndex        =   30
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox c_sucursal 
         Height          =   315
         ItemData        =   "com006.frx":0004
         Left            =   7440
         List            =   "com006.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox t_cotizacion 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1680
         Width           =   735
      End
      Begin VB.ComboBox c_tipocomp 
         Height          =   315
         ItemData        =   "com006.frx":0008
         Left            =   1680
         List            =   "com006.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Text            =   "c_prov"
         Top             =   720
         Width           =   6015
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
         Caption         =   "Zona:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5880
         TabIndex        =   31
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Punto Venta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5880
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Cotizacion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Tipo :"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Proveedor:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9960
      TabIndex        =   9
      Top             =   3120
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "com006.frx":0030
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "com006.frx":08B2
         Style           =   1  'Graphical
         TabIndex        =   7
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
      TabIndex        =   8
      Top             =   4260
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   13229
            MinWidth        =   13229
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "21/03/2013"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "9:19"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "COM_ajustesint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim EXISTE As String
Dim cantidadp As Double
Dim calcula_perc_ib As String
Dim alicuota_perc_ib As Single
Dim minimo_perc_ib As Double
Dim gcuit As String
Dim numint As Long
Dim cuentaact As Long
Dim abreviatura As String
Dim cantlineas As Integer
Dim ubicacionctacte As String

Sub limpia()
   
   t_subtotal = ""
   t_nograbado = ""
   t_perc = ""
   t_iva = ""
   t_total = ""
   Option1 = True
   
End Sub


Private Sub btnacepta_Click()
 Call iniciagraba



End Sub
Sub iniciagraba()
If Val(t_total) > 0 Then
 J = MsgBox("Graba Comprobante ", 4)
 If J = 6 Then
   If verificaperiodog(t_fecha) = "A" Then
     para.z_actual = 0
     Call graba
   End If
  Else
   MsgBox ("Periodo Cerrado. Imposible grabar comprobante")
  End If
 Else
 MsgBox ("Imposible emitir comprobante. El total del comprobante debe ser > 0 ")
End If
  
    

End Sub






Private Sub btnsale_Click()
Unload Me
End Sub





Private Sub c_prov_LostFocus()
If c_prov.ListIndex < 0 Then
    c_prov.ListIndex = 0
 End If

End Sub


Sub inicia()
     t_cotizacion = para.cotizacion
   t_alicuotaib = "0.00"
   T_PERCIB = "0.00"
 

End Sub


Private Sub c_sucursal_LostFocus()
If c_sucursal.ListIndex < 0 Then
  c_sucursal.ListIndex = buscaindice(c_sucursal, glo.sucursal)
End If
t_sucursal = Format$(c_sucursal, "0000")
t_numcomp = ""
End Sub

Private Sub c_tipocomp_GotFocus()
btnacepta.Enabled = False
End Sub

Private Sub c_tipocomp_LostFocus()
If c_tipocomp.ListIndex < 0 Then
  c_tipocomp.ListIndex = 0
End If
  
End Sub





Private Sub Form_Activate()
c_tipocomp.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     gen_tools.Show
  
End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call TabEnter2(Me, 6)
End If


End Sub

Private Sub Form_Load()
Call INICIALIZA2(Me)
Call carga_proveedores(c_prov)
c_prov.RemoveItem 0
c_prov.ListIndex = 0

Call carga_SUCURSALES(c_sucursal)
c_sucursal.ListIndex = buscaindice(c_sucursal, glo.sucursal)

c_tipocomp.ListIndex = 0


Call carga_zonas(c_zona)
c_zona.ListIndex = 0


Call barraesag(Me)
If para.moneda = "P" Then
  Option4 = True
Else
  Option3 = True
End If
t_sucursal = Format$(glo.sucursal, "0000")

gcuit = "0"


End Sub

Sub graba()
  
     ep = "S"
     cp = "ctdo"
     ssi = 0
     ctdo = "N"
     numint = saca_ultnumero_int_comp("C")
      
     If c_tipocomp.ListIndex = 0 Then
         tipoc = 24
         cc = "H"
     Else
         tipoc = 34
         cc = "D"
     End If
     
     Set cl_comp = New COMPROBANTES
     cl_comp.actual (tipoc)
      
         
      t_obs = RTrim$(t_observaciones) & " "
      
      If Option4 = True Then
        moneda = "P"
      Else
        moneda = "D"
      End If
      
      
      If Check3 Then
        t2 = 0
      Else
        t2 = Val(T_total2)
      End If
      
      
      Set cl_prov = New proveedores
      cl_prov.carga (c_prov.ItemData(c_prov.ListIndex))
      letrac = cl_prov.letra
      numc = saca_ultnumero_comp2(tipoc)

      
      
      cn1.BeginTrans
      
      QUERY = "INSERT INTO a5([num_int], [sucursal], [num_comprobante], [letra], [id_tipocomp], [id_proveedor], [fecha], [id_usuario], [subtotal], " & _
" [no_grabado], [percep_ret], [iva], [total], [fecha_prob_entrega], [fecha_recepcion], [estado], [id_codretgan], [id_cuenta], [stock], [ctacte], [grabado], " & _
" [estado_pago], [num_op], [id_codretib], [obs], [condiciones], [info_contacto], [moneda], [cotiz_dolar], [contado], [TOTAL_D], [monto_suj_ret], " & _
"[alicuota_ret], [ret_mes], [pagos_realizados], [pago_actual], [minimo_no_imp], [fecha_vto], [COMPRA], [saldo_impago], [zona], [cuit05], [proveedor05])"
      
 QUERY = QUERY & " VALUES (" & numint & ", " & Val(c_sucursal) & ", " & numc & ", '" & letrac & "', " & tipoc & _
 ", " & c_prov.ItemData(c_prov.ListIndex) & ", '" & t_fecha & "', " & para.id_usuario & ", " & Val(t_total) & ", " & Val(0) & ", " & Val(0) & ", " & Val(0) & _
 ", " & Val(t_total) & ", '" & Format$(Now, "dd/mm/yyyy") & "', '" & t_fecha & "', 'A', " & Val(1) & ", " & Val(0) & _
 ", '" & cl_comp.STOCK & "', '" & cc & "', '" & cl_comp.grabado & "', '" & ep & "', '" & cp & "', " & Val(1) & ", '" & t_observaciones & "', ' ', ' ', '" & moneda & "', " & _
 Val(t_cotizacion) & ", '" & ctdo & "', " & t2 & ", 0, 0, 0, 0, 0, 0, '" & t_fecha & "', '" & cl_comp.compra & "', " & ssi & ", " & c_zona.ListIndex + 1 & ", " & Val(com_proveedor.t_cuit) & ", '" & Left$(Trim$(com_proveedor.t_cli), 50) & "')"
 cn1.Execute QUERY
   
      
      
  
      'actualiza contador
      QUERY = "update g2 set  [ult_num]=" & numc
      QUERY = QUERY & " where [id_tipo_comp]= " & tipoc
      cn1.Execute QUERY
          
      
  cn1.CommitTrans
  Set cl_prov = Nothing
      
      
   
  Call INICIALIZA2(Me)
  c_tipocomp.SetFocus

     

Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  
  

End Sub





Private Sub Option3_Click()
Label13 = "Total $"
End Sub

Private Sub Option4_Click()
Label13 = "Total U$s"
End Sub

Private Sub Option4_GotFocus()
'Call keyform(Me, "A")


End Sub

Private Sub Option4_LostFocus()
'Call keyform(Me, "D")

End Sub



Private Sub t_cotizacion_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 1)
End Sub

Private Sub t_cotizacion_LostFocus()
If Val(t_cotizacion) <= 0 Then
   t_cotizacion = 1
End If

End Sub

Private Sub t_fecha_LostFocus()
If Not IsDate(t_fecha) Then
  t_fecha = Format$(Now, "dd/mm/yyyy")
Else
  t_fecha = Format$(t_fecha, "dd/mm/yyyy")
End If
Call verifica_fechacorte(t_fecha)

End Sub







Private Sub t_observaciones_LostFocus()
Call NULOS(t_observaciones)
End Sub





Private Sub t_total_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 1)
End Sub

Private Sub t_total_LostFocus()
t_total = Format$(t_total, "######0.00")
If Option4 = True Then
 T_total2 = Format$(Val(t_total) / Val(t_cotizacion), "######0.00")
Else
 T_total2 = Format$(Val(t_total) * Val(t_cotizacion), "######0.00")
End If
End Sub

Private Sub T_total2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 
 btnacepta.Enabled = True
 btnacepta.SetFocus
End If

End Sub


