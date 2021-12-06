VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fsc_tiqueNF 
   BackColor       =   &H00E0E0E0&
   Caption         =   "TIQUE NO FISCAL"
   ClientHeight    =   8490
   ClientLeft      =   300
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos Complementarios"
      Height          =   1095
      Left            =   120
      TabIndex        =   25
      Top             =   6240
      Visible         =   0   'False
      Width           =   6855
      Begin VB.ComboBox c_vend 
         Height          =   315
         Left            =   1560
         TabIndex        =   27
         Top             =   240
         Width           =   5055
      End
      Begin VB.ComboBox c_actividad 
         Height          =   315
         Left            =   1560
         TabIndex        =   26
         Top             =   600
         Width           =   5055
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Vendedor"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Actividad:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   8760
      TabIndex        =   22
      Top             =   0
      Width           =   3015
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Impresora Fiscal"
      Height          =   735
      Left            =   120
      TabIndex        =   20
      Top             =   7320
      Width           =   5895
      Begin VB.TextBox t_impfiscal 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      Height          =   1695
      Left            =   7320
      TabIndex        =   17
      Top             =   6480
      Width           =   4455
      Begin VB.TextBox t_total 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   6
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Parciales"
      Height          =   855
      Left            =   360
      TabIndex        =   14
      Top             =   7320
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox t_limite 
         Height          =   285
         Left            =   4320
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "IVA"
         Height          =   195
         Left            =   2760
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox t_iva 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox t_subtotal 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox t_nograbado 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "limite controlador"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   4560
         TabIndex        =   32
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Subtotal"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "No Grabado"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   6600
      TabIndex        =   11
      Top             =   7680
      Visible         =   0   'False
      Width           =   1815
      Begin VB.TextBox t_funcion 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   12
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
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   4815
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8493
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   5655
      Begin VB.TextBox t_letra 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox t_numcomp 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3720
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   0
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox t_sucursal 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2640
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Comprobante:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   8235
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   17639
            MinWidth        =   17639
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   6120
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
   End
End
Attribute VB_Name = "fsc_tiqueNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim EXISTE As String
Dim cantidadp As Double
Dim gcuit As String
Dim estadotique As String
Dim r As Boolean
Dim numint As Long
Dim cuentaact As Long
Dim gprueba As Integer '0 fiscal   '1 prueba
Dim gsucursalprueba As Integer

Dim Fiscaltq As Driver




Sub limpia()
   Call armagrid
   t_subtotal = ""
   t_nograbado = ""
   t_perc = ""
   t_iva = ""
   T_TOTAL = ""
   Option1 = True
   
End Sub


Sub grabaformapago()
  For i = 1 To fsc_formapago.msf2.Rows - 1
         If Val(fsc_formapago.msf2.TextMatrix(i, 0)) = 3 Then
                'ch. terceros
                q = "select * from cyb_03"
                Set rs = New ADODB.Recordset
                rs.Open q, cn1, adOpenDynamic, adLockOptimistic
                rs.AddNew
                 rs("fecha_emision") = t_fecha
                 rs("num_cheque") = Val(fsc_formapago.msf2.TextMatrix(i, 2))
                 rs("banco") = fsc_formapago.msf2.TextMatrix(i, 3)
                 rs("sucursal") = fsc_formapago.msf2.TextMatrix(i, 4)
                 rs("titular") = fsc_formapago.msf2.TextMatrix(i, 5)
                 rs("importe") = Val(fsc_formapago.msf2.TextMatrix(i, 6))
                 rs("estado") = "C"
                 rs("fecha_dif") = fsc_formapago.msf2.TextMatrix(i, 7)
                 rs("origen") = fsc_formapago.msf2.TextMatrix(i, 5)
                 rs("destino") = " "
                 rs("num_mov_banco_i") = 0
                 rs("num_mov_banco_e") = 0
                 rs("num_int_op") = 0
                 rs("num_int_rbo") = numint
                 rs("fecha_salida") = t_fecha
                 rs("fecha_ingreso") = t_fecha
                 rs("tipo_salida") = "C"
                rs.Update
                
                qr = "SELECT @@IDENTITY AS NewID"
                Set rs = cn1.Execute(qr)
                numintch = rs.Fields("NewID").Value

                
                Set rs = Nothing
         
         Else
           numintch = 0
         End If
         
         
         If Val(fsc_formapago.msf2.TextMatrix(i, 0)) = 4 Then
                q = "select * from cyb_04"
                Set rs = New ADODB.Recordset
                rs.Open q, cn1, adOpenDynamic, adLockOptimistic
                rs.AddNew
                 rs("id_banco") = Val(fsc_formapago.msf2.TextMatrix(i, 8))
                 rs("fecha") = fsc_formapago.msf2.TextMatrix(i, 7)
                 rs("importe") = Val(fsc_formapago.msf2.TextMatrix(i, 6))
                 rs("id_tipomov") = 60 'transf
                 rs("fecha_dif") = fsc_formapago.msf2.TextMatrix(i, 7)
                 rs("ubicacion") = "H"
                 rs("entro") = "N"
                 rs("fecha_acreed") = fsc_formapago.msf2.TextMatrix(i, 7)
                 rs("num_comp") = Val(vta_formapago.msf2.TextMatrix(i, 2))
                 rs("detalle") = "Transf." & Left$(fsc_formapago.msf2.TextMatrix(i, 5), 30)
                 rs("modulo") = "V"
                 rs("num_mov_int") = numint
                 rs("id_tipodbcr") = 1
                rs.Update
                
                Set rs = Nothing
         End If
         
         
         q = "select * from cyb_01 where [id_forma_pago] = " & Val(fsc_formapago.msf2.TextMatrix(i, 0))
         Set rs = New ADODB.Recordset
         rs.Open q, cn1
         If Not rs.EOF And Not rs.BOF Then
          If rs("CAJA") = "S" Then
             ctach = rs("id_cuenta_cont")
             QUERY = "INSERT INTO cyb_05([id_cuenta_caja], [id_cuenta_contra], [descripcion], [importe], [ubicacion], [fecha], [num_mov_int], [modulo], [operacion], [id_forma_pago], [num_int_ch_terc], [id_usuario])"
             QUERY = QUERY & " VALUES (" & ctach & ", " & cuentaact & ", '" & RTrim$(Left$("Tique Contado " & Format$(Val(t_numcomp), "0000000"), 49)) & " ', " & Val(fsc_formapago.msf2.TextMatrix(i, 6)) & ", 'D', '" & t_fecha & "', " & numint & ", 'V', '" & Left$(abreviatura, 5) & " " & t_letra & Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numcomp), "00000000") & "', " & Val(fsc_formapago.msf2.TextMatrix(i, 0)) & ", " & numintch & ", " & para.id_usuario & ")"
             cn1.Execute QUERY
          End If
         End If
         Set rs = Nothing

                 
        'formas de pago
        QUERY = "INSERT INTO vta_04([num_int], [secuencia], [id_formapago], [formapago], [num_ch], [detalle_banco], [sucursal], [titular], [importe], [fecha_dif], [num_int_fp])"
        QUERY = QUERY & " VALUES (" & numint & ", " & i & ", " & Val(fsc_formapago.msf2.TextMatrix(i, 0)) & ", '" & Left$(RTrim$(fsc_formapago.msf2.TextMatrix(i, 1)), 9) & _
        " ', " & Val(fsc_formapago.msf2.TextMatrix(i, 2)) & ", '" & RTrim$(fsc_formapago.msf2.TextMatrix(i, 3)) & " ', '" & RTrim$(fsc_formapago.msf2.TextMatrix(i, 4)) & " ', '" & RTrim$(fsc_formapago.msf2.TextMatrix(i, 5)) & " ', " & Val(fsc_formapago.msf2.TextMatrix(i, 6)) & ", '" & RTrim$(fsc_formapago.msf2.TextMatrix(i, 7)) & " ', " & numintch & ")"
        cn1.Execute QUERY

      Next i

End Sub


Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 11
msf1.ColWidth(0) = 500
msf1.ColWidth(1) = 700
msf1.ColWidth(2) = 5000
msf1.ColWidth(3) = 1100
msf1.ColWidth(4) = 900
msf1.ColWidth(5) = 1100
msf1.ColWidth(6) = 900
msf1.ColWidth(7) = 1100
msf1.ColWidth(8) = 1100
msf1.ColWidth(9) = 1100
msf1.ColWidth(10) = 1100
msf1.TextMatrix(0, 0) = "Reng."
msf1.TextMatrix(0, 1) = "Id.Prod."
msf1.TextMatrix(0, 2) = "Detalle"
msf1.TextMatrix(0, 3) = "Cantidad"
msf1.TextMatrix(0, 4) = "Unidad"
msf1.TextMatrix(0, 5) = "P.U."
msf1.TextMatrix(0, 6) = "% Iva"
msf1.TextMatrix(0, 7) = "Importe"
msf1.TextMatrix(0, 8) = "PU s/iva"
msf1.TextMatrix(0, 9) = "Iva"
msf1.TextMatrix(0, 10) = "% IB"

End Sub




Sub inicia()
espere.Show
espere.Label1 = "Inicializando Comprobante....."
espere.Refresh
Set cl_cli = New Clientes
cl_cli.carga (c_prov.ItemData(c_prov.ListIndex))
If cl_cli.id > 0 Then
   t_letra = cl_cli.letra
   't_sucursal = Format$(val(glo.sucursal, "0000")
   gcuit = cl_cli.CUIT
   Set cl_compvta = New comprobantes_venta
   cl_compvta.sucursal = Val(c_sucursal)
   cl_compvta.actual (c_tipocomp.ItemData(c_tipocomp.ListIndex))
   cl_compvta.letra = t_letra
   cl_compvta.SACANUMCOMP
   t_numcomp = Format$(cl_compvta.numcomp, "00000000")
   Set cl_compvta = Nothing
   t_cotizacion = para.cotizacion

   If para.calcula_perc_ib = "S" Then
     Set cl_padronib = New padron_ib
     cl_padronib.cuit_texto = cl_cli.CUIT
     cl_padronib.buscar
     t_alicuotaib = Format$(cl_padronib.tasa_percib, "##0.00")
     Select Case cl_padronib.estado_consulta
     Case Is = "OK"
       Label20 = "¡COMPROBANTE SUJETO A PERCEPCION IB! Consulta del Padron de IB Satistactoria"
     Case Is = "NO"
       Label20 = "¡ATENCION! El contribuyente NO se encuentra en el padron, si corresponde debera aplicarle una percpcion de IB del 3%"
     Case Is = "ER"
       Label20 = "¡CUIDADO! Numero de cuit con formato invalido. Padron NO consultado"
     End Select
     Frame11.Visible = True
     
     Set cl_padronib = Nothing
     
     
   Else
     t_alicuotaib = "0.00"
     T_PERCIB = "0.00"
     gcuit = "0"
   End If
   Call armagrid
   Unload espere
Else
  Unload espere
  MsgBox ("Error. No se puedo Inicializa el Cliente")
End If






End Sub






Sub CALCULATOTALES()
vta_facturacion2.armagrid
t = 0
s = 0
r8 = 0
tin = 0
For i = 1 To msf1.Rows - 1
      r8 = (Val(msf1.TextMatrix(i, 8)) * Val(msf1.TextMatrix(i, 3))) 'importe sin iva
      tin = tin + r8
      t = t + Val(msf1.TextMatrix(i, 7)) 'importe con iva
      
      
      'agrega en composicion de iva
      X = 1
      While X < vta_facturacion2.msf1.Rows
        If Val(vta_facturacion2.msf1.TextMatrix(X, 0)) = Val(msf1.TextMatrix(i, 6)) Then
           vta_facturacion2.msf1.TextMatrix(X, 1) = Format(Val(vta_facturacion2.msf1.TextMatrix(X, 1)) + r8, "######0.00")
           vta_facturacion2.msf1.TextMatrix(X, 2) = Format(Val(vta_facturacion2.msf1.TextMatrix(X, 2)) + (r8 * Val(msf1.TextMatrix(i, 6)) / 100), "######0.00")
           X = vta_facturacion2.msf1.Rows
        Else
           X = X + 1
        End If
      Wend

Next i
T_TOTAL = t
t_subtotal = tin
t_iva = t - tin
  
  
Call sacatotales
  
End Sub

Private Sub c_actividad_LostFocus()
If c_actividad.ListIndex < 0 Then
  c_actividad.ListIndex = 0
End If

End Sub

Private Sub c_vend_LostFocus()
If c_vend.ListIndex < 0 Then
  c_vend.ListIndex = 0
End If
  
End Sub

Sub cargarenglon2(t As String)

  ip = "(" & fsc_tiqueNF1.t_ip & ")"
  d = fsc_tiqueNF1.t_detalle
  cu = Format$(Val(fsc_tiqueNF1.t_cantidad), "#####0.000")
  ti = Format$(fsc_tiqueNF1.c_tasa, "####0.00")
  u = RTrim$(fsc_tiqueNF1.t_unidad)
  puf = Format$(Val(fsc_tiqueNF1.t_pu), "#####0.00")
  pu = Format$(Val(puf) / (1 + Val(fsc_tiqueNF1.c_tasa) / 100), "#####0.000")
  im = Format$(Val(puf) * Val(cu), "#####0.00")
  If u = "" Then
    u = " "
  End If
  
  r = msf1.Rows
  msf1.AddItem r & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & cu & Chr(9) & u & Chr$(9) & puf & Chr(9) & ti & Chr(9) & im & Chr(9) & pu & Chr(9) & (puf - pu) & Chr$(9) & t_tasaib
    
  CALCULATOTALES
  sacatotales
  para.producto_sel = 0
End Sub
 

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call TabEnter2(Me, 19)
End If


End Sub

Private Sub Form_Load()

Call INICIALIZA2(Me)
Call armagrid
t_sucursal = Format$(glo.sucursal, "0000")
t_letra = "B"
t_fecha = Format$(Now, "dd/mm/yyyy")
estadotique = "C" 'cerrado

Call carga_vendedores(c_vend)
c_vend.ListIndex = 0
Call carga_actividades(c_actividad)
c_actividad.ListIndex = 0
Load fsc_tiqueNF1
Load fsc_tiqueNF2
Load vta_facturacion2
Load fsc_formapago

t_limite = 1000000
t_limite.Tag = 1000000
  

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload fsc_tiqueNF1
Unload fsc_tiqueNF2
Unload vta_clientes
Unload vta_facturacion2
Unload fsc_formapago
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(1) = "[INS] Agrega  - [F5] Saca Item - [F6] Bonif.% - [F9] Cerrar Tique - [F4] Cancela Tique   "
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If
Me.KeyPreview = False

End Sub
Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyInsert Then
  If estadotique = "A" Then
   
   fsc_tiqueNF1.t_tipo = "N"
   fsc_tiqueNF1.t_renglon = ""
   fsc_tiqueNF1.t_cantidad = ""
   fsc_tiqueNF1.t_pu = ""
   fsc_tiqueNF1.t_importe = ""
   fsc_tiqueNF1.Show
  End If
End If



If KeyCode = vbKeyF5 Then
      If msf1.Rows > 2 Then
        msf1.RemoveItem (msf1.Row)
        Call renumera
      Else
        Call armagrid
      End If
      Call CALCULATOTALES
End If


If KeyCode = vbKeyF4 Then
 
   J = MsgBox("Confirma cancelar el tique actual(S/N)", 4)
   If J = 6 Then
     Call armagrid
   End If
   
 End If


If KeyCode = vbKeyF9 Then
  Call CALCULATOTALES
  Call sacatotales
  If Val(T_TOTAL) > 0 Then
   'J = MsgBox("Cierra Tiquet", 4)
   'If J = 6 Then
     Call renumera
     fsc_tiqueNF2.T_TOTAL = T_TOTAL
     fsc_tiqueNF2.Show
     fsc_tiqueNF2.Refresh
   'End If
  Else
    MsgBox ("El importe del Tique debe ser > 0 ")
  End If
End If



 
If KeyCode = vbKeyF6 Then
 If msf1.Rows > 1 And Val(T_TOTAL) > 0 Then
   J = InputBox$("Ingrese % a  bonificar, luego el tique se cerrará", " % BONICICACION", "")
   If Val(J) > 0 And Val(J) < 100 Then
          dto = (Val(T_TOTAL) * Val(J)) / 100
          r = msf1.Rows
          msf1.AddItem r & Chr(9) & Format$(1, "00000") & Chr(9) & "Bonificación " & J & "%" & Chr(9) & 1 & Chr(9) & "" & Chr$(9) & Format$(-Val(dto), "######0.00") & Chr(9) & 0 & Chr(9) & Format$(-Val(dto), "######0.00") & Chr(9) & Format$(-Val(dto), "######0.00") & Chr(9) & 0 & Chr$(9) & 0
           Call CALCULATOTALES
        ' SendKeys "{F9}" 'cierra el tique
   End If
 End If
End If


If KeyCode = vbKeyF12 Then
  gen_tools.Show
End If
End Sub
Sub renumera()
For i = 1 To msf1.Rows - 1
  msf1.TextMatrix(i, 0) = i
Next i


End Sub


Sub cierratique2()
Dim r As Boolean
If estadotique = "A" And Val(T_TOTAL) > 0 Then
  'cierro tique
  espere.Show
  espere.Label1 = "Espere Actualizando Contadores...."
  espere.Label1.Refresh
  
  'tique prueba
   Label2.Visible = False
   estadotique = "C"
   espere.Label1 = "Espere Grabando Tique...."
   espere.Label1.Refresh
   Call graba
   
   Unload espere
 Else
   MsgBox ("No Existe Tique Abierto")
 End If

End Sub

Sub graba()
   'On Error GoTo ERRORGRABA
   
  numint = saca_ultnumero_int_comp("V")
      
  Set cl_compvta = New comprobantes_venta
  'prueba
      cl_compvta.sucursal = Val(t_sucursal)
      cl_compvta.actual (310) 'tique
      cl_compvta.letra = "B"
      cl_compvta.SACANUMCOMP
      t_numcomp = Format$(cl_compvta.numcomp, "00000000")
      cl_compvta.ACTUALIZA_NUMERADOR
      
      para.z_actual = Val(Mid$(Format$(Now, "dd/mm/yy"), 7, 2) & Mid$(Format$(Now, "dd/mm/yy"), 4, 2) & Mid$(Format$(Now, "dd/mm/yy"), 1, 2))
        
    abreviatura = cl_compvta.abreviatura
  
       ep = "S"
       cp = "ctdo"
       contado = "S"
       cl_compvta.ctacte = "N"
      
      cl_compvta.ACTUALIZA_NUMERADOR
      
      moneda = "P"
 
      
      Set rs = New ADODB.Recordset
      q = "select * from g8 where [id_actividad] = " & c_actividad.ItemData(c_actividad.ListIndex)
      rs.Open q, cn1
      If Not rs.EOF And Not rs.BOF Then
       codact = rs("id_actividad")
       alicuotaib = rs("alicuota_ib")
       cuentaact = rs("cuenta_contable_venta")
      Else
       codact = 0
       alicuotaib = 0
       cuentaact = para.cuenta_ventas
      End If
      Set rs = Nothing
       
       codvend = c_vend.ItemData(c_vend.ListIndex)
                     
      tiporespiva = 3
      
    cn1.BeginTrans
    QUERY = "INSERT INTO vta_02([num_int], [sucursal], [num_comp], [letra], [id_tipocomp], [id_cliente], [fecha], [id_usuario], [subtotal], [impuestos], [iva], [total], [estado], [id_cuenta], [stock], [cta_cte], [grabado]," & _
    " [estado_pago], [recibo_Pago], [observaciones], [cotizacion_dolar], [total_otra_moneda], [moneda], [id_vendedor], [VENTA], [CONTADO], [perc_ib], [perc_gan], [perc_iva], [id_actividad], [alicuota_ib], " & _
    "[alicuota_perc_iva], [canje_cereal], [fecha_vto], [total_bultos], [valor_declarado], [transporte], [direccion_transp], [cuit_transp], [perc_ss], [sucursal_ingreso], [cliente02], [direccion02], [cuit02], [localidad02], [id_tipo_iva02], [saldo_impago02], [num_z])"

    QUERY = QUERY & " VALUES (" & numint & ", " & Val(t_sucursal) & ", " & Val(t_numcomp) & ", '" & t_letra & "', 310, 1" & _
    ", '" & t_fecha & "', " & para.id_usuario & ", " & Val(t_subtotal) & ", " & Val(t_nograbado) & ", " & Val(t_iva) & ", " & Val(T_TOTAL) & ", 'A', " & cuentaact & ", '" & cl_compvta.STOCK & "', '" & _
    cl_compvta.ctacte & "', '" & cl_compvta.grabado & "', '" & ep & "', '0000-00000000', 'Tq.Ctdo', " & para.cotizacion & ", " & Format(Val(T_TOTAL) / para.cotizacion, "######0.00") & ", '" & moneda & "', " & codvend & ", '" & _
    cl_compvta.venta & "', '" & contado & "', " & Val(t_perc) & ", 0, " & Val(t_perciva) & ", " & codact & ", " & Val(t_alicuotaib) & ", " & Val(t_alicuotaperciva) & ", 0 , '" & t_fecha & "', 0, 0, ' ', ' ', ' ', 0, " & Val(t_sucursal) & ", 'Tique Contado' , ' ', ' ', '00-00000000-0', 3, 0, " & para.z_actual & ")"
    
    cn1.Execute QUERY
      
      COSTOINV = 0
      For i = 1 To msf1.Rows - 1
        If Val(msf1.TextMatrix(i, 1)) > 1 Then
          Set cl_prod = New productos
          cl_prod.cargar (Val(msf1.TextMatrix(i, 1)))
          costo = cl_prod.precio_ult_compra
          Set cl_prod = Nothing
        Else
          costo = 0
        End If
        
        QUERY = "INSERT INTO vta_03([num_int], [RENGLON], [id_producto], [descripcion], [cantidad], [pu], [importe], [tasaiva], [impuesto], [costo], [cantidad_original], [tunidad], [pu_final], [tasaib])"
        QUERY = QUERY & " VALUES (" & numint & ", " & Val(msf1.TextMatrix(i, 0)) & ", " & Val(msf1.TextMatrix(i, 1)) & ", '" & msf1.TextMatrix(i, 2) & " ', " & Val(msf1.TextMatrix(i, 3)) & ", " & Val(msf1.TextMatrix(i, 5)) & ", " & Val(msf1.TextMatrix(i, 7)) & ", " & Val(msf1.TextMatrix(i, 6)) & ", 0, " & costo & ", " & Val(msf1.TextMatrix(i, 3)) & ", '" & msf1.TextMatrix(i, 4) & "', " & Val(msf1.TextMatrix(i, 8)) & ", " & Val(msf1.TextMatrix(i, 10)) & ")"
        cn1.Execute QUERY
      
        If cl_compvta.STOCK <> "N" Then
           QUERY = "INSERT INTO stk_01([fecha], [id_producto], [cantidad], [ubicacion], [comprobante], [descripcion], [num_mov_int], [modulo])"
           QUERY = QUERY & " VALUES ('" & t_fecha & "', " & Val(msf1.TextMatrix(i, 1)) & ", " & msf1.TextMatrix(i, 3) & ", '" & cl_compvta.STOCK & "', '" & cl_compvta.abreviatura & t_letra & Format$(t_sucursal, "0000") & "-" & Format$(t_numcomp, "00000000") & "', 'Tique Contado', " & numint & ",'V'" & ")"
           cn1.Execute QUERY
          
           If cl_compvta.STOCK = "E" Then
             c = Val(msf1.TextMatrix(i, 3))
             COSTOINV = COSTOINV + (costo * Val(msf1.TextMatrix(i, 3)))
           Else
             c = -Val(msf1.TextMatrix(i, 3))
             COSTOINV = COSTOINV - (costo * Val(msf1.TextMatrix(i, 3)))
           End If
           q = "update a2 set [stock] = [stock] + " & c & " where [id_producto] = " & Val(msf1.TextMatrix(i, 1))
           cn1.Execute q
        
        End If
        
        If cl_compvta.venta <> "N" Then
           ultvta = t_letra & Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numcomp), "00000000") & " | " & "Tique Ctdo." & " | " & t_fecha & " | " & Format$(Val(msf1.TextMatrix(i, 4)), "#####0.00")
           QUERY = "update a2 set  [ultima_venta]='" & ultvta & "'"
           QUERY = QUERY & " where [id_producto]= " & Val(msf1.TextMatrix(i, 1))
           cn1.Execute QUERY
        End If
      Next i
      
      
      'actualizo tasa de iva
      For i = 1 To 7
        If Val(vta_facturacion2.msf1.TextMatrix(i, 1)) > 0 Then
          QUERY = "INSERT INTO vta_09([num_int], [tasa_iva], [iva], [neto], [tipo_iva])"
          QUERY = QUERY & " VALUES (" & numint & ", " & Val(vta_facturacion2.msf1.TextMatrix(i, 0)) & ", " & Val(vta_facturacion2.msf1.TextMatrix(i, 2)) & ", " & Val(vta_facturacion2.msf1.TextMatrix(i, 1)) & ", " & tiporespiva & ")"
          cn1.Execute QUERY
          
        End If
      Next i
     
     
        'cobranza
        Call grabaformapago
        
      

    'contabilidad
    If Generaasientosauto Then
     If cl_compvta.contabilidad <> "N" Then
         numintcgr = saca_ultnumero_int_comp("G")

         cta = para.cuenta_caja
         
         u1 = cl_compvta.contabilidad
          
         If u1 = "D" Then
           u2 = "H"
         Else
           u2 = "D"
         End If
         
         Set rs = New ADODB.Recordset
         q = "select * from c_01 where [id_cuenta] = " & cta
         rs.Open q, cn1
         If Not rs.EOF And Not rs.BOF Then
           dcta = rs("descripcion")
         Else
           dcta = "Cuenta Inexistente"
         End If
         Set rs = Nothing
         
         'grabo asiento
         QUERY = "INSERT INTO c_02([num_interno], [fecha], [descripcion], [modulo], [num_mov_int], [debe], [haber], [id_USUARIO], [observaciones])"
         QUERY = QUERY & " VALUES (" & numintcgr & " ,'" & t_fecha & "', '[Ventas] " & cl_compvta.abreviatura & " " & t_letra & Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numcomp), "00000000") & "', 'V', " & numint & ", " & Val(T_TOTAL) & ", " & Val(T_TOTAL) & ", " & para.id_usuario & ", 'Tique Ctdo.')"
         cn1.Execute QUERY
      
         
         'ingresa forma de pago
          ic = 1
          For i = 1 To fsc_formapago.msf2.Rows - 1
               cta = Val(fsc_formapago.msf2.TextMatrix(i, 9))
               im = Format(Val(fsc_formapago.msf2.TextMatrix(i, 6)), "######0.00")
               dcta = fsc_formapago.msf2.TextMatrix(i, 3)
               QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
               QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & cta & ", '" & u1 & "', " & im & ", '" & dcta & "')"
               'MsgBox (QUERY)
               cn1.Execute QUERY
               ic = ic + 1
          Next i
         
         
         'ic = 1
         'cuenta madre ctacte o caja
         'QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         'QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & cta & ", '" & u1 & "', " & Val(t_total) & ", '" & dcta & "')"
         
        ' cn1.Execute QUERY
         'ic = ic + 1
      
         If Val(t_nograbado) > 0 Then
           'cuenta nogbra
           QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_conceptos_nograbados & ", '" & u2 & "', " & Val(t_nograbado) & ", 'No Grabado')"
           cn1.Execute QUERY
           ic = ic + 1
         End If
                   
         If Val(t_perc) > 0 Then
           'cuenta perc
           QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_perc_IB & ", '" & u2 & "', " & Val(t_perc) & ", 'Perc. IB')"
           cn1.Execute QUERY
           ic = ic + 1
         End If
          
          If Val(t_perciva) > 0 Then
           'cuenta perc
           QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_perc_iva & ", '" & u2 & "', " & Val(t_perciva) & ", 'Perc. IVA')"
           cn1.Execute QUERY
           ic = ic + 1
         End If
         
         If Val(t_iva) > 0 And cl_compvta.grabado <> "N" Then
           'cuenta perc
           QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_iva_ventas & ", '" & u2 & "', " & Val(t_iva) & ", 'IVA')"
           cn1.Execute QUERY
           ic = ic + 1
         End If
         
         'contrapartida
         If cl_compvta.grabado = "N" Then
           importe = Val(T_TOTAL)
         Else
           importe = Val(t_subtotal)
         End If
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & cuentaact & ", '" & u2 & "', " & importe & ", '" & "Ventas" & "')"
         cn1.Execute QUERY
         ic = ic + 1
      
      End If
      
      
      If COSTOINV <> 0 Then
         If COSTOINV > 0 Then
           u1 = "H"
           u2 = "D"
         Else
           u2 = "H"
           u1 = "D"
           COSTOINV = -COSTOINV
         End If
         tot = COSTOINV
         If cl_compvta.contabilidad = "N" Then
          'realizo asiento de costo aunque el doc. no mueva contabilidad
          numintcgr = saca_ultnumero_int_comp("G")
          ic = 1
          QUERY = "INSERT INTO c_02([num_interno], [fecha], [descripcion], [modulo], [num_mov_int], [debe], [haber], [id_USUARIO], [observaciones])"
          QUERY = QUERY & " VALUES (" & numintcgr & " ,'" & t_fecha & "', '[Ventas] " & cl_compvta.abreviatura & " " & t_letra & Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numcomp), "00000000") & "', 'V', " & numint & ", " & tot & ", " & tot & ", " & para.id_usuario & ", 'Tique Contado')"
          cn1.Execute QUERY
        
         End If
                   
         ic = ic + 1
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_inventario & ", '" & u2 & "', " & Format(COSTOINV, "#####0.00") & ", 'Inventario')"
         cn1.Execute QUERY
         
         ic = ic + 1
                           
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_costo & ", '" & u1 & "', " & Format(COSTOINV, "######0.00") & ", '" & "Costo Merc." & "')"
         cn1.Execute QUERY
      End If
    End If
      
      cn1.CommitTrans
      Set rs = Nothing
      Set cl_compvta = Nothing
      Set cl_cli = Nothing

      
      
      
      Call INICIALIZA2(Me)
      Call armagrid
      Call fsc_formapago.armagrid2
      t_sucursal = Format$(glo.sucursal, "0000")
      t_letra = "B"
      t_fecha = Format$(Now, "dd/mm/yyyy")
      estadotique = "C" 'cerrado
      t_numcomp.Enabled = True
      t_numcomp.SetFocus
Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos y vuelva a repetir la operacion")
  

End Sub
Private Sub msf1_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
 If estadotique = "A" Then
  J = MsgBox("Anula emision del Tique", 4)
  If J = 6 Then
    Call anulatique
  End If
 
 End If
End If


End Sub
Sub anulatique()
 'anula tique
 If gprueba = 0 Then
  Fiscaltq.CancelarComprobante
 
 End If
 Label2.Visible = False
 estadotique = "C"
 Call iniciatique
End Sub
Sub iniciatique()
  Call armagrid
  t_numcomp = ""
  T_TOTAL = ""
  t_numcomp.Enabled = True
  t_numcomp.SetFocus
End Sub
Private Sub msf1_LostFocus()
'Call barraesag(Me)
msf1.FocusRect = flexFocusLight
Me.KeyPreview = True
Call CALCULATOTALES
End Sub





Private Sub t_iva_LostFocus()
Call sacatotales

End Sub

Private Sub t_nograbado_LostFocus()
Call sacatotales

End Sub


Private Sub t_numcomp_GotFocus()
Me.StatusBar1.Panels.item(1) = "[F11] Abre Tique - [ESC] Sale Tique "

End Sub

Private Sub t_numcomp_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF11 Then
   'abre prueba
   t_limite = t_limite.Tag
   gprueba = 1
   t_sucursal = Format$(glo.sucursal, "0000")
   q = "select * from vta_06 where [sucursal] = " & Val(t_sucursal) & " and [id_tipocomp] = 310"
   Set rs = New ADODB.Recordset
   rs.Open q, cn1
   If Not rs.EOF And Not rs.BOF Then
         'continua
          Label2 = "Tique Abierto"
          Label2.Visible = True
      
           estadotique = "A"
           msf1.Enabled = True
           msf1.SetFocus
           t_numcomp = Format$(rs("ult_num_B") + 1, "00000000")
           t_numcomp.Enabled = False
           
           
            fsc_tiqueNF1.t_renglon = ""
            fsc_tiqueNF1.t_cantidad = ""
            fsc_tiqueNF1.t_pu = ""
            fsc_tiqueNF1.t_importe = ""
            fsc_tiqueNF1.Show
     Else
        MsgBox ("Error al generar el comprobante")
    End If
   
   
   
End If

End Sub

Sub abretique2()
  
  On Error GoTo DepuraErrores
 
  If Not Fiscaltq.Inicializar Then
    Err.Raise Fiscaltq.Error, "", Fiscaltq.ErrorDesc
  End If
  
  Fiscaltq.CancelarComprobante
    
  
  If Not Fiscaltq.AbrirComprobante(10) Then
     Err.Raise Fiscaltq.Error, "", Fiscaltq.ErrorDesc
  End If
  
  estadotique = "A"
  msf1.SetFocus
  
  Exit Sub

DepuraErrores:
  Fiscaltq.Finalizar
  MsgBox Fiscaltq.ErrorDesc
End Sub

Sub item()
If Not fiscal.ImprimirItem2g("Item 1", 1, 0.1, 21, 0, IFUniversal.Gravado, "0", 1, "7790001001054", "", IFUniversal.unidad) Then
     Err.Raise fiscal.Error, "", fiscal.ErrorDesc
  End If
  
  If Not fiscal.ImprimirDescuentoGeneral("Descuento General", 0.01) Then
     Err.Raise fiscal.Error, "", fiscal.ErrorDesc
  End If
  
  If Not fiscal.ImprimirPago2g("Efectivo", 5, "", IFUniversal.Efectivo, 1, "", "") Then
     Err.Raise fiscal.Error, "", fiscal.ErrorDesc
  End If
  
  fiscal.CerrarComprobante
  
  fiscal.Finalizar
  
  MsgBox ("Comprobante impreso exitosamente")
End Sub

Private Sub t_numcomp_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub t_numcomp_LostFocus()
c_actividad.ListIndex = buscaindice(c_actividad, sacaactividadsucursal(Val(t_sucursal)))
End Sub

Private Sub t_subtotal_LostFocus()
Call sacatotales
End Sub
Sub sacatotales()
t_subtotal = Format$(Val(t_subtotal), "######0.00")
't_nograbado = Format$(Val(t_nograbado), "######0.00")
't_perc = Format$(Val(t_perc), "######0.00")
t_iva = Format$(Val(t_iva), "######0.00")
't_perciva = Format$(Val(t_perciva), "######0.00")
T_TOTAL = Format$(Val(T_TOTAL), "######0.00")
End Sub

Private Sub t_sucursal_GotFocus()
t_sucursal = Format$(glo.sucursalf, "0000")
End Sub

Private Sub t_sucursal_LostFocus()
Call inicia
End Sub

Private Sub t_total_LostFocus()
T_TOTAL = Format$(T_TOTAL, "######0.00")
End Sub

