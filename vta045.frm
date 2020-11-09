VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form vta_planillaventa 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PLANILLA DE VENTA CONTADO"
   ClientHeight    =   8805
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox t_numfila 
      Height          =   285
      Left            =   360
      TabIndex        =   18
      Text            =   "Text2"
      Top             =   7800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
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
      Left            =   7440
      MaxLength       =   10
      TabIndex        =   6
      Top             =   7560
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   6015
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10610
      _Version        =   393216
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   11535
      Begin VB.TextBox t_cantidad 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   7560
         MaxLength       =   8
         TabIndex        =   12
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox t_pu 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   8640
         MaxLength       =   10
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox t_importe 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   9720
         MaxLength       =   11
         TabIndex        =   10
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox t_basico 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   120
         MaxLength       =   20
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox t_detalle 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   1560
         MaxLength       =   49
         TabIndex        =   8
         Top             =   600
         Width           =   5775
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Detalle"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Cantidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7440
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Pu"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8640
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Importe"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9840
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Basico"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10080
      TabIndex        =   2
      Top             =   7440
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta045.frx":0000
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
         Picture         =   "vta045.frx":0882
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
      Top             =   8550
      Width           =   11910
      _ExtentX        =   21008
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
            Object.Width           =   14111
            MinWidth        =   14111
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "27/02/2015"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:36"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      Caption         =   "Total diario:"
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
      Left            =   5400
      TabIndex        =   7
      Top             =   7560
      Width           =   1935
   End
End
Attribute VB_Name = "vta_planillaventa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim saldoanterior As Double
Sub carga()
  
  Call armagrid
  T2 = 0
  q = "select * from vta_02, vta_03  where [id_tipocomp] = 500 and vta_02.[num_int] = vta_03.[num_int] "
  c = " and "
    
  Set rs = New ADODB.Recordset
  'MsgBox (q)
  rs.Open q, cn1
  While Not rs.EOF
       If rs("bultos") = 0 Then
         'sin facturar / pendiente de facturacion
         b = "S"
         
       Else
         'facturado
         b = "F"
       End If
       
       msf1.AddItem rs("id_producto") & Chr(9) & rs("descripcion") & Chr(9) & rs("pu_final") & Chr(9) & rs("cantidad") & Chr(9) & rs("importe") & Chr(9) & b
       rs.MoveNext
  Wend
  
End Sub



Private Sub btnacepta_Click()

  Call carga
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub

Sub armagrid()
'armar grilla
  msf1.Clear
  msf1.Rows = 1
  msf1.Cols = 6
  msf1.ColWidth(0) = 1200
  msf1.ColWidth(1) = 4000
  msf1.ColWidth(2) = 1200
  msf1.ColWidth(3) = 1200
  msf1.ColWidth(4) = 1400
  msf1.ColWidth(5) = 800
  msf1.TextMatrix(0, 0) = "Basico"
  msf1.TextMatrix(0, 1) = "Producto"
  msf1.TextMatrix(0, 2) = "Cantidad"
  msf1.TextMatrix(0, 3) = "PU Final"
  msf1.TextMatrix(0, 4) = "Importe"
  msf1.TextMatrix(0, 5) = "Estado"
  
  For i = 0 To 5
    msf1.ColAlignment(i) = 1
  Next i
  msf1.ColAlignment(1) = 9
 
  
  
End Sub










Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub

Sub limpia()
t_cantidad = ""
t_basico = ""
t_detalle = ""
t_pu = ""
t_importe = ""
t_ip = ""
t_unidad = ""
End Sub

Sub cargarenglon(t As String)
  ip = Val(t_ip)
  d = t_detalle
  cu = Format$(Val(t_cantidad), "######0.00")
  ti = Format$(c_tasa, "####0.00")
  u = RTrim$(t_unidad)
   'se cargan todos los comprobantes sin iva, en las que no se descrimina solo se imprimen con iva
  'If para.tipoprecioventa = 1 Then
   pu = Format$(Val(t_pu) / (1 + Val(c_tasa) / 100), "#####0.000")
  ' im = Format$(Val(pu) * Val(cu), "#####0.00")
   puf = Format$(Val(t_pu), "#####0.00")
  'Else
  ' pu = Format$(Val(t_pu), "#####0.000")
  ' im = Format$(Val(pu) * Val(cu), "#####0.00")
  'puf = Format$(Val(t_pu) * (1 + Val(c_tasa) / 100), "#####0.000")
  'End If
  cr = Format$(Val(t_costo) * Val(cu), "#####0.00")
  If t = "A" Then
    'nueva linea
    r = msf1.Rows
    msf1.AddItem r & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & cu & Chr(9) & u & Chr$(9) & pu & Chr(9) & ti & Chr(9) & im & Chr(9) & puf & Chr(9) & Chr(9) & cr & Chr(9) & Format$(Val(t_tasaib), "####0.00")
   
  Else
    r = Val(t_numfila)
    msf1.AddItem r & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & cu & Chr$(9) & u & Chr$(9) & pu & Chr(9) & ti & Chr(9) & im & Chr(9) & puf & Chr(9) & Chr(9) & cr & Chr(9) & Format$(Val(t_tasaib), "####0.00"), r
    msf1.RemoveItem r + 1
  End If
   
  s = 0
  v = 0
  For i = 1 To vta_facturacion.msf1.Rows - 1
    If r > 0 Then
      r = Val(msf1.TextMatrix(i, 7))
      s = s + r
      v = v + (r * Val(msf1.TextMatrix(i, 6)) / 100)
    End If
  Next i
 t_subtotal = s
 t_iva = v
 'Call sacatotales
 'Call sacaperc
 'Call sacatotales

  para.producto_sel = 0
End Sub


Sub busca(tipo As String)
'tipo = I por id_producto tipo = B por cod_barra
Set rs = New ADODB.Recordset
q = "select * from a2, g5, g12 where a2.[id_unidad] = g5.[id_unidad] and a2.[id_tasaib] = g12.[id_tasaib] "
If tipo = "I" Then
  q = q & " and [id_producto] = " & Val(t_basico)
Else
  q = q & " and [cod_barra] = '" & RTrim$(t_basico) & "'"
End If
rs.MaxRecords = 1
rs.Open q, cn1
If Not rs.BOF And Not rs.EOF Then
  t_detalle = rs("descripcion")
  t_pu = rs("precio_final")
  c_tasa.ListIndex = rs("cod_tasaiva")
  t_ip = rs("id_producto")
  t_unidad = rs("unidad")
  t_costo = rs("costoreal")
  t_tasaib = rs("tasaib")
  
Else
  MsgBox ("Producto no Ingresado")
  t_basico.SetFocus
  t_costo = 0
End If
Set rs = Nothing
End Sub


Sub cargap()
If IsNumeric(t_basico) Then
 If Val(t_basico) <= 1 Then
    t_basico = 1
    t_ip = 1
    t_detalle.Enabled = True
    t_detalle.SetFocus
 Else
    If Len(t_basico) <= 5 Then
       Call busca("I") 'busca por id. producto
    Else
       Call busca("B") 'busca por cod. barra
    End If
 End If
Else
  Call busca("B")
End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 2)
  'Case Is = 27
  '      Me.Hide
End Select

End Sub

Private Sub Form_Load()


Call armagrid
Call barraesag(Me)

'Load com_faltantes1
't_numfila = 0 registro nuevo otro numero mofifica fila
t_numfila = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload com_faltantes1
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F7] Imprime - [F8] Elimina  - [F11] Excel - [INS] Agrega  "
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
  Call nivel_acceso(1)
  If para.id_grupo_modulo_actual >= 4 Then
    J = MsgBox("Prepare Impresora y Confirme", 4)
    If J = 6 Then
     Dim c(15) As Double

      
      c(0) = 0
      c(1) = 1
      c(2) = 2
      c(3) = 3
      c(4) = 4
      c(5) = 5
      c(6) = 6
      c(7) = 7
      
      For i = 8 To 14
        c(i) = -1
      Next i
     
     Call imprimegrid(msf1, c(), "PLANILLA VENTA CONTADO", "", "", "FECHA: " & t_fecha, 85, 7, True, False)

    End If
         
  End If
  
End If

If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If

If KeyCode = vbKeyF8 Then
 If Val(msf1.TextMatrix(msf1.Row, 8)) > 0 Then
  J = MsgBox("Confirma Eliminar articulo [" & msf1.TextMatrix(msf1.Row, 3) & "] de la Planilla de Venta", 4)
  If J = 6 Then
    'Set cl_prod = New productos
    'Call cl_prod.sacafaltante(Val(msf1.TextMatrix(msf1.Row, 8)))
    'Call carga
  End If
 End If
End If

If KeyCode = vbKeyInsert Then
  com_faltantes1.t_renglon = 0
  com_faltantes1.Show
 
End If

End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight

End Sub



Private Sub t_basico_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[ENTER] Acepta - [ESC] Sale - [F8]Lista Precios  "

t_detalle.Enabled = False
If para.producto_sel > 0 Then
  t_basico = para.producto_sel
End If
End Sub

Private Sub t_basico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF8 Then
  vta_listaprecios.Show
End If
End Sub

Private Sub t_basico_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call cargap
End If
End Sub

Private Sub t_basico_LostFocus()
Call barraesag(Me)
End Sub

Private Sub t_importe_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Val(t_numfila) = 0 Then
   Call cargarenglon("A")
   t_basico.SetFocus
   
  Else
   Call cargarenglon("M")
   t_basico.SetFocus
   Me.Hide
  End If
  Call limpia
  
Else
  Call solonum(KeyAscii, 1)
End If
End Sub

End Sub
