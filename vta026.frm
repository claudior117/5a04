VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_selremitos2 
   BackColor       =   &H00E0E0E0&
   Caption         =   "SELECCION DE REMITOS A FACTURAR"
   ClientHeight    =   5865
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   6210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5865
   ScaleWidth      =   6210
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   3615
      Begin VB.TextBox t_seleccionados 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Remitos Seleciionados:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "REMITOS PENDIENTES"
      Height          =   4215
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5895
      Begin MSFlexGridLib.MSFlexGrid msf1 
         Height          =   3855
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   6800
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FillStyle       =   1
         SelectionMode   =   1
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   4440
      TabIndex        =   2
      Top             =   4440
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta026.frx":0000
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
         Picture         =   "vta026.frx":0882
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
      Top             =   5610
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   10583
            MinWidth        =   10583
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_selremitos2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim habilitafacturaremito As Boolean

Function habilita() As Boolean
 If msf1.Rows > 1 Then
   h = 0
   For i = 1 To msf1.Rows - 1
      If msf1.TextMatrix(i, 0) = "**" Then
         h = 1
         i = msf1.Rows
      End If
   Next i
   If h = 0 Then
     habilita = False
     btnacepta.Enabled = False
   Else
     habilita = True
     btnacepta.Enabled = True
   End If
 Else
   habilita = False
   btnacepta.Enabled = False
 End If

End Function
Sub limpia()
Call armagrid
btnacepta.Enabled = False
habilitafacturaremitos = habilita
End Sub

Private Sub btnacepta_Click()
Call armafactura
End Sub

Private Sub btnsale_Click()
Me.Hide
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 5
msf1.AllowUserResizing = flexResizeNone
msf1.FixedCols = 0
msf1.SelectionMode = flexSelectionByRow
msf1.FocusRect = flexFocusNone
msf1.ColWidth(0) = 300
msf1.ColWidth(1) = 1400
msf1.ColWidth(2) = 1400
msf1.ColWidth(3) = 800
msf1.ColWidth(4) = 1200
msf1.TextMatrix(0, 1) = "Nro. Comprobante"
msf1.TextMatrix(0, 2) = "Fecha"
msf1.TextMatrix(0, 3) = "Tipo"
msf1.TextMatrix(0, 4) = "Nro. Interno"
For i = 0 To 3
 msf1.ColAlignment(i) = 1 'izq
Next i

msf1.FocusRect = flexFocusNone

End Sub

Private Sub Form_Activate()
Call cuenta
End Sub

Private Sub Form_Load()
Load vta_cc_detalle
Call limpia

End Sub

Sub carga()
   Call limpia
   q = "select * from vta_02 where [id_tipocomp] > 40 and [id_tipocomp] < 50 and estado = 'S' and [id_cliente] = " & vta_COMPVARIOS.c_prov.ItemData(vta_COMPVARIOS.c_prov.ListIndex)
   Set rs = New ADODB.Recordset
   rs.Open q, cn1
   While Not rs.EOF
     nc = Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
     If rs("id_tipocomp") = 45 Then
        t = "R"
     Else
        t = "D"
     End If
     F = Format$(rs("fecha"), "dd/mm/yyyy")
     msf1.AddItem "" & Chr$(9) & nc & Chr$(9) & F & Chr$(9) & t & Chr$(9) & rs("num_int")
     rs.MoveNext
   Wend
   
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Cancel = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload vta_cc_detalle
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(1) = "[Barra] Marca - [F5] Todos - [F9] Arma Factura - "
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF9 Then
  Call armafactura
End If

If KeyCode = vbKeyF5 Then
  If msf1.Rows > 1 Then
    For i = 1 To msf1.Rows - 1
      If msf1.TextMatrix(i, 0) = "**" Then
          msf1.TextMatrix(i, 0) = ""
      Else
         msf1.TextMatrix(i, 0) = "**"
      End If
    Next i
  End If
  habilitafacturaremito = habilita
  Call cuenta
End If

End Sub
Sub cuenta()
 If msf1.Rows > 1 Then
   h = 0
   For i = 1 To msf1.Rows - 1
      If msf1.TextMatrix(i, 0) = "**" Then
         h = h + 1
      End If
   Next i
    
 Else
  h = 0
 End If
 t_seleccionados = h
End Sub
Sub armafactura()
  vta_COMPVARIOS.armagrid
  i = 0
  listaremitos = ""
  'ubica = 0
  For i = 1 To msf1.Rows - 1
    If msf1.TextMatrix(i, 0) = "**" Then
       If msf1.TextMatrix(i, 3) = "R" Then
        'agrega linea remitos
        Call agregalinea(i)
        'listaremitos = listaremitos & "Rt " & Mid$(facturas.List(i), 6, 8) & "- "
       End If
    End If
  Next i
  'FACTURACION!t_remitos = listaremitos
  
  For i = 1 To vta_COMPVARIOS.msf1.Rows - 1
     vta_COMPVARIOS.msf1.TextMatrix(i, 0) = i
  Next i
  
  vta_COMPVARIOS.CALCULATOTALES
  Me.Hide

End Sub

'FIXIT: Declare 'k' con un tipo de datos de enlace en tiempo de compilación                FixIT90210ae-R1672-R1B8ZE
Sub agregalinea(ByVal k)
Dim CANTIDAD2 As Double
Dim PRECIOUNITARIO2 As Double
Dim IMPINTERNO2 As Double
Dim IMPORTE2 As Double
Dim PRECIOfinal2 As Double
Dim J As Single

'BUSCO REMITO
q = "select * from vta_02 where [num_int] = " & Val(msf1.TextMatrix(k, 4))
Set rs = New ADODB.Recordset
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
   nms = rs("num_int")
   monedar = rs("moneda")
   'BUSCA PROD. PEND. DE FACTURACION EN EL REMITO
   q = "select * from vta_03 where [num_int] = " & nms & " and [cantidad] > 0"
   Set rs1 = New ADODB.Recordset
   rs1.Open q, cn1
   While Not rs1.EOF
       CANTIDAD2 = 0
       PRECIOUNITARIO2 = 0
       PRECIOfinal2 = 0
       IMPINTERNO2 = 0
       IMPORTE2 = 0
       J = -1
       
     '  q = "select * from vta_03 where [num_int] = " & nms & " and [renglon] = " & rs1("secuencia")
      ' Set rs2 = New ADODB.Recordset
       'rs2.Open q, cn1
     '  If Not rs2.EOF And Not rs2.BOF Then
        If rs1("id_producto") > 1 Then
                J = BUSCOPRODFACTURA(rs1("id_producto"))
        End If
            
            'SI ENCUENTRA EL PROD. EN LA FACTURA
        If J <> -1 Then
               'SACA LOS valores de la factura
                CANTIDAD2 = Val(vta_facturacion.msf1.TextMatrix(J, 3))
                PRECIOUNITARIO2 = Val(vta_facturacion.msf1.TextMatrix(J, 4))
                PRECIOfinal2 = Val(vta_facturacion.msf1.TextMatrix(J, 8))
                IMPINTERNO2 = 0 'Val(vta_facturacion.msf1.TextMatrix(J, 6))
                IMPORTE2 = Val(vta_facturacion.msf1.TextMatrix(J, 6))
                vta_COMPVARIOS.msf1.RemoveItem J
        End If

            
        'si el importe del remito es 0 totma el de la b.d.
        If rs1("pu") = 0 Then
             Set rs3 = New ADODB.Recordset
             q = "select * from a2, g4 where [id_producto] = " & rs1("id_producto") & " and [cod_tasaiva] = [id_tasaiva]"
             rs3.Open q, cn1
             If Not rs3.EOF And Not rs3.BOF Then
                        I2 = rs3("pu")
                        ii = rs3("IMPUESTO")
                        MONEDALINEA = rs3("moneda")
                         ti = rs3("TASA")
                         f2 = rs3("precio_final")
              Else
                        I2 = 0
                        ii = 0
                        MONEDALINEA = monedar
                        ti = 0
                        f2 = 0
              End If
              Set rs3 = Nothing
         Else
                I2 = rs1("pu")
                ii = rs1("IMPUESTO")
                MONEDALINEA = monedar
                ti = rs1("TASAIVA")
                f2 = rs1("pu_final")
         End If

         Call CONVIERTEMONEDA
       

       CANTIDAD2 = CANTIDAD2 + rs1("CANTIDAD")
       IMPINTERNO2 = IMPINTERNO2 + (ii * CANTIDAD2)
       IMPORTE2 = IMPORTE2 + (I2 * CANTIDAD2)
       PRECIOUNITARIO2 = (I2 - ii)
       final2 = (f2)
       cp = Format$(rs1("id_producto"), "00000")
       dp = rs1("DESCRIPCION")
        u = rs1("tunidad")
        ct = Format$(CANTIDAD2, "#####0.00")
        pu = Format$(PRECIOUNITARIO2, "#####0.00")
        dt = Format$(rs1("tasaiva"), "###0.00")
        im = Format$(IMPORTE2, "######0.00")
        III = Format$(IMPINTERNO2, "###0.000")
        ubica = vta_facturacion.msf1.Rows - 1
        F = Format$(final2, "######0.00")
        vta_COMPVARIOS.msf1.AddItem ubica & Chr$(9) & cp & Chr$(9) & dp & Chr$(9) & ct & Chr$(9) & pu & Chr$(9) & dt & Chr$(9) & im & Chr$(9) & F
     
     rs1.MoveNext
    Wend
    Set rs1 = Nothing
  End If
Set rs = Nothing
End Sub

Sub CONVIERTEMONEDA()
If vta_COMPVARIOS.Option3 = True Then
  m = "P"
Else
  m = "D"
End If


If m <> MONEDALINEA Then
 If MONEDALINEA = "P" Then
    I2 = I2 / Val(vta_COMPVARIOS.t_cotizacion)
    ii = ii / Val(vta_COMPVARIOS.t_cotizacion)
    f2 = f2 / Val(vta_COMPVARIOS.t_cotizacion)
 Else
    I2 = I2 * Val(vta_COMPVARIOS.t_cotizacion)
    ii = ii * Val(vta_COMPVARIOS.t_cotizacion)
    f2 = f2 * Val(vta_COMPVARIOS.t_cotizacion)
 End If
End If

End Sub

'FIXIT: Declare 'cp' con un tipo de datos de enlace en tiempo de compilación               FixIT90210ae-R1672-R1B8ZE
Function BUSCOPRODFACTURA(ByVal cp) As Integer
 'DEVUELVE 0 SI NO LO ENCUENTRA O EL NRO. DE RENGLON
 t = 1
 e = -1
 While t < vta_COMPVARIOS.msf1.Rows - 1
  If Val(vta_COMPVARIOS.msf1.TextMatrix(t, 1)) = cp Then
    e = t
    t = vta_COMPVARIOS.msf1.Rows
  End If
  t = t + 1
 Wend
 
 BUSCOPRODFACTURA = e
End Function
Sub BUSCOPROD()
Set rs3 = New ADODB.Recordset
q = "select * from a2 "
TBA2.Seek "=", dynprod.Fields("cod-producto")
If Not TBA2.NoMatch Then
  ii = TBA2.Fields("impuesto")
  If FACTURACION!tipofacturacion = "P" Then
    I2 = TBA2.Fields("precio-unitario")
  Else
    I2 = TBA2.Fields("precio-unitario2")
  End If
  MONEDALINEA = TBA2.Fields("MONEDA")
Else
  ii = 0
  I2 = 0
  MONEDALINEA = FACTURACION!moneda
  ti = 0
End If
End Sub


Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Rows > 1 Then
    If Val(msf1.TextMatrix(msf1.Row, 4)) > 0 Then
       vta_cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 4)
       vta_cc_detalle.Show
    End If
  End If
  habilitafacturaremito = habilita
End If

If KeyAscii = vbKeySpace Then
  If Val(msf1.TextMatrix(msf1.Row, 4)) > 0 Then
      If msf1.TextMatrix(msf1.Row, 0) = "**" Then
          msf1.TextMatrix(msf1.Row, 0) = ""
      Else
         msf1.TextMatrix(msf1.Row, 0) = "**"
      End If
  End If
  habilitafacturaremitos = habilita
End If


End Sub

Private Sub msf1_LostFocus()
msf1.FocusRect = flexFocusNone

End Sub




