VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_selremitos 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SELECCION DE REMITOS A FACTURAR"
   ClientHeight    =   6480
   ClientLeft      =   0
   ClientTop       =   345
   ClientWidth     =   6210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6480
   ScaleWidth      =   6210
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   4320
      Width           =   4095
      Begin VB.CheckBox Check1 
         Caption         =   "Utiliza Lista de Precios  para Armar Factura"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   $"vta014.frx":0000
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   5520
      Width           =   4095
      Begin VB.TextBox t_r2 
         Height          =   285
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox t_r1 
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "de"
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Renglones Utilizados"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   4920
      Width           =   4095
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
         ToolTipText     =   $"vta014.frx":00A7
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
      Top             =   5160
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta014.frx":019D
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
         Picture         =   "vta014.frx":0A1F
         Style           =   1  'Graphical
         TabIndex        =   3
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
      Top             =   6225
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
Attribute VB_Name = "vta_selremitos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim habilitafacturaremito As Boolean
Dim grecargocc As Single

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
t_r1 = 0

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

Set rs = New ADODB.Recordset
q = "select [recargo_cc], [precio_remito_factura] from g0 where [sucursal] = 0"
rs.Open q, cn1
If Not rs.BOF And Not rs.EOF Then
  grecargocc = rs("recargo_cc")
  Check1 = rs("precio_remito_factura")
Else
  grecargocc = 0
  Check1 = 0
End If
Set rs = Nothing
End Sub

Sub carga()
   Call limpia
   q = "select [num_int], [sucursal], [num_comp], [id_tipocomp], [fecha]  from vta_02 where [id_tipocomp] > 40 and [id_tipocomp] < 50  and [id_cliente] = " & vta_facturacion.c_prov.ItemData(vta_facturacion.c_prov.ListIndex) & " and estado = 'S'"
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
Me.StatusBar1.Panels.item(1) = "[Barra] Marca - [F5] Todos - [F9] Arma Factura - "
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF9 Then
  Call armafactura
  Me.Hide
  vta_facturacion.msf1.SetFocus
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
  
  Call armafactura
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
  vta_facturacion.armagrid
  t_r1 = 0
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
  
  For i = 1 To vta_facturacion.msf1.Rows - 1
     vta_facturacion.msf1.TextMatrix(i, 0) = i
  Next i
  
  vta_facturacion.CALCULATOTALES
  'Me.Hide

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
q = "select [num_int], [moneda], [id_tipocomp] from vta_02 where [num_int] = " & Val(msf1.TextMatrix(k, 4))
Set rs = New ADODB.Recordset
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
   nms = rs("num_int")
   monedar = rs("moneda")
   'BUSCA PROD. PEND. DE FACTURACION EN EL REMITO
   q = "select [id_producto], [pu], [impuesto], [tasaiva], [pu_final], [cantidad], [descripcion], [tunidad]  from vta_03 where [num_int] = " & nms & " and [cantidad] > 0"
   Set rs1 = New ADODB.Recordset
   rs1.Open q, cn1
   While Not rs1.EOF
       
       J = -1
       
       If rs1("id_producto") > 1 Then
                J = BUSCOPRODFACTURA(rs1("id_producto"))
       
        End If
            
            'SI ENCUENTRA EL PROD. EN LA FACTURA
        'datos en la factura
        If J <> -1 Then
               'SACA LOS valores de la factura
                CANTIDAD2 = Val(vta_facturacion.msf1.TextMatrix(J, 3))
                PRECIOUNITARIO2 = Val(vta_facturacion.msf1.TextMatrix(J, 4))
                PRECIOfinal2 = Val(vta_facturacion.msf1.TextMatrix(J, 8))
                IMPINTERNO2 = 0 'Val(vta_facturacion.msf1.TextMatrix(J, 6))
                IMPORTE2 = Val(vta_facturacion.msf1.TextMatrix(J, 6))
                If vta_facturacion.msf1.Rows > 2 Then
                   vta_facturacion.msf1.RemoveItem J
                 Else
                   Call vta_facturacion.armagrid
                End If
                
        Else
                CANTIDAD2 = 0
                PRECIOUNITARIO2 = 0
                PRECIOfinal2 = 0
                IMPINTERNO2 = 0
                IMPORTE2 = 0
                t_r1 = Val(t_r1) + 1
        End If

            
        'datos en el remito
        'si el importe del remito es 0 totma el de la b.d.
        If rs1("id_producto") > 1 And (rs1("pu") = 0 Or Check1 = 1) Then
             Set rs3 = New ADODB.Recordset
             q = "select [pu], [impuesto], [moneda], [tasa], [precio_final] from a2, g4 where [id_producto] = " & rs1("id_producto") & " and [cod_tasaiva] = [id_tasaiva]"
             rs3.Open q, cn1
             If Not rs3.EOF And Not rs3.BOF Then
                                  
                        I2 = rs3("pu")
                        ii = rs3("IMPUESTO")
                        MONEDALINEA = rs3("moneda")
                         ti = rs3("TASA")
                         f2 = rs3("precio_final")
                         If grecargocc > 0 Then
                          r = (I2 * grecargocc) / 100
                          I2 = Format(I2 + r, "#####0.00")
                          f2 = Format(I2 * (1 + (ti / 100)), "#####0.00")
                         End If
                         
              Else
                        I2 = 0
                        ii = 0
                        MONEDALINEA = monedar
                        ti = 0
                        f2 = 0
              End If
              'Set rs3 = Nothing
         Else
                I2 = rs1("pu")
                ii = rs1("IMPUESTO")
                MONEDALINEA = monedar
                ti = rs1("TASAIVA")
                f2 = rs1("pu_final")
         End If

         Call CONVIERTEMONEDA
         c2 = rs1("CANTIDAD")
         it2 = I2 * c2
       'nuevos datos a factura
       
       'nota db o remito
       'If rs("id_tipocomp") = 45 Then
            CANTIDAD3 = CANTIDAD2 + c2
      ' Else
      '      CANTIDAD3 = CANTIDAD2 - c2
      ' End If
       IMPINTERNO3 = IMPINTERNO2 + (ii * c2)
       PRECIOUNITARIO3 = I2
       IMPORTE3 = (PRECIOUNITARIO3 * CANTIDAD3)
               
       final2 = (f2)
       cp = Format$(rs1("id_producto"), "00000")
       dp = rs1("DESCRIPCION")
        u = rs1("tunidad")
        ct = Format$(CANTIDAD3, "#####0.00")
        pu = Format$(PRECIOUNITARIO3, "#####0.00")
        dt = Format$(rs1("tasaiva"), "###0.00")
        im = Format$(IMPORTE3, "######0.00")
        III = Format$(IMPINTERNO3, "###0.000")
        ubica = vta_facturacion.msf1.Rows - 1
        F = Format$(final2, "######0.00")
        vta_facturacion.msf1.AddItem ubica & Chr$(9) & cp & Chr$(9) & dp & Chr$(9) & ct & Chr$(9) & u & Chr$(9) & pu & Chr$(9) & dt & Chr$(9) & im & Chr$(9) & F
     
     rs1.MoveNext
    Wend
    'Set rs1 = Nothing
  End If
'Set rs = Nothing
End Sub



Sub agregalinea2(ByVal k)
Dim CANTIDAD2 As Double
Dim PRECIOUNITARIO2 As Double
Dim IMPINTERNO2 As Double
Dim IMPORTE2 As Double
Dim PRECIOfinal2 As Double
Dim J As Single

'BUSCO REMITO
   'BUSCA PROD. PEND. DE FACTURACION EN EL REMITO
   q = "select vta_03.num_int, vta_03.id_producto, vta_03.pu, a2.pu   from vta_02, vta_03, a2, g4 where vta_02.[num_int] = vta_03.[num_int] and  vta_03.[id_producto] = a2.[id_producto] and [cod_tasaiva] = g4.[id_tasaiva] and vta_02.[num_int] = " & k & " and [cantidad] > 0"
   Set rs1 = New ADODB.Recordset
   rs1.Open q, cn1
   nms = rs1("num_int")
   'monedar = rs1("moneda")
   
   While Not rs1.EOF
       
       J = -1
       
       If rs1("id_producto") > 1 Then
                J = BUSCOPRODFACTURA(rs1("id_producto"))
       End If
            
        'SI ENCUENTRA EL PROD. EN LA FACTURA
        'datos en la factura
        If J <> -1 Then
               'SACA LOS valores de la factura
                CANTIDAD2 = Val(vta_facturacion.msf1.TextMatrix(J, 3))
                PRECIOUNITARIO2 = Val(vta_facturacion.msf1.TextMatrix(J, 4))
                PRECIOfinal2 = Val(vta_facturacion.msf1.TextMatrix(J, 8))
                IMPINTERNO2 = 0 'Val(vta_facturacion.msf1.TextMatrix(J, 6))
                IMPORTE2 = Val(vta_facturacion.msf1.TextMatrix(J, 6))
                If vta_facturacion.msf1.Rows > 2 Then
                   vta_facturacion.msf1.RemoveItem J
                 Else
                   Call vta_facturacion.armagrid
                End If
                
        Else
                CANTIDAD2 = 0
                PRECIOUNITARIO2 = 0
                PRECIOfinal2 = 0
                IMPINTERNO2 = 0
                IMPORTE2 = 0
                t_r1 = Val(t_r1) + 1
        End If

            
        'datos en el remito
        'si el importe del remito es 0 totma el de la b.d.
        If rs1("id_producto") > 1 And (rs1("vta_03.pu") = 0 Or Check1 = 1) Then
                        I2 = rs1("a2.pu")
                        ii = rs1("a2.IMPUESTO")
                        MONEDALINEA = rs1("a2.moneda")
                         ti = rs1("TASA")
                         f2 = rs1("precio_final")
                         If grecargocc > 0 Then
                          r = (I2 * grecargocc) / 100
                          I2 = Format(I2 + r, "#####0.00")
                          f2 = Format(I2 * (1 + (ti / 100)), "#####0.00")
                         End If
                         
         Else
                        I2 = 0
                        ii = 0
                        MONEDALINEA = rs1("vta_02.moneda")
                        ti = 0
                        f2 = 0
              End If
      
       Call CONVIERTEMONEDA
       c2 = rs1("CANTIDAD")
        it2 = I2 * c2
       'nuevos datos a factura
       CANTIDAD3 = CANTIDAD2 + c2
       IMPINTERNO3 = IMPINTERNO2 + (ii * c2)
       PRECIOUNITARIO3 = I2
       IMPORTE3 = (PRECIOUNITARIO3 * CANTIDAD3)
               
       final2 = (f2)
       cp = Format$(rs1("vta_03.id_producto"), "00000")
       dp = rs1("vta_03.DESCRIPCION")
        u = rs1("tunidad")
        ct = Format$(CANTIDAD3, "#####0.00")
        pu = Format$(PRECIOUNITARIO3, "#####0.00")
        dt = Format$(rs1("tasaiva"), "###0.00")
        im = Format$(IMPORTE3, "######0.00")
        III = Format$(IMPINTERNO3, "###0.000")
        ubica = vta_facturacion.msf1.Rows - 1
        F = Format$(final2, "######0.00")
        vta_facturacion.msf1.AddItem ubica & Chr$(9) & cp & Chr$(9) & dp & Chr$(9) & ct & Chr$(9) & u & Chr$(9) & pu & Chr$(9) & dt & Chr$(9) & im & Chr$(9) & F
     
     rs1.MoveNext
    Wend
    Set rs1 = Nothing
  
End Sub





Sub CONVIERTEMONEDA()
If vta_facturacion.Option3 = True Then
  m = "P"
Else
  m = "D"
End If


If m <> MONEDALINEA Then
 If MONEDALINEA = "P" Then
    I2 = I2 / Val(vta_facturacion.t_cotizacion)
    ii = ii / Val(vta_facturacion.t_cotizacion)
    f2 = f2 / Val(vta_facturacion.t_cotizacion)
 Else
    I2 = I2 * Val(vta_facturacion.t_cotizacion)
    ii = ii * Val(vta_facturacion.t_cotizacion)
    f2 = f2 * Val(vta_facturacion.t_cotizacion)
 End If
End If

End Sub

'FIXIT: Declare 'cp' con un tipo de datos de enlace en tiempo de compilación               FixIT90210ae-R1672-R1B8ZE
Function BUSCOPRODFACTURA(ByVal cp) As Integer
 'DEVUELVE 0 SI NO LO ENCUENTRA O EL NRO. DE RENGLON
 t = 1
 e = -1
 While t < vta_facturacion.msf1.Rows
  If Val(vta_facturacion.msf1.TextMatrix(t, 1)) = cp Then
    e = t
    t = vta_facturacion.msf1.Rows
  
  
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
      Call armafactura
  End If
  habilitafacturaremitos = habilita
End If


End Sub

Private Sub msf1_LostFocus()
msf1.FocusRect = flexFocusNone

End Sub




