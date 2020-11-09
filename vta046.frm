VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form vta_actudeuda 
   BackColor       =   &H00E0E0E0&
   Caption         =   "ACTUALIZACION DE DEUDA"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   4080
      TabIndex        =   23
      Top             =   1200
      Width           =   7695
      Begin VB.TextBox t_diasg 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6360
         MaxLength       =   10
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox t_tfm 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Dias Gracia:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   30
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Tasa Financiera Mensual:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   240
      TabIndex        =   20
      Top             =   1200
      Width           =   3735
      Begin VB.ComboBox c_sucursal 
         Height          =   315
         Left            =   2160
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Punto Venta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   8640
      TabIndex        =   14
      Top             =   600
      Width           =   3135
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pesos($)"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dolares(U$s)"
         Height          =   375
         Left            =   1680
         TabIndex        =   15
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   8640
      TabIndex        =   11
      Top             =   0
      Width           =   3135
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha vencimiento"
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha Comprobante"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   1335
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   2400
      TabIndex        =   10
      Top             =   360
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   140902401
      CurrentDate     =   38754
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5295
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9340
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   8295
      Begin VB.TextBox t_fechacorte 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6960
         MaxLength       =   10
         TabIndex        =   26
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Height          =   255
         Left            =   7800
         Picture         =   "vta046.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   17
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Text            =   "c_prov"
         Top             =   240
         Width           =   6135
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Corte:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5760
         TabIndex        =   27
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2880
         TabIndex        =   18
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Cliente:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   4
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta046.frx":0372
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "vta046.frx":0BF4
         Style           =   1  'Graphical
         TabIndex        =   5
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
      TabIndex        =   3
      Top             =   8235
      Width           =   11880
      _ExtentX        =   20955
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
            TextSave        =   "14/09/2012"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:30"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080FFFF&
      Caption         =   $"vta046.frx":1476
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   240
      TabIndex        =   28
      Top             =   7320
      Width           =   9375
   End
End
Attribute VB_Name = "vta_actudeuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim saldoanterior As Double
Sub carga()
  QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
  QUERY = QUERY & " VALUES ('Emision estado Cuenta: " & c_prov.ItemData(c_prov.ListIndex) & "', " & para.id_usuario & ", 'V', 0, '" & Now & "', '" & Left$(c_prov, 50) & "', 9, " & c_prov.ItemData(c_prov.ListIndex) & ")"
  cn1.BeginTrans
  cn1.Execute QUERY
  cn1.CommitTrans
  
  Call armagrid
  sa = 0
  da = 0
  ha = 0
  
  sao = 0
  dao = 0
  hao = 0
  t2 = 0
  If t_fecha <> "" Then
     q = "select * from vta_02 where [id_cliente] = " & c_prov.ItemData(c_prov.ListIndex) & " and [cta_cte] <> 'N' " & " and [contado] = " & "'N' "
     
     If c_sucursal.ListIndex > 0 Then
        q = q & " and [sucursal_ingreso] = " & Val(c_sucursal)
     End If
     
     If Option1 = True Then
        q = q & " and datevalue([fecha]) < datevalue('" & t_fecha & "')"
     Else
        q = q & " and datevalue([fecha_vto]) < datevalue('" & t_fecha & "')"
     End If
    Set rs = New ADODB.Recordset
    rs.Open q, cn1
    While Not rs.EOF
     If Option4 = True Then
      If rs("moneda") = "P" Then
        t = rs("total")
        t2 = rs("total_otra_moneda")
      Else
        t = rs("total_otra_moneda")
        t2 = rs("total")
      End If
     Else
      If rs("moneda") = "D" Then
        t = rs("total")
        t2 = rs("total_otra_moneda")
      Else
        t = rs("total_otra_moneda")
        t2 = rs("total")
      End If
     End If
     
     If rs("cta_cte") = "D" Then
        da = da + t
        dao = dao + t2
     Else
        ha = ha + t
        hao = hao + t2
     End If
     rs.MoveNext
    Wend
    sa = da - ha
    sao = dao - hao
  End If
  
  saldoanterior = sa
  saldoanterioro = sao
  If Check1 = 0 Then
    msf1.AddItem t_fecha & Chr(9) & "" & Chr(9) & "Saldo Ant." & Chr(9) & "" & Chr(9) & Format$(da, "######0.00") & Chr(9) & Format$(ha, "######0.00") & Chr(9) & Format$(sa, "######0.00")
  Else
    msf1.AddItem t_fecha & Chr(9) & "" & Chr(9) & "Saldo Ant." & Chr(9) & "" & Chr(9) & Format$(da, "######0.00") & Chr(9) & Format$(ha, "######0.00") & Chr(9) & Format$(sa, "######0.00") & Chr(9) & Format$(sao, "######0.00")
  End If
  q = "select * from vta_02, vta_06 where [id_cliente] = " & c_prov.ItemData(c_prov.ListIndex) & " and vta_02.[cta_cte] <> 'N'  and vta_02.[id_tipocomp] = vta_06.[id_tipocomp]  " & " and [contado] = " & "'N' and vta_02.[sucursal_ingreso] = vta_06.[sucursal]"
  If t_fecha <> "" Then
    If Option1 = True Then
       q = q & " and datevalue([fecha]) >= datevalue('" & t_fecha & "')"
    Else
       q = q & " and datevalue([fecha_vto]) >= datevalue('" & t_fecha & "')"
    End If
  End If
    
  If t_fecha2 <> "" Then
    If Option1 = True Then
       q = q & " and datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
    Else
       q = q & " and datevalue([fecha_vto]) <= datevalue('" & t_fecha2 & "')"
    End If
  End If
    
  If c_sucursal.ListIndex > 0 Then
        q = q & " and [sucursal_ingreso] = " & Val(c_sucursal)
  End If
     
  If Option1 = True Then
     q = q & " order by [fecha], vta_02.[id_tipocomp], [num_comp]"
  Else
      q = q & " order by [fecha_vto], vta_02.[id_tipocomp], [num_comp]"
  End If
    
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  s = sa
  s2 = sao
  sao = ""
  dao = ""
  hao = ""
  While Not rs.EOF
    If Option1 = True Then
         f = rs("fecha")
    Else
         f = rs("fecha_vto")
    End If
     CTC = Format$(rs("vta_02.ID_TIPOCOMP"), "000")
     tc = rs("descripcion")
     nc = rs("letra") & " " & Format$(rs("vta_02.sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
     If Option4 = True Then
      If rs("vta_02.moneda") = "P" Then
        t = rs("total")
        t2 = rs("total_otra_moneda")
      Else
        t = rs("total_otra_moneda")
        t2 = rs("total")
      End If
     Else
       If rs("vta_02.moneda") = "D" Then
        t = rs("total")
        t2 = rs("total_otra_moneda")
      Else
        t = rs("total_otra_moneda")
        t2 = rs("total")
      End If
     End If
     If rs("cta_cte") = "D" Then
       d = Format$(t, "######0.00")
       h = ""
       dao = Format$(t2, "######0.00")
       hao = ""
     Else
       h = Format$(t, "######0.00")
       d = ""
       hao = Format$(t2, "######0.00")
       dao = ""
     
     End If
     s = Format$(Val(s) + Val(d) - Val(h), "######0.00")
     s2 = Format$(Val(s2) + Val(dao) - Val(hao), "######0.00")
     ni = rs("num_int")
     o = rs("observaciones")
     If rs("vta_02.id_tipocomp") = 1 Then
       If rs("estado_pago") = "P" Then
        ep = "Cancelado"
       Else
        If rs("vta_02.moneda") = "P" Then
          ic = rs("total")
        Else
          ic = rs("total_otra_moneda")
        End If
        If rs("saldo_impago02") < ic Then
          ep = "Parcial"
        Else
          ep = "Impago"
        End If
       End If
     Else
        ep = " "
     End If
     If Check1 = 0 Then
       msf1.AddItem f & Chr(9) & CTC & Chr(9) & tc & Chr(9) & nc & Chr(9) & d & Chr(9) & h & Chr(9) & s & Chr(9) & o & Chr(9) & ep & Chr(9) & ni
     Else
        msf1.AddItem f & Chr(9) & CTC & Chr(9) & tc & Chr(9) & nc & Chr(9) & d & Chr(9) & h & Chr(9) & s & Chr(9) & s2 & Chr(9) & o & Chr(9) & ep & Chr(9) & ni
     End If
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
  msf1.clear
  msf1.Rows = 1
  msf1.Cols = 10
  msf1.ColWidth(0) = 1300
  msf1.ColWidth(1) = 1500
  msf1.ColWidth(2) = 1700
  msf1.ColWidth(3) = 1000
  msf1.ColWidth(4) = 1200
  msf1.ColWidth(5) = 1200
  msf1.ColWidth(6) = 1200
  msf1.ColWidth(7) = 1200
  msf1.ColWidth(8) = 1000
  msf1.ColWidth(9) = 500
  msf1.TextMatrix(0, 0) = "Fecha"
  msf1.TextMatrix(0, 1) = "Op."
  msf1.TextMatrix(0, 2) = "Nro.Comprobante"
  msf1.TextMatrix(0, 3) = "Dias Mora"
  msf1.TextMatrix(0, 4) = "Importe"
  msf1.TextMatrix(0, 5) = "Tasa"
  msf1.TextMatrix(0, 6) = "Mora"
  msf1.TextMatrix(0, 7) = "Deuda"
  msf1.TextMatrix(0, 8) = "Num.Int."
  msf1.TextMatrix(0, 9) = " "
  For i = 0 To 2
    msf1.ColAlignment(i) = 1
  Next i
  For i = 3 To 8
    msf1.ColAlignment(i) = 9
  Next i
  
  
  

End Sub







Private Sub c_prov_LostFocus()
If c_prov.ListIndex < 0 Then
  If Val(c_prov) > 0 Then
    c_prov.ListIndex = buscaindice(c_prov, Val(c_prov))
  Else
    c_prov.ListIndex = 0
  End If
End If
End Sub

Private Sub c_sucursal_LostFocus()
If c_sucursal.ListIndex < 0 Then
  c_sucursal.ListIndex = 0
End If
End Sub

Private Sub cal1_DblClick()
t_fecha = cal1
cal1.Visible = False
End Sub

Private Sub cal1_LostFocus()
t_fecha = cal1
cal1.Visible = False
End Sub

Private Sub Command5_Click()
vta_clientes.t_id = c_prov.ItemData(c_prov.ListIndex)
vta_clientes.carga
vta_clientes.Show

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
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

Call carga_clientes(c_prov)
c_prov.ListIndex = 0

Call carga_SUCURSALES(c_sucursal)
c_sucursal.AddItem "<Todas>", 0
c_sucursal.ListIndex = 0
t_sucursal = Format$(glo.sucursal, "0000")


Call armagrid
Call barraesag(Me)
cal1.Visible = False
Option1 = True
If para.moneda = "P" Then
 Option4 = True
Else
 Option3 = True
End If
Load vta_clientes

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload vta_clientes
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F7] Imprime - [ENTER] Visualiza Comprobante - [F11] Excel "
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

     If Check1 = 0 Then
      c(0) = 10
      c(1) = 0
      c(2) = 2
      c(3) = 3
      c(4) = 4
      c(5) = 5
      c(6) = 6
      c(7) = 7
      For i = 8 To 14
        c(i) = -1
      Next i
    Else
      
      c(0) = 11
      c(1) = 0
      c(2) = 2
      c(3) = 3
      c(4) = 4
      c(5) = 5
      c(6) = 6
      c(7) = 7
      c(8) = 8
      
      For i = 9 To 14
        c(i) = -1
      Next i
     End If
     
     If Check2 = 0 Then
        Call imprimegrid(msf1, c(), "ESTADO DE CUENTA", "", "Cliente: " & c_prov, "Periodo: " & t_fecha & "  " & t_fecha2, 85, 7, True, False)
     Else
        Call imprimegrid(msf1, c(), "ESTADO DE CUENTA", "", "Cliente: " & c_prov, "Periodo: " & t_fecha & "  " & t_fecha2, 50, 9, True, False, "H")
     End If
    End If
         
  End If
  
End If

If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If
End Sub
Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    Load vta_cc_detalle
    vta_cc_detalle.T_IDPROV = c_prov.ItemData(c_prov.ListIndex)
    vta_cc_detalle.t_prov = c_prov
    vta_cc_detalle.t_sucursal = Mid$(msf1.TextMatrix(msf1.Row, 3), 3, 4)
    vta_cc_detalle.t_letra = Mid$(msf1.TextMatrix(msf1.Row, 3), 1, 1)
    vta_cc_detalle.t_numcomp = Mid$(msf1.TextMatrix(msf1.Row, 3), 8, 8)
    vta_cc_detalle.t_tipocomp = msf1.TextMatrix(msf1.Row, 1)
    If Check1 = 0 Then
      vta_cc_detalle.t_NUMINT = msf1.TextMatrix(msf1.Row, 9)
    Else
       vta_cc_detalle.t_NUMINT = msf1.TextMatrix(msf1.Row, 10)
    End If
    vta_cc_detalle.Show
  End If
End If
End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight

End Sub

Private Sub Option3_Click()
Check1.Caption = "Muestra Saldo en $"
End Sub

Private Sub t_fecha_DblClick()
cal1.Visible = True
End Sub

Private Sub t_fecha_GotFocus()
t_fecha = ""
End Sub

Private Sub t_fecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call carga
End If
End Sub

Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
    t_fecha = ""
  End If
End If
  
End Sub
