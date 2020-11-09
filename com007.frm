VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form con_veroc 
   BackColor       =   &H00E0E0E0&
   Caption         =   "INFORME DE ORDENES DE COMPRAS EMITIDAS"
   ClientHeight    =   8685
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   12405
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Usar Fecha"
      Height          =   975
      Left            =   240
      TabIndex        =   20
      Top             =   7320
      Width           =   7095
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha Prob. Entrega"
         Height          =   255
         Left            =   4440
         TabIndex        =   25
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha Vencimiento"
         Height          =   255
         Left            =   2160
         TabIndex        =   24
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha Emision"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cambiar"
      Height          =   975
      Left            =   8520
      TabIndex        =   18
      Top             =   7320
      Width           =   1095
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   120
         Picture         =   "com007.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5175
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9128
      _Version        =   393216
      AllowBigSelection=   0   'False
      FocusRect       =   0
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1935
      Left            =   240
      TabIndex        =   8
      Top             =   0
      Width           =   11535
      Begin VB.ComboBox c_zona 
         Height          =   315
         ItemData        =   "com007.frx":030A
         Left            =   8640
         List            =   "com007.frx":031A
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1440
         Width           =   1815
      End
      Begin VB.ComboBox c_pago 
         Height          =   315
         ItemData        =   "com007.frx":034B
         Left            =   8640
         List            =   "com007.frx":0358
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   600
         Width           =   2415
      End
      Begin VB.ComboBox c_cuenta 
         Height          =   315
         Left            =   8640
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox t_producto 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   8640
         MaxLength       =   20
         TabIndex        =   12
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Text            =   "c_prov"
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Zona:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6960
         TabIndex        =   22
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Estado:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   6960
         TabIndex        =   17
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Cuenta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6960
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Producto:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   6960
         TabIndex        =   13
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha  Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Proveedor:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   5
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "com007.frx":0387
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "com007.frx":0C09
         Style           =   1  'Graphical
         TabIndex        =   6
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
      TabIndex        =   4
      Top             =   8430
      Width           =   12405
      _ExtentX        =   21881
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
            TextSave        =   "20/11/2013"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "8:50"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "con_veroc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim indice As Long


Sub carga(par As Integer)
  Call armagrid
  q = "select * from a5, g2, a1 where [id_tipocomp] = [id_tipo_comp] and a5.[id_proveedor] = a1.[id_proveedor] "
  c = " and "
  If c_prov.ListIndex > 0 Then
     q = q & c & " a5.[id_proveedor] = " & c_prov.ItemData(c_prov.ListIndex)
  End If
  
 
  q = q & c & " [id_tipocomp] = 65 "
  
  
  Select Case par
  Case Is = 0 'fecha
   If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
   End If
  
   If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
   End If
   o = " order by [fecha]"
  
  Case Is = 1 'vencimiento
   If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha_vto]) >= datevalue('" & t_fecha & "')"
   End If
  
   If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha_vto]) <= datevalue('" & t_fecha2 & "')"
   End If
   o = " order by [fecha_vto]"
  
  Case Is = 2 'propbable entrega
    If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha_prob_entrega]) >= datevalue('" & t_fecha & "')"
    End If
  
    If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha_prob_entrega]) <= datevalue('" & t_fecha2 & "')"
    End If
    o = " order by [fecha_prob_entrega]"

  End Select
  
    
  If c_cuenta.ListIndex > 0 Then
    q = q & c & " [id_CUENTA] = " & c_cuenta.ItemData(c_cuenta.ListIndex)
  End If
  
  If c_pago.ListIndex > 0 Then
    q = q & c & " [estado] = '" & Mid$(c_pago, 2, 1) & "'"
  End If
  
 
  
  If c_zona.ListIndex > 0 Then
    q = q & c & " [zona] = " & c_zona.ListIndex
  End If
  
  
  
  q = q & o
  Set rs = New ADODB.Recordset
  rs.Open q, cn1, adOpenDynamic, adLockOptimistic, 1
  
  If Not rs.EOF And Not rs.BOF Then
     espere!ProgressBar1.Max = 100
     espere!ProgressBar1.Min = 1
     espere.Show
     espere.Refresh
  End If
  t = 0
  pb = 1
  While Not rs.EOF
     espere!ProgressBar1 = pb
     
     Select Case par
     Case Is = 0
       F = rs("fecha")
     Case Is = 1
       F = rs("fecha_vto")
     Case Is = 2
       F = rs("fecha_prob_entrega")
     End Select
     CTC = Format$(rs("ID_TIPOCOMP"), "000")
     tc = rs("descripcion")
     nc = rs("letra") & " " & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comprobante"), "00000000")
     d = Format$(rs("total"), "######0.00")
     cp = Format$(rs("a5.id_proveedor"), "0000")
     p = rs("proveedor05")
     t = t + Val(d)
     ni = rs("num_int")
     If rs("contado") = "S" Then
       tp = "CTDO"
     Else
       tp = "CTACTE"
     End If
     msf1.AddItem F & Chr(9) & cp & Chr(9) & p & Chr(9) & CTC & Chr(9) & tc & Chr(9) & nc & Chr(9) & d & Chr(9) & rs("estado") & Chr(9) & rs("num_int") & Chr(9) & rs("Moneda")

    rs.MoveNext
    pb = pb + 1
    If pb > 100 Then
      pb = 1
    End If
  Wend
  msf1.AddItem ""
  msf1.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "Totales:" & Chr(9) & Format$(t, "#####0.00") & Chr(9) & ""
 espere.Hide
   
End Sub
Sub carga2(par As Integer)
  Call armagrid
  q = "select * from a5, a6, g2, a1 where [id_tipocomp] = [id_tipo_comp] and a5.[id_proveedor] = a1.[id_proveedor] and a5.[num_int] = a6.[num_int] "
  c = " and "
  If c_prov.ListIndex > 0 Then
     q = q & c & " a5.[id_proveedor] = " & c_prov.ItemData(c_prov.ListIndex)
  End If
  

    q = q & c & " [id_tipocomp] = 65"
 
  
 Select Case par
  Case Is = 0 'fecha
   If IsDate(t_fecha) Then
     q = q & c & " datevalue(a5.[fecha]) >= datevalue('" & t_fecha & "')"
   End If
  
   If IsDate(t_fecha2) Then
     q = q & c & " datevalue(a5.[fecha]) <= datevalue('" & t_fecha2 & "')"
   End If
   o = " order by a5.[fecha]"
  
  Case Is = 1 'vencimiento
   If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha_vto]) >= datevalue('" & t_fecha & "')"
   End If
  
   If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha_vto]) <= datevalue('" & t_fecha2 & "')"
   End If
   o = " order by [fecha_vto]"
  
  Case Is = 2 'propbable entrega
    If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha_prob_entrega]) >= datevalue('" & t_fecha & "')"
    End If
  
    If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha_prob_entrega]) <= datevalue('" & t_fecha2 & "')"
    End If
    o = " order by [fecha_prob_entrega]"

  End Select
    
  If c_cuenta.ListIndex > 0 Then
    q = q & c & " [id_CUENTA] = " & c_cuenta.ItemData(c_cuenta.ListIndex)
  End If

   If c_pago.ListIndex > 0 Then
    q = q & c & " [estado] = '" & Mid$(c_pago, 2, 1) & "'"
   End If
  
  If t_producto <> "" Then
    q = q & c & "[detalle] like '%" & t_producto & "%'"
  End If
  
  If c_zona.ListIndex > 0 Then
    q = q & c & " [zona] = " & c_zona.ListIndex
  End If
  

  q = q & o
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  maximo = rs.RecordCount + 3
  espere!ProgressBar1.Max = maximo
  espere!ProgressBar1.Min = 1
  espere.Show
  espere.Refresh
  t = 0
  pb = 1
  While Not rs.EOF
     If pb >= maximo Then
       pb = 1
     End If
     espere!ProgressBar1 = pb
     Select Case par
     Case Is = 0
          F = rs("a5.fecha")
     Case Is = 1
          F = rs("fecha_vto")
     Case Is = 2
          F = rs("fecha_prob_entrega")
          
     End Select
     CTC = Format$(rs("ID_TIPOCOMP"), "000")
     tc = rs("descripcion")
     nc = rs("letra") & " " & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comprobante"), "00000000")
     d = Format$(rs("total"), "######0.00")
     cp = Format$(rs("a5.id_proveedor"), "0000")
     p = rs("proveedor05")
     
       t = t + Val(d)
     
     ni = rs("a5.num_int")
     msf1.AddItem F & Chr(9) & cp & Chr(9) & p & Chr(9) & CTC & Chr(9) & tc & Chr(9) & nc & Chr(9) & d & Chr(9) & rs("estado") & Chr(9) & rs("a5.num_int") & Chr(9) & rs("Moneda")
    rs.MoveNext
    pb = pb + 1
  Wend
  msf1.AddItem ""
  msf1.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "Totales:" & Chr(9) & Format$(t, "#####0.00") & Chr(9) & ""
  espere.Hide
End Sub

Private Sub btnacepta_Click()
If t_producto <> "" Then
     If Option1 = True Then
       Call carga2(0)
   Else
    If Option2 = True Then
      Call carga2(1)
    Else
      Call carga2(2)
    End If
    
   End If
Else
   
   If Option1 = True Then
       Call carga(0)
   Else
    If Option2 = True Then
      Call carga(1)
    Else
      Call carga(2)
    End If
    
   End If
End If
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 10
msf1.ColWidth(0) = 1300
msf1.ColWidth(1) = 700 'cod prov
msf1.ColWidth(2) = 3500
msf1.ColWidth(3) = 500
msf1.ColWidth(4) = 1700
msf1.ColWidth(5) = 1700
msf1.ColWidth(6) = 1200
msf1.ColWidth(7) = 1200
msf1.ColWidth(8) = 1000
msf1.ColWidth(9) = 800



msf1.TextMatrix(0, 0) = "Fecha"
msf1.TextMatrix(0, 1) = ""
msf1.TextMatrix(0, 2) = "Proveedor"
msf1.TextMatrix(0, 3) = ""
msf1.TextMatrix(0, 4) = "Operacion"
msf1.TextMatrix(0, 5) = "Nro.Comprobante"
msf1.TextMatrix(0, 6) = "Total"
msf1.TextMatrix(0, 7) = "Estado"
msf1.TextMatrix(0, 8) = "Num.Int."
msf1.TextMatrix(0, 9) = "Moneda"

For i = 0 To 8
  msf1.ColAlignment(i) = 1 'izq
Next i
msf1.ColAlignment(6) = 9 'der
msf1.ColAlignment(8) = 9 'der


End Sub








Private Sub C_cc_LostFocus()
If C_cc.ListIndex < 0 Then
   C_cc.ListIndex = 0
End If
End Sub

Private Sub c_cuenta_LostFocus()
If c_cuenta.ListIndex < 0 Then
  If Val(c_cuenta) > 0 Then
    c_cuenta.ListIndex = buscaindice(c_cuenta, Val(c_cuenta))
  Else
    c_cuenta.ListIndex = 0
  End If
End If
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

Private Sub Command1_Click()
gen_seleccionarimp.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub

Private Sub Form_Load()
Load espere
Call carga_proveedores(c_prov)
c_prov.AddItem "<Todos>", 0
c_prov.ListIndex = 0



Call carga_cuentas_cont(c_cuenta, "C", "D")
c_cuenta.AddItem "<Todas>", 0
c_cuenta.ListIndex = 0

c_pago.ListIndex = 0

Call armagrid
Call barraesag(Me)

Call carga_zonas(c_zona)
c_zona.AddItem "<Todas>", 0
c_zona.ListIndex = 0
Option1 = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload espere
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F1] Cambia Estado -[F7] Imprime - [ENTER] Detalla - [F8] Borra - [F6] Arch. Texto - [F11] Excel "
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF1 Then
   cp = Val(msf1.TextMatrix(msf1.RowSel, 8))
   If cp > 1 Then
    J = MsgBox("Confirmas cambiar estado O.C ", 4)
    If J = 6 Then
      Set rs = New ADODB.Recordset
      q = "select * from a5 where [num_int] = " & cp
      rs.Open q, cn1, adOpenDynamic, adLockOptimistic
      If Not rs.EOF And Not rs.BOF Then
        If rs("estado") = "P" Then
          rs("estado") = "F"
        Else
          rs("estado") = "P"
        End If
        rs.Update
        MsgBox ("Modificacion de estado realizada")
     End If
      Set rs = Nothing
    End If
    
   End If
End If

If KeyCode = vbKeyF7 Then
  Dim c(15) As Double
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    c(0) = 0
    c(1) = 2
    c(2) = 4
    c(3) = 5
    c(4) = 6
    c(5) = 8

    For i = 6 To 14
      c(i) = -1
    Next i
    
    Call imprimegrid(msf1, c(), "ORDENES DE COMPRA EMITIDAS", "", "", "", 72, 8, True, False)
  End If

End If


If KeyCode = vbKeyF6 Then
  Dim c2(15) As Double
    c2(0) = 0
    c2(1) = 2
    c2(2) = 4
    c2(3) = 5
    c2(4) = 6
    c2(5) = 8

    For i = 6 To 14
      c2(i) = -1
    Next i
    Call exportagrid(msf1, c2(), "ORDENES DE COMPRAS EMITIDAS", "", "", "", True, False, para.archivo_exportacion)

End If





 If KeyCode = vbKeyF8 Then
  Call nivel_acceso(2)
  If para.id_grupo_modulo_actual >= 8 Then
    indice = msf1.RowSel
      Set cl_comp = New COMPROBANTES
      cl_comp.cargar2 (Val(msf1.TextMatrix(indice, 8)))
      cl_comp.borrar
      Set cl_comp = Nothing
     
  End If
End If


If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If


End Sub

Sub borracomp()

End Sub
Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
      
    Load cc_detalle
    cc_detalle.T_IDPROV = msf1.TextMatrix(msf1.Row, 1)
    cc_detalle.t_prov = msf1.TextMatrix(msf1.Row, 2)
    cc_detalle.t_sucursal = Mid$(msf1.TextMatrix(msf1.Row, 5), 3, 4)
    cc_detalle.t_letra = Mid$(msf1.TextMatrix(msf1.Row, 5), 1, 1)
    cc_detalle.t_numcomp = Mid$(msf1.TextMatrix(msf1.Row, 5), 8, 8)
    cc_detalle.t_tipocomp = msf1.TextMatrix(msf1.Row, 3)
    cc_detalle.t_NUMINT = msf1.TextMatrix(msf1.Row, 8)
    cc_detalle.Show
  End If
End If

End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight

End Sub

Private Sub t_fecha_GotFocus()
t_fecha = ""
End Sub

Private Sub t_fecha2_GotFocus()
t_fecha2 = ""
End Sub

Private Sub t_producto_GotFocus()
t_producto = ""
End Sub
