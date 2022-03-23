VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_verremitos 
   BackColor       =   &H00E0E0E0&
   Caption         =   "INFORME DE REMITOS y NOTAS DEVOLUCION"
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
      Caption         =   "Moneda"
      Height          =   615
      Left            =   3720
      TabIndex        =   29
      Top             =   7320
      Width           =   3015
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "$"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "U$s"
         Height          =   255
         Left            =   1560
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ordenados por:"
      Height          =   615
      Left            =   240
      TabIndex        =   26
      Top             =   7320
      Width           =   3015
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cliente"
         Height          =   255
         Left            =   1560
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.ComboBox c_vend 
      Height          =   315
      Left            =   8760
      TabIndex        =   23
      Top             =   1560
      Width           =   2895
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   4815
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8493
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   2295
      Left            =   240
      TabIndex        =   9
      Top             =   0
      Width           =   11535
      Begin VB.TextBox t_obs 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   32
         ToolTipText     =   "Busca Remitos que en la observacion tenga el texto ingresado"
         Top             =   1920
         Width           =   2895
      End
      Begin VB.ComboBox c_transp 
         Height          =   315
         Left            =   1680
         TabIndex        =   20
         Text            =   "c_transp"
         Top             =   1560
         Width           =   4815
      End
      Begin VB.TextBox T_prod 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   18
         Top             =   1200
         Width           =   2895
      End
      Begin VB.ComboBox c_prod 
         Height          =   315
         Left            =   1680
         TabIndex        =   16
         Text            =   "c_prod"
         Top             =   1080
         Width           =   4815
      End
      Begin VB.ComboBox c_estado 
         Height          =   315
         ItemData        =   "vta017.frx":0000
         Left            =   8520
         List            =   "vta017.frx":000D
         TabIndex        =   14
         Text            =   "c_estado"
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox c_tipocomp 
         Height          =   315
         ItemData        =   "vta017.frx":0034
         Left            =   8520
         List            =   "vta017.frx":0041
         TabIndex        =   1
         Text            =   "c_tipocomp"
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5160
         MaxLength       =   10
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   2
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
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Observaciones:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6960
         TabIndex        =   33
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Vendedor:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6960
         TabIndex        =   25
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   9960
         TabIndex        =   22
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Transporte:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Desc. Prod.:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6960
         TabIndex        =   19
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Producto:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Estado:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   6960
         TabIndex        =   15
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Tipo Comprobante:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   6960
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Cliente:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   6
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta017.frx":006B
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "vta017.frx":08ED
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
      TabIndex        =   5
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
            TextSave        =   "11/03/2022"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "08:51 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800080&
      Caption         =   "Transporte:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   24
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "vta_verremitos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer


Sub carga()
  If Option3 = True Then
    mp = "D"
  Else
    mp = "P"
  End If
  
  
  Call armagrid
  q = "select * from vta_02, vta_06, vta_01 where vta_02.[id_tipocomp] = vta_06.[id_tipocomp] and vta_02.[id_cliente] = vta_01.[id_cliente] and vta_02.[sucursal] = vta_06.[sucursal] "
  c = " and "
  If c_prov.ListIndex > 0 Then
     q = q & c & " vta_02.[id_cliente] = " & c_prov.ItemData(c_prov.ListIndex)
  End If
  
  If c_tipocomp.ListIndex > 0 Then
    q = q & c & " vta_02.[id_tipocomp] = " & Val(Mid$(c_tipocomp, 1, 2))
  Else
    q = q & c & " vta_02.[id_tipocomp] >= 40 and  vta_02.[id_tipocomp] < 50 "
  End If
  
  If c_estado.ListIndex > 0 Then
    q = q & c & " [estado] = '" & Mid$(c_estado, 1, 1) & "'"
  End If
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  
  If c_transp.ListIndex > 0 Then
     q = q & c & " [Id_transporte] = " & c_transp.ItemData(c_transp.ListIndex)
  End If
  
  
  If c_vend.ListIndex > 0 Then
     q = q & c & " vta_02.[Id_vendedor] = " & c_vend.ItemData(c_vend.ListIndex)
  End If
  
  If t_obs <> "" Then
    q = q & c & " vta_02.[observaciones] like '%" & t_obs & "%'"
  End If
  
  If Option1 = True Then
    q = q & " order by [fecha], [num_comp]"
  Else
    q = q & " order by [denominacion], [fecha], [num_comp]"
  End If
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  t = 0
  reg = 0
  tfa = 0
  tpe = 0
  While Not rs.EOF
     F = rs("fecha")
     CTC = Format$(rs("vta_02.ID_TIPOCOMP"), "000")
     tc = rs("abreviatura")
     nc = rs("letra") & " " & Format$(rs("vta_02.sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
     d = Format$(rs("total"), "######0.00")
     cp = Format$(rs("vta_02.id_cliente"), "0000")
     p = rs("denominacion")
     m = rs("vta_02.moneda")
     
     ni = rs("num_int")
     
     If rs("estado") = "S" Then
     'calculo pendientes
     Set rs1 = New ADODB.Recordset
     q = "select * from vta_03 where [num_int] = " & rs("num_int")
     rs1.Open q, cn1
     fa = 0
     pe = 0
     While Not rs1.EOF
       pe = pe + (rs1("cantidad") * rs1("pu_final"))
       rs1.MoveNext
     Wend
     Set rs1 = Nothing
     fa = Val(d) - pe
    
  Else
    fa = Val(d)
    pe = 0
  End If
  If m <> mp Then
    If m = "D" Then
       fa = fa * rs("cotizacion_dolar")
       pe = pe * rs("cotizacion_dolar")
        d = Format$(Val(d) * rs("cotizacion_dolar"), "#####0.00")
    Else
       fa = fa / rs("cotizacion_dolar")
       pe = pe / rs("cotizacion_dolar")
       d = Format$(Val(d) / rs("cotizacion_dolar"), "#####0.00")
    End If
  End If
  If rs("vta_02.id_tipocomp") = 45 Then
       t = t + Val(d)
  Else
       t = t - Val(d)
  End If
  msf1.AddItem F & Chr(9) & cp & Chr(9) & p & Chr(9) & CTC & Chr(9) & tc & Chr(9) & nc & Chr(9) & d & Chr(9) & Format$(fa, "#####0.00") & Chr$(9) & Format$(pe, "#####0.00") & Chr(9) & rs("estado") & Chr(9) & rs("num_int") & Chr(9) & m & Chr(9) & rs("vta_02.observaciones")
  reg = reg + 1
  Label9 = reg
  Label9.Refresh
  tfa = tfa + fa
  tpe = tpe + pe
  rs.MoveNext
 Wend

  msf1.AddItem ""
  msf1.AddItem "" & Chr(9) & "" & Chr(9) & "Comprobantes: " & reg & Chr(9) & "" & Chr(9) & "" & Chr(9) & "Totales:" & Chr(9) & Format$(t, "#####0.00") & Chr(9) & Format$(tfa, "#####0.00") & Chr(9) & Format$(tpe, "#####0.00") & Chr(9) & ""

  
 
End Sub

Private Sub btnacepta_Click()
Call carga
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 13
msf1.ColWidth(0) = 1200
msf1.ColWidth(1) = 700 'cod prov
msf1.ColWidth(2) = 3500
msf1.ColWidth(3) = 500
msf1.ColWidth(4) = 700
msf1.ColWidth(5) = 1700
msf1.ColWidth(6) = 1100
msf1.ColWidth(7) = 1100
msf1.ColWidth(8) = 1100
msf1.ColWidth(9) = 500
msf1.ColWidth(10) = 1000
msf1.ColWidth(11) = 500
msf1.ColWidth(12) = 1500


msf1.TextMatrix(0, 0) = "Fecha"
msf1.TextMatrix(0, 1) = ""
msf1.TextMatrix(0, 2) = "Cliente"
msf1.TextMatrix(0, 3) = ""
msf1.TextMatrix(0, 4) = "Op."
msf1.TextMatrix(0, 5) = "Nro.Comprobante"
msf1.TextMatrix(0, 6) = "Imp.Final"
msf1.TextMatrix(0, 7) = "Fact."
msf1.TextMatrix(0, 8) = "Pend."
msf1.TextMatrix(0, 9) = "Estado"
msf1.TextMatrix(0, 10) = "Num.Int."
msf1.TextMatrix(0, 11) = "Mda"
msf1.TextMatrix(0, 12) = "Observaciones "
For i = 0 To 8
    msf1.ColAlignment(i) = 1 'izq
Next i
msf1.ColAlignment(6) = 9 'der
msf1.ColAlignment(7) = 9 'der
msf1.ColAlignment(8) = 9 'der

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

Private Sub c_vend_LostFocus()
If c_vend.ListIndex < 0 Then
  c_vend.ListIndex = 0
End If
End Sub

Private Sub cal1_DblClick()
If cal1.Tag = "1" Then
   t_fecha = cal1
Else
   t_fecha2 = cal1
End If
cal1.Visible = False

End Sub

Private Sub cal1_LostFocus()
If cal1.Tag = "1" Then
   t_fecha = cal1
Else
   t_fecha2 = cal1
End If
cal1.Visible = False

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
    Call TabEnter2(Me, 4)
  'Case Is = 27
  '      Me.Hide
End Select

End Sub

Private Sub Form_Load()
Call carga_clientes(c_prov)
c_prov.AddItem "<Todos>", 0
c_prov.ListIndex = 0

'Call carga_productos(c_prod)
c_prod.AddItem "<Todos>", 0
c_prod.ListIndex = 0

Call carga_vendedores(c_vend)
c_vend.AddItem "<Todos>", 0
c_vend.ListIndex = 0

T_prod = ""

c_estado.ListIndex = 0

c_tipocomp.ListIndex = 0

Call carga_transporte(c_transp)
c_transp.AddItem "<Todos>", 0
c_transp.ListIndex = 0

Option1 = True

If para.moneda = "P" Then
  Option4 = True
Else
  Option3 = True
End If

Call armagrid
Call barraesag(Me)


End Sub


Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F7] Imprime - [ENTER] Detalla - [F11] Excel"
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
  Dim c(15) As Double
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    c(0) = 0
    c(1) = 2
    c(2) = 4
    c(3) = 5
    c(4) = 6
    c(5) = 7
    c(6) = 8
    c(7) = 9
    c(8) = 11
    c(9) = 12
    For i = 10 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "REMITOS EMITIDOS", "Cliente:" & c_prov & "           Estado: " & c_estado, "Fecha desde: " & t_fecha & "  Fecha hasta: " & t_fecha2, "Vendedor: " & c_vend, 45, 8, True, False, "H")
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
    vta_cc_detalle.t_idprov = msf1.TextMatrix(msf1.Row, 1)
    vta_cc_detalle.t_prov = msf1.TextMatrix(msf1.Row, 2)
    vta_cc_detalle.t_sucursal = Mid$(msf1.TextMatrix(msf1.Row, 5), 3, 4)
    vta_cc_detalle.t_letra = Mid$(msf1.TextMatrix(msf1.Row, 5), 1, 1)
    vta_cc_detalle.t_numcomp = Mid$(msf1.TextMatrix(msf1.Row, 5), 8, 8)
    vta_cc_detalle.t_tipocomp = msf1.TextMatrix(msf1.Row, 3)
    vta_cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 10)
    vta_cc_detalle.Show
  End If
End If

End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight

End Sub

Private Sub t_fecha_DblClick()
cal1.Visible = True
cal1.Tag = "1"
End Sub

Private Sub t_fecha_GotFocus()
t_fecha = ""
End Sub

Private Sub t_fecha2_DblClick()
cal1.Visible = True
cal1.Tag = "2"

End Sub

Private Sub t_fecha2_GotFocus()
t_fecha2 = ""
End Sub

Private Sub t_obs_GotFocus()
t_obs = ""
End Sub

Private Sub t_prod_GotFocus()
T_prod = ""
End Sub
