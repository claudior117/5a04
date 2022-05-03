VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form con_subdiarioc 
   BackColor       =   &H00E0E0E0&
   Caption         =   "SUBDIARIO DE COMPRAS"
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
      Caption         =   "Tipo:"
      Height          =   615
      Left            =   3960
      TabIndex        =   20
      Top             =   7080
      Width           =   3255
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Detallado"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Resumido"
         Height          =   255
         Left            =   1680
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Agrupado por:"
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   7080
      Width           =   3255
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cuenta"
         Height          =   255
         Left            =   1680
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   3600
      TabIndex        =   11
      Top             =   480
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   115539969
      CurrentDate     =   38750
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1815
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   11655
      Begin VB.ComboBox c_zona 
         Height          =   315
         Left            =   8640
         TabIndex        =   18
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox c_cuenta2 
         Height          =   315
         Left            =   1680
         TabIndex        =   17
         Top             =   1320
         Width           =   4575
      End
      Begin VB.ComboBox c_cuenta1 
         Height          =   315
         Left            =   1680
         TabIndex        =   16
         Top             =   960
         Width           =   4575
      End
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Zona:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7080
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Cuenta Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Cuenta Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   3
      Top             =   7080
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "Con004.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "Con004.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   4
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
      TabIndex        =   2
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
            TextSave        =   "29/04/2022"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "04:27 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   4695
      Left            =   0
      TabIndex        =   12
      Top             =   2280
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   8281
      _Version        =   393216
   End
End
Attribute VB_Name = "con_subdiarioc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984


Sub carga()
   
  Call armagrid
  
  espere.Show
  espere.Label1 = "Espere procesando información..."
  espere.Refresh
  q = "select * from a5, g2, a1, g3, c_01 where [grabado] <> 'N' and  [id_tipocomp] = [id_tipo_comp] and a5.[id_proveedor] = a1.[id_proveedor] and a1.[cod_tipoiva] = g3.[cod_tipoiva] and a5.[id_cuenta] = c_01.[id_cuenta]"
  c = " and "
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  
  If c_cuenta1.ListIndex > 0 Then
    q = q & c & " [a5.id_cuenta] >= " & c_cuenta1.ItemData(c_cuenta1.ListIndex)
  End If
  
  If c_cuenta2.ListIndex > 0 Then
    q = q & c & " [a5.id_cuenta] <= " & c_cuenta2.ItemData(c_cuenta2.ListIndex)
  End If
  
  If c_zona.ListIndex > 0 Then
    q = q & c & " [zona] = " & c_zona.ListIndex
  End If
  
  If Option1 = True Then
   q = q & " order by [fecha], [a5.id_cuenta]"
  Else
   q = q & " order by [a5.id_cuenta], [fecha] "
  End If
  
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  tt = 0
  ti = 0
  ts = 0
  tng = 0
  trp = 0
  
  ttg = 0
  tig = 0
  tsg = 0
  tngg = 0
  trpg = 0
  
  pasada = 0
  
  While Not rs.EOF
    If pasada <> 0 Then
       If Option1 = True Then
         compara = rs("fecha")
         
       Else
         compara = rs("a5.id_cuenta")
       End If
       
       If compara <> corte Then
         
         'muestra totales por corte
          If Option3 = True Then
              'resumido
              If Option1 = True Then
                'por fecha
                 f4 = corte
                 c4 = ""
              Else
                 c4 = dc
                 f4 = ""
              End If
              msf1.AddItem f4 & Chr(9) & " " & Chr(9) & " " & Chr(9) & destipo & Chr(9) & tsg & Chr(9) & trpg & Chr(9) & tig & Chr(9) & tngg & Chr(9) & ttg & Chr(9) & " " & Chr(9) & c4 & Chr(9) & " "
              'msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & ""
          
          Else
              msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & "" & Chr(9) & "______________________" & Chr(9) & "______________________" & Chr(9) & "______________________" & Chr(9) & "______________________" & Chr(9) & "______________________"
              msf1.AddItem "" & Chr(9) & " " & Chr(9) & " " & Chr(9) & destipo & Chr(9) & tsg & Chr(9) & trpg & Chr(9) & tig & Chr(9) & tngg & Chr(9) & ttg & Chr(9) & " " & Chr(9) & " " & Chr(9) & " "
              msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & ""
          End If
         corte = compara
         ttg = 0
         tig = 0
         tsg = 0
         tngg = 0
         trpg = 0
 
               
       
       End If
       
     Else
       pasada = 1
       If Option1 = True Then
         corte = rs("fecha")
         desctipo = "Totales por Dia"
       Else
         corte = rs("a5.id_cuenta")
         desctipo = "Totales por Cuenta"
       End If
     End If
     F = Format$(rs("fecha"), "dd/mm/yy")
     tc = rs("g2.abreviatura")
     nc = rs("letra") & " " & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comprobante"), "00000000")
     idc = rs("a5.id_cuenta")
     dc = rs("c_01.descripcion")
     If rs("moneda") = "P" Then
       m5 = 1
     Else
       m5 = rs("cotiz_dolar")
     End If
     If rs("grabado") = "S" Then
       t = Format$(rs("total") * m5, "######0.00")
       i = Format$(rs("a5.iva") * m5, "######0.00")
       s = Format$(rs("subtotal") * m5, "######0.00")
       ng = Format$(rs("no_grabado") * m5, "######0.00")
       rp = Format$(rs("percep_ret") * m5, "######0.00")
     Else
       t = Format$(-rs("total") * m5, "######0.00")
       i = Format$(-rs("a5.iva") * m5, "######0.00")
       s = Format$(-rs("subtotal") * m5, "######0.00")
       ng = Format$(-rs("no_grabado") * m5, "######0.00")
       rp = Format$(-rs("percep_ret") * m5, "######0.00")
     End If
     tt = tt + Val(t)
     ti = ti + Val(i)
     ts = ts + Val(s)
     tng = tng + Val(ng)
     trp = trp + Val(rp)
     
     ttg = ttg + Val(t)
     tig = tig + Val(i)
     tsg = tsg + Val(s)
     tngg = tngg + Val(ng)
     trpg = trpg + Val(rp)
     
     If Option4 = True Then
       msf1.AddItem F & Chr(9) & rs("proveedor05") & Chr(9) & rs("cuit") & " " & rs("g3.abreviatura") & Chr(9) & tc & " " & nc & Chr(9) & s & Chr(9) & rp & Chr(9) & i & Chr(9) & ng & Chr(9) & t & Chr(9) & Format$(rs("a5.id_cuenta"), "000000") & Chr(9) & rs("c_01.descripcion") & Chr(9) & Format$(rs("num_int"), "00000")
     End If
    rs.MoveNext
  Wend
  
 If Option3 = True Then
              'resumido
              If Option1 = True Then
                'por fecha
                 f4 = corte
                 c4 = ""
              Else
                 c4 = dc
                 f4 = ""
              End If
              msf1.AddItem f4 & Chr(9) & " " & Chr(9) & " " & Chr(9) & destipo & Chr(9) & tsg & Chr(9) & trpg & Chr(9) & tig & Chr(9) & tngg & Chr(9) & ttg & Chr(9) & " " & Chr(9) & c4 & Chr(9) & " "
              'msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & ""
          
 Else
              msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & "" & Chr(9) & "______________________" & Chr(9) & "______________________" & Chr(9) & "______________________" & Chr(9) & "______________________" & Chr(9) & "______________________"
              msf1.AddItem "" & Chr(9) & " " & Chr(9) & " " & Chr(9) & destipo & Chr(9) & tsg & Chr(9) & trpg & Chr(9) & tig & Chr(9) & tngg & Chr(9) & ttg & Chr(9) & " " & Chr(9) & " " & Chr(9) & " "
              msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & ""
End If
 msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & ""
 msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & "" & Chr(9) & "______________________" & Chr(9) & "______________________" & Chr(9) & "______________________" & Chr(9) & "______________________" & Chr(9) & "______________________"
 msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & "Totales:" & Chr(9) & Format$(ts, "######0.00") & Chr(9) & Format$(trp, "######0.00") & Chr(9) & Format$(ti, "######0.00") & Chr(9) & Format$(tng, "######0.00") & Chr(9) & Format$(tt, "######0.00")
   
 Unload espere
End Sub

Private Sub btnacepta_Click()
Call carga

End Sub

Private Sub btnsale_Click()
Unload Me
End Sub










Private Sub c_cuenta1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  t_cuenta = c_cuenta1.ItemData(c_cuenta1.ListIndex)
End If
End Sub

Private Sub c_cuenta1_LostFocus()
If c_cuenta1.ListIndex < 0 Then
  If Val(c_cuenta1) > 0 Then
    c_cuenta1.ListIndex = buscaindice(c_cuenta1, Val(c_cuenta1))
  Else
    c_cuenta1.ListIndex = 0
    t_cuenta1 = 0
  End If
End If

End Sub

Private Sub c_cuenta2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  t_cuenta2 = c_cuenta2.ItemData(c_cuenta2.ListIndex)
End If
End Sub

Private Sub c_cuenta2_LostFocus()
If c_cuenta2.ListIndex < 0 Then
  If Val(c_cuenta2) > 0 Then
    c_cuenta2.ListIndex = buscaindice(c_cuenta2, Val(c_cuenta2))
  Else
    c_cuenta2.ListIndex = 0
    t_cuenta2 = 0
  End If
End If

End Sub

Private Sub cal1_DblClick()
If cal1.Tag = "1" Then
  t_fecha = cal1.Value
Else
  t_fecha2 = cal1.Value
End If
cal1.Visible = False
End Sub

Private Sub cal1_LostFocus()
cal1.Visible = False
End Sub

Private Sub Form_Activate()
cal1.Visible = False
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 12
msf1.ColWidth(0) = 900
msf1.ColWidth(1) = 2000
msf1.ColWidth(2) = 1500
msf1.ColWidth(3) = 2000
msf1.ColWidth(4) = 1000
msf1.ColWidth(5) = 1000
msf1.ColWidth(6) = 1000
msf1.ColWidth(7) = 1000
msf1.ColWidth(8) = 1000
msf1.ColWidth(9) = 900
msf1.ColWidth(10) = 1500
msf1.ColWidth(11) = 700


msf1.TextMatrix(0, 0) = "Fecha"
msf1.TextMatrix(0, 1) = "Proveedor"
msf1.TextMatrix(0, 2) = "Cuit "
msf1.TextMatrix(0, 3) = "Tipo y Nro.Comprob."
msf1.TextMatrix(0, 4) = "Subtotal  "
msf1.TextMatrix(0, 5) = "Ret./Perc."
msf1.TextMatrix(0, 6) = "Iva"
msf1.TextMatrix(0, 7) = "No Grab."
msf1.TextMatrix(0, 8) = "Total"
msf1.TextMatrix(0, 9) = "Cuenta"
msf1.TextMatrix(0, 10) = "Detalle Cuenta"
msf1.TextMatrix(0, 11) = "Num.Int."

For i = 0 To 3
  msf1.ColAlignment(i) = 1 'izq
Next i
For i = 4 To 8
  msf1.ColAlignment(i) = 9 'der
Next i

End Sub


Private Sub Form_Load()

Call barraesag(Me)
Me.Option1 = True
Call carga_cuentas_cont(c_cuenta1, "C", "C")
Call carga_cuentas_cont(c_cuenta2, "C", "C")
c_cuenta1.AddItem "<Sin Especificar>", 0
c_cuenta2.AddItem "<Sin Especificar>", 0
c_cuenta1.ListIndex = 0
c_cuenta2.ListIndex = 0
Call armagrid

Call carga_zonas(c_zona)
c_zona.AddItem "<Todas>", 0
c_zona.ListIndex = 0
Option4 = True
End Sub




Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F7] Imprime - [F11] Excel"

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
  Dim c(15) As Double
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    c(0) = 0
    c(1) = 1
    c(2) = 2
    c(3) = 3
    c(4) = 4
    c(5) = 5
    c(6) = 6
    c(7) = 7
    c(8) = 8
    c(9) = 9
    c(10) = 10
    For i = 11 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "SUBDIARIO DE COMPRAS", "", "Periodo: " & t_fecha & " : " & t_fecha2, "", 95, 6, True, False)
  End If

End If


If KeyCode = vbKeyF6 Then
  Dim c2(15) As Double
    c2(0) = 0
    c2(1) = 1
    c2(2) = 2
    c2(3) = 3
    c2(4) = 4
    c2(5) = 5
    c2(6) = 6
    c2(7) = 7
    c2(8) = 8
    c2(9) = 9
    c2(10) = 10
    For i = 11 To 14
      c2(i) = -1
    Next i
    Call exportagrid(msf1, c2(), "SUBDIARIO DE COMPRAS", "", "Periodo: " & t_fecha & " : " & t_fecha2, "", True, False, para.archivo_exportacion)

End If

If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If

End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    Load cc_detalle
    cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 11)
    cc_detalle.Show
  End If
End If

End Sub

Private Sub t_fecha_DblClick()
cal1.Visible = True
cal1.Tag = "1"
End Sub

Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
    t_fecha = Format$(Now, "dd/mm/yyyy")
  End If
End If
End Sub

Private Sub t_fecha2_DblClick()
cal1.Visible = True
cal1.Tag = "2"
End Sub

Private Sub t_fecha2_LostFocus()
If t_fecha2 <> "" Then
  If Not IsDate(t_fecha2) Then
    t_fecha2 = Format$(Now, "dd/mm/yyyy")
  End If
End If
End Sub
