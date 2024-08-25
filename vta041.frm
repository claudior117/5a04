VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form vta_retyperc_realizadas 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INFORME DE PERCEPCIONES REALIZADAS(No Usar)"
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18315
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   18315
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1815
      Left            =   7080
      TabIndex        =   11
      Top             =   0
      Width           =   8175
      Begin VB.ComboBox c_imp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   17
         Top             =   1200
         Width           =   5535
      End
      Begin VB.ComboBox c_cli 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   15
         Top             =   720
         Width           =   5535
      End
      Begin VB.ComboBox c_comp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "vta041.frx":0000
         Left            =   2160
         List            =   "vta041.frx":0002
         TabIndex        =   13
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "Impuesto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Comprobantes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   4680
      TabIndex        =   9
      Top             =   120
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   179306497
      CurrentDate     =   38750
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   4335
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   16440
      TabIndex        =   3
      Top             =   8280
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta041.frx":0004
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
         Picture         =   "vta041.frx":0886
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
      Top             =   9330
      Width           =   18315
      _ExtentX        =   32306
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   14111
            MinWidth        =   14111
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   12347
            MinWidth        =   12347
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "25/08/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:15 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   6135
      Left            =   0
      TabIndex        =   10
      Top             =   1920
      Width           =   18135
      _ExtentX        =   31988
      _ExtentY        =   10821
      _Version        =   393216
   End
End
Attribute VB_Name = "vta_retyperc_realizadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984


Sub carga()
Call armagrid
Call buscaperc
  
  
  
   
   
   
End Sub
Sub buscaperc()
  'seleccionar todos los comprobantes comprobantes
'percepciones ingresos brutos
If (c_comp.ListIndex = 0 Or c_comp.ListIndex = 1) And (c_imp.ListIndex = 0 Or c_imp.ListIndex = 2) Then
  q = "select * from VTA_02, vta_06, VTA_01 where  vta_02.[id_tipocomp] = vta_06.[id_tipocomp] and VTA_02.[id_CLIENTE] = VTA_01.[id_CLIENTE] and vta_02.[sucursal] = vta_06.[sucursal] "
  c = " and "
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  q = q & " and perc_ib <> 0"
  q = q & " order by [fecha]"
  
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  tr = 0
  ti = 0
  ts = 0
  p = 0
  ta = 0
  While Not rs.EOF
     
     If p = 0 Then
       p = 1
       msf1.AddItem " " & Chr(9) & "Percepcion IB"
     End If
     F = Format$(rs("fecha"), "dd/mm/yy")
     tc = rs("abreviatura")
     nc = rs("letra") & " " & Format$(rs("vta_02.sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
     If rs("vta_02.moneda") = "P" Then
       c5 = 1
     Else
       c5 = rs("cotizacion_dolar")
     End If
     c = rs("cuit")
     If rs("vta_06.iva") = "S" Then
      i = Format$(rs("perc_ib") * c5, "######0.00")
      s = Format$(rs("subtotal") * c5, "######0.00")
     Else
      i = Format$(-rs("perc_ib") * c5, "######0.00")
      s = Format$(-rs("subtotal") * c5, "######0.00")
     End If
     ts = ts + Val(s)
     ti = ti + Val(i)
     sa = sa + Val(s)
     ia = ia + Val(i)
     msf1.AddItem F & Chr(9) & "" & Chr$(9) & rs("denominacion") & Chr$(9) & c & Chr(9) & tc & " " & nc & Chr(9) & Format$(Val(s), para.formato_numerico) & Chr(9) & Format$(Val(i), para.formato_numerico) & Chr(9) & Format$(rs("num_int"), "00000")
    rs.MoveNext
  Wend
  If ti > 0 Then
     msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & "" & Chr(9) & "______________________" & Chr(9) & "______________________"
     msf1.AddItem "" & Chr(9) & "" & Chr(9) & "Total Percepciones IB " & Chr(9) & " " & Chr$(9) & Chr(9) & Format$(sa, para.formato_numerico) & Chr$(9) & Format$(ia, para.formato_numerico)
     msf1.AddItem ""
  End If
  Set rs = Nothing
End If


If (c_comp.ListIndex = 0 Or c_comp.ListIndex = 1) And (c_imp.ListIndex = 0 Or c_imp.ListIndex = 1) Then
  q = "select * from VTA_02, vta_06, VTA_01 where  vta_02.[id_tipocomp] = vta_06.[id_tipocomp] and VTA_02.[id_CLIENTE] = VTA_01.[id_CLIENTE] and vta_02.[sucursal] = vta_06.[sucursal]"
  c = " and "
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  q = q & " and perc_iva <> 0"
  q = q & " order by [fecha]"
  
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  tr = 0
  ti = 0
  ts = 0
  p = 0
  ta = 0
  While Not rs.EOF
     If p = 0 Then
       p = 1
       msf1.AddItem " " & Chr(9) & "Percepcion IVA"
     End If
     F = Format$(rs("fecha"), "dd/mm/yy")
     tc = rs("abreviatura")
     nc = rs("letra") & " " & Format$(rs("vta_02.sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
     If rs("vta_02.moneda") = "P" Then
       c5 = 1
     Else
       c5 = rs("cotizacion_dolar")
     End If
     If rs("vta_06.iva") = "S" Then
      i = Format$(rs("perc_iva") * c5, "######0.00")
      s = Format$(rs("subtotal") * c5, "######0.00")
     Else
      i = Format$(-rs("perc_iva") * c5, "######0.00")
      s = Format$(-rs("subtotal") * c5, "######0.00")
     End If
     
     
     c = rs("cuit")
     ts = ts + Val(s)
     ti = ti + Val(i)
     sa = sa + Val(s)
     ia = ia + Val(i)
     msf1.AddItem F & Chr(9) & "" & Chr$(9) & rs("denominacion") & Chr$(9) & c & Chr(9) & tc & " " & nc & Chr(9) & Format$(Val(s), para.formato_numerico) & Chr(9) & Format$(Val(i), para.formato_numerico) & Chr(9) & Format$(rs("num_int"), "00000")
    rs.MoveNext
  Wend
  If ti > 0 Then
     msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & "" & Chr(9) & "______________________" & Chr(9) & "______________________"
     msf1.AddItem "" & Chr(9) & "" & Chr(9) & "Total Percepciones IVA " & " " & Chr(9) & Chr$(9) & Chr(9) & Format$(sa, para.formato_numerico) & Chr$(9) & Format$(ia, para.formato_numerico)
     msf1.AddItem ""
  End If
  Set rs = Nothing
End If


If (c_comp.ListIndex = 0 Or c_comp.ListIndex = 1) And (c_imp.ListIndex = 0 Or c_imp.ListIndex = 3) Then
  q = "select * from VTA_02, vta_06, VTA_01 where  vta_02.[id_tipocomp] = vta_06.[id_tipocomp] and VTA_02.[id_CLIENTE] = VTA_01.[id_CLIENTE] and vta_02.[sucursal] = vta_06.[sucursal]"
  c = " and "
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  q = q & " and perc_gan <> 0"
  q = q & " order by [fecha]"
  
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  tr = 0
  ti = 0
  ts = 0
  p = 0
  ta = 0
  While Not rs.EOF
     If p = 0 Then
       p = 1
       msf1.AddItem " " & Chr(9) & "Percepcion GAN"
     End If
     F = Format$(rs("fecha"), "dd/mm/yy")
     tc = rs("abreviatura")
     nc = rs("letra") & " " & Format$(rs("vta_02.sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
     If rs("vta_02.moneda") = "P" Then
       c5 = 1
     Else
       c5 = rs("cotizacion_dolar")
     End If
     
       If rs("vta_06.iva") = "S" Then
      i = Format$(rs("perc_gan") * c5, "######0.00")
      s = Format$(rs("subtotal") * c5, "######0.00")
     Else
      i = Format$(-rs("perc_gan") * c5, "######0.00")
      s = Format$(-rs("subtotal") * c5, "######0.00")
     End If
   
     c = rs("cuit")
     ts = ts + Val(s)
     ti = ti + Val(i)
     sa = sa + Val(s)
     ia = ia + Val(i)
     msf1.AddItem F & Chr(9) & "" & Chr$(9) & rs("denominacion") & Chr(9) & c & Chr$(9) & tc & " " & nc & Chr(9) & Format$(Val(s), para.formato_numerico) & Chr(9) & Format$(Val(i), para.formato_numerico) & Chr(9) & Format$(rs("num_int"), "00000")
    rs.MoveNext
  Wend
  If ti > 0 Then
     msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & "" & Chr(9) & "______________________" & Chr(9) & "______________________"
     msf1.AddItem "" & Chr(9) & "" & Chr(9) & "Total Percepciones IB " & Chr(9) & Chr$(9) & Chr(9) & Format$(sa, para.formato_numerico) & Chr$(9) & Format$(ia, para.formato_numerico)
     msf1.AddItem ""
  End If
  Set rs = Nothing
End If


End Sub
Private Sub btnacepta_Click()
espere.Show
espere.Refresh
Call carga
Unload espere

End Sub
Sub cargaib()
     End Sub
Private Sub btnsale_Click()
Unload Me
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



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 9
msf1.ColWidth(0) = 1300
msf1.ColWidth(1) = 2100
msf1.ColWidth(2) = 3300
msf1.ColWidth(3) = 1900
msf1.ColWidth(4) = 3300
msf1.ColWidth(5) = 2100
msf1.ColWidth(6) = 2100
msf1.ColWidth(7) = 2100
msf1.ColWidth(8) = 500



msf1.TextMatrix(0, 0) = "Fecha"
msf1.TextMatrix(0, 1) = "Tipo Impuesto"
msf1.TextMatrix(0, 2) = "Cliente"
msf1.TextMatrix(0, 3) = "Cuit"
msf1.TextMatrix(0, 4) = "Tipo y Nro.Comprob."
msf1.TextMatrix(0, 5) = "Imponible"
msf1.TextMatrix(0, 6) = "Impuesto"
msf1.TextMatrix(0, 7) = "Num.Int."
msf1.TextMatrix(0, 8) = " "

For i = 0 To 4
  msf1.ColAlignment(i) = 1 'izq
Next i
For i = 5 To 7
  msf1.ColAlignment(i) = 9 'der
Next i

End Sub

Private Sub Form_Load()
Call barraesag(Me)
cal1.Visible = False
Call armagrid
Call carga_clientes(c_cli)
c_cli.AddItem "<Todos>", 0
c_cli.ListIndex = 0
Call cargaret

End Sub
Sub cargaret()
'impuestos
c_imp.clear
c_imp.AddItem "<Todos>", 0
c_imp.AddItem "Iva", 1
c_imp.AddItem "Ing.Brutos", 2
c_imp.AddItem "Ganancias", 3
c_imp.AddItem "Seg. Social", 4
c_imp.ListIndex = 0

c_comp.clear
c_comp.AddItem "<Todos>", 0
c_comp.AddItem "Percepciones", 1

c_comp.ListIndex = 0
End Sub


Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F7] Imprime - "

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
  Dim c(15) As Double
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    c(0) = 8
    c(1) = 0
    c(2) = 1
    c(3) = 2
    c(4) = 3
    c(5) = 4
    c(6) = 5
    c(7) = 6
    c(8) = 7
    For i = 9 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "LISTADO DE RETENCIONES y PERCEPCIONES RECIBIDAS por VENTAS", "", "Periodo: " & t_fecha & " : " & t_fecha2, "", 95, 6, True, False)
  End If

End If

End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    Load cc_detalle
    vta_cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 7)
    vta_cc_detalle.Show
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
