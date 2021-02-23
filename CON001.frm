VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form con_estadocuenta 
   BackColor       =   &H00E0E0E0&
   Caption         =   "ESTADO DE CUENTA POR PROVEDOR"
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
      Height          =   495
      Left            =   8640
      TabIndex        =   19
      Top             =   1440
      Width           =   2535
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Muestra Saldo en U$s"
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   8640
      TabIndex        =   14
      Top             =   840
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
      Height          =   855
      Left            =   8640
      TabIndex        =   11
      Top             =   0
      Width           =   3135
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha vencimiento"
         Height          =   495
         Left            =   1680
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha Comprobante"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   3600
      TabIndex        =   10
      Top             =   120
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   111738881
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
      Height          =   1575
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   8295
      Begin VB.ComboBox c_zona 
         Height          =   315
         ItemData        =   "CON001.frx":0000
         Left            =   2160
         List            =   "CON001.frx":0002
         TabIndex        =   22
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CommandButton Command5 
         Height          =   255
         Left            =   7080
         Picture         =   "CON001.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   17
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   2160
         TabIndex        =   0
         Text            =   "c_prov"
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
         Caption         =   "Zona:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   18
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Provedor:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1935
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
         Picture         =   "CON001.frx":0376
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
         Picture         =   "CON001.frx":0BF8
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
            TextSave        =   "23/02/2021"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:28 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "con_estadocuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim saldoanterior As Double
Sub carga()
  
  Call armagrid
  sa = 0
  da = 0
  ha = 0
  
  sao = 0
  dao = 0
  hao = 0
  T2 = 0
  If t_fecha <> "" Then
     q = "select * from a5 where [id_proveedor] = " & c_prov.ItemData(c_prov.ListIndex) & " and [ctacte] <> 'N' " & " and [contado] = " & "'N'"
     If Option1 = True Then
        q = q & " and datevalue([fecha]) < datevalue('" & t_fecha & "')"
     Else
        q = q & " and datevalue([fecha_vto]) < datevalue('" & t_fecha & "')"
     End If
     
     If c_zona.ListIndex > 0 Then
       q = q & " and [zona] = " & c_zona.ListIndex
     End If
     
    Set rs = New adodb.Recordset
    rs.Open q, cn1
    While Not rs.EOF
     If Option4 = True Then
      If rs("moneda") = "P" Then
        t = rs("total")
        T2 = rs("total_d")
      Else
        t = rs("total_d")
        T2 = rs("total")
      End If
     Else
      If rs("moneda") = "D" Then
        t = rs("total")
        T2 = rs("total_d")
      Else
        t = rs("total_d")
        T2 = rs("total")
      End If
     End If
     
     If rs("ctacte") = "D" Then
        da = da + t
        dao = dao + T2
     Else
        ha = ha + t
        hao = hao + T2
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
  q = "select * from a5, g2 where [id_proveedor] = " & c_prov.ItemData(c_prov.ListIndex) & " and a5.[ctacte] <> 'N' and [contado] = 'N' and [id_tipocomp] = [id_tipo_comp]"
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
    
  If c_zona.ListIndex > 0 Then
       q = q & " and [zona] = " & c_zona.ListIndex
  End If
    
  If Option1 = True Then
     q = q & " order by [fecha], [id_tipocomp], [num_comprobaNTE]"
  Else
      q = q & " order by [fecha_vto], [id_tipocomp], [num_comprobante]"
  End If
    
  Set rs = New adodb.Recordset
  rs.Open q, cn1
  s = sa
  hao = ""
  dao = ""
  While Not rs.EOF
    If Option1 = True Then
         F = rs("fecha")
    Else
         F = rs("fecha_vto")
    End If
     CTC = Format$(rs("ID_TIPOCOMP"), "000")
     tc = rs("descripcion")
     nc = rs("letra") & " " & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comprobante"), "00000000")
     If Option4 = True Then
      If rs("moneda") = "P" Then
        t = rs("total")
        T2 = rs("total_d")
      Else
        t = rs("total_d")
        T2 = rs("total")
      End If
     Else
       If rs("moneda") = "D" Then
        t = rs("total")
        T2 = rs("total_d")
      Else
        t = rs("total_d")
        T2 = rs("total")
      End If
     End If
     If rs("a5.ctacte") = "D" Then
       d = Format$(t, "######0.00")
       h = ""
       dao = Format$(T2, "######0.00")
       hao = ""
     Else
       h = Format$(t, "######0.00")
       d = ""
       hao = Format$(T2, "######0.00")
       dao = ""
     
     End If
     s = Format$(Val(s) + Val(d) - Val(h), "######0.00")
     sao = Format$(Val(sao) + Val(dao) - Val(hao), "######0.00")
     ni = rs("num_int")
     o = rs("obs")
     If Check1 = 0 Then
       msf1.AddItem F & Chr(9) & CTC & Chr(9) & tc & Chr(9) & nc & Chr(9) & d & Chr(9) & h & Chr(9) & s & Chr(9) & o & Chr(9) & ni
     Else
        msf1.AddItem F & Chr(9) & CTC & Chr(9) & tc & Chr(9) & nc & Chr(9) & d & Chr(9) & h & Chr(9) & s & Chr(9) & sao & Chr(9) & o & Chr(9) & ni
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
'armar grilla
If Check1 = 0 Then
  msf1.clear
  msf1.Rows = 1
  msf1.Cols = 10
  msf1.ColWidth(0) = 1300
  msf1.ColWidth(1) = 500
  msf1.ColWidth(2) = 1500
  msf1.ColWidth(3) = 1700
  msf1.ColWidth(4) = 1200
  msf1.ColWidth(5) = 1200
  msf1.ColWidth(6) = 1200
  msf1.ColWidth(7) = 2000
  msf1.ColWidth(8) = 1000
  msf1.ColWidth(9) = 500
  msf1.TextMatrix(0, 0) = "Fecha"
  msf1.TextMatrix(0, 2) = "Op."
  msf1.TextMatrix(0, 3) = "Nro.Comprobante"
  If Option4 = True Then
   msf1.TextMatrix(0, 4) = "Debe($)"
   msf1.TextMatrix(0, 5) = "Haber($)"
   msf1.TextMatrix(0, 6) = "Saldo($)"
  Else
   msf1.TextMatrix(0, 4) = "Debe(U$s)"
   msf1.TextMatrix(0, 5) = "Haber(U$s)"
   msf1.TextMatrix(0, 6) = "Saldo(U$s)"
  End If

  msf1.TextMatrix(0, 7) = "Obs."
  msf1.TextMatrix(0, 8) = "Num.Int."
  msf1.TextMatrix(0, 9) = " "
  For i = 0 To 3
    msf1.ColAlignment(i) = 1
  Next i
  For i = 4 To 6
    msf1.ColAlignment(i) = 9
  Next i
  For i = 7 To 8
    msf1.ColAlignment(i) = 1
  Next i

  
  
Else
   msf1.clear
  msf1.Rows = 1
  msf1.Cols = 11
  msf1.ColWidth(0) = 1300
  msf1.ColWidth(1) = 500
  msf1.ColWidth(2) = 1500
  msf1.ColWidth(3) = 1700
  msf1.ColWidth(4) = 1200
  msf1.ColWidth(5) = 1200
  msf1.ColWidth(6) = 1200
  msf1.ColWidth(7) = 1200
  msf1.ColWidth(8) = 2000
  msf1.ColWidth(9) = 1000
  msf1.ColWidth(10) = 500
  msf1.TextMatrix(0, 0) = "Fecha"
  msf1.TextMatrix(0, 2) = "Op."
  msf1.TextMatrix(0, 3) = "Nro.Comprobante"
  msf1.TextMatrix(0, 4) = "Debe"
  msf1.TextMatrix(0, 5) = "Haber"
  msf1.TextMatrix(0, 6) = "Saldo "
  If Option4 = True Then
   msf1.TextMatrix(0, 4) = "Debe($)"
   msf1.TextMatrix(0, 5) = "Haber($)"
   msf1.TextMatrix(0, 6) = "Saldo($)"
   msf1.TextMatrix(0, 7) = "Saldo(U$s)"
  Else
   msf1.TextMatrix(0, 4) = "Debe(U$s)"
   msf1.TextMatrix(0, 5) = "Haber(U$s)"
   msf1.TextMatrix(0, 6) = "Saldo(U$s)"
   msf1.TextMatrix(0, 7) = "Saldo($)"
  End If
  msf1.TextMatrix(0, 8) = "Obs."
  msf1.TextMatrix(0, 9) = "Num.Int."
  msf1.TextMatrix(0, 10) = " "
 For i = 0 To 3
    msf1.ColAlignment(i) = 1
  Next i
  For i = 4 To 7
    msf1.ColAlignment(i) = 9
  Next i
  For i = 8 To 9
    msf1.ColAlignment(i) = 1
  Next i
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

Private Sub cal1_DblClick()
t_fecha = cal1
cal1.Visible = False
End Sub

Private Sub cal1_LostFocus()
t_fecha = cal1
cal1.Visible = False
End Sub

Private Sub Command5_Click()
com_proveedor.t_id = c_prov.ItemData(c_prov.ListIndex)
com_proveedor.carga
com_proveedor.Show

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

Call carga_proveedores(c_prov)
c_prov.ListIndex = 0
t_sucursal = Format$(glo.sucursal, "0000")
Call armagrid
Call barraesag(Me)
cal1.Visible = False
Option1 = True
Option4 = True
'Load vta_clientes
Call carga_zonas(c_zona)
c_zona.AddItem "<Todas>", 0
c_zona.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Unload vta_clientes
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F7] Imprime - [ENTER] Visualiza - [F10] Ajuste CtaCte - [F11] Excel  "
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
      c(0) = 9
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
      
      c(0) = 10
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
     Call imprimegrid(msf1, c(), "ESTADO DE CUENTA PROVEDORES", "", "Provedor: " & c_prov, "Periodo: " & t_fecha & "  " & t_fecha2, 85, 7, True, False)

    End If
         
  End If
  
End If

If KeyCode = vbKeyF10 Then
  Load COM_ajustesint
  COM_ajustesint.c_prov.ListIndex = buscaindice(COM_ajustesint.c_prov, c_prov.ItemData(c_prov.ListIndex))
  COM_ajustesint.Show
  
End If




If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If

End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    Load cc_detalle
    cc_detalle.t_idprov = c_prov.ItemData(c_prov.ListIndex)
    cc_detalle.t_prov = c_prov
    cc_detalle.t_sucursal = Mid$(msf1.TextMatrix(msf1.Row, 3), 3, 4)
    cc_detalle.t_letra = Mid$(msf1.TextMatrix(msf1.Row, 3), 1, 1)
    cc_detalle.t_numcomp = Mid$(msf1.TextMatrix(msf1.Row, 3), 8, 8)
    cc_detalle.t_tipocomp = msf1.TextMatrix(msf1.Row, 1)
    If Check1 = 0 Then
      cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 8)
    Else
      cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 9)
    End If
    cc_detalle.Show
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

Private Sub Option4_Click()
Check1.Caption = "Muestra Saldo en U$s"
End Sub

Private Sub t_fecha_Click()
t_fecha = ""
End Sub

Private Sub t_fecha_DblClick()
cal1.Visible = True
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

Private Sub t_fecha2_Click()
t_fecha2 = ""
End Sub
