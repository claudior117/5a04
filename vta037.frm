VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form vta_verfletes 
   BackColor       =   &H00E0E0E0&
   Caption         =   "INFORME DE FLETES"
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
      Caption         =   "Cambiar"
      Height          =   975
      Left            =   8400
      TabIndex        =   17
      Top             =   7200
      Width           =   1095
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   120
         Picture         =   "vta037.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ordenados por:"
      Height          =   615
      Left            =   240
      TabIndex        =   12
      Top             =   7080
      Width           =   3615
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cliente"
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5295
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9340
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   11535
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   1440
         TabIndex        =   25
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox t_dest 
         Height          =   285
         Left            =   6720
         MaxLength       =   25
         TabIndex        =   23
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox t_origen 
         Height          =   285
         Left            =   6720
         MaxLength       =   25
         TabIndex        =   22
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox t_merc 
         Height          =   285
         Left            =   6720
         MaxLength       =   25
         TabIndex        =   21
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox t_chofer 
         Height          =   285
         Left            =   6720
         MaxLength       =   25
         TabIndex        =   20
         Top             =   120
         Width           =   2895
      End
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Cliente:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Destino:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5160
         TabIndex        =   19
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Origen:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5160
         TabIndex        =   16
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Mercaderia"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5160
         TabIndex        =   15
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   9960
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Chofer:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5160
         TabIndex        =   10
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   4
      Top             =   7200
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta037.frx":030A
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
         Picture         =   "vta037.frx":0B8C
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
            TextSave        =   "03/03/2011"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "05:11 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_verfletes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer


Sub carga()
  espere.Show
  espere.Label1 = "Cargando flete realizados...."
  espere.Refresh
  Call armagrid
  q = "select * from vta_02, vta_011 where vta_02.[num_int] = vta_011.[num_int]  "
  c = " and "
  If t_fecha <> "" And IsDate(t_fecha) Then
     q = q & c & " datevalue(vta_011.[fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If t_fecha2 <> "" And IsDate(t_fecha2) Then
     q = q & c & " datevalue(vta_011.[fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  If t_chofer <> "" Then
    q = q & c & " [chofer] like '%" & t_chofer & "%'"
  End If
  
  If t_merc <> "" Then
    q = q & c & " [detalle] like '%" & t_merc & "%'"
  End If
  
  If t_origen <> "" Then
    q = q & c & " [origen] like '%" & t_origen & "%'"
  End If
  
  If t_dest <> "" Then
    q = q & c & " [destino] like '%" & t_dest & "%'"
  End If
  
  If c_prov.ListIndex > 0 Then
     q = q & c & " [id_cliente] = " & c_prov.ItemData(c_prov.ListIndex)
  End If
  
    
 If Option2 = True Then
    q = q & " order by vta_011.[fecha], [num_comp]"
 Else
    q = q & " order by [cliente02], vta_011.[fecha], [num_comp]"
 End If
 
  Set rs = New ADODB.Recordset

  rs.Open q, cn1
  t = 0
  reg = 0
  While Not rs.EOF
     f = rs("vta_011.fecha")
     nc = rs("letra") & " " & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
     t = Format$(rs("importe_total"), "######0.00")
     c = rs("cliente02")
     tot = tot + Val(t)
     ni = rs("vta_02.num_int")
     msf1.AddItem f & Chr(9) & rs("detalle") & Chr(9) & rs("chofer") & Chr(9) & rs("origen") & Chr(9) & rs("destino") & Chr(9) & t & Chr(9) & rs("toneladas") & Chr(9) & rs("kmts") & Chr(9) & rs("tarifa") & Chr(9) & rs("carta_porte") & Chr(9) & c & Chr(9) & nc & Chr(9) & rs("vta_02.fecha")
     reg = reg + 1
     Label5 = reg
     Label5.Refresh
    rs.MoveNext
  Wend
  tt = Format$(suma_msflexgrid(msf1, 6), "#####0.00")
  tk = Format$(suma_msflexgrid(msf1, 7), "#####0.00")
  tf = Format$(suma_msflexgrid(msf1, 8), "#####0.00")
  msf1.AddItem "" & Chr(9) & "" & Chr(9) & " " & Chr(9) & "" & Chr(9) & " " & Chr(9) & "-----------------------------" & Chr(9) & "-----------------------------" & Chr(9) & "-----------------------------" & Chr(9) & "-----------------------------"

  msf1.AddItem "" & Chr(9) & "" & Chr(9) & "Fletes: " & reg & Chr(9) & "" & Chr(9) & "Totales:" & Chr(9) & Format$(tot, "#####0.00") & Chr(9) & tt & Chr(9) & tk
  msf1.AddItem ""
  If reg > 0 Then
    msf1.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "Promedios:" & Chr(9) & Format$(tot / reg, "#####0.00") & Chr(9) & Format$(Val(tt) / reg, "#####0.00") & Chr(9) & Format$(Val(tk) / reg, "#####0.00") & Chr(9) & Format$(Val(tf) / reg, "#####0.00")
  End If
  Unload espere
   
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
msf1.Cols = 14
msf1.ColWidth(0) = 1100
msf1.ColWidth(1) = 1500
msf1.ColWidth(2) = 1500
msf1.ColWidth(3) = 1500
msf1.ColWidth(4) = 1500
msf1.ColWidth(5) = 1200
msf1.ColWidth(6) = 900
msf1.ColWidth(7) = 900
msf1.ColWidth(8) = 900
msf1.ColWidth(9) = 1000
msf1.ColWidth(10) = 1700
msf1.ColWidth(11) = 1500
msf1.ColWidth(12) = 1100
msf1.ColWidth(13) = 0

msf1.TextMatrix(0, 0) = "Fecha"
msf1.TextMatrix(0, 1) = "Mercaderia"
msf1.TextMatrix(0, 2) = "Chofer/camion"
msf1.TextMatrix(0, 3) = "Origen"
msf1.TextMatrix(0, 4) = "Destino"
msf1.TextMatrix(0, 5) = "Imp. c/iva"
msf1.TextMatrix(0, 6) = "Ton."
msf1.TextMatrix(0, 7) = "Kmts"
msf1.TextMatrix(0, 8) = "Tarifa"
msf1.TextMatrix(0, 9) = "C.Porte"
msf1.TextMatrix(0, 10) = "Cliente"
msf1.TextMatrix(0, 11) = "Nro.Fact"
msf1.TextMatrix(0, 12) = "Fecha Fact"

For i = 0 To 4
    msf1.ColAlignment(i) = 1 'izq
Next i
For i = 5 To 9
    msf1.ColAlignment(i) = 9 'der
Next i
For i = 10 To 12
    msf1.ColAlignment(i) = 1 'izq
Next i

End Sub






Private Sub c_prov_LostFocus()
If c_prov.ListIndex < 0 Then
  c_prov.ListIndex = 0
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
Call carga_clientes(c_prov)
c_prov.AddItem "<Todos>", 0
c_prov.ListIndex = 0


Call armagrid
Call barraesag(Me)
Option1 = True
Option2 = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload vta_clientes
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F1] Cliente -  [F8] Borra - [F3] Cambia Datos - [F11] Excel "
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
    c(11) = 11
    c(12) = 12
    c(13) = -1
    c(14) = -1
    
    'For i = 6 To 14
    '  c(i) = -1
    'Next i
    Call imprimegrid(msf1, c(), "INFORME DE FLETES FACTURADOS", "Cliente:" & c_prov & "           Chofer: " & t_chofer, "Fecha desde: " & t_fecha & "  Fecha hasta: " & t_fecha2, "Origen/Destino: " & t_origen & "/" & t_destino, 55, 8, True, False, "H")
  End If

End If



 If KeyCode = vbKeyF8 Then
  Call nivel_acceso(2)
  If para.id_grupo_modulo_actual >= 8 Then
   J = MsgBox("Confirma Eliminar Comprobante Nro." & msf1.TextMatrix(msf1.RowSel, 5), 4)
   If J = 6 Then
      indice = msf1.RowSel
      Set cl_compvta = New comprobantes_venta
      cl_compvta.cargar2 (Val(msf1.TextMatrix(indice, 8)))
      cl_compvta.borrar
      Set cl_compvta = Nothing
      MsgBox ("Operacion Terminada")
      Call carga
   End If
  End If
End If


If KeyCode = vbKeyF5 Then
 J = MsgBox("Prepare Impresora y Confirme", 4)
 If J = 6 Then
        Call nivel_acceso(2)
        If para.id_grupo_modulo_actual >= 6 Then
           Set cl_compvta = New comprobantes_venta
           cl_compvta.cargar2 (Val(msf1.TextMatrix(msf1.Row, 8)))
           cl_compvta.imprimir
        End If
  End If
End If

If KeyCode = vbKeyF1 Then
  If Val(msf1.TextMatrix(msf1.Row, 1)) > 0 Then
     vta_clientes.t_id = Val(msf1.TextMatrix(msf1.Row, 1))
     vta_clientes.carga
     vta_clientes.Show
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
    vta_cc_detalle.T_IDPROV = msf1.TextMatrix(msf1.Row, 1)
    vta_cc_detalle.t_prov = msf1.TextMatrix(msf1.Row, 2)
    vta_cc_detalle.t_sucursal = Mid$(msf1.TextMatrix(msf1.Row, 5), 3, 4)
    vta_cc_detalle.t_letra = Mid$(msf1.TextMatrix(msf1.Row, 5), 1, 1)
    vta_cc_detalle.t_numcomp = Mid$(msf1.TextMatrix(msf1.Row, 5), 8, 8)
    vta_cc_detalle.t_tipocomp = msf1.TextMatrix(msf1.Row, 3)
    vta_cc_detalle.t_NUMINT = msf1.TextMatrix(msf1.Row, 8)
    vta_cc_detalle.Show
  End If
End If

End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight

End Sub


Private Sub T_chofer_GotFocus()
t_chofer = ""
End Sub

Private Sub t_dest_GotFocus()
t_dest = ""
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

Private Sub t_importe_GotFocus()
t_importe = ""
End Sub

Private Sub t_merc_GotFocus()
t_merc = ""
End Sub

Private Sub t_origen_GotFocus()
t_origen = ""
End Sub
