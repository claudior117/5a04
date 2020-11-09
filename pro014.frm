VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form prod_verempaque 
   BackColor       =   &H00E0E0E0&
   Caption         =   "VER ORDENES DE EMPAQUE"
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
      TabIndex        =   13
      Top             =   7200
      Width           =   1095
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   120
         Picture         =   "pro014.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   2520
      TabIndex        =   12
      Top             =   1200
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   99549185
      CurrentDate     =   38754
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5295
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9340
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1335
      Left            =   240
      TabIndex        =   8
      Top             =   0
      Width           =   8295
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   2
         Top             =   720
         Width           =   1335
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
         Width           =   6615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2880
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   5
      Top             =   7200
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "pro014.frx":030A
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
         Picture         =   "pro014.frx":0B8C
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
            TextSave        =   "06/08/2017"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "06:08 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "prod_verempaque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer


Sub carga()
  espere.Show
  espere.Label1 = "Cargando ordenes de empaque...."
  espere.Refresh
  Call armagrid
  q = "select * from vta_02, vta_06, vta_01 where vta_02.[id_tipocomp] = vta_06.[id_tipocomp] and vta_02.[id_cliente] = vta_01.[id_cliente] and vta_02.[sucursal_ingreso] = vta_06.[sucursal]"
  c = " and "
  If c_prov.ListIndex > 0 Then
     q = q & c & " vta_02.[id_cliente] = " & c_prov.ItemData(c_prov.ListIndex)
  End If
  
 q = q & c & " vta_02.[id_tipocomp] = 150"
  
  If t_fecha <> "" And IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If t_fecha2 <> "" And IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
    
   
   
  If t_cliente <> "" Then
    q = q & c & " [cliente02] like '%" & t_cliente & "%'"
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
  While Not rs.EOF
     F = rs("fecha")
     CTC = Format$(rs("vta_02.ID_TIPOCOMP"), "000")
     tc = rs("descripcion")
     nc = rs("letra") & " " & Format$(rs("vta_02.sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
     d = Format$(rs("total"), "######0.00")
     cp = Format$(rs("vta_02.id_cliente"), "0000")
     p = rs("cliente02")
     ni = rs("num_int")
     msf1.AddItem F & Chr(9) & cp & Chr(9) & p & Chr(9) & CTC & Chr(9) & tc & Chr(9) & nc & Chr(9) & rs("vta_02.observaciones") & Chr$(9) & "" & Chr$(9) & rs("num_int")
     reg = reg + 1
     Label5 = reg
    rs.MoveNext
  Wend
  msf1.AddItem ""
  msf1.AddItem "" & Chr(9) & "" & Chr(9) & "Comprobantes: " & reg & Chr(9) & "" & Chr(9) & "" & Chr(9) & "Totales:" & Chr(9) & Format$(t, "#####0.00") & Chr(9) & ""
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
msf1.Cols = 9
msf1.ColWidth(0) = 1300
msf1.ColWidth(1) = 700 'cod prov
msf1.ColWidth(2) = 3500
msf1.ColWidth(3) = 500
msf1.ColWidth(4) = 1700
msf1.ColWidth(5) = 1700
msf1.ColWidth(6) = 3000
msf1.ColWidth(7) = 500
msf1.ColWidth(8) = 1000


msf1.TextMatrix(0, 0) = "Fecha"
msf1.TextMatrix(0, 1) = ""
msf1.TextMatrix(0, 2) = "Cliente"
msf1.TextMatrix(0, 3) = ""
msf1.TextMatrix(0, 4) = "Operacion"
msf1.TextMatrix(0, 5) = "Nro.Comprobante"

msf1.TextMatrix(0, 6) = "Observaciones"
msf1.TextMatrix(0, 7) = ""
msf1.TextMatrix(0, 8) = "Num.Int."


For i = 0 To 8
    msf1.ColAlignment(i) = 1 'izq
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
cal1.Visible = False
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
Me.StatusBar1.Panels.Item(2) = "[F1] Cliente -  [F8] Borra - [F11] Excel "
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
    c(5) = 8

    For i = 6 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "COMPROBANTES EMITIDOS", "Cliente:" & c_prov & "           Estado: " & c_estado, "Fecha desde: " & t_fecha & "  Fecha hasta: " & t_fecha2, "Vendedor: " & c_vend, 72, 8, True, False)
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
    vta_cc_detalle.t_idprov = msf1.TextMatrix(msf1.Row, 1)
    vta_cc_detalle.t_prov = msf1.TextMatrix(msf1.Row, 2)
    vta_cc_detalle.t_sucursal = Mid$(msf1.TextMatrix(msf1.Row, 5), 3, 4)
    vta_cc_detalle.t_letra = Mid$(msf1.TextMatrix(msf1.Row, 5), 1, 1)
    vta_cc_detalle.t_numcomp = Mid$(msf1.TextMatrix(msf1.Row, 5), 8, 8)
    vta_cc_detalle.t_tipocomp = msf1.TextMatrix(msf1.Row, 3)
    vta_cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 8)
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

