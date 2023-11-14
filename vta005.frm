VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_vercomp 
   BackColor       =   &H00E0E0E0&
   Caption         =   "COMPROBANTES EMITIDOS"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   5400
      TabIndex        =   42
      Top             =   7560
      Width           =   2895
      Begin VB.ComboBox c_moneda 
         Height          =   315
         ItemData        =   "vta005.frx":0000
         Left            =   1200
         List            =   "vta005.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Moneda:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipo Clientes"
      Height          =   615
      Left            =   5400
      TabIndex        =   39
      Top             =   6960
      Width           =   2895
      Begin VB.CheckBox Check2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Exportacion"
         Height          =   255
         Left            =   1560
         TabIndex        =   41
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nacionales"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Venta"
      Height          =   615
      Left            =   2640
      TabIndex        =   37
      Top             =   7560
      Width           =   2655
      Begin VB.ComboBox c_v 
         Height          =   315
         ItemData        =   "vta005.frx":002A
         Left            =   240
         List            =   "vta005.frx":0037
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cambiar"
      Height          =   975
      Left            =   8400
      TabIndex        =   32
      Top             =   7200
      Width           =   1095
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   120
         Picture         =   "vta005.frx":0062
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ordenados por:"
      Height          =   615
      Left            =   240
      TabIndex        =   25
      Top             =   7560
      Width           =   2295
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cliente"
         Height          =   255
         Left            =   1080
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Importe Total "
      Height          =   615
      Left            =   240
      TabIndex        =   15
      Top             =   6960
      Width           =   5055
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox t_importe 
         Height          =   285
         Left            =   3960
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mayores"
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Menores"
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Iguales"
         Height          =   255
         Left            =   3000
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   2520
      TabIndex        =   14
      Top             =   1200
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   179568641
      CurrentDate     =   38754
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5295
      Left            =   240
      TabIndex        =   4
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
      TabIndex        =   9
      Top             =   0
      Width           =   11535
      Begin VB.ComboBox c_c 
         Height          =   315
         ItemData        =   "vta005.frx":036C
         Left            =   9240
         List            =   "vta005.frx":0379
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox t_cliente 
         Height          =   285
         Left            =   5280
         TabIndex        =   34
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox c_estadopago 
         Height          =   315
         ItemData        =   "vta005.frx":0397
         Left            =   9240
         List            =   "vta005.frx":03A4
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   840
         Width           =   2175
      End
      Begin VB.ComboBox c_suc 
         Height          =   315
         Left            =   9240
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   480
         Width           =   2175
      End
      Begin VB.ComboBox c_vend 
         Height          =   315
         Left            =   1320
         TabIndex        =   23
         Text            =   "c"
         Top             =   1200
         Width           =   3975
      End
      Begin VB.CommandButton Command5 
         Height          =   255
         Left            =   7320
         Picture         =   "vta005.frx":03C1
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   255
      End
      Begin VB.ComboBox c_tipocomp 
         Height          =   315
         Left            =   9240
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   120
         Width           =   2175
      End
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Text            =   "c_prov"
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Contado o Ctacte:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7680
         TabIndex        =   36
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Estado Pago:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7680
         TabIndex        =   31
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Sucursal:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7680
         TabIndex        =   29
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Vendedor:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   9960
         TabIndex        =   21
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Tipo Comprobante:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7680
         TabIndex        =   13
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2880
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
         Width           =   1095
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
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   6
      Top             =   7200
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta005.frx":0733
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
         Picture         =   "vta005.frx":0FB5
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
            TextSave        =   "10/11/2023"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:38 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_vercomp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer


Sub carga()
  espere.Show
  espere.Label1 = "Cargando comprobantes emitidos...."
  espere.Refresh
  Call armagrid
  
  
  q = "select * from vta_02, vta_06, vta_01 where vta_02.[id_tipocomp] = vta_06.[id_tipocomp] and vta_02.[id_cliente] = vta_01.[id_cliente] and vta_02.[sucursal_ingreso] = vta_06.[sucursal] "
  c = " and "
  If c_prov.ListIndex > 0 Then
     q = q & c & " vta_02.[id_cliente] = " & c_prov.ItemData(c_prov.ListIndex)
  End If
  
  If c_tipocomp.ListIndex > 0 Then
    q = q & c & " vta_02.[id_tipocomp] = " & c_tipocomp.ItemData(c_tipocomp.ListIndex)
  End If
  
  If t_fecha <> "" And IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If t_fecha2 <> "" And IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
    
   If c_moneda.ListIndex > 0 Then
    If c_moneda.ListIndex = 1 Then
        q = q & c & " vta_02.[moneda] = 'P'"
    Else
         q = q & c & " vta_02.[moneda] = 'D'"
    End If
   End If
  
    
   If c_vend.ListIndex > 0 Then
    q = q & c & " vta_02.[id_vendedor] = " & c_vend.ItemData(c_vend.ListIndex)
   End If

   If c_suc.ListIndex > 0 Then
    q = q & c & " vta_02.[sucursal_ingreso] = " & Val(c_suc)
   End If

   If Check1 = 1 Then
    q = q & c & " [id_tipoiva] <> 8"
   End If

   If Check2 = 1 Then
     q = q & c & " [id_tipoiva] = 8"
   End If
   
   
   If c_estadopago.ListIndex > 0 Then
      If c_estadopago.ListIndex = 1 Then
          q = q & c & " [estado_pago] = 'P'"
      Else
          q = q & c & " [estado_pago] <> 'P'"
      End If
   End If

  If t_cliente <> "" Then
    q = q & c & " [cliente02] like '%" & t_cliente & "%'"
  End If
    
  If Option1 = False Then
    If Option4 = True Then
       q = q & c & " [total] >= " & Val(t_importe)
    Else
      If Option5 = True Then
       q = q & c & " [total] <= " & Val(t_importe)
      Else
        q = q & c & " [total] = " & Val(t_importe)
      End If
   End If
 End If
 
 If c_c.ListIndex > 0 Then
  If c_c.ListIndex = 1 Then
     'contado
     q = q & c & "[contado] = 'S'"
     c = " and "
  Else
     q = q & c & " [contado] = 'N'  and [cta_cte] <> 'N'"
     c = " and "
  End If
 End If
 
 If c_v.ListIndex > 0 Then
  If c_v.ListIndex = 1 Then
     'solo ventas
     q = q & c & " vta_02.[venta] <> 'N'"
     c = " and "
  Else
     q = q & c & " vta_02.[venta] = 'N' "
     c = " and "
  End If
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
     If rs("cta_cte") <> "H" Then
       d = Format$(rs("total"), "######0.00")
       s = Format$(rs("subtotal"), "######0.00")
       i = Format$(rs("vta_02.iva"), "######0.00")
     Else
       d = Format$(-rs("total"), "######0.00")
       s = Format$(-rs("subtotal"), "######0.00")
       i = Format$(-rs("vta_02.iva"), "######0.00")
     
     End If
     t = t + Val(d)
     ni = rs("num_int")
     If rs("vta_02.moneda") = "P" Then
      m = "$"
     Else
      m = "U$s"
     End If
     msf1.AddItem F & Chr(9) & cp & Chr(9) & p & Chr(9) & CTC & Chr(9) & tc & Chr(9) & nc & Chr(9) & d & Chr(9) & rs("estado") & Chr(9) & rs("num_int") & Chr(9) & m & Chr(9) & rs("estado_pago") & Chr(9) & rs("vta_02.observaciones") & Chr(9) & rs("contado") & Chr(9) & s & Chr(9) & i
     reg = reg + 1
     Label5 = reg
     Label5.Refresh
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
msf1.Cols = 15
msf1.ColWidth(0) = 1300
msf1.ColWidth(1) = 700 'cod prov
msf1.ColWidth(2) = 3500
msf1.ColWidth(3) = 500
msf1.ColWidth(4) = 1700
msf1.ColWidth(5) = 1700
msf1.ColWidth(6) = 1200
msf1.ColWidth(7) = 500
msf1.ColWidth(8) = 1000
msf1.ColWidth(9) = 700
msf1.ColWidth(10) = 700
msf1.ColWidth(11) = 2500
msf1.ColWidth(12) = 700
msf1.ColWidth(13) = 1200
msf1.ColWidth(14) = 1200


msf1.TextMatrix(0, 0) = "Fecha"
msf1.TextMatrix(0, 1) = ""
msf1.TextMatrix(0, 2) = "Cliente"
msf1.TextMatrix(0, 3) = ""
msf1.TextMatrix(0, 4) = "Operacion"
msf1.TextMatrix(0, 5) = "Nro.Comprobante"
msf1.TextMatrix(0, 6) = "Total"
msf1.TextMatrix(0, 7) = "Estado"
msf1.TextMatrix(0, 8) = "Num.Int."
msf1.TextMatrix(0, 9) = "Moneda"
msf1.TextMatrix(0, 10) = "Cobrado"
msf1.TextMatrix(0, 11) = "Observaciones"
msf1.TextMatrix(0, 12) = "Ctdo"
msf1.TextMatrix(0, 13) = "Neto Grav."
msf1.TextMatrix(0, 14) = "Iva"


For i = 0 To 10
    msf1.ColAlignment(i) = 1 'izq
Next i
msf1.ColAlignment(6) = 9 'der
msf1.ColAlignment(8) = 9 'der
msf1.ColAlignment(13) = 9 'der
msf1.ColAlignment(14) = 9 'der

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

Private Sub c_suc_LostFocus()
If c_suc.ListIndex < 0 Then
  c_suc.ListIndex = 0
End If
End Sub

Private Sub c_tipocomp_LostFocus()
If c_tipocomp.ListIndex < 0 Then
  c_tipocomp.ListIndex = 0
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


Private Sub Command1_Click()
gen_seleccionarimp.Show
End Sub

Private Sub Command5_Click()
If c_prov.ListIndex > 0 Then
 vta_clientes.t_id = c_prov.ItemData(c_prov.ListIndex)
 vta_clientes.carga
 vta_clientes.Show
End If
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

c_moneda.ListIndex = 0


Call carga_vendedores(c_vend)
c_vend.AddItem "<Todos>", 0
c_vend.ListIndex = 0

Call carga_SUCURSALES(c_suc)
c_suc.AddItem "<Todas>", 0
c_suc.ListIndex = 0

Set rs = New ADODB.Recordset
q = "select * from vta_06 where [sucursal] = " & glo.sucursal
rs.Open q, cn1
Call llena_combo(rs, "descripcion", "id_tipocomp", c_tipocomp, True)
Set rs = Nothing
c_tipocomp.AddItem "<Todos>", 0
c_tipocomp.ListIndex = 0
c_estadopago.ListIndex = 0
c_c.ListIndex = 0
c_v.ListIndex = 0

Call armagrid
Call barraesag(Me)
Option1 = True
Option2 = True
Load vta_clientes
End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload vta_clientes
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F1] Cliente -  [F8] Borra - [F11] Excel "
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

Private Sub t_cliente_GotFocus()
t_cliente = ""
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
