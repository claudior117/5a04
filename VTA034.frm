VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form vta_informevta4 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CIERRE DE VENTAS DIARIO"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   12120
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Moneda"
      Height          =   495
      Left            =   8160
      TabIndex        =   14
      Top             =   120
      Width           =   3495
      Begin VB.OptionButton Option5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "U$s Dolares"
         Height          =   195
         Left            =   1680
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "$ Pesos"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   4200
      TabIndex        =   11
      Top             =   120
      Width           =   3255
      Begin VB.ComboBox c_sucursal 
         Height          =   315
         ItemData        =   "VTA034.frx":0000
         Left            =   1680
         List            =   "VTA034.frx":0002
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Punto Venta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   4320
      TabIndex        =   9
      Top             =   0
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   200933377
      CurrentDate     =   38750
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   3615
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
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
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "VTA034.frx":0004
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
         Picture         =   "VTA034.frx":0886
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
      Top             =   8340
      Width           =   12120
      _ExtentX        =   21378
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
            TextSave        =   "21/08/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:20 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5775
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   10186
      _Version        =   393216
      FixedRows       =   0
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
Attribute VB_Name = "vta_informevta4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
'FIXIT: Declare 'ti' con un tipo de datos de enlace en tiempo de compilación               FixIT90210ae-R1672-R1B8ZE
Dim ti, t As Double
'FIXIT: Declare 'reg' con un tipo de datos de enlace en tiempo de compilación              FixIT90210ae-R1672-R1B8ZE
Dim reg, regi As Integer
Dim tipomoneda As String


Sub carga()
  
  Call armagrid
  'consulta ventas cuenta corriente
  
   espere.Show
   espere.Label1 = "Espere....  Obteniendo ventas en Cuenta Corriente"
   espere.Refresh
  q = "select * from vta_02,  vta_06 where vta_02.[id_tipocomp] = vta_06.[id_tipocomp] and vta_02.[venta] <> 'N' and vta_02.[sucursal] = vta_06.[sucursal]"
  c = " and "
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  
  If c_sucursal.ListIndex > 0 Then
     q = q & c & " [sucursal_ingreso] = " & Val(c_sucursal)
  End If
  
  q1 = q & " and [cta_cte] <> 'N' and [contado] = 'N' "
  
  Set rs = New ADODB.Recordset
  rs.Open q1, cn1
  tdcc = 0
  tccc = 0
  While Not rs.EOF
     If rs("vta_02.moneda") = "P" Then
       m = 1
     Else
       m = rs("cotizacion_dolar")
     End If
     If rs("ctacte") = "D" Then
       tdcc = tdcc + (rs("total") * m)
     Else
       tccc = tccc + (rs("total") * m)
     End If
     rs.MoveNext
  Wend
  Set rs = Nothing
  
  msf1.TextMatrix(2, 1) = Format$(tdcc, "########0.00")
  msf1.TextMatrix(2, 2) = Format$(tccc, "########0.00")
  msf1.TextMatrix(2, 3) = Format$(Val(msf1.TextMatrix(2, 1)) - Val(msf1.TextMatrix(2, 2)), "########0.00")
  
  espere.Label1 = "Espere....  Obteniendo ventas Contado"
  espere.Refresh
  q1 = q & " and [contado] = 'S'"
  Set rs = New ADODB.Recordset
  rs.Open q1, cn1
  tdc = 0
  tcc = 0
  While Not rs.EOF
     If rs("vta_02.moneda") = "P" Then
       m = 1
     Else
       m = rs("cotizacion_dolar")
     End If
     If rs("vta_02.venta") = "S" Then
       tdc = tdc + (rs("total") * m)
     Else
       tcc = tcc + (rs("total") * m)
     End If
     rs.MoveNext
  Wend
  Set rs = Nothing
  
  msf1.TextMatrix(3, 1) = Format$(tdc, "########0.00")
  msf1.TextMatrix(3, 2) = Format$(tcc, "########0.00")
  msf1.TextMatrix(3, 3) = Format$(Val(msf1.TextMatrix(3, 1)) - Val(msf1.TextMatrix(3, 2)), "########0.00")
     
  msf1.TextMatrix(5, 1) = Format$(Val(msf1.TextMatrix(2, 1)) + Val(msf1.TextMatrix(3, 1)), "########0.00")
  msf1.TextMatrix(5, 2) = Format$(Val(msf1.TextMatrix(2, 2)) + Val(msf1.TextMatrix(3, 2)), "########0.00")
  msf1.TextMatrix(5, 3) = Format$(Val(msf1.TextMatrix(2, 3)) + Val(msf1.TextMatrix(3, 3)), "########0.00")
  
  
   espere.Label1 = "Espere....  Obteniendo Recibos..."
  espere.Refresh
  
  q = "select * from vta_02 where [id_tipocomp] = 50"
  c = " and "
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  
  If c_sucursal.ListIndex > 0 Then
     q = q & c & " [sucursal_ingreso] = " & Val(c_sucursal)
  End If
  
  
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  tr = 0
  While Not rs.EOF
       tr = tr + (rs("total"))
     rs.MoveNext
  Wend
  Set rs = Nothing
  msf1.TextMatrix(9, 1) = Format$(tr, "########0.00")
  msf1.TextMatrix(10, 1) = Format$(Val(msf1.TextMatrix(3, 3)), "########0.00")
  msf1.TextMatrix(12, 1) = Format$(Val(msf1.TextMatrix(9, 1)) + Val(msf1.TextMatrix(10, 1)), "########0.00")
  Unload espere
   
End Sub
Private Sub btnacepta_Click()
   QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
  QUERY = QUERY & " VALUES ('Cierre de Ventas Diario " & "', " & para.id_usuario & ", 'V', 0, '" & Now & "', ' ', 18, " & 0 & ")"
  cn1.BeginTrans
  cn1.Execute QUERY
  cn1.CommitTrans
   Call carga
  
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub


Sub renglon2()
     F = rs("fecha")
     CTC = Format$(rs("vta_02.ID_TIPOCOMP"), "000")
     tc = rs("abreviatura")
     nc = rs("letra") & " " & Format$(rs("vta_02.sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
     If tipomoneda = rs("vta_02.moneda") Then
        d = Format$(rs("total"), "######0.00")
     Else
        d = Format$(rs("total_otra_moneda"), "######0.00")
     End If
     cp = Format$(rs("vta_02.id_cliente"), "0000")
     p = rs("vta_01.denominacion")
     v = rs("vta_05.denominacion")
     If rs("vta_02.venta") = "S" Then
       t = t + Val(d)
       ti = ti + Val(d)
     
     Else
       t = t - Val(d)
       ti = ti - Val(d)
       d = -d
     End If
     ni = rs("num_int")
     msf1.AddItem F & Chr(9) & cp & Chr(9) & p & Chr(9) & tc & " " & nc & Chr(9) & d & Chr(9) & v & Chr(9) & rs("num_int")
     reg = reg + 1
     regi = regi + 1
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
l = "====================================================================="
l2 = "---------------------------------------------------------------------"

msf1.clear
msf1.Rows = 14
msf1.Cols = 5

msf1.ColWidth(0) = 4000
msf1.ColWidth(1) = 1700
msf1.ColWidth(2) = 1700
msf1.ColWidth(3) = 1700
msf1.ColWidth(4) = 400
msf1.TextMatrix(0, 0) = "VENTAS"
msf1.TextMatrix(0, 1) = "Facturacion"
msf1.TextMatrix(0, 2) = "N.C."
msf1.TextMatrix(0, 3) = "Totales"

msf1.TextMatrix(2, 0) = "Total Ventas en Cuenta Corriente"
msf1.TextMatrix(3, 0) = "Total Ventas Contado"

For i = 1 To 3
 msf1.TextMatrix(4, i) = l2
Next i
msf1.TextMatrix(5, 0) = "Total Ventas"

msf1.TextMatrix(8, 0) = "INGRESOS"
msf1.TextMatrix(9, 0) = "Ingresos en Cuenta Corriente(Recibos)"
msf1.TextMatrix(10, 0) = "Ingreso por Ventas Contado"
msf1.TextMatrix(11, 1) = l2
msf1.TextMatrix(12, 0) = "Total de Ingresos"



For J = 1 To 3
    msf1.ColAlignment(J) = 9 'der
Next J
End Sub

Private Sub Form_Load()
Call barraesag(Me)
cal1.Visible = False
Call armagrid
Call carga_SUCURSALES(c_sucursal)
c_sucursal.AddItem "<Todas>", 0
c_sucursal.ListIndex = 0

End Sub




Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F7] Imprime - [F11] Excel"

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
  Dim c(15) As Double
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    c(0) = 4
    c(1) = 0
    c(2) = 1
    c(3) = 2
    c(4) = 3
    For i = 5 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "CIERRE DIARIO VENTAS", "", "Periodo: " & t_fecha & " : " & t_fecha2, "Sucursal: " & c_sucursal, 60, 9, True, False)
  End If

End If

If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If
End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    Load cc_detalle
    vta_cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 6)
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
