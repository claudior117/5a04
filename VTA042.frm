VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_int_mora 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ACTUALIZACION DEUDA(CALCULO de INTERESES POR MORA)"
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
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Utiliza Fecha"
      Height          =   495
      Left            =   8640
      TabIndex        =   23
      Top             =   1200
      Width           =   3135
      Begin VB.OptionButton Option7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Comprobante"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Vencimiento"
         Height          =   195
         Left            =   1800
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos para raliar calculo de Intereses"
      Height          =   735
      Left            =   0
      TabIndex        =   22
      Top             =   7200
      Width           =   9495
      Begin VB.TextBox t_fechac 
         Height          =   285
         Left            =   4560
         TabIndex        =   30
         ToolTipText     =   $"VTA042.frx":0000
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox t_interes 
         Height          =   285
         Left            =   7920
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox t_dias 
         Height          =   285
         Left            =   1440
         TabIndex        =   27
         ToolTipText     =   $"VTA042.frx":0096
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Fecha Calculo:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   31
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Interes diario:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6600
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Dias Tolerancia:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Moneda"
      Height          =   495
      Left            =   5520
      TabIndex        =   19
      Top             =   1200
      Width           =   2775
      Begin VB.OptionButton Option5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "U$s Dolares"
         Height          =   195
         Left            =   1200
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "$ Pesos"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   240
      TabIndex        =   16
      Top             =   1080
      Width           =   3255
      Begin VB.ComboBox c_sucursal 
         Height          =   315
         Left            =   1680
         TabIndex        =   17
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
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   5520
      TabIndex        =   11
      Top             =   0
      Width           =   6255
      Begin VB.ComboBox c_vend 
         Height          =   315
         Left            =   1440
         TabIndex        =   15
         Top             =   720
         Width           =   4575
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00800080&
         Caption         =   "Vendedor:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800080&
         Caption         =   "Cliente:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   2520
      TabIndex        =   9
      Top             =   0
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   169410561
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
         Picture         =   "VTA042.frx":0122
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
         Picture         =   "VTA042.frx":09A4
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
            TextSave        =   "12/04/2021"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:40 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5295
      Left            =   0
      TabIndex        =   10
      Top             =   1800
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9340
      _Version        =   393216
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000FFFF&
      Caption         =   "Label9"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   8040
      Width           =   9375
   End
End
Attribute VB_Name = "vta_int_mora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
'FIXIT: Declare 'ti' con un tipo de datos de enlace en tiempo de compilación               FixIT90210ae-R1672-R1B8ZE
Dim ti, t As Double
'FIXIT: Declare 'reg' con un tipo de datos de enlace en tiempo de compilación              FixIT90210ae-R1672-R1B8ZE
Dim reg, regi As Integer
Dim tipomoneda, p, v As String


Sub carga()
  
  
  Call armagrid
  q = "select * from vta_02,  vta_06, vta_05  where vta_02.[id_tipocomp] = vta_06.[id_tipocomp] and  vta_02.[id_vendedor] = vta_05.[id_vendedor] and vta_02.[venta] <> 'N' and vta_02.[sucursal] = vta_06.[sucursal]"
  c = " and "
  If c_prov.ListIndex > 0 Then
     q = q & c & " vta_02.[id_cliente] = " & c_prov.ItemData(c_prov.ListIndex)
  End If
  
  
  If Option7 = True Then
   'fecha comp
   If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
   End If
  
   If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
   End If
  
  Else
   
    'fecha vencimiento
   If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha_vto]) >= datevalue('" & t_fecha & "')"
   End If
  
   If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha_vto]) <= datevalue('" & t_fecha2 & "')"
   End If
  
    
  End If
  
  If c_vend.ListIndex > 0 Then
     q = q & c & " vta_02.[Id_vendedor] = " & c_vend.ItemData(c_vend.ListIndex)
  End If
  
  If c_sucursal.ListIndex > 0 Then
     q = q & c & " [sucursal_ingreso] = " & Val(c_sucursal)
  End If
  
  If Option7 = True Then
    q = q & " order by [fecha], vta_02.[id_tipocomp], [num_comp]"
  Else
    q = q & " order by [fecha_vto], vta_02.[id_tipocomp], [num_comp]"
  End If
  
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  t = 0
  ti = 0
  reg = 0
  
  While Not rs.EOF
     Call renglon2
     rs.MoveNext
  Wend
  msf1.AddItem ""
  msf1.AddItem "" & Chr(9) & "Comprobantes: " & reg & Chr(9) & "" & Chr(9) & "Totales:" & Chr(9) & Format$(t, "#####0.00") & Chr(9) & "" & Chr(9) & Format$(ti, "#####0.00")

     
  
  
   
   
   
End Sub
Private Sub btnacepta_Click()
 espere.Show
 espere.Refresh
  QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
  QUERY = QUERY & " VALUES ('Calculo Interes por Mora " & "', " & para.id_usuario & ", 'V', 0, '" & Now & "', ' ', 23, " & 0 & ")"
  cn1.BeginTrans
  cn1.Execute QUERY
  cn1.CommitTrans

If Option4 = True Then
    tipomoneda = "P"
  Else
   tipomoneda = "D"
End If
   
 Call carga
 
 Unload espere
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub


Sub renglon2()
     If Option7 = True Then
       F = rs("fecha")
     Else
       F = rs("fecha_vto")
     End If
     
     
     fc = Format$(DateValue(t_fechac) - Val(t_dias), "dd-mm-yy")
     diasmora = DateValue(fc) - F
     CTC = Format$(rs("vta_02.ID_TIPOCOMP"), "000")
     tc = rs("abreviatura")
     nc = rs("letra") & " " & Format$(rs("vta_02.sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
     If tipomoneda = rs("vta_02.moneda") Then
        d = Format$(rs("total"), "######0.00")
     Else
        d = Format$(rs("total_otra_moneda"), "######0.00")
     End If
     If tipomoneda = "P" Then
       si2 = Format$(rs("saldo_impago02"), "#####0.00")
     Else
       si2 = Format$(rs("saldo_impago02") / rs("cotizacion_dolar"), "#####0.00")
     End If
     p = rs("cliente02")
     v = rs("denominacion")
     ni = rs("num_int")
    
     If diasmora > 0 And si2 > 0 Then
       interes = Format$((diasmora * Val(t_interes) * si2) / 100, "####0.00")
       msf1.AddItem F & Chr(9) & p & Chr(9) & tc & " " & nc & Chr(9) & d & Chr(9) & si2 & Chr(9) & diasmora & Chr(9) & interes & Chr(9) & v & Chr(9) & rs("num_int")
       reg = reg + 1
       t = t + si2
       ti = ti + interes
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

Sub mensaje()
If Option6 = True Then
  kp = " Vencimiento "
Else
  kp = " Comprobante "
End If
Label9 = "Se considerarán comprobantes EN MORA aquellos que tengan fecha de" & kp & "anterior a: " & Format$(DateValue(t_fechac) - Val(t_dias), "dd/mm/yy")

End Sub
Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 9
msf1.ColWidth(0) = 1100
msf1.ColWidth(1) = 3000
msf1.ColWidth(2) = 2200
msf1.ColWidth(3) = 1200
msf1.ColWidth(4) = 1200
msf1.ColWidth(5) = 900
msf1.ColWidth(6) = 1100
msf1.ColWidth(7) = 2000
msf1.ColWidth(8) = 1100

If Option6 = True Then
  msf1.TextMatrix(0, 0) = "Fecha Vto."
Else
  msf1.TextMatrix(0, 0) = "Fecha Comp."
End If
 
 
If Option4 = True Then
  m = " $ "
Else
  m = "U$s"
End If
 
msf1.TextMatrix(0, 1) = "Cliente"
msf1.TextMatrix(0, 2) = "Tipo y Nro.Comprob."
msf1.TextMatrix(0, 3) = "Total " & m
msf1.TextMatrix(0, 4) = "Deuda " & m
msf1.TextMatrix(0, 5) = "Dias Mora"
msf1.TextMatrix(0, 6) = "Int. " & m
msf1.TextMatrix(0, 7) = "Vendedor"
msf1.TextMatrix(0, 8) = "Num.Int."

For i = 0 To 2
  msf1.ColAlignment(i) = 1 'izq
Next i
For i = 3 To 6
  msf1.ColAlignment(i) = 9 'der
Next i
For i = 7 To 8
  msf1.ColAlignment(i) = 1 'der
Next i


End Sub

Private Sub Form_Load()
Call barraesag(Me)
cal1.Visible = False
Call armagrid
Call carga_clientes(c_prov)
c_prov.AddItem "<Todos>", 0
c_prov.ListIndex = 0
Call carga_SUCURSALES(c_sucursal)
c_sucursal.AddItem "<Todas>", 0
c_sucursal.ListIndex = 0
Call carga_vendedores(c_vend)
c_vend.AddItem "<Todos>", 0
c_vend.ListIndex = 0

Option1 = True
Option4 = True
Option6 = True
Frame7.Visible = True

Set rs = New ADODB.Recordset
q = "select [tasa_financiera] from g0 where [sucursal] = 0"
rs.Open q, cn1
t_interes = Format$(rs("tasa_financiera") / 30, "##0.000")
Set rs = Nothing

t_fechac = Format$(Now, "dd/mm/yy")
Call mensaje

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
    
    For i = 8 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), Space$(60) & "CALCULO INETRESES POR MORA", "     Fecha Calculo: " & t_fechac & "          Dias Tolerancia: " & t_dias & "          Interes diario: " & t_interes, "", "     " & Label9, 45, 9, True, False, "H")
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
    vta_cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 8)
    vta_cc_detalle.Show
  End If
End If

End Sub



Private Sub t_dias_LostFocus()
Call mensaje
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

Private Sub t_dias_GotFocus()
t_dias = ""
End Sub

Private Sub t_fechac_LostFocus()
If t_fechac <> "" Then
  If Not IsDate(t_fechac) Then
    t_fechac = Format$(Now, "dd/mm/yyyy")
  End If
Else
  t_fechac = Format$(Now, "dd/mm/yyyy")
End If
Call mensaje
End Sub

Private Sub t_interes_GotFocus()
t_interes = ""

End Sub
