VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form vta_informevta 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INFORME DE VENTAS"
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
      Caption         =   "Tipo"
      Height          =   495
      Left            =   8640
      TabIndex        =   27
      Top             =   1200
      Width           =   2775
      Begin VB.OptionButton Option7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Detallado"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Solo totales"
         Height          =   195
         Left            =   1320
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   4560
      TabIndex        =   26
      Top             =   7200
      Width           =   4935
      Begin VB.TextBox t_t2 
         Height          =   285
         Left            =   3120
         TabIndex        =   32
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox t_t1 
         Height          =   285
         Left            =   1320
         TabIndex        =   31
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Totales entre:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Moneda"
      Height          =   495
      Left            =   5520
      TabIndex        =   23
      Top             =   1200
      Width           =   2775
      Begin VB.OptionButton Option5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "U$s Dolares"
         Height          =   195
         Left            =   1200
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "$ Pesos"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   240
      TabIndex        =   20
      Top             =   1080
      Width           =   3255
      Begin VB.ComboBox c_sucursal 
         Height          =   315
         Left            =   1680
         TabIndex        =   21
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
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   0
      TabIndex        =   16
      Top             =   7200
      Width           =   4335
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Agrupados por Vendedor"
         Height          =   315
         Left            =   2880
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Agrupados por Clientes"
         Height          =   315
         Left            =   1440
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sin Agrupar"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1215
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
         BackColor       =   &H00C00000&
         Caption         =   "Vendedor:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
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
      StartOfWeek     =   113704961
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
         Picture         =   "VTA019.frx":0000
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
         Picture         =   "VTA019.frx":0882
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
            TextSave        =   "24/07/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "05:12 p.m."
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
End
Attribute VB_Name = "vta_informevta"
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
  q = "select * from vta_02,  vta_06, vta_05, vta_01 where vta_02.[id_tipocomp] = vta_06.[id_tipocomp] and vta_02.[id_cliente] = vta_01.[id_cliente]  and vta_02.[id_vendedor] = vta_05.[id_vendedor] and vta_02.[venta] <> 'N' and vta_02.[sucursal] = vta_06.[sucursal]"
  c = " and "
  If c_prov.ListIndex > 0 Then
     q = q & c & " vta_02.[id_cliente] = " & c_prov.ItemData(c_prov.ListIndex)
  End If
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  
  If c_vend.ListIndex > 0 Then
     q = q & c & " vta_02.[Id_vendedor] = " & c_vend.ItemData(c_vend.ListIndex)
  End If
  
  If c_sucursal.ListIndex > 0 Then
     q = q & c & " [sucursal_ingreso] = " & Val(c_sucursal)
  End If
  
  
  q = q & " order by [fecha], vta_02.[id_tipocomp], [num_comp]"
  
  Set rs = New adodb.Recordset
  rs.Open q, cn1
  t = 0
  reg = 0
  
  While Not rs.EOF
     Call renglon2
     rs.MoveNext
  Wend
  msf1.AddItem ""
  msf1.AddItem "" & Chr(9) & "" & Chr(9) & "Comprobantes: " & reg & Chr(9) & "Totales:" & Chr(9) & Format$(t, "#####0.00") & Chr(9) & ""

     
  
  
   
   
   
End Sub
Private Sub btnacepta_Click()
 espere.Show
 espere.Refresh
  QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
  QUERY = QUERY & " VALUES ('Informe de Ventas " & "', " & para.id_usuario & ", 'V', 0, '" & Now & "', ' ', 15, " & 0 & ")"
  cn1.BeginTrans
  cn1.Execute QUERY
  cn1.CommitTrans

If Option4 = True Then
    tipomoneda = "P"
  Else
   tipomoneda = "D"
End If
If Option1 = True Then
   Call carga
Else
  If Option2 = True Then
    Call carga2
  Else
    Call carga3
  End If
End If
 Unload espere
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub


Sub carga2()
  Call armagrid
  q = "select * from vta_02,  vta_06, vta_05, vta_01 where vta_02.[id_tipocomp] = vta_06.[id_tipocomp] and vta_02.[id_cliente] = vta_01.[id_cliente]  and vta_02.[id_vendedor] = vta_05.[id_vendedor] and vta_02.[venta] <> 'N' and vta_02.[sucursal] = vta_06.[sucursal]"
  c = " and "
  If c_prov.ListIndex > 0 Then
     q = q & c & " vta_02.[id_cliente] = " & c_prov.ItemData(c_prov.ListIndex)
  End If
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  
  If c_vend.ListIndex > 0 Then
     q = q & c & " vta_02.[Id_vendedor] = " & c_vend.ItemData(c_vend.ListIndex)
  End If
  
   If c_sucursal.ListIndex > 0 Then
     q = q & c & " [sucursal_ingreso] = " & Val(c_sucursal)
  End If
 
  q = q & " order by vta_02.[id_cliente], [fecha], vta_02.[id_tipocomp], [num_comp]"
  
  Set rs = New adodb.Recordset
  rs.Open q, cn1
  t = 0
  reg = 0
  ti = 0
  regi = 0
  a = 0
  T2 = 0
  While Not rs.EOF
   If a = 0 Then
     cc = rs("vta_02.id_cliente")
     ti = 0
     regi = 0
     a = 1
   End If
   
   If cc = rs("vta_02.id_cliente") Then
       Call renglon2
   Else
       If Option7 = True Then
        msf1.AddItem Chr(9) & Chr(9) & Chr(9) & Chr(9) & "_____________________"
       End If
      If verificasiimprime Then
        msf1.AddItem "" & Chr(9) & "" & Chr(9) & p & Chr(9) & "Totales:" & Chr(9) & Format$(ti, "#####0.00") & Chr(9) & "Comprobantes: " & regi
        msf1.AddItem ""
        T2 = T2 + Val(ti)
      End If
       ti = 0
       regi = 0
       Call renglon2
       cc = rs("vta_02.id_cliente")
   End If
   rs.MoveNext
  Wend
 If Option7 = True Then
     msf1.AddItem Chr(9) & Chr(9) & Chr(9) & Chr(9) & "_____________________"
  End If
  If verificasiimprime Then
    msf1.AddItem "" & Chr(9) & "" & Chr(9) & p & Chr(9) & "Totales:" & Chr(9) & Format$(ti, "#####0.00") & Chr(9) & "Comprobantes: " & regi
    msf1.AddItem ""
     T2 = T2 + Val(ti)
  End If
  msf1.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "Totales:" & Chr(9) & Format$(T2, "#####0.00") & Chr(9) & ""
End Sub
Function verificasiimprime() As Boolean
If Option6 = True Then
 m1 = 0
 m2 = 0
 If Val(t_t1) > 0 Or Val(t_t2) > 0 Then
  If Val(t_t1) > 0 Then
    If Val(ti) >= Val(t_t1) Then
      m1 = 1
    End If
  Else
    m1 = 1
  End If
  
  If Val(t_t2) > 0 Then
    If Val(ti) <= Val(t_t2) Then
      m2 = 1
    End If
  Else
    m2 = 1
  End If
  
 Else
  m1 = 1
  m2 = 1
 End If
 If m1 = 1 And m2 = 1 Then
   verificasiimprime = True
 Else
   verificasiimprime = False
 End If
Else
 verificasiimprime = True
End If
End Function
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
    If Option7 = True Then
     msf1.AddItem F & Chr(9) & cp & Chr(9) & p & Chr(9) & tc & " " & nc & Chr(9) & d & Chr(9) & v & Chr(9) & rs("num_int")
    End If
    reg = reg + 1
    regi = regi + 1
End Sub

Sub carga3()
  Call armagrid
  q = "select * from vta_02,  vta_06, vta_05, vta_01 where vta_02.[id_tipocomp] = vta_06.[id_tipocomp] and vta_02.[id_cliente] = vta_01.[id_cliente]  and vta_02.[id_vendedor] = vta_05.[id_vendedor] and vta_02.[venta] <> 'N' and vta_02.[sucursal] = vta_06.[sucursal]"
  c = " and "
  If c_prov.ListIndex > 0 Then
     q = q & c & " vta_02.[id_cliente] = " & c_prov.ItemData(c_prov.ListIndex)
  End If
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  
  If c_vend.ListIndex > 0 Then
     q = q & c & " vta_02.[Id_vendedor] = " & c_vend.ItemData(c_vend.ListIndex)
  End If
  
   If c_sucursal.ListIndex > 0 Then
     q = q & c & " [sucursal_ingreso] = " & Val(c_sucursal)
  End If
 
  q = q & " order by vta_02.[id_vendedor], vta_02.[id_cliente], [fecha], vta_02.[id_tipocomp], [num_comp]"

  
  Set rs = New adodb.Recordset
  rs.Open q, cn1
  t = 0
  reg = 0
  ti = 0
  regi = 0
  p = 0
  T2 = 0
  While Not rs.EOF
   If p = 0 Then
     cc = rs("vta_02.id_vendedor")
     ti = 0
     regi = 0
     p = 1
   End If
   
   If cc = rs("vta_02.id_vendedor") Then
       Call renglon2
   Else
       If Option7 = True Then
         msf1.AddItem Chr(9) & Chr(9) & Chr(9) & Chr(9) & "_____________________"
       End If
       If verificasiimprime Then
        msf1.AddItem "" & Chr(9) & "" & Chr(9) & v & Chr(9) & "Totales:" & Chr(9) & Format$(ti, "#####0.00") & Chr(9) & "Comprobantes: " & regi
        msf1.AddItem ""
        T2 = T2 + Val(ti)
       End If
       ti = 0
       regi = 0
       Call renglon2
       cc = rs("vta_02.id_vendedor")
   End If
   rs.MoveNext
  Wend
  If Option7 = True Then
         msf1.AddItem Chr(9) & Chr(9) & Chr(9) & Chr(9) & "_____________________"
  End If
  If verificasiimprime Then
        msf1.AddItem "" & Chr(9) & "" & Chr(9) & v & Chr(9) & "Totales:" & Chr(9) & Format$(ti, "#####0.00") & Chr(9) & "Comprobantes: " & regi
        msf1.AddItem ""
        T2 = T2 + Val(ti)
  End If
  msf1.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "Totales:" & Chr(9) & Format$(T2, "#####0.00") & Chr(9) & ""



      
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

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 7
msf1.ColWidth(0) = 1100
msf1.ColWidth(1) = 500
msf1.ColWidth(2) = 3000
msf1.ColWidth(3) = 2200
msf1.ColWidth(4) = 1100
msf1.ColWidth(5) = 2500
msf1.ColWidth(6) = 1100
msf1.TextMatrix(0, 0) = "Fecha"
msf1.TextMatrix(0, 1) = "Id."
msf1.TextMatrix(0, 2) = "Cliente"
msf1.TextMatrix(0, 3) = "Tipo y Nro.Comprob."
msf1.TextMatrix(0, 4) = "Total"
msf1.TextMatrix(0, 5) = "Vendedor"
msf1.TextMatrix(0, 6) = "Num.Int."

For i = 0 To 3
  msf1.ColAlignment(i) = 1 'izq
Next i
For i = 4 To 6
  msf1.ColAlignment(i) = 9 'der
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
Option7 = True
Frame7.Visible = True
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
    For i = 7 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "INFORME DE VENTAS", "", "Periodo: " & t_fecha & " : " & t_fecha2, "", 90, 7, True, False)
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



Private Sub Option6_Click()
Frame7.Visible = True
End Sub

Private Sub Option7_Click()
Frame7.Visible = False
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

Private Sub t_t1_GotFocus()
t_t1 = ""
End Sub

Private Sub t_t2_GotFocus()
t_t2 = ""

End Sub
