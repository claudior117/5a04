VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cja_informecierre 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INFORME CIERRE CAJA"
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
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   1335
      Left            =   120
      TabIndex        =   18
      Top             =   6960
      Width           =   9855
      Begin VB.Label Label9 
         BackColor       =   &H0080FFFF&
         Caption         =   "(4) Verificar las Ventas Totales en el Z menos las NC debe concidir con este importe."
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   960
         Width           =   9135
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080FFFF&
         Caption         =   "(3) Ventas en Cuenta Corriente. No juegan en la caja es solo informativo."
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   9135
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080FFFF&
         Caption         =   "(2) Ventas de Contado, tiques, facturas, nc. "
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   9135
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080FFFF&
         Caption         =   "(1) Total de Recibos. Para verificar puede emitir en Ventas ""Cierre diario"" o ""Ver comprobantes"""
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   9135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos de Caja del la Jornada"
      Height          =   1695
      Left            =   4200
      TabIndex        =   13
      Top             =   120
      Width           =   7695
      Begin VB.TextBox t_dia 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox t_cierre 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5400
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox t_monedas 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3480
         MaxLength       =   10
         TabIndex        =   3
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox t_billetes 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3480
         MaxLength       =   10
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox t_inicio 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3480
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Total en efectivo al cerrar Jornada:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF0000&
         Caption         =   "Total en Monedas contadas al cerrar Jornada"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF0000&
         Caption         =   "Total en Billetes contados al cerrar Jornada"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF0000&
         Caption         =   "Total Efectivo en Caja  al Inicio Jornada:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   3615
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Cierre:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10320
      TabIndex        =   8
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "cja_006.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "cja_006.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   9
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
      TabIndex        =   7
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
            TextSave        =   "27/02/2015"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:39"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   4935
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   8705
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
Attribute VB_Name = "cja_informecierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
'FIXIT: Declare 'ti' con un tipo de datos de enlace en tiempo de compilación               FixIT90210ae-R1672-R1B8ZE
Dim ti, t As Double
'FIXIT: Declare 'reg' con un tipo de datos de enlace en tiempo de compilación              FixIT90210ae-R1672-R1B8ZE
Dim totalingresos As Double
Dim totalegresos As Double
Dim saldoinicio As Double
Dim saldocierre As Double


Sub carga()
  
  Call armagrid
  'consulta ventas cuenta corriente
  
   espere.Show
   espere.Label1 = "Espere....  Obteniendo Ingresos"
   espere.Refresh
   
   msf1.AddItem "COMPOSICION DE SALDO ACTUAL DE CAJA"
   msf1.AddItem "Saldo en efectivo al iniciar jornada" & Chr$(9) & t_inicio
   msf1.AddItem "Total en Billetes al Cerrar Jornada" & Chr$(9) & t_billetes
   msf1.AddItem "Total en Monedas al Cerrar Jornada" & Chr$(9) & t_monedas
   msf1.AddItem "" & Chr$(9) & "--------------------------------"
   msf1.AddItem "Totales en Efectivo al cerrar Jornada" & Chr$(9) & t_dia & Chr$(9) & t_cierre
   msf1.AddItem ""
   msf1.AddItem ""
   msf1.AddItem "INGRESOS DEL DIA REGISTRADOS" & Chr$(9) & "Cta.Cte (1)" & Chr$(9) & "Contado (2)" & Chr$(9) & "Ing. Tot." & Chr$(9) & "Ventas C.C.(3)" & Chr$(9) & "Tot.Z-NC (4)"
   
   saldoinicio = Val(t_incio)
   
   q = "select * from cyb_05, vta_02 where datevalue(cyb_05.[fecha]) = datevalue('" & t_fecha & "') and   [modulo] = 'V' and [num_mov_int] = [num_int] order by [sucursal_ingreso]"
   Set rs = New ADODB.Recordset
   rs.Open q, cn1
   suc = 1
   tsc = 0
   tscc = 0
   ttsc = 0
   ttscc = 0
   cuentac = 0
   totc = 0
   While Not rs.EOF
    If suc <> rs("sucursal_ingreso") Then
       cc = sacactacte(suc)
      'muestro totales e inicializo contadores
       msf1.AddItem "Punto de Venta Nro. " & Format$(suc, "0000") & Chr$(9) & Format$(tscc, "#######0.00") & Chr$(9) & Format$(tsc, "#######0.00") & Chr$(9) & Format$(tscc + tsc, "#######0.00") & Chr$(9) & Format$(cc, "#######0.00") & Chr$(9) & Format$(cc + tsc, "#######0.00")
       ttscc = ttscc + tscc
       ttsc = ttsc + tsc
       suc = rs("sucursal_ingreso")
       tsc = 0
       tscc = 0
    End If
    If rs("contado") = "S" Then
         tsc = tsc + rs("importe")
    Else
         tscc = tscc + rs("importe")
    End If
    
    rs.MoveNext
   Wend
   
   Set rs = Nothing
   ttscc = ttscc + tscc
   ttsc = ttsc + tsc
   msf1.AddItem "Punto de Venta Nro. " & Format$(suc, "0000") & Chr$(9) & Format$(tscc, "#######0.00") & Chr$(9) & Format$(tsc, "#######0.00") & Chr$(9) & Format$(tscc + tsc, "#######0.00")
  
   q = "select * from cyb_05 where  [modulo] <> 'V' and [ubicacion] = 'D' and datevalue([fecha])= datevalue('" & t_fecha & "')"
   Set rs = New ADODB.Recordset
   rs.Open q, cn1
   toi = 0
   While Not rs.EOF
      toi = toi + rs("importe")
      rs.MoveNext
   Wend
   Set rs = Nothing
   msf1.AddItem "Otros Ingresos  " & Chr$(9) & "" & Chr$(9) & Format$(toi, "#######0.00") & Chr$(9) & Format$(toi, "#######0.00")
   msf1.AddItem "" & Chr$(9) & "------------------------------------" & Chr$(9) & "------------------------------------" & Chr$(9) & "------------------------------------"
   msf1.AddItem "Ingresos tot. registrados en el dia " & Chr$(9) & Format$(ttscc, "#######0.00") & Chr$(9) & Format$(ttsc + toi, "#######0.00") & Chr$(9) & Format$(ttscc + ttsc + toi, "#######0.00")
   msf1.AddItem ""
   msf1.AddItem "Ingresos Tot. Registrados en el dia" & Chr$(9) & Format$(ttscc + ttsc + toi, "#######0.00")
   
   totalingresos = ttscc + ttsc + toi
   
   msf1.AddItem ""
   msf1.AddItem ""
   msf1.AddItem "EGRESOS del DIA"
   Call sacaegresos
   
   msf1.AddItem ""
   msf1.AddItem "SALDO DE CIERRE" & Chr$(9) & Format$(saldoinicio + totalingresos - totalegresos, "######0.00")
   msf1.AddItem "DINERO EN CAJA" & Chr$(9) & Format$(Val(t_cierre), "######0.00")
   
   diferencia = Format$(saldoinicio + totalingresos - totalegresos - Val(t_cierre), "######0.00")
   msf1.AddItem "" & Chr$(9) & "---------------------------------"
   Select Case Val(diferencia)
    Case Is = 0
      
      msf1.AddItem "¡¡¡¡ CAJA EXACTA !!!!!" & Chr$(9) & diferencia
    Case Is > 0
      msf1.AddItem "¡¡¡¡ FALTANTE DE CAJA  !!!!!" & Chr$(9) & diferencia
    Case Is < 0
      msf1.AddItem "¡¡¡¡ SOBRANTE DE CAJA  !!!!!" & Chr$(9) & -diferencia
    End Select
   
   
   
   Unload espere
   
End Sub
Sub sacaegresos()
   q = "select * from cyb_05 where  [ubicacion] <> 'D' and datevalue([fecha])= datevalue('" & t_fecha & "')"
   Set rs = New ADODB.Recordset
   rs.Open q, cn1
   te = 0
   While Not rs.EOF
      te = te + rs("importe")
      msf1.AddItem rs("descripcion") & Chr$(9) & Format$(rs("importe"), "######0.00")
      rs.MoveNext
   Wend
   Set rs = Nothing
   msf1.AddItem "" & Chr$(9) & "---------------------------------------"
   msf1.AddItem "Total de egresos del dia" & Chr$(9) & Format$(te, "#######0.00")
   totalegresos = te
End Sub
Private Sub btnacepta_Click()
   Call carga
  
End Sub
Function sacactacte(ByVal s As Integer) As Double
  'saca las ventas en ctacte para la sucursal
  q = "select * from vta_02 where [sucursal_ingreso] = " & s & " and datevalue(fecha) = datevalue('" & t_fecha & "') and [contado] = 'N' and [cta_cte] <> 'N' and [venta] <> 'N'"
  Set rs1 = New ADODB.Recordset
  rs1.Open q, cn1
  tot = 0
  While Not rs1.EOF
    If rs1("cta_cte") <> "H" Then
      t = rs1("total")
    Else
      t = -rs1("total")
    End If
    'MsgBox (t)
    tot = tot + t
    rs1.MoveNext
  Wend
  Set rs1 = Nothing
  sacactacte = tot
End Function
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








Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub

Sub armagrid()

msf1.clear
msf1.Rows = 1
msf1.Cols = 6

msf1.ColWidth(0) = 4300
msf1.ColWidth(1) = 1300
msf1.ColWidth(2) = 1300
msf1.ColWidth(3) = 1300
msf1.ColWidth(4) = 1300
msf1.ColWidth(5) = 1300


For J = 1 To 3
    msf1.ColAlignment(J) = 9 'der
Next J
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 6)
  
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Call armagrid

End Sub




Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F7] Imprime - [F11] Excel"

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
    
    For i = 6 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "CIERRE de CAJA", "", "Fecha: " & t_fecha, " ", 60, 8, True, False)
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



Private Sub t_billetes_LostFocus()
Call cuentaplata
End Sub

Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
    t_fecha = Format$(Now, "dd/mm/yyyy")
  End If
Else
  t_fecha = Format$(Now, "dd/mm/yyyy")
End If
End Sub

Sub cuentaplata()
  t_inicio = Format$(Val(t_inicio), "#######0.00")
  t_billetes = Format$(Val(t_billetes), "#######0.00")
  t_monedas = Format$(Val(t_monedas), "#######0.00")
    t_cierre = Format$(Val(t_inicio) + Val(t_billetes) + Val(t_monedas), "######0.00")
  t_dia = Format$(Val(t_billetes) + Val(t_monedas), "######0.00")

End Sub
Private Sub t_inicio_LostFocus()
Call cuentaplata
End Sub

Private Sub t_monedas_LostFocus()
Call cuentaplata
End Sub
