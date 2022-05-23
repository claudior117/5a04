VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form emp_saldos 
   BackColor       =   &H00E0E0E0&
   Caption         =   "SALDOS EMPLEADOS"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12270
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8700
   ScaleWidth      =   12270
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   360
      TabIndex        =   14
      Top             =   840
      Width           =   6975
      Begin VB.TextBox t_cliente 
         Height          =   285
         Left            =   1440
         TabIndex        =   16
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Empleado:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Orden"
      Height          =   735
      Left            =   3960
      TabIndex        =   10
      Top             =   120
      Width           =   3375
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Apellido"
         Height          =   195
         Left            =   1560
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Id. Legajo"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   7440
      TabIndex        =   7
      Top             =   120
      Width           =   3375
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Muestra Empleados de Baja"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fecha Desde - Hasta"
      Height          =   735
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   3495
      Begin VB.TextBox t_fecha2 
         Height          =   330
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox t_fecha 
         Height          =   330
         Left            =   120
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10080
      TabIndex        =   3
      Top             =   7200
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "emp004.frx":0000
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
         Picture         =   "emp004.frx":0882
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
      Top             =   8445
      Width           =   12270
      _ExtentX        =   21643
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
            TextSave        =   "20/05/2022"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:32 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   4695
      Left            =   360
      TabIndex        =   9
      Top             =   1920
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8281
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   10440
      TabIndex        =   13
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "emp_saldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private EXISTE As String
Private saldoant As Double
Private saldoact As Double
'FIXIT: Declare 'saf' and 'df' and 'hf' and 'sf' and 'sof' and 'd' and 'h' and 'sa' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim saf, df, hf, sf, sof, d, h, sa, s As Double




Private Sub btnacepta_Click()
If verifica Then
  Call carga
End If
End Sub
Function verifica() As Boolean
  verifica = True
  If t_fecha <> "" Then
    If Not IsDate(t_fecha) Then
      verifica = False
    End If
  Else
    verifica = False
  End If
  
  If t_fecha2 <> "" Then
    If Not IsDate(t_fecha2) Then
      verifica = False
    End If
  Else
    verifica = False
  End If
  
  If verifica = False Then
    MsgBox ("Error en las Fechas Ingresadas")
  End If
  
End Function
Sub carga()
Dim r As Integer
Call armagrid

Load espere

pb = 1
Set rs1 = New ADODB.Recordset
QUERY = "select * from emp_01"
X = " where "

If t_cliente <> "" Then
  QUERY = QUERY & X & " [denominacion]  like '%" & t_cliente & "%'"
  X = " and "
End If

If Check1 = 0 Then
  QUERY = QUERY & X & " [estado] = 'A'"
  X = " and "
End If

If Option1 = True Then
  QUERY = QUERY & " order by [id_legajo]"
Else
  QUERY = QUERY & " order by [denominacion]"
End If
  
rs1.Open QUERY, cn1, adOpenStatic, adLockOptimistic, 1
If Not rs1.EOF And Not rs1.BOF Then
  espere!ProgressBar1.Max = rs1.RecordCount + 1
  espere!ProgressBar1.Min = 1
  espere.Show
  espere.Refresh
  saf = 0
  df = 0
  hf = 0
  sf = 0
  sof = 0
  r = 0
  While Not rs1.EOF
   
   espere!ProgressBar1 = pb
    
   Call sacasaldos(rs1("id_legajo"))
    saf = saf + sa
    df = df + d
    hf = hf + h
    msf1.AddItem rs1("id_legajo") & Chr$(9) & rs1("denominacion") & Chr$(9) & Format$(sa, "#####0.00") & Chr$(9) & Format$(d, "#####0.00") & Chr$(9) & Format$(h, "#####0.00") & Chr$(9) & Format$(s, "######0.00")
    rs1.MoveNext
    pb = pb + 1
    Label5 = pb
    Label5.Refresh
    r = r + 1
  Wend
  sf = saf + df - hf
  msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & "________________" & Chr$(9) & "________________" & Chr$(9) & "________________" & Chr$(9) & "________________"
  msf1.AddItem "" & Chr$(9) & "Total Empleados: " & r & Chr$(9) & Format$(saf, "#####0.00") & Chr$(9) & Format$(df, "######0.00") & Chr$(9) & Format$(hf, "######0.00") & Chr$(9) & Format$(sf, "######0.00")
  
  Unload espere
End If
Set rs1 = Nothing
Set rs2 = Nothing


End Sub
Sub sacasaldos(ByVal l As Long)
q = "select * from emp_02 where [id_legajo] = " & l & " and datevalue([fecha]) < datevalue('" & t_fecha & "')"
Set rss = New ADODB.Recordset
rss.Open q, cn1
sa = 0
While Not rss.EOF
 If rss("tipo_movimiento") = 1 Then
   sa = sa + rss("importe")
 Else
   sa = sa - rss("importe")
 End If
 rss.MoveNext
Wend
Set rss = Nothing


q = "select * from emp_02 where [id_legajo] = " & l & " and datevalue([fecha]) >= datevalue('" & t_fecha & "')" & " and datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
Set rss = New ADODB.Recordset
rss.Open q, cn1
d = 0
h = 0
s = 0
While Not rss.EOF
 If rss("tipo_movimiento") = 1 Then
   d = d + rss("importe")
 Else
   h = h + rss("importe")
 End If
 rss.MoveNext
Wend
Set rss = Nothing

s = sa + d - h

   
End Sub


Private Sub btnsale_Click()

Unload Me
End Sub






Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     gen_tools.Show
End Select

End Sub
Sub armagrid()
'armar grilla
  msf1.clear
  msf1.Rows = 1
  msf1.Cols = 7
  msf1.ColWidth(0) = 600
  msf1.ColWidth(1) = 4000
  msf1.ColWidth(2) = 1400
  msf1.ColWidth(3) = 1400
  msf1.ColWidth(4) = 1400
  msf1.ColWidth(5) = 1400
  msf1.ColWidth(6) = 500
  msf1.TextMatrix(0, 0) = "Legajo"
  msf1.TextMatrix(0, 1) = "Empleado"
  msf1.TextMatrix(0, 2) = "Saldo Ant."
  msf1.TextMatrix(0, 3) = "Debe($)"
  msf1.TextMatrix(0, 4) = "Haber($)"
  msf1.TextMatrix(0, 5) = "Saldo($)"
  msf1.TextMatrix(0, 6) = " "
  For i = 0 To 6
    msf1.ColAlignment(i) = 9 'der
  Next i
  msf1.ColAlignment(1) = 1 'izq
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 2)
  
End Select


End Sub

Private Sub Form_Load()
Load vta_estadocuenta
Call barra(Me)

Check1 = 0
Option1 = True
Call armagrid
End Sub




Private Sub Form_Unload(Cancel As Integer)
Unload vta_estadocuenta
End Sub






Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F7] Imprime - [F11] Exporta Excel  "
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
      c(0) = 6
      c(1) = 0
      c(2) = 1
      c(3) = 2
      c(4) = 3
      c(5) = 4
      c(6) = 5
      
      p = "Periodo: " & t_fecha & "  " & t_fecha2
      For i = 7 To 14
        c(i) = -1
      Next i
      Call imprimegrid(msf1, c(), "SALDOS por EMPLEADOS", p, " ", v, 72, 8, True, False)
  End If

End If

If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If

End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And msf1.Rows > 1 Then
   c = Val(msf1.TextMatrix(msf1.Row, 0))
   If c > 0 Then
      Load emp_estadocuenta
      emp_estadocuenta.c_prov.ListIndex = buscaindice(emp_estadocuenta.c_prov, c)
      emp_estadocuenta.Show
   End If
End If
End Sub


Private Sub t_cliente_GotFocus()
t_cliente = ""
End Sub

Private Sub t_fecha_DblClick()
cal1.Visible = True
cal1.Tag = 1
cal1.SetFocus
End Sub

Private Sub t_fecha_GotFocus()
t_fecha = ""
End Sub

Private Sub t_fecha2_DblClick()
cal1.Visible = True
cal1.Tag = 2
cal1.SetFocus

End Sub

Private Sub t_fecha2_GotFocus()
t_fecha2 = ""
End Sub

'FIXIT: t_fecha2_LinkOpen event no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
Private Sub t_fecha2_LinkOpen(Cancel As Integer)
Call solofecha(t_fecha2)
End Sub
