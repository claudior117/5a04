VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form gen_asistencia 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PROGRAMA EMERGENCIA AL TRABAJO y PRODUCCIÓN"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12060
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   12060
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Salida:"
      Height          =   615
      Left            =   0
      TabIndex        =   18
      Top             =   6840
      Width           =   9975
      Begin VB.CommandButton Command2 
         Caption         =   "Carpeta destino:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox t_carpeta 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   8055
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   0
      TabIndex        =   16
      Top             =   7440
      Width           =   9975
      Begin VB.Label Label4 
         Caption         =   "Se genera un archivo Excel"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   9375
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opciones"
      Height          =   615
      Left            =   7320
      TabIndex        =   14
      Top             =   720
      Width           =   3255
      Begin VB.CommandButton Command1 
         Caption         =   "Verifica Totales"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   7320
      TabIndex        =   11
      Top             =   120
      Width           =   3255
      Begin VB.ComboBox c_sucursal 
         Height          =   315
         Left            =   1680
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
      Left            =   1920
      TabIndex        =   9
      Top             =   120
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   216727553
      CurrentDate     =   38750
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
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
         Top             =   720
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
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
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
      Left            =   10200
      TabIndex        =   3
      Top             =   7440
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "gen039..frx":0000
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
         Picture         =   "gen039..frx":0882
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
      Top             =   8550
      Width           =   12060
      _ExtentX        =   21273
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
            TextSave        =   "14/04/2020"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "21:09"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5415
      Left            =   0
      TabIndex        =   10
      Top             =   1440
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9551
      _Version        =   393216
   End
End
Attribute VB_Name = "gen_asistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim c5 As Double

Sub carga()
 'Dim cr(2) As Long
 Dim nic As String
 espere.Show
 espere.Label1 = "Espere...... Generando Listado para CITI"
 espere.Refresh
 Call armagrid
  q = "select * from VTA_02, vta_01, vta_06, g3 where [grabado] <> 'N' and  vta_02.[id_tipocomp] = vta_06.[id_tipocomp] and vta_02.[id_cliente] = vta_01.[id_cliente] and vta_01.[id_tipoiva] = [cod_tipoiva] and vta_02.[sucursal_ingreso] = vta_06.[sucursal] "
  c = " and "
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  If c_sucursal.ListIndex > 0 Then
    q = q & c & " and [sucursal_ingreso] = " & Val(c_sucursal)
  End If
  
 q = q & c & "  (vta_02.[id_tipocomp] < 33 or vta_02.[id_tipocomp] > 200)"
  q = q & " order by [fecha], [letra], [num_comp]"
  Set rs = New adodb.Recordset
  
  rs.Open q, cn1
  tt = 0
  ti = 0
  ts = 0
  tng = 0
  trp = 0
  While Not rs.EOF
     er = ""
     obserr = ""
     F = Mid$(Format$(rs("fecha"), "dd/mm/yyyy"), 7, 4) & Mid$(Format$(rs("fecha"), "dd/mm/yyyy"), 4, 2) & Mid$(Format$(rs("fecha"), "dd/mm/yyyy"), 1, 2)
     Select Case rs("letra")
     Case Is = "A"
          tc = Format$(rs("cod_afip_a"), "000")
         
          
     Case Is = "B"
          tc = Format$(rs("cod_afip_b"), "000")
         
          
     Case Is = "E"
          tc = Format$(rs("cod_afip_e"), "000")
          
     Case Else
          tc = Format$(rs("cod_afip_a"), "000")
          
     End Select
          
     
     PtV = Format$(rs("vta_02.sucursal"), "00000")
     nc = Format$(rs("num_comp"), "00000000")
     If Val(PtV) <= 0 Or Val(nc) <= 0 Then
         er = er & "*"
         obserr = obserr & "PV/Nro.- "
     End If
     
     If rs("vta_02.moneda") = "P" Then
       c5 = 1
       moneda = "PES"
     Else
       c5 = rs("cotizacion_dolar")
       moneda = "DOL"
     End If
         cambio = Format$(c5, "###0.0000")
         pin = (rs("perc_iva") + rs("perc_gan")) * c5
         pip = rs("perc_ib") * c5
         impuestos = pin + pip + rs("VTA_02.iva") + rs("impuestos")
         t = Format$(rs("total") * c5, "######0.00")
         i = Format$(impuestos, "######0.00")
         s = Format$(rs("subtotal"), "######0.00")
         msf1.AddItem F & Chr(9) & tc & Chr(9) & PtV & Chr$(9) & nc & Chr$(9) & s & Chr(9) & i & Chr(9) & t
    rs.MoveNext
  Wend
  Unload espere
     
End Sub
Private Sub btnacepta_Click()
  Call carga
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



Private Sub Command1_Click()
If t_fecha = "" Or t_fecha2 = "" Then
  MsgBox ("Debe indicar un periodo de trabajo para realizar esta operacion")
  Exit Sub
End If
h = MsgBox("Verificacion de Totales. Asegurese de haber indicado correctamente el periodo de trabajo y No apague la maquina ni cancele este proceso. ¿Esta seguro que quiere actualizar? ", 4)
If h = 6 Then
espere.Show
espere.Refresh
qm = "select * from vta_02 where  [grabado] <> 'N'"
c = " and "
If IsDate(t_fecha) Then
    qm = qm & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
End If
  
If IsDate(t_fecha2) Then
   qm = qm & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
End If
  
If c_sucursal.ListIndex > 0 Then
    qm = qm & c & " and [sucursal_ingreso] = " & Val(c_sucursal)
End If
Set rs2 = New adodb.Recordset
rs2.Open qm, cn1
a = 1
While Not rs2.EOF
 Call verifica_tasa_iva(rs2("num_int"))
 
 rs2.MoveNext
Wend
Set rs2 = Nothing
Unload espere
MsgBox ("Proceso Terminado")
End If

End Sub

Private Sub Command2_Click()
Load gen_seleccionacarpeta
gen_seleccionacarpeta.t_llamada = "3"
gen_seleccionacarpeta.Show

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
msf1.ColWidth(0) = 2000
msf1.ColWidth(1) = 2000
msf1.ColWidth(2) = 2000
msf1.ColWidth(3) = 2000
msf1.ColWidth(4) = 2000
msf1.ColWidth(5) = 2000
msf1.ColWidth(6) = 2000



msf1.TextMatrix(0, 0) = "Fecha Comprobante"
msf1.TextMatrix(0, 1) = "Tipo Comprobante"
msf1.TextMatrix(0, 2) = "Punto Venta "
msf1.TextMatrix(0, 3) = "Numero Comprobante"
msf1.TextMatrix(0, 4) = "Importe Neto"
msf1.TextMatrix(0, 5) = "Impuestos"
msf1.TextMatrix(0, 6) = "Importe Total"

For i = 0 To 6
  msf1.ColAlignment(i) = 9
Next i


End Sub

Private Sub Form_Load()
Call carga_SUCURSALES(c_sucursal)
c_sucursal.AddItem "<Todas>", 0
c_sucursal.ListIndex = 0

Call barraesag(Me)
cal1.Visible = False
Call armagrid


t_carpeta = "c:\"
End Sub



Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F5] Exporta -  [F7] Imprime "

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
    Call imprimegrid(msf1, c(), "LISTADO DE IVA VENTAS", "", "Periodo: " & t_fecha & " : " & t_fecha2, "", 95, 6, True, False)
  End If

End If


If KeyCode = vbKeyF5 Then
 
  Call exportaexcel(msf1)
End If






End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    Load cc_detalle
    vta_cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 14)
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
