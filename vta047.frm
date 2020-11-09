VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_arba_corralones 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AICYC Informa ARBA para Empresas constructoras y Corralones"
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
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   16
      Top             =   7200
      Width           =   9375
      Begin VB.Label Label4 
         Caption         =   $"vta047.frx":0000
         Height          =   855
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   8895
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opciones"
      Height          =   855
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
         Width           =   2895
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
      Left            =   4800
      TabIndex        =   9
      Top             =   120
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   277610497
      CurrentDate     =   38750
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   240
      TabIndex        =   6
      Top             =   0
      Width           =   3615
      Begin VB.TextBox t_importe 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   18
         Top             =   1080
         Width           =   1335
      End
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
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "Importe minimo a informar"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
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
         Picture         =   "vta047.frx":00D0
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
         Picture         =   "vta047.frx":0952
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
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   17639
            MinWidth        =   17639
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "09/10/2014"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:09"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5415
      Left            =   0
      TabIndex        =   10
      Top             =   1680
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9551
      _Version        =   393216
   End
End
Attribute VB_Name = "vta_arba_corralones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim c5 As Double

Sub carga()
 'Dim cr(2) As Long
 espere.Show
 espere.Label1 = "Espere...... Generando Listado para AICYC"
 espere.Refresh
 Call armagrid
  q = "select * from VTA_02, vta_01, vta_06 where vta_02.[id_tipocomp] = 1 and  vta_02.[id_tipocomp] = vta_06.[id_tipocomp] and vta_02.[id_cliente] = vta_01.[id_cliente] and vta_02.[sucursal_ingreso] = vta_06.[sucursal]"
  c = " and "
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  If c_sucursal.ListIndex > 0 Then
    q = q & c & "  [sucursal_ingreso] = " & Val(c_sucursal)
  End If
  
  If Val(t_importe) > 0 Then
    q = q & c & " [total] >= " & Val(t_importe)
  End If
  
  'MsgBox (q)
  
  q = q & " order by [fecha], [letra], [num_comp]"
  Set rs = New ADODB.Recordset
  
  rs.Open q, cn1
  tt = 0
  ti = 0
  ts = 0
  tng = 0
  trp = 0
  While Not rs.EOF
     er = ""
     obserr = ""
     F = Format$(rs("fecha"), "dd/mm/yyyy")
     PtV = Format$(rs("vta_02.sucursal"), "0000")
     nc = Format$(rs("num_comp"), "00000000")
     letra = rs("letra")
     If rs("vta_02.moneda") = "P" Then
       c5 = 1
     Else
       c5 = rs("cotizacion_dolar")
     End If
     t = Format$(rs("total") * c5, "00000000.0")
     
     If rs("id_tipoiva") = 3 Then 'cf lleva dni el resto lleva cuit
       td = 1
     Else
       td = 7
     End If
     l = "Rojas"
     If rs("cuit") < 100000 Then
       cu = 1999999
     Else
       cu = rs("cuit")
     End If
     
     If Len(rs("direccion_local")) < 6 Then
       dl = "Cuartel 1 - Rojas"
     Else
       dl = rs("direccion_local")
    End If
      msf1.AddItem er & Chr$(9) & F & Chr(9) & rs("cliente02") & Chr(9) & td & Chr$(9) & cu & Chr$(9) & letra & Chr$(9) & PtV & Chr$(9) & nc & Chr(9) & t & Chr(9) & l & Chr$(9) & dl & Chr(9) & "2705" & Chr(9) & "0000000000" & Chr(9) & "2" & Chr(9) & rs("num_int")
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
Set rs2 = New ADODB.Recordset
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
msf1.Cols = 16
msf1.ColWidth(0) = 500
msf1.ColWidth(1) = 1100
msf1.ColWidth(2) = 3000
msf1.ColWidth(3) = 500
msf1.ColWidth(4) = 1200
msf1.ColWidth(5) = 500
msf1.ColWidth(6) = 800
msf1.ColWidth(7) = 1000
msf1.ColWidth(8) = 1100
msf1.ColWidth(9) = 1000
msf1.ColWidth(10) = 2000
msf1.ColWidth(11) = 1000
msf1.ColWidth(12) = 900
msf1.ColWidth(13) = 600
msf1.ColWidth(14) = 800
msf1.ColWidth(15) = 1500

msf1.TextMatrix(0, 0) = ""
msf1.TextMatrix(0, 1) = "Fecha"
msf1.TextMatrix(0, 2) = "Cliente"
msf1.TextMatrix(0, 3) = "Tipo"
msf1.TextMatrix(0, 4) = "Nro.Cuit/Dni "
msf1.TextMatrix(0, 5) = "Tipo"
msf1.TextMatrix(0, 6) = "Sucursal"
msf1.TextMatrix(0, 7) = "Numero "
msf1.TextMatrix(0, 8) = "Total"
msf1.TextMatrix(0, 9) = "Localidad"
msf1.TextMatrix(0, 10) = "Direccion"
msf1.TextMatrix(0, 11) = "C.P"
msf1.TextMatrix(0, 12) = "Nro. Partido"
msf1.TextMatrix(0, 13) = "Partido"
msf1.TextMatrix(0, 14) = "nro Int."
msf1.TextMatrix(0, 15) = "Errores"


For i = 0 To 1
  msf1.ColAlignment(i) = 1 'izq
Next i
For i = 2 To 14
  msf1.ColAlignment(i) = 9 'der
Next i

End Sub

Private Sub Form_Load()
Call carga_SUCURSALES(c_sucursal)
c_sucursal.AddItem "<Todas>", 0
c_sucursal.ListIndex = 0

Call barraesag(Me)
cal1.Visible = False
Call armagrid
End Sub



Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(1) = "[F2] Modifica Celda - [F4] Saca Fila - [F5] Archivo Exportacion - [F7] Imprime - [F11] Excel -"

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

    For i = 9 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "LISTADO DE IVA VENTAS", "", "Periodo: " & t_fecha & " : " & t_fecha2, "", 95, 6, True, False)
  End If

End If


If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If

If KeyCode = vbKeyF2 Then
  t = InputBox$(" ", "Cambia valor celda", msf1.TextMatrix(msf1.Row, msf1.col))
  If t <> "" Then
    't = t), "00000000000")
    msf1.TextMatrix(msf1.Row, msf1.col) = t
  End If
End If

If KeyCode = vbKeyF4 Then
  r = msf1.Row
  If r > 1 Then
   msf1.RemoveItem r
  End If
End If


If KeyCode = vbKeyF5 Then
  J = MsgBox("Confirma genera archivo para importar del aplicativo Citi Ventas. Archivo: c:\Aicyc.txt", 4)
  If J = 6 Then
    Call exporta
  End If
  
End If



End Sub


Sub exporta()
Dim c5 As String
k = 1
Open "c:\aicyc.txt" For Output As #1
While k <= msf1.Rows - 1
   c1 = Format$(msf1.TextMatrix(k, 2), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
   c2 = Format$(msf1.TextMatrix(k, 3), "0")
   If Val(msf1.TextMatrix(k, 3)) = 7 Then
       c3 = Format$(msf1.TextMatrix(k, 4), "00-00000000-0")
   Else
       c3 = Format$(msf1.TextMatrix(k, 4), "0000000000000")
   End If
   c4 = msf1.TextMatrix(k, 1)
   'c5 = Format$(Val(msf1.TextMatrix(k, 6)), "0000")
   c5 = Format$(msf1.TextMatrix(k, 6), "0000")
   c6 = Format$(msf1.TextMatrix(k, 7), "00000000")
   c7 = "0000000000000"
   c8 = Mid$(msf1.TextMatrix(k, 8), 1, 8) & Mid$(msf1.TextMatrix(k, 8), 10, 1)
   c9 = Format$(msf1.TextMatrix(k, 10), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
   c10 = "00000"
   c11 = "1"
   c12 = "000000"
   c13 = String(6, " ")
   c14 = String(3, " ")
   c15 = String(3, " ")
   c16 = String(4, " ")
   c17 = String(3, " ")
   c18 = "Rojas                                   "
   c19 = "02"
   c20 = "2705"
   c21 = String(5, " ")
   c22 = String(10, "0")
   c23 = String(5, " ")
   c24 = String(10, "0")
   c25 = String(40, " ")
   c26 = "La calle incluye el num. "
   c27 = "0000000000"
   c28 = "2"
   c29 = msf1.TextMatrix(k, 5)
      
      l = c1 & c2 & c3 & c4 & c5 & c6 & c7 & c8 & c9 & c10 & c11 & c12 & c13 & c14 & c15 & c16 & c17 & c18 & c19 & c20 & c21 & c22 & c23 & c24 & c25 & c26 & c27 & c28 & c29
      
   Print #1, l
   k = k + 1
Wend
Close #1
MsgBox ("Operacion Terminada. " & k - 1 & " registros generados. Ingrse al prgrama Aicyc del siap para importarlos")

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



Private Sub t_importe_LostFocus()
If Val(t_importe) < 0 Then
  t_importe = 0
End If
End Sub
