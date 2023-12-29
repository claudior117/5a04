VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form cyb_estadocuenta 
   BackColor       =   &H00E0E0E0&
   Caption         =   "ESTADO DE CUENTA BANCARIA"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Tools"
      Height          =   735
      Left            =   240
      TabIndex        =   19
      Top             =   7200
      Width           =   5655
      Begin VB.CommandButton Command3 
         Caption         =   "Depositos"
         Height          =   375
         Left            =   3720
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ch. Propios"
         Height          =   375
         Left            =   1920
         TabIndex        =   21
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cr. y Db. Bancarios"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mostrar movimientos segun estado conciliacion:"
      Height          =   735
      Left            =   7680
      TabIndex        =   15
      Top             =   840
      Width           =   3975
      Begin VB.ComboBox c_concilia 
         Height          =   315
         ItemData        =   "CYB009.frx":0000
         Left            =   120
         List            =   "CYB009.frx":000D
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ordenados por:"
      Height          =   855
      Left            =   7680
      TabIndex        =   11
      Top             =   0
      Width           =   3975
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha Acreditacion"
         Height          =   495
         Left            =   2640
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha Diferida"
         Height          =   495
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha Emision"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   4560
      TabIndex        =   10
      Top             =   120
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   114688001
      CurrentDate     =   38754
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
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   7215
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   17
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox c_banco 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Banco:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   4
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "CYB009.frx":0032
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
         Picture         =   "CYB009.frx":08B4
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
            Object.Width           =   353
            MinWidth        =   353
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   17639
            MinWidth        =   17639
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "21/12/2023"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:40 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "cyb_estadocuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim saldoanterior As Double
Sub carga()
  
  Call armagrid
  
  sa = 0
  da = 0
  ha = 0
  q = "select [importe], [ubicacion] from cyb_04 where [id_banco] = " & c_banco.ItemData(c_banco.ListIndex)
  c = " and "
  If t_fecha <> "" Then
   If Option1 = True Then
      q = q & c & " datevalue([fecha]) < datevalue('" & t_fecha & "')"
   Else
      If Option2 = True Then
        q = q & c & " datevalue([fecha_dif]) < datevalue('" & t_fecha & "')"
      Else
        q = q & c & " datevalue([fecha_acreed]) < datevalue('" & t_fecha & "')"
      End If
   End If
    
    
    Select Case c_concilia.ListIndex
     Case Is = 1 'pendientes
      q = q & " and [entro] = 'N'"
     Case Is = 2 'ingresados
      q = q & " and [entro] = 'S'"
    End Select
    
    Set rs = New ADODB.Recordset
    rs.Open q, cn1
    While Not rs.EOF
     If rs("ubicacion") = "D" Then
        da = da + rs("importe")
     Else
        ha = ha + rs("importe")
     End If
     rs.MoveNext
    Wend
    Set rs = Nothing
    sa = ha - da
  End If
  
  
  saldoanterior = sa
  msf1.AddItem t_fecha & Chr(9) & Chr(9) & Chr(9) & Format$(da, "######0.00") & Chr(9) & Format$(ha, "######0.00") & Chr(9) & Format$(sa, "######0.00") & Chr(9) & "Saldo Inicial"
  
  q = "select [fecha], [fecha_dif], [num_mov_banco], [abreviatura], cyb_04.[ubicacion], [importe], [detalle], [entro], [fecha_acreed], [num_comp], cyb_04.[id_tipomov] from cyb_04, cyb_06 where [id_banco] = " & c_banco.ItemData(c_banco.ListIndex) & " and cyb_04.[id_tipomov] = cyb_06.[id_tipomov]"
  Select Case c_concilia.ListIndex
     Case Is = 1 'pendientes
      q = q & " and [entro] = 'N'"
     Case Is = 2 'ingresados
      q = q & " and [entro] = 'S'"
    End Select
    
  If t_fecha <> "" Then
   If Option1 = True Then
    q = q & " and datevalue([fecha]) >= datevalue('" & t_fecha & "')"
    
   Else
    If Option2 = True Then
       q = q & " and datevalue([fecha_dif]) >= datevalue('" & t_fecha & "')"
    Else
       q = q & " and datevalue([fecha_acreed]) >= datevalue('" & t_fecha & "')"
    End If
   End If
  End If
  
  If t_fecha2 <> "" Then
   If Option1 = True Then
    q = q & " and datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
    
   Else
    If Option2 = True Then
       q = q & " and datevalue([fecha_dif]) <= datevalue('" & t_fecha2 & "')"
    Else
       q = q & " and datevalue([fecha_acreed]) <= datevalue('" & t_fecha2 & "')"
    End If
   End If
  End If
  
  
   If Option1 = True Then
     q = q & " order by [fecha]"
   Else
     If Option2 = True Then
        q = q & " order by [fecha_dif]"
     Else
        q = q & " order by [fecha_acreed]"
     End If
   End If

    
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  s = sa
  While Not rs.EOF
     F = rs("fecha")
     fd = rs("fecha_dif")
     ni = Format$(rs("num_mov_banco"), "00000")
     t = rs("abreviatura")
     If rs("ubicacion") = "D" Then
       d = Format$(rs("importe"), "######0.00")
       h = ""
     Else
         h = Format$(rs("importe"), "######0.00")
         d = ""
     End If
     s = Format$(Val(s) - Val(d) + Val(h), "######0.00")
     o = rs("detalle")
     If rs("entro") = "N" Then
       fa = ""
     Else
       fa = rs("fecha_acreed")
     End If
     nc = Format$(rs("num_comp"), "0000000000")
     msf1.AddItem fd & Chr(9) & t & Chr(9) & nc & Chr(9) & d & Chr(9) & h & Chr(9) & s & Chr(9) & o & Chr$(9) & rs("entro") & Chr$(9) & fa & Chr(9) & F & Chr(9) & ni & Chr(9) & rs("id_tipomov")
     
    rs.MoveNext
  Wend
  Set rs = Nothing
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
msf1.Cols = 12
msf1.ColWidth(0) = 1100
msf1.ColWidth(1) = 1200
msf1.ColWidth(2) = 1200
msf1.ColWidth(3) = 1100
msf1.ColWidth(4) = 1100
msf1.ColWidth(5) = 1300
msf1.ColWidth(6) = 4900
msf1.ColWidth(7) = 600
msf1.ColWidth(8) = 1100
msf1.ColWidth(9) = 1100
msf1.ColWidth(10) = 800
msf1.ColWidth(11) = 800
msf1.TextMatrix(0, 0) = "Fecha Dif."
msf1.TextMatrix(0, 1) = "Tipo"
msf1.TextMatrix(0, 2) = "Nro."
msf1.TextMatrix(0, 3) = "Debe"
msf1.TextMatrix(0, 4) = "Haber"
msf1.TextMatrix(0, 5) = "Saldo"
msf1.TextMatrix(0, 6) = "Detalle"
msf1.TextMatrix(0, 7) = "Entro"
msf1.TextMatrix(0, 8) = "Fecha Acred."
msf1.TextMatrix(0, 9) = "Fecha Op."
msf1.TextMatrix(0, 10) = "Nro. Int"
msf1.TextMatrix(0, 11) = "Tipo Op."
End Sub








Private Sub cal1_DblClick()
t_fecha = cal1
cal1.Visible = False
End Sub

Private Sub cal1_LostFocus()
t_fecha = cal1
cal1.Visible = False
End Sub

Private Sub Command1_Click()
 Call nivel_acceso(4)
    If para.id_grupo_modulo_actual >= 5 Then
      cyb_movbanco.Show
    Else
      Call sinpermisos
    End If
End Sub

Private Sub Command2_Click()
Call nivel_acceso(4)
    If para.id_grupo_modulo_actual >= 4 Then
      cyb_chpropios.Show
    Else
      Call sinpermisos
    End If
 End Sub

Private Sub Command3_Click()
Call nivel_acceso(4)
    If para.id_grupo_modulo_actual >= 5 Then
      cyb_depositoS.Show
    Else
      Call sinpermisos
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 2)
  'Case Is = 27
  '      Me.Hide
End Select

End Sub

Private Sub Form_Load()
Load cyb_cc_detalleb
Load cyb_concilia
Call carga_formas_pago(c_banco, "B")
c_banco.ListIndex = 0
Call armagrid
Call barraesag(Me)
cal1.Visible = False
Option2 = True
c_concilia.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload cyb_cc_detalleb
Unload cyb_concilia
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F7] Imprime - [F8] Borra Mov. - [F5] Concilia - [ENTER] Detalle - [F3] Abrir O.P. -[F11] Excel"
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

    For i = 8 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "ESTADO CUENTA BANCARIO", "", "Periodo...: " & t_fecha, "Banco.....:" & c_banco, 85, 7, True, False)
  End If
    
End If


If KeyCode = vbKeyF8 Then
 'borrar mov.
 J = MsgBox("Confirma Borrar Movimiento", 4)
 If J = 6 Then
    ni = Val(msf1.TextMatrix(msf1.Row, 10))
    Set cl_banco = New bancos
    cl_banco.borrar (ni)
    Set cl_banco = Nothing
 End If
    
End If


If KeyCode = vbKeyF5 Then
    cyb_concilia.t_id = Val(msf1.TextMatrix(msf1.Row, 10))
    cyb_concilia.t_mov = Val(msf1.TextMatrix(msf1.Row, 3)) + Val(msf1.TextMatrix(msf1.Row, 4))
    cyb_concilia.t_entro = msf1.TextMatrix(msf1.Row, 7)
    If cyb_concilia.t_entro = "S" Then
      cyb_concilia.t_fecha = msf1.TextMatrix(msf1.Row, 8)
      cyb_concilia.t_entro = "N"
    Else
      cyb_concilia.t_fecha = msf1.TextMatrix(msf1.Row, 0)
      cyb_concilia.t_entro = "S"
    End If
    cyb_concilia.t_tipomov = msf1.TextMatrix(msf1.Row, 11)
    cyb_concilia.t_idbanco = c_banco.ItemData(c_banco.ListIndex)
    cyb_concilia.Show
End If


If KeyCode = vbKeyF3 Then
    q = "select * from cyb_04 where [num_mov_banco] = " & Val(msf1.TextMatrix(msf1.Row, 10))
    Set rs = New ADODB.Recordset
    rs.Open q, cn1
    If Not rs.EOF And Not rs.BOF Then
       If rs("modulo") = "C" And rs("num_mov_int") > 0 Then
          q = "select * from a5 where [num_int] = " & rs("num_mov_int")
          Set rs1 = New ADODB.Recordset
          rs1.Open q, cn1
          If Not rs1.EOF And Not rs1.BOF Then
            If rs1("id_tipocomp") = 50 Then
               Load cc_detalle
               cc_detalle.t_numint = rs1("num_int")
               cc_detalle.Show
            End If
          End If
          Set rs1 = Nothing
       End If
     End If
    Set rs = Nothing
End If

If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If

End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  cyb_cc_detalleb.t_numint = msf1.TextMatrix(msf1.Row, 10)
  cyb_cc_detalleb.Show
End If
End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight

End Sub

Private Sub t_fecha_DblClick()
cal1.Visible = True
End Sub

Private Sub t_fecha_GotFocus()
t_fecha = ""
End Sub

Private Sub t_fecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call carga
End If
End Sub

Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
    t_fecha = ""
  End If
End If
End Sub

Private Sub t_fecha2_LostFocus()
If t_fecha2 <> "" Then
  If Not IsDate(t_fecha2) Then
    t_fecha2 = ""
  End If
End If

End Sub
