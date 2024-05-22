VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form cyb_concilia2 
   BackColor       =   &H00E0E0E0&
   Caption         =   "CONCILIACION AVANZADA"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos para conciliacion"
      Height          =   1215
      Left            =   240
      TabIndex        =   21
      Top             =   6960
      Width           =   9015
      Begin VB.TextBox t_fecha3 
         Height          =   285
         Left            =   2040
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ingresar Fecha"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Usa fecha Diferida"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipo Movimiento"
      Height          =   735
      Left            =   7200
      TabIndex        =   15
      Top             =   840
      Width           =   4455
      Begin VB.ComboBox c_tipo 
         Height          =   315
         ItemData        =   "CYB023.frx":0000
         Left            =   120
         List            =   "CYB023.frx":000D
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Consulta por:"
      Height          =   855
      Left            =   7200
      TabIndex        =   11
      Top             =   0
      Width           =   4455
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha Acred."
         Height          =   495
         Left            =   3240
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha Diferida"
         Height          =   495
         Left            =   1800
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha Emision"
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   5640
      TabIndex        =   10
      Top             =   0
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   111804417
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
      Width           =   6855
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "¿Entro?"
         Height          =   735
         Left            =   4080
         TabIndex        =   19
         Top             =   600
         Width           =   1455
         Begin VB.ComboBox c_c1 
            Height          =   315
            ItemData        =   "CYB023.frx":0032
            Left            =   120
            List            =   "CYB023.frx":003F
            TabIndex        =   20
            Text            =   "Combo1"
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   17
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox c_banco 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha  Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   1695
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
         Width           =   1695
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
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   4
      Top             =   7200
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "CYB023.frx":0054
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
         Picture         =   "CYB023.frx":08D6
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
            TextSave        =   "21/05/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "07:14 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "cyb_concilia2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim saldoanterior As Double
Sub graba()
If msf1.Rows > 1 Then
   h = 1
   While h < msf1.Rows
      q = "select * from cyb_04 where [num_mov_banco] = " & Val(msf1.TextMatrix(h, 11))
      Set rs = New adodb.Recordset
      rs.Open q, cn1, adOpenDynamic, adLockOptimistic
     ' rs.MaxRecords = 1
      If Not rs.EOF And Not rs.BOF Then
         If msf1.TextMatrix(h, 0) = "N" Then
            rs("entro") = "N"
            rs.Update
         Else
            If msf1.TextMatrix(h, 0) = "S" Then
              rs("entro") = "S"
              rs("fecha_acreed") = msf1.TextMatrix(h, 1)
              rs.Update
            End If
         End If
      End If
      Set rs = Nothing
      h = h + 1
    Wend
End If

End Sub

Sub carga()
  
  Call armagrid
  
  
  q = "select [fecha], [fecha_dif], [num_mov_banco], [abreviatura], cyb_04.[ubicacion], [importe], [detalle], [entro], [fecha_acreed], [num_comp], cyb_04.[id_tipomov] from cyb_04, cyb_06 where [id_banco] = " & c_banco.ItemData(c_banco.ListIndex) & " and cyb_04.[id_tipomov] = cyb_06.[id_tipomov]"
  Select Case c_c1.ListIndex
     Case Is = 1 'pendientes
      q = q & " and [entro] = 'S'"
     Case Is = 2 'ingresados
      q = q & " and [entro] = 'N'"
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
  
  If c_tipo.ListIndex > 0 Then
    q = q & " and cyb_04.[id_tipomov] = " & c_tipo.ItemData(c_tipo.ListIndex)
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

    
  Set rs = New adodb.Recordset
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
     o = rs("detalle")
     If rs("entro") = "N" Then
       fa = ""
     Else
       fa = rs("fecha_acreed")
     End If
     nc = Format$(rs("num_comp"), "0000000000")
     msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & rs("entro") & Chr$(9) & fa & Chr$(9) & fd & Chr(9) & t & Chr(9) & nc & Chr(9) & d & Chr(9) & h & Chr(9) & o & Chr(9) & F & Chr(9) & ni & Chr(9) & rs("id_tipomov")
     
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
msf1.Cols = 13
msf1.ColWidth(0) = 400
msf1.ColWidth(1) = 1000
msf1.ColWidth(2) = 600
msf1.ColWidth(3) = 1100
msf1.ColWidth(4) = 1100
msf1.ColWidth(5) = 1200
msf1.ColWidth(6) = 1200
msf1.ColWidth(7) = 1100
msf1.ColWidth(8) = 1100
msf1.ColWidth(9) = 4900
msf1.ColWidth(10) = 1100
msf1.ColWidth(11) = 800
msf1.ColWidth(12) = 800
msf1.TextMatrix(0, 0) = "NE"
msf1.TextMatrix(0, 1) = "NF"
msf1.TextMatrix(0, 2) = "Entro"
msf1.TextMatrix(0, 3) = "Fecha Acred."
msf1.TextMatrix(0, 4) = "Fecha Dif."
msf1.TextMatrix(0, 5) = "Tipo"
msf1.TextMatrix(0, 6) = "Nro."
msf1.TextMatrix(0, 7) = "Debe"
msf1.TextMatrix(0, 8) = "Haber"
msf1.TextMatrix(0, 9) = "Detalle"
msf1.TextMatrix(0, 10) = "Fecha Op."
msf1.TextMatrix(0, 11) = "Nro. Int"
msf1.TextMatrix(0, 12) = "Tipo Op."
End Sub








Private Sub c_tipo_LostFocus()
If c_tipo.ListIndex < 0 Then
  c_tipo.ListIndex = 0
End If
End Sub

Private Sub cal1_DblClick()
t_fecha = cal1
cal1.Visible = False
End Sub

Private Sub cal1_LostFocus()
t_fecha = cal1
cal1.Visible = False
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
Call carga_formas_pago(c_banco, "B")
c_banco.ListIndex = 0
Call armagrid
Call barraesag(Me)
cal1.Visible = False
Option2 = True
Call carga_mov_banco(c_tipo)
c_tipo.AddItem "<Todos>", 0
c_tipo.ListIndex = 0
c_c1.ListIndex = 0
Option4 = True
t_fecha3.Visible = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload cyb_cc_detalleb
Unload cyb_concilia
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F2] Entro = S - [F4] Saca -  [F6] Entro = N -  [ESP] Cambia estado [F9] Graba "
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)


If KeyCode = vbKeyF4 Then
  r = msf1.Row
  p = msf1.Rows
  If p > 2 Then
    msf1.RemoveItem r
  Else
    Call armagrid
  End If
End If

If KeyCode = vbKeyF2 Then
  If msf1.Rows > 1 Then
    h = 1
    While h < msf1.Rows
      msf1.TextMatrix(h, 0) = "S"
      If Option4 = True Then
        msf1.TextMatrix(h, 1) = msf1.TextMatrix(h, 4)
      Else
        msf1.TextMatrix(h, 1) = t_fecha3
      End If
     h = h + 1
    Wend
  End If
End If


If KeyCode = vbKeyF6 Then
  If msf1.Rows > 1 Then
    h = 1
    While h < msf1.Rows
      msf1.TextMatrix(h, 0) = "N"
      msf1.TextMatrix(h, 1) = ""
      h = h + 1
    Wend
  End If
End If


If KeyCode = vbKeySpace Then
  h = msf1.Row
  If h > 0 Then
    If msf1.TextMatrix(h, 0) = "" Then
      If msf1.TextMatrix(h, 2) = "S" Then
        msf1.TextMatrix(h, 0) = "N"
        msf1.TextMatrix(h, 0) = ""
      Else
         msf1.TextMatrix(h, 0) = "S"
         If Option4 = True Then
           msf1.TextMatrix(h, 1) = msf1.TextMatrix(h, 4)
         Else
           msf1.TextMatrix(h, 1) = t_fecha3
         End If
       End If
    Else
      If msf1.TextMatrix(h, 0) = "S" Then
        msf1.TextMatrix(h, 0) = "N"
        msf1.TextMatrix(h, 1) = ""
      Else
        msf1.TextMatrix(h, 0) = "S"
        If Option4 = True Then
           msf1.TextMatrix(h, 1) = msf1.TextMatrix(h, 4)
        Else
           msf1.TextMatrix(h, 1) = t_fecha3
        End If
     End If
  End If
End If
End If


If KeyCode = vbKeyF9 Then
  J = MsgBox("Confirma Grabar Cociliacion", 4)
  If J = 6 Then
    Call graba
  End If
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

Private Sub Option4_Click()
t_fecha3.Visible = False
End Sub

Private Sub Option5_Click()
t_fecha3.Visible = True
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

Private Sub t_fecha3_LostFocus()
If t_fecha3 <> "" Then
  If Not IsDate(t_fecha3) Then
    t_fecha3 = ""
  End If
End If

End Sub
