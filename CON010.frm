VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form con_ley23966 
   BackColor       =   &H00E0E0E0&
   Caption         =   "REGISTRO LEY 23966  ART. 15"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   4320
      TabIndex        =   8
      Top             =   120
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   175964161
      CurrentDate     =   38754
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5895
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10398
      _Version        =   393216
      AllowUserResizing=   1
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
      Height          =   1095
      Left            =   240
      TabIndex        =   6
      Top             =   0
      Width           =   4335
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   9
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1935
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
         Width           =   1935
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
         Picture         =   "CON010.frx":0000
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
         Picture         =   "CON010.frx":0882
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
      Top             =   8235
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   12347
            MinWidth        =   12347
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "08/10/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:11 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "con_ley23966"
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
  T2 = 0
  If t_fecha <> "" Then
     q = "select * from a19 where datevalue([fecha]) < datevalue('" & t_fecha & "')"
     Set rs = New ADODB.Recordset
     rs.Open q, cn1
     While Not rs.EOF
       t = rs("importe")
       If rs("ubicacion") = "D" Then
        da = da + t
       Else
        ha = ha + t
       End If
       rs.MoveNext
     Wend
    sa = da - ha
  End If
  
  saldoanterior = sa
  msf1.AddItem t_fecha & Chr(9) & "Saldo Ant." & Chr(9) & Format$(da, "######0.00") & Chr(9) & Format$(ha, "######0.00") & Chr(9) & Format$(sa, "######0.00")
  
  q = "select * from a19 "
  c = " where "
  If t_fecha <> "" Then
       q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
       c = " and "
  End If
    
  If t_fecha2 <> "" Then
       q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  q = q & " order by [fecha] "
    
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  s = sa
  While Not rs.EOF
     F = rs("fecha")
     t = rs("importe")
     If rs("ubicacion") = "D" Then
       d = Format$(t, "######0.00")
       h = ""
       l = Format$(rs("litros"), "####0.00")
       ii = Format$(rs("pu_impuesto_int"), "##0.0000")
        
       Set rs1 = New ADODB.Recordset
       q = "select [letra], [num_comprobante], [sucursal] from a5 where [num_int] = " & rs("num_int")
       rs1.Open q, cn1
       If Not rs1.EOF And Not rs1.BOF Then
         nc = rs1("letra") & Format$(rs1("sucursal"), "0000") & "-" & Format$(rs1("num_comprobante"), "00000000")
       Else
         nc = ""
       End If
     Else
       h = Format$(t, "######0.00")
       d = ""
       l = ""
       ii = ""
       nc = ""
     End If
     s = Format$(Val(s) + Val(d) - Val(h), "######0.00")
     ni = rs("num_int")
     o = rs("detalle")
     msf1.AddItem F & Chr(9) & nc & Chr(9) & d & Chr(9) & h & Chr(9) & s & Chr(9) & l & Chr(9) & ii & Chr(9) & o & Chr(9) & rs("num_int")
     rs.MoveNext
  Wend
  
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
  msf1.Cols = 10
  msf1.ColWidth(0) = 1200
  msf1.ColWidth(1) = 1600
  msf1.ColWidth(2) = 1200
  msf1.ColWidth(3) = 1200
  msf1.ColWidth(4) = 1200
  msf1.ColWidth(5) = 1200
  msf1.ColWidth(6) = 1200
  msf1.ColWidth(7) = 2000
  msf1.TextMatrix(0, 0) = "Fecha"
  msf1.TextMatrix(0, 1) = "Comprobante"
  msf1.TextMatrix(0, 2) = "Debe($)"
  msf1.TextMatrix(0, 3) = "Haber($)"
  msf1.TextMatrix(0, 4) = "Saldo($)"
  msf1.TextMatrix(0, 5) = "Litros"
  msf1.TextMatrix(0, 6) = "Imp.Int($/lt)"
  msf1.TextMatrix(0, 7) = "Detalle"
  msf1.TextMatrix(0, 8) = "N.I."
  
  For i = 0 To 1
    msf1.ColAlignment(i) = 1
  Next i
  For i = 2 To 6
    msf1.ColAlignment(i) = 9
  Next i
  For i = 7 To 7
    msf1.ColAlignment(i) = 1
  Next i

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


Call armagrid
Call barraesag(Me)
cal1.Visible = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Unload vta_clientes
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F1]Ingresa Pago a cuenta -  [F7] Imprime - [F8] Borra Subsidio - [F11] Excel  "
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
  Call nivel_acceso(1)
  If para.id_grupo_modulo_actual >= 4 Then
    J = MsgBox("Prepare Impresora y Confirme", 4)
    If J = 6 Then
     Dim c(15) As Double

     If Check1 = 0 Then
      c(0) = 9
      c(1) = 0
      c(2) = 2
      c(3) = 3
      c(4) = 4
      c(5) = 5
      c(6) = 6
      c(7) = 7
      For i = 8 To 14
        c(i) = -1
      Next i
    Else
      
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
     End If
     Call imprimegrid(msf1, c(), "REGISTRO SUBSIDIO LEY 23966 ART.15", "", " ", "Periodo: " & t_fecha & "  " & t_fecha2, 85, 7, True, False)

    End If
         
  End If
  
End If

If KeyCode = vbKeyF8 Then
  Call nivel_acceso(2)
  If para.id_grupo_modulo_actual >= 8 Then
   J = MsgBox("Confirma Eliminar Subsidio Nro." & msf1.TextMatrix(msf1.RowSel, 8) & ". Nota: Solo se borra el subsidio y no el comprobante", 4)
   If J = 6 Then
      indice = Val(msf1.TextMatrix(msf1.RowSel, 8))
      Set rs = New ADODB.Recordset
      q = "select * from a19 where [num_int] = " & indice
      rs.Open q, cn1, adOpenDynamic, adLockOptimistic
      If Not rs.EOF And Not rs.BOF Then
        rs.Delete
        rs.Update
      End If
      Set rs = Nothing
      MsgBox ("Operacion Terminada")
      Call carga
   End If
  End If
End If


If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If

If KeyCode = vbKeyF1 Then
  con_ley23966A.Show
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
