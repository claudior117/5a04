VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form emp_estadocuenta 
   BackColor       =   &H00E0E0E0&
   Caption         =   "ESTADO DE CUENTA POR EMPLEADO"
   ClientHeight    =   8805
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12285
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   12285
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5655
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9975
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1455
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   7335
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   10
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
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   2160
         TabIndex        =   0
         Text            =   "c_prov"
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
         TabIndex        =   11
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
         Caption         =   "Empleado:"
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
         Picture         =   "emp003.frx":0000
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
         Picture         =   "emp003.frx":0882
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
      Top             =   8550
      Width           =   12285
      _ExtentX        =   21669
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
            TextSave        =   "11:28 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "emp_estadocuenta"
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
  If t_fecha <> "" Then
    q = "select * from emp_02 where [id_legajo] = " & c_prov.ItemData(c_prov.ListIndex)
    q = q & " and datevalue([fecha]) < datevalue('" & t_fecha & "')"
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
  msf1.AddItem t_fecha & Chr(9) & "" & Chr(9) & "Saldo Ant." & Chr(9) & Format$(da, "######0.00") & Chr(9) & Format$(ha, "######0.00") & Chr(9) & Format$(sa, "######0.00")
  
  q = "select * from emp_02 where [id_legajo] = " & c_prov.ItemData(c_prov.ListIndex)
  If t_fecha <> "" Then
       q = q & " and datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
    
  If t_fecha2 <> "" Then
       q = q & " and datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
    
  q = q & " order by [fecha]"
    
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  s = sa
  
  sao = ""
  dao = ""
  hao = ""
  While Not rs.EOF
     F = rs("fecha")
     Select Case rs("tipo_movimiento")
     Case Is = 1
        tm = "Adelanto"
     Case Is = 20
        tm = "Gasto"
     Case Is = 100
        tm = "Rbo Sueldo"
    End Select
     
     ob = rs("observaciones")
     t = rs("importe")
     If rs("ubicacion") = "D" Then
       d = Format$(t, "######0.00")
       h = ""
     Else
       h = Format$(t, "######0.00")
       d = ""
     End If
     s = Format$(Val(s) + Val(d) - Val(h), "######0.00")
     ni = rs("num_mov_int")
     If t > 0 Then
      msf1.AddItem F & Chr(9) & tm & Chr(9) & ob & Chr(9) & d & Chr(9) & h & Chr(9) & s & Chr(9) & ni
     End If
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
  msf1.Cols = 8
  msf1.ColWidth(0) = 1300
  msf1.ColWidth(1) = 2000
  msf1.ColWidth(2) = 3500
  msf1.ColWidth(3) = 1200
  msf1.ColWidth(4) = 1200
  msf1.ColWidth(5) = 1200
  msf1.ColWidth(6) = 1000
  msf1.ColWidth(7) = 500
  
  msf1.TextMatrix(0, 0) = "Fecha"
  msf1.TextMatrix(0, 1) = "Tipo Mov."
  msf1.TextMatrix(0, 2) = "Observaciones"
  msf1.TextMatrix(0, 3) = "Debe($)"
  msf1.TextMatrix(0, 4) = "Haber($)"
  msf1.TextMatrix(0, 5) = "Saldo($)"
  msf1.TextMatrix(0, 6) = "Num.Mov."
  msf1.TextMatrix(0, 7) = " "
End Sub







Private Sub c_prov_LostFocus()
If c_prov.ListIndex < 0 Then
  c_prov.ListIndex = 0
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

Call carga_empleados(c_prov)
c_prov.ListIndex = 0
Call armagrid
Call barraesag(Me)

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload vta_clientes
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F7] Imprime - [F8] Borra Mov. - [F11] Exporta Excel  "
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

      c(0) = 7
      c(1) = 0
      c(2) = 1
      c(3) = 2
      c(4) = 3
      c(5) = 4
      c(6) = 5
      For i = 7 To 14
        c(i) = -1
      Next i
     Call imprimegrid(msf1, c(), "ESTADO DE CUENTA EMPLEADOS", "", "Empleado: " & c_prov, "Periodo: " & t_fecha & "  " & t_fecha2, 85, 7, True, False)

    End If
         
  End If
  
End If


If KeyCode = vbKeyF8 Then
  Call nivel_acceso(1)
  If para.id_grupo_modulo_actual >= 8 Then
    J = MsgBox("Confirma Borrar todo el movimiento en el empleado", 4)
    If J = 6 Then
      QUERY = "DELETE FROM emp_02 WHERE [num_mov_int] = " & Val(msf1.TextMatrix(msf1.Row, 6)) & " and [id_legajo] = " & c_prov.ItemData(c_prov.ListIndex)
      cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans
    End If
    
  Else
    Call sinpermisos
  End If
End If

If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
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
