VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form cgr_aGRUPA 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Agrupa Asientos por dia"
   ClientHeight    =   8355
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   11925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Excluir asientos apertura/cierre"
      Height          =   1095
      Left            =   3480
      TabIndex        =   20
      Top             =   120
      Width           =   3255
      Begin VB.TextBox t_ainicio 
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox t_acierre 
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   21
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF0000&
         Caption         =   "Asiento Apertura:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
         Caption         =   "Asiento Cierre:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Agrupa"
      Height          =   1095
      Left            =   6840
      TabIndex        =   17
      Top             =   120
      Width           =   2775
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Un asiento por Mes"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Un asiento por dia"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1815
      Left            =   720
      TabIndex        =   15
      Top             =   1680
      Width           =   9615
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   $"CGR019.frx":0000
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   8775
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   975
      Left            =   360
      TabIndex        =   10
      Top             =   7440
      Visible         =   0   'False
      Width           =   8295
      Begin VB.TextBox t_idperiodo 
         Height          =   285
         Left            =   6120
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox t_tothaber 
         Height          =   285
         Left            =   4200
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox t_totdebe 
         Height          =   285
         Left            =   2280
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox t_fecha 
         Height          =   285
         Left            =   360
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fechas  Corte"
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3255
      Begin VB.TextBox t_f2 
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox t_f1 
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000FF&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   4935
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8705
      _Version        =   393216
      FixedCols       =   0
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10080
      TabIndex        =   3
      Top             =   240
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "CGR019.frx":01E2
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
         Picture         =   "CGR019.frx":0A64
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
      Top             =   8100
      Width           =   11925
      _ExtentX        =   21034
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
            TextSave        =   "17/11/2019"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:47"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "cgr_aGRUPA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Private Sub btnacepta_Click()
If Option1 = True Then
  Call pordia
Else
  Call pormes
End If
  

End Sub

Sub pordia()
If verifica Then
  J = MsgBox("Este proceso es irreversible, realice un backup antes de ejecutarlo", 4)
  If J = 6 Then
   h = MsgBox("Confirma realmente realizar este proceso", 4)
   If h = 6 Then
     Call armagrid
     Call limpia
   End If
  End If
Else
  MsgBox ("Verifique el periodo ingresado")
End If
End Sub

Sub pormes()
If verifica Then
  J = MsgBox("Este proceso es irreversible, realice un backup antes de ejecutarlo", 4)
  If J = 6 Then
   h = MsgBox("Confirma realmente realizar este proceso", 4)
   If h = 6 Then
     Call armagrid
     Call limpia2
   End If
  End If
Else
  MsgBox ("Verifique el periodo ingresado")
End If
End Sub


Function verifica()
   v = True
   If t_f1 <> "" Then
     If Not IsDate(t_f1) Then
       v = False
     Else
       t_f1 = Format$(t_f1, "dd/mm/yyyy")
     End If
  Else
    v = False
  End If
  
   
   If t_f2 <> "" Then
     If Not IsDate(t_f2) Then
       v = False
     Else
      t_f2 = Format$(t_f2, "dd/mm/yyyy")
     End If
  Else
    v = False
  End If
   
  If v Then
   If DateValue(t_f2) < DateValue(t_f1) Then
    v = False
   End If
  End If
  verifica = v
  
   
End Function


Private Sub btnsale_Click()
Unload Me
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub
Sub limpia()
espere.Show
espere.Refresh
Dim q As String
d = 0
fa = DateValue(t_f1) + d

While fa <= DateValue(t_f2)
  espere.Label1 = "Agrupando Asiento Fecha: " & fa
  espere.Refresh
  Call armagrid
  q = "select * from c_11 where datevalue(fecha) = '" & fa & "'"
  Set rs = New ADODB.Recordset
  rs.Open q, cn1, adOpenDynamic, adLockOptimistic
  p = 0
  While Not rs.EOF
    If rs("num_asiento") <> Val(t_ainicio) And rs("num_asiento") <> Val(t_acierre) Then
  
      If p = 0 Then
        t_fecha = rs("fecha")
        t_idperiodo = rs("id_periodo")
        T_mes = rs("mes")
        t_año = rs("año")
        p = 1
      End If
      Set rs1 = New ADODB.Recordset
      q = "select * from c_12 where c_12.[id_asiento] = " & rs("id_asiento")
      rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
      While Not rs1.EOF
        Call busca(rs1("id_cuenta"), rs1("ubicacion"), rs1("importe"))
        rs1.Delete
        rs1.MoveNext
      Wend
      rs.Delete
     End If
     rs.MoveNext
  Wend
  If p <> 0 Then
    Call graba
  End If
  d = d + 1
  fa = DateValue(t_f1) + d
Wend
Call suma

Unload espere
End Sub

Sub limpia2()
espere.Show
espere.Refresh
Dim q As String
mi = Mid$(t_f1, 4, 2)
ai = Mid$(t_f1, 7, 4)

mf = Mid$(t_f2, 4, 2)
af = Mid$(t_f2, 7, 4)

d = 0
m = Val(mi)
a = Val(ai)

While d = 0
  espere.Label1 = "Agrupando Asiento Mes: " & Format$(m & "/" & a, "@@@@@@@@@@")
  espere.Refresh
  Call armagrid
  
  'intervalo
  f1 = "01" & "/" & Format$(m, "00") & "/" & Format$(a, "0000")
  
  Select Case m
   Case Is = 1, Is = 3, Is = 5, Is = 7, Is = 8, Is = 10, Is = 12
       dia = 31
   Case Is = 4, Is = 6, Is = 9, Is = 11
      dia = 30
   Case Is = 2
     If IsDate("29/2/" & Format$(a, "0000")) Then
         dia = 29
     Else
         dia = 28
     End If
  End Select
  
  f2 = Format$(dia, "00") & "/" & Format$(m, "00") & "/" & Format$(a, "0000")
  'MsgBox (f1 & " ---- " & f2)
  q = "select * from c_11 where datevalue(fecha) >= datevalue('" & f1 & "') and datevalue(fecha) <= datevalue('" & f2 & "')"
  'MsgBox (q)
  Set rs = New ADODB.Recordset
  rs.Open q, cn1, adOpenDynamic, adLockOptimistic
  p = 0
  While Not rs.EOF
   If rs("num_asiento") <> Val(t_ainicio) And rs("num_asiento") <> Val(t_acierre) Then
  
      If p = 0 Then
        t_fecha = f2
        t_idperiodo = rs("id_periodo")
        T_mes = m
        t_año = a
        p = 1
      End If
      Set rs1 = New ADODB.Recordset
      q = "select * from c_12 where c_12.[id_asiento] = " & rs("id_asiento")
      rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
      'MsgBox (q)
      While Not rs1.EOF
        Call busca(rs1("id_cuenta"), rs1("ubicacion"), rs1("importe"))
        rs1.Delete
        rs1.MoveNext
      Wend
      Set rs1 = Nothing
      rs.Delete
     End If
     rs.MoveNext
  Wend
  Set rs = Nothing
  
  If p <> 0 Then
    Call graba
  End If
  m = m + 1
  If m > 12 Then
     m = 1
     a = a + 1
  End If
   
  If (m > Val(mf) And a = Val(af)) Or a > Val(af) Then
     d = 1
   End If
  
Wend
Call suma

'renumeramos el asiento de cierre para que sea el ultimo
If Val(t_acierre) > 0 Then
  Set rs1 = New ADODB.Recordset
  q = "select * from c_11 where [num_asiento] = " & Val(t_acierre)
  rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
  If Not rs1.EOF And Not rs.BOF Then
    na = Val(Mid$(t_acierre, 1, 6) & "999")
    rs1("num_asiento") = na
    rs1.Update
  End If
  Set rs1 = Nothing
End If
Unload espere
End Sub


Sub graba()
    'saco numero
        a = Format$(Val(Mid$(t_fecha, 7, 4)), "0000")
        m = Format$(Val(Mid$(t_fecha, 4, 2)), "00")
       a1 = Val(a & m & "000")
       a2 = Val(a & m & "999")
      
       Set rs = New ADODB.Recordset
       q = "select * from c_11 where [año] = " & Val(a) & " and [mes] = " & Val(m)
       rs.Open q, cn1, adOpenDynamic, adLockOptimistic
       If Not rs.EOF And Not rs.BOF Then
         rs.MoveLast
         na = rs("num_asiento") + 1
       Else
         na = Val(a & m & "001")
       End If
       Set rs = Nothing
      
      
      cn1.BeginTrans
      QUERY = "INSERT INTO c_11([num_asiento], [fecha], [descripcion], [id_periodo], [importe], [año], [mes])"
      QUERY = QUERY & " VALUES (" & na & ", '" & t_fecha & "', 'As. Resumen', " & Val(t_idperiodo) & ", " & Val(t_totdebe) & ", " & Val(a) & ", " & Val(m) & ")"
      cn1.Execute QUERY
      
      qr = "SELECT @@IDENTITY AS NewID"
      Set rs = cn1.Execute(qr)
      nic = rs.Fields("NewID").Value

      
      s = 1
      For i = 1 To msf1.Rows - 1
        If Val(msf1.TextMatrix(i, 1)) > 0 Then
          QUERY = "INSERT INTO c_12([id_asiento], [secuencia], [id_cuenta], [importe], [descripcion], [ubicacion])"
          QUERY = QUERY & " VALUES (" & nic & ", " & s & ", " & Val(msf1.TextMatrix(i, 0)) & ", " & Val(msf1.TextMatrix(i, 1)) & ", '" & msf1.TextMatrix(i, 3) & "', 'D')"
          cn1.Execute QUERY
          s = s + 1
        End If
      Next i
      
      For i = 1 To msf1.Rows - 1
       If Val(msf1.TextMatrix(i, 2)) > 0 Then
        QUERY = "INSERT INTO c_12([id_asiento], [secuencia], [id_cuenta], [importe], [descripcion], [ubicacion])"
        QUERY = QUERY & " VALUES (" & nic & ", " & s & ", " & Val(msf1.TextMatrix(i, 0)) & ", " & Val(msf1.TextMatrix(i, 2)) & ", '" & msf1.TextMatrix(i, 3) & "', 'H')"
        cn1.Execute QUERY
        s = s + 1
       End If
      Next i
      cn1.CommitTrans
  
End Sub
Sub suma()
i = 0
d = 0
h = 0
While i < msf1.Rows - 1
   d = d + Val(msf1.TextMatrix(i, 1))
   h = h + Val(msf1.TextMatrix(i, 2))
   i = i + 1
Wend
msf1.AddItem "" & Chr$(9) & "--------------------" & Chr$(9) & "--------------------"
msf1.AddItem "" & Chr$(9) & d & Chr$(9) & h
t_totdebe = d
t_tothaber = h

End Sub
Sub busca(ByVal c, ByVal u, ByVal importe)
i = 0
e = 0
th = ""
td = ""
While i < msf1.Rows - 1
   If Val(msf1.TextMatrix(i, 0)) = c Then
     If u = "D" Then
       msf1.TextMatrix(i, 1) = Val(msf1.TextMatrix(i, 1)) + importe
     Else
        msf1.TextMatrix(i, 2) = Val(msf1.TextMatrix(i, 2)) + importe
    End If
    i = msf1.Rows
    e = 1
  End If
  i = i + 1
Wend
If e = 0 Then
   Set rs3 = New ADODB.Recordset
   q = "select * from c_01 where [id_cuenta] = " & c
   rs3.Open q, cn1
   If Not rs3.EOF And Not rs3.BOF Then
     d = rs3("descripcion")
   Else
     d = "Cuenta Inexistente"
   End If
   Set rs3 = Nothing
   
   If u = "D" Then
          td = importe
          th = ""
   Else
          td = ""
          th = importe
   End If
        
   msf1.AddItem c & Chr$(9) & td & Chr$(9) & th & Chr$(9) & d
  
End If

End Sub
Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 5
msf1.AllowUserResizing = flexResizeNone
msf1.FixedCols = 0
msf1.SelectionMode = flexSelectionByRow
msf1.FocusRect = flexFocusNone
msf1.ColWidth(0) = 2000
msf1.ColWidth(1) = 2000
msf1.ColWidth(2) = 2000
msf1.ColWidth(3) = 2000

msf1.TextMatrix(0, 0) = "Cuenta"
msf1.TextMatrix(0, 1) = "Debe"
msf1.TextMatrix(0, 2) = "Haber"
msf1.TextMatrix(0, 3) = "Desc. Cuenta"
msf1.TextMatrix(0, 4) = ""

For i = 0 To 1
 msf1.ColAlignment(i) = 1 'izq
Next i

For i = 2 To 3
 msf1.ColAlignment(i) = 9 'izq
Next i

End Sub

Private Sub Form_Load()
Call barracgr(Me)
Call armagrid
Option1 = True
Option3 = True
Option6 = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload abm_asientos
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
    If Option1 Then
      t = "Detallado"
    Else
      t = "Resumido"
    End If
    Call imprimegrid(msf1, c(), "Balance de Sumas y Saldos", "", "Periodo......: " & t_f1 & "  " & t_f2, "Tipo.........:" & t, 85, 7, True, False)
  End If
    
    
End If
End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 'On Error GoTo e1
 'If Val(msf1.TextMatrix(msf1.Row, 0)) > 0 Then
 'Call nivel_acceso(1)
 'If para.id_grupo_modulo_actual >= 5 Then
 '  Load abm_asientos0
 '  abm_asientos0.t_f3 = msf1.TextMatrix(msf1.Row, 2)
 '  abm_asientos0.t_f4 = msf1.TextMatrix(msf1.Row, 2)
 '  abm_asientos0.limpia
 '  abm_asientos0.Show
 'Else
 '  Call sinpermisos
 'End If
End If

Exit Sub
e1:
 Exit Sub
'End If
End Sub

Private Sub t_f1_GotFocus()
t_f1 = ""

End Sub

Private Sub t_f1_LostFocus()
If t_f1 <> "" Then
 If Not IsDate(t_f1) Then
   t_f1 = ""
 End If
End If
 
End Sub

Private Sub t_f2_GotFocus()
t_f2 = ""

End Sub

Private Sub t_f2_LostFocus()
If t_f2 <> "" Then
 If Not IsDate(t_f2) Then
   t_f2 = ""
 End If
End If

End Sub


