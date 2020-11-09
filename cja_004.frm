VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cja_detallemov2 
   BackColor       =   &H00E0E0E0&
   Caption         =   "CASH FLOW"
   ClientHeight    =   8715
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   11970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Meses"
      Height          =   855
      Left            =   5280
      TabIndex        =   17
      Top             =   0
      Width           =   1335
      Begin VB.TextBox t_meses 
         Height          =   495
         Left            =   120
         MaxLength       =   2
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   240
         Width           =   735
      End
      Begin MSComCtl2.UpDown UpDown3 
         Height          =   495
         Left            =   960
         TabIndex        =   19
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   873
         _Version        =   393216
         Enabled         =   -1  'True
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   7080
      TabIndex        =   13
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton Command1 
         Caption         =   "Calcular"
         Height          =   255
         Left            =   3600
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox t_saldoant 
         Height          =   435
         Left            =   1440
         TabIndex        =   15
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Saldo Anterior"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox t_ultimodia 
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   7800
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox t_primerdia 
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   7320
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Periodo"
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   0
      Width           =   4695
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   495
         Left            =   4320
         TabIndex        =   12
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   873
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   495
         Left            =   2280
         TabIndex        =   11
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   873
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.TextBox T_FECHA2 
         Height          =   435
         Left            =   2880
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         Height          =   450
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Periodo Desde:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   1
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "cja_004.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "cja_004.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   2
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
      TabIndex        =   0
      Top             =   8460
      Width           =   11970
      _ExtentX        =   21114
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   13229
            MinWidth        =   13229
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
            TextSave        =   "09:41"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5415
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9551
      _Version        =   393216
      BackColorBkg    =   12632256
      AllowBigSelection=   0   'False
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
End
Attribute VB_Name = "cja_detallemov2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim o As Integer
Dim gcol As Integer
Dim gii, gfi, fti, fte, ftr As Integer  'intervalo ingresos

Private Sub btnacepta_Click()

 
Call limpia
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub




Sub limpia()
  Call armagrid
  espere.Show
  espere.Label1 = "ESPERE........  [Generando Informe]"
  espere.Refresh
  Call opcion1
  Unload espere
  msf1.SetFocus
  End Sub

Sub opcion1()
   
  q = "SELECT * FROM cyb_05, C_01 WHERE  [ID_cuenta_contra] = [ID_cuenta] "
  q = q & " and datevalue([fecha]) >= datevalue('" & t_primerdia & "') and datevalue([fecha]) <= datevalue('" & t_ultimodia & "') and [ubicacion] = 'D'"
  q = q & " order by [id_cuenta_contra], [fecha]"
  Set rs = New ADODB.Recordset
  'MsgBox (q)
  rs.Open q, cn1
  pasada = 0
  'ingresos
  ingresos = 0
  F = 1
  cuenta = 0
  msf1.AddItem "*** INGRESOS *** "
  While Not rs.EOF
    If cuenta <> rs("id_cuenta_contra") Then
      cuenta = rs("id_cuenta_contra")
      pasada = 1
      msf1.AddItem rs("c_01.descripcion")
      F = msf1.Rows - 1
      For i = 1 To Val(t_meses) + 2
        msf1.TextMatrix(F, i) = ""
      Next i
       'imprimo
    Else
       p = Format$(Month(rs("fecha")), "00") & "/" & Format$(Year(rs("fecha")), "0000")
       For i = 1 To 12
         If p = msf1.TextMatrix(0, i) Then
             msf1.TextMatrix(F, i) = Format$(Val(msf1.TextMatrix(F, i)) + rs("cyb_05.importe"), "######0.00")
             i = 14
         End If
       Next i
       rs.MoveNext
    End If
    
  Wend
  Set rs = Nothing
 ' gii = 2
 ' gfi = msf1.Rows - 1
  
  'CALCULA TOTALES POR CUENTA y Promedio
  If msf1.Rows >= 2 Then
    For i = 2 To msf1.Rows - 1
      tc = 0
      For J = 1 To Val(t_meses)
        tc = tc + Val(msf1.TextMatrix(i, J))
      Next J
      msf1.TextMatrix(i, Val(t_meses) + 1) = Format$(tc, "######0.00")
      msf1.TextMatrix(i, Val(t_meses) + 2) = Format$(tc / Val(t_meses), "######0.00")
    Next i
  End If
  
  
  'CALCULA TOTALES POR periodo
  linea = "--------------------"
  If msf1.Rows >= 2 Then
    msf1.AddItem ""
    For i = 1 To Val(t_meses) + 1
      msf1.TextMatrix(msf1.Rows - 1, i) = linea
    Next i
    msf1.AddItem "***TOTAL INGRESOS***"
    
    For i = 1 To Val(t_meses) + 1
      tc = 0
      For J = 2 To msf1.Rows - 3
        tc = tc + Val(msf1.TextMatrix(J, i))
      Next J
      msf1.TextMatrix(msf1.Rows - 1, i) = Format$(tc, "######0.00")
    Next i
    msf1.TextMatrix(msf1.Rows - 1, Val(t_meses) + 2) = Format$(Val(msf1.TextMatrix(msf1.Rows - 1, Val(t_meses) + 1)) / Val(t_meses), "######0.00")

    fti = msf1.Rows - 1
   
  Else
    fti = 0
  
  End If
    
  'intervalo de ingresos
  gii = 2
  gfi = msf1.Rows - 3
  
  
'EGRESOS
  q = "SELECT * FROM cyb_05, C_01 WHERE  [ID_cuenta_contra] = [ID_cuenta] "
  q = q & " and datevalue([fecha]) >= datevalue('" & t_primerdia & "') and datevalue([fecha]) <= datevalue('" & t_ultimodia & "') and [ubicacion] = 'H'"
  q = q & " order by [id_cuenta_contra], [fecha]"
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  pasada = 0
  egresos = 0
  
  cuenta = 0
  msf1.AddItem ""
  msf1.AddItem ""
  msf1.AddItem "=== EGRESOS === "
  FCE = msf1.Rows
  While Not rs.EOF
    If cuenta <> rs("id_cuenta_contra") Then
      cuenta = rs("id_cuenta_contra")
      pasada = 1
      msf1.AddItem rs("c_01.descripcion")
      F = msf1.Rows - 1
      For i = 1 To Val(t_meses) + 1
        msf1.TextMatrix(F, i) = ""
      Next i
       'imprimo
    Else
       p = Format$(Month(rs("fecha")), "00") & "/" & Format$(Year(rs("fecha")), "0000")
       For i = 1 To 12
         If p = msf1.TextMatrix(0, i) Then
             msf1.TextMatrix(F, i) = Format$(Val(msf1.TextMatrix(F, i)) + rs("cyb_05.importe"), "######0.00")
             i = 14
         End If
       Next i
       rs.MoveNext
    End If
    
  Wend
  Set rs = Nothing
  
  'CALCULA TOTALES POR CUENTA
  If FCE < msf1.Rows - 1 Then
    For i = FCE To msf1.Rows - 1
      tc = 0
      For J = 1 To Val(t_meses)
        tc = tc + Val(msf1.TextMatrix(i, J))
      Next J
      msf1.TextMatrix(i, Val(t_meses) + 1) = Format$(tc, "######0.00")
      msf1.TextMatrix(i, Val(t_meses) + 2) = Format$(tc / Val(t_meses), "######0.00")
    Next i
  
    'CALCULA TOTALES POR periodo
    linea = "--------------------"
    msf1.AddItem ""
    For i = 1 To Val(t_meses) + 1
      msf1.TextMatrix(msf1.Rows - 1, i) = linea
    Next i
    msf1.AddItem "===TOTAL EGRESOS==="
    
    For i = 1 To Val(t_meses) + 1
      tc = 0
      For J = FCE To msf1.Rows - 3
        tc = tc + Val(msf1.TextMatrix(J, i))
      Next J
      msf1.TextMatrix(msf1.Rows - 1, i) = Format$(tc, "######0.00")
    Next i
    msf1.TextMatrix(msf1.Rows - 1, Val(t_meses) + 2) = Format$(Val(msf1.TextMatrix(msf1.Rows - 1, Val(t_meses) + 1)) / Val(t_meses), "######0.00")
    
    fte = msf1.Rows - 1
  
  Else
   fte = 0
  End If
  
  msf1.AddItem ""
  msf1.AddItem "xxx RESULTADO xxx"
  For i = 1 To Val(t_meses) + 1
      msf1.TextMatrix(msf1.Rows - 1, i) = linea
  Next i
  
  For i = 1 To Val(t_meses) + 1
    If fti > 0 Then
      ti = Val(msf1.TextMatrix(fti, i))
    Else
      ti = 0
    End If
    
    If fte > 0 Then
      te = Val(msf1.TextMatrix(fte, i))
    Else
      te = 0
    End If
     
    r = Format$(ti - te, "######0.00")
    msf1.TextMatrix(msf1.Rows - 1, i) = r
    
  Next i
  ftr = msf1.Rows - 1
  msf1.TextMatrix(msf1.Rows - 1, Val(t_meses) + 2) = Format$(Val(msf1.TextMatrix(msf1.Rows - 1, Val(t_meses) + 1)) / Val(t_meses), "######0.00")

  
  msf1.AddItem "Saldo Anterior ---->"
  msf1.TextMatrix(msf1.Rows - 1, 1) = Format$(Val(t_saldoant), "#####0.00")
  msf1.AddItem "xxx ACUMULADO xxx"
  a = Val(t_saldoant)
  For i = 1 To Val(t_meses)
      a = a + Val(msf1.TextMatrix(msf1.Rows - 3, i))
      msf1.TextMatrix(msf1.Rows - 1, i) = Format$(a, "######0.00")
  Next i
  msf1.Refresh
  
End Sub




Sub armagrid()
'armar grilla
gcol = Val(t_meses) + 3
msf1.clear
msf1.Rows = 1
msf1.Cols = gcol
'msf1.FixedCols = 0
'msf1.SelectionMode = flexSelectionByRow
msf1.FocusRect = flexFocusNone
msf1.ColWidth(0) = 2500
For i = 1 To Val(t_meses)
msf1.ColWidth(i) = 900
Next i
'msf1.ColWidth(2) = 800
'msf1.ColWidth(3) = 800
'msf1.ColWidth(4) = 800
'msf1.ColWidth(5) = 800
'msf1.ColWidth(6) = 800
'msf1.ColWidth(7) = 800
'msf1.ColWidth(8) = 800
'msf1.ColWidth(9) = 800
'msf1.ColWidth(10) = 800
'msf1.ColWidth(11) = 800
'msf1.ColWidth(12) = 800
msf1.ColWidth(Val(t_meses) + 1) = 1000
msf1.ColWidth(Val(t_meses) + 2) = 1000

msf1.TextMatrix(0, 0) = "Concepto"
msf1.TextMatrix(0, Val(t_meses) + 1) = "Totales"
msf1.TextMatrix(0, Val(t_meses) + 2) = "Promedios"
mp = Val(t_fecha)
ap = Val(t_fecha2)
For i = 1 To Val(t_meses) + 1
  If i <= Val(t_meses) Then
   msf1.TextMatrix(0, i) = Format$(mp, "00") & "/" & Format$(ap, "00")
  Else
   t_ultimodia = Format$(DateSerial(ap, mp, 1) - 1, "dd/mm/yyyy")
  End If
  mp = mp + 1
  If mp > 12 Then
    mp = 1
    ap = ap + 1
  End If
Next i
t_primerdia = "01/" & msf1.TextMatrix(0, 1)

End Sub




Private Sub Command1_Click()
F = "01/" & t_fecha & "/" & t_fecha2
t_saldoant = Format$(saldoanterior(F, 0, 0), "######0.00")

End Sub

Private Sub Form_Load()
Call barraesag(Me)
t_fecha = "01"
t_fecha2 = Format$(Now, "yyyy")
t_meses = 12
Call armagrid

End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F7] Imprime - [F10] Grafica - [F11] EXCEL"

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)


If KeyCode = vbKeyF7 Then
Dim c(15) As Double
Call nivel_acceso(3)
If para.id_grupo_modulo_actual >= 5 Then
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
   For i = 0 To Val(t_meses) + 2
    c(i) = i
   Next i
   For i = Val(t_meses) + 3 To 14
      c(i) = -1
    Next i
    
    t1 = "Periodo desde:......: " & t_fecha & " / " & t_fecha2
    T2 = "Meses...............: " & t_meses
    t3 = "Saldo Inicial.......: " & Format$(t_saldoant, "#######0.00")
    
    Call imprimegrid(msf1, c(), "CASH FLOW", T2, t1, t3, 55, 8, True, False, "H")
  End If
Else
 Call sinpermisos
End If
End If


If KeyCode = vbKeyF10 Then
  J = MsgBox("Confirma armado de grafica", 4)
  If J = 6 Then
   Load gen_graficaresultado
   gen_graficaresultado.MSChart1.Title = "Flujo de Fondos"
   gen_graficaresultado.MSChart1.chartType = VtChChartType2dLine
   
   X = InputBox$("Seleciona Tipo: [1] Solo resultado   -  [2] Solo E-I    -   [3] Todo]   -   [0]salir", "Tipo de Grafico", "1")
   Select Case Val(X)
    Case Is = 3
          gen_graficaresultado.MSChart1.DataGrid.ColumnCount = 3
          gen_graficaresultado.MSChart1.DataGrid.RowCount = msf1.Cols - 3
          gen_graficaresultado.MSChart1.DataGrid.ColumnLabel(1, 1) = "Ingresos"
          gen_graficaresultado.MSChart1.DataGrid.ColumnLabel(2, 1) = "Egresos"
          gen_graficaresultado.MSChart1.DataGrid.ColumnLabel(3, 1) = "Resultado"
   
          For i = 1 To msf1.Cols - 3
              gen_graficaresultado.MSChart1.DataGrid.RowLabel(i, 1) = msf1.TextMatrix(0, i)
              gen_graficaresultado.MSChart1.DataGrid.SetData i, 1, msf1.TextMatrix(fti, i), 0
              gen_graficaresultado.MSChart1.DataGrid.SetData i, 2, msf1.TextMatrix(fte, i), 0
              gen_graficaresultado.MSChart1.DataGrid.SetData i, 3, msf1.TextMatrix(ftr, i), 0
          Next i
          gen_graficaresultado.Show
    Case Is = 2
          gen_graficaresultado.MSChart1.DataGrid.ColumnCount = 2
          gen_graficaresultado.MSChart1.DataGrid.RowCount = msf1.Cols - 3
          gen_graficaresultado.MSChart1.DataGrid.ColumnLabel(1, 1) = "Ingresos"
          gen_graficaresultado.MSChart1.DataGrid.ColumnLabel(2, 1) = "Egresos"
          
          For i = 1 To msf1.Cols - 3
              gen_graficaresultado.MSChart1.DataGrid.RowLabel(i, 1) = msf1.TextMatrix(0, i)
              gen_graficaresultado.MSChart1.DataGrid.SetData i, 1, msf1.TextMatrix(fti, i), 0
              gen_graficaresultado.MSChart1.DataGrid.SetData i, 2, msf1.TextMatrix(fte, i), 0
           Next i
          gen_graficaresultado.Show
   Case Is = 1
          gen_graficaresultado.MSChart1.DataGrid.ColumnCount = 1
          gen_graficaresultado.MSChart1.DataGrid.RowCount = msf1.Cols - 3
          gen_graficaresultado.MSChart1.DataGrid.ColumnLabel(1, 1) = "Resultado"
          
          For i = 1 To msf1.Cols - 3
              gen_graficaresultado.MSChart1.DataGrid.RowLabel(i, 1) = msf1.TextMatrix(0, i)
              gen_graficaresultado.MSChart1.DataGrid.SetData i, 1, msf1.TextMatrix(ftr, i), 0
          Next i
          gen_graficaresultado.Show
   End Select
  
  End If
End If


If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If

End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If msf1.Row >= gii And msf1.Row <= gfi Then
  h = InputBox$("Ingrese nuevo valor para la celda", "MODIFICACION CASH FLOW", msf1.TextMatrix(msf1.Row, msf1.col))
  msf1.TextMatrix(msf1.Row, msf1.col) = Format$(Val(h), "######0.00")
  Call CALCULATOTALES
 End If
End If
End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
End Sub


Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Val(t_fecha) < 1 Or Val(t_fecha) > 12 Then
    t_fecha = "01"
  End If
Else
   t_fecha = "01"
End If
End Sub

Private Sub t_fecha2_LostFocus()
If t_fecha2 <> "" Then
  If Val(t_fecha2) < 2000 Or Val(t_fecha2) > 2100 Then
    t_fecha2 = Format$(Now, "yyyy")
  End If
End If
End Sub




Private Sub t_meses_LostFocus()
If Val(t_meses) < 1 Or Val(t_meses) > 12 Then
  t_meses = 12
End If
End Sub

Private Sub t_saldoant_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 1)
End Sub

Private Sub t_saldoant_LostFocus()
t_saldoant = Format$(Val(t_saldoant), "######0.00")
End Sub

Private Sub UpDown1_DownClick()
If Val(t_fecha) > 1 Then
  t_fecha = Format$(Val(t_fecha) - 1, "00")
Else
 t_fecha = "12"
End If
End Sub

Private Sub UpDown1_UpClick()
If Val(t_fecha) < 12 Then
  t_fecha = Format$(Val(t_fecha) + 1, "00")
Else
 t_fecha = "01"
End If

End Sub

Private Sub UpDown2_DownClick()
  t_fecha2 = Format$(Val(t_fecha2) - 1, "00")
End Sub

Private Sub UpDown2_UpClick()
 t_fecha2 = Format$(Val(t_fecha2) + 1, "00")
End Sub

Private Sub UpDown3_DownClick()
If Val(t_meses) > 1 Then
  t_meses = Val(t_meses) - 1
End If
End Sub

Private Sub UpDown3_UpClick()
If Val(t_meses) < 12 Then
  t_meses = Val(t_meses) + 1
End If

End Sub
Sub CALCULATOTALES()
'CALCULA TOTALES POR CUENTA y Promedio
  If gii >= 2 Then
    For i = 2 To gfi
      tc = 0
      For J = 1 To Val(t_meses)
        tc = tc + Val(msf1.TextMatrix(i, J))
      Next J
      msf1.TextMatrix(i, Val(t_meses) + 1) = Format$(tc, "######0.00")
      msf1.TextMatrix(i, Val(t_meses) + 2) = Format$(tc / Val(t_meses), "######0.00")
    Next i
  End If
  
  
  If gii >= 2 Then
    For i = 1 To Val(t_meses) + 1
      tc = 0
      For J = 2 To gfi
        tc = tc + Val(msf1.TextMatrix(J, i))
      Next J
      msf1.TextMatrix(gfi + 2, i) = Format$(tc, "######0.00")
    Next i
   ' msf1.TextMatrix(gfi, Val(t_meses) + 2) = Format$(Val(msf1.TextMatrix(gfi, Val(t_meses) + 1)) / Val(t_meses), "######0.00")
  End If
End Sub
