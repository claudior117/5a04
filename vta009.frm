VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_saldoscli 
   BackColor       =   &H00E0E0E0&
   Caption         =   "SALDOS CLIENTES"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame12 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Formato resumen de cuenta"
      Height          =   615
      Left            =   240
      TabIndex        =   40
      Top             =   7440
      Width           =   7575
      Begin VB.OptionButton Option8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Impresion Normal"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Impresion Apaisada"
         Height          =   255
         Left            =   2640
         TabIndex        =   42
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Impresion Media Hoja"
         Height          =   255
         Left            =   5280
         TabIndex        =   41
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   8040
      TabIndex        =   37
      Top             =   1800
      Width           =   2295
      Begin VB.CheckBox Check6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Muestra Saldo a Favor"
         Height          =   315
         Left            =   120
         TabIndex        =   38
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipo Cliente"
      Height          =   855
      Left            =   6960
      TabIndex        =   34
      Top             =   6600
      Width           =   3015
      Begin VB.CheckBox Check5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sin Clientes Nacionales"
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sin Clientes Exportaciones"
         Height          =   375
         Left            =   1440
         TabIndex        =   35
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Comprobante"
      Height          =   855
      Left            =   3960
      TabIndex        =   31
      Top             =   6600
      Width           =   2175
      Begin VB.CheckBox Check3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Con Detalle"
         Height          =   375
         Left            =   1200
         TabIndex        =   33
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sin Detalle"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipo informe"
      Height          =   855
      Left            =   240
      TabIndex        =   28
      Top             =   6600
      Width           =   3375
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha Comprobante"
         Height          =   495
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha vencimiento"
         Height          =   495
         Left            =   1680
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   1335
      Left            =   360
      TabIndex        =   21
      Top             =   840
      Width           =   7575
      Begin VB.ComboBox c_tiposaldo 
         Height          =   315
         ItemData        =   "vta009.frx":0000
         Left            =   1440
         List            =   "vta009.frx":000D
         TabIndex        =   26
         Top             =   960
         Width           =   2775
      End
      Begin VB.ComboBox c_VEND 
         Height          =   315
         Left            =   1440
         TabIndex        =   25
         Top             =   600
         Width           =   5055
      End
      Begin VB.TextBox t_cliente 
         Height          =   285
         Left            =   1440
         TabIndex        =   23
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   "Tipo Saldos:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "Vendedor:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Cliente:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Orden"
      Height          =   735
      Left            =   3960
      TabIndex        =   17
      Top             =   120
      Width           =   3975
      Begin VB.OptionButton Option5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Saldo"
         Height          =   195
         Left            =   2880
         TabIndex        =   39
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Razon Social"
         Height          =   195
         Left            =   1320
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Id. Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   8040
      TabIndex        =   14
      Top             =   1320
      Width           =   2295
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Muestra Saldo en U$s"
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   2055
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   4440
      TabIndex        =   13
      Top             =   2040
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   106954753
      CurrentDate     =   38803
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Muestra Saldos en Cero"
      Height          =   615
      Left            =   8040
      TabIndex        =   10
      Top             =   720
      Width           =   2295
      Begin VB.OptionButton O_cero 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Si"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Tag             =   "P"
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton O_nocero 
         BackColor       =   &H00E0E0E0&
         Caption         =   "No"
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Tag             =   "D"
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fecha Desde - Hasta"
      Height          =   735
      Left            =   360
      TabIndex        =   9
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
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Moneda"
      Height          =   615
      Left            =   8040
      TabIndex        =   6
      Top             =   120
      Width           =   2295
      Begin VB.OptionButton O_dolares 
         BackColor       =   &H00E0E0E0&
         Caption         =   "U$s"
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Tag             =   "D"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton O_pesos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pesos"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Tag             =   "P"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10200
      TabIndex        =   3
      Top             =   7080
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta009.frx":0030
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
         Picture         =   "vta009.frx":08B2
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
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   15875
            MinWidth        =   15875
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "07/04/2022"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:35 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   4095
      Left            =   240
      TabIndex        =   16
      Top             =   2400
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7223
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   10560
      TabIndex        =   20
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "vta_saldoscli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private EXISTE As String
Private saldoant As Double
Private saldoact As Double
'FIXIT: Declare 'saf' and 'df' and 'hf' and 'sf' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim saf, df, hf, sf, sof As Double




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
QUERY = "select * from VTA_01 where [id_cliente] > 1"
X = " and "

If t_cliente <> "" Then
  QUERY = QUERY & X & " [denominacion]  like '%" & t_cliente & "%'"
  X = " and "
End If

If c_vend.ListIndex > 0 Then
  QUERY = QUERY & X & " [id_vendedor] = " & c_vend.ItemData(c_vend.ListIndex)
  X = " and "
End If

If c_tiposaldo.ListIndex > 0 Then
  If Mid$(c_tiposaldo, 1, 1) = "I" Then
     m2 = "S"
  Else
     m2 = "N"
  End If
  QUERY = QUERY & X & " [saldo_incobrable] = '" & m2 & "'"
  X = " and "
End If

If Check4 = 1 Then
   QUERY = QUERY & X & " [id_tipoiva] <> 8"
   X = " and "
End If

If Check5 = 1 Then
   QUERY = QUERY & X & " [id_tipoiva] = 8"
   X = " and "
End If

If Option1 = True Then
  QUERY = QUERY & " order by [id_cliente]"
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
   Set cl_cli = New Clientes
   cl_cli.carga (rs1("id_CLIENTE"))
   If Option4 = True Then
     saldoant = cl_cli.saldo(False, t_fecha, O_pesos)
     saldoact = cl_cli.saldoentrefechas(t_fecha, t_fecha2, O_pesos)
   Else
     saldoant = cl_cli.saldov(False, t_fecha, O_pesos)
     saldoact = cl_cli.saldoentrefechasv(t_fecha, t_fecha2, O_pesos)
   End If
   If Val(Format$(saldoact, "######0.00")) = 0 Then
     If O_cero = True Then
        Call agrega(r)
        r = r + 1
     End If
   Else
    If Val(Format$(saldoact, "######0.00")) < 0 Then
     If Check6 = 1 Then
       Call agrega(r)
       r = r + 1
     End If
    Else
      Call agrega(r)
      r = r + 1
    End If
   End If
   Set cl_cli = Nothing
   rs1.MoveNext
   pb = pb + 1
   Label5 = pb
   Label5.Refresh
  Wend
  If Option5 = True Then
     
    msf1.col = 5 'Desde que columna iniciar la ordenacion
    msf1.ColSel = 5 'Hasta que columna terminar la ordenacion

    msf1.Row = 1 'Primer renglon del MsFlex a sortear
    msf1.RowSel = msf1.Rows - 1 'Ultimo renglon del msflex a sortear

    msf1.Sort = 4 'metodo de sorteo deseado Numerico descendente
  End If
  
  msf1.Refresh
  If Check1 = 0 Then
      msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & "________________" & Chr$(9) & "________________" & Chr$(9) & "________________" & Chr$(9) & "________________"
      msf1.AddItem "" & Chr$(9) & "Total Clientes: " & r & Chr$(9) & Format$(saf, "#####0.00") & Chr$(9) & Format$(df, "######0.00") & Chr$(9) & Format$(hf, "######0.00") & Chr$(9) & Format$(sf, "######0.00")
  Else
      msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & "________________" & Chr$(9) & "________________" & Chr$(9) & "________________" & Chr$(9) & "________________" & Chr$(9) & "________________"
      msf1.AddItem "" & Chr$(9) & "Total Clientes: " & r & Chr$(9) & Format$(saf, "#####0.00") & Chr$(9) & Format$(df, "######0.00") & Chr$(9) & Format$(hf, "######0.00") & Chr$(9) & Format$(sf, "######0.00") & Chr$(9) & Format$(sof, "######0.00")
  End If

  Unload espere
End If
Set rs1 = Nothing
Set rs2 = Nothing


End Sub






Sub agrega(r As Integer)
    saf = saf + saldoant
    df = df + cl_cli.DEBE
    hf = hf + cl_cli.HABER
    If Option4 = True Then
      sf = sf + cl_cli.saldocli
    Else
      sf = sf + cl_cli.saldocliv
    End If
    If Check1 = 0 Then
      msf1.AddItem cl_cli.id & Chr$(9) & cl_cli.razonsocial & Chr$(9) & Format$(saldoant, "#####0.00") & Chr$(9) & Format$(cl_cli.DEBE, "#####0.00") & Chr$(9) & Format$(cl_cli.HABER, "#####0.00") & Chr$(9) & Format$(cl_cli.saldocli, "######0.00")
    Else
        d = cl_cli.DEBE
        h = cl_cli.HABER
        
        If O_pesos Then
         If Option4 = True Then
          s = cl_cli.saldocli
          so = cl_cli.saldo(True, t_fecha2, False)
         Else
           s = cl_cli.saldocliv
           so = cl_cli.saldov(True, t_fecha2, False)
         End If
        Else
         If Option4 = True Then
          s = cl_cli.saldocli
          so = cl_cli.saldo(True, t_fecha2, True)
         Else
          s = cl_cli.saldocliv
          so = cl_cli.saldov(True, t_fecha2, True)
         End If
        End If
        msf1.AddItem cl_cli.id & Chr$(9) & cl_cli.razonsocial & Chr$(9) & Format$(saldoant, "#####0.00") & Chr$(9) & Format$(d, "######0.00") & Chr$(9) & Format$(h, "######0.00") & Chr$(9) & Format$(s, "######0.00") & Chr$(9) & Format$(so, "######0.00")
        sof = sof + so
    End If
    If Check2 = 1 Then
       Call muestracomp(cl_cli.id)
    End If
End Sub

Sub muestracomp(ByVal ic As Long)
'ic = id cliente
If Option4 = True Then 'fecha comp
    QUERY = "SELECT * FROM VTA_02 where [id_CLIENTE] = " & ic
    QUERY = QUERY & " and datevalue(fecha) >= " & "DateValue('" & t_fecha & "') "
    QUERY = QUERY & " and datevalue(fecha) <= " & "DateValue('" & t_fecha2 & "') "
    QUERY = QUERY & " and  [cta_cte] <> " & "'N'" & " and [contado] = " & "'N'"
Else 'fecha vto
    QUERY = "SELECT * FROM VTA_02 where [id_CLIENTE] = " & ic
    QUERY = QUERY & " and datevalue(fecha_vto) >= " & "DateValue('" & t_fecha & "') "
    QUERY = QUERY & " and datevalue(fecha_vto) <= " & "DateValue('" & t_fecha2 & "') "
    QUERY = QUERY & " and  [cta_cte] <> " & "'N'" & " and [contado] = " & "'N'"
End If
Set rs2 = New ADODB.Recordset
rs2.Open QUERY, cn1
While Not rs2.EOF
  Set cl_compvta = New comprobantes_venta
  cl_compvta.sucursal = rs2("sucursal_ingreso")
  cl_compvta.actual (rs2("id_tipocomp"))
  comp = cl_compvta.abreviatura
  Set cl_compvta = Nothing
  comp = comp & " " & Format$(rs2("sucursal"), "0000") & "-" & Format$(rs2("num_comp"), "00000000")
  If O_pesos = True Then
    mm = "P"
  Else
    mm = "D"
  End If
  If rs2("cta_cte") = "D" Then
   If rs2("moneda") = mm Then
     d = Format$(rs2("total"), "######0.00")
   Else
     d = Format$(rs2("total_otra_moneda"), "######0.00")
   End If
   h = ""
  Else
   If rs2("moneda") = mm Then
      h = Format$(rs2("total"), "######0.00")
   Else
      h = Format$(rs2("total_otra_moneda"), "######0.00")
   End If
   d = ""
  End If
  msf1.AddItem "" & Chr$(9) & "  --->   " & comp & Chr$(9) & "" & Chr$(9) & d & Chr$(9) & h
  If Check3 = 1 Then
    If rs2("id_tipocomp") <> 50 Then
     Call muestradetalle(rs2("num_int"))
    Else
     Call muestradetallerbo(rs2("num_int"))
    End If
  End If
 rs2.MoveNext
Wend
Set rs2 = Nothing
End Sub

Sub muestradetalle(ByVal numint)
q = "select * from vta_03 where [num_int] = " & numint
Set rs4 = New ADODB.Recordset
rs4.Open q, cn1
While Not rs4.EOF
   det = Format$(Left$(rs4("descripcion"), 40), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!") & "(" & Format$(rs4("cantidad"), "####0.00") & rs4("tunidad") & ")"
   msf1.AddItem "" & Chr$(9) & "             ++" & det
  
  rs4.MoveNext
Wend
Set rs4 = Nothing
End Sub

Sub muestradetallerbo(ByVal numint)
q = "select * from vta_010, vta_02 where [num_int_rbo] = " & numint & " and [num_int_comp] = [num_int]"
Set rs4 = New ADODB.Recordset
rs4.Open q, cn1
While Not rs4.EOF
   det = Format$(rs4("sucursal"), "0000") & Format$(rs4("num_comp"), "00000000") & "      (" & Format$(rs4("importe_pagado"), "#####0.00") & ")"
   msf1.AddItem "" & Chr$(9) & "             **" & det
  
  rs4.MoveNext
Wend
Set rs4 = Nothing
End Sub

Private Sub btnsale_Click()

Unload Me
End Sub




Private Sub c_tiposaldo_LostFocus()
If c_tiposaldo.ListIndex < 0 Then
  c_tiposaldo.ListIndex = 0
End If
End Sub

Private Sub c_vend_LostFocus()
If c_vend.ListIndex < 0 Then
  c_vend.ListIndex = 0
End If
End Sub

Private Sub cal1_LostFocus()
cal1.Visible = False
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyUp
     Call tabup(Me)
End Select

End Sub
Sub armagrid()
'armar grilla
If Check1 = 0 Then
  msf1.clear
  msf1.Rows = 1
  msf1.Cols = 8
  msf1.ColWidth(0) = 500
  msf1.ColWidth(1) = 5000
  msf1.ColWidth(2) = 1300
  msf1.ColWidth(3) = 1300
  msf1.ColWidth(4) = 1300
  msf1.ColWidth(5) = 1300
  msf1.ColWidth(6) = 500
  msf1.ColWidth(7) = 200
    msf1.TextMatrix(0, 0) = "Id."
  msf1.TextMatrix(0, 1) = "Cliente"
  msf1.TextMatrix(0, 2) = "Saldo Ant."
  If O_pesos Then
   msf1.TextMatrix(0, 3) = "Debe($)"
   msf1.TextMatrix(0, 4) = "Haber($)"
   msf1.TextMatrix(0, 5) = "Saldo($)"
  Else
   msf1.TextMatrix(0, 3) = "Debe(U$s)"
   msf1.TextMatrix(0, 4) = "Haber(U$s)"
   msf1.TextMatrix(0, 5) = "Saldo(U$s)"
  End If
  msf1.TextMatrix(0, 6) = " "
  For i = 0 To 6
    msf1.ColAlignment(i) = 9 'der
  Next i
  msf1.ColAlignment(1) = 1 'izq
    
  
Else
   
  msf1.clear
  msf1.Rows = 1
  msf1.Cols = 8
  msf1.ColWidth(0) = 500
  msf1.ColWidth(1) = 3800
  msf1.ColWidth(2) = 1100
  msf1.ColWidth(3) = 1100
  msf1.ColWidth(4) = 1100
  msf1.ColWidth(5) = 1100
  msf1.ColWidth(6) = 1100
  msf1.ColWidth(7) = 200
  msf1.TextMatrix(0, 0) = "Id."
  msf1.TextMatrix(0, 1) = "Cliente"
  msf1.TextMatrix(0, 2) = "Saldo Ant."
  If O_pesos Then
   msf1.TextMatrix(0, 3) = "Debe($)"
   msf1.TextMatrix(0, 4) = "Haber($)"
   msf1.TextMatrix(0, 5) = "Saldo($)"
   msf1.TextMatrix(0, 6) = "Saldo(U$s)"
  Else
   msf1.TextMatrix(0, 3) = "Debe(U$s)"
   msf1.TextMatrix(0, 4) = "Haber(U$s)"
   msf1.TextMatrix(0, 5) = "Saldo(U$s)"
   msf1.TextMatrix(0, 6) = "Saldo($)"
  End If
  msf1.TextMatrix(0, 7) = " "
  For i = 0 To 7
    msf1.ColAlignment(i) = 9
  Next i
  msf1.ColAlignment(1) = 1
 End If
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

Call carga_vendedores(c_vend)
c_vend.AddItem "<Todos>", 0
c_vend.ListIndex = 0
If para.moneda = "P" Then
  O_pesos = Checked
Else
  O_dolares = Checked
End If
O_nocero = Checked

cal1.Visible = False
Check1 = 0
Check2 = 0
Check3 = 0
Call armagrid

Option1 = True
Option4 = True
Option7 = True
c_tiposaldo.ListIndex = 0
End Sub




Private Sub Form_Unload(Cancel As Integer)
Unload vta_estadocuenta
End Sub



Private Sub cal1_DblClick()
If cal1.Tag = "1" Then
  t_fecha = cal1
  t_fecha.SetFocus
Else
  t_fecha2 = cal1
  t_fecha2.SetFocus
End If
cal1.Visible = False
End Sub



Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(1) = "[F7] Imprime - [F10] Ajuste Ctacte - [F11] Excel -[ENTER] Estado Cuenta - [esp] Selecciona - [F5] Todos -[F2] Resumen Cta "
End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF2 Then
  Call nivel_acceso(1)
  If para.id_grupo_modulo_actual >= 4 Then
   k = MsgBox("Prepare impresora y confirme emitir todos los resumenes de cuenta seleccionados", 4)
   If k = 6 Then
     
     Load vta_estadocuenta
     J = 0
     While J < msf1.Rows - 1
       
        c = Val(msf1.TextMatrix(J, 0))
        If c > 0 And msf1.TextMatrix(J, 6) = "**" Then
           vta_estadocuenta.c_prov.ListIndex = buscaindice(vta_estadocuenta.c_prov, c)
           vta_estadocuenta.t_fecha = t_fecha
           vta_estadocuenta.t_fecha2 = t_fecha2
           vta_estadocuenta.Option1 = Option4
           vta_estadocuenta.Option4 = O_pesos
           'media hoja despues hacer por parametro
           
              vta_estadocuenta.Option5 = Option8
              vta_estadocuenta.Option6 = Option6
              vta_estadocuenta.Option7 = Option7
                        
           vta_estadocuenta.carga
           vta_estadocuenta.imprime
        End If
        J = J + 1
     Wend
     Unload vta_estadocuenta
    End If
   Else
    MsgBox ("No tiene permisos suficientes para esta operacion")
  End If
End If

If KeyCode = vbKeyF5 Then
  If msf1.Rows > 1 Then
    For i = 1 To msf1.Rows - 1
      If msf1.TextMatrix(i, 6) = "**" Then
          msf1.TextMatrix(i, 6) = ""
      Else
         msf1.TextMatrix(i, 6) = "**"
      End If
    Next i
  End If
End If


If KeyCode = vbKeyF7 Then
  Dim c2(15) As Double
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    If O_pesos Then
      m = "Pesos ($)"
    Else
      m = "Dolares (U$s)"
    End If
    
    If c_vend.ListIndex > 0 Then
       v = "Vendedor: " & c_vend
    Else
       v = " "
    End If
      
    If Check1 = 0 Then
      c2(0) = 6
      c2(1) = 0
      c2(2) = 1
      c2(3) = 2
      c2(4) = 3
      c2(5) = 4
      c2(6) = 5

      For i = 7 To 14
        c2(i) = -1
      Next i
      Call imprimegrid(msf1, c2(), "SALDOS CLIENTES", "Periodo: " & t_fecha & " - " & t_fecha2, "Moneda: " & m, v, 72, 8, True, False)
    Else
      c2(0) = 7
      c2(1) = 0
      c2(2) = 1
      c2(3) = 2
      c2(4) = 3
      c2(5) = 4
      c2(6) = 5
      c2(7) = 6
      For i = 8 To 14
        c2(i) = -1
      Next i
      Call imprimegrid(msf1, c2(), "SALDOS CLIENTES", "Periodo: " & t_fecha & " - " & t_fecha2, "Moneda: " & m, v, 72, 8, True, False)
    End If
  End If

End If


If KeyCode = vbKeyF10 Then
      Load vta_ajustesint
      If msf1.Rows > 1 Then
        k = Val(msf1.TextMatrix(msf1.Row, 0))
        If k > 0 Then
         vta_ajustesint.c_prov.ListIndex = buscaindice(vta_ajustesint.c_prov, k)
        End If
      End If
      vta_ajustesint.Show
End If




If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If

End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And msf1.Rows > 1 Then
   c = Val(msf1.TextMatrix(msf1.Row, 0))
   If c > 0 Then
      Load vta_estadocuenta
      vta_estadocuenta.c_prov.ListIndex = buscaindice(vta_estadocuenta.c_prov, c)
      vta_estadocuenta.Show
   End If
End If

If KeyAscii = vbKeySpace Then
  If Val(msf1.TextMatrix(msf1.Row, 0)) > 0 Then
      If msf1.TextMatrix(msf1.Row, 6) = "**" Then
          msf1.TextMatrix(msf1.Row, 6) = ""
      Else
         msf1.TextMatrix(msf1.Row, 6) = "**"
      End If
     End If
  
End If

End Sub

Private Sub O_dolares_Click()
Check1.Caption = "Muestra Saldo en $"
End Sub

Private Sub O_pesos_Click()
Check1.Caption = "Muestra Saldo en U$s"
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
