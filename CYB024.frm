VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form cyb_cuenta_ch_vend 
   BackColor       =   &H00E0E0E0&
   Caption         =   "CHEQUES VENDIDOS PENDIENTES DE ACREDITACION"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   360
      TabIndex        =   11
      Top             =   7080
      Width           =   9135
      Begin VB.Label Label4 
         Caption         =   $"CYB024.frx":0000
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   8655
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   5040
      TabIndex        =   10
      Top             =   120
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   115539969
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
      Height          =   1335
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   7215
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
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Corte:"
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
         Picture         =   "CYB024.frx":00A0
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
         Picture         =   "CYB024.frx":0922
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
            TextSave        =   "29/04/2022"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "04:27 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "cyb_cuenta_ch_vend"
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
  q = "select [fecha], [num_cheque], cyb_03.[fecha_dif], [titular], [origen], cyb_03.[importe]  FROM CYB_03, cyb_04 where [estado] = 'V' and cyb_03.[num_mov_banco_e] = cyb_04.[num_mov_banco] "
  q = q & " and [id_banco] = " & c_banco.ItemData(c_banco.ListIndex)
  
  c = " and "
  If t_fecha <> "" Then
      q = q & c & " datevalue(cyb_03.[fecha_dif]) >= datevalue('" & t_fecha & "')"
  End If
    
  q = q & " order by cyb_03.[fecha_dif]"
  
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  While Not rs.EOF
      sa = sa + rs("importe")
      FV2 = rs("fecha")
      fd = rs("fecha_dif")
      nc = rs("num_cheque")
      t = rs("titular")
      o = rs("origen")
      i = Format$(rs("importe"), "#######0.00")
      msf1.AddItem FV2 & Chr(9) & nc & Chr(9) & fd & Chr(9) & t & Chr(9) & o & Chr(9) & i & Chr(9) & Format$(sa, "#######0.00")
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
msf1.Cols = 8
msf1.ColWidth(0) = 1100
msf1.ColWidth(1) = 1000
msf1.ColWidth(2) = 1100
msf1.ColWidth(3) = 2600
msf1.ColWidth(4) = 3000
msf1.ColWidth(5) = 1100
msf1.ColWidth(6) = 1300
msf1.ColWidth(7) = 200
msf1.TextMatrix(0, 0) = "Fecha Vta."
msf1.TextMatrix(0, 1) = "Nro.Ch"
msf1.TextMatrix(0, 2) = "Fecha Dif."
msf1.TextMatrix(0, 3) = "Titular"
msf1.TextMatrix(0, 4) = "Origen"
msf1.TextMatrix(0, 5) = "Importe"
msf1.TextMatrix(0, 6) = "Saldo"
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
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F7] Imprime - "
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

