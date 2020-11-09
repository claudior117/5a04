VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form stk_movprod2 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MOVIMIENTOS DE PRODUCTOS POR FECHA"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12135
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8565
   ScaleWidth      =   12135
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tools"
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   7080
      Width           =   2535
      Begin VB.CommandButton Command1 
         Caption         =   "Ajuste Stock"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   4920
      TabIndex        =   9
      Top             =   120
      Width           =   5055
      Begin VB.ComboBox c_tipo 
         Height          =   315
         ItemData        =   "stk011.frx":0000
         Left            =   1440
         List            =   "stk011.frx":000D
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Tipo Movimiento:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10200
      TabIndex        =   5
      Top             =   7080
      Width           =   1575
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "stk011.frx":002D
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "stk011.frx":08AF
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha"
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3375
      Begin VB.TextBox t_fecha2 
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox t_fecha 
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Fecha hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   8205
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   635
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
            TextSave        =   "27/02/2015"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:39"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5295
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9340
      _Version        =   393216
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
Attribute VB_Name = "stk_movprod2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Sub carga()
 Call armagrid
 espere.Show
 espere.Label1 = "Espere... Procesando reporte"
 espere.Refresh
 Set rs2 = New ADODB.Recordset
 q = "select * from stk_01 where [id_producto] > 1"
 c = " and "
 
 If t_fecha <> "" Then
   q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
   'c = " and "
 End If
 
 If t_fecha2 <> "" Then
   q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
   'c = " and "
 End If
 
 Select Case c_tipo.ListIndex
 Case Is = 1
    'entradas
     q = q & c & "[ubicacion] = 'E'"
   
 Case Is = 2
    'salidas
     q = q & c & "[ubicacion] = 'S'"
    
 End Select
 q = q & " order by [id_producto], [fecha]"
 
 
 rs2.Open q, cn1
 
 If Not rs2.EOF And Not rs2.BOF Then
      Call procesa(rs2)
 End If
 Unload espere

End Sub

Sub procesa(ByVal r As ADODB.Recordset)
cp = r("id_producto")
e = 0
s = 0
sant = 0
c = 0
While Not r.EOF
  espere.Label1 = "Espere.... procesando registro " & c
  If cp = r("id_producto") Then
     If r("ubicacion") = "E" Then
       e = e + r("cantidad")
     Else
       s = s + r("cantidad")
     End If
  Else
     Set cl_prod = New productos
     cl_prod.cargar (cp)
     If t_fecha <> "" Then
        sant = cl_prod.stock_anterior(cp, t_fecha)
     Else
        sant = 0
     End If
     msf1.AddItem cp & Chr$(9) & cl_prod.Detalle & Chr$(9) & Format$(sant, "######0.00") & Chr$(9) & Format$(e, "######0.00") & Chr$(9) & Format$(s, "######0.00") & Chr$(9) & Format$(sant + e - s, "######0.00")
     cp = r("id_producto")
     e = 0
     s = 0
     sant = 0
    If r("ubicacion") = "E" Then
       e = e + r("cantidad")
     Else
       s = s + r("cantidad")
     End If
     Set cl_prod = Nothing
  
  End If
  r.MoveNext
  c = c + 1
Wend
Set cl_prod = New productos
cl_prod.cargar (cp)
If t_fecha <> "" Then
      sant = cl_prod.stock_anterior(cp, t_fecha)
Else
      sant = 0
End If
msf1.AddItem cp & Chr$(9) & cl_prod.Detalle & Chr$(9) & Format$(sant, "######0.00") & Chr$(9) & Format$(e, "######0.00") & Chr$(9) & Format$(s, "######0.00") & Chr$(9) & Format$(sant + e - s, "######0.00")
Set cl_prod = Nothing
End Sub

Private Sub btnacepta_Click()

 Call carga

End Sub

Private Sub btnsale_Click()
Unload Me

End Sub



Private Sub c_tipo_LostFocus()
If c_tipo.ListIndex < 0 Then
  c_tipo.ListIndex = 0
End If

End Sub





Private Sub Command1_Click()
stk_movint.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 1)
  Case Is = 27
        Unload Me
End Select

End Sub
Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 6

msf1.ColWidth(0) = 1200
msf1.ColWidth(1) = 4600
msf1.ColWidth(2) = 1400
msf1.ColWidth(3) = 1400
msf1.ColWidth(4) = 1400
msf1.ColWidth(5) = 1400

msf1.TextMatrix(0, 0) = "Id."
msf1.TextMatrix(0, 1) = "Producto"
msf1.TextMatrix(0, 2) = "Mov.Ant."
msf1.TextMatrix(0, 3) = "Entrada"
msf1.TextMatrix(0, 4) = "Salida"
msf1.TextMatrix(0, 5) = "Resultdo"

For i = 0 To 2
    msf1.ColAlignment(i) = 1 'izq
Next i
For i = 3 To 5
    msf1.ColAlignment(i) = 9 'der
Next i


End Sub

Private Sub Form_Load()
Call barraesag(Me)
Call armagrid

c_tipo.ListIndex = 0

End Sub


Private Sub msf1_GotFocus()
Me.KeyPreview = False
Me.StatusBar1.Panels.Item(2) = "[F7] Imprime - [F11] Excel - [ENTER] Detalle"

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
    Call imprimegrid(msf1, c(), Space$(50) & "MOVIMIENTOS DE PRODUCTOS POR FECHA", "     Periodo.....: " & t_fecha & " - " & t_fceha2, t, " ", 80, 8, True, False, "V")
  End If
End If


If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If
End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    Load stk_movprod
    stk_movprod.t_id = msf1.TextMatrix(msf1.Row, 0)
    stk_movprod.t_fecha = t_fecha
    stk_movprod.Show
  End If
End If
End Sub

Private Sub msf1_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub t_fecha_LostFocus()
Call solofecha(t_fecha)
End Sub



Private Sub t_fecha2_LostFocus()
Call solofecha(t_fecha2)
End Sub
