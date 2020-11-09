VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form com_ivacompras 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IVA COMPRAS"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10545
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6600
   ScaleWidth      =   10545
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   8520
      TabIndex        =   7
      Top             =   7200
      Width           =   3255
      Begin VB.CommandButton Command2 
         Caption         =   "&Salir"
         Height          =   495
         Left            =   1800
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Mostrar"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   12632256
      Appearance      =   1
      StartOfWeek     =   24969217
      CurrentDate     =   38750
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo"
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3375
      Begin VB.TextBox t_fecha2 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   10
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox t_fecha 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Saldo Anterior (Composicion)"
      Height          =   5055
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   11655
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4680
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   11415
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6240
      Width           =   10545
      _ExtentX        =   18600
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
            TextSave        =   "02/02/2006"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "08:02 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "com_ivacompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub carga()
 ct = Space$(10)
 List1.clear
 Call cabecera
 Set cl_prod = New productos
 s = cl_prod.stock_anterior(c_prod.ItemData(c_prod.ListIndex), t_fecha)
 f = Format$(t_fecha, "dd/mm/yyyy")
 c = "S.I. " & " 0000-00000000"
 dE = Format$("Saldo Anterior", "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 RSet ct = Format$(s, "######0.00")
 If s >= 0 Then
    d = ct
    h = "      0.00"
 Else
    h = ct
    d = "      0.00"
 End If
 List1.AddItem f & "  " & c & "  " & dE & "  " & d & "  " & h & "  " & ct
 
 Set rs = New ADODB.Recordset
 q = "select * from stk_01 where [id_producto] = " & c_prod.ItemData(c_prod.ListIndex) & " and datevalue([fecha]) >= datevalue('" & t_fecha & "')"
 rs.Open q, cn1
 saldo = s
 While Not rs.EOF
   s = rs("cantidad")
   f = Format$(rs("fecha"), "dd/mm/yyyy")
   c = rs("comprobante")
  dE = Format$(rs("descripcion"), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
  RSet ct = Format$(rs("CANTIDAD"), "######0.00")
  If s >= 0 Then
    d = ct
    h = "      0.00"
  Else
    h = ct
    d = "      0.00"
  End If
  saldo = saldo + Val(d) - Val(h)
  RSet ct = Format$(saldo, "######0.00")
  List1.AddItem f & "  " & c & "  " & dE & "  " & d & "  " & h & "  " & ct
  rs.MoveNext
 Wend



End Sub

Sub cabecera()
lc = "-----------------------------------------------------------------------------------------------------------------------"
List1.AddItem "MOVIMIENTOS DE PRODUCTOS"
List1.AddItem ""
List1.AddItem "Producto:   " & Format$(c_prod.ItemData(c_prod.ListIndex), "00000") & " " & c_prod
List1.AddItem ""
List1.AddItem "Fecha Desde:" & t_fecha
List1.AddItem ""
List1.AddItem lc
List1.AddItem "Fecha        Comprobante         Detalle                           Ingresos     Egresos      Saldo"
List1.AddItem lc


End Sub
Private Sub c_prod_LostFocus()
If c_prod.ListIndex < 0 Then
  c_prod.ListIndex = 0
End If
End Sub

Private Sub cal1_DblClick()
If cal1.Tag = "1" Then
 t_fecha = cal1
Else
 t_fecha2 = cal1
End If
cal1.Visible = False
End Sub

Private Sub cal1_LostFocus()
t_fecha = cal1
cal1.Visible = False
End Sub

Private Sub Command1_Click()
Call carga

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  Unload Me
End If

End Sub

Private Sub Form_Load()
Call barraesag(Me)
t_fecha = Format$(Now, "dd/mm/yyyy")
t_fecha2 = Format$(Now, "dd/mm/yyyy")

cal1.Visible = False

End Sub

  




Private Sub List1_LostFocus()
List1.ListIndex = -1
End Sub


Private Sub t_fecha_DblClick()
cal1.Visible = True
cal1.Tag = 1
End Sub

Private Sub t_fecha_LostFocus()
If Not IsDate(t_fecha) Then
 t_fecha = Format$(Now, "dd/mm/yyyy")
End If
End Sub

Private Sub t_fecha2_Change()

End Sub

Private Sub t_fecha2_DblClick()
cal1.Visible = True
cal1.Tag = 2

End Sub

Private Sub t_fecha2_LostFocus()
If Not IsDate(t_fecha2) Then
 t_fecha2 = Format$(Now, "dd/mm/yyyy")
End If
End Sub
