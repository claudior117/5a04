VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form cgr_vermayores 
   BackColor       =   &H00E0E0E0&
   Caption         =   "MUESTRA MAYORES TEMPORALES"
   ClientHeight    =   8760
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4785
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   11535
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   11535
      Begin VB.ComboBox c_cuenta2 
         Height          =   315
         Left            =   2760
         TabIndex        =   14
         Text            =   "c_cuenta2"
         Top             =   720
         Width           =   4815
      End
      Begin VB.TextBox t_cuenta2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox t_cuenta 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   9960
         MaxLength       =   10
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   9960
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox c_cuenta 
         Height          =   315
         Left            =   2760
         TabIndex        =   0
         Text            =   "c_cuenta"
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Cuenta Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8400
         TabIndex        =   10
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8400
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Cuenta Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1455
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
         Picture         =   "CGR003.frx":0000
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
         Picture         =   "CGR003.frx":0882
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
      Top             =   8505
      Width           =   12015
      _ExtentX        =   21193
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
            TextSave        =   "21/01/2010"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "06:24 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "cgr_vermayores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Sub cabecera()
List1.clear
Call cabeceralist(List1)
List1.AddItem ""
List1.AddItem "LIBRO MAYOR"
List1.AddItem ""
List1.AddItem "Fecha Desde.....:" & t_fecha
List1.AddItem "Fecha Hasta.....:" & t_fech2

End Sub

Sub carga()
  Call cabecera
  'filtro cuentas
  q = "select * from c_01"
  c = " where "
  If t_cuenta <> "" Then
     q = q & c & " [id_cuenta] >= " & Val(t_cuenta)
     c = " and "
  End If
  
  If t_cuenta2 <> "" Then
     q = q & c & " [id_cuenta] <= " & Val(t_cuenta2)
     c = " and "
  End If
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  While Not rs.EOF
      Call armomayor(rs("id_cuenta"))
      rs.MoveNext
  Wend
  Set rs = Nothing
  
  
   
End Sub

Sub armomayor(ByVal i As Long)
q = "select * from c_03, c_02, c_01 where c_03.[id_cuenta] = " & i & " and c_03.[num_interno] = c_02.[num_interno] and c_03.[id_cuenta] = c_01.[Id_cuenta]"
If t_fecha <> "" Then
  q = q & " and datevalue([fecha]) >= datevalue('" & t_fecha & "')"
Else
  sa = 0
End If

If t_fecha2 <> "" Then
   q = q & " and datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
End If

q = q & " order by [fecha], c_02.[num_interno], [renglon]"
Set rs2 = New ADODB.Recordset
rs2.Open q, cn1
d = Space$(10)
h = Space$(10)
If Not rs2.BOF And Not rs2.EOF Then
  List1.AddItem ""
  List1.AddItem "Cuenta:..... [" & i & "] " & rs2("c_01.descripcion")
  td = 0
  th = 0
  List1.AddItem "================================================================================================="
  List1.AddItem "Nro. As.    Fecha    Operacion                                DEBE      HABER    Observaciones"
  List1.AddItem "================================================================================================="
  While Not rs2.EOF
    f = Format$(rs2("fecha"), "dd/mm/yyyy")
    ope = Format$(Left$(Trim$(rs2("c_02.descripcion")), 35), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    ni = Format$(rs2("c_02.num_interno"), "00000000")
    obs = Left$(RTrim$(rs2("observaciones")), 20)
    RSet d = Format$(rs2("c_03.importe"), "######0.00")
    If rs2("ubicacion") = "D" Then
      e = ""
      td = td + Val(d)
      e2 = "          "
    Else
      e = "          "
      e2 = ""
      th = th + Val(d)
    End If
    List1.AddItem ni & " " & f & " " & ope & "  " & e & d & e2 & "   " & obs
    rs2.MoveNext
   Wend
   s = Space$(10)
   RSet d = Format$(td, "######0.00")
   RSet h = Format$(th, "######0.00")
   RSet s = Format$(Val(td) - Val(th), "######0.00")
   List1.AddItem "                                                        *****************************************"
   List1.AddItem Space$(57) & d & " " & h & " ---> " & s
   
   List1.AddItem "                                                        *****************************************"
   
   List1.AddItem ""
   List1.AddItem ""
End If
End Sub
Private Sub btnacepta_Click()
Call carga
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub










Private Sub c_cuenta_LostFocus()
If c_cuenta.ListIndex < 0 Then
  c_cuenta.ListIndex = 0
  t_cuenta = ""
Else
  t_cuenta = c_cuenta.ItemData(c_cuenta.ListIndex)
End If
End Sub

Private Sub c_cuenta2_LostFocus()
If c_cuenta2.ListIndex < 0 Then
  c_cuenta2.ListIndex = 0
  t_cuenta2 = ""
Else
  t_cuenta2 = c_cuenta2.ItemData(c_cuenta2.ListIndex)
End If
End Sub

Private Sub Form_Load()


Call carga_cuentas_cont(c_cuenta, "C", "D")
c_cuenta.AddItem "<Todas>", 0
c_cuenta.ListIndex = 0

Call carga_cuentas_cont(c_cuenta2, "C", "D")
c_cuenta2.AddItem "<Todas>", 0
c_cuenta2.ListIndex = 0

Call limpia
Call barraesag(Me)


End Sub
Sub limpia()
t_fecha = ""
t_fecha2 = ""
t_cuenta = ""
t_cuenta2 = ""
List1.clear
End Sub

Private Sub List1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F7] Imprime "

End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
  J = MsgBox("Prepare Impresora y Confirme", 4)
  If J = 6 Then
    Call imprimelist(List1, 70, 9, False, True, 10)
  End If

End If

If KeyCode = vbKeyF8 Then
 Call nivel_acceso(2)
 If para.id_grupo_modulo_actual >= 8 Then
   
  If Val(Mid$(List1.List(List1.ListIndex), 1, 8)) > 0 Then
    J = MsgBox("Confirma Eliminar Asiento " & Mid$(List1.List(List1.ListIndex), 1, 8), 4)
    If J = 6 Then
      nicgr = Val(Mid$(List1.List(List1.ListIndex), 1, 8))
      cn1.BeginTrans
      QUERY = "DELETE FROM c_02 WHERE [num_interno] = " & nicgr
      cn1.Execute QUERY
      
      QUERY = "DELETE FROM c_03 WHERE [num_interno] = " & nicgr
      cn1.Execute QUERY
      cn1.CommitTrans
      
      Call carga
    End If
  End If
 Else
   Call sinpermisos
 End If
End If

End Sub

Private Sub t_cuenta_LostFocus()
If t_cuenta <> "" Then
  c_cuenta.ListIndex = buscaindice(c_cuenta, Val(t_cuenta))
  If c_cuenta.ListIndex = 0 Then
    t_cuenta = ""
  End If
End If
End Sub

Private Sub t_cuenta2_LostFocus()
If t_cuenta2 <> "" Then
  c_cuenta2.ListIndex = buscaindice(c_cuenta2, Val(t_cuenta2))
  If c_cuenta2.ListIndex = 0 Then
    t_cuenta2 = ""
  End If
End If
End Sub

Private Sub t_fecha_GotFocus()
t_fecha = ""
End Sub

Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
    t_fecha = ""
  End If
End If
End Sub

Private Sub t_fecha2_GotFocus()
t_fecha2 = ""
End Sub

Private Sub t_fecha2_LostFocus()
If t_fecha2 <> "" Then
  If Not IsDate(t_fecha2) Then
    t_fecha2 = ""
  End If
End If

End Sub
