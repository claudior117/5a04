VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form gen_borra 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Elimina Movimientos en lote"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3120
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   2295
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   9015
      Begin VB.ComboBox c_zona 
         Height          =   315
         Left            =   6480
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox c_suc 
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   720
         Width           =   2175
      End
      Begin VB.ComboBox c_comp 
         Height          =   315
         Left            =   2160
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   1680
         Width           =   5895
      End
      Begin VB.TextBox t_f2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6480
         MaxLength       =   12
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox t_f1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   12
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox c_cli 
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   1200
         Width           =   5895
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Zona:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4800
         TabIndex        =   16
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Sucursal:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   480
         TabIndex        =   15
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Comprobantes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   14
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   4800
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   480
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   480
         TabIndex        =   11
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9840
      TabIndex        =   7
      Top             =   1680
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Height          =   615
         Left            =   840
         Picture         =   "gen030.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "gen030.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   8
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
      TabIndex        =   6
      Top             =   2865
      Width           =   11910
      _ExtentX        =   21008
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
Attribute VB_Name = "gen_borra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private EXISTE As String



Private Sub btnacepta_Click()
J = MsgBox("Este proceso es irreversible. Esta seguro que quiere eliminar comprobantes por lote", 4)
If J = 6 Then
  k = MsgBox("Confirme nuevamente eliminar comprobantes por lotes", 4)
  If k = 6 Then
   Call graba
  End If
End If
End Sub

Sub graba()
      q = "select * from vta_02 "
      c = " where "
      
      If t_f1 <> "" Then
       q = q & c & " datevalue([fecha]) >= datevalue('" & t_f1 & "')"
      c = " and "
      End If
      
      If t_f2 <> "" Then
       q = q & c & " datevalue([fecha]) <= datevalue('" & t_f2 & "')"
       c = " and "
      End If
      
      If c_cli.ListIndex > 0 Then
        q = q & c & " [id_cliente] = " & c_cli.ItemData(c_cli.ListIndex)
        c = " and "
      End If
      
      If c_comp.ListIndex > 0 Then
        q = q & c & " [id_tipocomp] = " & c_comp.ItemData(c_comp.ListIndex)
        c = " and "
      End If
      
      If c_suc.ListIndex > 0 Then
         q = q & c & " [sucursal_ingreso] = " & Val(c_suc)
         c = " and "
      End If

      espere.Show
      c = 1
      Set rs1 = New ADODB.Recordset
      rs1.Open q, cn1
      Set cl_compvta = New comprobantes_venta
      While Not rs1.EOF
        espere.Label1 = "Borrando registro.... " & c
        espere.Label1.Refresh
        cl_compvta.cargar2 (rs1("num_int"))
        rs1.MoveNext
        c = c + 1
        cl_compvta.borrar
     Wend
     Set cl_compvta = Nothing
     Set rs1 = Nothing
      Unload espere
End Sub

Private Sub btnsale_Click()
Me.Hide
End Sub





Private Sub c_cli_LostFocus()
If c_cli.ListIndex < 0 Then
  c_cli.ListIndex = 0
End If

End Sub

Private Sub c_comp_LostFocus()
If c_comp.ListIndex < 0 Then
  c_comp.ListIndex = 0
End If

End Sub

Private Sub c_suc_LostFocus()
If c_suc.ListIndex < 0 Then
  c_suc.ListIndex = 0
End If
End Sub

Private Sub c_zona_LostFocus()
If c_zona.ListIndex < 0 Then
  c_zona.ListIndex = 0
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 5)
  
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Call carga_clientes(c_cli)
c_cli.AddItem "<Todos>", 0
c_cli.ListIndex = 0
 
Call carga_SUCURSALES(c_suc)
c_suc.AddItem "<Todas>", 0
c_suc.ListIndex = 0

Set rs = New ADODB.Recordset
q = "select * from vta_06 where [sucursal] = " & glo.sucursal
rs.Open q, cn1
Call llena_combo(rs, "descripcion", "id_tipocomp", c_comp, True)
Set rs = Nothing
c_comp.AddItem "<Todos>", 0
c_comp.ListIndex = 0

c_zona.clear
c_zona.AddItem "<Todas>", 0
c_zona.AddItem "Zona 1", 1
c_zona.AddItem "Zona 2", 2
c_zona.ListIndex = 0


End Sub



Private Sub t_f1_GotFocus()
t_f1 = ""
End Sub

Private Sub t_f1_LostFocus()
If t_f1 <> "" Then
  If Not IsDate(t_f1) Then
    t_f1 = ""
  Else
    t_f1 = Format$(t_f1, "dd/mm/yyyy")
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
  Else
    t_f2 = Format$(t_f2, "dd/mm/yyyy")
  End If
End If

End Sub
