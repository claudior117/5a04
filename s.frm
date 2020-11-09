VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form abm_periodos1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PERIODOS CONTABLES"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4320
   ScaleWidth      =   8880
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Width           =   2535
      Begin VB.TextBox t_funcion 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label10 
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Funcion"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   2415
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   8295
      Begin VB.ComboBox c_estado 
         Height          =   315
         ItemData        =   "s.frx":0000
         Left            =   2160
         List            =   "s.frx":000A
         TabIndex        =   3
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox t_fc 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   2
         ToolTipText     =   "Ingrese los digitos del 4 al 7 del cod. de barra de algun articulo de la marca"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox t_fi 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   1
         ToolTipText     =   "Ingrese los digitos del 4 al 7 del cod. de barra de algun articulo de la marca"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox t_id 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox t_descripcion 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   150
         TabIndex        =   0
         Top             =   720
         Width           =   5895
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Fecha Cierre:"
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
         TabIndex        =   17
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Fecha Inicio:"
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
         TabIndex        =   16
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Estado:"
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
         TabIndex        =   15
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Id. Periodo:"
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
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Descripcion:"
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
         TabIndex        =   10
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   6840
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Height          =   615
         Left            =   840
         Picture         =   "s.frx":0020
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "s.frx":08A2
         Style           =   1  'Graphical
         TabIndex        =   6
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
      TabIndex        =   4
      Top             =   4065
      Width           =   8880
      _ExtentX        =   15663
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
            TextSave        =   "27/02/2015"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:43"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "abm_periodos1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private EXISTE As String



Private Sub btnacepta_Click()
Call graba
End Sub

Sub graba()
J = MsgBox("Confirma Valores para Grabar", 4)
If J = 6 Then
   'On Error GoTo ERRORGRABA
   Select Case t_funcion
     
   Case "A"
      QUERY = "INSERT INTO c_10([DEscripcion], [fecha_inicio], [fecha_cierre], [estado])"
      QUERY = QUERY & " VALUES ('" & t_descripcion & "', '" & t_fi & "', '" & t_fc & "', '" & Mid$(c_estado, 1, 1) & "')"
      cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans
   Case "M"
      'verifico que todos los asientos del periodo esten dentro de las fecha de inicio y cierre
      
      Set rs = New ADODB.Recordset
      q = "select * from c_11 where (datevalue([fecha]) < datevalue('" & t_fi & "') or datevalue([fecha]) > datevalue('" & t_fc & "')) and [id_periodo] = " & Val(t_id)
      rs.Open q, cn1
      If Not rs.EOF And Not rs.BOF Then
         MsgBox ("No se puede modificar Periodo porque hay asientos ingresados fuera de las fecha de Inicio/Cierre indicadas")
      Else
        QUERY = "update c_10 set  [Descripcion]='" & t_descripcion & "' , [fecha_inicio]='" & t_fi & "' , [fecha_cierre]='" & t_fc & "' , [estado]='" & Mid$(c_estado, 1, 1) & "'"
        QUERY = QUERY & " where [id_periodo]= " & Val(t_id)
        cn1.BeginTrans
        cn1.Execute QUERY
        cn1.CommitTrans
      End If
   Case "B"
      t = MsgBox("Eliminar el periodo implica eliminar todos los Asientos ingresados en dicho periodo. Confirma Eliminar", 4)
      If t = 6 Then
            cn1.BeginTrans
            QUERY = "DELETE FROM C_10 WHERE [id_periodo] = " & Val(t_id)
            cn1.Execute QUERY
            cn1.CommitTrans
            
       End If
   
   End Select
   
   ABM_periodos.DataGrid1.Refresh
   ABM_periodos.Show
   Me.Hide
    
End If

Exit Sub
ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  
End Sub

Private Sub btnsale_Click()
Me.Hide
End Sub



Private Sub c_estado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  btnacepta.SetFocus
End If

End Sub

Private Sub c_estado_LostFocus()
If c_estado.ListIndex < 0 Then
   c_estado.ListIndex = 0
End If
End Sub

Private Sub Form_Activate()
If t_funcion = "B" Then
  btnacepta.Enabled = True
  btnacepta.SetFocus
Else
  t_descripcion.SetFocus
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyUp
     Call tabup(Me)
   Case Is = vbKeyF9
     Call graba
         
End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 3)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)
End Sub

Private Sub t_codbarra_KeyPress(KeyAscii As Integer)

End Sub

Private Sub t_descripcion_LostFocus()
Call NULOS(t_descripcion)
End Sub


Private Sub t_fc_LostFocus()
If t_fc <> "" Then
  If Not IsDate(t_fi) Then
     t_fc = Format$(Now, "dd/mm/yyyy")
  End If
Else
  t_fc = Format$(Now, "dd/mm/yyyy")
End If

End Sub

Private Sub t_fi_LostFocus()
If t_fi <> "" Then
  If Not IsDate(t_fi) Then
     t_fi = Format$(Now, "dd/mm/yyyy")
  End If
Else
  t_fi = Format$(Now, "dd/mm/yyyy")
End If
End Sub
