VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form abm_PERC1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CODIGOS DE PERCEPCION"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9315
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3930
   ScaleWidth      =   9315
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   2760
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
      Begin VB.ComboBox c_impuesto 
         Height          =   315
         ItemData        =   "Arch012.frx":0000
         Left            =   2160
         List            =   "Arch012.frx":0013
         TabIndex        =   2
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox t_tipo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   1
         ToolTipText     =   "[P] Percepcion - [R] Retencion"
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   7440
         Picture         =   "Arch012.frx":0059
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1800
         Width           =   495
      End
      Begin VB.ComboBox c_cuenta 
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   1800
         Width           =   5055
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
         Caption         =   "Impuesto"
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
         TabIndex        =   18
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Tipo:"
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
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Cuenta Cont."
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
         Caption         =   "Id. Percepcion"
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
         Caption         =   "Descripcion"
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
         Picture         =   "Arch012.frx":0363
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
         Picture         =   "Arch012.frx":0BE5
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
      Top             =   3675
      Width           =   9315
      _ExtentX        =   16431
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
Attribute VB_Name = "abm_PERC1"
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
   On Error GoTo ERRORGRABA
   
   Select Case t_funcion
      
   Case "A"
      QUERY = "INSERT INTO a12([DEscripcion], [id_cuenta], [tipo12], [impuesto12])"
      QUERY = QUERY & " VALUES ('" & t_descripcion & "', " & c_cuenta.ItemData(c_cuenta.ListIndex) & ", '" & t_tipo & " ', '" & Mid$(c_impuesto, 2, 1) & "')"
      cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans
   
   
   Case "M"
   
      QUERY = "update a12 set  [Descripcion]='" & t_descripcion & "' , [id_cuenta]=" & c_cuenta.ItemData(c_cuenta.ListIndex) & ", [tipo12]= '" & t_tipo & "'" & ", [impuesto12]= '" & Mid$(c_impuesto, 2, 1) & "'"
      QUERY = QUERY & " where [id_percepcion]= " & Val(t_id)
      cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans
      
   Case "B"
      QUERY = "DELETE FROM a12 WHERE [id_percepcion] = " & Val(t_id)
      cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans
   
   
   End Select
   
   ABM_perc.DataGrid1.Refresh
   ABM_perc.Show
   Me.Hide
    
End If

Exit Sub
ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos. Modulo: Graba")
  
End Sub

Private Sub btnsale_Click()
Me.Hide
End Sub



Private Sub c_cuenta_KeyPress(KeyAscii As Integer)
If c_cuenta.ListIndex < 0 Then
  If Val(c_cuenta) > 0 Then
    c_cuenta.ListIndex = buscaindice(c_cuenta, Val(c_cuenta))
  Else
    c_cuenta.ListIndex = 0
  End If
End If
End Sub

Private Sub c_cuenta_LostFocus()
If c_cuenta.ListIndex < 0 Then
  c_cuenta.ListIndex = 0
End If
End Sub

Private Sub c_impuesto_LostFocus()
If c_impuesto.ListIndex < 0 Then
  c_impuesto.ListIndex = 0
End If
End Sub

Private Sub Command1_Click()
cgr_buscacuenta.Show
End Sub

Private Sub Form_Activate()
If t_funcion = "B" Then
  btnacepta.Enabled = True
  btnacepta.SetFocus
Else
  t_descripcion.SetFocus
End If
If para.cuenta_sel > 0 Then
  c_cuenta.ListIndex = buscaindice(c_cuenta, para.cuenta_sel)
  para.cuenta_sel = 0
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
Call carga_cuentas_cont(c_cuenta, "C", "D")
para.cuenta_sel = 0
End Sub


Private Sub t_descripcion_LostFocus()
If t_descripcion = "" Then
  t_descripcion = "*"
End If
End Sub


Private Sub Text1_Change()

End Sub

Private Sub Text1_LostFocus()

End Sub

Private Sub t_tipo_LostFocus()
If t_tipo = "" Then
  t_fipo = "P"
Else
 t_tipo = Format$(t_tipo, ">@")
 If t_tipo <> "P" And t_tipo <> "R" Then
   t_tipo = "P"
 End If
End If
 
End Sub
