VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form gen_agenda1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INGRESO  DE EVENTOS DIARIOS"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8040
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4230
   ScaleWidth      =   8040
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.MonthView mv_1 
      Height          =   2370
      Left            =   4920
      TabIndex        =   18
      Top             =   240
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   114491394
      CurrentDate     =   39637
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   9240
      TabIndex        =   11
      Top             =   120
      Width           =   2535
      Begin VB.TextBox t_funcion 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   12
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
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7815
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cal"
         Height          =   195
         Left            =   3840
         TabIndex        =   16
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox t_obs 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   1800
         MaxLength       =   79
         TabIndex        =   2
         Top             =   1680
         Width           =   5895
      End
      Begin VB.TextBox t_fecha 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   0
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox t_id 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox t_descripcion 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   1800
         MaxLength       =   79
         TabIndex        =   1
         Top             =   1200
         Width           =   5895
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   3240
         TabIndex        =   17
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         Caption         =   "Observaciones"
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
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         Caption         =   "Fecha"
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
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         Caption         =   "Id. Evento"
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
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         Caption         =   "Evento"
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
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   6120
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Height          =   615
         Left            =   840
         Picture         =   "ARCH014.frx":0000
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
         Picture         =   "ARCH014.frx":0882
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
      Top             =   3975
      Width           =   8040
      _ExtentX        =   14182
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
            TextSave        =   "31/03/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:34 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "gen_agenda1"
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
      QUERY = "INSERT INTO g7([descripcion], [fecha_vto], [id_usuario], [observaciones])"
      QUERY = QUERY & " VALUES ('" & t_descripcion & " ', '" & t_fecha & "' , " & para.id_usuario & " , '" & t_obs & " ')"
      cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans
   
   
   Case "M"
   
      QUERY = "update g7 set  [Descripcion]='" & t_descripcion & " ' , [fecha_vto]='" & t_fecha & "' , [id_usuario]=" & para.id_usuario & " , [observaciones]='" & t_obs & " '"
      QUERY = QUERY & " where [id_evento]= " & Val(t_id)
      cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans
      
   
   
   End Select
   
   gen_agenda.DataGrid1.Refresh
   gen_agenda.Show
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



Private Sub Check1_Click()
If Check1 = 1 Then
  mv_1.Visible = True
Else
  mv_1.Visible = False
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
    Call TabEnter2(Me, 2)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)
mv_1.Value = Now
mv_1.Visible = False
End Sub


Private Sub Mv_1_DblClick()
t_fecha = mv_1.Value

End Sub

Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
     t_fecha = Format$(Now, "dd/mm/yyyy")
  End If
Else
  t_fecha = Format$(Now, "dd/mm/yyyy")
End If
End Sub

Private Sub t_obs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  btnacepta.Enabled = True
  btnacepta.SetFocus
End If
End Sub

Private Sub UpDown1_DownClick()
If t_fecha <> "" Then
   t_fecha = DateValue(t_fecha) - 1
Else
   t_fecha = Format$(Now, "dd/mm/yyyy")
End If

End Sub

Private Sub UpDown1_UpClick()
If t_fecha <> "" Then
   t_fecha = DateValue(t_fecha) + 1
Else
   t_fecha = Format$(Now, "dd/mm/yyyy")
End If

End Sub
