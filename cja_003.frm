VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cja_cierremes 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CIERRE DE CAJA"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7170
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3510
   ScaleWidth      =   7170
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6735
      Begin VB.TextBox t_f1 
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox t_f2 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox c_estado 
         Height          =   315
         ItemData        =   "cja_003.frx":0000
         Left            =   1560
         List            =   "cja_003.frx":000A
         TabIndex        =   0
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Desde:"
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
         Left            =   360
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "El cierre de caja evita que se ingresen movimientos  una vez cerrado el periodo"
         ForeColor       =   &H000000FF&
         Height          =   1335
         Left            =   3840
         TabIndex        =   10
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label3 
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
         Left            =   360
         TabIndex        =   9
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Hasta:"
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
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   5280
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Height          =   615
         Left            =   840
         Picture         =   "cja_003.frx":0020
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
         Picture         =   "cja_003.frx":08A2
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
      Top             =   3255
      Width           =   7170
      _ExtentX        =   12647
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
            TextSave        =   "09:41"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "cja_cierremes"
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
If verifica Then
 J = MsgBox("Confirma Cambiar estado para el Periodo seleccionado", 4)
 If J = 6 Then
   'On Error GoTo ERRORGRABA
   seguir = True
   i = 0
   a = Val(Mid$(t_f1, 7, 4))
   m = Val(Mid$(t_f1, 4, 2))
   d = Val(Mid$(t_f1, 1, 2))
   
   While seguir
      fecha = DateSerial(a, m, d + i)
      If DateValue(fecha) <= DateValue(t_f2) Then
         q = "select * from  cyb_09 where datevalue([fecha]) = datevalue('" & fecha & "')"
         Set rs = New ADODB.Recordset
         rs.Open q, cn1, adOpenDynamic, adLockOptimistic
         If Not rs.EOF And Not rs.BOF Then
           'existe
           rs("estado") = Mid$(c_estado, 1, 1)
           rs.Update
         Else
           rs.AddNew
           rs("fecha") = Format$(DateValue(fecha), "dd/mm/yyyy")
           rs("estado") = Mid$(c_estado, 1, 1)
           rs.Update
         End If
         i = i + 1
      Else
        seguir = False
      End If
   Wend
   MsgBox ("El estado de la caja en el periodo ha sido cambiado")
   Unload Me
 End If
End If

Exit Sub
ERRORGRABA:
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos. Modulo: Graba")
  
End Sub
Function verifica() As Boolean
v = 1
If t_f1 <> "" Then
  If Not IsDate(t_f1) Then
     v = 0
     MsgBox ("Error al Ingresar Fecha Desde")
  End If
Else
   v = 0
   MsgBox ("Debe Ingresar una Fecha")
End If

If t_f2 <> "" Then
  If Not IsDate(t_f2) Then
     v = 0
     MsgBox ("Error al Ingresar Fecha Hasta")
  Else
    If v = 1 Then
      If DateValue(t_f1) > DateValue(t_f2) Then
         v = 0
         MsgBox ("La Fecha Final debe ser posterior o igual  a la Inicial")
      End If
    End If
  End If
Else
  t_f2 = t_f1
End If

If v = 0 Then
  verifica = False
Else
 verifica = True
End If

End Function
Private Sub btnsale_Click()
Unload Me
End Sub



Private Sub c_cuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  btnacepta.SetFocus
End If
End Sub

Private Sub c_cuenta_LostFocus()
If c_cuenta.ListIndex < 0 Then
  c_cuenta.ListIndex = 0
End If
End Sub

Private Sub c_año_LostFocus()
If c_año.ListIndex < 0 Then
  Call actual
End If

End Sub

Private Sub c_estado_Change()
If c_estado.ListIndex < 0 Then
  c_estado.ListIndex = 0
End If

End Sub

Private Sub c_mes_LostFocus()
If c_mes.ListIndex < 0 Then
  Call actual
End If
End Sub

Private Sub Form_Load()
t_f1 = ""
t_f2 = ""
c_estado.ListIndex = 0
Call actual

End Sub
Sub actual()

t_f1 = Format$(Now, "dd/mm/yyyy")
t_f2 = Format$(Now, "dd/mm/yyyy")



End Sub



Private Sub t_f1_GotFocus()
t_f1 = ""
End Sub

Private Sub t_f2_GotFocus()
t_f2 = ""
End Sub
