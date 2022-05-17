VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form gen_cf 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INFORME ENTRE FECHAS CONTROLADOR FISCAL"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Salida:"
      Height          =   615
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   6855
      Begin VB.CommandButton Command2 
         Caption         =   "Carpeta destino:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox t_carpeta 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   4815
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   2760
      TabIndex        =   9
      Top             =   120
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   175702017
      CurrentDate     =   38750
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   3615
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   5520
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "gen042..frx":0000
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
         Picture         =   "gen042..frx":0882
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
      Top             =   3720
      Width           =   7350
      _ExtentX        =   12965
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
            TextSave        =   "14/05/2022"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:13 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label2 
      Caption         =   $"gen042..frx":1104
      Height          =   855
      Left            =   240
      TabIndex        =   13
      Top             =   2640
      Width           =   5055
   End
End
Attribute VB_Name = "gen_cf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim c5 As Double
Dim Fiscali1 As Driver


Private Sub btnacepta_Click()
 If glo.sucursalf <> 0 Then
   Call exporta
 Else
   MsgBox ("Terminal no habilitada para Controlador Fiscal")
End If
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub







Private Sub cal1_DblClick()
If cal1.Tag = "1" Then
  t_fecha = cal1.Value
Else
  t_fecha2 = cal1.Value
End If
cal1.Visible = False
End Sub

Private Sub cal1_LostFocus()
cal1.Visible = False
End Sub





Private Sub Command2_Click()
Load gen_seleccionacarpeta
gen_seleccionacarpeta.t_llamada = "6"
gen_seleccionacarpeta.Show

End Sub





Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     gen_tools.Show
End Select
End Sub



Private Sub Form_Load()

Call barraesag(Me)
cal1.Visible = False
t_carpeta = "c:\"
t_fecha = Format$(Now, "dd/mm/yy")
t_fecha2 = Format$(Now, "dd/mm/yy")

Set cl_fiscal = New fiscal
   cl_fiscal.carga (glo.sucursalf)
   If cl_fiscal.id > 0 Then
      cMODELO = cl_fiscal.idmodelo
      cPUERTO = cl_fiscal.puerto
      cBAUDIOS = cl_fiscal.baudios
      Set Fiscali1 = New Driver
      Fiscali1.Modelo = cMODELO
      Fiscali1.puerto = cPUERTO
      Fiscali1.baudios = cBAUDIOS
      
   Else
      MsgBox ("Impresora Fiscal No definida")
   End If



End Sub




Sub exporta()
If t_fecha <> "" And t_fecha2 <> "" Then
   f1 = Mid$(t_fecha, 1, 2) & Mid$(t_fecha, 4, 2) & Mid$(t_fecha, 7, 2)
   f2 = Mid$(t_fecha2, 1, 2) & Mid$(t_fecha2, 4, 2) & Mid$(t_fecha2, 7, 2)
    
    
      
  On Error GoTo DepuraErrores
  
 If Not Fiscali1.Inicializar Then
    Err.Raise Fiscali1.Error, "", Fiscali1.ErrorDesc
  End If
  
  Fiscali1.CancelarComprobante
  archivo22 = t_carpeta & "\reporte" & f1 & "-" & f2 & ".zip"
  
  If Fiscali1.ObtenerPrimerBloqueReporteElectronico(f1, f2, archivo22, 0) Then
    Do
    Loop While Fiscali1.ObtenerSiguienteBloqueReporteElectronico
  End If

  If Fiscali1.Error <> 0 Then
    Err.Raise Fiscali1.Error, "", Fiscali1.ErrorDesc
  End If
  
  Fiscali1.Finalizar
  
  MsgBox ("Reporte descargado exitosamente")
  
Else
   MsgBox ("Ingrese Fechas correctamente")
End If
  
  
  
  Exit Sub

DepuraErrores:
  Fiscali1.Finalizar
  MsgBox ("Error: " & Fiscali1.ErrorDesc)




End Sub


Private Sub t_fecha_DblClick()
cal1.Visible = True
cal1.Tag = "1"


End Sub

Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
    t_fecha = Format$(Now, "dd/mm/yy")
  Else
    t_fecha = Format$(t_fecha, "dd/mm/yy")
  End If
Else
   t_fecha = Format$(Now, "dd/mm/yy")
End If
End Sub

Private Sub t_fecha2_DblClick()
cal1.Visible = True
cal1.Tag = "2"

End Sub

Private Sub t_fecha2_LostFocus()
If t_fecha2 <> "" Then
  If Not IsDate(t_fecha2) Then
    t_fecha2 = Format$(Now, "dd/mm/yy")
  Else
    t_fecha2 = Format$(t_fecha2, "dd/mm/yy")
  End If
Else
  t_fecha2 = Format$(Now, "dd/mm/yy")
End If

End Sub
