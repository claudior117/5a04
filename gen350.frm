VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form gen_exportaPercibweb 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PROCESO PARA EXPORTAR PERCEPCIOONES IBBA REALIZADAS VIA WEB"
   ClientHeight    =   7620
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   7710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7620
   ScaleWidth      =   7710
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Archivo definitivo"
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   3120
      Width           =   7455
      Begin VB.TextBox t_camino 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   7215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Periodo"
      Height          =   2655
      Left            =   3960
      TabIndex        =   9
      Top             =   120
      Width           =   3735
      Begin VB.TextBox t_periodo 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   22
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox t_lote 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   12
         TabIndex        =   16
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox t_actividad 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   14
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         Caption         =   "Periodo:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Nombre Lote:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Codigo Actividad:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "SELECCIONE UBICACION ARCHIVO"
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3735
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   3135
      End
      Begin VB.DirListBox Dir1 
         Height          =   2115
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   6000
      TabIndex        =   1
      Top             =   6240
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "gen350.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "gen350.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   2
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
      TabIndex        =   0
      Top             =   7365
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   10583
            MinWidth        =   10583
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label11 
      Caption         =   "Codigo Actividad: Actividad por la cual percibe exepto 29, 7 quincenal y 17 de bancos"
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   25
      Top             =   6360
      Width           =   5535
   End
   Begin VB.Label Label10 
      Caption         =   "Nombre Lote: LOTE1, LOTE2, etc."
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   6000
      Width           =   5535
   End
   Begin VB.Label Label8 
      Caption         =   "Donde el periodo es: AAAAMMQ, año mes y quincena. Si la presentacion es mensual la quincena es 0, si no 1 o 2."
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   21
      Top             =   5400
      Width           =   5535
   End
   Begin VB.Label Label7 
      Caption         =   "El nombre de archivo esta compuesto por. AR-cuit-perodo-actividad-lote.txt"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   5040
      Width           =   5535
   End
   Begin VB.Label Label2 
      Caption         =   $"gen350.frx":1104
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   7455
   End
   Begin VB.Label Label1 
      Caption         =   "Este modulo genera un archivo .txt con las percepciones realizadas y las caracteristicas solicitadas por ARBA. "
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   7455
   End
End
Attribute VB_Name = "gen_exportaPercibweb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim habilitafacturaremito As Boolean
Dim t1 As String



Private Sub btnacepta_Click()
Dim l As String
If verifica Then
 Call sacaperiodo
 CUIT = Mid$(glo.CUIT, 1, 2) & Mid$(glo.CUIT, 4, 8) & Mid$(glo.CUIT, 13, 1)
 na = "AR-" & CUIT & "-" & t_periodo & "-" & t_actividad & "-" & t_lote & ".txt"
 t_camino = t_camino & na
 r = 0
 c = 0
 z = 0
 J = MsgBox("Confirma Generacion de Archivo de Percepciones " & na, 4)
 If J = 6 Then
   q = "select fecha, cuit02, num_comp, letra, id_tipocomp, sucursal, base_imponible, importe from vta_016, vta_02 where vta_016.num_int = vta_02.num_int and  id_percepcion = 2 and datevalue([fecha]) >= datevalue('" & t_fecha & "') and datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
   Set rs = New adodb.Recordset
  ' MsgBox (q)
   rs.Open q, cn1
   
   Open t_camino For Output As #1
   
  
   espere.Show
   espere.Refresh
   While Not rs.EOF
    CUIT = Format$(Left$(rs("cuit02"), 11), "00-00000000-0")
    fechac = Format$(rs("fecha"), "dd/mm/yyyy")
    Select Case rs("id_tipocomp")
      Case Is = 1
         tc = "F"
      Case Is = 2
         tc = "D"
      Case Is = 3
        tc = "C"
      Case Is = 30
        tc = "E"
      Case Is = 31
        tc = "I"
      Case Is = 32
        tc = "H"
      Case Else
       tc = "V"
    
    End Select
    numc = Format$(rs("num_comp"), "00000000")
    succ = Format$(rs("sucursal"), "0000")
    
    If rs("id_tipocomp") <> 3 And rs("id_tipocomp") <> 32 Then 'distinto NC
        impc = Format(rs("base_imponible"), "000000000.00")
        impp = Format(rs("importe"), "00000000.00")
    Else
        impc = Format(rs("base_imponible"), "-00000000.00")
        impp = Format(rs("importe"), "-0000000.00")
    End If
    linea = CUIT & fechac & tc & rs("letra") & succ & numc & impc & impp & "A"
    Print #1, linea
    If Len(linea) = 61 Then
      c = c + 1
    End If
    r = r + 1
    espere.Label1 = "Percepciones exportadas " & r
    espere.Label1.Refresh
    rs.MoveNext
   Wend
   Unload espere
   Close #1
 End If
 MsgBox ("Proceso finalizado. Percepciones Exportadas " & r & "   Percepciones Correctas: " & c)
End If
Set rs = Nothing


End Sub
Function verifica() As Boolean
If t_fecha <> "" And t_fecha2 <> "" And Val(t_actividad) > 0 And t_lote <> "" Then
  verifica = True
Else
  verifica = False
End If

  
End Function
Private Sub btnsale_Click()
Unload Me
End Sub



Private Sub Dir1_Change()

Call camino
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1
Call camino
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
End If

End Sub
Sub camino()
If Dir1 <> "C:\" Then
  t_camino = Dir1 & "\"
Else
  t_camino = Dir1
End If

End Sub



Private Sub Form_Load()
Call camino
End Sub

Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
   If Not IsDate(t_fecha) Then
      t_fecha = ""
   End If
End If
End Sub

Private Sub t_fecha2_LostFocus()
If t_fecha2 <> "" Then
   If Not IsDate(t_fecha2) Then
      t_fecha = ""
   Else
      If t_fecha <> "" Then
        If DateValue(t_fecha2) < DateValue(t_fecha) Then
            tf = t_fecha2
            t_fecha2 = t_fecha
            t_fecha = tf
        End If
        Call sacaperiodo
      Else
        t_fecha = t_fecha2
        t_fecha2 = ""
      End If
   End If
End If
End Sub
Sub sacaperiodo()
CD = DateValue(t_fecha2) - DateValue(t_fecha)
If CD > 15 Then
  q = 0
Else
  If Val(Mid$(t_fecha, 1, 2)) > 15 Then
    q = 2
  Else
    q = 1
  End If
End If
t_periodo = Mid$(t_fecha, 7, 4) & Mid$(t_fecha, 4, 2) & q

End Sub
