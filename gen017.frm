VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form gen_exportaretib 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PROCESO PARA EXPORTAR RETENCIONES IB REALIZADAS AL SICORE"
   ClientHeight    =   6420
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   7710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6420
   ScaleWidth      =   7710
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Periodo"
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   4080
      Width           =   3735
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "SELECCIONE ARCHIVO ORIGEN"
      Height          =   3855
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3735
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ubicacion definitiva de los archivos"
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   3000
         Width           =   3495
         Begin VB.TextBox t_camino 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   240
            Width           =   3255
         End
      End
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
      Left            =   5400
      TabIndex        =   1
      Top             =   5160
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "gen017.frx":0000
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
         Picture         =   "gen017.frx":0882
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
      Top             =   6165
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
   Begin VB.Label Label6 
      Caption         =   "Este modulo genera el archivo RETIB.TXT en el directorio seleccionado."
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   3960
      TabIndex        =   17
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label Label5 
      Caption         =   "Desde el Aplic. IB se podran importar las retenciones. "
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   3960
      TabIndex        =   16
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Desde el SICORE se podran importar en ese orden para pasar automaticamente las retenciones generadas en el periodo."
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   3960
      TabIndex        =   8
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Este modulo genera 2 archivos SUJETOS.TXT y RETENCIONES.TXT en el directorio seleccionado."
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   3960
      TabIndex        =   7
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "gen_exportaretib"
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
 r = 0
 c = 0
 z = 0
 J = MsgBox("Confirma Generacion de Archivo. ", 4)
 If J = 6 Then
   q = "select * from a5, a1  where [id_tipocomp] = 96 and datevalue([fecha]) >= datevalue('" & t_fecha & "') and datevalue([fecha]) <= datevalue('" & t_fecha2 & "') and a5.[id_proveedor] = a1.[id_proveedor] "
   Set rs = New adodb.Recordset
   'MsgBox (q)
   rs.Open q, cn1
   
   Open t_camino & "retib.txt" For Output As #1
   
  
   espere.Show
   espere.Refresh
   While Not rs.EOF
    'sujetos.txt
    CUIT = Format$(Left$(rs("cuit"), 11), "00-00000000-0")
    fechac = Format$(rs("fecha"), "dd/mm/yyyy")
    numc = Format$(rs("num_comprobante"), "00000000")
    succ = Format$(rs("sucursal"), "0000")
    impc = Format(rs("total"), "0000000.00")
    linea = CUIT & fechac & succ & numc & impc
    Print #1, linea
    If Len(linea) = 45 Then
      c = c + 1
    End If
    r = r + 1
    espere.Label1 = "Retenciones exportadas " & r
    espere.Label1.Refresh
    rs.MoveNext
   Wend
   Unload espere
   Close #1
 End If
 Label1 = "Fin"
 Label2 = "Comprobantes Exportados " & r & "   Retenciones Correctas: " & c
 
End If
Set rs = Nothing


End Sub
Function verifica() As Boolean
If t_fecha <> "" And t_fecha2 <> "" Then
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
Label1 = "En espera..."
Label2 = ""
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
      Else
        t_fecha = t_fecha2
        t_fecha2 = ""
      End If
   End If
End If
End Sub
