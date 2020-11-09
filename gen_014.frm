VERSION 5.00
Begin VB.Form gen_seleccionarimp 
   Caption         =   "Seleccionar Impresora Salida"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Nueva Configuracion"
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   8295
      Begin VB.ComboBox c_imp 
         Height          =   315
         Left            =   2880
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Impresoras Disponibles"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Configuracion Actual"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8295
      Begin VB.TextBox t_actual 
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   5175
      End
      Begin VB.TextBox t_predeterminada 
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label Label3 
         Caption         =   "Impresora Actual del Sistema:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Impresora Predeterminada Windows:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   7335
      Begin VB.CommandButton Command3 
         Caption         =   "Seleccionar Predeterminada"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         Height          =   255
         Left            =   5880
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Seleccionar Nueva"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
   End
End
Attribute VB_Name = "gen_seleccionarimp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  
    If c_imp.ListIndex <> -1 Then
       Call Establecer(c_imp.Text)
       MsgBox "Se usará la impresora: " & _
       Printer.DeviceName & " para imprimir ", vbInformation
       para.impresora_actual = Printer.DeviceName
       Call muestra
    End If
End Sub
  
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
       Call Establecer(para.IMPRESORA_PREDETERMINADA)
       MsgBox "Se usará la impresora: " & _
       Printer.DeviceName & " para imprimir ", vbInformation
       para.impresora_actual = Printer.DeviceName
       Call muestra

End Sub

Private Sub Form_Load()
    Obtener_Impresoras
    Call muestra
End Sub
Sub muestra()
  t_predeterminada = para.IMPRESORA_PREDETERMINADA
  t_actual = para.impresora_actual
End Sub

Public Function Obtener_Impresoras()
       
    Dim i As Integer
    ' recorre las impresoras del sistema y las añade a la lista
    For i = 0 To Printers.Count - 1
        c_imp.AddItem Printers(i).DeviceName
    Next
    If c_imp.ListCount > 0 Then
      c_imp.ListIndex = 0
    Else
      MsgBox ("No se encontraron impresoras en el sistema")
      Exit Function
    End If
    
End Function
  
Public Function Establecer(Nombre_Impresora As String)
  
Dim Prt As Printer
    ' Establece la impresora que se utilizará para imprimir
    For Each Prt In Printers
        If Prt.DeviceName = Nombre_Impresora Then
            Set Printer = Prt
        End If
    Next
End Function

