VERSION 5.00
Begin VB.Form cgr_info_periodo 
   Caption         =   "Seleccion de Periodo Contable a Trabajar"
   ClientHeight    =   960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Periodo Contable Actual"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton Command1 
         Caption         =   "Cambiar"
         Height          =   375
         Left            =   4320
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox c_periodo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   240
         Width           =   4095
      End
   End
End
Attribute VB_Name = "cgr_info_periodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If c_periodo.ListIndex < 0 Then
  c_periodo.ListIndex = buscaindice(c_periodo, para.id_periodo_contable)
End If
j = MsgBox("Confirma Seleccionar Periodo para Trabajar", 4)
If j = 4 Then
   Set rs = New ADODB.Recordset
   q = "select * from g0 where [sucursal] = 0"
   rs.Open q, cn1, adOpenDynamic, adLockOptimistic
   If Not rs.EOF And rs.BOF Then
     rs("id_periodo_contable") = c_periodo.ItemData(c_periodo.ListIndex)
     rs.Update
   
     para.id_periodo_contable = c_periodo.ItemData(c_periodo.ListIndex)
   End If
End If
End Sub

Private Sub Form_Load()
Call carga_periodos(c_periodo)
c_periodo.ListIndex = buscaindice(c_periodo, para.id_periodo_contable)

End Sub
