VERSION 5.00
Object = "{0A6BE9FC-5039-11D5-98EC-0800460222F0}#1.0#0"; "IFEpson.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_importa_rto 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Importa Remitos desde Excel"
   ClientHeight    =   7200
   ClientLeft      =   2175
   ClientTop       =   1485
   ClientWidth     =   9375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   9375
   Begin VB.Frame Frame8 
      Caption         =   "Datos de la Hoja de Excel"
      Height          =   2415
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   7815
      Begin VB.TextBox t_filaini 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   32
         Text            =   "65536"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Utilizar id. producto proveedor"
         Height          =   615
         Left            =   6360
         TabIndex        =   31
         Top             =   960
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.ComboBox c_colcant 
         Height          =   315
         Left            =   1680
         TabIndex        =   29
         Text            =   "Combo1"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.ComboBox c_colp 
         Height          =   315
         Left            =   4920
         TabIndex        =   27
         Text            =   "Combo1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox c_coln 
         Height          =   315
         Left            =   1680
         TabIndex        =   25
         Text            =   "Combo1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox t_filafin 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4920
         MaxLength       =   5
         TabIndex        =   21
         Text            =   "65536"
         Top             =   2040
         Width           =   1095
      End
      Begin VB.ComboBox c_colf 
         Height          =   315
         Left            =   4920
         TabIndex        =   20
         Text            =   "Combo1"
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox c_colc 
         Height          =   315
         Left            =   1680
         TabIndex        =   19
         Text            =   "Combo1"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox t_nombrehoja 
         Height          =   285
         Left            =   1680
         MaxLength       =   25
         TabIndex        =   16
         Text            =   "Hoja1"
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Nro.  Fila Inicial"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Columna Cantidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   30
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Columna Id. Producto"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3240
         TabIndex        =   28
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Columna Nro.Comprobante"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Nro Fila Fin"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   22
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Columna Fecha:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   18
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Columna Cuit"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Nombre de la Hoja"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Planilla Excel con Remitos"
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   7815
      Begin VB.CommandButton Command1 
         Caption         =   "Examinar"
         Height          =   375
         Left            =   6600
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox t_path 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   6255
      End
   End
   Begin EPSON_Impresora_Fiscal.PrinterFiscal epson1 
      Left            =   0
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos para la actualizacion"
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   7815
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   23
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Text            =   "c_prov"
         Top             =   240
         Width           =   6015
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3480
         TabIndex        =   24
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Cliente:"
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
      Left            =   7680
      TabIndex        =   4
      Top             =   5880
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta058.frx":0000
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
         Picture         =   "vta058.frx":0882
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
      TabIndex        =   3
      Top             =   6945
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   13229
            MinWidth        =   13229
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "25/11/2014"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "18:53"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8280
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label9 
      Caption         =   $"vta058.frx":1104
      ForeColor       =   &H00000080&
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   5400
      Width           =   7095
   End
   Begin VB.Label Label6 
      Caption         =   "Este proceso genera el ingreso de remitos al sistema desde una  planilla excel siguiendo la siguiente estructura:"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   4920
      Width           =   7095
   End
End
Attribute VB_Name = "vta_importa_rto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984




Private Sub btnacepta_Click()
J = MsgBox("Este proceso es irreversible, confirma la  generacion de Remitos", 4)
If J = 6 Then
   J = MsgBox("Ha confirmado la generacion de Remitos. Esta seguro?", 4)
   If J = 6 Then
     If verifica Then
       Call actualiza
     End If
   End If
End If



End Sub


Function verifica() As Boolean
v = True
If Len(Dir(t_path)) = 0 Then
       MsgBox "No se ha encontrado el archivo: " & t_path, vbCritical
       v = False
End If

If Val(t_filafin) < 1 Or Val(t_filafin) > 65536 Then
   MsgBox ("La ultima fila no puede ser menor a 1 ni mayor a 65536")
   v = False
End If

If Val(t_filaini) < 1 Or Val(t_filaini) > 65536 Then
   MsgBox ("La fila Inicial no puede ser menor a 1 ni mayor a 65536")
   v = False
End If


If Val(t_filaini) > Val(t_filafin) Then
   MsgBox ("La  fila Inicial no puede ser mayor que la fila Final")
   v = False
End If


verifica = v
End Function

Sub actualiza()

    Dim o_Excel     As Object
    Dim o_Libro     As Object
    Dim o_Hoja      As Object
    Dim F        As Long
    Dim c     As Long
    Dim CODIGO As String
   Dim nombrehoja As String
   
   nombrehoja = t_nombrehoja
     
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    On Error GoTo e1
    Set o_Libro = o_Excel.Workbooks.Open(t_path, True, True, , “”)
    On Error GoTo e2
    Set o_Hoja = o_Libro.Worksheets(nombrehoja)

 
 fin = Val(t_filafin)
 
 cc = c_colc.ListIndex + 1 'cuit
 cp = c_colp.ListIndex + 1 'fecha
 a = 0
 cerr = 0
 espere.Show
  
 'On Error GoTo ERRORGRABA

 For F = 1 To fin
    espere.Label1 = F
    espere.Label1.Refresh
    actu = 0
    CODIGO = o_Hoja.Cells(F, cc).Value
    precio = o_Hoja.Cells(F, cp).Value
     
    If RTrim$(UCase(CODIGO)) = "" Then
       
    Else
      'actualiza
      Set rs = New ADODB.Recordset
      q = "select [tasa], [porc_utilidad], [dto_compra], [dto_compra2], [flete_compra], [precio_ult_compra], [dto_compra], [dto_compra2], [costoreal], [flete_compra], [pu], [precio_final], [fecha_actu_precio_venta],[envase] from a2, g4 where [id_proveedor] = " & _
      c_prov.ItemData(c_prov.ListIndex) & " and [id_prod_prov] = '" & RTrim$(UCase(CODIGO)) & "' and [cod_tasaiva] = [id_tasaiva]"
     ' MsgBox (q)
      
      On Error GoTo SALTARACTU
      rs.Open q, cn1, adOpenDynamic, adLockOptimistic
      If Not rs.EOF And Not rs.BOF Then
        tiva = rs("tasa")
        
        If Check6 Then
          precio = Val(precio) / rs("envase")
        End If
        
        If Option4 = True Then
            pf = Val(precio)
            psi = Val(precio) / (1 + (tiva / 100))
        Else
            psi = Val(precio)
            pf = Val(precio) * (1 + (tiva / 100))
        End If
      
        If Check1 Then 'precio compra
           rs("precio_ult_compra") = Format(psi, "######0.00")
           actu = 1
        End If
      
        If Check2 Then
          rs("dto_compra") = Val(t_dto1)
          rs("dto_compra2") = Val(t_dto2)
          rs("flete_compra") = Val(t_recargo)
          actu = 1
        Else
          t_dto1 = rs("dto_compra")
          t_dto2 = rs("dto_compra2")
          t_recargo = rs("flete_compra")
        End If
      
         If Check3 Then
          rs("porc_utilidad") = Val(t_utilidad)
          actu = 1
        Else
          t_utilidad = rs("porc_utilidad")
        End If
        
        If Check5 Then 'costo real
           d1 = psi * Val(t_dto1) / 100
           n1 = psi - d1
           d2 = n1 * Val(t_dto2) / 100
           n2 = n1 - d2
           rf = n2 * Val(t_recargo) / 100
           n3 = n2 + rf
           costoreal = Format(n3, "######0.00")
           rs("costoreal") = costoreal
           actu = 1
        Else
          costoreal = rs("costoreal")
        End If
      
      
        If Check4 Then  'precio de venta
          pvsi = Format(costoreal * (1 + (Val(t_utilidad) / 100)), "######0.00")
          pvf = pvsi * (1 + (tiva / 100))
           
          If Option1 = True Then
              'dos decimales
              pvf = Format(pvf, "######0.00")
          Else
             If Option2 = True Then
               'entero
               pvf = Format(pvf, "######0")
             Else
               pvf = Format(pvf, "######0.0")
             End If
          End If
          rs("pu") = pvsi
          rs("precio_final") = pvf
          rs("fecha_actu_precio_venta") = Format$(t_fecha, "dd/mm/yyyy")
          actu = 1
        End If
      
      End If
    End If
    If actu = 1 Then
      a = a + 1
      rs.Update
    End If
    Set rs = Nothing
SALTARACTU:
 
Next F
 Unload espere
 o_Libro.Close SaveChanges:=False
 o_Excel.Quit
 
'reset variables de los objetos
 Set o_Hoja = Nothing
 Set o_Libro = Nothing
 Set o_Excel = Nothing
 MsgBox ("Actualizacion Terminada. Items Actualizados: " & a)
Exit Sub
ERRORGRABA:
MsgBox ("Error en la actualizacion")
Exit Sub

e1:
MsgBox ("La planilla no existe o no puede abrirse")
Exit Sub

e2:
MsgBox ("El nombre de hoja no es correcto")
Exit Sub
End Sub




Private Sub btnsale_Click()
Unload Me
End Sub





Private Sub c_colc_LostFocus()
If c_colc.ListIndex < 0 Then
  c_colc.ListIndex = 0
End If

End Sub

Private Sub c_colp_LostFocus()
If c_colp.ListIndex < 0 Then
  c_colp.ListIndex = 1
End If

End Sub

Private Sub c_prov_LostFocus()
If c_prov.ListIndex < 0 Then
    c_prov.ListIndex = 0
 End If

End Sub








Private Sub Command1_Click()
X = seleccion(t_path)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     gen_tools.Show
  
End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call TabEnter2(Me, 4)
End If


End Sub


Function seleccion(filename As String) As Boolean
On Error GoTo err_sel


CommonDialog1.DialogTitle = "Selecciona Archivo"
CommonDialog1.InitDir = "C:\"
CommonDialog1.filename = filename
CommonDialog1.CancelError = True
CommonDialog1.ShowOpen
filename = CommonDialog1.filename
t_path = filename

Exit Function
err_sel:
t_path = filename
End Function
Private Sub Form_Load()
Call INICIALIZA2(Me)
Call carga_clientes(c_prov)
c_prov.RemoveItem 0
c_prov.AddItem "<Todos>", 0
c_prov.ListIndex = 0

For i = 1 To 26
  c_colc.AddItem Chr$(64 + i)
  c_colf.AddItem Chr$(64 + i)
  c_coln.AddItem Chr$(64 + i)
  c_colcant.AddItem Chr$(64 + i)
  c_colp.AddItem Chr$(64 + i)

Next i
c_colc.ListIndex = 1
c_colf.ListIndex = 4
c_coln.ListIndex = 3
c_colcant.ListIndex = 7
c_colp.ListIndex = 5





t_fecha = Format$(Now, "dd/mm/yyyy")
t_fecha2 = Format$(Now, "dd/mm/yyyy")
t_filafin = "65536"
t_filaini = "1"
t_nombrehoja = "Hoja1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload vta_facturacion1
Unload vta_facturacion2
Unload vta_selremitos
Unload vta_clientes
Unload vta_formapago
End Sub











Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
 If Not IsDate(t_fecha) Then
  t_fecha = Format$(Now, "dd/mm/yyyy")
 Else
  t_fecha = Format$(t_fecha, "dd/mm/yyyy")
 End If
Else
 t_fecha = Format$(Now, "dd/mm/yyyy")
End If

End Sub







Private Sub t_nombrehoja_LostFocus()
If t_nombrehoja = "" Then
   t_nombrehoja = "Hoja1"
End If

End Sub
