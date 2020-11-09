VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form com_excel 
   BackColor       =   &H00E0E0E0&
   Caption         =   "IMPORTA COMPROBANTES de COMPRAS DE EXCEL"
   ClientHeight    =   4890
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   7995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4890
   ScaleWidth      =   7995
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Nro. Ultima fila ocupada"
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   2055
      Begin VB.TextBox t_nuf 
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Seleccionar"
      Height          =   495
      Left            =   6960
      TabIndex        =   7
      Top             =   1920
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Planilla de calculo a Importar"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   6735
      Begin VB.TextBox t_path 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   6495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   6120
      TabIndex        =   1
      Top             =   3480
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "cap003.frx":0000
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
         Picture         =   "cap003.frx":0882
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
      Top             =   4635
      Width           =   7995
      _ExtentX        =   14102
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
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3960
      Width           =   5775
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   5775
   End
   Begin VB.Label Label2 
      Caption         =   "3) Las columnas de datos que debe contener son en ese orden: FECHA/PROVEEDOR/CUIT/COMPROBANTE/TOTAL"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   7695
   End
   Begin VB.Label Label1 
      Caption         =   "2)Se debe seleccionar la planilla e indicar el ultimo numero de fila ocupado"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   7695
   End
   Begin VB.Label Label4 
      Caption         =   "1) La planilla de calculo debe tener los datos en el libro1 y comenzar en la celda A1 sin encabezados.  "
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7695
   End
End
Attribute VB_Name = "com_excel"
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
 J = MsgBox("Confirma Importar Comprobantes ", 4)
 If J = 6 Then
  Call Excel_a_Access(Val(t_nuf))
 End If
 Label5 = "Fin"
End If
End Sub

  
Private Sub Excel_a_Access(ByVal Filas As Integer)
  
  
Dim Obj_Excel As Object
Dim Obj_Hoja As Object
Dim Fila_Actual As Integer
Dim Columna_Actual As Integer
Dim Dato As Variant
  
Screen.MousePointer = vbHourglass
  
espere.Show
espere.Label1 = "Importando Comprobantes desde Excel"
espere.Refresh
'Nueva instancia de Excel
Set Obj_Excel = CreateObject("Excel.Application")
  
' Abre el libro de Excel
Obj_Excel.Workbooks.Open filename:=t_path
  
' si es la versión de Excel 97, asigna la hoja activa ( ActiveSheet )
If Val(Obj_Excel.Application.version) >= 8 Then
     Set Obj_Hoja = Obj_Excel.ActiveSheet
 Else
     Set Obj_Hoja = Obj_Excel
 End If
       
       
 'Nuevo objeto recordset
  espere.Label1 = "Eliminando registros anteriores..."
  espere.Label1.Refresh
  Set rs = New ADODB.Recordset
  q = "select * from tc"
  rs.Open q, cnib, adOpenDynamic, adLockOptimistic
    
  
  'borro archivo
  While Not rs.EOF
    rs.Delete
    rs.MoveNext
  Wend
    
 ' Recorre las filas y columnas de la hoja
  
  
  For Fila_Actual = 1 To Filas
      'Nuevo registro
        espere.Label1 = "Cargando nuevos registros...  Fila " & Fila_Actual
        espere.Label1.Refresh
           
        f = Trim$(Obj_Hoja.Cells(Fila_Actual, 1)) 'fecha
        If Len(f) >= 8 Then
           p = Trim$(Obj_Hoja.Cells(Fila_Actual, 2)) 'proveedor
           c = Trim$(Obj_Hoja.Cells(Fila_Actual, 3)) 'cuit
           cp = Trim$(Obj_Hoja.Cells(Fila_Actual, 4)) 'comprobante
           t = Trim$(Obj_Hoja.Cells(Fila_Actual, 5)) 'total
           
          If Len(c) = 13 Then 'cuit con guines
            c = Mid$(c, 1, 2) & Mid$(c, 4, 8) & Mid$(c, 13, 1)
          End If
    
          rs.AddNew
          rs("fecha") = Format$(DateValue(f), "dd/mm/yyyy")
          rs("proveedor") = Left$(p, 50)
          rs("cuit") = Val(c)
          rs("comprobante") = Left$(cp, 50)
          rs("total") = Val(t)
         rs.Update
      End If
   
   Next
       
    Unload espere
    Obj_Excel.ActiveWorkbook.Close False
    Obj_Excel.Quit
    Set Obj_Hoja = Nothing
    Set Obj_Excel = Nothing
    Screen.MousePointer = vbDefault
    MsgBox " Datos copiados ", vbInformation
  
Exit Sub
  
'Error
ErrSub:
  
Obj_Excel.ActiveWorkbook.Close False
Obj_Excel.Quit
Set Obj_Hoja = Nothing
Set Obj_Excel = Nothing
MsgBox Err.Description, vbCritical
Screen.MousePointer = vbDefault
       
End Sub
  
    
  
'Descarga los objetos y los cierra
Sub Descargar_Objetos(rst_Ado As ADODB.Recordset, cn_Ado As ADODB.Connection, _
                                      Obj_Excel As Object, Obj_Hoja As Object)
  
       
    Obj_Excel.ActiveWorkbook.Close False
    Obj_Excel.Quit
    Set Obj_Hoja = Nothing
    Set Obj_Excel = Nothing
  
End Sub


Function verifica() As Boolean
On Error GoTo errorib
verifica = True
Exit Function

errorib:
  MsgBox ("Archivo Inexistente o Invalido")
  verifica = False
  Close #1
  Exit Function
  
End Function
Private Sub btnsale_Click()
Unload Me
End Sub




Private Sub Command1_Click()
x = seleccion(t_path)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
End If

End Sub
Function seleccion(filename As String) As Boolean
On Error GoTo err_sel
CommonDialog1.Filter = "Apps *.xls"
CommonDialog1.DefaultExt = "xls"
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




