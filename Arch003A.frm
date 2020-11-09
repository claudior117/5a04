VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form abm_solmat 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PEDIDOS de MATERIALES PENDIENTES"
   ClientHeight    =   8865
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12255
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8865
   ScaleWidth      =   12255
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar"
      Height          =   1695
      Left            =   240
      TabIndex        =   4
      Top             =   0
      Width           =   11535
      Begin VB.ComboBox c_tipo 
         Height          =   315
         ItemData        =   "Arch003A.frx":0000
         Left            =   8760
         List            =   "Arch003A.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1320
         Width           =   2535
      End
      Begin VB.ComboBox c_Obra 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1320
         Width           =   5535
      End
      Begin VB.ComboBox c_estado 
         Height          =   315
         ItemData        =   "Arch003A.frx":003B
         Left            =   1680
         List            =   "Arch003A.frx":0048
         TabIndex        =   16
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox t_fecha2 
         Height          =   285
         Left            =   8760
         MaxLength       =   10
         TabIndex        =   14
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox t_fecha1 
         Height          =   285
         Left            =   8760
         MaxLength       =   10
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox c_usuario 
         Height          =   315
         Left            =   8760
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   2535
      End
      Begin VB.ComboBox c_prod 
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         Top             =   960
         Width           =   5535
      End
      Begin VB.ComboBox c_estado2 
         Height          =   315
         ItemData        =   "Arch003A.frx":0077
         Left            =   1680
         List            =   "Arch003A.frx":0084
         TabIndex        =   8
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label Label8 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   7440
         TabIndex        =   21
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Obra:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha hasta:"
         Height          =   375
         Left            =   7440
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha desde:"
         Height          =   375
         Left            =   7440
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Usuario"
         Height          =   375
         Left            =   7440
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Estado Recepcion"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Producto"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Estado Pedido"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   1
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "Arch003A.frx":00B3
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
         Picture         =   "Arch003A.frx":0935
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1320
      Top             =   7320
      Visible         =   0   'False
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   661
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   2
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "claudio"
      Password        =   "0969"
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8610
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   35278
            MinWidth        =   35278
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Arch003A.frx":11B7
      Height          =   5295
      Left            =   240
      TabIndex        =   17
      Top             =   1800
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9340
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      TabAction       =   2
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "PEDIDOS DE MATERIA PRIMA y PRODUCTOS"
      ColumnCount     =   14
      BeginProperty Column00 
         DataField       =   "num_referencia"
         Caption         =   "Ref."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "id_producto"
         Caption         =   "Id.Prod."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "00000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "detalle"
         Caption         =   "Producto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "total_pedido"
         Caption         =   "Pedido"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "total_oc"
         Caption         =   "En O.C."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "total_recibido"
         Caption         =   "Recibido"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "observaciones"
         Caption         =   "Observaciones"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "pro_04.id_usuario"
         Caption         =   "Id. usuario"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "pro_04.id_obra"
         Caption         =   "Id. Obra"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "fecha_esperado"
         Caption         =   "Esperado"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "fecha"
         Caption         =   "Pedido"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "usuario"
         Caption         =   "Ingresado por:"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "descripcion"
         Caption         =   "Obra"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "total_facturado"
         Caption         =   "Facturado"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   5
         SizeMode        =   1
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            Object.Visible         =   -1  'True
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
         EndProperty
         BeginProperty Column13 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "abm_solmat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim gquery As String

Private Sub btnacepta_Click()
If c_tipo.ListIndex = 1 Then
   Call limpia
   DataGrid1.SetFocus
Else
   If para.id_grupo_modulo_compras > 7 Then
        Call limpia
        DataGrid1.SetFocus
   Else
    MsgBox ("Operacion No autorizada")
   End If
End If
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub



Private Sub c_estado_LostFocus()
If c_estado.ListIndex < 0 Then
  c_estado.ListIndex = 1
End If
End Sub

Private Sub DataGrid1_GotFocus()
Me.StatusBar1.Panels.Item(1) = "[ENTER]Detalle - [F3] Arma Sol.Cot. - [F4] Modificacion Manual - [F6]Arma Rto. - [F7] Imprime - [F8] Borra Ingreso - [F9] Arma O.C. - [F5] Imp. Comp."

End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
Set rs = Adodc1.Recordset
gcolumna = ColIndex
If rs.Sort <> DataGrid1.Columns(ColIndex).DataField & " ASC" Then
   rs.Sort = DataGrid1.Columns(ColIndex).DataField & " ASC"
Else
   rs.Sort = DataGrid1.Columns(ColIndex).DataField & " DESC"
End If
Set rs = Nothing

End Sub



Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF9 Then
 If para.id_grupo_modulo_compras > 6 Then
     J = MsgBox("Confirma Armar O.C. con los materiales seleccionados", 4)
     If J = 6 Then
       Load ABM_OC
       Call armaoc
       Unload Me
     End If
 End If
End If


If KeyCode = vbKeyF4 Then
 If para.id_grupo_modulo_compras >= 8 Then
    Call armamanual
    
      
 End If
End If


If KeyCode = vbKeyF5 Then
 If para.id_grupo_modulo_compras >= 6 Then
  J = MsgBox("Imprime Comprobante", 4)
      If J = 6 Then
         Set rs2 = New ADODB.Recordset
         q = "select [num_int] from pro_02 where [num_referencia] = " & DataGrid1.Columns(0).CellValue(DataGrid1.Bookmark)
         rs2.MaxRecords = 1
         rs2.Open q, cn1
         If Not rs2.EOF And Not rs2.BOF Then
           Set cl_compprod = New comprobantes_produccion
           cl_compprod.cargar2 (rs2("num_int"))
           If cl_compprod.numint > 0 Then
              cl_compprod.imprimir
         
              J = MsgBox("Imprime Parte para Taller", 4)
              If J = 6 Then
                k = InputBox$("Cantidad de Copias", "Imprimir", 1)
                If Val(k) > 0 Then
                cl_compprod.imprimesmt (Val(J))
              End If
           End If
          End If
          Set rs2 = Nothing
        End If
    End If
 End If
End If

If KeyCode = vbKeyF6 Then
 If para.id_grupo_modulo_compras > 6 Then
     J = MsgBox("Confirma Armar Remito con los materiales seleccionados", 4)
     If J = 6 Then
       Load ABM_COMP_COMPRA
       Call armarto
       Unload Me
     End If
 End If
End If

If KeyCode = vbKeyF3 Then
 If para.id_grupo_modulo_compras > 6 Then
     J = MsgBox("Confirma Armar Solicitud de Cotizacion con los materiales seleccionados", 4)
     If J = 6 Then
       Load ABM_cotizacion
       Call armacotiz
       Unload Me
     End If
 End If
End If

If KeyCode = vbKeyF8 Then
 If para.id_grupo_modulo_actual >= 6 Or para.id_grupo_modulo_compras >= 8 Then
     J = MsgBox("Confirma Borrar producto de requisicion y todos sus movimientos", 4)
     If J = 6 Then
        nr = DataGrid1.Columns(0).CellValue(DataGrid1.Bookmark)
        If nr > 0 Then
          Set cl_compprod = New comprobantes_produccion
          cl_compprod.borraproductorequisicion (nr)
          Set cl_compprod = Nothing
          MsgBox ("Operacion terminada")
          Call limpia
          DataGrid1.SetFocus
        End If
     
     
     
     End If
 End If
End If



If KeyCode = vbKeyF7 Then
 If para.id_grupo_modulo_compras > 4 Then
  J = MsgBox("Prepare Impresora y Confirme", 4)
    If J = 6 Then
      Dim c(15) As Double
      c(0) = 1
      c(1) = 2
      c(2) = 3
      c(3) = 4
      c(4) = 5
      c(5) = 6
      c(6) = 9
      c(7) = 10
      c(8) = 11
      c(9) = 12
      
      
      For i = 10 To 14
        c(i) = -1
      Next i
      Call imprimedatagrid2(DataGrid1, c(), "PEDIDO DE MATERIALES", "Estado: " & c_estado, "Periodo: " & t_fecha1 & " al " & t_fecha2, "Producto: " & c_producto, 60, 7, True, False, gquery, "H")
     End If
         
  End If
 End If

End Sub

Sub armaoc()
'genero un arrary para guardar por cada fila de la oc la cadena
'de ordenes de req. que la compnen

i = 1
'FIXIT: Declare 'varBmk' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Dim varBmk As Variant
For Each varBmk In DataGrid1.SelBookmarks
      Adodc1.Recordset.Bookmark = varBmk
      ip = DataGrid1.Columns(1).CellValue(DataGrid1.Bookmark)
      d = DataGrid1.Columns(2).CellValue(DataGrid1.Bookmark)
      b = Format$(1, "######0.00")
      cu = Format$(Val(DataGrid1.Columns(3).CellValue(DataGrid1.Bookmark)) - Val(DataGrid1.Columns(4).CellValue(DataGrid1.Bookmark)), "######0.00")
     crq = Format$(Val(DataGrid1.Columns(0).CellValue(DataGrid1.Bookmark)), "00000000")
       o = DataGrid1.Columns(6).CellValue(DataGrid1.Bookmark)
       idobra = DataGrid1.Columns(8).CellValue(DataGrid1.Bookmark)
       obra = DataGrid1.Columns(12).CellValue(DataGrid1.Bookmark)
       r = ABM_OC.msf1.Rows
      
      If Val(cu) > 0 Then
        ABM_OC.msf1.AddItem r & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & crq & Chr(9) & o & Chr(9) & cu & Chr(9) & "0.00" & Chr(9) & "  " & Chr(9) & Format$(para.tasageneral, "#0.00") & Chr(9) & "0.00" & Chr(9) & obra & Chr(9) & idobra
      End If
     
    
     
 Next
ABM_OC.Show
End Sub
Sub armamanual()
If para.id_grupo_modulo_compras >= 8 Then
 If DataGrid1.Bookmark > 0 Then
  ip = Val(DataGrid1.Columns(0).CellValue(DataGrid1.Bookmark))
  Set rs2 = New ADODB.Recordset
  q = "select * from pro_04 where [num_referencia] = " & ip
  rs2.Open q, cn1
  If Not rs2.EOF And Not rs2.BOF Then
  Load prod_manual
  prod_manual.t_id = rs2("num_referencia")
  prod_manual.t_descripcion = rs2("detalle")
  prod_manual.t_p = rs2("total_pedido")
  prod_manual.t_r = rs2("total_recibido")
  prod_manual.t_o = rs2("total_oc")
  prod_manual.t_f = rs2("total_facturado")
  prod_manual.t_idprod = DataGrid1.Columns(1).CellValue(DataGrid1.Bookmark)
  prod_manual.t_descprod = DataGrid1.Columns(2).CellValue(DataGrid1.Bookmark)
  prod_manual.t_idobra = rs2("id_obra")
  
  prod_manual.Show
 End If
 Set rs2 = Nothing
 End If
End If


End Sub
Sub armarto()
'genero un arrary para guardar por cada fila de la oc la cadena
'de ordenes de req. que la compnen

i = 1
'FIXIT: Declare 'varBmk' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Dim varBmk As Variant
For Each varBmk In DataGrid1.SelBookmarks
      Adodc1.Recordset.Bookmark = varBmk
      ip = DataGrid1.Columns(1).CellValue(DataGrid1.Bookmark)
      d = DataGrid1.Columns(2).CellValue(DataGrid1.Bookmark)
      b = Format$(1, "######0.00")
      cu = Format$(Val(DataGrid1.Columns(13).CellValue(DataGrid1.Bookmark)) - Val(DataGrid1.Columns(4).CellValue(DataGrid1.Bookmark)), "######0.00")
     crq = Format$(Val(DataGrid1.Columns(0).CellValue(DataGrid1.Bookmark)), "00000000")
       o = DataGrid1.Columns(6).CellValue(DataGrid1.Bookmark)
       idobra = DataGrid1.Columns(8).CellValue(DataGrid1.Bookmark)
       obra = DataGrid1.Columns(12).CellValue(DataGrid1.Bookmark)
       r = ABM_COMP_COMPRA.msf1.Rows
      
      If Val(cu) > 0 Then
        ABM_COMP_COMPRA.msf1.AddItem r & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & cu & Chr(9) & "0.00" & Chr(9) & Format$(para.tasageneral, "#0.00") & Chr(9) & "0.00" & Chr(9) & "0.00" & Chr(9) & crq
      End If
     
    
     
 Next
ABM_COMP_COMPRA.c_tipocomp.ListIndex = buscaindice(ABM_COMP_COMPRA.c_tipocomp, 45)
ABM_COMP_COMPRA.Show

End Sub

Sub armacotiz()
'genero un arrary para guardar por cada fila de la oc la cadena
'de ordenes de req. que la compnen

i = 1
'FIXIT: Declare 'varBmk' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Dim varBmk As Variant
For Each varBmk In DataGrid1.SelBookmarks
      Adodc1.Recordset.Bookmark = varBmk
      ip = DataGrid1.Columns(1).CellValue(DataGrid1.Bookmark)
      d = DataGrid1.Columns(2).CellValue(DataGrid1.Bookmark)
      b = Format$(1, "######0.00")
      cu = Format$(Val(DataGrid1.Columns(3).CellValue(DataGrid1.Bookmark)) - Val(DataGrid1.Columns(4).CellValue(DataGrid1.Bookmark)), "######0.00")
     crq = Format$(Val(DataGrid1.Columns(0).CellValue(DataGrid1.Bookmark)), "00000000")
       o = DataGrid1.Columns(6).CellValue(DataGrid1.Bookmark)
       idobra = DataGrid1.Columns(8).CellValue(DataGrid1.Bookmark)
       obra = DataGrid1.Columns(12).CellValue(DataGrid1.Bookmark)
       r = ABM_cotizacion.msf1.Rows
      
      If Val(cu) > 0 Then
        ABM_cotizacion.msf1.AddItem r & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & o & Chr(9) & cu & Chr(9) & " " & Chr(9) & obra & Chr(9) & idobra
      End If
     
    
     
 Next
ABM_cotizacion.Show

End Sub
Function busca_prod_oc(cp As Long) As Long
  i = 1
  r = 0
  While i <= ABM_OC.msf1.Rows - 1
     If cp = Val(ABM_OC.msf1.TextMatrix(i, 1)) Then
        r = i
        i = ABM_OC.msf1.Rows
     End If
     i = i + 1
  Wend
  busca_prod_oc = r
End Function

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    prod_detalle_pedidos.t_referencia = DataGrid1.Columns(0).CellValue(DataGrid1.Bookmark)
    prod_detalle_pedidos.t_idproducto = DataGrid1.Columns(1).CellValue(DataGrid1.Bookmark)
    prod_detalle_pedidos.t_producto = DataGrid1.Columns(2).CellValue(DataGrid1.Bookmark)
    prod_detalle_pedidos.Show
End If
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub
Sub limpia()
Dim q As String
 q = "select  *   from pro_04, g1, a4 where pro_04.[id_usuario] = g1.[id_usuario] and pro_04.[id_obra] = a4.[id_obra]  "
 c = " and "


If c_tipo.ListIndex > 0 Then
   q = q & c & " [tipo04] = " & c_tipo.ListIndex
End If

'estado pedido
Select Case c_estado.ListIndex
  Case Is = 1 'completo
    q = q & c & " [total_oc] >= [total_pedido]"
    c = " and "
  Case Is = 2 'incompleto
    q = q & c & " [total_oc] < [total_pedido]"
    c = " and "

End Select

'estado recepcion
Select Case c_estado2.ListIndex
  Case Is = 1 'completo
    q = q & c & " [total_recibido] >= [total_pedido] and [total_recibido] >= [total_facturado]"
    c = " and "
  Case Is = 2 'incompleto
    q = q & c & " ([total_recibido] < [total_pedido] or  [total_recibido] < [total_facturado])"
    c = " and "

End Select


Select Case c_prod.ListIndex
 Case Is > 0
    q = q & c & " [id_producto] = " & c_prod.ItemData(c_prod.ListIndex)
    c = " and "
 Case Is < 0
  If c_prod <> "" Then
    q = q & c & " [detalle] like '%" & c_prod & "%'"
    c = " and "
  End If
End Select

If c_usuario.ListIndex > 0 Then
  q = q & c & " pro_04.[id_usuario] = " & c_usuario.ListIndex
  c = " and "
End If

If t_fecha1 <> "" Then
  q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha1 & "') "
  c = " and "
End If

If t_fecha2 <> "" Then
  q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "') "
  c = " and "
End If

If c_Obra.ListIndex > 0 Then
  q = q & c & " pro_04.[id_obra] = " & c_Obra.ItemData(c_Obra.ListIndex)
  c = " and "
End If



q = q & " order by [num_referencia]"

gquery = q
Call conectaradodc(Adodc1, q, cn1)
DataGrid1.Refresh

End Sub

Private Sub Form_Load()
Call carga_productos(c_prod)
c_prod.AddItem "<Todos>", 0
c_prod.ListIndex = 0

Call carga_usuarios(c_usuario)
c_usuario.AddItem "<Todos>", 0
c_usuario.ListIndex = 0

Call carga_obras(c_Obra, "A")
c_Obra.AddItem "<Todas>", 0
c_Obra.ListIndex = 0


Call barraesag(Me)
Load ver_PROD_oc

c_estado.ListIndex = 2
c_estado2.ListIndex = 0

Load prod_detalle_pedidos
gfilas = 0
c_tipo.ListIndex = 1
End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload prod_detalle_pedidos

End Sub

Private Sub t_fecha1_LostFocus()
If Not IsDate(t_fecha1) And t_fecha1 <> "" Then
   t_fecha1 = ""
End If
End Sub

Private Sub t_fecha2_LostFocus()
If Not IsDate(t_fecha2) And t_fecha2 <> "" Then
   t_fecha2 = ""
End If

End Sub
