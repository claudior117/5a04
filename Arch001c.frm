VERSION 5.00
Begin VB.Form abm_comp_compra3 
   Caption         =   "Graba Comprobante"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Opciones de Actualizacion de estructura de Precios"
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.OptionButton Option4 
         Caption         =   "No actualiza Nada"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   1440
         Width           =   3615
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Solo actualiza Precio de VENTA"
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   1080
         Width           =   3615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Solo actualiza Precio de COMPRA"
         Height          =   255
         Left            =   600
         TabIndex        =   2
         Top             =   720
         Width           =   3615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Actualiza Precio de COMPRA y de VENTA"
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
   End
End
Attribute VB_Name = "abm_comp_compra3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
