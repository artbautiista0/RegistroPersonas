VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5970
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstPersonas 
      Height          =   1815
      Left            =   2520
      TabIndex        =   5
      Top             =   480
      Width           =   3135
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   360
      Left            =   960
      TabIndex        =   2
      Top             =   1920
      Width           =   990
   End
   Begin VB.TextBox txtEdad 
      Height          =   405
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtNombre 
      Height          =   405
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblEdad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edad"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form1.frm
Option Explicit

Dim validador As ClsValidador
Dim gestor As ClsGestorPersonas

Private Sub Form_Load()
    Set validador = New ClsValidador
    Set gestor = New ClsGestorPersonas
End Sub

Private Sub cmdAgregar_Click()
    Dim nombre As String
    Dim edadTexto As String
    Dim persona As ClsPersona
    
    nombre = txtNombre.Text
    edadTexto = txtEdad.Text
    
    If Not validador.EsTextoValido(nombre) Then
        MsgBox "El nombre es obligatorio.", vbExclamation
        Exit Sub
    End If
    
    If Not validador.EsNumeroValido(edadTexto) Then
        MsgBox "Edad inválida.", vbExclamation
        Exit Sub
    End If
    
    Set persona = New ClsPersona
    persona.nombre = nombre
    persona.Edad = Val(edadTexto)
    
    gestor.Agregar persona
    
    lstPersonas.AddItem persona.Descripcion
    
    txtNombre.Text = ""
    txtEdad.Text = ""
    txtNombre.SetFocus
End Sub

