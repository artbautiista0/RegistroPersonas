VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ClsPersona.cls
Option Explicit
'Comentario
Private mNombre As String
Private mEdad As Integer

Public Property Let nombre(ByVal value As String)
    mNombre = value
End Property

Public Property Get nombre() As String
    nombre = mNombre
End Property

Public Property Let Edad(ByVal value As Integer)
    mEdad = value
End Property

Public Property Get Edad() As Integer
    Edad = mEdad
End Property

Public Function Descripcion() As String
    Descripcion = mNombre & " - " & mEdad & " años"
End Function

