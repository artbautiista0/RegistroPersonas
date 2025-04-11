VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsGestorPersonas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ClsGestorPersonas.cls
Option Explicit

Private lista As Collection

Private Sub Class_Initialize()
    Set lista = New Collection
End Sub

Public Sub Agregar(persona As ClsPersona)
    lista.Add persona 'Agregamos a la lista a la persona
End Sub

Public Function ObtenerTodas() As Collection
    Set ObtenerTodas = lista
End Function

