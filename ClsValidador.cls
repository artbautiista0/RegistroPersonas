VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsValidador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ClsValidador.cls
Option Explicit

Public Function EsTextoValido(texto As String) As Boolean
    EsTextoValido = (Len(Trim(texto)) > 0)
End Function

Public Function EsNumeroValido(texto As String) As Boolean
    EsNumeroValido = IsNumeric(texto) And Val(texto) > 0
End Function

