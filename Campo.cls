VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Campo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Nombre As String
Private valor As String
Private valorPorDefecto As String
Private tipoDato As String
Private verificador As IVerificador
Private ErrMessage As String
Private NombreTabla As String


Public Sub Constructor(ByVal sNombre As String, ByVal sNombreTabla As String, ByVal sValorPorDefecto As String, ByVal sTipoDato As String, ByVal oVerificador As IVerificador)
    
    Me.setNombre sNombre
    Me.setNombreTabla sNombreTabla
    Me.setValor sValorPorDefecto
    Me.setValorPorDefecto sValorPorDefecto
    Me.setTipoDato sTipoDato
    Me.setVerificador oVerificador
    
End Sub

Public Sub setNombreTabla(ByVal sNombreTabla As String)
    NombreTabla = sNombreTabla
End Sub

Public Function getNombreTabla() As String
    getNombreTabla = NombreTabla
End Function

Public Sub setErrMessage(ByVal sMessage As String)
    ErrMessage = sMessage
End Sub
Public Function getErrMessage() As String
    getErrMessage = ErrMessage
End Function


Public Sub setNombre(ByVal sNombre As String)
    Nombre = sNombre
End Sub

Public Function getNombre() As String
    getNombre = Nombre
End Function


Public Sub setValor(ByVal sValor As String)
    valor = sValor
End Sub

Public Function getValor() As String
    getValor = valor
End Function


Public Sub setValorPorDefecto(ByVal sValorPorDefecto As String)
    valorPorDefecto = sValorPorDefecto
End Sub

Public Function getValorPorDefecto() As String
    getValorPorDefecto = valorPorDefecto
End Function

Public Sub setTipoDato(ByVal sTipoDato As String)
    tipoDato = sTipoDato
End Sub

Public Function getTipoDato() As String
    getTipoDato = tipoDato
End Function

Public Sub setVerificador(ByVal oVerificador As IVerificador)
    Set verificador = oVerificador
    oVerificador.setSource Me
End Sub

Public Function getVerificador() As IVerificador
    Set getVerificador = verificador
End Function

Public Sub cambiarAValorPorDefecto()
    Me.setValor Me.getValorPorDefecto
End Sub
