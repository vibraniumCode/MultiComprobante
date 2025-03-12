Attribute VB_Name = "modConexion"
Option Explicit

' Módulo para manejar la conexión global
Public conn As New ADODB.Connection

Public Sub ConectarBD()
    ' Cadena de conexión
    Dim strConn As String
    strConn = "Provider=SQLOLEDB;Data Source=facturacion_neumaticos.mssql.somee.com;" & _
              "Initial Catalog=facturacion_neumaticos;User ID=mlopez_cliente_2_SQLLogin_2;" & _
              "Password=khuwpknwob;Persist Security Info=False;" & _
              "TrustServerCertificate=True"

    ' Abrir conexión
    On Error GoTo ErrHandler
    If conn.State = 0 Then
        conn.Open strConn
    End If
    Exit Sub

ErrHandler:
    MsgBox "Error al conectar: " & Err.Description, vbCritical, "Error"
End Sub

Public Sub DesconectarBD()
    ' Cierra la conexión si está abierta
    If conn.State = 1 Then
        conn.Close
    End If
End Sub


