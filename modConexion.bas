Attribute VB_Name = "modConexion"
Option Explicit

' M�dulo para manejar la conexi�n global
Public conn As New ADODB.Connection

Public Sub ConectarBD()
    ' Cadena de conexi�n
    Dim strConn As String
    strConn = "Provider=SQLOLEDB;Data Source=facturacion_neumaticos.mssql.somee.com;" & _
              "Initial Catalog=facturacion_neumaticos;User ID=mlopez_cliente_2_SQLLogin_2;" & _
              "Password=khuwpknwob;Persist Security Info=False;" & _
              "TrustServerCertificate=True"

    ' Abrir conexi�n
    On Error GoTo ErrHandler
    If conn.State = 0 Then
        conn.Open strConn
    End If
    Exit Sub

ErrHandler:
    MsgBox "Error al conectar: " & Err.Description, vbCritical, "Error"
End Sub

Public Sub DesconectarBD()
    ' Cierra la conexi�n si est� abierta
    If conn.State = 1 Then
        conn.Close
    End If
End Sub


