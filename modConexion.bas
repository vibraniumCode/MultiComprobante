Attribute VB_Name = "modConexion"
Option Explicit

' M�dulo para manejar la conexi�n global
Public conn As New ADODB.Connection

Public Sub ConectarBD()
    
    Dim rutaIni As String
    Dim strConn As String
    
    ' Obtener ruta del archivo INI (en el mismo directorio que la aplicaci�n)
    rutaIni = App.Path & "\login.ini"
    
    ' Leer configuraci�n
    Dim Provider As String
    Dim DataSource As String
    Dim InitialCatalog As String
    Dim UserID As String
    Dim Password As String
    
    Provider = LeerIni("DatabaseConfig", "Provider", rutaIni)
    DataSource = LeerIni("DatabaseConfig", "DataSource", rutaIni)
    InitialCatalog = LeerIni("DatabaseConfig", "InitialCatalog", rutaIni)
    UserID = LeerIni("DatabaseConfig", "UserID", rutaIni)
    Password = LeerIni("DatabaseConfig", "Password", rutaIni)
    
    ' Armar cadena de conexi�n
    strConn = "Provider=" & Provider & ";" & _
              "Data Source=" & DataSource & ";" & _
              "Initial Catalog=" & InitialCatalog & ";" & _
              "User ID=" & UserID & ";" & _
              "Password=" & Password & ";" & _
              "Persist Security Info=False;" & _
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


