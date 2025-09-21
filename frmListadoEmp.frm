VERSION 5.00
Begin VB.Form frmListadoEmp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresion de comprobantes"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5550
   Icon            =   "frmListadoEmp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.ComboBox cboEmpresas 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   5055
      End
      Begin VB.CommandButton btnCargar 
         Caption         =   "&Imprimir comprobante"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Picture         =   "frmListadoEmp.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmListadoEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub CargarEmpresas()
    Dim rs As New ADODB.Recordset
    
    ' Conectar a la base de datos
    Call ConectarBD
    
    On Error GoTo ErrHandler
    
    cboEmpresas.Clear
    
    rs.Open "SELECT id, empresa FROM empresas ORDER BY empresa", conn, adOpenStatic, adLockReadOnly
    
    ' Cargar los meses desde la base de datos al ComboBox
    Do While Not rs.EOF
        ' Puedes guardar el ID en ItemData si querés usarlo después
        cboEmpresas.AddItem rs("empresa")
        cboEmpresas.ItemData(cboEmpresas.NewIndex) = rs("id")
        rs.MoveNext
    Loop
    
    If cboEmpresas.ListCount > 0 Then
        cboEmpresas.ListIndex = 0
    End If
    
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
    Exit Sub
    
ErrHandler:
    MsgBox "Error al cargar el listado de combustible: " & Err.Description, vbCritical, "Error"
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD

End Sub

Private Sub btnCargar_Click()
Dim idEmpresa As Integer

idEmpresa = cboEmpresas.ItemData(cboEmpresas.ListIndex)

Select Case idEmpresa
    Case 1
        Call ImpresionPapeleraBP
    Case 2
        Call ImpresionPCdelCentro
    Case 3
        Call ImpresionVagol
    Case 4
        Call ImpresionTransItalia
    Case 5
        Call ImpresionMadera
    Case 6
        Call ImpresionNeumaticosArturo
    Case 7
        Call ImpresionFederalParts
    Case 8
        Call ImpresionMaterialNuciari
    Case 9
        Call ImpresionFletesRattaro
    Case 11
        Call ImpresionFerreteriaSabul
    Case 12
        Call ImpresionMecanica
    Case 13
        Call ImpresionTrole
    Case 14
        Call ImpresionPyramar
End Select
End Sub

Private Sub Form_Load()
Call CargarEmpresas
End Sub

Private Sub ImpresionNeumaticosArturo()
    GenerarComprobante Facturacion.txtSubtotal.Text, Facturacion.txtIva.Text, Facturacion.txtTotal.Text, Facturacion
End Sub

Private Sub ImpresionMecanica()
    GenerarComprobanteMecanica Facturacion.txtSubtotal.Text, Facturacion.txtIva.Text, Facturacion.txtTotal.Text, Facturacion
End Sub

Private Sub ImpresionFederalParts()
    GenerarComprobante2 Facturacion.txtSubtotal.Text, Facturacion.txtIva.Text, Facturacion.txtTotal.Text, Facturacion
End Sub

Private Sub ImpresionMaterialNuciari()
    GenerarComprobante3 Facturacion.txtSubtotal.Text, Facturacion.txtIva.Text, Facturacion.txtTotal.Text, Facturacion
End Sub

Private Sub ImpresionPapeleraBP()
Dim rs As New ADODB.Recordset
Dim sql As String

Call ConectarBD
sql = "select descripcion, cantidad, precio_unitario as punitario, precio_neto as neto from productos_ventas where factura = " & nroFactura
rs.Open sql, conn, adOpenStatic, adLockReadOnly

If Not rs.EOF Then
    With papeleraBP
        Set .DataSource = rs
        .Sections("Sección4").Controls("Etiqueta19").Caption = Facturacion.txtNombre.Text
        .Sections("Sección4").Controls("Etiqueta29").Caption = Facturacion.txtDireccion.Text
        .Sections("Sección4").Controls("Etiqueta31").Caption = Facturacion.txtCuit.Text
        .Sections("Sección4").Controls("Etiqueta21").Caption = Facturacion.txtLocalidad.Text
        .Sections("Sección4").Controls("Etiqueta30").Caption = Facturacion.txtProvincia.Text
        .Sections("Sección4").Controls("factura").Caption = Format(Facturacion.txtPunto.Text, "0000") & "-" & Format(nroFactura, "00000000")
        .Sections("Sección4").Controls("fecha").Caption = Facturacion.fecEmision.Value
        
        .Sections("Sección3").Controls("subtotal").Caption = Facturacion.txtSubtotal
        .Sections("Sección3").Controls("iva").Caption = Facturacion.txtIva
        .Sections("Sección3").Controls("total").Caption = Facturacion.txtTotal
        .Sections("Sección3").Controls("totalp").Caption = Facturacion.txtTotal
        
        .Sections("Sección3").Controls("vencimiento").Caption = "Fecha de Vencimiento: " & Format(DateAdd("d", 7, Facturacion.fecEmision.Value), "dd/mm/yyyy")
    
        .Show vbModal
    End With
Else
    MsgBox "No hay datos para mostrar.", vbInformation
End If
End Sub

Private Sub ImpresionPCdelCentro()
Dim rs As New ADODB.Recordset
Dim sql As String

Call ConectarBD
sql = "select descripcion, cantidad, precio_unitario as punitario, precio_neto as neto from productos_ventas where factura = " & nroFactura
rs.Open sql, conn, adOpenStatic, adLockReadOnly

If Not rs.EOF Then
    With pcCentro
        Set .DataSource = rs
            .Sections("Sección4").Controls("cliente").Caption = Facturacion.txtNombre.Text
            .Sections("Sección4").Controls("direccion").Caption = Facturacion.txtDireccion.Text
            .Sections("Sección4").Controls("localidad").Caption = Facturacion.txtCP.Text + " " + Facturacion.txtProvincia.Text
            .Sections("Sección4").Controls("cuit").Caption = Facturacion.txtCuit.Text
            .Sections("Sección4").Controls("factura").Caption = Format(Facturacion.txtPunto.Text, "0000") & "-" & Format(nroFactura, "00000000")
            .Sections("Sección4").Controls("fecha").Caption = "Fecha: " & Facturacion.fecEmision.Value
            
            .Sections("Sección5").Controls("subtotal").Caption = Facturacion.txtSubtotal
            .Sections("Sección5").Controls("iva").Caption = Facturacion.txtIva
            .Sections("Sección5").Controls("total").Caption = Facturacion.txtTotal
            
            .Sections("Sección5").Controls("vencimiento").Caption = "Fecha de Vencimiento: " & Format(DateAdd("d", 7, Facturacion.fecEmision.Value), "dd/mm/yyyy")
            
            .Show vbModal
    End With
Else
    MsgBox "No hay datos para mostrar.", vbInformation
End If
End Sub

Private Sub ImpresionVagol()
Dim rs As New ADODB.Recordset
Dim sql As String

Call ConectarBD
sql = "select descripcion, cantidad, precio_unitario as punitario, precio_neto as neto from productos_ventas where factura = " & nroFactura
rs.Open sql, conn, adOpenStatic, adLockReadOnly

If Not rs.EOF Then
    With vagol
        Set .DataSource = rs
        .Sections("Sección4").Controls("cliente").Caption = Facturacion.txtNombre.Text
        .Sections("Sección4").Controls("direccion").Caption = Facturacion.txtDireccion.Text + " " + Facturacion.txtLocalidad.Text + " " + Facturacion.txtProvincia.Text
        .Sections("Sección4").Controls("Etiqueta14").Caption = Facturacion.txtCuit.Text
        .Sections("Sección4").Controls("factura").Caption = Format(Facturacion.txtPunto.Text, "0000") & "-" & Format(nroFactura, "00000000")
        .Sections("Sección4").Controls("fecha").Caption = "Fecha: " & Facturacion.fecEmision.Value
        
        .Sections("Sección5").Controls("subtotal").Caption = Facturacion.txtSubtotal
        .Sections("Sección5").Controls("iva").Caption = Facturacion.txtIva
        .Sections("Sección5").Controls("total").Caption = Facturacion.txtTotal
        
        .Sections("Sección5").Controls("vencimiento").Caption = "Fecha de Vencimiento: " & Format(DateAdd("d", 7, Facturacion.fecEmision.Value), "dd/mm/yyyy")
        
        .Show vbModal
    End With
Else
    MsgBox "No hay datos para mostrar.", vbInformation
End If
End Sub

Private Sub ImpresionTransItalia()
Dim rs As New ADODB.Recordset
Dim sql As String

Call ConectarBD
sql = "select descripcion, cantidad, precio_unitario as punitario, precio_neto as neto from productos_ventas where factura = " & nroFactura
rs.Open sql, conn, adOpenStatic, adLockReadOnly

If Not rs.EOF Then
    With transporteItalia
        Set .DataSource = rs
        .Sections("Sección4").Controls("cliente").Caption = Facturacion.txtNombre.Text
        .Sections("Sección4").Controls("direccion").Caption = Facturacion.txtDireccion.Text
        .Sections("Sección4").Controls("localidad").Caption = Facturacion.txtCP.Text + " " + Facturacion.txtLocalidad.Text + " " + Facturacion.txtProvincia.Text
        .Sections("Sección4").Controls("cuit").Caption = Facturacion.txtCuit.Text
        .Sections("Sección4").Controls("factura").Caption = Format(Facturacion.txtPunto.Text, "0000") & "-" & Format(nroFactura, "00000000")
        .Sections("Sección4").Controls("fecha").Caption = "Fecha: " & Facturacion.fecEmision.Value
        
        .Sections("Sección5").Controls("subtotal").Caption = Facturacion.txtSubtotal
        .Sections("Sección5").Controls("iva").Caption = Facturacion.txtIva
        .Sections("Sección5").Controls("total").Caption = Facturacion.txtTotal
        
        .Sections("Sección5").Controls("vencimiento").Caption = "Fecha de Vencimiento: " & Format(DateAdd("d", 7, Facturacion.fecEmision.Value), "dd/mm/yyyy")
        
        .Show vbModal
    End With
Else
    MsgBox "No hay datos para mostrar.", vbInformation
End If
End Sub

Private Sub ImpresionMadera()
Dim rs As New ADODB.Recordset
Dim sql As String

Call ConectarBD
sql = "select descripcion, cantidad, precio_unitario as punitario, precio_neto as neto from productos_ventas where factura = " & nroFactura
rs.Open sql, conn, adOpenStatic, adLockReadOnly

If Not rs.EOF Then
    With mymMadera
        Set .DataSource = rs
        .Sections("Sección4").Controls("cliente").Caption = Facturacion.txtNombre.Text
        .Sections("Sección4").Controls("direccion").Caption = Facturacion.txtDireccion.Text + " " + Facturacion.txtCP.Text + " " + Facturacion.txtLocalidad.Text + " " + Facturacion.txtProvincia.Text
        .Sections("Sección4").Controls("cuit").Caption = Facturacion.txtCuit.Text
        .Sections("Sección4").Controls("numero").Caption = Format(Facturacion.txtPunto.Text, "0000") & "-" & Format(nroFactura, "00000000")
        .Sections("Sección4").Controls("fecha").Caption = "Fecha: " & Facturacion.fecEmision.Value
        
        .Sections("Sección5").Controls("subtotal").Caption = Facturacion.txtSubtotal
        .Sections("Sección5").Controls("iva").Caption = Facturacion.txtIva
        .Sections("Sección5").Controls("total").Caption = Facturacion.txtTotal
        
        .Sections("Sección5").Controls("vencimiento").Caption = "Fecha de Vencimiento: " & Format(DateAdd("d", 7, Facturacion.fecEmision.Value), "dd/mm/yyyy")
        
        .Show vbModal
    End With
Else
    MsgBox "No hay datos para mostrar.", vbInformation
End If
End Sub

Private Sub ImpresionFletesRattaro()
Dim rs As New ADODB.Recordset
Dim sql As String

Call ConectarBD
sql = "select descripcion, cantidad, precio_unitario as punitario, precio_neto as neto from productos_ventas where factura = " & nroFactura
rs.Open sql, conn, adOpenStatic, adLockReadOnly

If Not rs.EOF Then
    With FletesRattaro
        Set .DataSource = rs
        .Sections("Sección4").Controls("cliente").Caption = Facturacion.txtNombre.Text
        .Sections("Sección4").Controls("direccion").Caption = Facturacion.txtDireccion.Text + " " + Facturacion.txtCP.Text + " " + Facturacion.txtLocalidad.Text + " " + Facturacion.txtProvincia.Text
        .Sections("Sección4").Controls("localidad").Caption = Facturacion.txtLocalidad.Text + " (" + Facturacion.txtCP.Text + ") - " + Facturacion.txtProvincia.Text
        .Sections("Sección4").Controls("cuit").Caption = Facturacion.txtCuit.Text
        .Sections("Sección4").Controls("factura").Caption = Format(Facturacion.txtPunto.Text, "0000") & "-" & Format(nroFactura, "00000000")
        .Sections("Sección4").Controls("fecha").Caption = "Fecha: " & Facturacion.fecEmision.Value
        
        .Sections("Sección5").Controls("subtotal").Caption = Facturacion.txtSubtotal
        .Sections("Sección5").Controls("iva").Caption = Facturacion.txtIva
        .Sections("Sección5").Controls("total").Caption = Facturacion.txtTotal
        
        .Sections("Sección5").Controls("vencimiento").Caption = "Fecha de Vencimiento: " & Format(DateAdd("d", 7, Facturacion.fecEmision.Value), "dd/mm/yyyy")
        
        .Show vbModal
    End With
Else
    MsgBox "No hay datos para mostrar.", vbInformation
End If
End Sub


Private Sub ImpresionFerreteriaSabul()
Dim rs As New ADODB.Recordset
Dim sql As String

Call ConectarBD
sql = "select descripcion, cantidad, precio_unitario as punitario, precio_neto as neto from productos_ventas where factura = " & nroFactura
rs.Open sql, conn, adOpenStatic, adLockReadOnly

If Not rs.EOF Then
    With ferrateriaSabul
        Set .DataSource = rs
        .Sections("Sección4").Controls("cliente").Caption = Facturacion.txtNombre.Text
        .Sections("Sección4").Controls("direccion").Caption = Facturacion.txtDireccion.Text + " (" + Facturacion.txtCP.Text + ") " + Facturacion.txtLocalidad.Text + " " + Facturacion.txtProvincia.Text
        .Sections("Sección4").Controls("cuit").Caption = Facturacion.txtCuit.Text
        .Sections("Sección4").Controls("factura").Caption = Format(Facturacion.txtPunto.Text, "0000") & "-" & Format(nroFactura, "00000000")
        .Sections("Sección4").Controls("fecha").Caption = "Fecha: " & Facturacion.fecEmision.Value
        
        .Sections("Sección5").Controls("subtotal").Caption = Facturacion.txtSubtotal
        .Sections("Sección5").Controls("iva").Caption = Facturacion.txtIva
        .Sections("Sección5").Controls("total").Caption = Facturacion.txtTotal
        
        .Sections("Sección5").Controls("vencimiento").Caption = "Fecha de Vencimiento: " & Format(DateAdd("d", 7, Facturacion.fecEmision.Value), "dd/mm/yyyy")
        
        .Show vbModal
    End With
Else
    MsgBox "No hay datos para mostrar.", vbInformation
End If
End Sub

Private Sub ImpresionTrole()
Dim rs As New ADODB.Recordset
Dim sql As String

Call ConectarBD
sql = "select descripcion, cantidad, precio_unitario as punitario, precio_neto as neto from productos_ventas where factura = " & nroFactura
rs.Open sql, conn, adOpenStatic, adLockReadOnly

If Not rs.EOF Then
    With Trole
        Set .DataSource = rs
        .Sections("Sección4").Controls("cliente").Caption = Facturacion.txtNombre.Text
        .Sections("Sección4").Controls("direccion").Caption = Facturacion.txtDireccion.Text
        .Sections("Sección4").Controls("localidad").Caption = "(" + Facturacion.txtCP.Text + ") " + Facturacion.txtLocalidad.Text + " " + Facturacion.txtProvincia.Text
        .Sections("Sección4").Controls("cuit").Caption = Facturacion.txtCuit.Text
        .Sections("Sección4").Controls("factura").Caption = Format(Facturacion.txtPunto.Text, "0000") & "-" & Format(nroFactura, "00000000")
        .Sections("Sección4").Controls("fecha").Caption = "Fecha: " & Facturacion.fecEmision.Value
        
        .Sections("Sección3").Controls("subtotal").Caption = Facturacion.txtSubtotal
        .Sections("Sección3").Controls("iva").Caption = Facturacion.txtIva
        .Sections("Sección3").Controls("total").Caption = Facturacion.txtTotal
        
        .Sections("Sección3").Controls("vencimiento").Caption = "Fecha de Vencimiento: " & Format(DateAdd("d", 7, Facturacion.fecEmision.Value), "dd/mm/yyyy")
        
        .Show vbModal
    End With
Else
    MsgBox "No hay datos para mostrar.", vbInformation
End If
End Sub

Private Sub ImpresionPyramar()
Dim rs As New ADODB.Recordset
Dim sql As String

Call ConectarBD
sql = "select '' as codigo, descripcion, cantidad, precio_unitario as punitario, precio_neto as neto from productos_ventas where factura = " & nroFactura
rs.Open sql, conn, adOpenStatic, adLockReadOnly

If Not rs.EOF Then
    With parymar
        Set .DataSource = rs
        .Sections("Sección4").Controls("cliente").Caption = Facturacion.txtNombre.Text
        .Sections("Sección4").Controls("direccion").Caption = Facturacion.txtDireccion.Text
        .Sections("Sección4").Controls("localidad").Caption = "(" + Facturacion.txtCP.Text + ") " + Facturacion.txtLocalidad.Text + " " + Facturacion.txtProvincia.Text
        .Sections("Sección4").Controls("cuit").Caption = Facturacion.txtCuit.Text
        .Sections("Sección4").Controls("factura").Caption = Format(Facturacion.txtPunto.Text, "0000") & "-" & Format(nroFactura, "00000000")
        .Sections("Sección4").Controls("fecha").Caption = "Fecha: " & Facturacion.fecEmision.Value
        
        .Sections("Sección5").Controls("subtotal").Caption = Facturacion.txtSubtotal
        .Sections("Sección5").Controls("iva").Caption = Facturacion.txtIva
        .Sections("Sección5").Controls("total").Caption = Facturacion.txtTotal
        
        .Sections("Sección5").Controls("vencimiento").Caption = "Fecha de Vencimiento: " & Format(DateAdd("d", 7, Facturacion.fecEmision.Value), "dd/mm/yyyy")
        
        .Show vbModal
    End With
Else
    MsgBox "No hay datos para mostrar.", vbInformation
End If
End Sub
