Attribute VB_Name = "ModuloUtilidades"
Option Explicit

Global nroFactura As Long

' Convierte un valor de texto en número limpio
Public Function LimpiarValor(ByVal texto As String) As Double
    Dim valorLimpio As String
    valorLimpio = Replace(texto, "$", "")
    valorLimpio = Replace(valorLimpio, ",", "")
    valorLimpio = Trim(valorLimpio)
    
    If IsNumeric(valorLimpio) Then
        LimpiarValor = Val(valorLimpio)
    Else
        LimpiarValor = 0
    End If
End Function

' Formatea un número como precio
Public Function FormatoPrecio(ByVal Valor As Double) As String
    FormatoPrecio = "$" & Format$(Valor, "#,##0.00")
End Function

' Muestra una alerta estándar
Public Sub MostrarAlerta(ByVal mensaje As String)
    MsgBox mensaje, vbExclamation, "Advertencia"
End Sub

Public Function FormatearNumeroFactura(ByVal numero As Long) As String
    FormatearNumeroFactura = Format(numero, "00000000")
End Function

Public Function GenerarComprobante(ByVal subTotal As Double, iva As Integer, total As Double, formulario As Object) As String
    Dim archivo As Integer
    Dim ruta As String
    Dim i As Integer
'    Dim subTotal As Double, iva As Double, total As Double
    ruta = "C:\comprobante.html"
    archivo = FreeFile()
    Open ruta For Output As #archivo
    
    Print #archivo, "<html><head><title>Factura</title></head><body>"
    Print #archivo, "<style>"
    Print #archivo, "body { font-family: Arial, sans-serif; width: 800px; margin: auto; }"
    Print #archivo, ".container { border: 1px solid #000; padding: 0; }"
    Print #archivo, ".header { display: flex; border-bottom: 1px solid #000; position: relative; }"
    Print #archivo, ".logo-section { width: 50%; padding: 10px; }"
    Print #archivo, ".logo { width: 350px; display: block; }"
    Print #archivo, ".company-info { font-size: 12px; margin: 0; }"
    Print #archivo, ".company-space { margin-bottom: 15px; }"
    Print #archivo, ".company-margin { margin: 0 50px;} "
    Print #archivo, ".company-name { font-weight: bold; }"
    Print #archivo, ".email { color: blue; }"
    Print #archivo, ".letter-box {  width: 7%; display: flex; flex-direction: column; height: 100%; /* Añadir posición relativa */ }"
    Print #archivo, ".vertical-blue-line { position: absolute; top: 48%; left: 50%; width: 1px; height: 111px; background-color: black; print-color-adjust: exact; -webkit-print-color-adjust: exact; }"
    Print #archivo, ".letter { font-size: 36px; "
    Print #archivo, "  text-align: center;"
    Print #archivo, "  height: 100%; /* Ocupa todo el alto disponible */"
    Print #archivo, "  display: flex;"
    Print #archivo, "  flex-direction: column;"
    Print #archivo, "  justify-content: center;}"
    Print #archivo, ".letter-border { "
    Print #archivo, "  border-left: 1px solid #000;"
    Print #archivo, "  border-right: 1px solid #000;"
    Print #archivo, "  border-bottom: 1px solid #000;"
    Print #archivo, "  height: 100%; /* Ocupa todo el alto disponible */"
    Print #archivo, "  display: flex;"
    Print #archivo, "  align-items: center; /* Centra el contenido verticalmente */"
    Print #archivo, "  justify-content: center; /* Centra el contenido horizontalmente */"
    Print #archivo, "  min-height: 80px; /* Establece una altura mínima */"
    Print #archivo, "  padding: 10px 0; /* Añade espacio arriba y abajo */} "
    Print #archivo, ".factura-section { width: 50%; padding: 10px; display: flex; flex-direction: column; }"
    Print #archivo, ".factura-label { font-size: 14px; text-align: left; margin-bottom: 10px; }"
    Print #archivo, ".factura-num { font-size: 23px; margin-bottom: 20px; }"
    Print #archivo, ".factura-num-dos {margin-right: 6px;}"
    Print #archivo, ".factura-num-font {font-size: 28px; font-family: 'Courier New', monospace; color: gray}"
    Print #archivo, ".factura-info { text-align: left; margin-top: 35px;}"
    Print #archivo, ".vertical-line { position: absolute; top: 70px; left: 76%; width: 1px; height: calc(100% - 70px); background-color: #000; }"
    Print #archivo, ".items-table { width: 100%; border-collapse: collapse; }"
    Print #archivo, ".items-table th, .items-table td {"
    Print #archivo, "  width: 25%;"
    Print #archivo, "  text-align: center;"
    Print #archivo, "  font-size: 12px;"
    Print #archivo, "}"
    Print #archivo, ".watermark { position: relative; height: 400px; }"
    Print #archivo, ".watermark-content { position: absolute; width: 100%; height: 100%; display: flex; justify-content: center; align-items: center; opacity: 0.1; }"
    Print #archivo, ".watermark-logos { display: flex; justify-content: center; }"
    Print #archivo, ".watermark-logo { margin: 0 20px; }"
    Print #archivo, ".totals-section { "
    Print #archivo, "  display: flex;"
    Print #archivo, "  margin-bottom: 10px;"
    Print #archivo, "}"
    Print #archivo, ".totals-section-border {"
    Print #archivo, "  position: relative;"
    Print #archivo, "  height: 10px; /* Altura del espacio que contiene el borde */"
    Print #archivo, "  margin: 0 auto;"
    Print #archivo, "  width: 94%;"
    Print #archivo, "  border-top: 1px dashed #000;"
    Print #archivo, "}"
    Print #archivo, ".totals-section-border::before {"
    Print #archivo, "  content: '';"
    Print #archivo, "  position: absolute;"
    Print #archivo, "  top: 0;"
    Print #archivo, "  left: 10%; /* Ajusta este valor para controlar dónde comienza el borde */"
    Print #archivo, "  width: 95%; /* Ajusta este valor para controlar el ancho del borde */"
    Print #archivo, "  height: 1px;"
    Print #archivo, "  background-color: transparent; /* Elimina el fondo sólido */"
    Print #archivo, "}"
    Print #archivo, ".subtotal-section { "
    Print #archivo, "  width: 40%; "
    Print #archivo, "  padding: 10px; "
    Print #archivo, "  font-size: 12px; "
    Print #archivo, "}"
    Print #archivo, ".total-section { "
    Print #archivo, "  width: 60%; "
    Print #archivo, "  padding: 10px; "
    Print #archivo, "  text-align: right; "
    Print #archivo, "  font-size: 14px; "
    Print #archivo, "}"
    Print #archivo, ".footer {padding: 0 10px; font-size: 10px; text-align: left; margin: 15px 0;}"
    Print #archivo, ".footer p { margin: 0; }"
    Print #archivo, "@media print {"
    Print #archivo, "  .no-print, .no-print * {"
    Print #archivo, "    display: none !important;"
    Print #archivo, "    visibility: hidden !important;"
    Print #archivo, "  }"
    Print #archivo, "  .vertical-blue-line {"
    Print #archivo, "    position: absolute;"
    Print #archivo, "    top: 49.5%;"
    Print #archivo, "    left: 50%;"
    Print #archivo, "    width: 1px;"
    Print #archivo, "    height: 103px;"
    Print #archivo, "    background-color: black;"
    Print #archivo, "    print-color-adjust: exact;"
    Print #archivo, "    -webkit-print-color-adjust: exact;"
    Print #archivo, "}"
    Print #archivo, "  @page {"
    Print #archivo, "    margin: 0.5cm;"
    Print #archivo, "    size: auto;"
    Print #archivo, "  }"
    Print #archivo, "  html {"
    Print #archivo, "    -webkit-print-color-adjust: exact !important;"
    Print #archivo, "    print-color-adjust: exact !important;"
    Print #archivo, "  }"
    Print #archivo, "  body > *:not(.container) {"
    Print #archivo, "    display: none !important;"
    Print #archivo, "  }"
    Print #archivo, "  .header-date, .print-timestamp, .browser-info {"
    Print #archivo, "    display: none !important;"
    Print #archivo, "  }"
    Print #archivo, "}"
    Print #archivo, "</style>"

    Print #archivo, "<div class='container'>"
    Print #archivo, "<div class='header'>"
    Print #archivo, "  <div class='logo-section'>"
    Print #archivo, "    <button onclick='window.print()' class='no-print'>Imprimir</button>"
    Print #archivo, "    <img src='C:\Users\mlopez\Desktop\mlopez\CLIENTE 2/LOGO1.jpg' class='logo'>"
    Print #archivo, "    <div class='company-margin'>"
    Print #archivo, "      <p class='company-info company-name'>NEUMATICOS ARTURO S.R.L.</p>"
    Print #archivo, "      <p class='company-info'>AV. SAN MARTIN 1695 - 1678 CASEROS</p>"
    Print #archivo, "      <p class='company-info'>TEL.: (011) 4734 - 8476</p>"
    Print #archivo, "      <p class='company-info email'>ventas@neumaticosarturo.com.ar</p>"
    Print #archivo, "      <p class='company-info'>I.V.A.: Responsable Inscripto</p>"
    Print #archivo, "    </div>"
    Print #archivo, "  </div>"
    Print #archivo, "  <div class='letter-box'>"
    Print #archivo, "    <div class='flex letter'>"
    Print #archivo, "      <div class='letter-border'>"
    Print #archivo, "          A"
    Print #archivo, "      </div>"
    Print #archivo, "      <div class='vertical-blue-line'></div>"
    Print #archivo, "    </div>"
    Print #archivo, "  </div>"
    Print #archivo, "  <div class='factura-section'>"
    Print #archivo, "    <div class='factura-label'>FACTURA</div>"
    Print #archivo, "    <div class='factura-num'><span class='factura-num-dos'>N°0001- " & "</span> <span class='factura-num-Font'>" & nroFactura & "</span></div>"
    Print #archivo, "    <div class='factura-info'>"
    Print #archivo, "      <p class='company-info company-space'>FECHA: 17/03/2025</p>"
    Print #archivo, "      <p class='company-info'>C.U.I.T.: 33-71457404-9</p>"
    Print #archivo, "      <p class='company-info'>ING. BRUTOS: 33-71457404-9</p>"
    Print #archivo, "      <p class='company-info'>INICIO DE ACTIVIDADES: 08/2014</p>"
    Print #archivo, "    </div>"
    Print #archivo, "  </div>"
    Print #archivo, "</div>"
    Print #archivo, "<div style='border-bottom: 1px solid #000; height: 60px'></div>"
    Print #archivo, "<table class='items-table'>"
    Print #archivo, "  <tr>"
    Print #archivo, "    <th style='padding: 5px; border-bottom: none; width: 25%; text-align: center; color: gray'>CANTIDAD</th>"
    Print #archivo, "    <th style='padding: 5px; border-bottom: none; width: 32%; text-align: center; color: gray'>DESCRIPCION</th>"
    Print #archivo, "    <th style='padding: 5px; border-bottom: none; width: 25%; text-align: center; color: gray'>P.UNITARIO</th>"
    Print #archivo, "    <th style='padding: 5px; border-bottom: none; width: 25%; text-align: center; color: gray'>TOTAL</th>"
    Print #archivo, "  </tr>"
    Print #archivo, "<tr><td colspan='4'>"
    Print #archivo, "<div class='watermark'>"
    Print #archivo, "  <div class='watermark-content'>"
    Print #archivo, "    <div class='watermark-logos'>"
    Print #archivo, "      <img src='C:\Users\mlopez\Desktop\mlopez\CLIENTE 2\LOGO3.jpg' class='watermark-logo' style='width: 300px'>"
    Print #archivo, "      <img src='C:\Users\mlopez\Desktop\mlopez\CLIENTE 2\LOGO2.jpg' class='watermark-logo' style='width: 300px'>"
    Print #archivo, "    </div>"
    Print #archivo, "  </div>"
    Print #archivo, "  <table style='width: 100%; border: none'>"

    ' Recorrer ListView y agregar filas a la tabla
    With formulario
    For i = 1 To .Grilla.ListItems.Count
        Dim Cantidad As Integer
        Dim Descripcion As String
        Dim PrecioUnitario As Double
        Dim totalProducto As Double

        Cantidad = .Grilla.ListItems(i).SubItems(2)
        Descripcion = .Grilla.ListItems(i).SubItems(1)
        PrecioUnitario = .Grilla.ListItems(i).SubItems(3)
        totalProducto = .Grilla.ListItems(i).SubItems(4)

        Print #archivo, "<tr><td style='width: 25%; text-align: center'>" & Cantidad & "</td>"
        Print #archivo, "<td style='width: 32%; text-align: left; text-transform: uppercase'>" & Descripcion & "</td>"
        Print #archivo, "<td style='width: 25%; text-align: center'>" & Format(PrecioUnitario, "#,##0.00") & "</td>"
        Print #archivo, "<td style='width: 25%; text-align: center'>" & Format(totalProducto, "#,##0.00") & "</td></tr>"

'        subTotal = subTotal + totalProducto
    Next i
    End With
    Print #archivo, "  </table>"
    Print #archivo, "</div>"
    Print #archivo, "</td></tr>"
    Print #archivo, "</table>"
    Print #archivo, "<div style='display: flex'>"
    Print #archivo, "  <div class='subtotal-section'>"
    Print #archivo, "    <div style='display: flex'> "
    Print #archivo, "    <div style='margin-right: 80px'>SUBTOTAL</div>"
    Print #archivo, "    <div>IVA</div>"
    Print #archivo, "  </div>"
    Print #archivo, "<div style='display: flex; margin-top: 5px;'>"
    Print #archivo, "<div style='margin-right: 80px;'>" & Format$(subTotal, "#,##0.00") & "</div>"
    Print #archivo, "<div>(21.00)</div>"
    Print #archivo, "</div>"
    Print #archivo, "  </div>"
    Print #archivo, "  <div class='total-section'>"
    Print #archivo, "    <div>CONCEPTOS NO GRABADOS</div>"
    Print #archivo, "    <br><br>"
    Print #archivo, "    <div style='display: flex; justify-content: space-between; font-size: 16px; font-weight: bold; text-align: left; padding-left: 50%;'>TOTAL $ <span>" & Format$(total, "#,##0.00") & "</span></div>"
    Print #archivo, "  </div>"
    Print #archivo, "</div>"
    Print #archivo, "<div class='totals-section-border'></div>"
    Print #archivo, "<div class='footer'>"
    Print #archivo, "  <p>C.A.I. N 2175-19625194739</p>"
    Print #archivo, "  <p>Fecha de Vencimiento: 17/04/2025</p>"
    Print #archivo, "  <p style='margin-left: 20px;'>REGISTRO PHG3523973</p>"
    Print #archivo, "</div>"
    Print #archivo, "</div>"
    Print #archivo, "</body></html>"

    Close #archivo
    
    ' Abrir el comprobante en el navegador
    Shell "cmd /c start " & ruta, vbNormalFocus
End Function
