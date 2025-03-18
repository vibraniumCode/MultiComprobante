VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Facturacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formulario"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   13590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame8 
      Height          =   615
      Left            =   120
      TabIndex        =   33
      Top             =   8520
      Width           =   13335
      Begin VB.CommandButton btnFinalizar 
         Caption         =   "&Finalizar Venta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11160
         TabIndex        =   37
         Top             =   160
         Width           =   2055
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "@vibraniumcode - mlopez developer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   3285
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13335
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   5520
         TabIndex        =   38
         Top             =   480
         Width           =   1335
      End
      Begin VB.Frame Frame7 
         Height          =   1335
         Left            =   120
         TabIndex        =   22
         Top             =   6960
         Width           =   13095
         Begin VB.CommandButton ImpFederal 
            Caption         =   "&Comprobante federal parts"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9600
            TabIndex        =   32
            Top             =   840
            Width           =   3375
         End
         Begin VB.CommandButton ImpArturo 
            Caption         =   "&Comprobante neumaticos arturo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9600
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   240
            Width           =   3375
         End
         Begin VB.TextBox txtTotal 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6000
            TabIndex        =   30
            Text            =   "$00.00"
            Top             =   600
            Width           =   3495
         End
         Begin VB.TextBox txtIva 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3240
            TabIndex        =   29
            Text            =   "$00.00"
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox txtSubtotal 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   120
            TabIndex        =   28
            Text            =   "$00.00"
            Top             =   600
            Width           =   3015
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000080&
            BorderWidth     =   3
            DrawMode        =   14  'Copy Pen
            Index           =   1
            X1              =   6000
            X2              =   8280
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00008000&
            BorderWidth     =   3
            DrawMode        =   14  'Copy Pen
            X1              =   3240
            X2              =   4920
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line Line2 
            BorderColor     =   &H8000000D&
            BorderWidth     =   3
            DrawMode        =   14  'Copy Pen
            Index           =   0
            X1              =   120
            X2              =   2160
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "TOTAL"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6000
            TabIndex        =   27
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "ALICUOTA IVA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3240
            TabIndex        =   24
            Top             =   240
            Width           =   1320
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "SUBTOTAL"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   945
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "CARGA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1815
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   13095
         Begin VB.CommandButton btnActualizarproducto 
            Caption         =   "&Actualizar producto"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   10920
            TabIndex        =   36
            Top             =   960
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.CommandButton btnIngresarproducto 
            Caption         =   "&Ingresar producto"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   10920
            TabIndex        =   26
            Top             =   1000
            Width           =   2055
         End
         Begin VB.Frame Frame6 
            Height          =   735
            Left            =   9240
            TabIndex        =   19
            Top             =   840
            Width           =   1575
            Begin VB.TextBox iva 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   600
               TabIndex        =   21
               Text            =   "21.00"
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label11 
               Caption         =   "IVA"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   25
               Top             =   360
               Width           =   375
            End
            Begin VB.Label Label8 
               Caption         =   "IVA"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -840
               TabIndex        =   20
               Top             =   1080
               Width           =   1335
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "PRECIOS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   2520
            TabIndex        =   14
            Top             =   840
            Width           =   6615
            Begin VB.TextBox precioNeto 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   375
               Left            =   4680
               Locked          =   -1  'True
               TabIndex        =   18
               Text            =   "$00.00"
               Top             =   240
               Width           =   1815
            End
            Begin VB.TextBox Preciouni 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1680
               TabIndex        =   16
               Text            =   "$00.00"
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label Label6 
               Caption         =   "PRECIO NETO"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3600
               TabIndex        =   17
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label7 
               Caption         =   "PRECIO UNITARIO"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   360
               Width           =   1455
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "CANTIDAD"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   2295
            Begin VB.CommandButton restar 
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   600
               TabIndex        =   13
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton sumar 
               Caption         =   "+"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   120
               TabIndex        =   12
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox btnCantidad 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   11
               Text            =   "1"
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.TextBox txtDescripcion 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1320
            TabIndex        =   9
            Top             =   240
            Width           =   11655
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPCIÓN:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Registros"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   4095
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   13095
         Begin MSComctlLib.ListView Grilla 
            Height          =   3735
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   12855
            _ExtentX        =   22675
            _ExtentY        =   6588
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.TextBox factura 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label fecha 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   11400
         TabIndex        =   5
         Top             =   600
         Width           =   75
      End
      Begin VB.Label Label4 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de factura:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1695
      End
      Begin VB.Line Line1 
         X1              =   6600
         X2              =   13200
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label2 
         Caption         =   "FACTURA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Menu mnuListView 
      Caption         =   "&mnuListView"
      Visible         =   0   'False
      Begin VB.Menu Eliminar 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu Actualizar 
         Caption         =   "Actualizar"
      End
   End
End
Attribute VB_Name = "Facturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Producto As New ClaseProducto
Dim alertaMostrada As Boolean
Dim idProducto As Long
Dim nroFactura As Long

Private Sub Actualizar_Click()
    Dim test As String
    
    test = Grilla.SelectedItem.Text
    
    ' Verificar si hay elementos en la grilla
    If Grilla.ListItems.Count = 0 Then
        MsgBox "No hay productos para actualizar", vbExclamation
        Exit Sub
    End If
    
    ' Verificar si hay un elemento seleccionado
    On Error Resume Next
    
    If Err.Number <> 0 Then
        MsgBox "Por favor, seleccione un producto para actualizar", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Una vez confirmado que hay un elemento seleccionado, cargar sus datos
    Call CargarDatosParaActualizar
    
    ' Cambiar visibilidad de los botones (si es necesario)
    btnIngresarproducto.Visible = False
    btnActualizarproducto.Visible = True
End Sub

Private Sub CargarDatosParaActualizar()
    ' Obtener el ID del producto seleccionado
    idProducto = CLng(Grilla.SelectedItem.Text)
    
    ' Guardar el ID en el Tag del formulario
    Me.Tag = CStr(idProducto)
    
    ' Cargar datos en los TextBox
    txtDescripcion.Text = Grilla.SelectedItem.SubItems(1)
    btnCantidad.Text = Grilla.SelectedItem.SubItems(2)
    Preciouni.Text = Grilla.SelectedItem.SubItems(3)
    precioNeto.Text = Grilla.SelectedItem.SubItems(4)
    ' Agregar más campos según sea necesario
End Sub

Private Sub btnActualizarproducto_Click()
' Validar los datos
    If Trim(txtDescripcion.Text) = "" Then
        MsgBox "La descripción no puede estar vacía", vbExclamation
        txtDescripcion.SetFocus
        Exit Sub
    End If
    
    ' Conectar a la base de datos
    Call ConectarBD
    
    If txtDescripcion.Text = "" Then
        MostrarAlerta "Ingrese una descripción del producto."
        Exit Sub
    ElseIf btnCantidad = 0 Then
        MostrarAlerta "La cantidad no puede ser cero. Ingrese un valor válido."
        Exit Sub
    ElseIf Preciouni.Text = 0 Then
        MostrarAlerta "El precio unitario no puede ser cero. Ingrese un valor válido."
        Exit Sub
    End If
    
    ' Actualizar el registro en la base de datos
    On Error GoTo ErrHandler
    conn.Execute "UPDATE PRODUCTOS_VENTAS SET " & _
                "DESCRIPCION = '" & Replace(txtDescripcion.Text, "'", "''") & "', " & _
                "CANTIDAD = " & Replace(btnCantidad.Text, ",", ".") & ", " & _
                "PRECIO_UNITARIO = " & Producto.PrecioUnitario & ", " & _
                "PRECIO_UNITARIO = " & Producto.PrecioUnitario & ", " & _
                "PRECIO_NETO = " & Producto.precioNeto & _
                " WHERE ID = " & idProducto
    
    ' Desconectar de la base de datos
    Call DesconectarBD
    
    ' Actualizar la grilla
    Call CargarGrilla
    Call CalculoGral
    ' Limpiar los campos y restablecer botones
    LimpiarCampos
    btnIngresarproducto.Visible = True
    btnActualizarproducto.Visible = False
    
    MsgBox "Producto actualizado correctamente", vbInformation
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error al actualizar el producto: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
End Sub

Private Sub Eliminar_Click()
    Dim idProducto As Long
    
    ' Verificar si hay un elemento seleccionado
    If Grilla.SelectedItem Is Nothing Then
        MsgBox "Por favor, seleccione un elemento para eliminar", vbExclamation
        Exit Sub
    End If
    
    ' Obtener el ID del producto desde el ListView
    idProducto = CLng(Grilla.SelectedItem.Text)
    
    ' Conectar a la base de datos
    Call ConectarBD
    
    ' Eliminar el producto de la base de datos usando Execute
    On Error GoTo ErrHandler
    conn.Execute "DELETE FROM PRODUCTOS_VENTAS WHERE id = " & idProducto
    
    ' Desconectar de la base de datos
    Call DesconectarBD
    
    ' Eliminar el item seleccionado del ListView
    Grilla.ListItems.Remove Grilla.SelectedItem.Index
    
    MsgBox "Registro eliminado correctamente", vbInformation
    Call CalculoGral
    Exit Sub
    
ErrHandler:
    MsgBox "Error al eliminar el producto: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
End Sub

Private Sub Form_Load()
    Dim Cantidad As Integer
    
    fecha = "Fecha: " & Date
    CargarNumeroFactura

    Producto.Cantidad = 0
    Producto.PrecioUnitario = 0
    Producto.precioNeto = 0
    Producto.Descripcion = ""
    
    ' Conectar a la base de datos
    Call ConectarBD

    ' Configuración del ListView
    With Grilla
        .View = lvwReport
        .ColumnHeaders.Add , , "Id", 1000
        .ColumnHeaders.Add , , "Descripción", 2000
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Precio Unitario", 1500
        .ColumnHeaders.Add , , "Precio Neto", 1500
    End With

    ' Cargar datos en el ListView
    Call CargarGrilla
    

    Cantidad = Grilla.ListItems.Count
    If Cantidad > 0 Then

        Call CalculoGral
    End If
End Sub

Private Sub btnIngresarproducto_Click()
Dim cmd As New ADODB.Command
Dim facturaNumero As Double

facturaNumero = Val(factura.Text)
If Not IsNumeric(Producto.Cantidad) Or Not IsNumeric(Producto.PrecioUnitario) Then
    MsgBox "La cantidad y el precio deben ser números válidos.", vbCritical, "Error"
    Exit Sub
End If

' Conectar a la base de datos utilizando el módulo de conexión
    Call ConectarBD

    If Producto.Descripcion = "" Then
        MostrarAlerta "Ingrese una descripción del producto."
        Exit Sub
    ElseIf Producto.Cantidad = 0 Then
        MostrarAlerta "La cantidad no puede ser cero. Ingrese un valor válido."
        Exit Sub
    ElseIf Producto.PrecioUnitario = 0 Then
        MostrarAlerta "El precio unitario no puede ser cero. Ingrese un valor válido."
        Exit Sub
    End If

 'Preparar comando SQL para insertar datos
    On Error GoTo ErrHandler
    With cmd
        .ActiveConnection = conn
        .CommandType = adCmdText
        .CommandText = "INSERT INTO PRODUCTOS_VENTAS (DESCRIPCION, CANTIDAD, PRECIO_UNITARIO, PRECIO_NETO, FACTURA) VALUES (?, ?, ?, ?, ?)"
        .Parameters.Append .CreateParameter("DESCRIPCION", adVarChar, adParamInput, 255, Producto.Descripcion)
        .Parameters.Append .CreateParameter("CANTIDAD", adInteger, adParamInput, , Producto.Cantidad)
        .Parameters.Append .CreateParameter("PRECIO_UNITARIO", adDouble, adParamInput, , Producto.PrecioUnitario)
        .Parameters.Append .CreateParameter("PRECIO_NETO", adDouble, adParamInput, , Producto.precioNeto)
        .Parameters.Append .CreateParameter("FACTURA", adDouble, adParamInput, , facturaNumero)
        .Execute
    End With

    ' Actualizar el ListView después de la inserción
    Call CargarGrilla
    Call CalculoGral
    LimpiarCampos
    
    ' Desconectar
    Call DesconectarBD
    Exit Sub

ErrHandler:
    MsgBox "Error al insertar: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
End Sub

Private Sub Grilla_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 ' Mostrar el menú contextual solo si se hace clic derecho
    If Button = vbRightButton Then
        ' Mostrar el menú emergente
        PopupMenu mnuListView
    End If
End Sub

Private Sub ImpArturo_Click()
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ManejadorErrores
    
    ' Conectar a la base de datos utilizando el módulo de conexión
    Call ConectarBD
    
    ' Verificar que la conexión esté abierta
    If conn.State = adStateClosed Then
        MsgBox "Error: La conexión a la base de datos no se estableció correctamente.", vbCritical, "Error de conexión"
        GoTo Finalizar
    End If
    
    ' Obtener datos de la base
    rs.Open "SELECT cantidad, descripcion, precio_unitario, precio_neto AS total FROM PRODUCTOS_VENTAS", conn, adOpenStatic, adLockReadOnly
    
    ' Cargar datos en el reporte
    If rs.State = adStateOpen And Not rs.EOF Then
        With RptArturo
            Set .DataSource = rs
            .Refresh    ' Añadido para asegurar que el reporte se actualiza
            .Show
        End With
    Else
        MsgBox "No se pudieron cargar los datos. Verifique la consulta o la conexión.", vbExclamation, "Aviso"
    End If
    
    GoTo Finalizar
    
ManejadorErrores:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error en reporte"
    
Finalizar:
    ' Cerrar el recordset si está abierto
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    
    ' Desconectar
    Call DesconectarBD
    Exit Sub
End Sub



Private Sub ImpFederal_Click()
 Call GenerarComprobante
End Sub

Private Sub GenerarComprobante()
    Dim archivo As Integer
    Dim ruta As String
    Dim i As Integer
    Dim subtotal As Double, iva As Double, total As Double
    ruta = "C:\comprobante.html"
    archivo = FreeFile()
    Open ruta For Output As #archivo
    
    ' Encabezado del HTML con diseño exacto al de la imagen
'    Print #archivo, "<html><head><title>Factura</title></head><body>"
'    Print #archivo, "<style>"
'    Print #archivo, "body { font-family: Arial, sans-serif; width: 800px; margin: auto; }"
'    Print #archivo, ".container { border: 1px solid #000; padding: 0; }"
'    Print #archivo, ".header { display: flex; border-bottom: 1px solid #000; position: relative; }"
'    Print #archivo, ".logo-section { width: 70%; border-right: 1px solid #000; padding: 10px; }"
'    Print #archivo, ".logo { width: 250px; display: block; margin-bottom: 10px; }"
'    Print #archivo, ".company-info { font-size: 12px; margin: 3px 0; }"
'    Print #archivo, ".company-name { font-weight: bold; }"
'    Print #archivo, ".email { color: blue; }"
'    Print #archivo, ".letter-box { position: absolute; top: 0; left: 70%; width: 6%; height: 70px; border-right: 1px solid #000; display: flex; align-items: center; justify-content: center; }"
'    Print #archivo, ".letter { font-size: 36px; font-weight: bold; }"
'    Print #archivo, ".factura-section { width: 30%; padding: 10px; display: flex; flex-direction: column; }"
'    Print #archivo, ".factura-label { font-size: 14px; text-align: left; margin-bottom: 5px; }"
'    Print #archivo, ".factura-num { font-size: 16px; font-weight: bold; margin-bottom: 20px; }"
'    Print #archivo, ".factura-info { text-align: left; }"
'    Print #archivo, ".vertical-line { position: absolute; top: 70px; left: 76%; width: 1px; height: calc(100% - 70px); background-color: #000; }"
'    Print #archivo, ".items-table { width: 100%; border-collapse: collapse; }"
'    Print #archivo, ".items-table th { background-color: #f5f5f5; font-size: 12px; border-bottom: 1px solid #000; padding: 8px 5px; text-align: center; }"
'    Print #archivo, ".items-table td { padding: 5px; text-align: center; font-size: 12px; }"
'    Print #archivo, ".watermark { position: relative; height: 400px; }"
'    Print #archivo, ".watermark-content { position: absolute; width: 100%; height: 100%; display: flex; justify-content: center; align-items: center; opacity: 0.1; }"
'    Print #archivo, ".watermark-logos { display: flex; justify-content: center; }"
'    Print #archivo, ".watermark-logo { margin: 0 20px; }"
'    Print #archivo, ".totals-section { border-top: 1px solid #000; display: flex; }"
'    Print #archivo, ".subtotal-section { width: 40%; padding: 10px; border-right: 1px solid #000; font-size: 12px; }"
'    Print #archivo, ".total-section { width: 60%; padding: 10px; text-align: right; font-size: 14px; }"
'    Print #archivo, ".footer { border-top: 1px solid #000; padding: 10px; font-size: 10px; text-align: center; }"
'    Print #archivo, "</style>"
    
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
Print #archivo, "  padding: 5px;"
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
    
    
'    Print #archivo, "<div class='container'>"
'
'    ' Header con disposición exacta según la imagen
'    Print #archivo, "<div class='header'>"
'    Print #archivo, "  <div class='logo-section'>"
'    Print #archivo, "    <img src='C:\Users\progr\Downloads\CLIENTE 2\LOGO1.jpg' class='logo'>"
'    Print #archivo, "    <p class='company-info company-name'>NEUMÁTICOS ARTURO S.R.L.</p>"
'    Print #archivo, "    <p class='company-info'>AV. SAN MARTÍN 1695 - 1678 CASEROS</p>"
'    Print #archivo, "    <p class='company-info'>TEL.: (011) 4734 - 8476</p>"
'    Print #archivo, "    <p class='company-info email'>ventas@neumaticosarturo.com.ar</p>"
'    Print #archivo, "    <p class='company-info'>I.V.A.: Responsable Inscripto</p>"
'    Print #archivo, "  </div>"
'    Print #archivo, "  <div class='letter-box'>"
'    Print #archivo, "    <div class='letter'>A</div>"
'    Print #archivo, "  </div>"
'    Print #archivo, "  <div class='vertical-line'></div>"
'    Print #archivo, "  <div class='factura-section'>"
'    Print #archivo, "    <div class='factura-label'>FACTURA</div>"
'    Print #archivo, "    <div class='factura-num'>N°0001- " & factura.Text & "</div>"
'    Print #archivo, "    <div class='factura-info'>"
'    Print #archivo, "      <p class='company-info'>FECHA: " & Format(Date, "dd/mm/yyyy") & "</p>"
'    Print #archivo, "      <p class='company-info'>C.U.I.T.: 33-71457404-9</p>"
'    Print #archivo, "      <p class='company-info'>ING. BRUTOS: 33-71457404-9</p>"
'    Print #archivo, "      <p class='company-info'>INICIO DE ACTIVIDADES: 08/2014</p>"
'    Print #archivo, "    </div>"
'    Print #archivo, "  </div>"
'    Print #archivo, "</div>"
'
'    ' Table for items
'    Print #archivo, "<table class='items-table'>"
'    Print #archivo, "<tr><th>CANTIDAD</th><th>DESCRIPCION</th><th>P.UNITARIO</th><th>TOTAL</th></tr>"
'
'    ' Watermark section with items
'    Print #archivo, "<tr><td colspan='4'>"
'    Print #archivo, "<div class='watermark'>"
'    Print #archivo, "  <div class='watermark-content'>"
'    Print #archivo, "    <div class='watermark-logos'>"
'    Print #archivo, "      <img src='C:\Users\progr\Downloads\CLIENTE 2\LOGO1.jpg' class='watermark-logo' style='width: 200px;'>"
'    Print #archivo, "      <img src='C:\Users\progr\Downloads\CLIENTE 2\LOGO2.jpg' class='watermark-logo' style='width: 200px;'>"
'    Print #archivo, "      <img src='C:\Users\progr\Downloads\CLIENTE 2\LOGO3.jpg' class='watermark-logo' style='width: 200px;'>"
'    Print #archivo, "    </div>"
'    Print #archivo, "  </div>"
'
'    ' Content table (inside the watermark area)
'    Print #archivo, "  <table style='width: 100%; border: none;'>"
'
'    ' Recorrer ListView y agregar filas a la tabla
'    For i = 1 To Grilla.ListItems.Count
'        Dim Cantidad As Integer
'        Dim Descripcion As String
'        Dim PrecioUnitario As Double
'        Dim totalProducto As Double
'
'        Cantidad = Val(Grilla.ListItems(i).SubItems(1))
'        Descripcion = Grilla.ListItems(i).Text
'        PrecioUnitario = Val(Grilla.ListItems(i).SubItems(2))
'        totalProducto = Cantidad * PrecioUnitario
'
'        Print #archivo, "<tr><td style='width: 15%; text-align: center;'>" & Cantidad & "</td>"
'        Print #archivo, "<td style='width: 45%; text-align: left;'>" & Descripcion & "</td>"
'        Print #archivo, "<td style='width: 20%; text-align: right;'>$" & Format(PrecioUnitario, "#,##0.00") & "</td>"
'        Print #archivo, "<td style='width: 20%; text-align: right;'>$" & Format(totalProducto, "#,##0.00") & "</td></tr>"
'
'        subtotal = subtotal + totalProducto
'    Next i
'
'    Print #archivo, "  </table>"
'    Print #archivo, "</div>"
'    Print #archivo, "</td></tr>"
'    Print #archivo, "</table>"
'
'    ' Calcular IVA y total
'    iva = subtotal * 0.21  ' Suponiendo 21% de IVA
'    total = subtotal + iva
'
'    ' Totals section
'    Print #archivo, "<div class='totals-section'>"
'    Print #archivo, "  <div class='subtotal-section'>"
'    Print #archivo, "    <div>SUBTOTAL</div>"
'    Print #archivo, "    <div>IVA</div>"
'    Print #archivo, "  </div>"
'    Print #archivo, "  <div class='total-section'>"
'    Print #archivo, "    <div>CONCEPTOS NO GRABADOS</div>"
'    Print #archivo, "    <br><br>"
'    Print #archivo, "    <div style='font-size: 16px; font-weight: bold;'>TOTAL $" & Format(total, "#,##0.00") & "</div>"
'    Print #archivo, "  </div>"
'    Print #archivo, "</div>"
'
'    ' Footer
'    Print #archivo, "<div class='footer'>"
'    Print #archivo, "  <p>C.A.I. N 2175-19625194739</p>"
'    Print #archivo, "  <p>Fecha de Vencimiento: " & Format(DateAdd("m", 1, Date), "dd/mm/yyyy") & "</p>"
'    Print #archivo, "  <p>REGISTRO PHG3523973</p>"
'    Print #archivo, "</div>"
'
'    Print #archivo, "</div>" ' End container
'
'    ' Cerrar HTML
'    Print #archivo, "</body></html>"


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
Print #archivo, "    <div class='factura-num'><span class='factura-num-dos'>N°0001- " & "</span> <span class='factura-num-Font'>" & factura.Text & "</span></div>"
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
Print #archivo, "    <th style='border-bottom: none; width: 25%; text-align: center; color: gray'>CANTIDAD</th>"
Print #archivo, "    <th style='border-bottom: none; width: 25%; text-align: center; color: gray'>DESCRIPCION</th>"
Print #archivo, "    <th style='border-bottom: none; width: 25%; text-align: center; color: gray'>P.UNITARIO</th>"
Print #archivo, "    <th style='border-bottom: none; width: 25%; text-align: center; color: gray'>TOTAL</th>"
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
'Print #archivo, "    <tr>"
'Print #archivo, "      <td style='width: 25%; text-align: center'>21</td>"
'Print #archivo, "      <td style='width: 25%; text-align: center'>NEUMATICO 2.0 STR</td>"
'Print #archivo, "      <td style='width: 25%; text-align: center'>$3,00</td>"
'Print #archivo, "      <td style='width: 25%; text-align: center'>$0,00</td>"
'Print #archivo, "    </tr>"
'Print #archivo, "    <tr>"
'Print #archivo, "      <td style='width: 25%; text-align: center'>21</td>"
'Print #archivo, "      <td style='width: 25%; text-align: center'>NEUMATICO 2.0 STR</td>"
'Print #archivo, "      <td style='width: 25%; text-align: center'>$2,00</td>"
'Print #archivo, "      <td style='width: 25%; text-align: center'>$0,00</td>"
'Print #archivo, "    </tr>"

    ' Recorrer ListView y agregar filas a la tabla
    For i = 1 To Grilla.ListItems.Count
        Dim Cantidad As Integer
        Dim Descripcion As String
        Dim PrecioUnitario As Double
        Dim totalProducto As Double

        Cantidad = Grilla.ListItems(i).SubItems(2)
        Descripcion = Grilla.ListItems(i).SubItems(1)
        PrecioUnitario = Grilla.ListItems(i).SubItems(3)
        totalProducto = Grilla.ListItems(i).SubItems(4)

        Print #archivo, "<tr><td style='width: 25%; text-align: center'>" & Cantidad & "</td>"
        Print #archivo, "<td style='width: 25%; text-align: left; text-transform: uppercase'>" & Descripcion & "</td>"
        Print #archivo, "<td style='width: 25%; text-align: center'>" & Format(PrecioUnitario, "#,##0.00") & "</td>"
        Print #archivo, "<td style='width: 25%; text-align: center'>" & Format(totalProducto, "#,##0.00") & "</td></tr>"

        subtotal = subtotal + totalProducto
    Next i

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
Print #archivo, "  </div>"
Print #archivo, "  <div class='total-section'>"
Print #archivo, "    <div>CONCEPTOS NO GRABADOS</div>"
Print #archivo, "    <br><br>"
Print #archivo, "    <div style='font-size: 16px; font-weight: bold; text-align: left; padding-left: 50%;'>TOTAL $</div>"
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
End Sub

Private Sub precioNeto_LostFocus()
    Producto.precioNeto = LimpiarValor(precioNeto.Text)
    precioNeto.Text = FormatoPrecio(Producto.precioNeto)
End Sub

Private Sub Preciouni_KeyPress(KeyAscii As Integer)
    ' Permitir solo números, el signo de dólar, el punto decimal, la retroceso y la barra espaciadora
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 32) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Preciouni_LostFocus()
    Producto.PrecioUnitario = LimpiarValor(Preciouni.Text)
    Preciouni.Text = FormatoPrecio(Producto.PrecioUnitario)
    Producto.Cantidad = Val(btnCantidad.Text)

    Call ActualizarPrecio
End Sub

Private Sub btnCantidad_Change()
    alertaMostrada = False
    Producto.Cantidad = Val(btnCantidad.Text)
    Call ActualizarPrecio
End Sub

Private Sub CargarPrecio()
    precioNeto.Text = FormatoPrecio(Producto.CalcularPrecioNeto())
End Sub

Private Sub sumar_Click()
    alertaMostrada = False
    Producto.Cantidad = Producto.Cantidad + 1
    btnCantidad.Text = Producto.Cantidad
    Call ActualizarPrecio
End Sub

Private Sub restar_Click()
    Producto.Cantidad = Producto.Cantidad - 1
    btnCantidad.Text = Producto.Cantidad
    Call ActualizarPrecio
End Sub

Private Sub ActualizarPrecio()
    precioNeto.Text = FormatoPrecio(Producto.CalcularPrecioNeto())
    precioNeto_LostFocus
End Sub

Private Sub txtDescripcion_LostFocus()
    Producto.Descripcion = txtDescripcion.Text
End Sub

Private Sub CargarGrilla()
    Dim rs As New ADODB.Recordset

    ' Conectar a la base de datos utilizando el módulo de conexión
    Call ConectarBD

    ' Limpiar el ListView antes de agregar los nuevos datos
    Grilla.ListItems.Clear

    ' Obtener datos de la base
    On Error GoTo ErrHandler

    rs.Open "SELECT id, descripcion, cantidad, precio_unitario, precio_neto FROM PRODUCTOS_VENTAS WHERE FACTURA = " & nroFactura, conn, adOpenStatic, adLockReadOnly

    ' Cargar datos en el ListView
    If Not rs.EOF Then
        Do While Not rs.EOF
            With Grilla.ListItems.Add(, , rs("id"))
                .SubItems(1) = rs("descripcion")
                .SubItems(2) = rs("cantidad")
                .SubItems(3) = "$" & Format(rs("precio_unitario"), "#,##0.00")
                .SubItems(4) = "$" & Format(rs("precio_neto"), "#,##0.00")
            End With
            rs.MoveNext
        Loop
    Else
        MsgBox "No hay productos registrados.", vbExclamation, "Aviso"
    End If

    ' Cerrar el recordset
    rs.Close
    
    ' Desconectar
    Call DesconectarBD
    Exit Sub

ErrHandler:
    MsgBox "Error al cargar datos: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
End Sub
Private Sub CalculoGral()
    Dim rs As New ADODB.Recordset

    ' Conectar a la base de datos utilizando el módulo de conexión
    Call ConectarBD

    ' Obtener datos de la base
    On Error GoTo ErrHandler
    rs.Open "SELECT " & _
                "SUM(PRECIO_NETO) SUBTOTAL, " & _
                "SUM(PRECIO_NETO * 0.21) AS IVA, " & _
                "SUM(PRECIO_NETO) + SUM(PRECIO_NETO * 0.21) AS TOTAL " & _
                "FROM PRODUCTOS_VENTAS Where factura = " & nroFactura, conn, adOpenStatic, adLockReadOnly
    
    txtSubtotal.Text = Format(rs(0), "$0.00")
    txtIva.Text = Format(rs(1), "$0.00")
    txtTotal.Text = Format(rs(2), "$0.00")
    
    ' Cerrar el recordset
    rs.Close
    
    ' Desconectar
    Call DesconectarBD
    Exit Sub

ErrHandler:
    MsgBox "Error al cargar datos: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
End Sub

' Procedimiento para limpiar los campos
Private Sub LimpiarCampos()
    txtDescripcion.Text = ""
    btnCantidad.Text = 1
    Preciouni.Text = "$" & Format(0, "#,##0.00")
    precioNeto.Text = "$" & Format(0, "#,##0.00")
    Me.Tag = ""  ' Limpiar el ID guardado
End Sub

Private Sub CargarNumeroFactura()
    Dim rs As New ADODB.Recordset

    ' Conectar a la base de datos
    Call ConectarBD

    ' Obtener el último número de factura
    On Error GoTo ErrHandler
    rs.Open "SELECT MAX(factura) AS UltimoNro FROM FACTURAS", conn, adOpenStatic, adLockReadOnly

    ' Verificar si hay datos
    If Not rs.EOF Then
        nroFactura = rs("UltimoNro")
        factura.Text = FormatearNumeroFactura(nroFactura)
    Else
        MsgBox "No hay facturas registradas.", vbExclamation, "Aviso"
    End If

    ' Cerrar el recordset y desconectar
    rs.Close
    Call DesconectarBD
    Exit Sub

ErrHandler:
    MsgBox "Error al obtener el número de factura: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
End Sub

