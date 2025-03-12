VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Facturacion 
   Caption         =   "Formulario"
   ClientHeight    =   9135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   ScaleHeight     =   9135
   ScaleWidth      =   13590
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame8 
      Height          =   615
      Left            =   120
      TabIndex        =   34
      Top             =   8520
      Width           =   13335
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
         TabIndex        =   35
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
      Begin VB.Frame Frame7 
         Height          =   1335
         Left            =   120
         TabIndex        =   23
         Top             =   6960
         Width           =   13095
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   615
            Left            =   6360
            TabIndex        =   36
            Top             =   480
            Width           =   975
         End
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
            TabIndex        =   33
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
            TabIndex        =   32
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
            Left            =   4080
            TabIndex        =   31
            Text            =   "$00.00"
            Top             =   600
            Width           =   1455
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
            Left            =   1800
            TabIndex        =   30
            Text            =   "$00.00"
            Top             =   600
            Width           =   1695
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
            TabIndex        =   29
            Text            =   "$00.00"
            Top             =   600
            Width           =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000080&
            BorderWidth     =   3
            DrawMode        =   14  'Copy Pen
            Index           =   1
            X1              =   4080
            X2              =   5520
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00008000&
            BorderWidth     =   3
            DrawMode        =   14  'Copy Pen
            X1              =   1800
            X2              =   3480
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line Line2 
            BorderColor     =   &H8000000D&
            BorderWidth     =   3
            DrawMode        =   14  'Copy Pen
            Index           =   0
            X1              =   120
            X2              =   1320
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
            Left            =   4080
            TabIndex        =   28
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
            Left            =   1800
            TabIndex        =   25
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
            TabIndex        =   24
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
         TabIndex        =   8
         Top             =   1080
         Width           =   13095
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
            TabIndex        =   27
            Top             =   1000
            Width           =   2055
         End
         Begin VB.Frame Frame6 
            Height          =   735
            Left            =   9240
            TabIndex        =   20
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
               TabIndex        =   22
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
               TabIndex        =   26
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
               TabIndex        =   21
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
            TabIndex        =   15
            Top             =   840
            Width           =   6615
            Begin VB.TextBox precioNeto 
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
               ForeColor       =   &H00004000&
               Height          =   375
               Left            =   4680
               TabIndex        =   19
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
               TabIndex        =   17
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
               TabIndex        =   18
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
               TabIndex        =   16
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
            TabIndex        =   11
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
               TabIndex        =   14
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
               TabIndex        =   13
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
               TabIndex        =   12
               Text            =   "0"
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
            TabIndex        =   10
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
            TabIndex        =   9
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
            Height          =   3615
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   12855
            _ExtentX        =   22675
            _ExtentY        =   6376
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.TextBox Text1 
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
         Text            =   "00000000000000000000000000000000001"
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label5 
         Caption         =   "FECHA: 08/03/2025"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11400
         TabIndex        =   5
         Top             =   600
         Width           =   1815
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
End
Attribute VB_Name = "Facturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Producto As New ClaseProducto
Dim alertaMostrada As Boolean

Private Sub Form_Load()
    Producto.Cantidad = 0
    Producto.PrecioUnitario = 0
    Producto.Descripcion = ""
End Sub


Private Sub Command1_Click()
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset

    ' Cadena de conexión
    Dim strConn As String
    strConn = "Provider=SQLOLEDB;Data Source=facturacion_neumaticos.mssql.somee.com;" & _
              "Initial Catalog=facturacion_neumaticos;User ID=mlopez_cliente_2_SQLLogin_2;" & _
              "Password=khuwpknwob;Persist Security Info=False;" & _
              "TrustServerCertificate=True"

    ' Intentar la conexión
    On Error GoTo ErrHandler
    conn.Open strConn
    
    ' Consulta SQL
    Dim sql As String
    sql = "select id, nombre from usuarios" ' Ajusta según tu tabla

    ' Ejecutar consulta
    rs.Open sql, conn, adOpenStatic, adLockReadOnly

    ' Recorrer los resultados
    If Not rs.EOF Then
        Do While Not rs.EOF
            MsgBox "ID: " & rs("id") & " - Nombre: " & rs("nombre"), vbInformation, "Registro"
            rs.MoveNext
        Loop
    Else
        MsgBox "No hay registros.", vbExclamation, "Consulta"
    End If

    ' Cerrar conexión
    rs.Close
    conn.Close
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub btnIngresarproducto_Click()

    If Producto.Descripcion = "" Then
        MostrarAlerta "Ingrese una descripción del producto."
    ElseIf Producto.Cantidad = 0 Then
        MostrarAlerta "La cantidad no puede ser cero. Ingrese un valor válido."
    ElseIf Producto.PrecioUnitario = 0 Then
        MostrarAlerta "El precio unitario no puede ser cero. Ingrese un valor válido."
    Else
        MsgBox "Paso"
    End If
    
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
End Sub

Private Sub txtDescripcion_LostFocus()
    Producto.Descripcion = txtDescripcion.Text
End Sub
