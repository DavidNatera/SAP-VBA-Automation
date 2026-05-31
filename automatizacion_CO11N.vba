Option Explicit         ' Se asegura una declaración explícita de todas las variables.
                        ' Mejora la detección de errores y el mantenimiento del código.

Sub AccesoSAP_Optimizado()

    ' === Declaración de Variables del Tipo Objeto de Excel ===
    Dim EJECUTABLE As Workbook          ' Libro actual donde se ejecuta el código
    Dim NOTIFICACIONES As Worksheet     ' Hoja donde se encuentra la tabla para las notificaciones
    Dim Tbl_DATOS As ListObject         ' Tabla donde se encuentran los datos a notificar
    
    ' === Declaración de Variables de la MATRIZ ===
    Dim arrDatos As Variant             ' Matriz para cargar los datos en memoria
    Dim i As Long                       ' Contador para el ciclo de la matriz
    Dim cORDEN As Long                  ' Indicador de la columna ORDEN
    Dim cOP As Long                     ' Indicador de la columna OP
    Dim cCLASE As Long                  ' Indicador de la columna CLASE
    Dim cCANT As Long                   ' Indicador de la columna CANT A NOTIFICAR
    Dim cCOLAB As Long                  ' Indicador de la columna COLABORADOR
    Dim cFINI As Long                   ' Indicador de la columna FECHA INICIAL
    Dim cFFIN As Long                   ' Indicador de la columna FECHA FIN
    Dim cHTRAB As Long                  ' Indicador de la columna HORA TRAB
    Dim cHMAQ As Long                   ' Indicador de la columna HORA MAQ
    Dim cHPARADA As Long                ' Indicador de la columna HORA PARADA
    Dim cESTATUS As Long                ' Indicador de la columna ESTATUS
    Dim cMENSAJE As Long                ' Indicador de la columna MENSAJE
    
    ' === Declaración de Variables del Tipo Objeto del SAP ===
    Dim SapGuiAuto As Object            ' Objeto para scripting SAP GUI
    Dim App As Object                   ' Referencia al motor de scripting de SAP
    Dim Connection As Object            ' Conexión activa al servidor SAP.
    Dim Session As Object               ' Sesión activa de SAP
    
    ' === Variables para Mensajes de SAP y Lógica ===
    Dim tipoMsj As String
    Dim textoMsj As String
    Dim horasCent As Double             ' Variable para la hora centesimal
    Dim horaFinConvertida As String     ' Variable para la hora ya en formato reloj
    
    ' === Asignación de Objetos y Validaciones ===
    Set EJECUTABLE = ThisWorkbook
    Set NOTIFICACIONES = EJECUTABLE.Worksheets("NOTIFICACIONES")
    Set Tbl_DATOS = NOTIFICACIONES.ListObjects("Tbl_DATOS")
    
    ' === Verificación que la Tbl_DATOS contenga datos ===
    ' 1. Verifica si la tabla esta vacia (0 filas)
    If Tbl_DATOS.DataBodyRange Is Nothing Then
        MsgBox "La tabla no tiene filas de datos. Ingrese información antes de continuar.", vbInformation, "Sin Datos"
        Exit Sub
    End If
    
    ' 2. Verifica si la tabla tiene filas, pero TODAS sus celdas están vacías (CountA = 0)
    If Application.WorksheetFunction.CountA(Tbl_DATOS.DataBodyRange) = 0 Then
        MsgBox "Las filas de la tabla están completamente vacías. Ingrese datos válidos.", vbInformation, "Sin Datos"
        Exit Sub
    End If

    ' === Optimización de Entorno Excel ===
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.EnableCancelKey = xlErrorHandler
    
    ' === Manejador de errores general, por si la Macro, Excel o SAP colapsan ===
    On Error GoTo ManejadorErrores

    ' === Manejador de errores para SAP Abierto y Conectado ===
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    On Error GoTo ManejadorErrores
    
    ' === Verificación que el SAP GUI esté abierto ===
    If SapGuiAuto Is Nothing Then
        RestaurarEntorno
        MsgBox "SAP no está abierto. Inicie SAP y abra una sesión.", vbCritical
        Exit Sub
    End If
    
    Set App = SapGuiAuto.GetScriptingEngine
    
    ' === Verificación que hay conexión de SAP activa ===
    If App.Connections.Count = 0 Then
        RestaurarEntorno
        MsgBox "No hay ninguna conexión activa en SAP.", vbCritical
        Exit Sub
    End If

    Set Connection = App.Connections(0)
    
    ' === Verificación que hay menos de 6 sesiones de SAP activa ===
    If Connection.Sessions.Count < 6 Then
        Connection.Sessions(0).CreateSession
        Application.Wait (Now() + TimeValue("0:00:03"))
        Set Session = Connection.Sessions(Connection.Sessions.Count - 1)
    Else
        RestaurarEntorno
        MsgBox "Se alcanzó el máximo de sesiones en SAP (6).", vbExclamation
        Exit Sub
    End If
      
    ' === Carga de la Tbl_DATOS en una Matriz ===
    ' 1. Identificamos en qué posición está cada columna
    With Tbl_DATOS.ListColumns
        cORDEN = .Item("ORDEN").Index
        cOP = .Item("OP").Index
        cCLASE = .Item("CLASE NOTI").Index
        cCANT = .Item("CANT NOTI").Index
        cCOLAB = .Item("COLABORADOR").Index
        cFINI = .Item("FECHA INI").Index
        cFFIN = .Item("FECHA FIN").Index
        cHTRAB = .Item("HORA TRAB").Index
        cHMAQ = .Item("HORA MAQ").Index
        cHPARADA = .Item("HORA PARADA").Index
        cESTATUS = .Item("ESTATUS").Index
        cMENSAJE = .Item("MENSAJE SAP").Index
    End With
    
    ' 2. Cargamos toda la tabla a la RAM
    arrDatos = Tbl_DATOS.DataBodyRange.Value
    
    ' === Ciclo de Notificación ===
    For i = LBound(arrDatos, 1) To UBound(arrDatos, 1)
        
        Session.StartTransaction "CO11N"
        
        With Session.findById("wnd[0]")
            .FindByName("AFRUD-AUFNR", "GuiCTextField").Text = CStr(arrDatos(i, cORDEN))
            .FindByName("AFRUD-VORNR", "GuiCTextField").Text = CStr(arrDatos(i, cOP))
            .FindByName("AFRUD-AUERU", "GuiComboBox").Value = Trim(CStr(arrDatos(i, cCLASE)))
                      
            ' Lógica de compensación
            If UCase(Trim(CStr(arrDatos(i, cCLASE)))) <> "NOTIFICACIÓN PARCIAL" Then
                .FindByName("AFRUD-AUSOR", "GuiCheckBox").Selected = True
            Else
                .FindByName("AFRUD-AUSOR", "GuiCheckBox").Selected = False
            End If
            
            .FindByName("AFRUD-LMNGA", "GuiTextField").Text = CStr(arrDatos(i, cCANT))
            
            ' Boton (Datos Reales)
            .FindByName("btn[13]", "GuiButton").press
            
            ' Evaluamos la barra de estado tras pisar el botón 13
            tipoMsj = .findById("sbar").MessageType
            textoMsj = .findById("sbar").Text
            
            ' Si es Error (E), Aborto (A) o Advertencia (W - como la tolerancia de tu imagen)
            If tipoMsj = "E" Or tipoMsj = "A" Or tipoMsj = "W" Then
                arrDatos(i, cESTATUS) = "ERROR/ADVERTENCIA"
                arrDatos(i, cMENSAJE) = textoMsj
                GoTo SiguienteFila ' Aborta esta fila y salta a la etiqueta de abajo
            End If
            
            ' -------------------------------------------------------------
            ' 1. Horas Trabajadas y Máquina a 3 decimales (Corrección de coma decimal)
            Dim valTRAB As Double
            Dim valMAQ As Double
            
            ' Validamos que haya un número real. Si lo hay, usamos CDbl que respeta la coma (,)
            ' Si está vacío, le asignamos 0
            If IsNumeric(arrDatos(i, cHTRAB)) And arrDatos(i, cHTRAB) <> "" Then
                valTRAB = CDbl(arrDatos(i, cHTRAB))
            Else
                valTRAB = 0
            End If
            
            If IsNumeric(arrDatos(i, cHMAQ)) And arrDatos(i, cHMAQ) <> "" Then
                valMAQ = CDbl(arrDatos(i, cHMAQ))
            Else
                valMAQ = 0
            End If
            
            .FindByName("AFRUD-ISM03", "GuiTextField").Text = Format(valTRAB, "0.000")
            .FindByName("AFRUD-ISM04", "GuiTextField").Text = Format(valMAQ, "0.000")
            .FindByName("AFRUD-PERNR", "GuiCTextField").Text = CStr(arrDatos(i, cCOLAB))
            .FindByName("AFRUD-ISDD", "GuiCTextField").Text = CStr(arrDatos(i, cFINI))
            .FindByName("AFRUD-IEDD", "GuiCTextField").Text = CStr(arrDatos(i, cFFIN))
            .FindByName("AFRUD-ISDZ", "GuiCTextField").Text = "00:00:00"
            
            ' Conversión de Hora Parada (Centesimal a Reloj)
            If IsNumeric(arrDatos(i, cHPARADA)) And arrDatos(i, cHPARADA) <> "" Then
                horasCent = CDbl(arrDatos(i, cHPARADA))
                horaFinConvertida = Format(horasCent / 24, "hh:mm:ss")
            Else
                horaFinConvertida = "00:00:00"
            End If
            .FindByName("AFRUD-IEDZ", "GuiCTextField").Text = horaFinConvertida
            
            ' 2. Presionar botón Guardar
            .FindByName("btn[11]", "GuiButton").press
            
            ' --- CAPTURA FINAL DESPUÉS DE GUARDAR ---
            tipoMsj = .findById("sbar").MessageType
            textoMsj = .findById("sbar").Text
            
            ' Evaluamos "Falsos Positivos" y Advertencias
            If tipoMsj = "S" Then
                ' Buscamos la palabra "ERROR" dentro del texto SAP
                ' (Convertimos todo a mayúscula con UCase para evitar problemas con minúscula/mayúscula)
                If InStr(1, UCase(textoMsj), "ERROR") > 0 Then
                    arrDatos(i, cESTATUS) = "ERROR (Mov. Mercancías)"
                Else
                    arrDatos(i, cESTATUS) = "OK"
                End If
                
            ElseIf tipoMsj = "W" Or tipoMsj = "I" Then
                ' SÍ es amarillo (W) o azul (I), SAP se detuvo y no guardó
                arrDatos(i, cESTATUS) = "ADVERTENCIA (No guardó)"
            Else
                ' SÍ es un mensaje E o A genuino
                arrDatos(i, cESTATUS) = "ERROR"
            End If
            
            arrDatos(i, cMENSAJE) = textoMsj
        End With
        
SiguienteFila:
    Next i
    
    ' === Volcado de la matriz actualizada a Excel en un solo paso ===
    Tbl_DATOS.DataBodyRange.Value = arrDatos
    
    ' === Cierre y Limpieza ===
    Connection.CloseSession (Session.Name)
    RestaurarEntorno
    
    ' Mensaje final con reporte de resultados
    MsgBox "Proceso Concluido. Revise las columnas de ESTATUS y MENSAJES SAP.", vbInformation
    Exit Sub

' === Subrutina para reactivar Excel SÍ algo falla de manera general y captura del error ===
ManejadorErrores:
    
    RestaurarEntorno
    
    ' Evaluamos si el error fue provocado por el usuario presionando ESC (Error 18)
    If Err.Number = 18 Then
        MsgBox "El usuario canceló el proceso manualmente.", vbExclamation, "Proceso Interrumpido"
    Else
        MsgBox "Se produjo un error inesperado en la macro." & vbCrLf & _
               "Número de Error: " & Err.Number & vbCrLf & _
               "Detalle: " & Err.Description, vbCritical, "Error del Sistema"
    End If
End Sub
' === Subrutina para reactivar Excel al terminar el proceso Y/O si falla por alguna razón de manera general ===
Sub RestaurarEntorno()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
