'// software elaborado con objetos COM(component object model) de referencias externa utilizando la libreria de autodesk 
'//de clases para la ejecucion de comandos desde este mismo software(ejecucion por medio de una interfaz paralela a la de autocada) 
''// *version de(17.0) libreria de tipos acax17enu.tlb de los archivos Autodesk  shared!!
''//dejando claro que este software podra utilizar cada proceso rutina funcion o calses (tipos) si se cuenta con esta version de libreria..
Public Class Form3
    '// variables globales de la version (2.0)
    Dim cargas(100), cargasil(100), n_carga(100), n_cargasil(100), c_cargat(100), c_cargail(100)
    Dim n_carga220(100) As String : Dim c_carga220(100) As Integer : Dim carga220(100) As Integer
    Dim n_carga220spe(100) As String : Dim c_carga220spe(100) As Integer : Dim carga220spe(100) As Integer
    Dim vecestoma, vecesiluminacion, veces220, veces220spe As Integer ' me indica la cantidad de circuito y veces q me calculara carga por circuito
    Dim i, j 'variables utilizadas en ciclos for
    Dim deptoma(100), depiuminacion(100), deptoma220(100), deptoma220spe(100) As String
    Dim acumuladort, acumuladori, acumulador220, acumulador220spe
    Dim seleccion As Autodesk.AutoCAD.Interop.AcadSelectionSet
    Dim vectorA(100), vectorB(100) As Integer : Dim n_vectorA(100), n_vectorB(100) As String : Dim c_vectorA(100), c_vectorB(100) As Integer
    Dim v_definitivo(100), v_deficode(100), v_definombre(100), cantidad_circuito_automatico
    Dim l As Integer = 1 : Dim ll As Integer = 1  ' indica cantidad total (iluminacion + toma)
    Dim valorres(100), c_valorres(100) As Integer : Dim n_valorres(100) As String ' para guardar variables
    Dim on_off As Integer = 0 ' switch del munu "automatic"
    '////////////////////////////////////////////// _
    'variables globales: de la version (1.0) _
    '*********************************************************************************
    Dim encender As Double
    Dim exe As Microsoft.Office.Interop.Excel.Application
    Dim uu
    Dim objautocad, objCAD As Autodesk.AutoCAD.Interop.AcadApplication ' inicia toda lasaplicaciones de autocad utilizadas (referencia q aplique a mi proyecto)..
    Dim BP As Integer
    Dim KV As Double
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        '*********************************************************************************
        'DECLARACION DE VARIABLES PRINCIPALES DEL CALCULO DE TABLERO Y DIBUJADO EN AUTOCAD*
        '********************************************************************************************************************************************************************************************

        Dim line As Autodesk.AutoCAD.Interop.Common.AcadLine 'declaro la linea
        Dim PtoIn(2) As Double 'declaro punto inicio x,y,z
        Dim PtoFin(2) As Double 'declaro puntofinal x,y,z

        Dim text As Autodesk.AutoCAD.Interop.Common.AcadText   'Declare text object
        Dim insPoint(2) As Double 'Declare insertion point
        Dim textHeight As Double       'Declare text height
        Dim textStr As String         'Declare text string

        Dim arco As Autodesk.AutoCAD.Interop.Common.AcadArc ' declaro el arco
        Dim centro(2) As Double 'declaro centro en x,y,z
        Dim radio, anginic, angfinal As Double ' decalro radio , angulo inicial y final  del arco

        Dim circulo As Autodesk.AutoCAD.Interop.Common.AcadCircle
        Dim centroo(2) As Double 'declaro centro en x,y,z
        Dim radioo As Double 'declaro radio del circulo

        Dim t, color, g, kk, i, lon, tt1, tt2, ttt1, ttt2 As Double
        Dim l1, l2 As Integer
        Dim ncircuitos, ncircuitos2 As Double 'numero de circuitos

        Dim nombre 'nombre del tablero's 

        '*********************************************************************************************+++************************************************************
        'arbrimos hoja de calculo en excel si se ha esojido !! en checkbox3!!
        '*****************************************************************
        If CheckBox3.Checked = True Then
            'direccion del archivo..
            uu = "C:\estudio de carga .xlsx" ' se coloca la direccion del documento excel q se utiliza...
            'se abre excel
            exe = New Microsoft.Office.Interop.Excel.Application
            ' se abre un espacio de trabajo
            exe.Workbooks.Open(uu)
            'este es visible
            exe.Visible = True
            exe.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMinimized
        End If

        '***********************************************************************************************************************************************************
        'INICIO DEL PROCESO
        '********************************************************************************************************************
        Dim si

        'parte del programa q ve si existe esppacio para dibujar si no existe da la opcion de crear uno ahora mismo
        On Error Resume Next

        objautocad = GetObject(, "AutoCAD.Application")

        If Err.Number <> 0 Then
            Err.Clear()
            si = MsgBox("no existe espacio para dibujar desea crear uno ahora mismo?", MsgBoxStyle.Information + MsgBoxStyle.YesNo)

            If si = vbYes Then
                objautocad = CreateObject("autoCAD.application", "")
                'instruccion q se utiliza para minimizar el autocad  cuando inicie
                objautocad.WindowState = Autodesk.AutoCAD.Interop.Common.AcWindowState.acMin
                'objautocad.Visible = False hace q aparesca el autocada pero en modo de  q no s evea pero igual en el se trabaja
                'para q el close se active del menu de herrmientas...
                encender = 1
                If encender = 1 Then
                    AbrirToolStripMenuItem.Enabled = True
                End If

            Else
                Exit Sub ' funcion q actua como el break..

            End If
        End If

        '*******************************************************************************************************************
        'PARTE DEL CODIGO EN DONDE SE PREGUNTA: FACTOR DE POTENCIA(FP), NUMERO DE CIRCUITOS(ncircuitos) Y NOMBRE DEL TABLERO*
        '***********************************************************************************************************************************************************
        If ((RadioButton1.Checked Or RadioButton2.Checked) = False) Then
            t = MsgBox(" no ha seleccinado el tipo de tablero", vbCritical, "tableros")

        Else

            Dim fp As Double   'variable q contiene factor de potencia 

            On Error Resume Next
oo:
            fp = InputBox("agrege factor de potencia(FP)", "FACTOR DE POTENCIA")

            If Err.Number <> 0 Then 'corrector de errores
                Err.Clear()
                MsgBox("error ha introducido un  valor caracter", MsgBoxStyle.Critical)
                GoTo oo
            End If

            On Error Resume Next
o2:
            ncircuitos = InputBox("indique numero de circuitos", "tablero")


            If Err.Number <> 0 Then
                Err.Clear()
                MsgBox("error ha introducido un valor caracter", MsgBoxStyle.Critical)
                GoTo o2
            End If

            ncircuitos2 = ncircuitos

            lon = 87.5 * ncircuitos ' me indica la longuitud vertical de las barras monofasico y trifasico
            'donde 87.5 es el valor promedio para crearlas ( valor aproximado )
            ncircuitos = ncircuitos / 2 ' indica elnumero de barras neutro utilizadas

            nombre = InputBox("introduzca nombre de tablero a realizar")
            'excel escribe en campo el titulo del tablero..
            exe.Cells(2, 1) = UCase(nombre)
            ' color = InputBox("ingrese color")
            Dim ubicacion As String
            ubicacion = InputBox("ingrese la ubicacion de tablero")
            exe.Cells(2, 8) = UCase(ubicacion)


            If CheckBox1.Checked = True Then
                nombre = UCase(nombre) 'instruccion importante(vuelve mayuscula los string (cadenas de caractere)

            Else
            End If

            insPoint(0) = -300     'Set insertion point x coordinate
            insPoint(1) = (lon + 200)   'Set insertion point y coordinate
            insPoint(2) = 0      'Set insertion point z coordinate
            textHeight = 20                'Set text height to 1.0
            textStr = nombre       'Set the text string
            'Create Text object
            text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
            text.color = 2
            '**********************************************************************************************************************************************************

            '**************************************************************************************************************************************
            'PARTE DEL PROCESO EN EL CUAL SEGUN EL BOTON ELEGIDO (MONOFASICO O TRIFASICO) ESTE DIBUJA DE CUANTAS BARRAS VERTICALES SERA EL TABLERO*
            '**************************************************************************************************************************************



            If RadioButton1.Checked = True Then

                exe.Cells(3, 4) = 3
                PtoIn(0) = 375 : PtoIn(1) = lon : PtoIn(2) = 0
                PtoFin(0) = 375 : PtoFin(1) = 0 : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.Lineweight = 0.2
                line.color = 150

                PtoIn(0) = 450 : PtoIn(1) = lon : PtoIn(2) = 0
                PtoFin(0) = 450 : PtoFin(1) = 0 : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.Lineweight = 0.2
                line.color = 150
                'espesor de linea
                'line.Lineweight = Autodesk.AutoCAD.Interop.Common.ACAD_LWEIGHT.acLnWt020

                PtoIn(0) = 525 : PtoIn(1) = lon : PtoIn(2) = 0
                PtoFin(0) = 525 : PtoFin(1) = 0 : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.Lineweight = 0.2
                line.color = 150

            ElseIf RadioButton2.Checked = True Then
                exe.Cells(3, 4) = 2
                PtoIn(0) = 375 : PtoIn(1) = lon : PtoIn(2) = 0
                PtoFin(0) = 375 : PtoFin(1) = 0 : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.Lineweight = 0.2
                line.color = 150

                PtoIn(0) = 525 : PtoIn(1) = lon : PtoIn(2) = 0
                PtoFin(0) = 525 : PtoFin(1) = 0 : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.Lineweight = 0.2
                line.color = 150

            End If
            '*******************************************************************************************************************************************************************

            '*****************************************************************************************************************************************************************************
            'EN ESTA PARTE DEL PROCESO SE CREA BARRAS DE NEUTRO(ncircuitos=ncircuitos/2), BARRITAS PEQUEÑAS DE AMBOS LASDOS DEL TABLERO, Y SE PREGUNTA DE COLOR SERAN ESTAS AL DIBUJARCE"*
            '*****************************************************************************************************************************************************************************
            'color = InputBox("de que color desea las barras") variables para asignar un color a dibujo(conjelada temporarmente)

            Dim hg ' contador para incrementar los numeros de circuitos es deacir enumera cada fase
            hg = 0

            g = lon
            For i = 1 To Val(ncircuitos)

                g = g - 150

                PtoIn(0) = 80.4375 : PtoIn(1) = g : PtoIn(2) = 0
                PtoFin(0) = 0 : PtoFin(1) = g : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.color = 150

                'escribir en autocad"
                hg = hg + 1
                insPoint(0) = -40    'Set insertion point x coordinate
                insPoint(1) = g     'Set insertion point y coordinate
                insPoint(2) = 0      'Set insertion point z coordinate
                textHeight = 20                'Set text height to 1.0
                textStr = hg   'Set the text string
                'Create Text object
                text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                text.color = 2 'es color amarillo

                PtoIn(0) = 751.8125 : PtoIn(1) = g : PtoIn(2) = 0
                PtoFin(0) = 148.1875 : PtoFin(1) = g : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.color = 150


                PtoIn(0) = 900 : PtoIn(1) = g : PtoIn(2) = 0
                PtoFin(0) = 819.5625 : PtoFin(1) = g : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.color = 150

                'escribir en autocad"
                hg = hg + 1
                insPoint(0) = 930    'Set insertion point x coordinate
                insPoint(1) = g     'Set insertion point y coordinate
                insPoint(2) = 0      'Set insertion point z coordinate
                textHeight = 20                'Set text height to 1.0
                textStr = hg   'Set the text string
                'Create Text object
                text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                text.color = 2 'es color amarillo
            Next
            '*****************************************************************************************************************************************************************************************


            '**********************************************************************************'NOTA: PARA LA ACTUALIZACION...
            'EN ESTA PARTE DEL PROCESO SE DIBUJA LOS NODOS(CIRCULOS) PARA MONOFASICO SOLAMENTE*
            '**********************************************************************************
            'Dim relleno As Autodesk.AutoCAD.Interop.Common.AcadHatch
            ' Dim lugar, lugar2 As Double
            ' lugar = lon
            ' lugar2 = lon
            ' For i = 1 To (Val(ncircuitos) / 2)

            'primera bolita
            'If i = 1 Then
            ' lugar = lugar - 150

            ' centroo(0) = 375.0 : centroo(1) = lugar : centroo(2) = 0
            ' radioo = 17.125
            ' circulo = objautocad.ActiveDocument.ModelSpace.AddCircle(centroo, radioo)
            'relleno = objautocad.ActiveDocument.ModelSpace.AddHatch(, radioo, circulo, 1, circulo)
            '  Else
            '  lugar = lugar - 300
            '  centroo(0) = 375.0 : centroo(1) = lugar : centroo(2) = 0
            '  radioo = 17.125
            '   circulo = objautocad.ActiveDocument.ModelSpace.AddCircle(centroo, radioo)

            ' End If
            'segunda bolita 
            ' lugar2 = lugar2 - 300
            ' centroo(0) = 525.0 : centroo(1) = lugar2 : centroo(2) = 0
            ' radioo = 17.125
            '   circulo = objautocad.ActiveDocument.ModelSpace.AddCircle(centroo, radioo)

            '  Next
            '****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
            'PARTE DEL PROGRAMA  Q PREGUNTA POR CADA CIRCUITO OPCION A ELEGIR(MONOFASICO(220),TRIFASICO,ILUMUNACION,RESERVA,TOMACORRIENTE,FIN), EN ESTA PARTTE TABN EL TABLERO TIENE LA INTELIGENCIA DE QUE UNA PARTE ESTE LLENA PREVIAMENTE POR UN (MONOFASICO OTRIFASICO) ESTE SE DA CUENTA Y PASA APREGUNTAR AL QUE NO OCUPA DIBUJO, Y INDICA EN Q POCICION VAEL USUARIO,TABN SE CALCULA VAMP Y ES COLOCADO EN EL TABLERO*
            '****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
            '
            If RadioButton3.Checked = True Then
                exe.Cells(3, 2) = "208V"
            Else
                exe.Cells(3, 2) = "240V"
            End If

            If RadioButton5.Checked = True Then
                exe.Cells(2, 4) = "NLAB"
            Else
                exe.Cells(2, 4) = "NAB"
            End If

            tt1 = lon ' `primera colunna
            ttt1 = lon

            tt2 = lon ' segunda colunna
            ttt2 = lon

            Dim tt11, tt12
            tt11 = lon 'variable utilizada en arcos
            tt12 = lon

            Dim ii, iii 'variable utilizada en kk= 5 y 4

            Dim yy, yyy 'variables utilizadas en trifasico algoritmo
            yyy = 0
            yy = 0

            Dim increment, incrementt, incrementtt ' para saber cuanto va en cada circuito,para saber cual circuito estoy cubriendo en monofasico y el ultimo para saber cual es cicuitos estoy cubriendo en truifasico
            increment = 0

            Dim o, o2, o3, o4 As Integer 'apagadores y encendedores en trifasico y monofasico..

            Dim suma_iluminacion, suma_tomacorriente, suma_monofasico_trifasico As Double 'variables utilizadas para acumular cargas va

            suma_iluminacion = 0
            suma_tomacorriente = 0
            suma_monofasico_trifasico = 0

            Dim factor_demanda_tomacorriente, factor_demanda_iluminacion, factor_demanda_fuerza As Double 'variable q contienen FACTOR DE DEAMANDA de (TOMACORRIENTES,ALUMBRAADOS,FUERZA(MONOFASICO_TRIFASICO))
            Dim demanda_tomacorriente, demanda_ilumiinacion, demanda_fuerza, demanda_total As Double 'variables q contienen la demanda de tomacorriente,iluminacion y fuerza(monofasico-trifasico)y la suma de todas esta q es la demanda total
            Dim I_total As Double ' variable q contiene la corriente total del tablero 

            Dim KTT, KTT2 As Double ' VARIABLES Q ME INDICAN FASES EN COMPUTO SI ES TRIIFASICO(A,B,C) O MONOFASICO(A,B)
            KTT = 1
            KTT2 = 1
            'variable utilizadas para las cell de excel
            Dim gg As Integer
            gg = 11

            For i = 1 To Val(ncircuitos)

                Dim jab, jabb, jak, jakk, jal, jall

                Dim va, amp 'VARIABLE Q GUARDA CARGA voltio amperio

                If i = 1 Then

                    jab = 184.3424
                    jabb = 206.25
                    'para el de kk= 2
                    jak = 334.34
                    jakk = 355.95
                    'para el de kk =3
                    jal = 484.34
                    jall = 505.64

                Else
                    jab = 150
                    jabb = 149.695
                    'para el de kk= 2
                    jak = 300
                    jakk = 300
                    'para el de kk =3
                    jal = 450
                    jall = 450
                End If

                If o = 1 Then ' explicar a papa el proceso de como trabaja este algoritmo
                    o = 0
                    increment = increment + 1
                    ' incrementt = increment + 2
                    'incrementtt = increment + 3
                    GoTo nm

                End If

                If o3 = 3 Then 'algoritmo para movilizar un trifasico
                    yy = yy + 1
                    increment = increment + 1
                    'incrementt = increment + 2
                    'incrementtt = increment + 3
                    If yy = 2 Then
                        o3 = 0
                        yy = 0
                    End If
                    GoTo nm
                End If 'fin del algoritmo

                'tk= comando para viajar sin trasar linea
                increment = increment + 1
                incrementt = increment + 2
                incrementtt = increment + 4
                gg = gg + 1

                On Error Resume Next
o22:
                kk = InputBox("Otra carga de 110V (4), Iluminacion (0), Dos fases(2), Tres fases(3),Tomacorriente(1),Reserva(5) ", "circuito numero " & increment)
                'corrector de errores.. si se introduce un tipo de variable q no relaciona con la vaibale repite l apregunta
                If Err.Number <> 0 Then
                    Err.Clear()
                    MsgBox("error ha introducido un valor caracter", MsgBoxStyle.Critical)
                    GoTo o22
                End If

                If kk = 4 Then

                    exe.Cells(gg, 1) = increment
                    exe.Cells(gg, 3) = "X"

                    tt1 = tt1 - jab

                    PtoIn(0) = 336.41 : PtoIn(1) = tt1 : PtoIn(2) = 0
                    PtoFin(0) = 0 : PtoFin(1) = tt1 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6

                    ttt1 = ttt1 - jabb ' segunda barrita
                    PtoIn(0) = 317.25 : PtoIn(1) = ttt1 : PtoIn(2) = 0
                    PtoFin(0) = 0 : PtoFin(1) = ttt1 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6

                    'en este parte de esta opcion se calcula los (amp) en el programa
repit:              ' hacer alos demascuando los amps son myor a 100repetirde nuevo esta pregunta
                    va = InputBox("agrege carga(VA) de circuito " & increment, "carga(VA)")
                    exe.Cells(gg, 4) = va
                    suma_tomacorriente = suma_tomacorriente + va 'acumula cargas va de tomacorriente y tomacorriente personalizado
                    amp = (va) / (110) 'ecuacion q es utilizada para calcaular los amper(alumbrado,tomacorriente,(1))

                    Dim texpp
                    texpp = tt1 'distancia en el momento q esta parado( inicio ejemplo tt1= tt1 - jab) en ese momento
                    insPoint(0) = 164.25    'Set insertion point x coordinate
                    insPoint(1) = (texpp + 38) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    If (amp > 0 And amp <= 20) Then
                        textStr = ("20A-1P")
                    ElseIf (amp > 20 And amp <= 30) Then
                        textStr = "30A-1P"
                    ElseIf (amp > 30 And amp <= 50) Then
                        textStr = "50A-1P"
                    ElseIf (amp > 50 And amp <= 60) Then
                        textStr = "60A-1P"
                    ElseIf (amp > 60 And amp <= 80) Then
                        textStr = "80A-1P"
                    ElseIf (amp > 80 And amp <= 100) Then
                        textStr = "100A-1P"
                    ElseIf (amp > 100 And amp <= 150) Then
                        textStr = "150A-1P"
                    ElseIf (amp > 150 And amp <= 175) Then
                        textStr = "175A-1P"
                    ElseIf (amp > 175 And amp <= 200) Then
                        textStr = "200A-1P"
                    ElseIf (amp > 200 And amp <= 225) Then
                        textStr = "225A-1P"
                    ElseIf (amp > 225 And amp <= 250) Then
                        textStr = "250A-1P"
                    ElseIf (amp > 250 And amp <= 275) Then
                        textStr = "275A-1P"
                    ElseIf (amp > 275 And amp <= 300) Then
                        textStr = "300A-1P"
                    ElseIf (amp > 300 And amp <= 350) Then
                        textStr = "350A-1P"
                    ElseIf (amp > 350 And amp <= 400) Then
                        textStr = "400A-1P"
                    ElseIf (amp > 400 And amp <= 450) Then
                        textStr = "450A-1P"
                    Else
                        GoTo repit
                    End If
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                    text.color = 2
                    exe.Cells(gg, 6) = UCase(textStr)

                    If RadioButton1.Checked = True Then
                        If KTT = 1 Then
                            exe.Cells(gg, 7) = "A"
                        ElseIf KTT = 2 Then
                            exe.Cells(gg, 7) = "B"
                        Else
                            exe.Cells(gg, 7) = "C"
                        End If
                    End If

                    If RadioButton2.Checked = True Then
                        If KTT2 = 1 Then
                            exe.Cells(gg, 7) = "A"
                        ElseIf KTT2 = 2 Then
                            exe.Cells(gg, 7) = "B"
                        End If
                    End If

                    tt11 = tt11 - 150 ' arco de las fases NOTA: mejorar la creacion del arco..
                    centro(0) = 335 - (246.81 - 26.1) : centro(1) = tt11 : centro(2) = 0  'arco
                    radio = 33.88
                    anginic = 0 ' angulos trabajan contra las abujas del relog
                    angfinal = -9.4012
                    arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
                    arco.color = 6

                    Dim nombrem, tesxm

                    nombrem = InputBox("introduzca nombre de circuito")
                    exe.Cells(gg, 8) = UCase(nombrem)

                    If CheckBox1.Checked = True Then ' condicion q si se escoje mayuscula los caracteres salgan en mayuscula
                        nombrem = UCase(nombrem)
                    End If

                    tesxm = tt1
                    insPoint(0) = -336.41    'Set insertion point x coordinate
                    insPoint(1) = (tesxm + 35.26) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    textStr = nombrem       'Set the text string
                    'Create Text object
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                    text.color = 2

                ElseIf kk = 0 Then
                    exe.Cells(gg, 2) = "X"
                    exe.Cells(gg, 1) = increment

                    If i = 1 Then
                        l1 = 1
                    End If 'nda q ver con el codigo q se aplica en esta opcion dento del algoritmo


                    tt1 = tt1 - jab
                    ttt1 = ttt1 - jabb ' barrita que no es colocada pero q de igual forma se tiene q ir restando para el proximo que lo requiera

                    PtoIn(0) = 336.41 : PtoIn(1) = tt1 : PtoIn(2) = 0
                    PtoFin(0) = 0 : PtoFin(1) = tt1 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6
                    'en este parte de esta opcion se calcula los (amp) en el programa
repit5:
                    va = InputBox("agrege carga(VA) de circuito " & increment, "carga(VA)")
                    exe.Cells(gg, 4) = va
                    suma_iluminacion = suma_iluminacion + va 'acumulador de cargas va para iluminacion
                    amp = (va) / (110) 'ecuacion q es utilizada para calcaular los amper(alumbrado,tomacorriente,(1))

                    Dim textt
                    textt = tt1 'distancia en el momento q esta parado( inicio ejemplo tt1= tt1 - jab) en ese momento
                    insPoint(0) = 164.25      'Set insertion point x coordinate
                    insPoint(1) = (textt + 38) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    If (amp > 0 And amp <= 20) Then
                        textStr = ("20A-1P")
                    ElseIf (amp > 20 And amp <= 30) Then
                        textStr = "30A-1P"
                    ElseIf (amp > 30 And amp <= 50) Then
                        textStr = "50A-1P"
                    ElseIf (amp > 50 And amp <= 60) Then
                        textStr = "60A-1P"
                    ElseIf (amp > 60 And amp <= 80) Then
                        textStr = "80A-1P"
                    ElseIf (amp > 80 And amp <= 100) Then
                        textStr = "100A-1P"
                    ElseIf (amp > 100 And amp <= 150) Then
                        textStr = "150A-1P"
                    ElseIf (amp > 150 And amp <= 175) Then
                        textStr = "175A-1P"
                    ElseIf (amp > 175 And amp <= 200) Then
                        textStr = "200A-1P"
                    ElseIf (amp > 200 And amp <= 225) Then
                        textStr = "225A-1P"
                    ElseIf (amp > 225 And amp <= 250) Then
                        textStr = "250A-1P"
                    ElseIf (amp > 250 And amp <= 275) Then
                        textStr = "275A-1P"
                    ElseIf (amp > 275 And amp <= 300) Then
                        textStr = "300A-1P"
                    ElseIf (amp > 300 And amp <= 350) Then
                        textStr = "350A-1P"
                    ElseIf (amp > 350 And amp <= 400) Then
                        textStr = "400A-1P"
                    ElseIf (amp > 400 And amp <= 450) Then
                        textStr = "450A-1P"
                    Else
                        GoTo repit5
                    End If
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight) 'Create Text object
                    text.color = 2
                    exe.Cells(gg, 6) = UCase(textStr)

                    If RadioButton1.Checked = True Then
                        If KTT = 1 Then
                            exe.Cells(gg, 7) = "A"
                        ElseIf KTT = 2 Then
                            exe.Cells(gg, 7) = "B"
                        Else
                            exe.Cells(gg, 7) = "C"
                        End If
                    End If

                    If RadioButton2.Checked = True Then
                        If KTT2 = 1 Then
                            exe.Cells(gg, 7) = "A"
                        ElseIf KTT2 = 2 Then
                            exe.Cells(gg, 7) = "B"
                        End If
                    End If

                    tt11 = tt11 - 150 ' arco de las fases NOTA: mejorar la creacion del arco..
                    centro(0) = 335 - (246.81 - 26.1) : centro(1) = tt11 : centro(2) = 0  'arco
                    radio = 33.88
                    anginic = 0 ' angulos trabajan contra las abujas del relog
                    angfinal = -9.4012
                    arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
                    arco.color = 6

                    Dim tesxms

                    tesxms = tt1

                    insPoint(0) = -336.41    'Set insertion point x coordinate
                    insPoint(1) = (tesxms + 35.26) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0

                    If CheckBox1.Checked = True Then ' condicion q si se escoje mayuscula los caracteres salgan en mayuscula
                        textStr = ("ILUMINACION")   'Set the text string
                    Else
                        textStr = ("iluminacion")   'Set the text string
                    End If
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight) 'Create Text object
                    text.color = 2
                    exe.Cells(gg, 8) = UCase(textStr)
                    'ElseIf kk = "fin" Then exit for

                ElseIf kk = 2 Then
                    exe.Cells(gg, 3) = "X"
                    exe.Cells(gg, 1) = increment & "," & incrementt

                    tt1 = tt1 - jak
                    ttt1 = ttt1 - jakk
                    o = 1

                    If CheckBox2.Checked = False Then
                        PtoIn(0) = 317.25 : PtoIn(1) = tt1 : PtoIn(2) = 0
                        PtoFin(0) = 0 : PtoFin(1) = tt1 : PtoFin(2) = 0
                        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                        line.color = 6
                    Else
                        'se selecciona la casilla "principal" en la interfacee 
                        PtoIn(0) = 336.41 : PtoIn(1) = tt1 : PtoIn(2) = 0
                        PtoFin(0) = 0 : PtoFin(1) = tt1 : PtoFin(2) = 0
                        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                        line.color = 6

                        PtoIn(0) = 317.25 : PtoIn(1) = ttt1 : PtoIn(2) = 0
                        PtoFin(0) = 0 : PtoFin(1) = ttt1 : PtoFin(2) = 0
                        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                        line.color = 6
                    End If
                    'en este parte de esta opcion se calcula los (amp) en el programa
repit1:
                    va = InputBox("agrege carga(VA) de circuito " & increment & "," & incrementt, "carga(VA)")
                    exe.Cells(gg, 4) = va
                    suma_monofasico_trifasico = suma_monofasico_trifasico + va 'sumatoria de todas las cargas va en monofasico-trifasico

                    If RadioButton3.Checked = True Then ' SI ES SELECCIONADO 208 V
                        amp = (va) / (208) 'ecuacion q es utilizada para calcaular los amper(alumbrado,tomacorriente,(1))
                    ElseIf RadioButton4.Checked = True Then ' SI ES SELECCIONADO 204
                        amp = (va) / (240)
                    End If

                    Dim texk
                    texk = tt1 'distancia en el momento q esta parado( inicio ejemplo tt1= tt1 - jab) en ese momento
                    insPoint(0) = 164.25    'Set insertion point x coordinate
                    insPoint(1) = (texk + 188) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    If (amp > 0 And amp <= 20) Then
                        textStr = ("20A-2P")
                    ElseIf (amp > 20 And amp <= 30) Then
                        textStr = "30A-2P"
                    ElseIf (amp > 30 And amp <= 50) Then
                        textStr = "50A-2P"
                    ElseIf (amp > 50 And amp <= 60) Then
                        textStr = "60A-2P"
                    ElseIf (amp > 60 And amp <= 80) Then
                        textStr = "80A-2P"
                    ElseIf (amp > 80 And amp <= 100) Then
                        textStr = "100A-2P"
                    ElseIf (amp > 100 And amp <= 150) Then
                        textStr = "150A-2P"
                    ElseIf (amp > 150 And amp <= 175) Then
                        textStr = "175A-2P"
                    ElseIf (amp > 175 And amp <= 200) Then
                        textStr = "200A-2P"
                    ElseIf (amp > 200 And amp <= 225) Then
                        textStr = "225A-2P"
                    ElseIf (amp > 225 And amp <= 250) Then
                        textStr = "250A-2P"
                    ElseIf (amp > 250 And amp <= 275) Then
                        textStr = "275A-2P"
                    ElseIf (amp > 275 And amp <= 300) Then
                        textStr = "300A-2P"
                    ElseIf (amp > 300 And amp <= 350) Then
                        textStr = "350A-2P"
                    ElseIf (amp > 350 And amp <= 400) Then
                        textStr = "400A-2P"
                    ElseIf (amp > 400 And amp <= 450) Then
                        textStr = "450A-2P"

                    Else
                        GoTo repit1
                    End If
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight) 'Create Text object
                    text.color = 2
                    exe.Cells(gg, 6) = UCase(textStr)

                    If RadioButton1.Checked = True Then
                        If KTT = 1 Then
                            exe.Cells(gg, 7) = "A,B"
                        ElseIf KTT = 2 Then
                            exe.Cells(gg, 7) = "B,C"
                        Else
                            exe.Cells(gg, 7) = "C,A"
                        End If
                    End If

                    If RadioButton2.Checked = True Then
                        If KTT2 = 1 Then
                            exe.Cells(gg, 7) = "A,B"
                        ElseIf KTT2 = 2 Then
                            exe.Cells(gg, 7) = "B,A"
                        End If
                    End If

                    For iii = 1 To 2
                        tt11 = tt11 - 150 ' arco de las fases NOTA: mejorar la creacion del arco..
                        centro(0) = 335 - (246.81 - 26.1) : centro(1) = tt11 : centro(2) = 0  'arco
                        radio = 33.88
                        anginic = 0 ' angulos trabajan contra las abujas del relog
                        angfinal = -9.4012
                        arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
                        arco.color = 6
                    Next iii

                    PtoIn(0) = 114.29 : PtoIn(1) = tt11 + 183.88 : PtoIn(2) = 0 '( coloca la barra vertical entre los acos monofasicos )
                    PtoFin(0) = 114.29 : PtoFin(1) = tt1 + 68.22 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6

                    Dim nombreet, tesxt

                    tesxt = tt1
                    nombreet = InputBox("introduzca nombre de circuito")
                    exe.Cells(gg, 8) = UCase(nombreet)
                    If CheckBox1.Checked = True Then ' condicion q si se escoje mayuscula los caracteres salgan en mayuscula
                        nombreet = UCase(nombreet)
                    Else
                    End If

                    insPoint(0) = -336.41    'Set insertion point x coordinate
                    insPoint(1) = (tesxt + 109.34) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    textStr = nombreet       'Set the text string
                    'Create Text object
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                    text.color = 2

                ElseIf kk = 8 Then
                    Exit For

                ElseIf kk = 3 Then
                    exe.Cells(gg, 3) = "X"
                    exe.Cells(gg, 1) = increment & "," & incrementt & "," & incrementtt
                    tt1 = tt1 - jal
                    ttt1 = ttt1 - jall
                    o3 = 3
                    If CheckBox2.Checked = False Then
                        PtoIn(0) = 317.25 : PtoIn(1) = tt1 : PtoIn(2) = 0
                        PtoFin(0) = 0 : PtoFin(1) = tt1 : PtoFin(2) = 0
                        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                        line.color = 6
                    Else
                        'se selecciona la casilla "principal" en la interfacee 
                        PtoIn(0) = 336.41 : PtoIn(1) = tt1 : PtoIn(2) = 0
                        PtoFin(0) = 0 : PtoFin(1) = tt1 : PtoFin(2) = 0
                        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                        line.color = 6
                        PtoIn(0) = 317.25 : PtoIn(1) = ttt1 : PtoIn(2) = 0
                        PtoFin(0) = 0 : PtoFin(1) = ttt1 : PtoFin(2) = 0
                        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                        line.color = 6
                    End If
                    'en este parte de esta opcion se calcula los (amp) en el programa
repit4:

                    va = InputBox("agrege carga(VA) de circuito " & increment & "," & incrementt & "," & incrementtt, "carga(VA)")
                    exe.Cells(gg, 4) = va
                    suma_monofasico_trifasico = suma_monofasico_trifasico + va 'sumatoria de todas las cargas va en monofasico-trifasico
                    amp = (va) / (Math.Sqrt(3) * 208) 'ecuacion q es utilizada para calcaular los amper(alumbrado,tomacorriente,(1))

                    Dim texdss
                    texdss = tt1 'distancia en el momento q esta parado( inicio ejemplo tt1= tt1 - jab) en ese momento
                    insPoint(0) = 164.25     'Set insertion point x coordinate
                    insPoint(1) = (texdss + 338) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    If (amp > 0 And amp <= 20) Then
                        textStr = ("20A-3P")
                    ElseIf (amp > 20 And amp <= 30) Then
                        textStr = "30A-3P"
                    ElseIf (amp > 30 And amp <= 50) Then
                        textStr = "50A-3P"
                    ElseIf (amp > 50 And amp <= 60) Then
                        textStr = "60A-3P"
                    ElseIf (amp > 60 And amp <= 80) Then
                        textStr = "80A-3P"
                    ElseIf (amp > 80 And amp <= 100) Then
                        textStr = "100A-3P"
                    ElseIf (amp > 100 And amp <= 150) Then
                        textStr = "150A-3P"
                    ElseIf (amp > 150 And amp <= 175) Then
                        textStr = "175A-3P"
                    ElseIf (amp > 175 And amp <= 200) Then
                        textStr = "200A-3P"
                    ElseIf (amp > 200 And amp <= 225) Then
                        textStr = "225A-3P"
                    ElseIf (amp > 225 And amp <= 250) Then
                        textStr = "250A-3P"
                    ElseIf (amp > 250 And amp <= 275) Then
                        textStr = "275A-3P"
                    ElseIf (amp > 275 And amp <= 300) Then
                        textStr = "300A-3P"
                    ElseIf (amp > 300 And amp <= 350) Then
                        textStr = "350A-3P"
                    ElseIf (amp > 350 And amp <= 400) Then
                        textStr = "400A-3P"
                    ElseIf (amp > 400 And amp <= 450) Then
                        textStr = "450A-3P"
                    Else
                        GoTo repit4
                    End If
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight) 'Create Text object
                    text.color = 2
                    exe.Cells(gg, 6) = UCase(textStr)

                    If RadioButton1.Checked = True Then
                        If KTT = 1 Then
                            exe.Cells(gg, 7) = "A,B,C"
                        ElseIf KTT = 2 Then
                            exe.Cells(gg, 7) = "B,C,A"
                        Else
                            exe.Cells(gg, 7) = "C,A,B"
                        End If
                    End If

                    For ii = 1 To 3
                        tt11 = tt11 - 150 ' arco de las fases NOTA: mejorar la creacion del arco..
                        centro(0) = 335 - (246.81 - 26.1) : centro(1) = tt11 : centro(2) = 0  'arco
                        radio = 33.88
                        anginic = 0 ' angulos trabajan contra las abujas del relog
                        angfinal = -9.4012
                        arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
                        arco.color = 6
                    Next ii

                    PtoIn(0) = 114.29 : PtoIn(1) = tt11 + 333.88 : PtoIn(2) = 0  '( coloca las lineas entre los arcos trifasicos)
                    PtoFin(0) = 114.29 : PtoFin(1) = tt1 + 68.22 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6

                    Dim nombree, tesx
                    tesx = tt1
                    nombree = InputBox("introduzca nombre de circuito")
                    exe.Cells(gg, 8) = UCase(nombree)
                    If CheckBox1.Checked = True Then ' condicion q si se escoje mayuscula los caracteres salgan en mayuscula
                        nombree = UCase(nombree)
                    Else
                    End If

                    insPoint(0) = -336.41    'Set insertion point x coordinate
                    insPoint(1) = (tesx + 184.34) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    textStr = nombree       'Set the text string
                    'Create Text object
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                    text.color = 2

                ElseIf kk = 1 Then
                    exe.Cells(gg, 3) = "X"
                    exe.Cells(gg, 1) = increment

                    tt1 = tt1 - jab

                    PtoIn(0) = 336.41 : PtoIn(1) = tt1 : PtoIn(2) = 0
                    PtoFin(0) = 0 : PtoFin(1) = tt1 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6

                    ttt1 = ttt1 - jabb ' segunda barrita
                    PtoIn(0) = 317.25 : PtoIn(1) = ttt1 : PtoIn(2) = 0
                    PtoFin(0) = 0 : PtoFin(1) = ttt1 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6
                    'en esta parte de esta opcion se calcula los (amp) en el programa
repit9:
                    va = InputBox("agrege carga(VA) de circuito " & increment, "carga(VA)")
                    exe.Cells(gg, 4) = va
                    suma_tomacorriente = suma_tomacorriente + va 'acumula cargas va de tomacorriente y tomacorriente personalizado                    amp = (va * 100) / (110 * fp) 'ecuacion q es utilizada para calcaular los amper(alumbrado,tomacorriente,(1))
                    amp = (va) / (110) 'ecuacion q es utilizada para calcaular los amper(alumbrado,tomacorriente,(1))

                    Dim textll

                    textll = tt1 'distancia en el momento q esta parado( inicio ejemplo tt1= tt1 - jab) en ese momento
                    insPoint(0) = 164.25    'Set insertion point x coordinate
                    insPoint(1) = (textll + 38) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    If (amp > 0 And amp <= 20) Then
                        textStr = ("20A-1P")
                    ElseIf (amp > 20 And amp <= 30) Then
                        textStr = "30A-1P"
                    ElseIf (amp > 30 And amp <= 50) Then
                        textStr = "50A-1P"
                    ElseIf (amp > 50 And amp <= 60) Then
                        textStr = "60A-1P"
                    ElseIf (amp > 60 And amp <= 80) Then
                        textStr = "80A-1P"
                    ElseIf (amp > 80 And amp <= 100) Then
                        textStr = "100A-1P"
                    ElseIf (amp > 100 And amp <= 150) Then
                        textStr = "150A-1P"
                    ElseIf (amp > 150 And amp <= 175) Then
                        textStr = "175A-1P"
                    ElseIf (amp > 175 And amp <= 200) Then
                        textStr = "200A-1P"
                    ElseIf (amp > 200 And amp <= 225) Then
                        textStr = "225A-1P"
                    ElseIf (amp > 225 And amp <= 250) Then
                        textStr = "250A-1P"
                    ElseIf (amp > 250 And amp <= 275) Then
                        textStr = "275A-1P"
                    ElseIf (amp > 275 And amp <= 300) Then
                        textStr = "300A-1P"
                    ElseIf (amp > 300 And amp <= 350) Then
                        textStr = "350A-1P"
                    ElseIf (amp > 350 And amp <= 400) Then
                        textStr = "400A-1P"
                    ElseIf (amp > 400 And amp <= 450) Then
                        textStr = "450A-1P"

                    Else
                        GoTo repit9
                    End If
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight) 'Create Text object
                    text.color = 2
                    exe.Cells(gg, 6) = UCase(textStr)

                    If RadioButton1.Checked = True Then
                        If KTT = 1 Then
                            exe.Cells(gg, 7) = "A"
                        ElseIf KTT = 2 Then
                            exe.Cells(gg, 7) = "B"
                        Else
                            exe.Cells(gg, 7) = "C"
                        End If
                    End If

                    If RadioButton2.Checked = True Then
                        If KTT2 = 1 Then
                            exe.Cells(gg, 7) = "A"
                        ElseIf KTT2 = 2 Then
                            exe.Cells(gg, 7) = "B"
                        End If
                    End If

                    tt11 = tt11 - 150 ' arco de las fases NOTA: mejorar la creacion del arco..
                    centro(0) = 335 - (246.81 - 26.1) : centro(1) = tt11 : centro(2) = 0  'arco
                    radio = 33.88
                    anginic = 0 ' angulos trabajan contra las abujas del relog
                    angfinal = -9.4012
                    arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
                    arco.color = 6

                    Dim tesxmh
                    tesxmh = tt1
                    insPoint(0) = -440   'Set insertion point x coordinate
                    insPoint(1) = (tesxmh + 35.26) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0

                    If CheckBox1.Checked = True Then ' condicion q si se escoje mayuscula los caracteres salgan en mayuscula
                        textStr = "TOMACORRIENTES"      'Set the text string
                    Else
                        textStr = "tomacorriente"
                    End If
                    exe.Cells(gg, 8) = UCase(textStr)
                    'Create Text object
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                    text.color = 2

                ElseIf kk = 5 Then

                    exe.Cells(gg, 1) = increment

                    tt1 = tt1 - jab
                    ttt1 = ttt1 - jabb
                    tt11 = tt11 - 150 ' reserva no lleva arco pero de todas formas es utilizado para ovilizar si este esta presente antes de un final
                    Dim tesxmhj

                    tesxmhj = tt1
                    insPoint(0) = -330  'Set insertion point x coordinate
                    insPoint(1) = (tesxmhj + 35.26) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0

                    If CheckBox1.Checked = True Then ' condicion q si se escoje mayuscula los caracteres salgan en mayuscula
                        textStr = "RESERVA"      'Set the text string
                    Else
                        textStr = "reserva"      'Set the text string
                    End If
                    exe.Cells(gg, 8) = UCase(textStr)
                    'Create Text object
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                    text.color = 2

                Else
                    tt1 = tt1 - jab
                    ttt1 = ttt1 - jabb

                End If


nm:

                If o4 = 4 Then 'algoritmo para movilizar un trifasico
                    yyy = yyy + 1
                    increment = increment + 1
                    'incrementt = increment + 2
                    'incrementtt = increment + 3
                    If yyy = 2 Then
                        o4 = 0
                        yyy = 0
                    End If
                    GoTo nmk
                End If 'fin del algoritmo

                If o2 = 2 Then
                    o2 = 0
                    increment = increment + 1
                    ' incrementt = increment + 2
                    '  incrementtt = increment + 3
                    GoTo nmk
                End If

                'segunda parte

                increment = increment + 1
                incrementt = increment + 2
                incrementtt = increment + 4
                gg = gg + 1

                On Error Resume Next ' iniciador de detector de errores
p22:
                kk = InputBox("Otra carga de 110V (4), Iluminacion (0), Dos fases(2), Tres fases(3),Tomacorriente(1),Reserva(5) ", "circuito numero " & increment)

                'corrector de errores.. si se introduce un tipo de variable q no relaciona con la vaibale repite l apregunta
                If Err.Number <> 0 Then
                    Err.Clear()
                    MsgBox("error ha introducido un valor caracter", MsgBoxStyle.Critical)
                    GoTo p22 ' vuelve a preguntar!!
                End If
                If kk = 4 Then

                    exe.Cells(gg, 1) = increment
                    exe.Cells(gg, 3) = "X"
                    tt2 = tt2 - jab
                    PtoIn(0) = 900 : PtoIn(1) = tt2 : PtoIn(2) = 0
                    PtoFin(0) = 563.5888 : PtoFin(1) = tt2 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6

                    ttt2 = ttt2 - jabb
                    PtoIn(0) = 900 : PtoIn(1) = ttt2 : PtoIn(2) = 0
                    PtoFin(0) = 582.75 : PtoFin(1) = ttt2 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6

                    'en este parte de esta opcion se calcula los (amp) en el programa
repit6:
                    va = InputBox("agrege carga(VA) de circuito " & increment, "carga(VA)")
                    exe.Cells(gg, 4) = va
                    suma_tomacorriente = suma_tomacorriente + va 'acumula cargas va de tomacorriente y tomacorriente personalizado
                    amp = (va) / (110) 'ecuacion q es utilizada para calcaular los amper(alumbrado,tomacorriente,(1))

                    Dim textf
                    textf = tt2 'distancia en el momento q esta parado( inicio ejemplo tt1= tt1 - jab) en ese momento
                    insPoint(0) = 585.42 'Set insertion point x coordinate
                    insPoint(1) = (textf + 38) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    If (amp > 0 And amp <= 20) Then
                        textStr = ("20A-1P")
                    ElseIf (amp > 20 And amp <= 30) Then
                        textStr = "30A-1P"
                    ElseIf (amp > 30 And amp <= 50) Then
                        textStr = "50A-1P"
                    ElseIf (amp > 50 And amp <= 60) Then
                        textStr = "60A-1P"
                    ElseIf (amp > 60 And amp <= 80) Then
                        textStr = "80A-1P"
                    ElseIf (amp > 80 And amp <= 100) Then
                        textStr = "100A-1P"
                    ElseIf (amp > 100 And amp <= 150) Then
                        textStr = "150A-1P"
                    ElseIf (amp > 150 And amp <= 175) Then
                        textStr = "175A-1P"
                    ElseIf (amp > 175 And amp <= 200) Then
                        textStr = "200A-1P"
                    ElseIf (amp > 200 And amp <= 225) Then
                        textStr = "225A-1P"
                    ElseIf (amp > 225 And amp <= 250) Then
                        textStr = "250A-1P"
                    ElseIf (amp > 250 And amp <= 275) Then
                        textStr = "275A-1P"
                    ElseIf (amp > 275 And amp <= 300) Then
                        textStr = "300A-1P"
                    ElseIf (amp > 300 And amp <= 350) Then
                        textStr = "350A-1P"
                    ElseIf (amp > 350 And amp <= 400) Then
                        textStr = "400A-1P"
                    ElseIf (amp > 400 And amp <= 450) Then
                        textStr = "450A-1P"

                    Else
                        GoTo repit6
                    End If
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight) 'Create Text object
                    text.color = 2
                    exe.Cells(gg, 6) = UCase(textStr)

                    If RadioButton1.Checked = True Then
                        If KTT = 1 Then
                            exe.Cells(gg, 7) = "A"
                        ElseIf KTT = 2 Then
                            exe.Cells(gg, 7) = "B"
                        Else
                            exe.Cells(gg, 7) = "C"
                        End If
                    End If

                    If RadioButton2.Checked = True Then
                        If KTT2 = 1 Then
                            exe.Cells(gg, 7) = "A"
                        ElseIf KTT2 = 2 Then
                            exe.Cells(gg, 7) = "B"
                        End If
                    End If

                    tt12 = tt12 - 150 ' arco de las fases NOTA: mejorar la creacion del arco..
                    centro(0) = 785.68 : centro(1) = tt12 : centro(2) = 0 'arco
                    radio = 33.88
                    anginic = 0 ' angulos trabajan contra las abujas del relog
                    angfinal = -9.4012
                    arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
                    arco.color = 6

                    Dim nombreu, tesxtu
                    tesxtu = tt2
                    nombreu = InputBox("introduzca nombre de circuito")
                    exe.Cells(gg, 8) = UCase(nombreu)

                    If CheckBox1.Checked = True Then ' condicion q si se escoje mayuscula los caracteres salgan en mayuscula
                        nombreu = UCase(nombreu)
                    Else
                    End If

                    insPoint(0) = 1000    'Set insertion point x coordinate
                    insPoint(1) = (tesxtu + 35.26) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    textStr = nombreu       'Set the text string
                    'Create Text object
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                    text.color = 2

                ElseIf kk = 0 Then
                    exe.Cells(gg, 2) = "X"
                    exe.Cells(gg, 1) = increment

                    If i = 1 Then
                        l2 = 2
                    End If

                    ttt2 = ttt2 - jabb
                    tt2 = tt2 - jab
                    PtoIn(0) = 900 : PtoIn(1) = tt2 : PtoIn(2) = 0
                    PtoFin(0) = 563.5888 : PtoFin(1) = tt2 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6

                    'en este parte de esta opcion se calcula los (amp) en el programa
repit7:
                    va = InputBox("agrege carga(VA) de circuito " & increment, "carga(VA)")
                    exe.Cells(gg, 4) = va
                    suma_iluminacion = suma_iluminacion + va 'sumatoria de carga va para iluminacion

                    amp = (va) / (110) 'ecuacion q es utilizada para calcaular los amper(alumbrado,tomacorriente,(1))

                    Dim texth
                    texth = tt2 'distancia en el momento q esta parado( inicio ejemplo tt1= tt1 - jab) en ese momento
                    insPoint(0) = 585.42 'Set insertion point x coordinate
                    insPoint(1) = (texth + 38) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    If (amp > 0 And amp <= 20) Then
                        textStr = ("20A-1P")
                    ElseIf (amp > 20 And amp <= 30) Then
                        textStr = "30A-1P"
                    ElseIf (amp > 30 And amp <= 50) Then
                        textStr = "50A-1P"
                    ElseIf (amp > 50 And amp <= 60) Then
                        textStr = "60A-1P"
                    ElseIf (amp > 60 And amp <= 80) Then
                        textStr = "80A-1P"
                    ElseIf (amp > 80 And amp <= 100) Then
                        textStr = "100A-1P"
                    ElseIf (amp > 100 And amp <= 150) Then
                        textStr = "150A-1P"
                    ElseIf (amp > 150 And amp <= 175) Then
                        textStr = "175A-1P"
                    ElseIf (amp > 175 And amp <= 200) Then
                        textStr = "200A-1P"
                    ElseIf (amp > 200 And amp <= 225) Then
                        textStr = "225A-1P"
                    ElseIf (amp > 225 And amp <= 250) Then
                        textStr = "250A-1P"
                    ElseIf (amp > 250 And amp <= 275) Then
                        textStr = "275A-1P"
                    ElseIf (amp > 275 And amp <= 300) Then
                        textStr = "300A-1P"
                    ElseIf (amp > 300 And amp <= 350) Then
                        textStr = "350A-1P"
                    ElseIf (amp > 350 And amp <= 400) Then
                        textStr = "400A-1P"
                    ElseIf (amp > 400 And amp <= 450) Then
                        textStr = "450A-1P"
                    Else
                        GoTo repit7
                    End If
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight) 'Create Text object
                    text.color = 2
                    exe.Cells(gg, 6) = UCase(textStr)

                    If RadioButton1.Checked = True Then
                        If KTT = 1 Then
                            exe.Cells(gg, 7) = "A"
                        ElseIf KTT = 2 Then
                            exe.Cells(gg, 7) = "B"
                        Else
                            exe.Cells(gg, 7) = "C"
                        End If
                    End If

                    If RadioButton2.Checked = True Then
                        If KTT2 = 1 Then
                            exe.Cells(gg, 7) = "A"
                        ElseIf KTT2 = 2 Then
                            exe.Cells(gg, 7) = "B"
                        End If
                    End If

                    tt12 = tt12 - 150 ' arco de las fases NOTA: mejorar la creacion del arco..
                    centro(0) = 785.68 : centro(1) = tt12 : centro(2) = 0 'arco
                    radio = 33.88
                    anginic = 0 ' angulos trabajan contra las abujas del relog
                    angfinal = -9.4012
                    arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
                    arco.color = 6

                    Dim tesxtukf
                    tesxtukf = tt2

                    insPoint(0) = 1000    'Set insertion point x coordinate
                    insPoint(1) = (tesxtukf + 35.26) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0

                    If CheckBox1.Checked = True Then ' condicion q si se escoje mayuscula los caracteres salgan en mayuscula
                        textStr = "ILUMINACION"     'Set the text string
                    Else
                        textStr = "iluminacion"     'Set the text string
                    End If
                    exe.Cells(gg, 8) = UCase(textStr)
                    'Create Text object
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                    text.color = 2

                    'ElseIf kk = "fin" Then Exit For

                ElseIf kk = 2 Then
                    exe.Cells(gg, 3) = "X"
                    exe.Cells(gg, 1) = increment & "," & incrementt

                    tt2 = tt2 - jak
                    ttt2 = ttt2 - jakk
                    o2 = 2

                    If CheckBox2.Checked = False Then
                        PtoIn(0) = 900 : PtoIn(1) = tt2 : PtoIn(2) = 0
                        PtoFin(0) = 582.75 : PtoFin(1) = tt2 : PtoFin(2) = 0
                        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                        line.color = 6
                    Else
                        'se selecciona la casilla "principal" en la interfacee 
                        PtoIn(0) = 900 : PtoIn(1) = tt2 : PtoIn(2) = 0
                        PtoFin(0) = 563.5888 : PtoFin(1) = tt2 : PtoFin(2) = 0
                        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                        line.color = 6

                        PtoIn(0) = 900 : PtoIn(1) = ttt2 : PtoIn(2) = 0
                        PtoFin(0) = 582.75 : PtoFin(1) = ttt2 : PtoFin(2) = 0
                        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                        line.color = 6
                    End If
                    'en este parte de esta opcion se calcula los (amp) en el programa
repit2:
                    va = InputBox("agrege carga(VA) de circuito " & increment & "," & incrementt, "carga(VA)")
                    exe.Cells(gg, 4) = va
                    suma_monofasico_trifasico = suma_monofasico_trifasico + va 'sumatoria de todas las cargas va en monofasico-trifasico

                    If RadioButton3.Checked = True Then
                        amp = (va) / (208) 'ecuacion q es utilizada para calcaular los amper(alumbrado,tomacorriente,(1))
                    ElseIf RadioButton4.Checked = True Then
                        amp = (va) / (240)
                    End If

                    Dim texkj
                    texkj = tt2 'distancia en el momento q esta parado( inicio ejemplo tt1= tt1 - jab) en ese momento
                    insPoint(0) = 585.42    'Set insertion point x coordinate
                    insPoint(1) = (texkj + 188) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    If (amp > 0 And amp <= 20) Then
                        textStr = ("20A-2P")
                    ElseIf (amp > 20 And amp <= 30) Then
                        textStr = "30A-2P"
                    ElseIf (amp > 30 And amp <= 50) Then
                        textStr = "50A-2P"
                    ElseIf (amp > 50 And amp <= 60) Then
                        textStr = "60A-2P"
                    ElseIf (amp > 60 And amp <= 80) Then
                        textStr = "80A-2P"
                    ElseIf (amp > 80 And amp <= 100) Then
                        textStr = "100A-2P"
                    ElseIf (amp > 100 And amp <= 150) Then
                        textStr = "150A-2P"
                    ElseIf (amp > 150 And amp <= 175) Then
                        textStr = "175A-2P"
                    ElseIf (amp > 175 And amp <= 200) Then
                        textStr = "200A-2P"
                    ElseIf (amp > 200 And amp <= 225) Then
                        textStr = "225A-2P"
                    ElseIf (amp > 225 And amp <= 250) Then
                        textStr = "250A-2P"
                    ElseIf (amp > 250 And amp <= 275) Then
                        textStr = "275A-2P"
                    ElseIf (amp > 275 And amp <= 300) Then
                        textStr = "300A-2P"
                    ElseIf (amp > 300 And amp <= 350) Then
                        textStr = "350A-2P"
                    ElseIf (amp > 350 And amp <= 400) Then
                        textStr = "400A-2P"
                    ElseIf (amp > 400 And amp <= 450) Then
                        textStr = "450A-2P"

                    Else
                        GoTo repit2
                    End If
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight) 'Create Text object
                    text.color = 2
                    exe.Cells(gg, 6) = UCase(textStr)


                    If RadioButton1.Checked = True Then
                        If KTT = 1 Then
                            exe.Cells(gg, 7) = "A,B"
                        ElseIf KTT = 2 Then
                            exe.Cells(gg, 7) = "B,C"
                        Else
                            exe.Cells(gg, 7) = "C,A"
                        End If
                    End If

                    If RadioButton2.Checked = True Then
                        If KTT2 = 1 Then
                            exe.Cells(gg, 7) = "A,B"
                        ElseIf KTT2 = 2 Then
                            exe.Cells(gg, 7) = "B,A"
                        End If
                    End If

                    For ii = 1 To 2
                        tt12 = tt12 - 150 ' arco de las fases NOTA: mejorar la creacion del arco..
                        centro(0) = 785.68 : centro(1) = tt12 : centro(2) = 0 'arco
                        radio = 33.88
                        anginic = 0 ' angulos trabajan contra las abujas del relog
                        angfinal = -9.4012
                        arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
                        arco.color = 6
                    Next ii

                    PtoIn(0) = 785.69 : PtoIn(1) = tt12 + 183.88 : PtoIn(2) = 0 '( coloca la barra vertical entre los acos monofasicos )
                    PtoFin(0) = 785.69 : PtoFin(1) = tt2 + 68.22 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6

                    Dim nombreer, tesxtt
                    tesxtt = tt2
                    nombreer = InputBox("introduzca nombre de circuito")
                    exe.Cells(gg, 8) = UCase(nombreer)
                    If CheckBox1.Checked = True Then ' condicion q si se escoje mayuscula los caracteres salgan en mayuscula
                        nombreer = UCase(nombreer)
                    Else
                    End If

                    insPoint(0) = 1000    'Set insertion point x coordinate
                    insPoint(1) = (tesxtt + 109.34) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    textStr = nombreer       'Set the text string
                    'Create Text object
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                    text.color = 2

                ElseIf kk = 3 Then
                    exe.Cells(gg, 3) = "X"
                    exe.Cells(gg, 1) = increment & "," & incrementt & "," & incrementtt

                    tt2 = tt2 - jal
                    ttt2 = ttt2 - jall
                    o4 = 4

                    If CheckBox2.Checked = False Then
                        PtoIn(0) = 900 : PtoIn(1) = tt2 : PtoIn(2) = 0
                        PtoFin(0) = 582.75 : PtoFin(1) = tt2 : PtoFin(2) = 0
                        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                        line.color = 6
                    Else
                        'se selecciona la casilla "principal" en la interfacee 
                        PtoIn(0) = 900 : PtoIn(1) = tt2 : PtoIn(2) = 0
                        PtoFin(0) = 563.5888 : PtoFin(1) = tt2 : PtoFin(2) = 0
                        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                        line.color = 6

                        PtoIn(0) = 900 : PtoIn(1) = ttt2 : PtoIn(2) = 0
                        PtoFin(0) = 582.75 : PtoFin(1) = ttt2 : PtoFin(2) = 0
                        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                        line.color = 6

                    End If
                    'en este parte de esta opcion se calcula los (amp) en el programa
repit3:
                    va = InputBox("agrege carga(VA) de circuito " & increment & "," & incrementt & "," & incrementtt, "carga(VA)")
                    exe.Cells(gg, 4) = va
                    suma_monofasico_trifasico = suma_monofasico_trifasico + va 'sumatoria de todas las cargas va en monofasico-trifasico
                    amp = (va) / (Math.Sqrt(3) * 208) 'ecuacion q es utilizada para calcaular los amper(alumbrado,tomacorriente,(1))

                    Dim texdsd

                    texdsd = tt2 'distancia en el momento q esta parado( inicio ejemplo tt1= tt1 - jab) en ese momento
                    insPoint(0) = 585.42    'Set insertion point x coordinate
                    insPoint(1) = (texdsd + 338) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    If (amp > 0 And amp <= 20) Then
                        textStr = ("20A-3P")
                    ElseIf (amp > 20 And amp <= 30) Then
                        textStr = "30A-3P"
                    ElseIf (amp > 30 And amp <= 50) Then
                        textStr = "50A-3P"
                    ElseIf (amp > 50 And amp <= 60) Then
                        textStr = "60A-3P"
                    ElseIf (amp > 60 And amp <= 80) Then
                        textStr = "80A-3P"
                    ElseIf (amp > 80 And amp <= 100) Then
                        textStr = "100A-3P"
                    ElseIf (amp > 100 And amp <= 150) Then
                        textStr = "150A-3P"
                    ElseIf (amp > 150 And amp <= 175) Then
                        textStr = "175A-3P"
                    ElseIf (amp > 175 And amp <= 200) Then
                        textStr = "200A-3P"
                    ElseIf (amp > 200 And amp <= 225) Then
                        textStr = "225A-3P"
                    ElseIf (amp > 225 And amp <= 250) Then
                        textStr = "250A-3P"
                    ElseIf (amp > 250 And amp <= 275) Then
                        textStr = "275A-3P"
                    ElseIf (amp > 275 And amp <= 300) Then
                        textStr = "300A-3P"
                    ElseIf (amp > 300 And amp <= 350) Then
                        textStr = "350A-3P"
                    ElseIf (amp > 350 And amp <= 400) Then
                        textStr = "400A-3P"
                    ElseIf (amp > 400 And amp <= 450) Then
                        textStr = "450A-3P"

                    Else
                        GoTo repit3
                    End If
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight) 'Create Text object
                    text.color = 2
                    exe.Cells(gg, 6) = UCase(textStr)

                    If RadioButton1.Checked = True Then
                        If KTT = 1 Then
                            exe.Cells(gg, 7) = "A,B,C"
                        ElseIf KTT = 2 Then
                            exe.Cells(gg, 7) = "B,C,A"
                        Else
                            exe.Cells(gg, 7) = "C,A,B"
                        End If
                    End If


                    For iii = 1 To 3
                        tt12 = tt12 - 150 ' arco de las fases NOTA: mejorar la creacion del arco..
                        centro(0) = 785.68 : centro(1) = tt12 : centro(2) = 0 'arco
                        radio = 33.88
                        anginic = 0 ' angulos trabajan contra las abujas del relog
                        angfinal = -9.4012
                        arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
                        arco.color = 6
                    Next iii

                    PtoIn(0) = 785.69 : PtoIn(1) = tt12 + 333.88 : PtoIn(2) = 0 '( coloca la barra vertical entre los acos monofasicos )
                    PtoFin(0) = 785.69 : PtoFin(1) = tt2 + 68.22 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6

                    Dim nombreee, tesxx

                    tesxx = tt2
                    nombreee = InputBox("introduzca nombre de circuito")
                    exe.Cells(gg, 8) = UCase(nombreee)
                    If CheckBox1.Checked = True Then ' condicion q si se escoje mayuscula los caracteres salgan en mayuscula
                        nombreee = UCase(nombreee)
                    Else
                    End If

                    insPoint(0) = 1000   'Set insertion point x coordinate
                    insPoint(1) = (tesxx + 184.34) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    textStr = nombreee       'Set the text string
                    'Create Text object
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                    text.color = 2

                ElseIf kk = 1 Then
                    exe.Cells(gg, 3) = "X"
                    exe.Cells(gg, 1) = increment

                    tt2 = tt2 - jab
                    PtoIn(0) = 900 : PtoIn(1) = tt2 : PtoIn(2) = 0
                    PtoFin(0) = 563.5888 : PtoFin(1) = tt2 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6

                    ttt2 = ttt2 - jabb
                    PtoIn(0) = 900 : PtoIn(1) = ttt2 : PtoIn(2) = 0
                    PtoFin(0) = 582.75 : PtoFin(1) = ttt2 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6
                    'en este parte de esta opcion se calcula los (amp) en el programa
repit8:
                    va = InputBox("agrege carga(VA) de circuito " & increment, "carga(VA)")
                    exe.Cells(gg, 4) = va
                    suma_tomacorriente = suma_tomacorriente + va 'acumula cargas va de tomacorriente y tomacorriente personalizado
                    amp = (va) / (110) 'ecuacion q es utilizada para calcaular los amper(alumbrado,tomacorriente,(1))

                    Dim textyt
                    textyt = tt2 'distancia en el momento q esta parado( inicio ejemplo tt1= tt1 - jab) en ese momento
                    insPoint(0) = 585.42 'Set insertion point x coordinate
                    insPoint(1) = (textyt + 38) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    If (amp > 0 And amp <= 20) Then
                        textStr = ("20A-1P")
                    ElseIf (amp > 20 And amp <= 30) Then
                        textStr = "30A-1P"
                    ElseIf (amp > 30 And amp <= 50) Then
                        textStr = "50A-1P"
                    ElseIf (amp > 50 And amp <= 60) Then
                        textStr = "60A-1P"
                    ElseIf (amp > 60 And amp <= 80) Then
                        textStr = "80A-1P"
                    ElseIf (amp > 80 And amp <= 100) Then
                        textStr = "100A-1P"
                    ElseIf (amp > 100 And amp <= 150) Then
                        textStr = "150A-1P"
                    ElseIf (amp > 150 And amp <= 175) Then
                        textStr = "175A-1P"
                    ElseIf (amp > 175 And amp <= 200) Then
                        textStr = "200A-1P"
                    ElseIf (amp > 200 And amp <= 225) Then
                        textStr = "225A-1P"
                    ElseIf (amp > 225 And amp <= 250) Then
                        textStr = "250A-1P"
                    ElseIf (amp > 250 And amp <= 275) Then
                        textStr = "275A-1P"
                    ElseIf (amp > 275 And amp <= 300) Then
                        textStr = "300A-1P"
                    ElseIf (amp > 300 And amp <= 350) Then
                        textStr = "350A-1P"
                    ElseIf (amp > 350 And amp <= 400) Then
                        textStr = "400A-1P"
                    ElseIf (amp > 400 And amp <= 450) Then
                        textStr = "450A-1P"

                    Else
                        GoTo repit8
                    End If
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight) 'Create Text object
                    text.color = 2
                    exe.Cells(gg, 6) = UCase(textStr)

                    If RadioButton1.Checked = True Then
                        If KTT = 1 Then
                            exe.Cells(gg, 7) = "A"
                        ElseIf KTT = 2 Then
                            exe.Cells(gg, 7) = "B"
                        Else
                            exe.Cells(gg, 7) = "C"
                        End If
                    End If

                    If RadioButton2.Checked = True Then
                        If KTT2 = 1 Then
                            exe.Cells(gg, 7) = "A"
                        ElseIf KTT2 = 2 Then
                            exe.Cells(gg, 7) = "B"
                        End If
                    End If

                    tt12 = tt12 - 150 ' arco de las fases NOTA: mejorar la creacion del arco..
                    centro(0) = 785.68 : centro(1) = tt12 : centro(2) = 0 'arco
                    radio = 33.88
                    anginic = 0 ' angulos trabajan contra las abujas del relog
                    angfinal = -9.4012
                    arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
                    arco.color = 6

                    Dim tesxtuk
                    tesxtuk = tt2
                    insPoint(0) = 1000    'Set insertion point x coordinate
                    insPoint(1) = (tesxtuk + 35.26) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    If CheckBox1.Checked = True Then ' condicion q si se escoje mayuscula los caracteres salgan en mayuscula
                        textStr = "TOMACORRIENTE"     'Set the text string
                    Else
                        textStr = "tomacorriente"     'Set the text string
                    End If
                    'Create Text object
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                    text.color = 2
                    exe.Cells(gg, 8) = UCase(textStr)
                ElseIf kk = 5 Then

                    exe.Cells(gg, 1) = increment

                    tt2 = tt2 - jab
                    ttt2 = ttt2 - jabb
                    tt12 = tt12 - 150 ' reserva no lleva arco pero de todas formas es utilizado para ovilizar si este esta presente antes de un final

                    Dim tesxtukl
                    tesxtukl = tt2
                    insPoint(0) = 1000    'Set insertion point x coordinate
                    insPoint(1) = (tesxtukl + 35.26) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    If CheckBox1.Checked = True Then ' condicion q si se escoje mayuscula los caracteres salgan en mayuscula
                        textStr = "RESERVA"     'Set the text string
                    Else
                        textStr = "reserva"
                    End If
                    exe.Cells(gg, 8) = UCase(textStr)
                    'Create Text object
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                    text.color = 2

                ElseIf kk = 8 Then
                    Exit For

                Else
                    tt2 = tt2 - jab
                    ttt2 = ttt2 - jabb

                End If
nmm:            ' si el primero es 4
nmk:            ' si el segundo es 4

                'PROCESO PARA PONER A,B,C EN EXCEL
                '*******************************************
                If RadioButton2.Checked = True Then
                    'CUANDO ES MONOFASICO
                    If KTT2 = 2 Then
                        KTT2 = 1
                        GoTo REINICIO
                    End If
                End If

                If RadioButton1.Checked = True Then
                    'CAUNDO ES TRIFASICO
                    If KTT = 3 Then
                        KTT = 1
                        GoTo REINICIO
                    End If
                End If
REINICIO:
            Next  'FINAL DEL CILCLO DEL ALGORITMO Q SABE CUANDO UN TRIFAASICO O MONOFACICO OUCAPA UN CIRCUITO!!
            '******************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
final:

            'instruccion q hace parte del esqueleto del tablero
            'barras dobles derecha e izquierda

            PtoIn(0) = 336.4112 : PtoIn(1) = (lon - 184.3425) : PtoIn(2) = 0 ' nota: en esta parte del programa  se formula para q la barra vertical de neutro siempre quede exacto con su barra horizontal y a su vez quede bn segun la longuitod q s eles de a las primeras tres barras verticales pricipales
            PtoFin(0) = 336.4112 : PtoFin(1) = (-lon + (lon - 97.8436)) : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6
            If l1 = 1 Then
                PtoIn(0) = 317.25 : PtoIn(1) = (lon - 355.95) : PtoIn(2) = 0
                PtoFin(0) = 317.25 : PtoFin(1) = -121.7616 : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.color = 6
            Else
                PtoIn(0) = 317.25 : PtoIn(1) = (lon - 206.25) : PtoIn(2) = 0
                PtoFin(0) = 317.25 : PtoFin(1) = -121.7616 : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.color = 6
            End If


            'SEGUNDA PARTE DEL TABLERO
            PtoIn(0) = 563.5888 : PtoIn(1) = (lon - 184.3425) : PtoIn(2) = 0 ' nota: en esta parte del programa  se formula para q la barra vertical de neutro siempre quede exacto con su barra horizontal y a su vez quede bn segun la longuitod q s eles de a las primeras tres barras verticales pricipales
            PtoFin(0) = 563.5888 : PtoFin(1) = (-lon + (lon - 97.8436)) : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6
            If l2 = 2 Then

                PtoIn(0) = 582.75 : PtoIn(1) = (lon - 335.95) : PtoIn(2) = 0
                PtoFin(0) = 582.75 : PtoFin(1) = -121.7616 : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.color = 6
            Else
                PtoIn(0) = 582.75 : PtoIn(1) = (lon - 206.25) : PtoIn(2) = 0
                PtoFin(0) = 582.75 : PtoFin(1) = -121.7616 : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.color = 6
            End If

            'lineas horizontales de neutro y tierra
            PtoIn(0) = 563.5888 : PtoIn(1) = (-lon + (lon - 97.8436)) : PtoIn(2) = 0
            PtoFin(0) = 336.4112 : PtoFin(1) = (-lon + (lon - 97.8436)) : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6

            PtoIn(0) = 582.75 : PtoIn(1) = (-lon + (lon - 121.7616)) : PtoIn(2) = 0
            PtoFin(0) = 0 : PtoFin(1) = (-lon + (lon - 121.7616)) : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6

            '*************************************************************************************************
            'CAJA COMPLETA DE NEUTRO Y PALABRA ("NUETRO")
            '****************************************************************************************************
            'de la linea vertical del neutro hasta toda la caja de neutro
            'linea vertical
            PtoIn(0) = 450 : PtoIn(1) = (-lon + (lon - 204.0488)) : PtoIn(2) = 0
            PtoFin(0) = 450 : PtoFin(1) = -97.8436 : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6

            'barra horizontal de caja de neutro
            PtoIn(0) = 527.8696 : PtoIn(1) = (-lon + (lon - 204.0488)) : PtoIn(2) = 0
            PtoFin(0) = 372.3104 : PtoFin(1) = (-lon + (lon - 204.0488)) : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6

            If CheckBox2.Checked = False Then 'SI  ES UN  SUBLTABLERO.
                PtoIn(0) = 527.8696 : PtoIn(1) = (-lon + (lon - 250.0502)) : PtoIn(2) = 0
                PtoFin(0) = 0 : PtoFin(1) = (-lon + (lon - 250.0502)) : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.color = 6
            Else
                'SI ES UN TABLERO PRINCIPAL..
                PtoIn(0) = 527.8696 : PtoIn(1) = (-lon + (lon - 250.0502)) : PtoIn(2) = 0
                PtoFin(0) = 372.3104 : PtoFin(1) = (-lon + (lon - 250.0502)) : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.color = 6
            End If

            insPoint(0) = 378  'Set insertion point x coordinate
            insPoint(1) = (-lon + (lon - 240.0502)) 'Set insertion point y coordinate
            insPoint(2) = 0      'Set insertion point z coordinate
            textHeight = 25                'Set text height to 1.0
            textStr = "NEUTRO"     'Set the text string
            'Create Text object
            text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
            text.color = 2

            ' barras verticales de caja de neutro
            PtoIn(0) = 372.134 : PtoIn(1) = (-lon + (lon - 250.0502)) : PtoIn(2) = 0
            PtoFin(0) = 372.134 : PtoFin(1) = -204.1113 : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6

            PtoIn(0) = 527.8696 : PtoIn(1) = (-lon + (lon - 250.0502)) : PtoIn(2) = 0
            PtoFin(0) = 527.8696 : PtoFin(1) = -204.1113 : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6
            '**************************************************************************************************************
            'FACTORES DE DEMANDA DE TMACORRIENTES,ILUMINACION Y FUERZA(MONOFASICO Y TRIFASICO)
            On Error Resume Next
p41:
            factor_demanda_tomacorriente = InputBox("ingrese factor de demanda de tomacorriente", "tableros")
            'corrector de errores.. si se introduce un tipo de variable q no relaciona con la vaibale repite l apregunta
            If Err.Number <> 0 Then
                Err.Clear()
                MsgBox("error ha introducido un valor caracter", MsgBoxStyle.Critical)
                GoTo p41
            End If

            On Error Resume Next
p31:
            factor_demanda_iluminacion = InputBox("ingrese factor de demanda de iluminacion", "tablaero")
            'corrector de errores.. si se introduce un tipo de variable q no relaciona con la vaibale repite l apregunta
            If Err.Number <> 0 Then
                Err.Clear()
                MsgBox("error ha introducido un valor caracter", MsgBoxStyle.Critical)
                GoTo p31
            End If

            On Error Resume Next
p21:
            factor_demanda_fuerza = InputBox("ingrese factor de demanda de fuerza", "tablero")
            'corrector de errores.. si se introduce un tipo de variable q no relaciona con la vaibale repite l apregunta
            If Err.Number <> 0 Then
                Err.Clear()
                MsgBox("error ha introducido un valor caracter", MsgBoxStyle.Critical)
                GoTo p21
            End If

            ' DEMANDAS DE TOMACORRIENTE,ILUMINACION Y FUERZA(MONOFASICO Y TRIFASICO)
            demanda_tomacorriente = (factor_demanda_tomacorriente / 100) * suma_tomacorriente
            exe.Cells(7, 3) = suma_tomacorriente & " x " & factor_demanda_tomacorriente & "%" & " = " & demanda_tomacorriente

            demanda_ilumiinacion = (factor_demanda_iluminacion / 100) * suma_iluminacion
            exe.Cells(8, 3) = suma_iluminacion & " x " & factor_demanda_iluminacion & "%" & " = " & demanda_ilumiinacion

            demanda_fuerza = (factor_demanda_fuerza / 100) * suma_monofasico_trifasico
            exe.Cells(9, 3) = suma_monofasico_trifasico & " x " & factor_demanda_fuerza & "%" & " = " & demanda_fuerza

            demanda_total = demanda_fuerza + demanda_ilumiinacion + demanda_tomacorriente
            exe.Cells(3, 8) = demanda_total & " VA"

            '**************************************************************************
            Dim KVA, CD, longitud, NAMP, NCD, j, MAC, R, X, DV, maxDV, INC As Double 'variables utilizadas para calculo de tablero y tipo de cable a utilizar!!

            On Error Resume Next
p61:
            maxDV = InputBox("ingrese maxima caida de tension permitida")
            'corrector de errores.. si se introduce un tipo de variable q no relaciona con la vaibale repite l apregunta
            If Err.Number <> 0 Then
                Err.Clear()
                MsgBox("error ha introducido un valor caracter", MsgBoxStyle.Critical)
                GoTo p61
            End If

            On Error Resume Next
p51:
            longitud = InputBox("ingrese longitud de alimentador")
            'corrector de errores.. si se introduce un tipo de variable q no relaciona con la vaibale repite l apregunta

            If Err.Number <> 0 Then
                Err.Clear()
                MsgBox("error ha introducido un valor caracter", MsgBoxStyle.Critical)
                GoTo p51
            End If

            KVA = demanda_total / 1000
            CD = longitud * KVA


            If RadioButton3.Checked = True Then
                KV = 0.208
            ElseIf RadioButton4.Checked = True Then
                KV = 0.24
            End If

            'se calcula la corriente total de la demanda..
            If RadioButton2.Checked = True Then

                I_total = KVA / KV

            ElseIf RadioButton1.Checked = True Then

                I_total = KVA / (Math.Sqrt(3) * KV)

            End If


            'luego criterio para selecion de cable neutro del alimentador...
            If RadioButton1.Checked = True And RadioButton3.Checked = True Then
                NAMP = 0.5 * I_total : NCD = 0.5 * CD

            ElseIf RadioButton2.Checked = True And RadioButton4.Checked = True And (I_total >= 200) Then
                NAMP = 0.7 * I_total : NCD = 0.7 * CD

            ElseIf RadioButton2.Checked = True And RadioButton4.Checked = True And (I_total < 200) Then
                NAMP = I_total : NCD = CD

            End If


            j = 1
            If (0 < NAMP) And (NAMP <= 25) Then
                MAC = 25

            ElseIf (25 < NAMP) And (NAMP <= 30) Then
                MAC = 30
                j = j + 2

            ElseIf (30 < NAMP) And (NAMP <= 50) Then
                MAC = 50
                j = j + 4

            ElseIf (50 < NAMP) And (NAMP <= 65) Then
                MAC = 65
                j = j + 6

            ElseIf (65 < NAMP) And (NAMP <= 85) Then
                MAC = 85
                j = j + 8

            ElseIf (85 < NAMP) And (NAMP <= 115) Then
                MAC = 115
                j = j + 10

            ElseIf (115 < NAMP) And (NAMP <= 150) Then
                MAC = 150
                j = j + 12

            ElseIf (150 < NAMP) And (NAMP <= 175) Then
                MAC = 175
                j = j + 14

            ElseIf (175 < NAMP) And (NAMP <= 200) Then
                MAC = 200
                j = j + 16

            ElseIf (200 < NAMP) And (NAMP <= 230) Then
                MAC = 230
                j = j + 18

            ElseIf (230 < NAMP) And (NAMP <= 255) Then
                MAC = 255
                j = j + 20

            ElseIf (255 < NAMP) And (NAMP <= 285) Then
                MAC = 285
                j = j + 22
                INC = 285
            ElseIf (285 < NAMP) And (NAMP <= 310) Then
                MAC = 310
                j = j + 24

            ElseIf (310 < NAMP) And (NAMP <= 335) Then
                MAC = 335
                j = j + 26

            ElseIf (335 < NAMP) And (NAMP <= 380) Then
                MAC = 380
                j = j + 28

            ElseIf (380 < NAMP) And (NAMP <= 420) Then
                MAC = 420
                j = j + 30

            ElseIf (420 < NAMP) And (NAMP <= 460) Then
                MAC = 460
                j = j + 32

            ElseIf (460 < NAMP) And (NAMP <= 475) Then
                MAC = 475
                j = j + 34

            ElseIf NAMP > 475 Then
                MsgBox("Se sugiere utilizar Conductores en Paralelo")
            End If
denuevo:
            If j = 1 Then
                R = 1968 : X = 58.4 : MAC = 25
                INC = 25
            ElseIf j = 3 Then
                R = 1230 : X = 56.4 : MAC = 30
                INC = 30
            ElseIf j = 5 Then
                R = 789 : X = 55.3 : MAC = 50
                INC = 50
            ElseIf j = 7 Then
                R = 490 : X = 51.2 : MAC = 65
                INC = 65
            ElseIf j = 9 Then
                R = 318 : X = 47.3 : MAC = 85
                INC = 85
            ElseIf j = 11 Then
                R = 203 : X = 43.8 : MAC = 115
                INC = 115
            ElseIf j = 13 Then
                R = 129 : X = 41.5 : MAC = 150
                INC = 150
            ElseIf j = 15 Then
                R = 103 : X = 40.9 : MAC = 175
                INC = 175
            ElseIf j = 17 Then
                R = 80.3 : X = 40.2 : MAC = 200
                INC = 200
            ElseIf j = 19 Then
                R = 66.6 : X = 39.1 : MAC = 230
                INC = 230
            ElseIf j = 21 Then
                R = 57.8 : X = 39 : MAC = 255
                INC = 255
            ElseIf j = 23 Then
                R = 50.1 : X = 38.7 : MAC = 285
                INC = 285
            ElseIf j = 25 Then
                R = 38 : X = 38.4 : MAC = 310
                INC = 310
            ElseIf j = 27 Then
                R = 35.6 : X = 38.1 : MAC = 335
                INC = 335
            ElseIf j = 29 Then
                R = 27.5 : X = 36.6 : MAC = 380
                INC = 380
            ElseIf j = 31 Then
                R = 24.1 : X = 36.4 : MAC = 420
                INC = 420
            ElseIf j = 33 Then
                R = 24.7 : X = 35.8 : MAC = 460
                INC = 460
            ElseIf j = 35 Then
                R = 19.8 : X = 35.3 : MAC = 475
                INC = 475
            ElseIf j > 35 Then
                MsgBox("Se sugiere utilizar Cables en Paralelo")

            End If

            If RadioButton1.Checked = True Then
                DV = (NCD * (R * (fp / 100) + X * Math.Sqrt(1 - ((fp / 100) ^ 2))) / (10 * (KV ^ 2))) * (0.0328083989 * 10 ^ -4)

            ElseIf RadioButton2.Checked = True Then
                DV = 2 * (NCD * (R * (fp / 100) + X * Math.Sqrt(1 - ((fp / 100) ^ 2))) / (10 * (KV ^ 2))) * (0.0328083989 * 10 ^ -4)
            End If



            If DV > maxDV Then
                j = j + 2 : GoTo denuevo

            End If
            exe.Cells(10, 7) = DV
            exe.Cells(8, 7) = (NAMP / INC) * 100 & "%"

            insPoint(0) = 650    'Set insertion point x coordinate
            insPoint(1) = lon - (lon + 400) 'Set insertion point y coordinate
            insPoint(2) = 0      'Set insertion point z coordinate
            textHeight = 25                'Set text height to 1.0
            If j = 1 Then
                textStr = "+ 1 CABLE THW  #12 AWG DE COBRE(NEUTRO)"
            ElseIf j = 3 Then
                textStr = ("+ 1 CABLE  THW #10 AWG DE COBRE(NEUTRO)")
            ElseIf j = 5 Then
                textStr = ("+ 1 CABLE  THW  #8  AWG DE COBRE(NEUTRO)")
            ElseIf j = 7 Then
                textStr = ("+ 1 CABLE THW  #6  AWG DE COBRE(NEUTRO)")
            ElseIf j = 9 Then
                textStr = ("+ 1 CABLE THW  #4  AWG DE COBRE(NEUTRO)")
            ElseIf j = 11 Then
                textStr = ("+ 1 CABLE  THW  #2  AWG DE COBRE(NEUTRO)")
            ElseIf j = 13 Then
                textStr = ("+ 1 CABLE  THW  #1/0 AWG DE COBRE(NEUTRO)")
            ElseIf j = 15 Then
                textStr = ("+ 1 CABLE  THW  #2/0 AWG DE COBRE(NEUTRO)")
            ElseIf j = 17 Then
                textStr = ("+ 1 CABLE  THW  #3/0 AWG DE COBRE(NEUTRO)")
            ElseIf j = 19 Then
                textStr = ("+ 1 CABLE  THW  #4/0 AWG DE COBRE(NEUTRO)")
            ElseIf j = 21 Then
                textStr = ("+ 1 CABLE  THW  #250 MCM DE COBRE(NEUTRO)")
            ElseIf j = 23 Then
                textStr = ("+ 1 CABLE  THW  #300 MCM DE COBRE(NEUTRO)")
            ElseIf j = 25 Then
                textStr = ("+ 1 CABLE  THW  #350 MCM DE COBRE(NEUTRO)")
            ElseIf j = 27 Then
                textStr = ("+ 1 CABLE  THW #400 MCM DE COBRE(NEUTRO)")
            ElseIf j = 29 Then
                textStr = ("+ 1 CABLE  THW #500 MCM DE COBRE(NEUTRO)")
            ElseIf j = 31 Then
                textStr = ("+ 1 CABLE  THW #600 MCM DE COBRE(NEUTRO)")
            ElseIf j = 33 Then
                textStr = ("+ 1 CABLE  THW #700 MCM DE COBRE(NEUTRO)")
            ElseIf j = 35 Then
                textStr = ("+ 1 CABLE  THW #750 MCM DE COBRE(NEUTRO)")
            End If
            text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
            text.color = 2

            exe.Cells(7, 8) = textStr


            i = 1
            If (0 < I_total) And (I_total <= 25) Then
                MAC = 20

            ElseIf (25 < I_total) And (I_total <= 30) Then
                MAC = 30
                i = i + 2

            ElseIf (30 < I_total) And (I_total <= 50) Then
                MAC = 50
                i = i + 4

            ElseIf (50 < I_total) And (I_total <= 65) Then
                MAC = 60
                i = i + 6

            ElseIf (65 < I_total) And (I_total <= 85) Then
                MAC = 80
                i = i + 8

            ElseIf (85 < I_total) And (I_total <= 115) Then
                MAC = 100
                i = i + 10

            ElseIf (115 < I_total) And (I_total <= 150) Then
                MAC = 150
                i = i + 12

            ElseIf (150 < I_total) And (I_total <= 175) Then
                MAC = 175
                i = i + 14

            ElseIf (175 < I_total) And (I_total <= 200) Then
                MAC = 200
                i = i + 16

            ElseIf (200 < I_total) And (I_total <= 230) Then
                MAC = 225
                i = i + 18

            ElseIf (230 < I_total) And (I_total <= 255) Then
                MAC = 250
                i = i + 20

            ElseIf (255 < I_total) And (I_total <= 285) Then
                MAC = 275
                i = i + 22

            ElseIf (285 < I_total) And (I_total <= 310) Then
                MAC = 300
                i = i + 24

            ElseIf (310 < I_total) And (I_total <= 335) Then
                MAC = 300
                i = i + 26

            ElseIf (335 < I_total) And (I_total <= 380) Then
                MAC = 350
                i = i + 28

            ElseIf (380 < I_total) And (I_total <= 420) Then
                MAC = 400
                i = i + 30

            ElseIf (420 < I_total) And (I_total <= 460) Then
                MAC = 450
                i = i + 32

            ElseIf (460 < I_total) And (I_total <= 475) Then
                MAC = 450
                i = i + 34

            ElseIf I_total > 475 Then
                MsgBox("Se sugiere utilizar Conductores en Paralelo")
            End If
            Dim numero(40), p As String
            Dim l_caracter As String
            p = 0
denuevoo:
            'calculo de fases..
            If i = 1 Then

                R = 1968 : X = 58.4 : MAC = 20
                numero(1) = " 12 AWG"
                INC = 25
            ElseIf i = 3 Then
                R = 1230 : X = 56.4
                If p = 0 Then
                    MAC = 30
                End If
                numero(3) = "10 AWG"
                INC = 30
            ElseIf i = 5 Then
                R = 789 : X = 55.3
                If p = 0 Then
                    MAC = 50
                End If
                numero(5) = "8 AWG"
                INC = 50
            ElseIf i = 7 Then
                R = 490 : X = 51.2
                If p = 0 Then
                    MAC = 60
                End If
                numero(7) = "6 AWG"
                INC = 65
            ElseIf i = 9 Then
                R = 318 : X = 47.3
                If p = 0 Then
                    MAC = 80
                End If
                numero(9) = "4 AWG"
                INC = 85
            ElseIf i = 11 Then
                R = 203 : X = 43.8
                If p = 0 Then
                    MAC = 100
                End If
                numero(11) = "2 AWG"
                INC = 115
            ElseIf i = 13 Then
                R = 129 : X = 41.5
                If p = 0 Then
                    MAC = 150
                End If
                numero(13) = "1/0 AWG"
                INC = 150
            ElseIf i = 15 Then
                R = 103 : X = 40.9
                If p = 0 Then
                    MAC = 175
                End If
                numero(15) = "2/0 AWG"
                INC = 175
            ElseIf i = 17 Then
                R = 80.3 : X = 40.2
                If p = 0 Then
                    MAC = 200
                End If
                numero(17) = "3/0 AWG"
                INC = 200
            ElseIf i = 19 Then
                R = 66.6 : X = 39.1
                If p = 0 Then
                    MAC = 225
                End If
                numero(19) = "4/0 AWG"
                INC = 230
            ElseIf i = 21 Then
                R = 57.8 : X = 39
                If p = 0 Then
                    MAC = 250
                End If
                numero(21) = "250 MCM"
                INC = 255
            ElseIf i = 23 Then
                R = 50.1 : X = 38.7
                If p = 0 Then
                    MAC = 275
                End If
                numero(23) = "300 MCM"
                INC = 285
            ElseIf i = 25 Then
                R = 38 : X = 38.4
                If p = 0 Then
                    MAC = 300
                End If
                numero(25) = "350 MCM"
                INC = 310
            ElseIf i = 27 Then
                R = 35.6 : X = 38.1
                If p = 0 Then
                    MAC = 300
                End If
                numero(27) = "400 MCM"
                INC = 335
            ElseIf i = 29 Then
                If p = 0 Then
                    R = 27.5 : X = 36.6
                    MAC = 350
                End If
                numero(29) = "500 MCM"
                INC = 380
            ElseIf i = 31 Then
                R = 24.1 : X = 36.4
                If p = 0 Then
                    MAC = 400
                End If
                numero(31) = "600 MCM"
                INC = 420
            ElseIf i = 33 Then
                R = 24.7 : X = 35.8
                If p = 0 Then
                    MAC = 450
                End If
                numero(33) = "700 MCM"
                INC = 460
            ElseIf i = 35 Then
                R = 19.8 : X = 35.3
                If p = 0 Then
                    MAC = 450
                End If
                numero(35) = "750 MCM"
                INC = 475
            ElseIf i > 35 Then
                MsgBox("Se sugiere utilizar Cables en Paralelo")
            End If


            If RadioButton1.Checked = True Then
                DV = (CD * (R * (fp / 100) + X * Math.Sqrt(1 - ((fp / 100) ^ 2))) / (10 * (KV ^ 2))) * (0.0328083989 * 10 ^ -4)
            ElseIf RadioButton2.Checked = True Then
                DV = 2 * (CD * (R * (fp / 100) + X * Math.Sqrt(1 - ((fp / 100) ^ 2))) / (10 * (KV ^ 2))) * (0.0328083989 * 10 ^ -4)
            End If

            If DV > maxDV Then
                i = i + 2
                p = 1
                GoTo denuevoo
            End If
            exe.Cells(9, 7) = DV

            exe.Cells(7, 7) = (I_total / INC) * 100 & "%"

            If CheckBox2.Checked = False Then
                BP = MAC
            End If


            If RadioButton5.Checked = True Then
                l_caracter = "NLAB"
            ElseIf RadioButton6.Checked = True Then
                l_caracter = "NAB"
            End If

            If RadioButton2.Checked = True Then


                insPoint(0) = -600   'Set insertion point x coordinate
                insPoint(1) = lon - (lon + 500) 'Set insertion point y coordinate
                insPoint(2) = 0      'Set insertion point z coordinate
                textHeight = 25                'Set text height to 1.0
                textStr = "TABLERO  TIPO:  " & l_caracter & "3" & ncircuitos2 & " DE " & ncircuitos2 & " CIRCUITOS, 2 FASES,  CON  BARRAS  DE  NEUTRO Y TIERRA."
                text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                text.color = 2

                insPoint(0) = -600   'Set insertion point x coordinate
                insPoint(1) = lon - (lon + 400) 'Set insertion point y coordinate
                insPoint(2) = 0      'Set insertion point z coordinate
                textHeight = 25                'Set text height to 1.0
                textStr = ("ALIMENTADOR: 2 CABLES THW  # " & numero(i) & " DE COBRE(FASES)")
                text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                text.color = 2

                'resultado sin la palabra "alimentaro" para la hoja de excel
                textStr = (" 2 CABLES THW  # " & numero(i) & " DE COBRE(FASES)")
                exe.Cells(6, 8) = textStr

                insPoint(0) = 600   'Set insertion point x coordinate
                insPoint(1) = lon + 50 'Set insertion point y coordinate
                insPoint(2) = 0      'Set insertion point z coordinate
                textHeight = 25                'Set text height to 1.0
                textStr = (BP & "A-2P")
                text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                text.color = 2
                exe.Cells(5, 8) = textStr

            ElseIf RadioButton1.Checked = True Then
                insPoint(0) = -600    'Set insertion point x coordinate
                insPoint(1) = lon - (lon + 500) 'Set insertion point y coordinate
                insPoint(2) = 0      'Set insertion point z coordinate
                textHeight = 25                'Set text height to 1.0
                textStr = "TABLERO  TIPO:  " & l_caracter & "4" & ncircuitos2 & " DE " & ncircuitos2 & " CIRCUITOS, 3 FASES,  CON  BARRAS  DE  NEUTRO Y TIERRA."
                text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                text.color = 2

                insPoint(0) = -600    'Set insertion point x coordinate
                insPoint(1) = lon - (lon + 400) 'Set insertion point y coordinate
                insPoint(2) = 0      'Set insertion point z coordinate
                textHeight = 25                'Set text height to 1.0
                textStr = ("ALIMENTADOR: 3 CABLES THW #" & numero(i) & " DE COBRE(FASES)")
                text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                text.color = 2

                'resultado sin la palabra "alimentaro" para la hoja de excel
                textStr = (" 3 CABLES THW #" & numero(i) & " DE COBRE(FASES)")
                exe.Cells(6, 8) = textStr

                insPoint(0) = 600    'Set insertion point x coordinate
                insPoint(1) = lon + 50  'Set insertion point y coordinate
                insPoint(2) = 0      'Set insertion point z coordinate
                textHeight = 25                'Set text height to 1.0
                textStr = (BP & "A-3P")
                text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                text.color = 2
                exe.Cells(5, 8) = textStr

            End If


            insPoint(0) = -311    'Set insertion point x coordinate
            insPoint(1) = lon - (lon + 450) 'Set insertion point y coordinate
            insPoint(2) = 0      'Set insertion point z coordinate
            textHeight = 25                'Set text height to 1.0
            If 0 < BP And BP <= 20 Then
                textStr = "+ THW # 12 AWG COBRE(TIERRA) EN TUBERIA Ø    PVC."
            ElseIf 20 < BP And BP <= 60 Then
                textStr = "+ THW # 10 AWG COBRE(TIERRA) EN TUBERIA Ø    PVC."
            ElseIf 60 < BP And BP <= 100 Then
                textStr = "+ THW # 8  AWG COBRE(TIERRA) EN TUBERIA Ø    PVC."
            ElseIf 100 < BP And BP <= 200 Then
                textStr = "+ THW # 6  AWG COBRE(TIERRA) EN TUBERIA Ø    PVC."
            ElseIf 200 < BP And BP <= 300 Then
                textStr = "+ THW # 4  AWG COBRE(TIERRA) EN TUBERIA Ø    PVC."
            ElseIf 300 < BP And BP <= 500 Then
                textStr = "+ THW # 2  AWG COBRE(TIERRA) EN TUBERIA Ø    PVC."
            ElseIf 500 < BP And BP <= 800 Then
                textStr = "+ THW # 1/0 AWG COBRE(TIERRA) EN TUBERIA Ø    PVC."
            ElseIf 800 < BP And BP <= 1000 Then
                textStr = "+ THW # 2/0 AWG COBRE(TIERRA) EN TUBERIA Ø    PVC."
            ElseIf 1000 < BP And BP <= 1200 Then
                textStr = "+ THW # 3/0 AWG COBRE(TIERRA) EN TUBERIA Ø    PVC."
            ElseIf 1200 < BP And BP <= 1600 Then
                textStr = "+ THW # 4/0 AWG COBRE(TIERRA) EN TUBERIA Ø    PVC."
            ElseIf 1600 < BP And BP <= 2000 Then
                textStr = "+ THW # 250 MCM COBRE(TIERRA) EN TUBERIA Ø    PVC."
            ElseIf 2000 < BP And BP <= 2500 Then
                textStr = "+ THW # 350 MCM  AWG COBRE(TIERRA) EN TUBERIA Ø    PVC."
            ElseIf 2500 < BP And BP <= 3000 Then
                textStr = "+ THW # 400 MCM COBRE(TIERRA) EN TUBERIA Ø    PVC."
            ElseIf 3000 < BP And BP <= 4000 Then
                textStr = "+ THW # 500 MCM COBRE(TIERRA) EN TUBERIA Ø    PVC."
            ElseIf 4000 < BP And BP <= 5000 Then
                textStr = "+ THW # 700 MCM COBRE(TIERRA) EN TUBERIA Ø    PVC."
            ElseIf 5000 < BP And BP <= 6000 Then
                textStr = "+ THW # 800 MCM COBRE(TIERRA) EN TUBERIA Ø    PVC."
            End If
            text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
            text.color = 2
            exe.Cells(8, 8) = textStr

            If CheckBox2.Checked = False Then
                insPoint(0) = -550  'Set insertion point x coordinate
                insPoint(1) = lon - (lon + 110) 'Set insertion point y coordinate
                insPoint(2) = 0      'Set insertion point z coordinate
                textHeight = 25                'Set text height to 1.0
                textStr = "A BARRA DE TIERRA EN TABLERO PRINCIPAL"
                text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                text.color = 2

                insPoint(0) = -550  'Set insertion point x coordinate
                insPoint(1) = lon - (lon + 240) 'Set insertion point y coordinate
                insPoint(2) = 0      'Set insertion point z coordinate
                textHeight = 25                'Set text height to 1.0
                textStr = "A BARRA DE NEUTRO EN TABLERO PRINCIPAL"
                text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                text.color = 2
            Else
                'barrita vertical de la figura triangular cuando es principal 
                PtoIn(0) = 0 : PtoIn(1) = (lon - (lon + 170.25)) : PtoIn(2) = 0
                PtoFin(0) = 0 : PtoFin(1) = (lon - (lon + 121.7616)) : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.color = 6

                'primera barra
                PtoIn(0) = 118.5625 : PtoIn(1) = (lon - (lon + 170.25)) : PtoIn(2) = 0
                PtoFin(0) = 0 : PtoFin(1) = (lon - (lon + 170.25)) : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.color = 6

                PtoIn(0) = -118.5625 : PtoIn(1) = (lon - (lon + 170.25)) : PtoIn(2) = 0
                PtoFin(0) = 0 : PtoFin(1) = (lon - (lon + 170.25)) : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.color = 6

                'segunda barra

                PtoIn(0) = 94.85 : PtoIn(1) = (lon - (lon + 170.25 + 14.5)) : PtoIn(2) = 0
                PtoFin(0) = 0 : PtoFin(1) = (lon - (lon + 170.25 + 14.5)) : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.color = 6

                PtoIn(0) = -94.85 : PtoIn(1) = (lon - (lon + 170.25 + 14.5)) : PtoIn(2) = 0
                PtoFin(0) = 0 : PtoFin(1) = (lon - (lon + 170.25 + 14.5)) : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.color = 6

                'tercera barrita
                PtoIn(0) = 71.1375 : PtoIn(1) = (lon - (lon + 170.25 + (14.5 * 2))) : PtoIn(2) = 0
                PtoFin(0) = 0 : PtoFin(1) = (lon - (lon + 170.25 + (14.5 * 2))) : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.color = 6

                PtoIn(0) = -71.1375 : PtoIn(1) = (lon - (lon + 170.25 + (14.5 * 2))) : PtoIn(2) = 0
                PtoFin(0) = 0 : PtoFin(1) = (lon - (lon + 170.25 + (14.5 * 2))) : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.color = 6

                'cuarta barrita
                PtoIn(0) = 47.425 : PtoIn(1) = (lon - (lon + 170.25 + (14.5 * 3))) : PtoIn(2) = 0
                PtoFin(0) = 0 : PtoFin(1) = (lon - (lon + 170.25 + (14.5 * 3))) : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.color = 6

                PtoIn(0) = -47.425 : PtoIn(1) = (lon - (lon + 170.25 + (14.5 * 3))) : PtoIn(2) = 0
                PtoFin(0) = 0 : PtoFin(1) = (lon - (lon + 170.25 + (14.5 * 3))) : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.color = 6

                'quinta barrita

                PtoIn(0) = 23.7125 : PtoIn(1) = (lon - (lon + 170.25 + (14.5 * 4))) : PtoIn(2) = 0
                PtoFin(0) = 0 : PtoFin(1) = (lon - (lon + 170.25 + (14.5 * 4))) : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.color = 6

                'sexta barrita
                PtoIn(0) = 3.2759 : PtoIn(1) = (lon - (lon + 170.25 + (14.5 * 5))) : PtoIn(2) = 0
                PtoFin(0) = 0 : PtoFin(1) = (lon - (lon + 170.25 + (14.5 * 5))) : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.color = 6

                PtoIn(0) = -3.2759 : PtoIn(1) = (lon - (lon + 170.25 + (14.5 * 5))) : PtoIn(2) = 0
                PtoFin(0) = 0 : PtoFin(1) = (lon - (lon + 170.25 + (14.5 * 5))) : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.color = 6

            End If

            'ARCOS Q VAN ARRIBA DEL TABLERO
            centro(0) = 375 : centro(1) = lon + 33.88 : centro(2) = 0 'arco
            radio = 33.88
            anginic = 300 ' angulos trabajan contra  las abujas del relog
            angfinal = -300
            arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
            arco.color = 6

            'si es trifasico....
            If RadioButton1.Checked = True Then
                centro(0) = 375 + 75 : centro(1) = lon + 33.88 : centro(2) = 0 'arco
                radio = 33.88
                anginic = 300 ' angulos trabajan contra  las abujas del relog
                angfinal = -300
                arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
                arco.color = 6
            End If

            centro(0) = 375 + 75 * 2 : centro(1) = lon + 33.88 : centro(2) = 0 'arco
            radio = 33.88
            anginic = 300 ' angulos trabajan contra  las abujas del relog
            angfinal = -300
            arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
            arco.color = 6

            'LINEA Q PASA POR LOS TRES ARCOS
            PtoIn(0) = 558.88 : PtoIn(1) = lon + 33.88 : PtoIn(2) = 0
            PtoFin(0) = 408.88 : PtoFin(1) = lon + 33.88 : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6

            'lineas q  estan arriba de los arcos..
            PtoIn(0) = 375 : PtoIn(1) = lon + 124 : PtoIn(2) = 0
            PtoFin(0) = 375 : PtoFin(1) = lon + 67.76 : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 150

            If RadioButton1.Checked = True Then
                PtoIn(0) = 375 + 75 : PtoIn(1) = lon + 124 : PtoIn(2) = 0
                PtoFin(0) = 375 + 75 : PtoFin(1) = lon + 67.76 : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.color = 150
            End If

            PtoIn(0) = 375 + 75 * 2 : PtoIn(1) = lon + 124 : PtoIn(2) = 0
            PtoFin(0) = 375 + 75 * 2 : PtoFin(1) = lon + 67.76 : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 150

            insPoint(0) = 365 'Set insertion point x coordinate
            insPoint(1) = lon + 140 'Set insertion point y coordinate
            insPoint(2) = 0      'Set insertion point z coordinate
            textHeight = 25                'Set text height to 1.0
            textStr = "A"
            text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
            text.color = 2

            If RadioButton1.Checked = True Then
                insPoint(0) = 365 + 75 'Set insertion point x coordinate
                insPoint(1) = lon + 140 'Set insertion point y coordinate
                insPoint(2) = 0      'Set insertion point z coordinate
                textHeight = 25                'Set text height to 1.0
                textStr = "B"
                text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                text.color = 2
            End If

            insPoint(0) = 365 + 75 * 2 'Set insertion point x coordinate
            insPoint(1) = lon + 140 'Set insertion point y coordinate
            insPoint(2) = 0      'Set insertion point z coordinate
            textHeight = 25                'Set text height to 1.0
            If RadioButton1.Checked = True Then
                textStr = "C"
            Else
                textStr = "B"
            End If
            text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
            text.color = 2

            centroo(0) = 375 : centroo(1) = lon + 124 : centroo(2) = 0
            radioo = 7.8125
            circulo = objautocad.ActiveDocument.ModelSpace.AddCircle(centroo, radioo)
            circulo.color = 150

            If RadioButton1.Checked = True Then
                centroo(0) = 375 + 75 : centroo(1) = lon + 124 : centroo(2) = 0
                radioo = 7.8125
                circulo = objautocad.ActiveDocument.ModelSpace.AddCircle(centroo, radioo)
                circulo.color = 150
            End If

            centroo(0) = 375 + 75 * 2 : centroo(1) = lon + 124 : centroo(2) = 0
            radioo = 7.8125
            circulo = objautocad.ActiveDocument.ModelSpace.AddCircle(centroo, radioo)
            circulo.color = 150

            'instruccion para centrar el dibujo en el plano..
            objautocad.ZoomExtents()

            MsgBox("Dibujo hecho satisfactoriamente")
            ' Y POR ULTIMO MAXIMIZA LA VENTANA DE DIBUJO(AUTOCAD)

            objautocad.WindowState = Autodesk.AutoCAD.Interop.Common.AcWindowState.acMax
            exe.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized
        End If

    End Sub

    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        MsgBox("version de programa: " & My.Application.Info.Version.Major)

        ' shell :funcion q se uiliza para abrir cualquier archivo

        'para saber en q fecha pasa algo
        'Label2.Text = Today (fecha)

        'If Today = ("24 / 3 / 2011") Then
        'Box("hey ya tienes q pagar")

        'End If

        'sender abre la hoja de calculo  estandar..

        'crea un nuevo espacio de trabajo en excel
        'exe.Workbooks.Add()
        'comando que guarda el archivo segun la ruta en q lo coloques
        ' exe.Application.ActiveWorkbook.Save()
        ' para guardar en autocad
        'objautocad.Application.ActiveDocument.Save()
        'para guardar utiliando objeto de file..
        'objautocad.Application.ActiveDocument.Save()



        'codigo para password
        '  Dim pass As String
        ' Dim h As Integer
        'h = 1
        'Do While h <= 3
        'pass = InputBox("Ingrese pass para entrar", "tableros electrico " & "intento: " & h)

        '       If pass = "1991semeco" Then
        'MsgBox("pass correcto", MsgBoxStyle.Exclamation)
        'Exit Do
        'Else
        'MsgBox(" pass erroneo", MsgBoxStyle.Critical)

        'End If
        'h = h + 1
        'Loop
        'If h = 4 Then
        'End
        'End If

        ' agrega los capos al combobox
        ComboBox1.Items.Add(" 3X15 KVA")
        ComboBox1.Items.Add(" 3X25 KVA")
        ComboBox1.Items.Add(" 3X37,5 KVA")
        ComboBox1.Items.Add(" 3X50 KVA")
        ComboBox1.Items.Add(" 3X75 KVA")
        ComboBox1.Items.Add(" 3X100 KVA")
        ComboBox1.Items.Add(" 3X125 KVA")
        ComboBox1.Items.Add(" 3X300 KVA")
        ComboBox1.Items.Add(" 3X500 KVA")
        ComboBox1.Items.Add(" 3X1000 KVA")
        ComboBox1.Items.Add("OTRA CAPACIDAD...")

        ' agrega los capos al combobox
        ComboBox2.Items.Add("15 KVA")
        ComboBox2.Items.Add("25 KVA")
        ComboBox2.Items.Add("37.5 KVA")
        ComboBox2.Items.Add("50 KVA")
        ComboBox2.Items.Add("75 KVA")
        ComboBox2.Items.Add("OTRA CAPACIDAD...")


    End Sub

    Private Sub MenuStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
     
    End Sub

    Private Sub SalirToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalirToolStripMenuItem.Click
        'se crea un nuevo es pacio para dibujar...

        objautocad = CreateObject("autoCAD.application", "")
        'instruccion q se utiliza para minimizar el autocad  cuando inicie
        objautocad.WindowState = Autodesk.AutoCAD.Interop.Common.AcWindowState.acMin
        'para q el close se active
        encender = 1
        If encender = 1 Then
            AbrirToolStripMenuItem.Enabled = True
        End If


    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged

        'metodo para q se active el trifasico y monofacio (combobox.visibl) si se maraca conjunto con checkbox "principal"
        If RadioButton1.Checked = True Then
            RadioButton3.Checked = True
        Else
            RadioButton3.Checked = False

        End If

        If RadioButton1.Checked = True And CheckBox2.Checked = True Then
            ComboBox1.Visible = True
            ComboBox2.Visible = False
        Else
            ComboBox1.Visible = False

            If CheckBox2.Checked = True Then

                ComboBox2.Visible = True

            End If

        End If


    End Sub

    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged

    End Sub

    Private Sub ToolStripProgressBar1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TableLayoutPanel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs)

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub SalirToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalirToolStripMenuItem1.Click
        'paarar abrir cualquier archivo tipido dwg en el programa..
        Dim obj As Autodesk.AutoCAD.Interop.AcadApplication


        OpenFileDialog1.Filter = (" ACAD (*.dwg)|*.dwg")
        OpenFileDialog1.ShowDialog()
        On Error Resume Next
        obj = New Autodesk.AutoCAD.Interop.AcadApplication

        obj.Application.Documents.Open(OpenFileDialog1.FileName)
        obj.Visible = True

        If Err.Number <> 0 Then
            Err.Clear()
            objautocad.Quit()
        End If


    End Sub



    Private Sub GroupBox2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

        ' se da valor a cada campo del combobox1...
        Dim TRR As Double

        If RadioButton3.Checked = True Then
            KV = 0.208
        ElseIf RadioButton4.Checked = True Then
            KV = 0.24
        Else
            MsgBox("seleccione un voltaje", MsgBoxStyle.Critical)
            Exit Sub
        End If


        Select Case ComboBox1.SelectedIndex
            Case 0
                BP = 125
            Case 1
                BP = 200
            Case 2
                BP = 300
            Case 3
                BP = 400
            Case 4
                BP = 600
            Case 5
                BP = 800
            Case 6
                BP = 1000
            Case 7
                BP = 2500
            Case 8
                BP = 4000
            Case 9
                BP = 8000
            Case 10
                TRR = InputBox("ingrese capacidad")
                BP = TRR / (Math.Sqrt(3) * KV)

        End Select


    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True And CheckBox2.Checked = True Then
            ComboBox2.Visible = True
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged

        'metodo para q se active el trifasico y monofacio (combobox.visibl) si se maraca conjunto con checkbox "principal"

        If RadioButton1.Checked = True And CheckBox2.Checked = True Then
            ComboBox1.Visible = True

        ElseIf RadioButton2.Checked = True And CheckBox2.Checked = True Then
            ComboBox2.Visible = True
        Else
            ComboBox1.Visible = False
            ComboBox2.Visible = False

        End If


    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged

        ' se da valor a cada campo del combobox1...
        If RadioButton3.Checked = True Then
            KV = 0.208
        ElseIf RadioButton4.Checked = True Then
            KV = 0.24
        Else
            MsgBox("seleccione un voltaje", MsgBoxStyle.Critical)
            Exit Sub
        End If

        Dim TR As Double

        Select Case ComboBox2.SelectedIndex
            Case 0
                BP = 60
            Case 1
                BP = 100
            Case 2
                BP = 150
            Case 3
                BP = 200
            Case 4
                BP = 300
            Case 5
                TR = InputBox("ingrese capacidad")
                BP = TR / KV

        End Select

    End Sub

    Private Sub Form3_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDoubleClick

    End Sub

    Private Sub AbrirToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AbrirToolStripMenuItem.Click

        'cierra el espacioen(autocad)
        objautocad.Quit()
        AbrirToolStripMenuItem.Enabled = False
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)



    End Sub

    Private Sub Button2_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        On Error Resume Next
otravez:
        Dim a, b, r As Integer
        a = InputBox("ingrese numero")
        b = InputBox("ingrese numero")

        r = a + b

        If Err.Number <> 0 Then

            MsgBox("error")
            Err.Clear() : GoTo otravez

        End If
        MsgBox(r)

    End Sub

    Private Sub CloseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseToolStripMenuItem.Click
        End
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Label1.Text = TimeOfDay() 'tiempo


    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        Shell("C:\Archivos de programa\iTunes\itunes.exe")
    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged

    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)

    End Sub

    Private Sub Button3_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        'EN ESTE MODULO SE DESARROLLA EL RASTREO DE TODOS LOS TEXT DE LOS CIRCUITOS (NOMBRE DE CIRCUITO ASIGNADO)_
        'SEGUN EL COLOR Q ESTOS PRESENTE: RASTREA LOS ROJOS(1)= ALUMBRADO O ILUMINACION; NARANJA(40) = TOMACORRIENTES;NARANJA MAS OSCURO(43)=220 Y NARANJA MARRON(42)= 220 TRIFASICO;
        ' AL REUNIR TODOS ESTOS NOMBRES SE PROSIGUE A BUSCAR EL NUMERO DE OBJETOS CON REFERENCIA A CADA NOMBRE Q SE HA GUARDADO EN UNOS ARREGLOS POR SEPARADO 
        ' Y SE CALCULA  LAS CARGAS ELECTRICAS QUE CADA CIRCUITO PRESENTA SEGU EL NUMERO DE COMPONENTES QUE ESTE POSEE Y SIENDO (ILUMINACION Y TOMACORRIENTES) MULTIPLICADO POR UNA CONSTANTE VOLTIO AMPERIO

        Dim text As Autodesk.AutoCAD.Interop.Common.AcadText

        On Error Resume Next
        objCAD = GetObject(, "autocad.application")

        If Err.Number <> 0 Then
            Err.Clear()
            objCAD = CreateObject("autocad.application", "")
            objCAD.Visible = True
        End If


        acumuladort = 0
        acumuladori = 0

        Dim filtertype(1) As Short
        Dim filterdata(1) As Object
        i = 1
        'objCAD.ActiveDocument.SelectionSets.Item("new222").Delete()
        For j = 1 To 4

            seleccion = objCAD.ActiveDocument.SelectionSets.Add("crear")
            ' tomas (palabras q van de color naranja)
            If j = 1 Then
                filtertype(0) = 0
                filterdata(0) = "text"
                filtertype(1) = 62
                filterdata(1) = 40 'naranja

                seleccion.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , filtertype, filterdata)

                vecestoma = seleccion.Count

                MsgBox("CANTIDAD DE CIRCUITOS(TOMAS): " & vecestoma)

                For Each text In seleccion
                    'MsgBox(text.TextString)
                    deptoma(i) = text.TextString
                    i = i + 1 'incrementa
                Next

                'instruccion para modificar letras en autocad
                ' For Each text In seleccion
                'text.TextString = "hola"
                'text.Update()
                'Next

            End If
            ' iluminacion (palabras q van de color rojo)
            If j = 2 Then
                filtertype(0) = 0
                filterdata(0) = "text"
                filtertype(1) = 62
                filterdata(1) = 1 'rojo

                seleccion.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , filtertype, filterdata)
                vecesiluminacion = seleccion.Count

                MsgBox("CANTIDAD DE CIRCUITOS(ILUMINACION): " & vecesiluminacion)

                For Each text In seleccion
                    ' MsgBox(text.TextString)
                    depiuminacion(i) = text.TextString
                    i = i + 1 'incrementa
                Next

            End If

            If j = 3 Then
                filtertype(0) = 0
                filterdata(0) = "text"
                filtertype(1) = 62
                filterdata(1) = 43 'rojo

                seleccion.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , filtertype, filterdata)
                veces220 = seleccion.Count

                MsgBox("CANTIDAD DE CIRCUITOS(220): " & veces220)

                For Each text In seleccion
                    ' MsgBox(text.TextString)
                    deptoma220(i) = text.TextString
                    i = i + 1 'incrementa
                Next
            End If

            If j = 4 Then
                filtertype(0) = 0
                filterdata(0) = "text"
                filtertype(1) = 62
                filterdata(1) = 42

                seleccion.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , filtertype, filterdata)
                veces220spe = seleccion.Count

                MsgBox("CANTIDAD DE CIRCUITOS(220) especial: " & veces220spe)

                For Each text In seleccion
                    ' MsgBox(text.TextString)
                    deptoma220spe(i) = text.TextString
                    i = i + 1 'incrementa
                Next
            End If

            objCAD.ActiveDocument.SelectionSets.Item("crear").Delete()
            i = 1 'para el proximo
        Next j
        '///////////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////////////////////////////////////////////////////////////////////////
        ' SUMINISTRO DE TODAS LAS CARGAS (ILUMINACION,TOMACORRIENTES,220 Y TRIFASICOS)

        'condicion si no hay nada y todos los componentes son cero
        ' If (veces220 = 0 And veces220spe = 0 And vecesiluminacion = 0 And vecestoma = 0 And vecestoma) Then
        'MsgBox("NO HAY COMPONENTES PARA HACER CALCULO DE CARGA DE TABLERO", MsgBoxStyle.Critical)
        ' Exit Sub ' se sale del proceso

        ' End If

        '1) ILUMINACION://////////////////////////////////////////////////////////////////////////
        Dim ilumvat : ilumvat = 100 'constante
        'Dim obj As Autodesk.AutoCAD.Interop.Common.AcadEntity
        Dim filtertypEe(0) As Short
        Dim filterdatAa(0) As Object

        For i = 1 To vecesiluminacion

            seleccion = objCAD.ActiveDocument.SelectionSets.Add("cargas_iluminacion")

            filtertypEe(0) = 2 ' codigo para filtrado de objetos 
            filterdatAa(0) = depiuminacion(i) '' nombre del objeto

            seleccion.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , filtertypEe, filterdatAa)
            ' MsgBox("cantidad de objetos seleccionados en " & depiuminacion(i) & " :" & seleccion.Count, MsgBoxStyle.Information)

            If depiuminacion(i).Count = 3 Then
                'fluorecente 3*40
                cargasil(i) = seleccion.Count * 120
            ElseIf depiuminacion(i).Count = 1 Then
                'fluorecente 2*40
                cargasil(i) = seleccion.Count * 80
            ElseIf depiuminacion(i).Count = 4 Then
                'fluorecente 4*40
                cargasil(i) = seleccion.Count * 160
            Else
                cargasil(i) = seleccion.Count * ilumvat
            End If

            n_cargasil(i) = depiuminacion(i)
            c_cargail(i) = 0

            acumuladori += seleccion.Count
            'For Each obj In seleccion
            'obj.color = Autodesk.AutoCAD.Interop.Common.ACAD_COLOR.acBlue
            'obj.Update()
            'Next
            objCAD.ActiveDocument.SelectionSets.Item("cargas_iluminacion").Delete()
        Next i

        '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        '2) TOMACORRIENTES:

        Dim vatio : vatio = 200
        ' Dim obj As Autodesk.AutoCAD.Interop.Common.AcadEntity
        Dim filtertypeW(0) As Short
        Dim filterdataW(0) As Object

        For i = 1 To vecestoma  ' CANTIDAD DE CIRCUITOS QUE HAY EN EL DIBUJO( MENOS UNO PORQ UTILIZO EL ESPACIO (0) EN EL VECTOR Q ATRAPA LOS NOMBRES)

            seleccion = objCAD.ActiveDocument.SelectionSets.Add("cargas_tomacorrientes")

            filtertypeW(0) = 2
            filterdataW(0) = deptoma(i)

            seleccion.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , filtertypeW, filterdataW)
            ' MsgBox("cantidad de objetos seleccionados en " & deptoma(i) & " :" & seleccion.Count, MsgBoxStyle.Information)

            If seleccion.Count = 1 Then ' cuadno son cargas 110 especiales !!
                n_carga(i) = deptoma(i)
                cargas(i) = InputBox("Ingrese carga de: " & deptoma(i), "TOMACORRIENTES")
                c_cargat(i) = 4
                acumuladort = seleccion.Count + acumuladort
            Else
                '//seria mi estructura de datos
                n_carga(i) = deptoma(i)
                cargas(i) = seleccion.Count * vatio
                c_cargat(i) = 1
                acumuladort = seleccion.Count + acumuladort
            End If
            'For Each obj In seleccion
            'obj.color = Autodesk.AutoCAD.Interop.Common.ACAD_COLOR.acBlue
            'obj.Update()
            'Next
            objCAD.ActiveDocument.SelectionSets.Item("cargas_tomacorrientes").Delete()

        Next i


        ' For i = 1 To vecestoma
        'MsgBox("Cargas: " & cargas(i) & " de " & n_carga(i))
        '  Next
        '////////////////////////////////////////////////////////////////////////////////////////

        'tomacorrientes 220:

        ' Dim obj As Autodesk.AutoCAD.Interop.Common.AcadEntity
        Dim filtertypea(0) As Short
        Dim filterdatar(0) As Object

        For i = 1 To veces220   ' CANTIDAD DE CIRCUITOS QUE HAY EN EL DIBUJO( MENOS UNO PORQ UTILIZO EL ESPACIO (0) EN EL VECTOR Q ATRAPA LOS NOMBRES)

            seleccion = objCAD.ActiveDocument.SelectionSets.Add("220")

            filtertypea(0) = 2
            filterdatar(0) = deptoma220(i)

            seleccion.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , filtertypea, filterdatar)
            ' MsgBox("cantidad de objetos seleccionados en " & deptoma220(i) & " :" & seleccion.Count, MsgBoxStyle.Information)
            '//seria mi estructura de datos

            carga220(i) = InputBox("Indique la carga de: " & deptoma220(i), "CARGAS 220")
            n_carga220(i) = deptoma220(i)
            c_carga220(i) = 2

            acumulador220 = seleccion.Count + acumulador220

            'For Each obj In seleccion
            'obj.color = Autodesk.AutoCAD.Interop.Common.ACAD_COLOR.acBlue
            'obj.Update()
            'Next
            objCAD.ActiveDocument.SelectionSets.Item("220").Delete()
        Next i


        ' For i = 1 To vecestoma
        'MsgBox("Cargas: " & cargas(i) & " de " & n_carga(i))
        '  Next
        '/////////////////////////////////////////////////////////////////////////////////////////

        '220 TRIFASICO:
        ' Dim obj As Autodesk.AutoCAD.Interop.Common.AcadEntity
        Dim filtertypeY(0) As Short
        Dim filterdataY(0) As Object

        For i = 1 To veces220spe   ' CANTIDAD DE CIRCUITOS QUE HAY EN EL DIBUJO( MENOS UNO PORQ UTILIZO EL ESPACIO (0) EN EL VECTOR Q ATRAPA LOS NOMBRES)

            seleccion = objCAD.ActiveDocument.SelectionSets.Add("220spe")

            filtertypeY(0) = 2
            filterdataY(0) = deptoma220spe(i)

            seleccion.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , filtertypeY, filterdataY)
            ' MsgBox("cantidad de objetos seleccionados en " & deptoma220(i) & " :" & seleccion.Count, MsgBoxStyle.Information)
            '//seria mi estructura de datos

            carga220spe(i) = InputBox("Indique la carga de: " & deptoma220spe(i), "CARGAS 220 TRIFASICO")
            n_carga220spe(i) = deptoma220spe(i)
            c_carga220spe(i) = 3

            acumulador220spe += seleccion.Count  ' necesita poner el mismo nombre alos componetes """OJOOOO"corregir

            'For Each obj In seleccion
            'obj.color = Autodesk.AutoCAD.Interop.Common.ACAD_COLOR.acBlue
            'obj.Update()
            'Next
            objCAD.ActiveDocument.SelectionSets.Item("220spe").Delete()
        Next i



        ' For i = 1 To vecestoma
        'MsgBox("Cargas: " & cargas(i) & " de " & n_carga(i))
        '  Next
        '//////////////////////////////////////////////////////////////////////////////

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ''NOTA: hacerlo personalizado para toma

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'NOTA: hacerlo personalizado para toma


    End Sub

    Sub ordenar_val_equilibrio_monofasico(ByRef valores_carg() As Integer, ByRef nombre_cargas() As String, ByRef codigo_cargas() As Integer, ByRef veces As Integer)
        'subrutina utilizando byref el cual modifica los valores originales q le pasan informacion a los parametros 
        Dim m As Integer, mmm As Integer : Dim mm As String

        For j = 1 To veces ' ordena mi estructura de datos
            m = valores_carg(j) 'carga
            mm = nombre_cargas(j) 'nombre
            mmm = codigo_cargas(j) 'codigo
            For i = 1 + j To veces
                If valores_carg(j) > valores_carg(i) Then
                    valores_carg(j) = valores_carg(i) : nombre_cargas(j) = nombre_cargas(i) : codigo_cargas(j) = codigo_cargas(i)
                    valores_carg(i) = m : nombre_cargas(i) = mm : codigo_cargas(i) = mmm
                    m = valores_carg(j) : mm = nombre_cargas(j) : mmm = codigo_cargas(j)
                End If
            Next
        Next
    End Sub
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim text As Autodesk.AutoCAD.Interop.Common.AcadText

        'equilibrio de trifasico trifasico
        If RadioButton1.Checked = True Then

            Dim aux2 As Integer = 1
            Dim vector220T(100) As Integer : Dim n_vector220T(100) As String : Dim c_vector220T(100)
            Dim aux As Integer = 1
            Dim cantidad_total_ilu_toma As Integer

            Do While cargas(l) <> 0
                l = l + 1
            Loop
            l = l - 1
            For ix = 1 To l
                valorres(aux2) = cargas(ix)
                n_valorres(aux2) = n_carga(ix)
                c_valorres(aux2) = c_cargat(ix)
                aux2 = aux2 + 1
            Next
            '//////////////////////////
            Do While cargasil(ll) <> 0
                ll = ll + 1
            Loop
            ll = ll - 1

            For ix = 1 To ll
                valorres(aux2) = cargasil(ix)
                n_valorres(aux2) = n_cargasil(ix)
                c_valorres(aux2) = c_cargail(ix)
                aux2 = aux2 + 1
            Next
            '//////////////////////////
            l = ll + l

            ordenar_val_equilibrio_monofasico(valorres, n_valorres, c_valorres, l)
            'ordenar_val_equilibrio_monofasico(carga220, n_carga220, c_carga220, veces220)

            '////////////////////////PARTE DEL CODIGO DONDE SE ENCARGA DE APROXIMAR LOS VALORES 110 CON LOS 220
            Dim aux_c, aux_b, aux_a, hijo_puta : aux_b = 1 : aux_c = 1 : aux_a = 1 : hijo_puta = 0
            Dim avisame As Integer = 0 : Dim avisado As Integer = 1

            Do While aux <= l + veces220

                For h = 1 To 2
                    If carga220(aux_c) = 0 Then
                        Exit For
                    End If
                    v_definitivo(aux) = carga220(aux_c)
                    v_definombre(aux) = n_carga220(aux_c)
                    v_deficode(aux) = c_carga220(aux_c)
                    '   MsgBox(aux)
                    maldito(valorres, n_valorres, c_valorres, aux_c) ' busca la aproximacion de las cargas 110 con respecto a las 220
                    aux += 1
                    ' MsgBox(carga220(aux_c))
                    aux_c += 1
                    avisame += 1 ' esta variable es util cuando  el numero de cargas 220 es impar
                Next

                ' MsgBox(" si paso por aki y hay un error")
                ' como es impar las cargasa avisame =1 coloco que avise sea =2 para
                '///////////////////////////////que coloque la carga 110 a ese unico 220 que queda y luego pase a ordenar el resto de las cargas 220 para mayor equilibrio.
                If avisame = 1 Then
                    ' esta parte del codigome evita que en mi algoritmo de asignamiento del tablero halla una confucion a la hora de ser las cant 220 IMPAR!!

                    Dim retenedor, n_retenedor, c_retenedor
                    retenedor = valorres(aux_b)
                    n_retenedor = n_valorres(aux_b)
                    c_retenedor = c_valorres(aux_b)

                    valorres(aux_b) = valorres(aux_b + 2)
                    n_valorres(aux_b) = n_valorres(aux_b + 2)
                    c_valorres(aux_b) = c_valorres(aux_b + 2)

                    valorres(aux_b + 2) = retenedor
                    n_valorres(aux_b + 2) = n_retenedor
                    c_valorres(aux_b + 2) = c_retenedor

                    avisado = 2
                    For ix = 1 To 2
                        If valorres(aux_b) = 0 Then
                            Exit For
                        End If
                        ' MsgBox("pase por aki 2")
                        v_definitivo(aux) = valorres(aux_b)
                        v_definombre(aux) = n_valorres(aux_b)
                        v_deficode(aux) = c_valorres(aux_b)
                        aux += 1
                        'MsgBox(valorres(aux_b))
                        aux_b += 1
                    Next
                    For ix = 1 To 2 ' me va eliminando las casillas q no necesito
                        valorres(aux_a) = Nothing
                        aux_a += 1
                    Next
                End If

                For ix = avisado To 2
                    If valorres(aux_b) = 0 Then
                        Exit For
                    End If
                    ' MsgBox("pase por aki 2")
                    v_definitivo(aux) = valorres(aux_b)
                    v_definombre(aux) = n_valorres(aux_b)
                    v_deficode(aux) = c_valorres(aux_b)
                    aux += 1
                    'MsgBox(valorres(aux_b))
                    aux_b += 1
                    hijo_puta = 1
                Next

                If hijo_puta = 1 Then
                    For ix = avisado To 2 ' me va eliminando las casillas q no necesito
                        valorres(aux_a) = Nothing
                        aux_a += 1
                    Next
                    hijo_puta = 0
                End If

                '// ordena los 110 restantes nuevamente y hace el equilibrio.. a continuacion
                If avisado = 2 Then
                    ordenar_val_equilibrio_monofasico(valorres, n_valorres, c_valorres, l)
                    ' MsgBox(" ordene el resto jejej va funcionar ya veras")
                    avisado = 1
                End If

                avisame = 0 ' apago el avisame 
            Loop
            'MsgBox("alfin salii de ese lugar :S")
            ' MsgBox(aux)

            '////////////////////////////////////////////////////////////////////////////////////////

            '////////// ASIGNACION DE CARGAS EN EL MODELSPACE(PLANO) VALORES (220) Y (110) *VALIDO PARA TRIFASICO ESTE ALGORITMO
            'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
            Dim filtertype(1) As Short
            Dim filterdata(1) As Object
            Dim bandera As Integer = 1
            Dim aux_posicion As Integer = 1  ' me indica por cual posicion del vector definitivo el esta
            Dim si_paso As Integer = 0 ' variable q me inca cuantas veces pasa por la rutina de 220 en asignacion
            Dim repetir As Integer = 0
            Dim paso_akimaldio As Integer = 0


repetir_proceso:
            'seleccion 220  
            MsgBox("veces220: " & l + veces220)
            For ix = 1 To 2
                seleccion = objCAD.ActiveDocument.SelectionSets.Add("220color")

                filtertype(0) = 62
                filterdata(0) = 43
                filtertype(1) = 0
                filterdata(1) = "text"

                seleccion.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , filtertype, filterdata)

                For Each text In seleccion
                    If text.TextString = v_definombre(aux_posicion) Then
                        ' MsgBox("pase por aki lol")
                        text.TextString = "C" & bandera & "-" & bandera + 2
                        text.Update()
                        bandera += 1
                        aux_posicion += 1
                        si_paso += 1
                    End If
                Next
                objCAD.ActiveDocument.SelectionSets.Item("220color").Delete()
                '  MsgBox(" sali del foreach")
            Next

            ' MsgBox(" aux : " & aux_posicion)
            If si_paso = 2 Then bandera += 2 ' para queda en la posicion dep de haber puesto dos valores en tablero(220)

            For ix = 1 To 2
                'seleccion (110)
                Dim filtertyp(0) As Short
                Dim filterdat(0) As Object

                seleccion = objCAD.ActiveDocument.SelectionSets.Add("110color2")

                filtertyp(0) = 0
                filterdat(0) = "text"

                seleccion.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , filtertyp, filterdat)

                For Each text In seleccion
                    If text.TextString = v_definombre(aux_posicion) Then
                        '  MsgBox("valor a modificar " & v_definombre(aux_posicion) & "and" & text.TextString)
                        text.TextString = "C" & bandera ' 
                        text.Update()
                        ' MsgBox("el valor modificado: " & text.TextString)
                        aux_posicion += 1
                        bandera += 1
                        paso_akimaldio += 1
                        'MsgBox("ando por estos lados")
                        If si_paso = 1 Then ' caso especial de impar
                            bandera += 1 ' se agrega otro
                            si_paso = 0 ' se termina y se coloca el resto de 110 normales
                        End If
                    End If
                    If paso_akimaldio = 2 Then Exit For
                Next

                objCAD.ActiveDocument.SelectionSets.Item("110color2").Delete()
                '  MsgBox("sali del segundo maldito foreach")

                If paso_akimaldio = 2 Then
                    paso_akimaldio = 0
                    Exit For ' termina y sale
                End If
            Next
            si_paso = 0

            If aux_posicion > l + veces220 Then GoTo fin_proceso

            GoTo repetir_proceso
Fin_proceso:

            'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA

            '  MsgBox(aux_posicion) '77 en elplano da 22
            '   MsgBox(bandera) '// 27
            '///////////////////////////////////////////////////////////////////////////////////////

            'ASIGNACION PARA 220 ESPECIALES

            seleccion = objCAD.ActiveDocument.SelectionSets.Add("220scolor")

            filtertype(0) = 62
            filterdata(0) = 42
            filtertype(1) = 0
            filterdata(1) = "text"

            seleccion.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , filtertype, filterdata)

            Dim aux_definitvo As Integer = aux_posicion
            Dim bandera_avisa As Integer = 1

            ' NOTA IMPORTANTE: en esta parte del codigo el asigna 220 a lo q resta del vector difinitivo !! este misma seleccion que se hace al principio del programa
            ' es la misma seleccion y tiene la mismas posiciones por lo tanto se logra la equivalencia sin identificar cual es el tipo de texto

            For ix = 1 To seleccion.Count ' aki asigno los 220 especiales 
                v_definitivo(aux_definitvo) = carga220spe(ix)
                v_definombre(aux_definitvo) = n_carga220spe(ix)
                v_deficode(aux_definitvo) = c_carga220spe(ix)
                aux_definitvo += 1
            Next

            For ix = aux_definitvo To 100
                v_deficode(ix) = 5
            Next

            If bandera Mod 2 <> 0 Then

                For Each text In seleccion ' esta seleccion tiene equivalencia ya con los datos guardados anteriormente

                    text.TextString = "C" & bandera & "," & bandera + 2 & "," & bandera + 4
                    text.Update()
                    bandera += 1

                    If bandera_avisa Mod 2 = 0 Then
                        bandera += 4
                    End If
                    bandera_avisa += 1

                Next
                objCAD.ActiveDocument.SelectionSets.Item("220scolor").Delete()
            Else
                For Each text In seleccion

                    text.TextString = "C" & bandera & "," & bandera + 2 & "," & bandera + 4
                    text.Update()

                    bandera += 1

                    If bandera_avisa Mod 2 = 0 Then
                        bandera += 4
                    End If
                    bandera_avisa += 1

                Next
                objCAD.ActiveDocument.SelectionSets.Item("220scolor").Delete()
            End If
            '///////////////////////////////////////////////////////////////////////////////////

            MsgBox(bandera) ' SON 32 Y DA 33 POR EL CONTADOR 

            If bandera Mod 2 = 0 Then ' esto forza a que si el numero es par entonces lo haga par y me deje RESERVA en el tablero
                bandera += 6
            Else
                bandera += 5
            End If

            cantidad_circuito_automatico = bandera
        End If ' fin de equilibrio para trifasico
        '///////////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////////////////
        '//equilibrio para monofasico
        If RadioButton2.Checked = True Then
            Dim bandera, toma_bandera, ilu_bandera : bandera = 1 : toma_bandera = 1 : ilu_bandera = 1

            Dim aux2 : aux2 = 1

            ' cantidad total de cargas!! importante

            Do While cargas(l) <> 0
                l = l + 1
            Loop
            l = l - 1
            For ix = 1 To l
                valorres(aux2) = cargas(ix)
                n_valorres(aux2) = n_carga(ix)
                c_valorres(aux2) = c_cargat(ix)
                aux2 = aux2 + 1
            Next
            '//////////////////////////
            Do While cargasil(ll) <> 0
                ll = ll + 1
            Loop
            ll = ll - 1

            For ix = 1 To ll
                valorres(aux2) = cargasil(ix)
                n_valorres(aux2) = n_cargasil(ix)
                c_valorres(aux2) = c_cargail(ix)
                aux2 = aux2 + 1
            Next
            '//////////////////////////
            l = ll + l


            ' MsgBox(l) ' cantidad total de cargas

            'metodo de ordenamiento por seleccion(visual basic)
            ordenar_val_equilibrio_monofasico(valorres, n_valorres, c_valorres, l) ' arreglos de datos

            Dim o, oo, totalA, totalB, casoimpar
            o = 1 : totalA = 0 : casoimpar = 0
            oo = 1 : totalB = 0

            '  sub-trex : me aviza si el total es par o impar ....
            ' en el caso de ser impar hago una excepcion
            If l Mod 2 <> 0 Then
                casoimpar = 1
            End If
            '////////////////////////////////////////// se establece un valores vector a y vector b
            If casoimpar = 1 Then ' en el caso de que las cargas introducidas sean impar 
                For ix = 2 To l ' en este caso el primer dato que es el menor para luego sumarselo al menor valor de  A O B
                    If ix Mod 2 = 0 Then
                        vectorA(o) = valorres(ix)
                        n_vectorA(o) = n_valorres(ix)
                        c_vectorA(o) = c_valorres(ix)
                        o = o + 1
                    Else
                        vectorB(oo) = valorres(ix)
                        n_vectorB(oo) = n_valorres(ix)
                        c_vectorB(oo) = c_valorres(ix)
                        oo = oo + 1
                    End If
                Next
            Else
                For ix = 1 To l  ' en el caso de que la cantidad de cargas sean par
                    If ix Mod 2 = 0 Then
                        vectorA(o) = valorres(ix)
                        n_vectorA(o) = n_valorres(ix)
                        c_vectorA(o) = c_valorres(ix)
                        o = o + 1
                    Else
                        vectorB(oo) = valorres(ix)
                        n_vectorB(oo) = n_valorres(ix)
                        c_vectorB(oo) = c_valorres(ix)
                        oo = oo + 1
                    End If
                Next
            End If
            '//////////////////////////////////////////aki se hace la suma de las columnas

            Dim aux, k, kkx, porcentaje, n_aux, c_aux
            porcentaje = 2
denuevo:
            k = l / 2 ' CANTIDAD DE CARGAS EN  AMBOS VECTORES
            kkx = k

            For j = 1 To kkx '// inicio del : ALGORITMO DE EQUILIBRIO!!

                If j = 1 Then
                    aux = vectorA(k) : n_aux = n_vectorA(k) : c_aux = c_vectorA(k)    ' cambia la posicion del final( el de mayor peso)
                    vectorA(k) = vectorB(k) : n_vectorA(k) = n_vectorB(k) : c_vectorA(k) = c_vectorB(k)
                    vectorB(k) = aux : n_vectorB(k) = n_aux : c_vectorB(k) = c_aux
                ElseIf j = 2 Then
                    k = 1           ' luego comienza desde el inicio a cambiar
                    aux = vectorA(k) : n_aux = n_vectorA(k) : c_aux = c_vectorA(k)    ' cambia la posicion del final( el de mayor peso)
                    vectorA(k) = vectorB(k) : n_vectorA(k) = n_vectorB(k) : c_vectorA(k) = c_vectorB(k)
                    vectorB(k) = aux : n_vectorB(k) = n_aux : c_vectorB(k) = c_aux
                Else                    'luego sigue la siguiente fila y se cambia su posicion
                    k = k + 1
                    aux = vectorA(k) : n_aux = n_vectorA(k) : c_aux = c_vectorA(k)    ' cambia la posicion del final( el de mayor peso)
                    vectorA(k) = vectorB(k) : n_vectorA(k) = n_vectorB(k) : c_vectorA(k) = c_vectorB(k)
                    vectorB(k) = aux : n_vectorB(k) = n_aux : c_vectorB(k) = c_aux
                End If

                For ix = 1 To o - 1
                    totalA = vectorA(ix) + totalA
                Next

                For ix = 1 To oo - 1
                    totalB = vectorB(ix) + totalB
                Next

                ' Esto ocurre si la cantidad de cargas es impar ( almenor valor de la sumatoria se le asigna el primer valor)
                If casoimpar = 1 Then
                    If totalA <= totalB Then
                        totalA = totalA + valorres(1)
                        vectorA(kkx + 1) = valorres(1) '// estar pendiente con esto asignado (impar)
                        n_vectorA(kkx + 1) = n_valorres(1)
                        c_vectorA(kkx + 1) = c_valorres(1)
                    Else
                        totalB = totalB + valorres(1)
                        vectorB(kkx + 1) = valorres(1)
                        n_vectorB(kkx + 1) = n_valorres(1)
                        c_vectorB(kkx + 1) = c_valorres(1)
                    End If
                End If
                '//////////////////////////////////////////

                ' For ix = 1 To kkx
                'MsgBox("carga: " & vectorA(ix) & "nombre: " & n_vectorA(ix) & "codifo: " & c_vectorA(ix))
                'Next
                'For ix = 1 To kkx
                'MsgBox("carga: " & vectorB(ix) & "nombre: " & n_vectorB(ix) & "codifo: " & c_vectorB(ix))
                ' Next

                ' SI EL RESUALTADO ES EL OPTIMO entonces :
                If porcentaje = 2 And Math.Abs(((totalA - totalB) / (totalA + totalB))) * 100 <= 2 Then
                    Exit For
                ElseIf porcentaje = 3 And Math.Abs((totalA - totalB) / (totalA + totalB)) * 100 <= 3 Then
                    Exit For
                ElseIf porcentaje = 4 And Math.Abs((totalA - totalB) / (totalA + totalB)) * 100 <= 4 Then
                    Exit For
                ElseIf porcentaje = 5 And Math.Abs((totalA - totalB) / (totalA + totalB)) * 100 <= 5 Then
                    Exit For
                ElseIf porcentaje = 6 And Math.Abs((totalA - totalB) / (totalA + totalB)) * 100 <= 6 Then
                    Exit For
                ElseIf porcentaje = 7 And Math.Abs((totalA - totalB) / (totalA + totalB)) * 100 <= 7 Then
                    Exit For
                ElseIf porcentaje = 8 And Math.Abs((totalA - totalB) / (totalA + totalB)) * 100 <= 8 Then
                    Exit For
                ElseIf porcentaje = 9 And Math.Abs((totalA - totalB) / (totalA + totalB)) * 100 <= 9 Then
                    Exit For
                ElseIf porcentaje = 10 And Math.Abs((totalA - totalB) / (totalA + totalB)) * 100 <= 10 Then
                    Exit For
                ElseIf porcentaje = 11 Then
                    totalA = 0 : totalB = 0 : vectorA(kkx + 1) = 0 : vectorB(kkx + 1) = 0
                    GoTo terminado
                End If

                If j > 1 Then 'si el cambi no resulta o entra dentro del porcentaje se regresa a sus posicion normal y sigue con la otra fila
                    aux = vectorA(k) : n_aux = n_vectorA(k) : c_aux = c_vectorA(k)    ' cambia la posicion del final( el de mayor peso)
                    vectorA(k) = vectorB(k) : n_vectorA(k) = n_vectorB(k) : c_vectorA(k) = c_vectorB(k)
                    vectorB(k) = aux : n_vectorB(k) = n_aux : c_vectorB(k) = c_aux
                End If

                totalA = 0 : totalB = 0 : vectorA(kkx + 1) = 0 : vectorB(kkx + 1) = 0

            Next
            '//////////////////////////////////////////

            '   MsgBox("Total de A: " & totalA & " y Total B: " & totalB) ' respuesta final

            If totalA = 0 And totalB = 0 Then
                porcentaje += 1
                '    MsgBox("Los valores de equilibrio no entran en el porcentaje: " & porcentaje - 1 & " se le asignara un porcentaje de: " & porcentaje, MsgBoxStyle.Critical)
                GoTo denuevo
            End If

terminado:
            If totalA = 0 And totalB = 0 Then
                '  MsgBox("No se pudo encontrar equilibrio", MsgBoxStyle.Information)
                Exit Sub
                'dejar aki un (exit sub) 
            End If


            '/////////////////////////////
            'ACTUALIZAR NOMBRE DE CARGAS EN EL DIBUJO: ALGORITMOS DE ASIGNACION

            '// NOTA HACER TRABAJAR ESTE ALGORITMO CON MAS PRUEBAS Y CON CARGAS IMPARES !!!

            Dim filtertype(0) As Short
            Dim filterdata(0) As Object
segundo:
            'MsgBox("primera parte")
            seleccion = objCAD.ActiveDocument.SelectionSets.Add("cambio212")
            filtertype(0) = 0
            filterdata(0) = "text"
            seleccion.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , filtertype, filterdata)

            For Each text In seleccion
                'MsgBox(text.TextString & "=" & n_vectorA(toma_bandera))
                If text.TextString = n_vectorA(toma_bandera) Then

                    '' MsgBox(n_vectorA(toma_bandera) & "y bandera: " & bandera)
                    text.TextString = "C" & bandera
                    text.Update()
                    v_definitivo(bandera) = vectorA(toma_bandera)
                    v_deficode(bandera) = c_vectorA(toma_bandera)
                    v_definombre(bandera) = n_vectorA(toma_bandera)
                    toma_bandera += 1 ' para movilizar toma bandera
                    bandera += 1

                    If bandera Mod 2 = 0 Then
                        objCAD.ActiveDocument.SelectionSets.Item("cambio212").Delete()
                        GoTo segundo
                    Else
                        objCAD.ActiveDocument.SelectionSets.Item("cambio212").Delete()
                        'MsgBox("pase aki")
                        GoTo segundo2
                    End If
                End If
            Next
            objCAD.ActiveDocument.SelectionSets.Item("cambio212").Delete()
            '////////////////////////////////////////////////////////////
segundo2:
            'MsgBox("paso a la segunda parte")
            seleccion = objCAD.ActiveDocument.SelectionSets.Add("cambio2333332")
            filtertype(0) = 0
            filterdata(0) = "text"
            seleccion.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , filtertype, filterdata)

            For Each text In seleccion
                If text.TextString = n_vectorB(ilu_bandera) Then
                    ' MsgBox(n_vectorB(ilu_bandera) & "y bandera: " & bandera)
                    text.TextString = "C" & bandera
                    text.Update()
                    v_definitivo(bandera) = vectorB(ilu_bandera)
                    v_deficode(bandera) = c_vectorB(ilu_bandera)
                    v_definombre(bandera) = n_vectorB(ilu_bandera)
                    ilu_bandera += 1
                    bandera += 1

                    If bandera Mod 2 = 0 Then
                        objCAD.ActiveDocument.SelectionSets.Item("cambio2333332").Delete()
                        GoTo segundo2
                    Else
                        objCAD.ActiveDocument.SelectionSets.Item("cambio2333332").Delete()
                        GoTo segundo
                    End If
                End If
            Next
            objCAD.ActiveDocument.SelectionSets.Item("cambio2333332").Delete()

            If bandera - 1 = l Then '//' final de este algoritmo de asignamiento de cargas
                ' MsgBox("Asignacion de cargas completa")
                'Exit Sub
            End If

            'For i = 1 To l
            'MsgBox("carga: " & v_definitivo(i) & "codeigo: " & v_deficode(i))
            'Next

            '////////////PARA AASIGNAR CODIGO A LOS TOMAS 220

            Dim filtertyp(1) As Short
            Dim filterdat(1) As Object
            Dim bandera220 : bandera220 = 1

            seleccion = objCAD.ActiveDocument.SelectionSets.Add("cambio220")

            filtertyp(0) = 62
            filterdat(0) = 43
            filtertyp(1) = 0
            filterdat(1) = "text"

            seleccion.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , filtertyp, filterdat)
            MsgBox(seleccion.Count)
            '///////////////////////////////OJO OJO OJO OJO 
            Dim defini_bandera As Integer = bandera

            For ix = 1 To seleccion.Count
                v_definitivo(defini_bandera) = carga220(ix) ' nota importante: en esta parte del codigo el asigna 220 a lo q resta del vector difinitivo !! este misma seleccion que se hace al principio del programa
                v_deficode(defini_bandera) = c_carga220(ix) ' es la misma seleccion y tiene la mismas posiciones por lo tanto se logra la equivalencia sin identificar cual es el tipo de texto
                v_definombre(defini_bandera) = n_carga220(ix)
                defini_bandera += 1
            Next

            For ix = defini_bandera To 100 ' con elfin de volver los demas reserva
                v_deficode(defini_bandera) = 5
                defini_bandera += 1
            Next
            '//////////////////////////////////////////////////
            If bandera Mod 2 <> 0 Then ' en este punto de evalua el valor de la variable bandera para ver por donde va el valor(numerico) dps del equilibrio y colocar los 220 monofasicos
                For Each text In seleccion
                    text.TextString = "C" & (bandera) & "," & (bandera + 2)
                    text.Update()

                    bandera = bandera + 1

                    If bandera220 Mod 2 = 0 Then
                        bandera = bandera + 2
                    End If
                    bandera220 = bandera220 + 1
                Next
            Else
                For Each text In seleccion

                    text.TextString = "C" & (bandera) & "," & (bandera + 2)
                    text.Update()
                    bandera = bandera + 1

                    If bandera220 Mod 2 = 0 Then
                        bandera += 2
                    End If
                    bandera220 = bandera220 + 1
                Next
            End If

            objCAD.ActiveDocument.SelectionSets.Item("cambio220").Delete()
            ' MsgBox("Asignacion de cargas completa")

            cantidad_circuito_automatico = bandera + 2 ' cantidad de circuitos con 220
        End If '77fin de equilibrio para monofasico





        '*********************************************************************************
        'DECLARACION DE VARIABLES PRINCIPALES DEL CALCULO DE TABLERO Y DIBUJADO EN AUTOCAD*
        '********************************************************************************************************************************************************************************************

        Dim line As Autodesk.AutoCAD.Interop.Common.AcadLine 'declaro la linea
        Dim PtoIn(2) As Double 'declaro punto inicio x,y,z
        Dim PtoFin(2) As Double 'declaro puntofinal x,y,z

        Dim insPoint(2) As Double 'Declare insertion point
        Dim textHeight As Double       'Declare text height
        Dim textStr As String         'Declare text string

        Dim arco As Autodesk.AutoCAD.Interop.Common.AcadArc ' declaro el arco
        Dim centro(2) As Double 'declaro centro en x,y,z
        Dim radio, anginic, angfinal As Double ' decalro radio , angulo inicial y final  del arco

        Dim circulo As Autodesk.AutoCAD.Interop.Common.AcadCircle
        Dim centroo(2) As Double 'declaro centro en x,y,z
        Dim radioo As Double 'declaro radio del circulo

        Dim t, color, g, kk, i, lon, tt1, tt2, ttt1, ttt2 As Double
        Dim l1, l2 As Integer
        Dim ncircuitos, ncircuitos2 As Double 'numero de circuitos

        Dim nombre 'nombre del tablero's 

        '*********************************************************************************************+++************************************************************
        'arbrimos hoja de calculo en excel si se ha esojido !! en checkbox3!!
        '*****************************************************************
        If CheckBox3.Checked = True Then
            'direccion del archivo..
            uu = "C:\estudio de carga .xlsx" ' se coloca la direccion del documento excel q se utiliza...
            'se abre excel
            exe = New Microsoft.Office.Interop.Excel.Application
            ' se abre un espacio de trabajo
            exe.Workbooks.Open(uu)
            'este es visible
            exe.Visible = True
            exe.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMinimized
        End If
        '***********************************************************************************************************************************************************
        'INICIO DEL PROCESO
        '********************************************************************************************************************
        Dim si
        'parte del programa q ve si existe esppacio para dibujar si no existe da la opcion de crear uno ahora mismo
        On Error Resume Next

        objautocad = GetObject(, "AutoCAD.Application")

        If Err.Number <> 0 Then
            Err.Clear()
            si = MsgBox("no existe espacio para dibujar desea crear uno ahora mismo?", MsgBoxStyle.Information + MsgBoxStyle.YesNo)

            If si = vbYes Then
                objautocad = CreateObject("autoCAD.application", "")
                'instruccion q se utiliza para minimizar el autocad  cuando inicie
                objautocad.WindowState = Autodesk.AutoCAD.Interop.Common.AcWindowState.acMin
                'objautocad.Visible = False hace q aparesca el autocada pero en modo de  q no s evea pero igual en el se trabaja

                'para q el close se active del menu de herrmientas...
                encender = 1
                If encender = 1 Then
                    AbrirToolStripMenuItem.Enabled = True
                End If

            Else
                Exit Sub ' funcion q actua como el break..

            End If
        End If

        '*******************************************************************************************************************
        'PARTE DEL CODIGO EN DONDE SE PREGUNTA: FACTOR DE POTENCIA(FP), NUMERO DE CIRCUITOS(ncircuitos) Y NOMBRE DEL TABLERO*
        '***********************************************************************************************************************************************************
        If ((RadioButton1.Checked Or RadioButton2.Checked) = False) Then
            t = MsgBox(" no ha seleccinado el tipo de tablero", vbCritical, "tableros")

        Else

            Dim fp As Double   'variable q contiene factor de potencia 

            On Error Resume Next
oo:
            fp = InputBox("agrege factor de potencia(FP)", "FACTOR DE POTENCIA")

            If Err.Number <> 0 Then 'corrector de errores
                Err.Clear()
                MsgBox("error ha introducido un  valor caracter", MsgBoxStyle.Critical)
                GoTo oo
            End If

            On Error Resume Next
o2:         'reparar esto(OJO 2.0)
            '//////////
            If RadioButton2.Checked = True Then
                If cantidad_circuito_automatico Mod 2 <> 0 Then
                    cantidad_circuito_automatico += 1
                End If
                ncircuitos = cantidad_circuito_automatico  ' cantidad de circuitos
            ElseIf RadioButton1.Checked = True Then
                ncircuitos = cantidad_circuito_automatico  ' cantidad de circuitos
            End If
            '/////////////

            If Err.Number <> 0 Then
                Err.Clear()
                MsgBox("error ha introducido un valor caracter", MsgBoxStyle.Critical)
                GoTo o2
            End If

            ncircuitos2 = ncircuitos


            lon = 87.5 * ncircuitos ' me indica la longuitud vertical de las barras monofasico y trifasico
            'donde 87.5 es el valor promedio para crearlas ( valor aproximado )
            ncircuitos = ncircuitos / 2 ' indica elnumero de barras neutro utilizadas

            nombre = InputBox("introduzca nombre de tablero a realizar")
            'excel escribe en campo el titulo del tablero..
            exe.Cells(2, 1) = UCase(nombre)
            ' color = InputBox("ingrese color")
            Dim ubicacion As String
            ubicacion = InputBox("ingrese la ubicacion de tablero")
            exe.Cells(2, 8) = UCase(ubicacion)


            If CheckBox1.Checked = True Then
                nombre = UCase(nombre) 'instruccion importante(vuelve mayuscula los string (cadenas de caractere)
            Else
            End If

            insPoint(0) = -300     'Set insertion point x coordinate
            insPoint(1) = (lon + 200)   'Set insertion point y coordinate
            insPoint(2) = 0      'Set insertion point z coordinate
            textHeight = 20                'Set text height to 1.0
            textStr = nombre       'Set the text string
            'Create Text object
            text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
            text.color = 2
            '**********************************************************************************************************************************************************

            '**************************************************************************************************************************************
            'PARTE DEL PROCESO EN EL CUAL SEGUN EL BOTON ELEGIDO (MONOFASICO O TRIFASICO) ESTE DIBUJA DE CUANTAS BARRAS VERTICALES SERA EL TABLERO*
            '**************************************************************************************************************************************


            If RadioButton1.Checked = True Then

                exe.Cells(3, 4) = 3
                PtoIn(0) = 375 : PtoIn(1) = lon : PtoIn(2) = 0
                PtoFin(0) = 375 : PtoFin(1) = 0 : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.Lineweight = 0.2
                line.color = 150

                PtoIn(0) = 450 : PtoIn(1) = lon : PtoIn(2) = 0
                PtoFin(0) = 450 : PtoFin(1) = 0 : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.Lineweight = 0.2
                line.color = 150
                'espesor de linea
                'line.Lineweight = Autodesk.AutoCAD.Interop.Common.ACAD_LWEIGHT.acLnWt020

                PtoIn(0) = 525 : PtoIn(1) = lon : PtoIn(2) = 0
                PtoFin(0) = 525 : PtoFin(1) = 0 : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.Lineweight = 0.2
                line.color = 150

            ElseIf RadioButton2.Checked = True Then

                exe.Cells(3, 4) = 2
                PtoIn(0) = 375 : PtoIn(1) = lon : PtoIn(2) = 0
                PtoFin(0) = 375 : PtoFin(1) = 0 : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.Lineweight = 0.2
                line.color = 150

                PtoIn(0) = 525 : PtoIn(1) = lon : PtoIn(2) = 0
                PtoFin(0) = 525 : PtoFin(1) = 0 : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.Lineweight = 0.2
                line.color = 150

            End If
            '*******************************************************************************************************************************************************************
            '*****************************************************************************************************************************************************************************
            'EN ESTA PARTE DEL PROCESO SE CREA BARRAS DE NEUTRO(ncircuitos=ncircuitos/2), BARRITAS PEQUEÑAS DE AMBOS LASDOS DEL TABLERO, Y SE PREGUNTA DE COLOR SERAN ESTAS AL DIBUJARCE"*
            '*****************************************************************************************************************************************************************************
            'color = InputBox("de que color desea las barras") variables para asignar un color a dibujo(conjelada temporarmente)

            Dim hg ' contador para incrementar los numeros de circuitos es deacir enumera cada fase
            hg = 0

            g = lon

            For i = 1 To Val(ncircuitos)

                g = g - 150

                PtoIn(0) = 80.4375 : PtoIn(1) = g : PtoIn(2) = 0
                PtoFin(0) = 0 : PtoFin(1) = g : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.color = 150

                'escribir en autocad"
                hg = hg + 1
                insPoint(0) = -40    'Set insertion point x coordinate
                insPoint(1) = g     'Set insertion point y coordinate
                insPoint(2) = 0      'Set insertion point z coordinate
                textHeight = 20                'Set text height to 1.0
                textStr = hg   'Set the text string
                'Create Text object
                text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                text.color = 2 'es color amarillo

                PtoIn(0) = 751.8125 : PtoIn(1) = g : PtoIn(2) = 0
                PtoFin(0) = 148.1875 : PtoFin(1) = g : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.color = 150


                PtoIn(0) = 900 : PtoIn(1) = g : PtoIn(2) = 0
                PtoFin(0) = 819.5625 : PtoFin(1) = g : PtoFin(2) = 0
                line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                line.color = 150

                'escribir en autocad"
                hg = hg + 1
                insPoint(0) = 930    'Set insertion point x coordinate
                insPoint(1) = g     'Set insertion point y coordinate
                insPoint(2) = 0      'Set insertion point z coordinate
                textHeight = 20                'Set text height to 1.0
                textStr = hg   'Set the text string
                'Create Text object
                text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                text.color = 2 'es color amarillo
            Next

            MsgBox(ncircuitos) '////////

            '*****************************************************************************************************************************************************************************************
            '**********************************************************************************'NOTA: PARA LA ACTUALIZACION...
            'EN ESTA PARTE DEL PROCESO SE DIBUJA LOS NODOS(CIRCULOS) PARA MONOFASICO SOLAMENTE*
            '**********************************************************************************
            'Dim relleno As Autodesk.AutoCAD.Interop.Common.AcadHatch
            ' Dim lugar, lugar2 As Double
            ' lugar = lon
            ' lugar2 = lon
            ' For i = 1 To (Val(ncircuitos) / 2)

            'primera bolita
            'If i = 1 Then
            ' lugar = lugar - 150

            ' centroo(0) = 375.0 : centroo(1) = lugar : centroo(2) = 0
            ' radioo = 17.125
            ' circulo = objautocad.ActiveDocument.ModelSpace.AddCircle(centroo, radioo)
            'relleno = objautocad.ActiveDocument.ModelSpace.AddHatch(, radioo, circulo, 1, circulo)
            '  Else
            '  lugar = lugar - 300
            '  centroo(0) = 375.0 : centroo(1) = lugar : centroo(2) = 0
            '  radioo = 17.125
            '   circulo = objautocad.ActiveDocument.ModelSpace.AddCircle(centroo, radioo)

            ' End If
            'segunda bolita 
            ' lugar2 = lugar2 - 300
            ' centroo(0) = 525.0 : centroo(1) = lugar2 : centroo(2) = 0
            ' radioo = 17.125
            '   circulo = objautocad.ActiveDocument.ModelSpace.AddCircle(centroo, radioo)

            '  Next
            '****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
            'PARTE DEL PROGRAMA  Q PREGUNTA POR CADA CIRCUITO OPCION A ELEGIR(MONOFASICO(220),TRIFASICO,ILUMUNACION,RESERVA,TOMACORRIENTE,FIN), EN ESTA PARTTE TABN EL TABLERO TIENE LA INTELIGENCIA DE QUE UNA PARTE ESTE LLENA PREVIAMENTE POR UN (MONOFASICO OTRIFASICO) ESTE SE DA CUENTA Y PASA APREGUNTAR AL QUE NO OCUPA DIBUJO, Y INDICA EN Q POCICION VAEL USUARIO,TABN SE CALCULA VAMP Y ES COLOCADO EN EL TABLERO*
            '****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
            '
            If RadioButton3.Checked = True Then
                exe.Cells(3, 2) = "208V"
            Else
                exe.Cells(3, 2) = "240V"
            End If

            If RadioButton5.Checked = True Then
                exe.Cells(2, 4) = "NLAB"
            Else
                exe.Cells(2, 4) = "NAB"
            End If

            tt1 = lon ' `primera colunna
            ttt1 = lon

            tt2 = lon ' segunda colunna
            ttt2 = lon

            Dim tt11, tt12
            tt11 = lon 'variable utilizada en arcos
            tt12 = lon

            Dim ii, iii 'variable utilizada en kk= 5 y 4

            Dim yy, yyy 'variables utilizadas en trifasico algoritmo
            yyy = 0
            yy = 0

            Dim increment, incrementt, incrementtt ' para saber cuanto va en cada circuito,para saber cual circuito estoy cubriendo en monofasico y el ultimo para saber cual es cicuitos estoy cubriendo en truifasico
            increment = 0

            Dim o, o2, o3, o4 As Integer 'apagadores y encendedores en trifasico y monofasico..

            Dim suma_iluminacion, suma_tomacorriente, suma_monofasico_trifasico As Double 'variables utilizadas para acumular cargas va

            suma_iluminacion = 0
            suma_tomacorriente = 0
            suma_monofasico_trifasico = 0

            Dim factor_demanda_tomacorriente, factor_demanda_iluminacion, factor_demanda_fuerza As Double 'variable q contienen FACTOR DE DEAMANDA de (TOMACORRIENTES,ALUMBRAADOS,FUERZA(MONOFASICO_TRIFASICO))
            Dim demanda_tomacorriente, demanda_ilumiinacion, demanda_fuerza, demanda_total As Double 'variables q contienen la demanda de tomacorriente,iluminacion y fuerza(monofasico-trifasico)y la suma de todas esta q es la demanda total
            Dim I_total As Double ' variable q contiene la corriente total del tablero 

            Dim KTT, KTT2 As Double
            KTT = 1
            KTT2 = 1
            'variable utilizadas para las cell de excel
            Dim gg As Integer
            gg = 11
            Dim movera, moverb : movera = 1 : moverb = 1


            For i = 1 To Val(ncircuitos)

                Dim jab, jabb, jak, jakk, jal, jall

                Dim va, amp 'VARIABLE Q GUARDA CARGA voltio amperio


                If i = 1 Then

                    jab = 184.3424
                    jabb = 206.25
                    'para el de kk= 2
                    jak = 334.34
                    jakk = 355.95
                    'para el de kk =3
                    jal = 484.34
                    jall = 505.64

                Else
                    jab = 150
                    jabb = 149.695
                    'para el de kk= 2
                    jak = 300
                    jakk = 300
                    'para el de kk =3
                    jal = 450
                    jall = 450
                End If

                If o = 1 Then ' explicar a papa el proceso de como trabaja este algoritmo
                    o = 0
                    increment = increment + 1
                    ' incrementt = increment + 2
                    'incrementtt = increment + 3
                    GoTo nm

                End If

                If o3 = 3 Then 'algoritmo para movilizar un trifasico
                    yy = yy + 1
                    increment = increment + 1
                    'incrementt = increment + 2
                    'incrementtt = increment + 3
                    If yy = 2 Then
                        o3 = 0
                        yy = 0
                    End If
                    GoTo nm
                End If 'fin del algoritmo

                'tk= comando para viajar sin trasar linea
                increment = increment + 1
                incrementt = increment + 2
                incrementtt = increment + 4
                gg = gg + 1

                On Error Resume Next
o22:


                kk = v_deficode(movera) ' tengo q hacerqveaqueesun 110( preguntar a paoa)
                'corrector de errores.. si se introduce un tipo de variable q no relaciona con la vaibale repite l apregunta
                If Err.Number <> 0 Then
                    Err.Clear()
                    MsgBox("error ha introducido un valor caracter", MsgBoxStyle.Critical)
                    GoTo o22
                End If

                If kk = 4 Then

                    exe.Cells(gg, 1) = increment
                    exe.Cells(gg, 3) = "X"

                    tt1 = tt1 - jab

                    PtoIn(0) = 336.41 : PtoIn(1) = tt1 : PtoIn(2) = 0
                    PtoFin(0) = 0 : PtoFin(1) = tt1 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6

                    ttt1 = ttt1 - jabb ' segunda barrita
                    PtoIn(0) = 317.25 : PtoIn(1) = ttt1 : PtoIn(2) = 0
                    PtoFin(0) = 0 : PtoFin(1) = ttt1 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6

                    'en este parte de esta opcion se calcula los (amp) en el programa
repit:              ' hacer alos demascuando los amps son myor a 100repetirde nuevo esta pregunta
                    va = v_definitivo(movera)
                    exe.Cells(gg, 4) = va
                    suma_tomacorriente = suma_tomacorriente + va 'acumula cargas va de tomacorriente y tomacorriente personalizado
                    amp = (va) / (110) 'ecuacion q es utilizada para calcaular los amper(alumbrado,tomacorriente,(1))

                    Dim texpp
                    texpp = tt1 'distancia en el momento q esta parado( inicio ejemplo tt1= tt1 - jab) en ese momento
                    insPoint(0) = 164.25    'Set insertion point x coordinate
                    insPoint(1) = (texpp + 38) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    If (amp > 0 And amp <= 20) Then
                        textStr = ("20A-1P")
                    ElseIf (amp > 20 And amp <= 30) Then
                        textStr = "30A-1P"
                    ElseIf (amp > 30 And amp <= 50) Then
                        textStr = "50A-1P"
                    ElseIf (amp > 50 And amp <= 60) Then
                        textStr = "60A-1P"
                    ElseIf (amp > 60 And amp <= 80) Then
                        textStr = "80A-1P"
                    ElseIf (amp > 80 And amp <= 100) Then
                        textStr = "100A-1P"
                    ElseIf (amp > 100 And amp <= 150) Then
                        textStr = "150A-1P"
                    ElseIf (amp > 150 And amp <= 175) Then
                        textStr = "175A-1P"
                    ElseIf (amp > 175 And amp <= 200) Then
                        textStr = "200A-1P"
                    ElseIf (amp > 200 And amp <= 225) Then
                        textStr = "225A-1P"
                    ElseIf (amp > 225 And amp <= 250) Then
                        textStr = "250A-1P"
                    ElseIf (amp > 250 And amp <= 275) Then
                        textStr = "275A-1P"
                    ElseIf (amp > 275 And amp <= 300) Then
                        textStr = "300A-1P"
                    ElseIf (amp > 300 And amp <= 350) Then
                        textStr = "350A-1P"
                    ElseIf (amp > 350 And amp <= 400) Then
                        textStr = "400A-1P"
                    ElseIf (amp > 400 And amp <= 450) Then
                        textStr = "450A-1P"
                    Else
                        GoTo repit
                    End If
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                    text.color = 2
                    exe.Cells(gg, 6) = UCase(textStr)

                    If RadioButton1.Checked = True Then
                        If KTT = 1 Then
                            exe.Cells(gg, 7) = "A"
                        ElseIf KTT = 2 Then
                            exe.Cells(gg, 7) = "B"
                        Else
                            exe.Cells(gg, 7) = "C"
                        End If
                    End If

                    If RadioButton2.Checked = True Then
                        If KTT2 = 1 Then
                            exe.Cells(gg, 7) = "A"
                        ElseIf KTT2 = 2 Then
                            exe.Cells(gg, 7) = "B"
                        End If
                    End If

                    tt11 = tt11 - 150 ' arco de las fases NOTA: mejorar la creacion del arco..
                    centro(0) = 335 - (246.81 - 26.1) : centro(1) = tt11 : centro(2) = 0  'arco
                    radio = 33.88
                    anginic = 0 ' angulos trabajan contra las abujas del relog
                    angfinal = -9.4012
                    arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
                    arco.color = 6

                    Dim nombrem, tesxm

                    nombrem = v_definombre(movera)
                    movera += 1
                    exe.Cells(gg, 8) = UCase(nombrem)

                    If CheckBox1.Checked = True Then ' condicion q si se escoje mayuscula los caracteres salgan en mayuscula
                        nombrem = UCase(nombrem)
                    End If

                    tesxm = tt1
                    insPoint(0) = -336.41    'Set insertion point x coordinate
                    insPoint(1) = (tesxm + 35.26) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    textStr = nombrem       'Set the text string
                    'Create Text object
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                    text.color = 2

                ElseIf kk = 0 Then
                    exe.Cells(gg, 2) = "X"
                    exe.Cells(gg, 1) = increment

                    If i = 1 Then
                        l1 = 1
                    End If 'nda q ver con el codigo q se aplica en esta opcion dento del algoritmo


                    tt1 = tt1 - jab
                    ttt1 = ttt1 - jabb ' barrita que no es colocada pero q de igual forma se tiene q ir restando para el proximo que lo requiera

                    PtoIn(0) = 336.41 : PtoIn(1) = tt1 : PtoIn(2) = 0
                    PtoFin(0) = 0 : PtoFin(1) = tt1 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6
                    'en este parte de esta opcion se calcula los (amp) en el programa
repit5:
                    va = v_definitivo(movera)
                    movera = movera + 1

                    exe.Cells(gg, 4) = va
                    suma_iluminacion = suma_iluminacion + va 'acumulador de cargas va para iluminacion
                    amp = (va) / (110) 'ecuacion q es utilizada para calcaular los amper(alumbrado,tomacorriente,(1))

                    Dim textt
                    textt = tt1 'distancia en el momento q esta parado( inicio ejemplo tt1= tt1 - jab) en ese momento
                    insPoint(0) = 164.25      'Set insertion point x coordinate
                    insPoint(1) = (textt + 38) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    If (amp > 0 And amp <= 20) Then
                        textStr = ("20A-1P")
                    ElseIf (amp > 20 And amp <= 30) Then
                        textStr = "30A-1P"
                    ElseIf (amp > 30 And amp <= 50) Then
                        textStr = "50A-1P"
                    ElseIf (amp > 50 And amp <= 60) Then
                        textStr = "60A-1P"
                    ElseIf (amp > 60 And amp <= 80) Then
                        textStr = "80A-1P"
                    ElseIf (amp > 80 And amp <= 100) Then
                        textStr = "100A-1P"
                    ElseIf (amp > 100 And amp <= 150) Then
                        textStr = "150A-1P"
                    ElseIf (amp > 150 And amp <= 175) Then
                        textStr = "175A-1P"
                    ElseIf (amp > 175 And amp <= 200) Then
                        textStr = "200A-1P"
                    ElseIf (amp > 200 And amp <= 225) Then
                        textStr = "225A-1P"
                    ElseIf (amp > 225 And amp <= 250) Then
                        textStr = "250A-1P"
                    ElseIf (amp > 250 And amp <= 275) Then
                        textStr = "275A-1P"
                    ElseIf (amp > 275 And amp <= 300) Then
                        textStr = "300A-1P"
                    ElseIf (amp > 300 And amp <= 350) Then
                        textStr = "350A-1P"
                    ElseIf (amp > 350 And amp <= 400) Then
                        textStr = "400A-1P"
                    ElseIf (amp > 400 And amp <= 450) Then
                        textStr = "450A-1P"
                    Else
                        GoTo repit5
                    End If
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight) 'Create Text object
                    text.color = 2
                    exe.Cells(gg, 6) = UCase(textStr)

                    If RadioButton1.Checked = True Then
                        If KTT = 1 Then
                            exe.Cells(gg, 7) = "A"
                        ElseIf KTT = 2 Then
                            exe.Cells(gg, 7) = "B"
                        Else
                            exe.Cells(gg, 7) = "C"
                        End If
                    End If

                    If RadioButton2.Checked = True Then
                        If KTT2 = 1 Then
                            exe.Cells(gg, 7) = "A"
                        ElseIf KTT2 = 2 Then
                            exe.Cells(gg, 7) = "B"
                        End If
                    End If

                    tt11 = tt11 - 150 ' arco de las fases NOTA: mejorar la creacion del arco..
                    centro(0) = 335 - (246.81 - 26.1) : centro(1) = tt11 : centro(2) = 0  'arco
                    radio = 33.88
                    anginic = 0 ' angulos trabajan contra las abujas del relog
                    angfinal = -9.4012
                    arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
                    arco.color = 6

                    Dim tesxms

                    tesxms = tt1

                    insPoint(0) = -336.41    'Set insertion point x coordinate
                    insPoint(1) = (tesxms + 35.26) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0

                    If CheckBox1.Checked = True Then ' condicion q si se escoje mayuscula los caracteres salgan en mayuscula
                        textStr = ("ILUMINACION")   'Set the text string
                    Else
                        textStr = ("iluminacion")   'Set the text string
                    End If
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight) 'Create Text object
                    text.color = 2
                    exe.Cells(gg, 8) = UCase(textStr)
                    'ElseIf kk = "fin" Then exit for

                ElseIf kk = 2 Then
                    exe.Cells(gg, 3) = "X"
                    exe.Cells(gg, 1) = increment & "," & incrementt

                    tt1 = tt1 - jak
                    ttt1 = ttt1 - jakk
                    o = 1

                    If CheckBox2.Checked = False Then
                        PtoIn(0) = 317.25 : PtoIn(1) = tt1 : PtoIn(2) = 0
                        PtoFin(0) = 0 : PtoFin(1) = tt1 : PtoFin(2) = 0
                        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                        line.color = 6
                    Else
                        'se selecciona la casilla "principal" en la interfacee 
                        PtoIn(0) = 336.41 : PtoIn(1) = tt1 : PtoIn(2) = 0
                        PtoFin(0) = 0 : PtoFin(1) = tt1 : PtoFin(2) = 0
                        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                        line.color = 6

                        PtoIn(0) = 317.25 : PtoIn(1) = ttt1 : PtoIn(2) = 0
                        PtoFin(0) = 0 : PtoFin(1) = ttt1 : PtoFin(2) = 0
                        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                        line.color = 6
                    End If
                    'en este parte de esta opcion se calcula los (amp) en el programa
repit1:
                    va = v_definitivo(movera)

                    exe.Cells(gg, 4) = va
                    suma_monofasico_trifasico = suma_monofasico_trifasico + va 'sumatoria de todas las cargas va en monofasico-trifasico

                    If RadioButton3.Checked = True Then ' SI ES SELECCIONADO 208 V
                        amp = (va) / (208) 'ecuacion q es utilizada para calcaular los amper(alumbrado,tomacorriente,(1))
                    ElseIf RadioButton4.Checked = True Then ' SI ES SELECCIONADO 204
                        amp = (va) / (240)
                    End If

                    Dim texk
                    texk = tt1 'distancia en el momento q esta parado( inicio ejemplo tt1= tt1 - jab) en ese momento
                    insPoint(0) = 164.25    'Set insertion point x coordinate
                    insPoint(1) = (texk + 188) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    If (amp > 0 And amp <= 20) Then
                        textStr = ("20A-2P")
                    ElseIf (amp > 20 And amp <= 30) Then
                        textStr = "30A-2P"
                    ElseIf (amp > 30 And amp <= 50) Then
                        textStr = "50A-2P"
                    ElseIf (amp > 50 And amp <= 60) Then
                        textStr = "60A-2P"
                    ElseIf (amp > 60 And amp <= 80) Then
                        textStr = "80A-2P"
                    ElseIf (amp > 80 And amp <= 100) Then
                        textStr = "100A-2P"
                    ElseIf (amp > 100 And amp <= 150) Then
                        textStr = "150A-2P"
                    ElseIf (amp > 150 And amp <= 175) Then
                        textStr = "175A-2P"
                    ElseIf (amp > 175 And amp <= 200) Then
                        textStr = "200A-2P"
                    ElseIf (amp > 200 And amp <= 225) Then
                        textStr = "225A-2P"
                    ElseIf (amp > 225 And amp <= 250) Then
                        textStr = "250A-2P"
                    ElseIf (amp > 250 And amp <= 275) Then
                        textStr = "275A-2P"
                    ElseIf (amp > 275 And amp <= 300) Then
                        textStr = "300A-2P"
                    ElseIf (amp > 300 And amp <= 350) Then
                        textStr = "350A-2P"
                    ElseIf (amp > 350 And amp <= 400) Then
                        textStr = "400A-2P"
                    ElseIf (amp > 400 And amp <= 450) Then
                        textStr = "450A-2P"

                    Else
                        MsgBox("pase por aki") ' presenta ERROR ERROR ERROR
                        GoTo repit1
                    End If
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight) 'Create Text object
                    text.color = 2
                    exe.Cells(gg, 6) = UCase(textStr)

                    If RadioButton1.Checked = True Then
                        If KTT = 1 Then
                            exe.Cells(gg, 7) = "A,B"
                        ElseIf KTT = 2 Then
                            exe.Cells(gg, 7) = "B,C"
                        Else
                            exe.Cells(gg, 7) = "C,A"
                        End If
                    End If

                    If RadioButton2.Checked = True Then
                        If KTT2 = 1 Then
                            exe.Cells(gg, 7) = "A,B"
                        ElseIf KTT2 = 2 Then
                            exe.Cells(gg, 7) = "B,A"
                        End If
                    End If

                    For iii = 1 To 2
                        tt11 = tt11 - 150 ' arco de las fases NOTA: mejorar la creacion del arco..
                        centro(0) = 335 - (246.81 - 26.1) : centro(1) = tt11 : centro(2) = 0  'arco
                        radio = 33.88
                        anginic = 0 ' angulos trabajan contra las abujas del relog
                        angfinal = -9.4012
                        arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
                        arco.color = 6
                    Next iii

                    PtoIn(0) = 114.29 : PtoIn(1) = tt11 + 183.88 : PtoIn(2) = 0 '( coloca la barra vertical entre los acos monofasicos )
                    PtoFin(0) = 114.29 : PtoFin(1) = tt1 + 68.22 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6

                    Dim nombreet, tesxt

                    tesxt = tt1
                    nombreet = v_definombre(movera)
                    movera = movera + 1

                    exe.Cells(gg, 8) = UCase(nombreet)
                    If CheckBox1.Checked = True Then ' condicion q si se escoje mayuscula los caracteres salgan en mayuscula
                        nombreet = UCase(nombreet)
                    Else
                    End If

                    insPoint(0) = -336.41    'Set insertion point x coordinate
                    insPoint(1) = (tesxt + 109.34) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    textStr = nombreet       'Set the text string
                    'Create Text object
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                    text.color = 2

                ElseIf kk = 8 Then
                    Exit For

                ElseIf kk = 3 Then
                    exe.Cells(gg, 3) = "X"
                    exe.Cells(gg, 1) = increment & "," & incrementt & "," & incrementtt
                    tt1 = tt1 - jal
                    ttt1 = ttt1 - jall
                    o3 = 3
                    If CheckBox2.Checked = False Then
                        PtoIn(0) = 317.25 : PtoIn(1) = tt1 : PtoIn(2) = 0
                        PtoFin(0) = 0 : PtoFin(1) = tt1 : PtoFin(2) = 0
                        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                        line.color = 6
                    Else
                        'se selecciona la casilla "principal" en la interfacee 
                        PtoIn(0) = 336.41 : PtoIn(1) = tt1 : PtoIn(2) = 0
                        PtoFin(0) = 0 : PtoFin(1) = tt1 : PtoFin(2) = 0
                        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                        line.color = 6
                        PtoIn(0) = 317.25 : PtoIn(1) = ttt1 : PtoIn(2) = 0
                        PtoFin(0) = 0 : PtoFin(1) = ttt1 : PtoFin(2) = 0
                        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                        line.color = 6
                    End If
                    'en este parte de esta opcion se calcula los (amp) en el programa
repit4:

                    va = v_definitivo(movera)
                    exe.Cells(gg, 4) = va
                    suma_monofasico_trifasico = suma_monofasico_trifasico + va 'sumatoria de todas las cargas va en monofasico-trifasico
                    amp = (va) / (Math.Sqrt(3) * 208) 'ecuacion q es utilizada para calcaular los amper(alumbrado,tomacorriente,(1))

                    Dim texdss
                    texdss = tt1 'distancia en el momento q esta parado( inicio ejemplo tt1= tt1 - jab) en ese momento
                    insPoint(0) = 164.25     'Set insertion point x coordinate
                    insPoint(1) = (texdss + 338) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    If (amp > 0 And amp <= 20) Then
                        textStr = ("20A-3P")
                    ElseIf (amp > 20 And amp <= 30) Then
                        textStr = "30A-3P"
                    ElseIf (amp > 30 And amp <= 50) Then
                        textStr = "50A-3P"
                    ElseIf (amp > 50 And amp <= 60) Then
                        textStr = "60A-3P"
                    ElseIf (amp > 60 And amp <= 80) Then
                        textStr = "80A-3P"
                    ElseIf (amp > 80 And amp <= 100) Then
                        textStr = "100A-3P"
                    ElseIf (amp > 100 And amp <= 150) Then
                        textStr = "150A-3P"
                    ElseIf (amp > 150 And amp <= 175) Then
                        textStr = "175A-3P"
                    ElseIf (amp > 175 And amp <= 200) Then
                        textStr = "200A-3P"
                    ElseIf (amp > 200 And amp <= 225) Then
                        textStr = "225A-3P"
                    ElseIf (amp > 225 And amp <= 250) Then
                        textStr = "250A-3P"
                    ElseIf (amp > 250 And amp <= 275) Then
                        textStr = "275A-3P"
                    ElseIf (amp > 275 And amp <= 300) Then
                        textStr = "300A-3P"
                    ElseIf (amp > 300 And amp <= 350) Then
                        textStr = "350A-3P"
                    ElseIf (amp > 350 And amp <= 400) Then
                        textStr = "400A-3P"
                    ElseIf (amp > 400 And amp <= 450) Then
                        textStr = "450A-3P"
                    Else
                        GoTo repit4
                    End If
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight) 'Create Text object
                    text.color = 2
                    exe.Cells(gg, 6) = UCase(textStr)

                    If RadioButton1.Checked = True Then
                        If KTT = 1 Then
                            exe.Cells(gg, 7) = "A,B,C"
                        ElseIf KTT = 2 Then
                            exe.Cells(gg, 7) = "B,C,A"
                        Else
                            exe.Cells(gg, 7) = "C,A,B"
                        End If
                    End If

                    For ii = 1 To 3
                        tt11 = tt11 - 150 ' arco de las fases NOTA: mejorar la creacion del arco..
                        centro(0) = 335 - (246.81 - 26.1) : centro(1) = tt11 : centro(2) = 0  'arco
                        radio = 33.88
                        anginic = 0 ' angulos trabajan contra las abujas del relog
                        angfinal = -9.4012
                        arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
                        arco.color = 6
                    Next ii

                    PtoIn(0) = 114.29 : PtoIn(1) = tt11 + 333.88 : PtoIn(2) = 0  '( coloca las lineas entre los arcos trifasicos)
                    PtoFin(0) = 114.29 : PtoFin(1) = tt1 + 68.22 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6

                    Dim nombree, tesx
                    tesx = tt1
                    nombree = v_definombre(movera)
                    movera = movera + 1
                    exe.Cells(gg, 8) = UCase(nombree)
                    If CheckBox1.Checked = True Then ' condicion q si se escoje mayuscula los caracteres salgan en mayuscula
                        nombree = UCase(nombree)
                    Else
                    End If

                    insPoint(0) = -336.41    'Set insertion point x coordinate
                    insPoint(1) = (tesx + 184.34) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    textStr = nombree       'Set the text string
                    'Create Text object
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                    text.color = 2

                ElseIf kk = 1 Then
                    exe.Cells(gg, 3) = "X"
                    exe.Cells(gg, 1) = increment

                    tt1 = tt1 - jab

                    PtoIn(0) = 336.41 : PtoIn(1) = tt1 : PtoIn(2) = 0
                    PtoFin(0) = 0 : PtoFin(1) = tt1 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6

                    ttt1 = ttt1 - jabb ' segunda barrita
                    PtoIn(0) = 317.25 : PtoIn(1) = ttt1 : PtoIn(2) = 0
                    PtoFin(0) = 0 : PtoFin(1) = ttt1 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6
                    'en esta parte de esta opcion se calcula los (amp) en el programa
repit9:
                    va = v_definitivo(movera)
                    movera = movera + 1

                    exe.Cells(gg, 4) = va
                    suma_tomacorriente = suma_tomacorriente + va 'acumula cargas va de tomacorriente y tomacorriente personalizado                    amp = (va * 100) / (110 * fp) 'ecuacion q es utilizada para calcaular los amper(alumbrado,tomacorriente,(1))
                    amp = (va) / (110) 'ecuacion q es utilizada para calcaular los amper(alumbrado,tomacorriente,(1))

                    Dim textll

                    textll = tt1 'distancia en el momento q esta parado( inicio ejemplo tt1= tt1 - jab) en ese momento
                    insPoint(0) = 164.25    'Set insertion point x coordinate
                    insPoint(1) = (textll + 38) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    If (amp > 0 And amp <= 20) Then
                        textStr = ("20A-1P")
                    ElseIf (amp > 20 And amp <= 30) Then
                        textStr = "30A-1P"
                    ElseIf (amp > 30 And amp <= 50) Then
                        textStr = "50A-1P"
                    ElseIf (amp > 50 And amp <= 60) Then
                        textStr = "60A-1P"
                    ElseIf (amp > 60 And amp <= 80) Then
                        textStr = "80A-1P"
                    ElseIf (amp > 80 And amp <= 100) Then
                        textStr = "100A-1P"
                    ElseIf (amp > 100 And amp <= 150) Then
                        textStr = "150A-1P"
                    ElseIf (amp > 150 And amp <= 175) Then
                        textStr = "175A-1P"
                    ElseIf (amp > 175 And amp <= 200) Then
                        textStr = "200A-1P"
                    ElseIf (amp > 200 And amp <= 225) Then
                        textStr = "225A-1P"
                    ElseIf (amp > 225 And amp <= 250) Then
                        textStr = "250A-1P"
                    ElseIf (amp > 250 And amp <= 275) Then
                        textStr = "275A-1P"
                    ElseIf (amp > 275 And amp <= 300) Then
                        textStr = "300A-1P"
                    ElseIf (amp > 300 And amp <= 350) Then
                        textStr = "350A-1P"
                    ElseIf (amp > 350 And amp <= 400) Then
                        textStr = "400A-1P"
                    ElseIf (amp > 400 And amp <= 450) Then
                        textStr = "450A-1P"

                    Else
                        GoTo repit9
                    End If
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight) 'Create Text object
                    text.color = 2
                    exe.Cells(gg, 6) = UCase(textStr)

                    If RadioButton1.Checked = True Then
                        If KTT = 1 Then
                            exe.Cells(gg, 7) = "A"
                        ElseIf KTT = 2 Then
                            exe.Cells(gg, 7) = "B"
                        Else
                            exe.Cells(gg, 7) = "C"
                        End If
                    End If

                    If RadioButton2.Checked = True Then
                        If KTT2 = 1 Then
                            exe.Cells(gg, 7) = "A"
                        ElseIf KTT2 = 2 Then
                            exe.Cells(gg, 7) = "B"
                        End If
                    End If

                    tt11 = tt11 - 150 ' arco de las fases NOTA: mejorar la creacion del arco..
                    centro(0) = 335 - (246.81 - 26.1) : centro(1) = tt11 : centro(2) = 0  'arco
                    radio = 33.88
                    anginic = 0 ' angulos trabajan contra las abujas del relog
                    angfinal = -9.4012
                    arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
                    arco.color = 6

                    Dim tesxmh
                    tesxmh = tt1
                    insPoint(0) = -440   'Set insertion point x coordinate
                    insPoint(1) = (tesxmh + 35.26) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0

                    If CheckBox1.Checked = True Then ' condicion q si se escoje mayuscula los caracteres salgan en mayuscula
                        textStr = "TOMACORRIENTES"      'Set the text string
                    Else
                        textStr = "tomacorriente"
                    End If
                    exe.Cells(gg, 8) = UCase(textStr)
                    'Create Text object
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                    text.color = 2

                ElseIf kk = 5 Then

                    exe.Cells(gg, 1) = increment

                    tt1 = tt1 - jab
                    ttt1 = ttt1 - jabb
                    tt11 = tt11 - 150 ' reserva no lleva arco pero de todas formas es utilizado para ovilizar si este esta presente antes de un final
                    Dim tesxmhj

                    tesxmhj = tt1
                    insPoint(0) = -330  'Set insertion point x coordinate
                    insPoint(1) = (tesxmhj + 35.26) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0

                    If CheckBox1.Checked = True Then ' condicion q si se escoje mayuscula los caracteres salgan en mayuscula
                        textStr = "RESERVA"      'Set the text string
                    Else
                        textStr = "reserva"      'Set the text string
                    End If
                    exe.Cells(gg, 8) = UCase(textStr)
                    'Create Text object
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                    text.color = 2

                Else
                    tt1 = tt1 - jab
                    ttt1 = ttt1 - jabb

                End If


nm:

                If o4 = 4 Then 'algoritmo para movilizar un trifasico
                    yyy = yyy + 1
                    increment = increment + 1
                    'incrementt = increment + 2
                    'incrementtt = increment + 3
                    If yyy = 2 Then
                        o4 = 0
                        yyy = 0
                    End If
                    GoTo nmk
                End If 'fin del algoritmo

                If o2 = 2 Then
                    o2 = 0
                    increment = increment + 1
                    ' incrementt = increment + 2
                    '  incrementtt = increment + 3
                    GoTo nmk
                End If

                'segunda parte

                increment = increment + 1
                incrementt = increment + 2
                incrementtt = increment + 4
                gg = gg + 1

                On Error Resume Next ' iniciador de detector de errores
p22:

                kk = v_deficode(movera)

                'corrector de errores.. si se introduce un tipo de variable q no relaciona con la vaibale repite l apregunta
                If Err.Number <> 0 Then
                    Err.Clear()
                    MsgBox("error ha introducido un valor caracter", MsgBoxStyle.Critical)
                    GoTo p22 ' vuelve a preguntar!!
                End If
                If kk = 4 Then

                    exe.Cells(gg, 1) = increment
                    exe.Cells(gg, 3) = "X"
                    tt2 = tt2 - jab
                    PtoIn(0) = 900 : PtoIn(1) = tt2 : PtoIn(2) = 0
                    PtoFin(0) = 563.5888 : PtoFin(1) = tt2 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6

                    ttt2 = ttt2 - jabb
                    PtoIn(0) = 900 : PtoIn(1) = ttt2 : PtoIn(2) = 0
                    PtoFin(0) = 582.75 : PtoFin(1) = ttt2 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6

                    'en este parte de esta opcion se calcula los (amp) en el programa
repit6:
                    va = v_definitivo(movera)
                    exe.Cells(gg, 4) = va
                    suma_tomacorriente = suma_tomacorriente + va 'acumula cargas va de tomacorriente y tomacorriente personalizado
                    amp = (va) / (110) 'ecuacion q es utilizada para calcaular los amper(alumbrado,tomacorriente,(1))

                    Dim textf
                    textf = tt2 'distancia en el momento q esta parado( inicio ejemplo tt1= tt1 - jab) en ese momento
                    insPoint(0) = 585.42 'Set insertion point x coordinate
                    insPoint(1) = (textf + 38) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    If (amp > 0 And amp <= 20) Then
                        textStr = ("20A-1P")
                    ElseIf (amp > 20 And amp <= 30) Then
                        textStr = "30A-1P"
                    ElseIf (amp > 30 And amp <= 50) Then
                        textStr = "50A-1P"
                    ElseIf (amp > 50 And amp <= 60) Then
                        textStr = "60A-1P"
                    ElseIf (amp > 60 And amp <= 80) Then
                        textStr = "80A-1P"
                    ElseIf (amp > 80 And amp <= 100) Then
                        textStr = "100A-1P"
                    ElseIf (amp > 100 And amp <= 150) Then
                        textStr = "150A-1P"
                    ElseIf (amp > 150 And amp <= 175) Then
                        textStr = "175A-1P"
                    ElseIf (amp > 175 And amp <= 200) Then
                        textStr = "200A-1P"
                    ElseIf (amp > 200 And amp <= 225) Then
                        textStr = "225A-1P"
                    ElseIf (amp > 225 And amp <= 250) Then
                        textStr = "250A-1P"
                    ElseIf (amp > 250 And amp <= 275) Then
                        textStr = "275A-1P"
                    ElseIf (amp > 275 And amp <= 300) Then
                        textStr = "300A-1P"
                    ElseIf (amp > 300 And amp <= 350) Then
                        textStr = "350A-1P"
                    ElseIf (amp > 350 And amp <= 400) Then
                        textStr = "400A-1P"
                    ElseIf (amp > 400 And amp <= 450) Then
                        textStr = "450A-1P"

                    Else
                        GoTo repit6
                    End If
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight) 'Create Text object
                    text.color = 2
                    exe.Cells(gg, 6) = UCase(textStr)

                    If RadioButton1.Checked = True Then
                        If KTT = 1 Then
                            exe.Cells(gg, 7) = "A"
                        ElseIf KTT = 2 Then
                            exe.Cells(gg, 7) = "B"
                        Else
                            exe.Cells(gg, 7) = "C"
                        End If
                    End If

                    If RadioButton2.Checked = True Then
                        If KTT2 = 1 Then
                            exe.Cells(gg, 7) = "A"
                        ElseIf KTT2 = 2 Then
                            exe.Cells(gg, 7) = "B"
                        End If
                    End If

                    tt12 = tt12 - 150 ' arco de las fases NOTA: mejorar la creacion del arco..
                    centro(0) = 785.68 : centro(1) = tt12 : centro(2) = 0 'arco
                    radio = 33.88
                    anginic = 0 ' angulos trabajan contra las abujas del relog
                    angfinal = -9.4012
                    arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
                    arco.color = 6

                    Dim nombreu, tesxtu
                    tesxtu = tt2
                    nombreu = v_definombre(movera)
                    movera += 1
                    exe.Cells(gg, 8) = UCase(nombreu)

                    If CheckBox1.Checked = True Then ' condicion q si se escoje mayuscula los caracteres salgan en mayuscula
                        nombreu = UCase(nombreu)
                    Else
                    End If

                    insPoint(0) = 1000    'Set insertion point x coordinate
                    insPoint(1) = (tesxtu + 35.26) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    textStr = nombreu       'Set the text string
                    'Create Text object
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                    text.color = 2

                ElseIf kk = 0 Then
                    exe.Cells(gg, 2) = "X"
                    exe.Cells(gg, 1) = increment

                    If i = 1 Then
                        l2 = 2
                    End If

                    ttt2 = ttt2 - jabb
                    tt2 = tt2 - jab
                    PtoIn(0) = 900 : PtoIn(1) = tt2 : PtoIn(2) = 0
                    PtoFin(0) = 563.5888 : PtoFin(1) = tt2 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6

                    'en este parte de esta opcion se calcula los (amp) en el programa
repit7:
                    va = v_definitivo(movera)
                    movera = movera + 1
                    exe.Cells(gg, 4) = va
                    suma_iluminacion = suma_iluminacion + va 'sumatoria de carga va para iluminacion
                    amp = (va) / (110) 'ecuacion q es utilizada para calcaular los amper(alumbrado,tomacorriente,(1))

                    Dim texth
                    texth = tt2 'distancia en el momento q esta parado( inicio ejemplo tt1= tt1 - jab) en ese momento
                    insPoint(0) = 585.42 'Set insertion point x coordinate
                    insPoint(1) = (texth + 38) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    If (amp > 0 And amp <= 20) Then
                        textStr = ("20A-1P")
                    ElseIf (amp > 20 And amp <= 30) Then
                        textStr = "30A-1P"
                    ElseIf (amp > 30 And amp <= 50) Then
                        textStr = "50A-1P"
                    ElseIf (amp > 50 And amp <= 60) Then
                        textStr = "60A-1P"
                    ElseIf (amp > 60 And amp <= 80) Then
                        textStr = "80A-1P"
                    ElseIf (amp > 80 And amp <= 100) Then
                        textStr = "100A-1P"
                    ElseIf (amp > 100 And amp <= 150) Then
                        textStr = "150A-1P"
                    ElseIf (amp > 150 And amp <= 175) Then
                        textStr = "175A-1P"
                    ElseIf (amp > 175 And amp <= 200) Then
                        textStr = "200A-1P"
                    ElseIf (amp > 200 And amp <= 225) Then
                        textStr = "225A-1P"
                    ElseIf (amp > 225 And amp <= 250) Then
                        textStr = "250A-1P"
                    ElseIf (amp > 250 And amp <= 275) Then
                        textStr = "275A-1P"
                    ElseIf (amp > 275 And amp <= 300) Then
                        textStr = "300A-1P"
                    ElseIf (amp > 300 And amp <= 350) Then
                        textStr = "350A-1P"
                    ElseIf (amp > 350 And amp <= 400) Then
                        textStr = "400A-1P"
                    ElseIf (amp > 400 And amp <= 450) Then
                        textStr = "450A-1P"
                    Else
                        GoTo repit7
                    End If
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight) 'Create Text object
                    text.color = 2
                    exe.Cells(gg, 6) = UCase(textStr)

                    If RadioButton1.Checked = True Then
                        If KTT = 1 Then
                            exe.Cells(gg, 7) = "A"
                        ElseIf KTT = 2 Then
                            exe.Cells(gg, 7) = "B"
                        Else
                            exe.Cells(gg, 7) = "C"
                        End If
                    End If

                    If RadioButton2.Checked = True Then
                        If KTT2 = 1 Then
                            exe.Cells(gg, 7) = "A"
                        ElseIf KTT2 = 2 Then
                            exe.Cells(gg, 7) = "B"
                        End If
                    End If

                    tt12 = tt12 - 150 ' arco de las fases NOTA: mejorar la creacion del arco..
                    centro(0) = 785.68 : centro(1) = tt12 : centro(2) = 0 'arco
                    radio = 33.88
                    anginic = 0 ' angulos trabajan contra las abujas del relog
                    angfinal = -9.4012
                    arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
                    arco.color = 6

                    Dim tesxtukf
                    tesxtukf = tt2

                    insPoint(0) = 1000    'Set insertion point x coordinate
                    insPoint(1) = (tesxtukf + 35.26) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0

                    If CheckBox1.Checked = True Then ' condicion q si se escoje mayuscula los caracteres salgan en mayuscula
                        textStr = "ILUMINACION"     'Set the text string
                    Else
                        textStr = "iluminacion"     'Set the text string
                    End If
                    exe.Cells(gg, 8) = UCase(textStr)
                    'Create Text object
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                    text.color = 2

                    'ElseIf kk = "fin" Then Exit For

                ElseIf kk = 2 Then
                    exe.Cells(gg, 3) = "X"
                    exe.Cells(gg, 1) = (increment & "," & incrementt)

                    'MsgBox(increment & " , " & incrementt) pra verificar un error qpor aki hay 

                    tt2 = tt2 - jak
                    ttt2 = ttt2 - jakk
                    o2 = 2

                    If CheckBox2.Checked = False Then
                        PtoIn(0) = 900 : PtoIn(1) = tt2 : PtoIn(2) = 0
                        PtoFin(0) = 582.75 : PtoFin(1) = tt2 : PtoFin(2) = 0
                        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                        line.color = 6
                    Else
                        'se selecciona la casilla "principal" en la interfacee 
                        PtoIn(0) = 900 : PtoIn(1) = tt2 : PtoIn(2) = 0
                        PtoFin(0) = 563.5888 : PtoFin(1) = tt2 : PtoFin(2) = 0
                        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                        line.color = 6

                        PtoIn(0) = 900 : PtoIn(1) = ttt2 : PtoIn(2) = 0
                        PtoFin(0) = 582.75 : PtoFin(1) = ttt2 : PtoFin(2) = 0
                        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                        line.color = 6
                    End If
                    'en este parte de esta opcion se calcula los (amp) en el programa
repit2:
                    va = v_definitivo(movera) ' se introdiuce la carga automaticamente(metodo automatico)
                    exe.Cells(gg, 4) = va
                    suma_monofasico_trifasico = suma_monofasico_trifasico + va 'sumatoria de todas las cargas va en monofasico-trifasico

                    If RadioButton3.Checked = True Then
                        amp = (va) / (208) 'ecuacion q es utilizada para calcaular los amper(alumbrado,tomacorriente,(1))
                    ElseIf RadioButton4.Checked = True Then
                        amp = (va) / (240)
                    End If

                    Dim texkj
                    texkj = tt2 'distancia en el momento q esta parado( inicio ejemplo tt1= tt1 - jab) en ese momento
                    insPoint(0) = 585.42    'Set insertion point x coordinate
                    insPoint(1) = (texkj + 188) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    If (amp > 0 And amp <= 20) Then
                        textStr = ("20A-2P")
                    ElseIf (amp > 20 And amp <= 30) Then
                        textStr = "30A-2P"
                    ElseIf (amp > 30 And amp <= 50) Then
                        textStr = "50A-2P"
                    ElseIf (amp > 50 And amp <= 60) Then
                        textStr = "60A-2P"
                    ElseIf (amp > 60 And amp <= 80) Then
                        textStr = "80A-2P"
                    ElseIf (amp > 80 And amp <= 100) Then
                        textStr = "100A-2P"
                    ElseIf (amp > 100 And amp <= 150) Then
                        textStr = "150A-2P"
                    ElseIf (amp > 150 And amp <= 175) Then
                        textStr = "175A-2P"
                    ElseIf (amp > 175 And amp <= 200) Then
                        textStr = "200A-2P"
                    ElseIf (amp > 200 And amp <= 225) Then
                        textStr = "225A-2P"
                    ElseIf (amp > 225 And amp <= 250) Then
                        textStr = "250A-2P"
                    ElseIf (amp > 250 And amp <= 275) Then
                        textStr = "275A-2P"
                    ElseIf (amp > 275 And amp <= 300) Then
                        textStr = "300A-2P"
                    ElseIf (amp > 300 And amp <= 350) Then
                        textStr = "350A-2P"
                    ElseIf (amp > 350 And amp <= 400) Then
                        textStr = "400A-2P"
                    ElseIf (amp > 400 And amp <= 450) Then
                        textStr = "450A-2P"

                    Else
                        GoTo repit2
                    End If
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight) 'Create Text object
                    text.color = 2
                    exe.Cells(gg, 6) = UCase(textStr)


                    If RadioButton1.Checked = True Then
                        If KTT = 1 Then
                            exe.Cells(gg, 7) = "A,B"
                        ElseIf KTT = 2 Then
                            exe.Cells(gg, 7) = "B,C"
                        Else
                            exe.Cells(gg, 7) = "C,A"
                        End If
                    End If

                    If RadioButton2.Checked = True Then
                        If KTT2 = 1 Then
                            exe.Cells(gg, 7) = "A,B"
                        ElseIf KTT2 = 2 Then
                            exe.Cells(gg, 7) = "B,A"
                        End If
                    End If

                    For ii = 1 To 2
                        tt12 = tt12 - 150 ' arco de las fases NOTA: mejorar la creacion del arco..
                        centro(0) = 785.68 : centro(1) = tt12 : centro(2) = 0 'arco
                        radio = 33.88
                        anginic = 0 ' angulos trabajan contra las abujas del relog
                        angfinal = -9.4012
                        arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
                        arco.color = 6
                    Next ii

                    PtoIn(0) = 785.69 : PtoIn(1) = tt12 + 183.88 : PtoIn(2) = 0 '( coloca la barra vertical entre los acos monofasicos )
                    PtoFin(0) = 785.69 : PtoFin(1) = tt2 + 68.22 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6

                    Dim nombreer, tesxtt
                    tesxtt = tt2
                    nombreer = v_definombre(movera) ' se introduce el nombre del circuito automaticamente
                    movera = movera + 1 ' dps incrementa para el proximo polo
                    exe.Cells(gg, 8) = UCase(nombreer)
                    If CheckBox1.Checked = True Then ' condicion q si se escoje mayuscula los caracteres salgan en mayuscula
                        nombreer = UCase(nombreer)
                    Else
                    End If

                    insPoint(0) = 1000    'Set insertion point x coordinate
                    insPoint(1) = (tesxtt + 109.34) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    textStr = nombreer       'Set the text string
                    'Create Text object
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                    text.color = 2

                ElseIf kk = 3 Then
                    exe.Cells(gg, 3) = "X"
                    exe.Cells(gg, 1) = increment & "," & incrementt & "," & incrementtt

                    tt2 = tt2 - jal
                    ttt2 = ttt2 - jall
                    o4 = 4

                    If CheckBox2.Checked = False Then
                        PtoIn(0) = 900 : PtoIn(1) = tt2 : PtoIn(2) = 0
                        PtoFin(0) = 582.75 : PtoFin(1) = tt2 : PtoFin(2) = 0
                        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                        line.color = 6
                    Else
                        'se selecciona la casilla "principal" en la interfacee 
                        PtoIn(0) = 900 : PtoIn(1) = tt2 : PtoIn(2) = 0
                        PtoFin(0) = 563.5888 : PtoFin(1) = tt2 : PtoFin(2) = 0
                        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                        line.color = 6

                        PtoIn(0) = 900 : PtoIn(1) = ttt2 : PtoIn(2) = 0
                        PtoFin(0) = 582.75 : PtoFin(1) = ttt2 : PtoFin(2) = 0
                        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                        line.color = 6

                    End If
                    'en este parte de esta opcion se calcula los (amp) en el programa
repit3:
                    va = v_definitivo(movera)
                    exe.Cells(gg, 4) = va
                    suma_monofasico_trifasico = suma_monofasico_trifasico + va 'sumatoria de todas las cargas va en monofasico-trifasico
                    amp = (va) / (Math.Sqrt(3) * 208) 'ecuacion q es utilizada para calcaular los amper(alumbrado,tomacorriente,(1))

                    Dim texdsd

                    texdsd = tt2 'distancia en el momento q esta parado( inicio ejemplo tt1= tt1 - jab) en ese momento
                    insPoint(0) = 585.42    'Set insertion point x coordinate
                    insPoint(1) = (texdsd + 338) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    If (amp > 0 And amp <= 20) Then
                        textStr = ("20A-3P")
                    ElseIf (amp > 20 And amp <= 30) Then
                        textStr = "30A-3P"
                    ElseIf (amp > 30 And amp <= 50) Then
                        textStr = "50A-3P"
                    ElseIf (amp > 50 And amp <= 60) Then
                        textStr = "60A-3P"
                    ElseIf (amp > 60 And amp <= 80) Then
                        textStr = "80A-3P"
                    ElseIf (amp > 80 And amp <= 100) Then
                        textStr = "100A-3P"
                    ElseIf (amp > 100 And amp <= 150) Then
                        textStr = "150A-3P"
                    ElseIf (amp > 150 And amp <= 175) Then
                        textStr = "175A-3P"
                    ElseIf (amp > 175 And amp <= 200) Then
                        textStr = "200A-3P"
                    ElseIf (amp > 200 And amp <= 225) Then
                        textStr = "225A-3P"
                    ElseIf (amp > 225 And amp <= 250) Then
                        textStr = "250A-3P"
                    ElseIf (amp > 250 And amp <= 275) Then
                        textStr = "275A-3P"
                    ElseIf (amp > 275 And amp <= 300) Then
                        textStr = "300A-3P"
                    ElseIf (amp > 300 And amp <= 350) Then
                        textStr = "350A-3P"
                    ElseIf (amp > 350 And amp <= 400) Then
                        textStr = "400A-3P"
                    ElseIf (amp > 400 And amp <= 450) Then
                        textStr = "450A-3P"

                    Else
                        GoTo repit3
                    End If
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight) 'Create Text object
                    text.color = 2
                    exe.Cells(gg, 6) = UCase(textStr)

                    If RadioButton1.Checked = True Then
                        If KTT = 1 Then
                            exe.Cells(gg, 7) = "A,B,C"
                        ElseIf KTT = 2 Then
                            exe.Cells(gg, 7) = "B,C,A"
                        Else
                            exe.Cells(gg, 7) = "C,A,B"
                        End If
                    End If


                    For iii = 1 To 3
                        tt12 = tt12 - 150 ' arco de las fases NOTA: mejorar la creacion del arco..
                        centro(0) = 785.68 : centro(1) = tt12 : centro(2) = 0 'arco
                        radio = 33.88
                        anginic = 0 ' angulos trabajan contra las abujas del relog
                        angfinal = -9.4012
                        arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
                        arco.color = 6
                    Next iii

                    PtoIn(0) = 785.69 : PtoIn(1) = tt12 + 333.88 : PtoIn(2) = 0 '( coloca la barra vertical entre los acos monofasicos )
                    PtoFin(0) = 785.69 : PtoFin(1) = tt2 + 68.22 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6

                    Dim nombreee, tesxx

                    tesxx = tt2
                    nombreee = v_definombre(movera)
                    movera = movera + 1
                    exe.Cells(gg, 8) = UCase(nombreee)
                    If CheckBox1.Checked = True Then ' condicion q si se escoje mayuscula los caracteres salgan en mayuscula
                        nombreee = UCase(nombreee)
                    Else
                    End If

                    insPoint(0) = 1000   'Set insertion point x coordinate
                    insPoint(1) = (tesxx + 184.34) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    textStr = nombreee       'Set the text string
                    'Create Text object
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                    text.color = 2

                ElseIf kk = 1 Then
                    exe.Cells(gg, 3) = "X"
                    exe.Cells(gg, 1) = increment

                    tt2 = tt2 - jab
                    PtoIn(0) = 900 : PtoIn(1) = tt2 : PtoIn(2) = 0
                    PtoFin(0) = 563.5888 : PtoFin(1) = tt2 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6

                    ttt2 = ttt2 - jabb
                    PtoIn(0) = 900 : PtoIn(1) = ttt2 : PtoIn(2) = 0
                    PtoFin(0) = 582.75 : PtoFin(1) = ttt2 : PtoFin(2) = 0
                    line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
                    line.color = 6
                    'en este parte de esta opcion se calcula los (amp) en el programa
repit8:
                    va = v_definitivo(movera)
                    movera = movera + 1
                    exe.Cells(gg, 4) = va
                    suma_tomacorriente = suma_tomacorriente + va 'acumula cargas va de tomacorriente y tomacorriente personalizado
                    amp = (va) / (110) 'ecuacion q es utilizada para calcaular los amper(alumbrado,tomacorriente,(1))

                    Dim textyt
                    textyt = tt2 'distancia en el momento q esta parado( inicio ejemplo tt1= tt1 - jab) en ese momento
                    insPoint(0) = 585.42 'Set insertion point x coordinate
                    insPoint(1) = (textyt + 38) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    If (amp > 0 And amp <= 20) Then
                        textStr = ("20A-1P")
                    ElseIf (amp > 20 And amp <= 30) Then
                        textStr = "30A-1P"
                    ElseIf (amp > 30 And amp <= 50) Then
                        textStr = "50A-1P"
                    ElseIf (amp > 50 And amp <= 60) Then
                        textStr = "60A-1P"
                    ElseIf (amp > 60 And amp <= 80) Then
                        textStr = "80A-1P"
                    ElseIf (amp > 80 And amp <= 100) Then
                        textStr = "100A-1P"
                    ElseIf (amp > 100 And amp <= 150) Then
                        textStr = "150A-1P"
                    ElseIf (amp > 150 And amp <= 175) Then
                        textStr = "175A-1P"
                    ElseIf (amp > 175 And amp <= 200) Then
                        textStr = "200A-1P"
                    ElseIf (amp > 200 And amp <= 225) Then
                        textStr = "225A-1P"
                    ElseIf (amp > 225 And amp <= 250) Then
                        textStr = "250A-1P"
                    ElseIf (amp > 250 And amp <= 275) Then
                        textStr = "275A-1P"
                    ElseIf (amp > 275 And amp <= 300) Then
                        textStr = "300A-1P"
                    ElseIf (amp > 300 And amp <= 350) Then
                        textStr = "350A-1P"
                    ElseIf (amp > 350 And amp <= 400) Then
                        textStr = "400A-1P"
                    ElseIf (amp > 400 And amp <= 450) Then
                        textStr = "450A-1P"

                    Else
                        GoTo repit8
                    End If
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight) 'Create Text object
                    text.color = 2
                    exe.Cells(gg, 6) = UCase(textStr)

                    If RadioButton1.Checked = True Then
                        If KTT = 1 Then
                            exe.Cells(gg, 7) = "A"
                        ElseIf KTT = 2 Then
                            exe.Cells(gg, 7) = "B"
                        Else
                            exe.Cells(gg, 7) = "C"
                        End If
                    End If

                    If RadioButton2.Checked = True Then
                        If KTT2 = 1 Then
                            exe.Cells(gg, 7) = "A"
                        ElseIf KTT2 = 2 Then
                            exe.Cells(gg, 7) = "B"
                        End If
                    End If

                    tt12 = tt12 - 150 ' arco de las fases NOTA: mejorar la creacion del arco..
                    centro(0) = 785.68 : centro(1) = tt12 : centro(2) = 0 'arco
                    radio = 33.88
                    anginic = 0 ' angulos trabajan contra las abujas del relog
                    angfinal = -9.4012
                    arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
                    arco.color = 6

                    Dim tesxtuk
                    tesxtuk = tt2
                    insPoint(0) = 1000    'Set insertion point x coordinate
                    insPoint(1) = (tesxtuk + 35.26) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    If CheckBox1.Checked = True Then ' condicion q si se escoje mayuscula los caracteres salgan en mayuscula
                        textStr = "TOMACORRIENTE"     'Set the text string
                    Else
                        textStr = "tomacorriente"     'Set the text string
                    End If
                    'Create Text object
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                    text.color = 2
                    exe.Cells(gg, 8) = UCase(textStr)
                ElseIf kk = 5 Then

                    exe.Cells(gg, 1) = increment

                    tt2 = tt2 - jab
                    ttt2 = ttt2 - jabb
                    tt12 = tt12 - 150 ' reserva no lleva arco pero de todas formas es utilizado para ovilizar si este esta presente antes de un final

                    Dim tesxtukl
                    tesxtukl = tt2
                    insPoint(0) = 1000    'Set insertion point x coordinate
                    insPoint(1) = (tesxtukl + 35.26) 'Set insertion point y coordinate
                    insPoint(2) = 0      'Set insertion point z coordinate
                    textHeight = 25                'Set text height to 1.0
                    If CheckBox1.Checked = True Then ' condicion q si se escoje mayuscula los caracteres salgan en mayuscula
                        textStr = "RESERVA"     'Set the text string
                    Else
                        textStr = "reserva"
                    End If
                    exe.Cells(gg, 8) = UCase(textStr)
                    'Create Text object
                    text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
                    text.color = 2

                ElseIf kk = 8 Then
                    Exit For

                Else
                    tt2 = tt2 - jab
                    ttt2 = ttt2 - jabb

                End If
nmm:            ' si el primero es 4
nmk:            ' si el segundo es 4

                'PROCESO PARA PONER A,B,C EN EXCEL
                '*******************************************
                If RadioButton2.Checked = True Then
                    'CUANDO ES MONOFASICO
                    If KTT2 = 2 Then
                        KTT2 = 1
                        GoTo REINICIO
                    End If
                End If

                If RadioButton1.Checked = True Then
                    'CAUNDO ES TRIFASICO
                    If KTT = 3 Then
                        KTT = 1
                        GoTo REINICIO
                    End If
                End If
                KTT = KTT + 1
                KTT2 = KTT2 + 1
REINICIO:
            Next  'FINAL DEL CILCLO DEL ALGORITMO Q SABE CUANDO UN TRIFAASICO O MONOFACICO OUCAPA UN CIRCUITO!!
        '******************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
final:
        'instruccion q hace parte del esqueleto del tablero
        'barras dobles derecha e izquierda

        PtoIn(0) = 336.4112 : PtoIn(1) = (lon - 184.3425) : PtoIn(2) = 0 ' nota: en esta parte del programa  se formula para q la barra vertical de neutro siempre quede exacto con su barra horizontal y a su vez quede bn segun la longuitod q s eles de a las primeras tres barras verticales pricipales
        PtoFin(0) = 336.4112 : PtoFin(1) = (-lon + (lon - 97.8436)) : PtoFin(2) = 0
        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
        line.color = 6
        If l1 = 1 Then
            PtoIn(0) = 317.25 : PtoIn(1) = (lon - 355.95) : PtoIn(2) = 0
            PtoFin(0) = 317.25 : PtoFin(1) = -121.7616 : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6
        Else
            PtoIn(0) = 317.25 : PtoIn(1) = (lon - 206.25) : PtoIn(2) = 0
            PtoFin(0) = 317.25 : PtoFin(1) = -121.7616 : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6
        End If


        'SEGUNDA PARTE DEL TABLERO
        PtoIn(0) = 563.5888 : PtoIn(1) = (lon - 184.3425) : PtoIn(2) = 0 ' nota: en esta parte del programa  se formula para q la barra vertical de neutro siempre quede exacto con su barra horizontal y a su vez quede bn segun la longuitod q s eles de a las primeras tres barras verticales pricipales
        PtoFin(0) = 563.5888 : PtoFin(1) = (-lon + (lon - 97.8436)) : PtoFin(2) = 0
        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
        line.color = 6
        If l2 = 2 Then

            PtoIn(0) = 582.75 : PtoIn(1) = (lon - 335.95) : PtoIn(2) = 0
            PtoFin(0) = 582.75 : PtoFin(1) = -121.7616 : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6
        Else
            PtoIn(0) = 582.75 : PtoIn(1) = (lon - 206.25) : PtoIn(2) = 0
            PtoFin(0) = 582.75 : PtoFin(1) = -121.7616 : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6
        End If

        'lineas horizontales de neutro y tierra
        PtoIn(0) = 563.5888 : PtoIn(1) = (-lon + (lon - 97.8436)) : PtoIn(2) = 0
        PtoFin(0) = 336.4112 : PtoFin(1) = (-lon + (lon - 97.8436)) : PtoFin(2) = 0
        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
        line.color = 6

        PtoIn(0) = 582.75 : PtoIn(1) = (-lon + (lon - 121.7616)) : PtoIn(2) = 0
        PtoFin(0) = 0 : PtoFin(1) = (-lon + (lon - 121.7616)) : PtoFin(2) = 0
        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
        line.color = 6

        '*************************************************************************************************
        'CAJA COMPLETA DE NEUTRO Y PALABRA ("NUETRO")
        '****************************************************************************************************
        'de la linea vertical del neutro hasta toda la caja de neutro
        'linea vertical
        PtoIn(0) = 450 : PtoIn(1) = (-lon + (lon - 204.0488)) : PtoIn(2) = 0
        PtoFin(0) = 450 : PtoFin(1) = -97.8436 : PtoFin(2) = 0
        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
        line.color = 6

        'barra horizontal de caja de neutro
        PtoIn(0) = 527.8696 : PtoIn(1) = (-lon + (lon - 204.0488)) : PtoIn(2) = 0
        PtoFin(0) = 372.3104 : PtoFin(1) = (-lon + (lon - 204.0488)) : PtoFin(2) = 0
        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
        line.color = 6

        If CheckBox2.Checked = False Then 'SI  ES UN  SUBLTABLERO.
            PtoIn(0) = 527.8696 : PtoIn(1) = (-lon + (lon - 250.0502)) : PtoIn(2) = 0
            PtoFin(0) = 0 : PtoFin(1) = (-lon + (lon - 250.0502)) : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6
        Else
            'SI ES UN TABLERO PRINCIPAL..
            PtoIn(0) = 527.8696 : PtoIn(1) = (-lon + (lon - 250.0502)) : PtoIn(2) = 0
            PtoFin(0) = 372.3104 : PtoFin(1) = (-lon + (lon - 250.0502)) : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6
        End If

        insPoint(0) = 378  'Set insertion point x coordinate
        insPoint(1) = (-lon + (lon - 240.0502)) 'Set insertion point y coordinate
        insPoint(2) = 0      'Set insertion point z coordinate
        textHeight = 25                'Set text height to 1.0
        textStr = "NEUTRO"     'Set the text string
        'Create Text object
        text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
        text.color = 2

        ' barras verticales de caja de neutro
        PtoIn(0) = 372.134 : PtoIn(1) = (-lon + (lon - 250.0502)) : PtoIn(2) = 0
        PtoFin(0) = 372.134 : PtoFin(1) = -204.1113 : PtoFin(2) = 0
        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
        line.color = 6

        PtoIn(0) = 527.8696 : PtoIn(1) = (-lon + (lon - 250.0502)) : PtoIn(2) = 0
        PtoFin(0) = 527.8696 : PtoFin(1) = -204.1113 : PtoFin(2) = 0
        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
        line.color = 6
        '**************************************************************************************************************
        'FACTORES DE DEMANDA DE TMACORRIENTES,ILUMINACION Y FUERZA(MONOFASICO Y TRIFASICO)
        On Error Resume Next
p41:
        factor_demanda_tomacorriente = InputBox("ingrese factor de demanda de tomacorriente", "tableros")
        'corrector de errores.. si se introduce un tipo de variable q no relaciona con la vaibale repite l apregunta
        If Err.Number <> 0 Then
            Err.Clear()
            MsgBox("error ha introducido un valor caracter", MsgBoxStyle.Critical)
            GoTo p41
        End If
        On Error Resume Next
p31:
        factor_demanda_iluminacion = InputBox("ingrese factor de demanda de iluminacion", "tablaero")
        'corrector de errores.. si se introduce un tipo de variable q no relaciona con la vaibale repite l apregunta
        If Err.Number <> 0 Then
            Err.Clear()
            MsgBox("error ha introducido un valor caracter", MsgBoxStyle.Critical)
            GoTo p31
        End If
        On Error Resume Next
p21:
        factor_demanda_fuerza = InputBox("ingrese factor de demanda de fuerza", "tablero")
        'corrector de errores.. si se introduce un tipo de variable q no relaciona con la vaibale repite l apregunta
        If Err.Number <> 0 Then
            Err.Clear()
            MsgBox("error ha introducido un valor caracter", MsgBoxStyle.Critical)
            GoTo p21
        End If

        ' DEMANDAS DE TOMACORRIENTE,ILUMINACION Y FUERZA(MONOFASICO Y TRIFASICO)
        demanda_tomacorriente = (factor_demanda_tomacorriente / 100) * suma_tomacorriente
        exe.Cells(7, 3) = suma_tomacorriente & " x " & factor_demanda_tomacorriente & "%" & " = " & demanda_tomacorriente

        demanda_ilumiinacion = (factor_demanda_iluminacion / 100) * suma_iluminacion
        exe.Cells(8, 3) = suma_iluminacion & " x " & factor_demanda_iluminacion & "%" & " = " & demanda_ilumiinacion

        demanda_fuerza = (factor_demanda_fuerza / 100) * suma_monofasico_trifasico
        exe.Cells(9, 3) = suma_monofasico_trifasico & " x " & factor_demanda_fuerza & "%" & " = " & demanda_fuerza

        demanda_total = demanda_fuerza + demanda_ilumiinacion + demanda_tomacorriente
        exe.Cells(3, 8) = demanda_total & " VA"

        '**************************************************************************
        Dim KVA, CD, longitud, NAMP, NCD, j, MAC, R, X, DV, maxDV, INC As Double 'variables utilizadas para calculo de tablero y tipo de cable a utilizar!!

        On Error Resume Next
p61:
        maxDV = InputBox("ingrese maxima caida de tension permitida")
        'corrector de errores.. si se introduce un tipo de variable q no relaciona con la vaibale repite l apregunta
        If Err.Number <> 0 Then
            Err.Clear()
            MsgBox("error ha introducido un valor caracter", MsgBoxStyle.Critical)
            GoTo p61
        End If

        On Error Resume Next
p51:
        longitud = InputBox("ingrese longitud de alimentador")
        'corrector de errores.. si se introduce un tipo de variable q no relaciona con la vaibale repite l apregunta

        If Err.Number <> 0 Then
            Err.Clear()
            MsgBox("error ha introducido un valor caracter", MsgBoxStyle.Critical)
            GoTo p51
        End If

        KVA = demanda_total / 1000
        CD = longitud * KVA


        If RadioButton3.Checked = True Then
            KV = 0.208
        ElseIf RadioButton4.Checked = True Then
            KV = 0.24
        End If

        'se calcula la corriente total de la demanda..
        If RadioButton2.Checked = True Then

            I_total = KVA / KV

        ElseIf RadioButton1.Checked = True Then

            I_total = KVA / (Math.Sqrt(3) * KV)

        End If


        'luego criterio para selecion de cable neutro del alimentador...
        If RadioButton1.Checked = True And RadioButton3.Checked = True Then
            NAMP = 0.5 * I_total : NCD = 0.5 * CD

        ElseIf RadioButton2.Checked = True And RadioButton4.Checked = True And (I_total >= 200) Then
            NAMP = 0.7 * I_total : NCD = 0.7 * CD

        ElseIf RadioButton2.Checked = True And RadioButton4.Checked = True And (I_total < 200) Then
            NAMP = I_total : NCD = CD

        End If


        j = 1
        If (0 < NAMP) And (NAMP <= 25) Then
            MAC = 25

        ElseIf (25 < NAMP) And (NAMP <= 30) Then
            MAC = 30
            j = j + 2

        ElseIf (30 < NAMP) And (NAMP <= 50) Then
            MAC = 50
            j = j + 4

        ElseIf (50 < NAMP) And (NAMP <= 65) Then
            MAC = 65
            j = j + 6

        ElseIf (65 < NAMP) And (NAMP <= 85) Then
            MAC = 85
            j = j + 8

        ElseIf (85 < NAMP) And (NAMP <= 115) Then
            MAC = 115
            j = j + 10

        ElseIf (115 < NAMP) And (NAMP <= 150) Then
            MAC = 150
            j = j + 12

        ElseIf (150 < NAMP) And (NAMP <= 175) Then
            MAC = 175
            j = j + 14

        ElseIf (175 < NAMP) And (NAMP <= 200) Then
            MAC = 200
            j = j + 16

        ElseIf (200 < NAMP) And (NAMP <= 230) Then
            MAC = 230
            j = j + 18

        ElseIf (230 < NAMP) And (NAMP <= 255) Then
            MAC = 255
            j = j + 20

        ElseIf (255 < NAMP) And (NAMP <= 285) Then
            MAC = 285
            j = j + 22
            INC = 285
        ElseIf (285 < NAMP) And (NAMP <= 310) Then
            MAC = 310
            j = j + 24

        ElseIf (310 < NAMP) And (NAMP <= 335) Then
            MAC = 335
            j = j + 26

        ElseIf (335 < NAMP) And (NAMP <= 380) Then
            MAC = 380
            j = j + 28

        ElseIf (380 < NAMP) And (NAMP <= 420) Then
            MAC = 420
            j = j + 30

        ElseIf (420 < NAMP) And (NAMP <= 460) Then
            MAC = 460
            j = j + 32

        ElseIf (460 < NAMP) And (NAMP <= 475) Then
            MAC = 475
            j = j + 34

        ElseIf NAMP > 475 Then
            MsgBox("Se sugiere utilizar Conductores en Paralelo")
        End If
denuevox:
        If j = 1 Then
            R = 1968 : X = 58.4 : MAC = 25
            INC = 25
        ElseIf j = 3 Then
            R = 1230 : X = 56.4 : MAC = 30
            INC = 30
        ElseIf j = 5 Then
            R = 789 : X = 55.3 : MAC = 50
            INC = 50
        ElseIf j = 7 Then
            R = 490 : X = 51.2 : MAC = 65
            INC = 65
        ElseIf j = 9 Then
            R = 318 : X = 47.3 : MAC = 85
            INC = 85
        ElseIf j = 11 Then
            R = 203 : X = 43.8 : MAC = 115
            INC = 115
        ElseIf j = 13 Then
            R = 129 : X = 41.5 : MAC = 150
            INC = 150
        ElseIf j = 15 Then
            R = 103 : X = 40.9 : MAC = 175
            INC = 175
        ElseIf j = 17 Then
            R = 80.3 : X = 40.2 : MAC = 200
            INC = 200
        ElseIf j = 19 Then
            R = 66.6 : X = 39.1 : MAC = 230
            INC = 230
        ElseIf j = 21 Then
            R = 57.8 : X = 39 : MAC = 255
            INC = 255
        ElseIf j = 23 Then
            R = 50.1 : X = 38.7 : MAC = 285
            INC = 285
        ElseIf j = 25 Then
            R = 38 : X = 38.4 : MAC = 310
            INC = 310
        ElseIf j = 27 Then
            R = 35.6 : X = 38.1 : MAC = 335
            INC = 335
        ElseIf j = 29 Then
            R = 27.5 : X = 36.6 : MAC = 380
            INC = 380
        ElseIf j = 31 Then
            R = 24.1 : X = 36.4 : MAC = 420
            INC = 420
        ElseIf j = 33 Then
            R = 24.7 : X = 35.8 : MAC = 460
            INC = 460
        ElseIf j = 35 Then
            R = 19.8 : X = 35.3 : MAC = 475
            INC = 475
        ElseIf j > 35 Then
            MsgBox("Se sugiere utilizar Cables en Paralelo")

        End If

        If RadioButton1.Checked = True Then
            DV = (NCD * (R * (fp / 100) + X * Math.Sqrt(1 - ((fp / 100) ^ 2))) / (10 * (KV ^ 2))) * (0.0328083989 * 10 ^ -4)

        ElseIf RadioButton2.Checked = True Then
            DV = 2 * (NCD * (R * (fp / 100) + X * Math.Sqrt(1 - ((fp / 100) ^ 2))) / (10 * (KV ^ 2))) * (0.0328083989 * 10 ^ -4)
        End If


        If DV > maxDV Then
            j = j + 2 : GoTo denuevox

        End If
        exe.Cells(10, 7) = DV
        exe.Cells(8, 7) = (NAMP / INC) * 100 & "%"

        insPoint(0) = 650    'Set insertion point x coordinate
        insPoint(1) = lon - (lon + 400) 'Set insertion point y coordinate
        insPoint(2) = 0      'Set insertion point z coordinate
        textHeight = 25                'Set text height to 1.0
        If j = 1 Then
            textStr = "+ 1 CABLE THW  #12 AWG DE COBRE(NEUTRO)"
        ElseIf j = 3 Then
            textStr = ("+ 1 CABLE  THW #10 AWG DE COBRE(NEUTRO)")
        ElseIf j = 5 Then
            textStr = ("+ 1 CABLE  THW  #8  AWG DE COBRE(NEUTRO)")
        ElseIf j = 7 Then
            textStr = ("+ 1 CABLE THW  #6  AWG DE COBRE(NEUTRO)")
        ElseIf j = 9 Then
            textStr = ("+ 1 CABLE THW  #4  AWG DE COBRE(NEUTRO)")
        ElseIf j = 11 Then
            textStr = ("+ 1 CABLE  THW  #2  AWG DE COBRE(NEUTRO)")
        ElseIf j = 13 Then
            textStr = ("+ 1 CABLE  THW  #1/0 AWG DE COBRE(NEUTRO)")
        ElseIf j = 15 Then
            textStr = ("+ 1 CABLE  THW  #2/0 AWG DE COBRE(NEUTRO)")
        ElseIf j = 17 Then
            textStr = ("+ 1 CABLE  THW  #3/0 AWG DE COBRE(NEUTRO)")
        ElseIf j = 19 Then
            textStr = ("+ 1 CABLE  THW  #4/0 AWG DE COBRE(NEUTRO)")
        ElseIf j = 21 Then
            textStr = ("+ 1 CABLE  THW  #250 MCM DE COBRE(NEUTRO)")
        ElseIf j = 23 Then
            textStr = ("+ 1 CABLE  THW  #300 MCM DE COBRE(NEUTRO)")
        ElseIf j = 25 Then
            textStr = ("+ 1 CABLE  THW  #350 MCM DE COBRE(NEUTRO)")
        ElseIf j = 27 Then
            textStr = ("+ 1 CABLE  THW #400 MCM DE COBRE(NEUTRO)")
        ElseIf j = 29 Then
            textStr = ("+ 1 CABLE  THW #500 MCM DE COBRE(NEUTRO)")
        ElseIf j = 31 Then
            textStr = ("+ 1 CABLE  THW #600 MCM DE COBRE(NEUTRO)")
        ElseIf j = 33 Then
            textStr = ("+ 1 CABLE  THW #700 MCM DE COBRE(NEUTRO)")
        ElseIf j = 35 Then
            textStr = ("+ 1 CABLE  THW #750 MCM DE COBRE(NEUTRO)")
        End If
        text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
        text.color = 2

        exe.Cells(7, 8) = textStr


        i = 1
        If (0 < I_total) And (I_total <= 25) Then
            MAC = 20

        ElseIf (25 < I_total) And (I_total <= 30) Then
            MAC = 30
            i = i + 2

        ElseIf (30 < I_total) And (I_total <= 50) Then
            MAC = 50
            i = i + 4

        ElseIf (50 < I_total) And (I_total <= 65) Then
            MAC = 60
            i = i + 6

        ElseIf (65 < I_total) And (I_total <= 85) Then
            MAC = 80
            i = i + 8

        ElseIf (85 < I_total) And (I_total <= 115) Then
            MAC = 100
            i = i + 10

        ElseIf (115 < I_total) And (I_total <= 150) Then
            MAC = 150
            i = i + 12

        ElseIf (150 < I_total) And (I_total <= 175) Then
            MAC = 175
            i = i + 14

        ElseIf (175 < I_total) And (I_total <= 200) Then
            MAC = 200
            i = i + 16

        ElseIf (200 < I_total) And (I_total <= 230) Then
            MAC = 225
            i = i + 18

        ElseIf (230 < I_total) And (I_total <= 255) Then
            MAC = 250
            i = i + 20

        ElseIf (255 < I_total) And (I_total <= 285) Then
            MAC = 275
            i = i + 22

        ElseIf (285 < I_total) And (I_total <= 310) Then
            MAC = 300
            i = i + 24

        ElseIf (310 < I_total) And (I_total <= 335) Then
            MAC = 300
            i = i + 26

        ElseIf (335 < I_total) And (I_total <= 380) Then
            MAC = 350
            i = i + 28

        ElseIf (380 < I_total) And (I_total <= 420) Then
            MAC = 400
            i = i + 30

        ElseIf (420 < I_total) And (I_total <= 460) Then
            MAC = 450
            i = i + 32

        ElseIf (460 < I_total) And (I_total <= 475) Then
            MAC = 450
            i = i + 34

        ElseIf I_total > 475 Then
            MsgBox("Se sugiere utilizar Conductores en Paralelo")
        End If
        Dim numero(40), p As String
        Dim l_caracter As String
        p = 0
denuevoo:
        'calculo de fases..
        If i = 1 Then

            R = 1968 : X = 58.4 : MAC = 20
            numero(1) = " 12 AWG"
            INC = 25
        ElseIf i = 3 Then
            R = 1230 : X = 56.4
            If p = 0 Then
                MAC = 30
            End If
            numero(3) = "10 AWG"
            INC = 30
        ElseIf i = 5 Then
            R = 789 : X = 55.3
            If p = 0 Then
                MAC = 50
            End If
            numero(5) = "8 AWG"
            INC = 50
        ElseIf i = 7 Then
            R = 490 : X = 51.2
            If p = 0 Then
                MAC = 60
            End If
            numero(7) = "6 AWG"
            INC = 65
        ElseIf i = 9 Then
            R = 318 : X = 47.3
            If p = 0 Then
                MAC = 80
            End If
            numero(9) = "4 AWG"
            INC = 85
        ElseIf i = 11 Then
            R = 203 : X = 43.8
            If p = 0 Then
                MAC = 100
            End If
            numero(11) = "2 AWG"
            INC = 115
        ElseIf i = 13 Then
            R = 129 : X = 41.5
            If p = 0 Then
                MAC = 150
            End If
            numero(13) = "1/0 AWG"
            INC = 150
        ElseIf i = 15 Then
            R = 103 : X = 40.9
            If p = 0 Then
                MAC = 175
            End If
            numero(15) = "2/0 AWG"
            INC = 175
        ElseIf i = 17 Then
            R = 80.3 : X = 40.2
            If p = 0 Then
                MAC = 200
            End If
            numero(17) = "3/0 AWG"
            INC = 200
        ElseIf i = 19 Then
            R = 66.6 : X = 39.1
            If p = 0 Then
                MAC = 225
            End If
            numero(19) = "4/0 AWG"
            INC = 230
        ElseIf i = 21 Then
            R = 57.8 : X = 39
            If p = 0 Then
                MAC = 250
            End If
            numero(21) = "250 MCM"
            INC = 255
        ElseIf i = 23 Then
            R = 50.1 : X = 38.7
            If p = 0 Then
                MAC = 275
            End If
            numero(23) = "300 MCM"
            INC = 285
        ElseIf i = 25 Then
            R = 38 : X = 38.4
            If p = 0 Then
                MAC = 300
            End If
            numero(25) = "350 MCM"
            INC = 310
        ElseIf i = 27 Then
            R = 35.6 : X = 38.1
            If p = 0 Then
                MAC = 300
            End If
            numero(27) = "400 MCM"
            INC = 335
        ElseIf i = 29 Then
            If p = 0 Then
                R = 27.5 : X = 36.6
                MAC = 350
            End If
            numero(29) = "500 MCM"
            INC = 380
        ElseIf i = 31 Then
            R = 24.1 : X = 36.4
            If p = 0 Then
                MAC = 400
            End If
            numero(31) = "600 MCM"
            INC = 420
        ElseIf i = 33 Then
            R = 24.7 : X = 35.8
            If p = 0 Then
                MAC = 450
            End If
            numero(33) = "700 MCM"
            INC = 460
        ElseIf i = 35 Then
            R = 19.8 : X = 35.3
            If p = 0 Then
                MAC = 450
            End If
            numero(35) = "750 MCM"
            INC = 475
        ElseIf i > 35 Then
            MsgBox("Se sugiere utilizar Cables en Paralelo")
        End If


        If RadioButton1.Checked = True Then
            DV = (CD * (R * (fp / 100) + X * Math.Sqrt(1 - ((fp / 100) ^ 2))) / (10 * (KV ^ 2))) * (0.0328083989 * 10 ^ -4)
        ElseIf RadioButton2.Checked = True Then
            DV = 2 * (CD * (R * (fp / 100) + X * Math.Sqrt(1 - ((fp / 100) ^ 2))) / (10 * (KV ^ 2))) * (0.0328083989 * 10 ^ -4)
        End If

        If DV > maxDV Then
            i = i + 2
            p = 1
            GoTo denuevoo
        End If
        exe.Cells(9, 7) = DV

        exe.Cells(7, 7) = (I_total / INC) * 100 & "%"

        If CheckBox2.Checked = False Then
            BP = MAC
        End If


        If RadioButton5.Checked = True Then
            l_caracter = "NLAB"
        ElseIf RadioButton6.Checked = True Then
            l_caracter = "NAB"
        End If

        If RadioButton2.Checked = True Then


            insPoint(0) = -600   'Set insertion point x coordinate
            insPoint(1) = lon - (lon + 500) 'Set insertion point y coordinate
            insPoint(2) = 0      'Set insertion point z coordinate
            textHeight = 25                'Set text height to 1.0
            textStr = "TABLERO  TIPO:  " & l_caracter & "3" & ncircuitos2 & " DE " & ncircuitos2 & " CIRCUITOS, 2 FASES,  CON  BARRAS  DE  NEUTRO Y TIERRA."
            text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
            text.color = 2

            insPoint(0) = -600   'Set insertion point x coordinate
            insPoint(1) = lon - (lon + 400) 'Set insertion point y coordinate
            insPoint(2) = 0      'Set insertion point z coordinate
            textHeight = 25                'Set text height to 1.0
            textStr = ("ALIMENTADOR: 2 CABLES THW  # " & numero(i) & " DE COBRE(FASES)")
            text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
            text.color = 2

            'resultado sin la palabra "alimentaro" para la hoja de excel
            textStr = (" 2 CABLES THW  # " & numero(i) & " DE COBRE(FASES)")
            exe.Cells(6, 8) = textStr

            insPoint(0) = 600   'Set insertion point x coordinate
            insPoint(1) = lon + 50 'Set insertion point y coordinate
            insPoint(2) = 0      'Set insertion point z coordinate
            textHeight = 25                'Set text height to 1.0
            textStr = (BP & "A-2P")
            text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
            text.color = 2
            exe.Cells(5, 8) = textStr

        ElseIf RadioButton1.Checked = True Then
            insPoint(0) = -600    'Set insertion point x coordinate
            insPoint(1) = lon - (lon + 500) 'Set insertion point y coordinate
            insPoint(2) = 0      'Set insertion point z coordinate
            textHeight = 25                'Set text height to 1.0
            textStr = "TABLERO  TIPO:  " & l_caracter & "4" & ncircuitos2 & " DE " & ncircuitos2 & " CIRCUITOS, 3 FASES,  CON  BARRAS  DE  NEUTRO Y TIERRA."
            text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
            text.color = 2

            insPoint(0) = -600    'Set insertion point x coordinate
            insPoint(1) = lon - (lon + 400) 'Set insertion point y coordinate
            insPoint(2) = 0      'Set insertion point z coordinate
            textHeight = 25                'Set text height to 1.0
            textStr = ("ALIMENTADOR: 3 CABLES THW #" & numero(i) & " DE COBRE(FASES)")
            text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
            text.color = 2

            'resultado sin la palabra "alimentaro" para la hoja de excel
            textStr = (" 3 CABLES THW #" & numero(i) & " DE COBRE(FASES)")
            exe.Cells(6, 8) = textStr

            insPoint(0) = 600    'Set insertion point x coordinate
            insPoint(1) = lon + 50  'Set insertion point y coordinate
            insPoint(2) = 0      'Set insertion point z coordinate
            textHeight = 25                'Set text height to 1.0
            textStr = (BP & "A-3P")
            text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
            text.color = 2
            exe.Cells(5, 8) = textStr

        End If


        insPoint(0) = -311    'Set insertion point x coordinate
        insPoint(1) = lon - (lon + 450) 'Set insertion point y coordinate
        insPoint(2) = 0      'Set insertion point z coordinate
        textHeight = 25                'Set text height to 1.0
        If 0 < BP And BP <= 20 Then
            textStr = "+ THW # 12 AWG COBRE(TIERRA) EN TUBERIA Ø    PVC."
        ElseIf 20 < BP And BP <= 60 Then
            textStr = "+ THW # 10 AWG COBRE(TIERRA) EN TUBERIA Ø    PVC."
        ElseIf 60 < BP And BP <= 100 Then
            textStr = "+ THW # 8  AWG COBRE(TIERRA) EN TUBERIA Ø    PVC."
        ElseIf 100 < BP And BP <= 200 Then
            textStr = "+ THW # 6  AWG COBRE(TIERRA) EN TUBERIA Ø    PVC."
        ElseIf 200 < BP And BP <= 300 Then
            textStr = "+ THW # 4  AWG COBRE(TIERRA) EN TUBERIA Ø    PVC."
        ElseIf 300 < BP And BP <= 500 Then
            textStr = "+ THW # 2  AWG COBRE(TIERRA) EN TUBERIA Ø    PVC."
        ElseIf 500 < BP And BP <= 800 Then
            textStr = "+ THW # 1/0 AWG COBRE(TIERRA) EN TUBERIA Ø    PVC."
        ElseIf 800 < BP And BP <= 1000 Then
            textStr = "+ THW # 2/0 AWG COBRE(TIERRA) EN TUBERIA Ø    PVC."
        ElseIf 1000 < BP And BP <= 1200 Then
            textStr = "+ THW # 3/0 AWG COBRE(TIERRA) EN TUBERIA Ø    PVC."
        ElseIf 1200 < BP And BP <= 1600 Then
            textStr = "+ THW # 4/0 AWG COBRE(TIERRA) EN TUBERIA Ø    PVC."
        ElseIf 1600 < BP And BP <= 2000 Then
            textStr = "+ THW # 250 MCM COBRE(TIERRA) EN TUBERIA Ø    PVC."
        ElseIf 2000 < BP And BP <= 2500 Then
            textStr = "+ THW # 350 MCM  AWG COBRE(TIERRA) EN TUBERIA Ø    PVC."
        ElseIf 2500 < BP And BP <= 3000 Then
            textStr = "+ THW # 400 MCM COBRE(TIERRA) EN TUBERIA Ø    PVC."
        ElseIf 3000 < BP And BP <= 4000 Then
            textStr = "+ THW # 500 MCM COBRE(TIERRA) EN TUBERIA Ø    PVC."
        ElseIf 4000 < BP And BP <= 5000 Then
            textStr = "+ THW # 700 MCM COBRE(TIERRA) EN TUBERIA Ø    PVC."
        ElseIf 5000 < BP And BP <= 6000 Then
            textStr = "+ THW # 800 MCM COBRE(TIERRA) EN TUBERIA Ø    PVC."
        End If
        text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
        text.color = 2
        exe.Cells(8, 8) = textStr

        If CheckBox2.Checked = False Then
            insPoint(0) = -550  'Set insertion point x coordinate
            insPoint(1) = lon - (lon + 110) 'Set insertion point y coordinate
            insPoint(2) = 0      'Set insertion point z coordinate
            textHeight = 25                'Set text height to 1.0
            textStr = "A BARRA DE TIERRA EN TABLERO PRINCIPAL"
            text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
            text.color = 2

            insPoint(0) = -550  'Set insertion point x coordinate
            insPoint(1) = lon - (lon + 240) 'Set insertion point y coordinate
            insPoint(2) = 0      'Set insertion point z coordinate
            textHeight = 25                'Set text height to 1.0
            textStr = "A BARRA DE NEUTRO EN TABLERO PRINCIPAL"
            text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
            text.color = 2
        Else
            'barrita vertical de la figura triangular cuando es principal 
            PtoIn(0) = 0 : PtoIn(1) = (lon - (lon + 170.25)) : PtoIn(2) = 0
            PtoFin(0) = 0 : PtoFin(1) = (lon - (lon + 121.7616)) : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6

            'primera barra
            PtoIn(0) = 118.5625 : PtoIn(1) = (lon - (lon + 170.25)) : PtoIn(2) = 0
            PtoFin(0) = 0 : PtoFin(1) = (lon - (lon + 170.25)) : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6

            PtoIn(0) = -118.5625 : PtoIn(1) = (lon - (lon + 170.25)) : PtoIn(2) = 0
            PtoFin(0) = 0 : PtoFin(1) = (lon - (lon + 170.25)) : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6

            'segunda barra

            PtoIn(0) = 94.85 : PtoIn(1) = (lon - (lon + 170.25 + 14.5)) : PtoIn(2) = 0
            PtoFin(0) = 0 : PtoFin(1) = (lon - (lon + 170.25 + 14.5)) : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6

            PtoIn(0) = -94.85 : PtoIn(1) = (lon - (lon + 170.25 + 14.5)) : PtoIn(2) = 0
            PtoFin(0) = 0 : PtoFin(1) = (lon - (lon + 170.25 + 14.5)) : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6

            'tercera barrita
            PtoIn(0) = 71.1375 : PtoIn(1) = (lon - (lon + 170.25 + (14.5 * 2))) : PtoIn(2) = 0
            PtoFin(0) = 0 : PtoFin(1) = (lon - (lon + 170.25 + (14.5 * 2))) : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6

            PtoIn(0) = -71.1375 : PtoIn(1) = (lon - (lon + 170.25 + (14.5 * 2))) : PtoIn(2) = 0
            PtoFin(0) = 0 : PtoFin(1) = (lon - (lon + 170.25 + (14.5 * 2))) : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6

            'cuarta barrita
            PtoIn(0) = 47.425 : PtoIn(1) = (lon - (lon + 170.25 + (14.5 * 3))) : PtoIn(2) = 0
            PtoFin(0) = 0 : PtoFin(1) = (lon - (lon + 170.25 + (14.5 * 3))) : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6

            PtoIn(0) = -47.425 : PtoIn(1) = (lon - (lon + 170.25 + (14.5 * 3))) : PtoIn(2) = 0
            PtoFin(0) = 0 : PtoFin(1) = (lon - (lon + 170.25 + (14.5 * 3))) : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6

            'quinta barrita

            PtoIn(0) = 23.7125 : PtoIn(1) = (lon - (lon + 170.25 + (14.5 * 4))) : PtoIn(2) = 0
            PtoFin(0) = 0 : PtoFin(1) = (lon - (lon + 170.25 + (14.5 * 4))) : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6

            PtoIn(0) = -23.7125 : PtoIn(1) = (lon - (lon + 170.25 + (14.5 * 4))) : PtoIn(2) = 0
            PtoFin(0) = 0 : PtoFin(1) = (lon - (lon + 170.25 + (14.5 * 4))) : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6

            'sexta barrita
            PtoIn(0) = 3.2759 : PtoIn(1) = (lon - (lon + 170.25 + (14.5 * 5))) : PtoIn(2) = 0
            PtoFin(0) = 0 : PtoFin(1) = (lon - (lon + 170.25 + (14.5 * 5))) : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6

            PtoIn(0) = -3.2759 : PtoIn(1) = (lon - (lon + 170.25 + (14.5 * 5))) : PtoIn(2) = 0
            PtoFin(0) = 0 : PtoFin(1) = (lon - (lon + 170.25 + (14.5 * 5))) : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 6

        End If

        'ARCOS Q VAN ARRIBA DEL TABLERO
        centro(0) = 375 : centro(1) = lon + 33.88 : centro(2) = 0 'arco
        radio = 33.88
        anginic = 300 ' angulos trabajan contra  las abujas del relog
        angfinal = -300
        arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
        arco.color = 6

        'si es trifasico....
        If RadioButton1.Checked = True Then
            centro(0) = 375 + 75 : centro(1) = lon + 33.88 : centro(2) = 0 'arco
            radio = 33.88
            anginic = 300 ' angulos trabajan contra  las abujas del relog
            angfinal = -300
            arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
            arco.color = 6
        End If

        centro(0) = 375 + 75 * 2 : centro(1) = lon + 33.88 : centro(2) = 0 'arco
        radio = 33.88
        anginic = 300 ' angulos trabajan contra  las abujas del relog
        angfinal = -300
        arco = objautocad.ActiveDocument.ModelSpace.AddArc(centro, radio, anginic, angfinal)
        arco.color = 6

        'LINEA Q PASA POR LOS TRES ARCOS
        PtoIn(0) = 558.88 : PtoIn(1) = lon + 33.88 : PtoIn(2) = 0
        PtoFin(0) = 408.88 : PtoFin(1) = lon + 33.88 : PtoFin(2) = 0
        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
        line.color = 6

        'lineas q  estan arriba de los arcos..
        PtoIn(0) = 375 : PtoIn(1) = lon + 124 : PtoIn(2) = 0
        PtoFin(0) = 375 : PtoFin(1) = lon + 67.76 : PtoFin(2) = 0
        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
        line.color = 150

        If RadioButton1.Checked = True Then
            PtoIn(0) = 375 + 75 : PtoIn(1) = lon + 124 : PtoIn(2) = 0
            PtoFin(0) = 375 + 75 : PtoFin(1) = lon + 67.76 : PtoFin(2) = 0
            line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
            line.color = 150
        End If

        PtoIn(0) = 375 + 75 * 2 : PtoIn(1) = lon + 124 : PtoIn(2) = 0
        PtoFin(0) = 375 + 75 * 2 : PtoFin(1) = lon + 67.76 : PtoFin(2) = 0
        line = objautocad.ActiveDocument.ModelSpace.AddLine(PtoIn, PtoFin)
        line.color = 150

        insPoint(0) = 365 'Set insertion point x coordinate
        insPoint(1) = lon + 140 'Set insertion point y coordinate
        insPoint(2) = 0      'Set insertion point z coordinate
        textHeight = 25                'Set text height to 1.0
        textStr = "A"
        text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
        text.color = 2

        If RadioButton1.Checked = True Then
            insPoint(0) = 365 + 75 'Set insertion point x coordinate
            insPoint(1) = lon + 140 'Set insertion point y coordinate
            insPoint(2) = 0      'Set insertion point z coordinate
            textHeight = 25                'Set text height to 1.0
            textStr = "B"
            text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
            text.color = 2
        End If

        insPoint(0) = 365 + 75 * 2 'Set insertion point x coordinate
        insPoint(1) = lon + 140 'Set insertion point y coordinate
        insPoint(2) = 0      'Set insertion point z coordinate
        textHeight = 25                'Set text height to 1.0
        If RadioButton1.Checked = True Then
            textStr = "C"
        Else
            textStr = "B"
        End If
        text = objautocad.ActiveDocument.ModelSpace.AddText(textStr, insPoint, textHeight)
        text.color = 2

        centroo(0) = 375 : centroo(1) = lon + 124 : centroo(2) = 0
        radioo = 7.8125
        circulo = objautocad.ActiveDocument.ModelSpace.AddCircle(centroo, radioo)
        circulo.color = 150

        If RadioButton1.Checked = True Then
            centroo(0) = 375 + 75 : centroo(1) = lon + 124 : centroo(2) = 0
            radioo = 7.8125
            circulo = objautocad.ActiveDocument.ModelSpace.AddCircle(centroo, radioo)
            circulo.color = 150
        End If

        centroo(0) = 375 + 75 * 2 : centroo(1) = lon + 124 : centroo(2) = 0
        radioo = 7.8125
        circulo = objautocad.ActiveDocument.ModelSpace.AddCircle(centroo, radioo)
        circulo.color = 150

        'instruccion para centrar el dibujo en el plano..
        objautocad.ZoomExtents()

        MsgBox("Dibujo hecho satisfactoriamente")
        ' Y POR ULTIMO MAXIMIZA LA VENTANA DE DIBUJO(AUTOCAD)
        objautocad.WindowState = Autodesk.AutoCAD.Interop.Common.AcWindowState.acMax
        exe.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized

        ''// parte q libera todas las variables utilizadas en la version 2.0
        vaciar_variables()


        End If ' condicion del inicio este entra si escogio un tipo de tablero
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Sub maldito(ByRef valores_vector() As Integer, ByRef n_valores() As String, ByRef c_valores() As Integer, ByVal aux As Integer)

        ' metodo de orden listo revizar mañana y completarlo:
        '* logica que sigue: selecciona el primer 220 luego busco la diferencia menor y cambios sus puestos luego ya tengo mi primer valor
        '* segundo: busca el segundo del ciclo (2do 220) en la segunda posicion(aux) y de alli hacia adelante comapra e intercambia y asi sucesivamente

        Dim diferencia(100)
        'MsgBox(aux)
        For i = aux To l  ' con esto obtengo todas las diferencias 
            diferencia(i) = Math.Abs((carga220(aux) / 2) - valores_vector(i))
            '  MsgBox("valor: " & valores_vector(i) & " diferencia: " & diferencia(i))
        Next
        '//////////////
        Dim menor = 100000000, bandera, bandera_f, n_bandera, c_bandera

        For i = aux To l
            If menor > diferencia(i) Then
                menor = diferencia(i)
                bandera = i
            End If

        Next

        ' en esta parte consigo la posicion del menor valor de las diferencias y hago el intercambio de valores delvector
        ' depositando esta posicion el una vriable llamada: "bandera"

        bandera_f = valores_vector(aux)
        n_bandera = n_valores(aux)
        c_bandera = c_valores(aux)

        valores_vector(aux) = valores_vector(bandera)
        n_valores(aux) = n_valores(bandera)
        c_valores(aux) = c_valores(bandera)

        valores_vector(bandera) = bandera_f
        n_valores(bandera) = n_bandera
        c_valores(bandera) = c_bandera

        'MsgBox(menor)
        'For j = 1 To vecesiluminacion + vecestoma
        'MsgBox(valores_vector(j))
        ' Next

    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click

        objCAD = GetObject(, "autocad.application")

        Dim xy(2) As Double
        Dim xyf(2) As Double

        xy(0) = 2 : xy(1) = 10 : xy(2) = 0       ' la forma que debe ser 
        xyf(0) = 2 : xyf(1) = 0 : xyf(2) = 0

        ' laprimera columna contrala el eje x
        ' la segunda columna el eje y

        objCAD.ActiveDocument.ModelSpace.AddLine(xy, xyf)



        ' Dim inicio(2) As Double
        ' Dim final(2) As Double

        ' inicio(0) = 5 : inicio(1) = 5 : inicio(2) = 0
        'final(0) = 5 : final(1) = 15 : final(2) = 0

        '  objCAD.ActiveDocument.ModelSpace.AddLine(inicio, final)


        ' InitializeComponent()
        'SuspendLayout()

        'Dim boton_alfredo As Object
        'boton_alfredo = New System.Windows.Forms.Button

        '        boton_alfredo.Location = New System.Drawing.Point(20, 54)
        '       boton_alfredo.Name = "Button1"
        '      boton_alfredo.Size = New System.Drawing.Size(85, 35)
        '     boton_alfredo.TabIndex = 0
        '    boton_alfredo.Text = "alfredo"
        '   boton_alfredo.UseVisualStyleBackColor = True
        '  Controls.Add(boton_alfredo)

    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'NOTA: hacerlo personalizado para toma


    End Sub
    Public Sub vaciar_variables()

        l = 1 'variable que contiene el totar de cargas(iluminacion ,toma)
        veces220 = Nothing
        veces220spe = Nothing
        vecesiluminacion = Nothing
        vecestoma = Nothing
        For i = 1 To 100

            ' variables donde se colocan nombre, codigo y carga definitiva (con equilibrio hecho) para la creacion del tablero
            v_deficode(i) = Nothing
            v_definitivo(i) = Nothing
            v_definombre(i) = Nothing

            ' variables q contienen toda la informacion de TOMA
            cargas(i) = Nothing
            n_carga(i) = Nothing
            c_cargat(i) = Nothing

            ' variables q contienen toda la informacion de 220
            carga220(i) = Nothing
            n_carga220(i) = Nothing
            c_carga220(i) = Nothing

            ' variables q contienen toda la informacion de 220 especial
            carga220spe(i) = Nothing
            n_carga220spe(i) = Nothing
            c_carga220spe(i) = Nothing

            ' variables q contienen toda la informacion de iluminacion
            cargasil(i) = Nothing
            n_cargasil(i) = Nothing
            c_cargail(i) = Nothing

            'variables para separar en equilibrio de monofasicos
            vectorA(i) = Nothing
            n_vectorA(i) = Nothing
            c_vectorA(i) = Nothing

            vectorB(i) = Nothing
            n_vectorB(i) = Nothing
            c_vectorB(i) = Nothing

            ' variables utilizadas al inicio para la reloccion de datos tipo "text" en el model space
            depiuminacion(i) = Nothing
            deptoma(i) = Nothing
            deptoma220(i) = Nothing
            deptoma220spe(i) = Nothing

            ' variables utilizadas para ingresar todos los valores de toma e iluminacion
            valorres(i) = Nothing
            n_valorres(i) = Nothing
            c_valorres(i) = Nothing
        Next

    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub AtomaticToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AtomaticToolStripMenuItem.Click
        'BARRITA Q SE LLENA !
        Dim x As Long

        ToolStripProgressBar1.Maximum = 10000
        ToolStripProgressBar1.Minimum = 0
        ToolStripProgressBar1.Value = 0

        ' Generamos un ciclo For
        For x = ToolStripProgressBar1.Minimum To ToolStripProgressBar1.Maximum

            ' Mostramos la veriable x (el value) en Label1
            ' Label1.Text = x

            ' Asignamos en la propiedad Value del control ProgressBar _
            'el valor de x para ir incrementando la barra de progreso
            ToolStripProgressBar1.Value = x

        Next x

        MsgBox("programa CREADOR DE TABLEROS ELECTRICOS dibujados en AUTOCAD ")
        Form2.Show()

        ToolStripProgressBar1.Value = 0
    End Sub

    Private Sub OnToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OnToolStripMenuItem.Click

        If on_off = 0 Then ' si esta apagado se prende
            Button8.Visible = True
            on_off = 1

            OnToolStripMenuItem.Text = "OFF"
        ElseIf on_off = 1 Then ' si esta prendido se apaga
            Button8.Visible = False
            on_off = 0

            OnToolStripMenuItem.Text = "ON"
        End If

    End Sub

    Private Sub ComputoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComputoToolStripMenuItem.Click
        Dim contar As New contador
        Dim resultado As Integer = 0 : Dim total : total = 0
        Dim opcion
        Dim interruptor_S, interruptor_3S, interruptor_2S, interrupto_S3
        Dim b20_1, b20_2, b30_1, b30_2, b40_1, b40_2, b50_2, b75_2, b90_2, b100_2, b125_2, b40_3, b50_3, b75_3, b90_3, b100_3, b125_3, b175_3, b200_3


        'direccion del archivo..
        uu = "C:\Formato Computos Metricos xlsx.xlsx" ' se coloca la direccion del documento excel q se utiliza...
        'se abre excel
        exe = CreateObject("excel.application", "")
        ' se abre un espacio de trabajo
        exe.Workbooks.Open(uu)
        'este es visible
        exe.Visible = True

        '////inicio
        contar.espacio_dibujo()
        '///cantidad total de tomas
        exe.Cells(4, 4) = contar.tomacorriente
        '///cantidad total de iluminacion
        exe.Cells(7, 4) = contar.Alumbrado
        '///cantidad total de toma220
        exe.Cells(5, 4) = contar.toma220
        '///cantidad total de toma220spe
        exe.Cells(6, 4) = contar.toma220spe
        '///longuitud de cable tomacorriente
        exe.Cells(15, 4) = contar.canalizaciones_tomacorriente()
        '///longuitud de cable alumbrado
        exe.Cells(14, 4) = contar.canalizaciones_alumbrado
        '//cantidad total de fluorecente2*40
        exe.Cells(36, 4) = contar.fluorecente_doscuarenta
        '//cantidad total de fluorecente3*40
        exe.Cells(8, 4) = contar.fluorecente_trescuarenta
        '//cantidad total de fluorecente3*40
        exe.Cells(9, 4) = contar.fluorecente_cuatrocuarenta
        'cantidad de interruptores
        interruptor_S = contar.interruptores("S")
        interruptor_2S = contar.interruptores("2S")
        interruptor_3S = contar.interruptores("3S")
        interrupto_S3 = contar.interruptores("S3")
        exe.Cells(10, 4) = interruptor_S
        exe.Cells(11, 4) = interruptor_2S
        exe.Cells(12, 4) = interruptor_3S
        exe.Cells(13, 4) = interrupto_S3
        ' cantidad de breaker
        b20_1 = contar.breaker("20A-1P")
        b20_2 = contar.breaker("20A-2P")
        b30_1 = contar.breaker("30A-1P")
        b30_2 = contar.breaker("30A-2P")
        b40_1 = contar.breaker("40A-1P")
        b40_2 = contar.breaker("40A-2P")
        b50_2 = contar.breaker("50A-2P")
        b75_2 = contar.breaker("75A-2P")
        b90_2 = contar.breaker("90A-2P")
        b100_2 = contar.breaker("100A-2P")
        b125_2 = contar.breaker("125A-2P")
        b40_3 = contar.breaker("40A-3P")
        b50_3 = contar.breaker("50A-3P")
        b75_3 = contar.breaker("75A-3P")
        b90_3 = contar.breaker("90A-3P")
        b100_3 = contar.breaker("100A-3P")
        b125_3 = contar.breaker("125A-3P")
        b175_3 = contar.breaker("175A-3P")
        b200_3 = contar.breaker("200A-3P")

        '  MsgBox("20A-1P: " & b20_1 & " 20A-2P: " & b20_2 & " 30A-1P: " & b30_1 & " 30A-2P: " & b30_2 & _
        '         " 40A-1P: " & b40_1 & " 40A-2P: " & b40_2 & " 50A-2P: " & b50_2 & " 75A-2P: " & b75_2 & _
        '        " 90A-2P: " & b90_2 & " 100A-2P: " & b100_2 & " 125A-2P: " & b125_2 & " 40A-3P: " & b40_3 & _
        '       " 50A-3P: " & b50_2 & " 75A-3P: " & b75_2 & " 90A-3P: " & b90_3 & " 100A-3P: " & b100_3 & " 125A-3P: " & b125_3 & _
        '      " 175A-3P: " & b175_3 & " 200A-3P: " & b200_3, "CANTIDAD TOTAL DE BREAKERS")
        exe.Cells(17, 4) = b20_1
        exe.Cells(18, 4) = b20_2
        exe.Cells(19, 4) = b30_1
        exe.Cells(20, 4) = b30_2
        exe.Cells(21, 4) = b40_1
        exe.Cells(22, 4) = b40_2
        exe.Cells(23, 4) = b50_2
        exe.Cells(24, 4) = b75_2
        exe.Cells(25, 4) = b90_2
        exe.Cells(26, 4) = b100_2
        exe.Cells(27, 4) = b125_2
        exe.Cells(28, 4) = b40_3
        exe.Cells(29, 4) = b50_3
        exe.Cells(30, 4) = b75_3
        exe.Cells(31, 4) = b90_3
        exe.Cells(32, 4) = b100_3
        exe.Cells(33, 4) = b125_3
        exe.Cells(34, 4) = b175_3
        exe.Cells(35, 4) = b200_3

        MsgBox("Computo hecho satisfactoriamente", MsgBoxStyle.Information, "COMPUTO")


    End Sub

    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        exe = CreateObject("excel.application", "")

        exe.Workbooks.Open("C:\Formato Computos Metricos xlsx.xlsx")
        'este es visible
        exe.Visible = True

    End Sub
End Class

Class contador
    Dim texto As Autodesk.AutoCAD.Interop.Common.AcadText
    Dim acumulador_longuitud As Decimal = 0
    Dim arco As Autodesk.AutoCAD.Interop.Common.AcadArc
    Dim resultado
    Dim seleccion As Autodesk.AutoCAD.Interop.AcadSelectionSet
    Dim obj_cad As Autodesk.AutoCAD.Interop.AcadApplication  ' instancia de clase padre!!
    Public Sub espacio_dibujo()
        obj_cad = GetObject(, "Autocad.application")
    End Sub
    Public Function tomacorriente()
        Dim filtertype(1) As Short
        Dim filterdata(1) As Object
        seleccion = obj_cad.ActiveDocument.SelectionSets.Add("tomacorrientes")
        filtertype(0) = 0
        filterdata(0) = "circle"
        filtertype(1) = 62
        filterdata(1) = 40 'color naranja
        seleccion.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , filtertype, filterdata)
        resultado = seleccion.Count
        'MsgBox(seleccion.Count)
        obj_cad.ActiveDocument.SelectionSets.Item("tomacorrientes").Delete()
        Return (resultado) ' retono de la cantidad de tomas que haya en el dibujo!!
    End Function
    Public Function toma220spe()
        Dim filtertype(1) As Short
        Dim filterdata(1) As Object
        seleccion = obj_cad.ActiveDocument.SelectionSets.Add("toma220spe-")
        filtertype(0) = 0
        filterdata(0) = "circle"
        filtertype(1) = 62
        filterdata(1) = 42
        seleccion.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , filtertype, filterdata)
        resultado = seleccion.Count
        obj_cad.ActiveDocument.SelectionSets.Item("toma220spe-").Delete()
        Return (resultado)
    End Function
    Public Function toma220()
        Dim filtertype(1) As Short
        Dim filterdata(1) As Object
        seleccion = obj_cad.ActiveDocument.SelectionSets.Add("toma220-")
        filtertype(0) = 0
        filterdata(0) = "circle"
        filtertype(1) = 62
        filterdata(1) = 43
        seleccion.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , filtertype, filterdata)
        resultado = seleccion.Count
        obj_cad.ActiveDocument.SelectionSets.Item("toma220-").Delete()
        Return (resultado)
    End Function
    Public Function Alumbrado()
        Dim filtertype(1) As Short
        Dim filterdata(1) As Object
        seleccion = obj_cad.ActiveDocument.SelectionSets.Add("alumbrado")
        filtertype(0) = 0
        filterdata(0) = "circle"
        filtertype(1) = 62
        filterdata(1) = 1
        seleccion.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , filtertype, filterdata)
        resultado = seleccion.Count
        obj_cad.ActiveDocument.SelectionSets.Item("alumbrado").Delete()
        Return (resultado)
    End Function
    Public Function fluorecente_doscuarenta()
        Dim filtertype(1) As Short
        Dim filterdata(1) As Object
        seleccion = obj_cad.ActiveDocument.SelectionSets.Add("2*40")
        filtertype(0) = 0
        filterdata(0) = "circle"
        filtertype(1) = 62
        filterdata(1) = 2
        seleccion.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , filtertype, filterdata)
        resultado = seleccion.Count
        obj_cad.ActiveDocument.SelectionSets.Item("2*40").Delete()
        Return (resultado)

    End Function
    Public Function fluorecente_trescuarenta()
        Dim filtertype(1) As Short
        Dim filterdata(1) As Object
        seleccion = obj_cad.ActiveDocument.SelectionSets.Add("3*40")
        filtertype(0) = 0
        filterdata(0) = "circle"
        filtertype(1) = 62
        filterdata(1) = 4
        seleccion.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , filtertype, filterdata)
        resultado = seleccion.Count
        obj_cad.ActiveDocument.SelectionSets.Item("3*40").Delete()
        Return (resultado)

    End Function

    Public Function fluorecente_cuatrocuarenta()
        Dim filtertype(1) As Short
        Dim filterdata(1) As Object
        seleccion = obj_cad.ActiveDocument.SelectionSets.Add("4*40")
        filtertype(0) = 0
        filterdata(0) = "circle"
        filtertype(1) = 62
        filterdata(1) = 3
        seleccion.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , filtertype, filterdata)
        resultado = seleccion.Count
        obj_cad.ActiveDocument.SelectionSets.Item("4*40").Delete()
        Return (resultado)
    End Function
    Public Function canalizaciones_tomacorriente()
        Dim filtertype(1) As Short
        Dim filterdata(1) As Object
        seleccion = obj_cad.ActiveDocument.SelectionSets.Add("tomacorriente_canalizacionn")
        filtertype(0) = 0
        filterdata(0) = "arc"
        filtertype(1) = 62
        filterdata(1) = 40 'color naranja
        seleccion.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , filtertype, filterdata)
        For Each arco In seleccion
            acumulador_longuitud += arco.ArcLength
        Next
        resultado = acumulador_longuitud * 3
        obj_cad.ActiveDocument.SelectionSets.Item("tomacorriente_canalizacionn").Delete()
        Return (resultado)
    End Function

    Public Function interruptores(ByVal tipo_interruptor)
        Dim total_S As Integer = 0 : Dim total_2S As Integer = 0 : Dim total_3S As Integer = 0 : Dim total_S3 As Integer = 0
        Dim filtertype(0) As Short
        Dim filterdata(0) As Object
        seleccion = obj_cad.ActiveDocument.SelectionSets.Add("interruptor1")
        filtertype(0) = 0
        filterdata(0) = "text"
        seleccion.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , filtertype, filterdata)
        For Each texto In seleccion
            If texto.TextString = "S" Then
                total_S += 1
            End If
            If texto.TextString = "3S" Then
                total_3S += +1
            End If
            If texto.TextString = "2S" Then
                total_2S += 1
            End If
            If texto.TextString = "S3" Then
                total_S3 += +1
            End If
        Next
        obj_cad.ActiveDocument.SelectionSets.Item("interruptor1").Delete()
        If tipo_interruptor = "S" Then
            Return (total_S)
        ElseIf tipo_interruptor = "2S" Then
            Return (total_2S)
        ElseIf tipo_interruptor = "3S" Then
            Return (total_3S)
        ElseIf tipo_interruptor = "S3" Then
            Return (total_S3)
        End If
    End Function
    Public Function canalizaciones_alumbrado()
        Dim filtertype(1) As Short
        Dim filterdata(1) As Object
        seleccion = obj_cad.ActiveDocument.SelectionSets.Add("alumbradoo_canalizacion")
        filtertype(0) = 0
        filterdata(0) = "arc"
        filtertype(1) = 62
        filterdata(1) = 1 'color rojo
        seleccion.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , filtertype, filterdata)
        For Each arco In seleccion
            acumulador_longuitud += arco.ArcLength
        Next
        resultado = acumulador_longuitud * 2.5
        obj_cad.ActiveDocument.SelectionSets.Item("alumbradoo_canalizacion").Delete()
        Return (resultado)
    End Function
    Public Function breaker(ByVal tipo_breaker)
        Dim a, b, c, d, e, f, g, h, i, j, k, l, m, n, o, p, q, r, s As Integer
        a = 0 = b = c = d = e = f = g = h = i = j = k = l = m = n = o = p = q = r = s
        Dim filterdata(0) As Object
        Dim filtertype(0) As Short
        seleccion = obj_cad.ActiveDocument.SelectionSets.Add("interruptor2")
        filtertype(0) = 0
        filterdata(0) = "text"
        seleccion.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , filtertype, filterdata)

        For Each texto In seleccion
            If texto.TextString = "20A-1P" Then
                a += 1
            End If
            If texto.TextString = "20A-2P" Then
                b += 1
            End If
            If texto.TextString = "30A-1P" Then
                c += 1
            End If
            If texto.TextString = "30A-2P" Then
                d += 1
            End If
            If texto.TextString = "40A-1P" Then
                e += 1
            End If
            If texto.TextString = "40A-2P" Then
                f += 1
            End If
            If texto.TextString = "50A-2P" Then
                g += 1
            End If
            If texto.TextString = "75A-2P" Then
                h += 1
            End If
            If texto.TextString = "90A-2P" Then
                i += 1
            End If
            If texto.TextString = "100A-2P" Then
                j += 1
            End If
            If texto.TextString = "125A-2P" Then
                k += 1
            End If
            If texto.TextString = "40A-3P" Then
                l += 1
            End If
            If texto.TextString = "50A-3P" Then
                m += 1
            End If
            If texto.TextString = "75A-3P" Then
                n += 1
            End If
            If texto.TextString = "90A-3P" Then
                o += 1
            End If
            If texto.TextString = "100A-3P" Then
                p += 1
            End If
            If texto.TextString = "125A-3P" Then
                q += 1
            End If
            If texto.TextString = "175A-3P" Then
                r += 1
            End If
            If texto.TextString = "200A-3P" Then
                s += 1
            End If
        Next
        obj_cad.ActiveDocument.SelectionSets.Item("interruptor2").Delete()

        If tipo_breaker = "20A-1P" Then
            Return (a)
        ElseIf tipo_breaker = "20A-2P" Then
            Return (b)
        ElseIf tipo_breaker = "30A-1P" Then
            Return (c)
        ElseIf tipo_breaker = "30A-2P" Then
            Return (d)
        ElseIf tipo_breaker = "40A-1P" Then
            Return (e)
        ElseIf tipo_breaker = "40A-2P" Then
            Return (f)
        ElseIf tipo_breaker = "50A-2P" Then
            Return (g)
        ElseIf tipo_breaker = "75A-2P" Then
            Return (h)
        ElseIf tipo_breaker = "90A-2P" Then
            Return (i)
        ElseIf tipo_breaker = "100A-2P" Then
            Return (j)
        ElseIf tipo_breaker = "125A-2P" Then
            Return (k)
        ElseIf tipo_breaker = "40A-3P" Then
            Return (l)
        ElseIf tipo_breaker = "50A-3P" Then
            Return (m)
        ElseIf tipo_breaker = "75A-3P" Then
            Return (n)
        ElseIf tipo_breaker = "90A-3P" Then
            Return (o)
        ElseIf tipo_breaker = "100A-3P" Then
            Return (p)
        ElseIf tipo_breaker = "125A-3P" Then
            Return (q)
        ElseIf tipo_breaker = "175A-3P" Then
            Return (r)
        ElseIf tipo_breaker = "200A-3P" Then
            Return (s)
        End If
    End Function
End Class