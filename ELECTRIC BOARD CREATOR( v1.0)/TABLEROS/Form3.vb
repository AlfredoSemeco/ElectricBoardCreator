Public Class Form3
    'variables globales:
    '*********************************************************************************
    Dim encender As Double
    Dim exe As Microsoft.Office.Interop.Excel.Application
    Dim uu
    Dim objautocad As Autodesk.AutoCAD.Interop.AcadApplication ' inicia toda lasaplicaciones de autocad utilizadas (referencia q aplique a mi proyecto)..
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

            Dim KTT, KTT2 As Double
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
                'CUANDO ES MONOFASICO
                If KTT2 = 2 Then
                    KTT2 = 1
                    GoTo REINICIO
                End If
                'CAUNDO ES TRIFASICO
                If KTT = 3 Then
                    KTT = 1
                    GoTo REINICIO
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
p41:
            factor_demanda_tomacorriente = InputBox("ingrese factor de demanda de tomacorriente", "tableros")
            'corrector de errores.. si se introduce un tipo de variable q no relaciona con la vaibale repite l apregunta
            If Err.Number <> 0 Then
                Err.Clear()
                MsgBox("error ha introducido un valor caracter", MsgBoxStyle.Critical)
                GoTo p41
            End If
p31:
            factor_demanda_iluminacion = InputBox("ingrese factor de demanda de iluminacion", "tablaero")
            'corrector de errores.. si se introduce un tipo de variable q no relaciona con la vaibale repite l apregunta
            If Err.Number <> 0 Then
                Err.Clear()
                MsgBox("error ha introducido un valor caracter", MsgBoxStyle.Critical)
                GoTo p31
            End If
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
        End If

    End Sub

    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
       

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
        Dim pass As String
        Dim h As Integer
        h = 1
        Do While h <= 3
            pass = InputBox("Ingrese pass para entrar", "tableros electrico " & "intento: " & h)

            If pass = "1991semeco" Then
                MsgBox("pass correcto", MsgBoxStyle.Exclamation)
                Exit Do
            Else
                MsgBox(" pass erroneo", MsgBoxStyle.Critical)

            End If
            h = h + 1
        Loop
        If h = 4 Then
            End
        End If

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

    Private Sub RadioButton5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton5.CheckedChanged

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

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked

    End Sub
End Class
