Imports System.Math
Imports TheGA
Imports Microsoft.Office.Interop
Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel

'Ensure that Excel book is named book 1

'Initialze all the arrays
Module Module1

    Dim Row As Integer = 45


    Public Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Integer) As Integer
    Public Declare Auto Function FindWindow Lib "user32.dll" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Public Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As Integer) As Boolean
    Public Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Integer, ByVal nCmdShow As Integer) As Integer

    Public Const SW_RESTORE As Integer = 9
    Public Const SW_SHOW As Integer = 5


    Sub FocusWindow(ByVal strWindowCaption As String, ByVal strClassName As String)
        Dim hWnd As Integer
        hWnd = FindWindow(strClassName, strWindowCaption)

        If hWnd > 0 Then
            SetForegroundWindow(hWnd)

            If IsIconic(hWnd) Then  'Restore if minimized
                ShowWindow(hWnd, SW_RESTORE)
            Else
                ShowWindow(hWnd, SW_SHOW)
            End If
        End If
    End Sub

    Sub Main()
        Dim obj As Object
        obj = GetObject(, "StaadPro.OpenSTAAD")

        Dim objApp As Excel.Application
        Dim objBook As Excel._Workbook

        Dim objBooks As Excel.Workbooks
        Dim objSheets As Excel.Sheets
        Dim Active_sheet As Excel._Worksheet

        objApp = New Excel.Application()
        objBooks = objApp.Workbooks
        objBook = objBooks.Add
        objSheets = objBook.Worksheets
        Active_sheet = objSheets(1)

        Dim n1, n2, n3 As Integer
        n1 = 10 'Represents population size
        'n2 Represents members
        n3 = 8 'Represents gene size
        n2 = obj.Geometry.GetMemberCount()
        Dim Member_property(n2) As String
        Dim n_pipe, n_square, n_ISMB, n_taper, n_prismatic As Integer
        n_pipe = 0
        n_square = 0
        n_ISMB = 0
        n_taper = 0
        n_prismatic = 0


        For iter = 1 To n2
            Member_property(iter - 1) = obj.Property.GetBeamSectionName(iter)
            If Member_property(iter - 1) = "SQU 0.1X0.1X0.0" Then
                n_square += 1
                Member_property(iter - 1) = "Square"
            ElseIf Member_property(iter - 1) = "PIP2445H" Then
                n_pipe = +=1
                Member_property(iter - 1) = "Pipe"
            ElseIf Member_property(iter - 1) = "ISMB150" Then
                n_ISMB += 1
                Member_property(iter - 1) = "ISMB"
            ElseIf Member_property(iter - 1) = "Taper" Then
                n_taper += 1
                Member_property(iter - 1) = "Taper"
            ElseIf Member_property(iter - 1) = "Prismatic General" Then
                n_prismatic += 1
                Member_property(iter - 1) = "Prismatic"
            Else
                Console.Write("Unexpected member present: ")
                Console.WriteLine(iter)
            End If
        Next

        objApp.Visible = True
        objApp.UserControl = False


        Console.WriteLine("For writing the input population")
        'For Entering the data in excel sheet as array of length 10()populationx20(member)
        Console.ReadKey()
        Console.WriteLine("Population Entered")



        Dim Population_Array(n1 - 1, n2 - 1, n3 - 1) As Population_class


        'Populating the array
        'Enter the initial population
        Dim Gene(n3 - 1) As Integer
        Dim Prismatic_Gene((n3 / 2) - 1) As Integer
        Dim Counter1, Counter2, n_row As Integer
        n_row = 0
        Dim Ix, Iy As Double

        'Gene Size matters here (It needs to be even)
        For iter1 = 1 To Population_Array.GetLength(0)
            Counter1 = 1
            Counter2 = 1
            For iter2 = 1 To Population_Array.GetLength(1)
                If Member_property(iter2) = "Prismatic" Then

                    n_row = 18 + (iter1 * 2)
                    Ix = Active_sheet.Cells(n_row, Counter1).Value()
                    Iy = Active_sheet.Cells(n_row + 1, Counter1).Value()

                    Prismatic_Gene = Real_to_binary(Ix, 4, 0.00005, 0.0000005)

                    For iter3 = 1 To (n3 / 2)
                        Population_Array((iter1 - 1), (iter2 - 1), (iter3 - 1)) = New Population_class()
                        Population_Array((iter1 - 1), (iter2 - 1), (iter3 - 1)).Value() = Prismatic_Gene(iter3 - 1)
                    Next

                    Prismatic_Gene = Real_to_binary(Iy, 4, 0.00005, 0.0000005)

                    For iter3 = ((n3 / 2) + 1) To n3
                        Population_Array((iter1 - 1), (iter2 - 1), (iter3 - 1)) = New Population_class()
                        Population_Array((iter1 - 1), (iter2 - 1), (iter3 - 1)).Value() = Prismatic_Gene(iter3 - ((n3 / 2) + 1))
                    Next
                    Counter1 += 1
                Else
                    Gene = Real_to_binary(Active_sheet.Cells(iter1, Counter2).value(), 8, 10, 1) 'Gene size
                    For iter3 = 1 To Gene.Length
                        Population_Array((iter1 - 1), (iter2 - 1), (iter3 - 1)) = New Population_class()
                        Population_Array((iter1 - 1), (iter2 - 1), (iter3 - 1)).Value() = Gene(iter3 - 1)
                    Next
                    Counter2 += 1
                End If
            Next
        Next


        Dim Population_Instance(Population_Array.GetLength(1) - n_prismatic - 1, Population_Array.GetLength(2) - 1) As Population_class
        Dim Altered_thickness() As Double
        Dim Sheet_str As String
        Dim Ineria_instance(n_prismatic - 1, Population_Array.GetLength(2) - 1) As Population_class
        Dim Altered_Inertia(n_prismatic - 1, 1) As Double

        For gen = 1 To 20

            Console.WriteLine("********************************************")
            Console.Write("This is gen: ")
            Console.WriteLine(gen)
            Console.WriteLine("********************************************")

            Sheet_str = "Sheet" + CStr(gen + 1)
            objSheets.Add()
            Active_sheet = objSheets(Sheet_str)
            Active_sheet.Activate()

            Population_Array = Genetic_Algorithm_Application(Population_Array, obj, Active_sheet, Member_property)

            'all members initially should be thickness based and later one should be prismatic general

            For iter1 = 0 To (Population_Array.GetLength(0) - 1)

                For iter2 = 0 To (Population_Array.GetLength(1) - n_prismatic - 1)
                    For iter3 = 0 To (Population_Array.GetLength(2) - 1)
                        Population_Instance(iter2, iter3) = New Population_class
                        Population_Instance(iter2, iter3) = Population_Array(iter1, iter2, iter3)
                    Next
                Next
                Altered_thickness = Binary_to_Real(Population_Instance, 10, 1)
                For iter4 = 1 To Altered_thickness.Length
                    Active_sheet.Cells((iter1 + 1), iter4) = Altered_thickness(iter4 - 1)
                Next

                For iter2 = (Population_Array.GetLength(1) - n_prismatic) To (Population_Array.GetLength(1) - 1)
                    For iter3 = 0 To (Population_Array.GetLength(2) - 1)
                        Ineria_instance(iter2 - Population_Array.GetLength(1) - n_prismatic, iter3) = New Population_class
                        Ineria_instance(iter2 - Population_Array.GetLength(1) - n_prismatic, iter3) = Population_Array(iter1, iter2, iter3)
                    Next
                Next
                Altered_Inertia = Binary_to_Real2(Ineria_instance, 0.00005, 0.0000005)

                For iter4 = 1 To Altered_Inertia.GetLength(0)
                    Active_sheet.Cells(20 + 2 * iter1, iter4) = Altered_Inertia(iter4 - 1, 1)
                    Active_sheet.Cells(21 + 2 * iter1, iter4) = Altered_Inertia(iter4 - 1, 2)
                Next

            Next

        Next

        objApp.Visible = True
        objApp.UserControl = True

        Active_sheet = Nothing
        objSheets = Nothing
        objBooks = Nothing
        obj = Nothing

        Console.ReadKey()
        Console.Write("This is done")
        Console.ReadKey()



    End Sub

    'Binary to real converter for Inertia_instance array
    Private Function Binary_to_Real2(inertia_instance(,) As Population_class, u_limit As Double, L_limit As Double) As Double(,)
        Dim n_bits As Integer = inertia_instance.GetLength(1)
        Dim Inertia_values(inertia_instance.GetLength(0) - 1, 1) As Double
        Dim max_value, Decoded_value As Integer
        Dim diff As Integer
        diff = u_limit - L_limit
        max_value = Pow(2, (n_bits / 2)) - 1

        For iter1 = 0 To inertia_instance.GetLength(0) - 1
            Decoded_value = 0
            For iter2 = 0 To ((n_bits / 2) - 1)
                Decoded_value += inertia_instance(iter1, iter2).Value() * Pow(2, ((n_bits / 2) - 1 - iter2))
            Next
            Inertia_values(iter1, 0) = L_limit + ((diff / max_value) * Decoded_value)
            Decoded_value = 0
            For iter2 = (n_bits / 2) To n_bits - 1
                Decoded_value += inertia_instance(iter1, iter2).Value() * Pow(2, (n_bits - 1 - iter2))
            Next
            Inertia_values(iter1, 1) = L_limit + ((diff / max_value) * Decoded_value)
        Next

            Return Inertia_values
    End Function


    'Definition of fitness function
    Public Function FitnessFunction(Population_Array(,) As Population_class, ByRef obj As Object, Active_sheet As _Worksheet, Member_property As String(), n_prismatic As Integer) As Double
        Dim F_value, CV_Value As Double
        Dim Ratio_array(Population_Array.GetLength(0) - 1), Real_thickness() As Double
        Dim U_bound_thickness, L_bound_thickness, ratio As Double
        U_bound_thickness = 10.0
        L_bound_thickness = 1.0
        Dim Inertia_array(n_prismatic - 1, 1) As Double
        Dim array_For_thckness(Population_Array.GetLength(0) - n_prismatic - 1, Population_Array.GetLength(1) - 1), Array_for_inertia(n_prismatic - 1, Population_Array.GetLength(1) - 1) As Population_class
        Dim iter1, iter2 As Integer

        For iter1 = 0 To (Population_Array.GetLength(0) - n_prismatic - 1)
            For iter2 = 0 To (Population_Array.GetLength(1) - 1)
                array_For_thckness(iter1, iter2) = Population_Array(iter1, iter2)
            Next
        Next

        For iter1 = (Population_Array.GetLength(0) - n_prismatic) To (Population_Array.GetLength(0) - 1)
            For iter2 = 0 To (Population_Array.GetLength(1) - 1)
                Array_for_inertia(iter1 - Population_Array.GetLength(0) - n_prismatic, iter2) = Population_Array(iter1, iter2)
            Next
        Next

        Real_thickness = Binary_to_Real(Population_Array, U_bound_thickness, L_bound_thickness)
        Inertia_array = Binary_to_Real2(Array_for_inertia, 0.00005, 0.0000005)

        'U_bound and L_bound to be given in mm
        'Real_thickness = Binary_to_Real(Population_Array, U_bound, L_bound)
        Dim Last_property_ref As Integer
        Dim Result As Boolean
        Dim thickness_in_m As Double
        Last_property_ref = obj.Property.GetSectionPropertyCount()
        Dim PropertyArrayISMB(6), PropertyArrayTaper(6), PropertyArrayGeneral(9) As Double

        PropertyArrayISMB(0) = 0.15
        PropertyArrayISMB(1) = 0.007 'This is thickness of web
        PropertyArrayISMB(2) = 0.15
        PropertyArrayISMB(3) = 0.075
        PropertyArrayISMB(4) = 0.007 'This is thickness of flange
        PropertyArrayISMB(5) = 0.075
        PropertyArrayISMB(6) = 0.007 'This is thickness of bottom Flange

        PropertyArrayTaper(0) = 0.152
        PropertyArrayTaper(1) = 0.00571 'thickness
        PropertyArrayTaper(2) = 0.152
        PropertyArrayTaper(3) = 0.0762
        PropertyArrayTaper(4) = 0.00571 'thickness
        PropertyArrayTaper(5) = 0.0762
        PropertyArrayTaper(6) = 0.00571 'thickness

        PropertyArrayGeneral(0) = 0.001342
        PropertyArrayGeneral(1) = 0.000762
        PropertyArrayGeneral(2) = 0.0004354
        PropertyArrayGeneral(3) = 0.000000008215 'Ix
        PropertyArrayGeneral(4) = 0.000004161 'Iy
        PropertyArrayGeneral(5) = 0.000004161
        PropertyArrayGeneral(6) = 0.0889
        PropertyArrayGeneral(7) = 0.1852
        PropertyArrayGeneral(8) = 0
        PropertyArrayGeneral(9) = 0

        Dim MyThread As System.Threading.Thread

        For iter1 = 1 To (Population_Array.GetLength(0))
            'thickness input unit is meter
            If iter1 < Population_Array.GetLength(0) - n_prismatic Then
                thickness_in_m = (Real_thickness(iter1 - 1)) / 1000
                If Member_property(iter1) = "Square" Then
                    obj.Property.CreateTaperedTubeProperty(5, 0.0889, 0.0889, thickness_in_m)
                ElseIf Member_property(iter1) = "Pipe" Then
                    obj.Property.CreateTaperedTubeProperty(0, 0.2445, 0.2445, thickness_in_m)
                ElseIf Member_property(iter1) = "ISMB" Then
                    PropertyArrayISMB(1) = PropertyArrayISMB(4) = PropertyArrayISMB(6) = thickness_in_m
                    obj.Property.CreateTaperedIProperty(PropertyArrayISMB)
                ElseIf Member_property(iter1) = "Taper" Then
                    PropertyArrayTaper(1) = PropertyArrayTaper(4) = PropertyArrayTaper(6) = thickness_in_m
                    obj.Property.CreateTaperedIProperty(PropertyArrayTaper)
                End If
            Else
                If Member_property(iter1) = "Prismatic" Then
                    PropertyArrayGeneral(3) = Inertia_array(iter1 - (Population_Array.GetLength(0) - n_prismatic), 0)
                    PropertyArrayGeneral(4) = Inertia_array(iter1 - (Population_Array.GetLength(0) - n_prismatic), 1)
                    obj.Property.CreatePrismaticGeneralProperty(PropertyArrayGeneral)
                End If
            End If

            Last_property_ref += 1

            If iter1 = 1 Then
                MyThread = New System.Threading.Thread(AddressOf PressEnterStaad)
                MyThread.Start()
            End If

            obj.Property.AssignBeamProperty(iter1, Last_property_ref)


        Next

        Threading.Thread.Sleep(12000)

        obj.updateStructure()
        obj.Analyze()
        Console.WriteLine("Waiting For analysis")

        Threading.Thread.Sleep(360000)

        MyThread = New System.Threading.Thread(AddressOf PressEnterAnalysis)
        MyThread.Start()

        obj.View.SetInterfaceMode(1)

        MyThread = New System.Threading.Thread(AddressOf PressEnterResults)
        MyThread.Start()

        Threading.Thread.Sleep(20000)


        For iter1 = 1 To Population_Array.GetLength(0)
            Result = obj.Output.GetMemberSteelDesignRatio(iter1, ratio)
            If ratio > 0 Then
                Ratio_array(iter1 - 1) = ratio
            Else
                Ratio_array(iter1 - 1) = 0
                Console.WriteLine("wrong Ratio Recieved")
            End If
        Next


        Active_sheet.Activate()

        'Get ratio, See that ratios are greater than zero
        F_value = Function_Value(Ratio_array)
        CV_Value = Constraint_voilation(Ratio_array)

        'Row defined in first line of module 1 and needs to be modified if population size increases.

        Console.Write("Function value is: ")
        Console.WriteLine(F_value)
        Active_sheet.Cells(Row, 1) = F_value


        Console.Write("Constraint voilation is: ")
        Console.WriteLine(CV_Value)
        Active_sheet.Cells(Row, 2) = CV_Value

        Dim Fitness_function As Double = 100 / (1 + F_value + CV_Value)

        Console.Write("Fitness Function is: ")
        Console.WriteLine(Fitness_function)
        Active_sheet.Cells(Row, 3) = Fitness_function

        Row += 1

        Return Fitness_function

    End Function

    'Binary to real converter
    Private Function Binary_to_Real(population_Array(,) As Population_class, u_bound As Double, l_bound As Double) As Double()
        Dim Real_Array(population_Array.GetLength(0) - 1) As Double
        Dim iter1, iter2 As Integer
        Dim Decoded_value, Max_value As ULong
        Dim Diff As Double
        Diff = u_bound - l_bound
        Dim n_bits As Integer
        n_bits = population_Array.GetLength(1)
        Max_value = Pow(2, population_Array.GetLength(1)) - 1
        'Both may be defined as integer
        For iter1 = 0 To (population_Array.GetLength(0) - 1)
            Decoded_value = 0
            For iter2 = 0 To (n_bits - 1)
                'Define Multiplication in population class
                Decoded_value += population_Array(iter1, iter2).Value() * Pow(2, (n_bits - 1 - iter2))
            Next

            Real_Array(iter1) = l_bound + (Diff * Decoded_value / Max_value)
        Next
        Return Real_Array
    End Function


    'Definition of constraint voilation (R taken as 1000)
    Private Function Constraint_voilation(ratio_array() As Double) As Double
        Dim i As Integer
        Dim Ratio_Difference, CV_Value As Double

        For i = 0 To (ratio_array.Length - 1)
            Ratio_Difference = ratio_array(i) - 1
            If Ratio_Difference > 0 Then
                CV_Value += 100 * Pow(Ratio_Difference, 2)
            End If
        Next
        Return CV_Value
    End Function


    'Definition of Basic Function
    Public Function Function_Value(ratio_array() As Double) As Double
        Dim i As Integer
        Dim Ratio_Sum As Double
        Ratio_Sum = 0

        For i = 0 To (ratio_array.Length - 1)
            Ratio_Sum += ratio_array(i)
        Next
        Return (ratio_array.Length - Ratio_Sum)
    End Function


    'Genetic Algorithm Base application
    Public Function Genetic_Algorithm_Application(Population_Array(,,) As Population_class, ByRef obj As Object, Active_sheet As _Worksheet, Member_property As String()) As Population_class(,,)
        Dim Reproduced_Population(,,), Crossed_Population(,,), Mutated_Population(,,) As Population_class
        Dim Mutation_probability, Crossover_Probability As Double

        'Taking Mutation Probability as .01 as crossover Probability as 0.8
        Mutation_probability = 0.01
        Crossover_Probability = 0.8

        Reproduced_Population = Reproduction_Roullete_Wheel(Population_Array, obj, Active_sheet, Member_property)
        Crossed_Population = Simple_Crossover(Reproduced_Population, Crossover_Probability)
        Mutated_Population = Simple_Mutation(Crossed_Population, Mutation_probability)

        Return Mutated_Population

    End Function

    'Here Goes the code For Mutation of Crossed Population
    Private Function Simple_Mutation(crossed_Population(,,) As Population_class, Mutation_probability As Double) As Population_class(,,)
        Dim Mutated_population(crossed_Population.GetLength(0) - 1, crossed_Population.GetLength(1) - 1, crossed_Population.GetLength(2) - 1) As Population_class
        Dim iter1, iter2, iter3 As Integer
        Dim random_number As Double
        For iter1 = 0 To (crossed_Population.GetLength(0) - 1)
            For iter2 = 0 To (crossed_Population.GetLength(1) - 1)
                For iter3 = 0 To (crossed_Population.GetLength(2) - 1)
                    random_number = Rnd()
                    Mutated_population(iter1, iter2, iter3) = New Population_class()
                    If 0.23 < random_number And random_number > 0.22 Then
                        'need to define the modification in value
                        Mutated_population(iter1, iter2, iter3) = crossed_Population(iter1, iter2, iter3) + 1
                    Else
                        Mutated_population(iter1, iter2, iter3) = crossed_Population(iter1, iter2, iter3)
                    End If
                Next
            Next
        Next

        Return Mutated_population

    End Function

    ' Here goes the code for Crossover of reproduced population
    Private Function Simple_Crossover(reproduced_Population(,,) As Population_class, Crossover_Probability As Double) As Population_class(,,)
        Dim Crossover_population, Total_population As Integer
        'Considered To_be_crossed_population as crossed population array to be returned
        Dim Crossed_population(reproduced_Population.GetLength(0), reproduced_Population.GetLength(1), reproduced_Population.GetLength(2)) As Population_class
        Dim random_number1, random_number2, Extra_random_number As Integer
        Dim population_count As Integer = 0
        Dim Deciding_variable, Inner_decision_variable, Already_taken(reproduced_Population.GetLength(0) - 1) As Boolean
        Deciding_variable = True
        Inner_decision_variable = True
        Total_population = reproduced_Population.GetLength(0)

        'Something could be done for this
        For iter = 0 To (Already_taken.GetUpperBound(0))
            Already_taken(iter) = False
        Next
        Crossed_population = reproduced_Population

        Crossover_population = Crossover_Probability * Total_population

        If (Crossover_population Mod 2) <> 0 Then
            Crossover_population -= 1
        End If

        Dim To_be_crossed_population(Crossover_population - 1, reproduced_Population.GetLength(1) - 1, reproduced_Population.GetLength(2) - 1) As Population_class


        For iter1 = 1 To Int(Crossover_population / 2)
            While Deciding_variable
                random_number1 = Int(Rnd() * Total_population)
                random_number2 = Int(Rnd() * Total_population)
                If random_number1 <> random_number2 Then
                    If (Already_taken(random_number1) = False) And (Already_taken(random_number2) = False) Then
                        Already_taken(random_number1) = True
                        Already_taken(random_number2) = True
                        For iter2 = 0 To (reproduced_Population.GetLength(1) - 1)
                            For iter3 = 0 To (reproduced_Population.GetLength(2) - 1)
                                To_be_crossed_population(population_count, iter2, iter3) = New Population_class()
                                To_be_crossed_population((population_count + 1), iter2, iter3) = New Population_class()
                                To_be_crossed_population(population_count, iter2, iter3) = reproduced_Population(random_number1, iter2, iter3)
                                To_be_crossed_population((population_count + 1), iter2, iter3) = reproduced_Population(random_number2, iter2, iter3)
                            Next
                        Next
                        population_count += 2
                        Deciding_variable = False

                    ElseIf Already_taken(random_number1) = True Then
                        While Inner_decision_variable
                            Extra_random_number = Int(Rnd() * Total_population)
                            If (Extra_random_number <> random_number1) And (Already_taken(Extra_random_number) = False) Then
                                Already_taken(Extra_random_number) = True
                                Already_taken(random_number2) = True
                                For iter2 = 0 To (reproduced_Population.GetLength(1) - 1)
                                    For iter3 = 0 To (reproduced_Population.GetLength(2) - 1)
                                        To_be_crossed_population(population_count, iter2, iter3) = New Population_class()
                                        To_be_crossed_population((population_count + 1), iter2, iter3) = New Population_class()
                                        To_be_crossed_population(population_count, iter2, iter3) = reproduced_Population(Extra_random_number, iter2, iter3)
                                        To_be_crossed_population((population_count + 1), iter2, iter3) = reproduced_Population(random_number2, iter2, iter3)
                                    Next
                                Next
                                population_count += 2
                                Deciding_variable = False
                                Inner_decision_variable = False
                            End If
                        End While

                    ElseIf Already_taken(random_number2) = True Then
                        While Inner_decision_variable
                            Extra_random_number = Int(Rnd() * Total_population)
                            If (Extra_random_number <> random_number2) And (Already_taken(Extra_random_number) = False) Then
                                Already_taken(Extra_random_number) = True
                                Already_taken(random_number1) = True
                                For iter2 = 0 To (reproduced_Population.GetLength(1) - 1)
                                    For iter3 = 0 To (reproduced_Population.GetLength(2) - 1)
                                        To_be_crossed_population(population_count, iter2, iter3) = New Population_class()
                                        To_be_crossed_population((population_count + 1), iter2, iter3) = New Population_class()
                                        To_be_crossed_population(population_count, iter2, iter3) = reproduced_Population(Extra_random_number, iter2, iter3)
                                        To_be_crossed_population((population_count + 1), iter2, iter3) = reproduced_Population(random_number1, iter2, iter3)
                                    Next
                                Next
                                population_count += 2
                                Deciding_variable = False
                                Inner_decision_variable = False
                            End If
                        End While
                    End If
                End If
            End While
        Next


        'Defining 9 point crossover, iter2 below will change, Keep caution
        Dim Crossover_points(8) As Integer
        Dim string_length As Integer
        string_length = (reproduced_Population.GetLength(1)) * (reproduced_Population.GetLength(2))

        For iter1 = 0 To 8
            Crossover_points(iter1) = Int(Rnd() * string_length)
        Next
        Array.Sort(Crossover_points)

        'Dim Gene_position, First_array_coordinate, Second_array_coordinate As Integer
        Dim Temporary_instance As New Population_class()


        For iter1 As Integer = 0 To (To_be_crossed_population.GetLength(0) - 1) Step 2

            'Cautious, iter2 will change depending on even and odd nature of number of crossover point
            For iter2 As Integer = 1 To (Crossover_points.Length() - 1) Step 2

                For iter3 As Integer = Crossover_points(iter2) To Crossover_points(iter2 + 1)

                    Temporary_instance = To_be_crossed_population(iter1, Int(iter3 \ 7), iter3 Mod 7)
                    To_be_crossed_population(iter1, Int(iter3 \ 7), iter3 Mod 7) = To_be_crossed_population(iter1 + 1, Int(iter3 \ 7), iter3 Mod 7)
                    To_be_crossed_population(iter1 + 1, Int(iter3 \ 7), iter3 Mod 7) = Temporary_instance

                Next

            Next

        Next

        Dim count As Integer
        count = 0

        For iter1 = 0 To (reproduced_Population.GetLength(0) - 1)
            If Already_taken(iter1) = True Then
                For iter2 = 0 To (reproduced_Population.GetLength(1) - 1)
                    For iter3 = 0 To (reproduced_Population.GetLength(2) - 1)
                        Crossed_population(iter1, iter2, iter3) = To_be_crossed_population(count, iter2, iter3)
                    Next
                Next
                count += 1
            End If
        Next

        Return Crossed_population

    End Function

    'Here Goes the code of roullete wheel selection process for reproduction
    Private Function Reproduction_Roullete_Wheel(population_Array(,,) As Population_class, ByRef obj As Object, Active_sheet As _Worksheet, Member_property As String()) As Population_class(,,)
        Dim Fitness_values(population_Array.GetLength(0) - 1), Fitness_Sum As Double
        Dim iter1, iter2, iter3, iter4 As Integer
        Dim Population_Instance(population_Array.GetLength(1) - 1, population_Array.GetLength(2) - 1) As Population_class
        Fitness_Sum = 0
        Dim String_Probability(population_Array.GetLength(0) - 1) As Double
        Dim Cumulative_String_probability(population_Array.GetLength(0) - 1) As Double
        Dim Random_number As Double
        Dim Reproduced_Population(population_Array.GetLength(0) - 1, population_Array.GetLength(1) - 1, population_Array.GetLength(2) - 1) As Population_class


        'GetLength(0) Gives the number of rows
        For iter1 = 0 To (population_Array.GetLength(0) - 1)

            For iter2 = 0 To (population_Array.GetLength(1) - 1)
                For iter3 = 0 To (population_Array.GetLength(2) - 1)
                    Population_Instance(iter2, iter3) = New Population_class()
                    Population_Instance(iter2, iter3) = population_Array(iter1, iter2, iter3)

                Next

            Next

            Fitness_values(iter1) = FitnessFunction(Population_Instance, obj, Active_sheet, Member_property)
            Fitness_Sum += Fitness_values(iter1)
        Next

        For iter1 = 0 To (population_Array.GetLength(0) - 1)
            String_Probability(iter1) = Fitness_values(iter1) / Fitness_Sum
            If iter1 = 0 Then
                Cumulative_String_probability(iter1) = String_Probability(iter1)
            Else
                Cumulative_String_probability(iter1) += Cumulative_String_probability(iter1 - 1) + String_Probability(iter1)
            End If
        Next

        For iter1 = 0 To (population_Array.GetLength(0) - 1)
            Random_number = Rnd()
            For iter2 = 0 To (population_Array.GetLength(0) - 1)
                If Random_number <= Cumulative_String_probability(iter2) Then
                    For iter3 = 0 To (population_Array.GetLength(1) - 1)
                        For iter4 = 0 To (population_Array.GetLength(2) - 1)
                            Reproduced_Population(iter1, iter3, iter4) = New Population_class()
                            Reproduced_Population(iter1, iter3, iter4) = population_Array(iter2, iter3, iter4)
                        Next
                    Next
                End If
            Next
        Next
        Return Reproduced_Population
    End Function

    'real to binary converter, returns a bit array of specified size
    Private Function Real_to_binary(val As Double, nBits As Integer, uLimit As Double, lLimit As Double) As Integer()
        Dim decoded_value, Decoded_instance_value As Integer
        Dim Binary_array(nBits - 1) As Integer

        decoded_value = (((val - lLimit) / (uLimit - lLimit)) * (Pow(2, nBits) - 1))
        Decoded_instance_value = decoded_value

        For iter = 0 To (nBits - 1)
            Binary_array(iter) = Decoded_instance_value Mod 2
            Decoded_instance_value = Decoded_instance_value / 2
        Next

        Return Binary_array

    End Function


    Private Sub PressEnterStaad()

        Threading.Thread.Sleep(10000)
        FocusWindow("STAAD.Pro V8i (SELECTseries 6)/Academic", Nothing)
        Threading.Thread.Sleep(5000)
        SendKeys.SendWait("{ENTER}")
        Console.WriteLine("Enter on Staad main window on another thread")

    End Sub


    Private Sub PressEnterAnalysis()

        FocusWindow("STAAD Analysis and Design", Nothing)
        Threading.Thread.Sleep(10000)
        SendKeys.SendWait("{ENTER}")
        Console.WriteLine("STAAD Analysis and Design Executed on another thread")
    End Sub

    Private Sub PressEnterResults()

        Threading.Thread.Sleep(10000)
        FocusWindow("Results Setup", Nothing)
        Threading.Thread.Sleep(10000)
        SendKeys.SendWait("{ENTER}")

        Console.WriteLine("Results Setup Executed on another thread")
    End Sub

    Private Sub ShowExcel()

        'Alwasy had to be book1
        FocusWindow("Book1 - Excel", Nothing)
        Threading.Thread.Sleep(10000)
        SendKeys.SendWait("{ENTER}")

        Console.WriteLine("Excel book Entered")

    End Sub

End Module
