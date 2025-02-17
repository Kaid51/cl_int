Imports System.Data
Imports System.IO
Imports System.Net.Http
Imports System.Threading

Module AdvancedInterpreter
    Dim variables As New Dictionary(Of String, Object)()
    Dim functions As New Dictionary(Of String, Func(Of Object(), Object))()
    Dim classes As New Dictionary(Of String, ClassDefinition)()
    Dim lineNumber As Integer = 0

    Class ClassDefinition
        Public Properties As Dictionary(Of String, Object)
        Public Methods As Dictionary(Of String, Func(Of Object(), Object))

        Public Sub New()
            Properties = New Dictionary(Of String, Object)()
            Methods = New Dictionary(Of String, Func(Of Object(), Object))()
        End Sub
    End Class

    Sub Main()
        Console.Title = "Books | Coffee | Two Rooks SL"
        Console.WriteLine("Добро пожаловать! Поддерживаемые команды: PRINT, LET, IF, FOR, WHILE, DEFINE FUNCTION, CLASS, IMPORT, FILE, THREAD, TRY, CATCH, GENERATOR, DECORATOR, ASYNC, HTTP")
        Console.WriteLine()
        Console.WriteLine("Введите команды:")


        DefineFunction("LEN", Function(args) args(0).ToString().Length)
        DefineFunction("SUBSTR", Function(args) args(0).ToString().Substring(CInt(args(1)), CInt(args(2))))
        DefineFunction("TO_UPPER", Function(args) args(0).ToString().ToUpper())

        While True
            Console.Write("> ")
            Dim input As String = Console.ReadLine()

            Try
                lineNumber += 1
                ExecuteCommand(input)
            Catch ex As Exception
                Console.WriteLine($"Ошибка в строке {lineNumber}: {ex.Message}")
            End Try
        End While
    End Sub

    Sub ExecuteCommand(command As String)
        command = command.Trim()

        If command.ToUpper().StartsWith("PRINT ") Then
            Dim expression = command.Substring(6).Trim()
            Dim result = EvaluateExpression(expression)
            Console.WriteLine(result)

        ElseIf command.ToUpper().StartsWith("LET ") Then
            Dim rest = command.Substring(4).Trim()
            Dim parts = rest.Split("="c)
            If parts.Length = 2 Then
                Dim variable = parts(0).Trim()
                Dim expression = parts(1).Trim()
                Dim result = EvaluateExpression(expression)
                variables(variable) = result
            Else
                Throw New Exception("Недопустимая команда LET")
            End If
        ElseIf command.ToUpper().StartsWith("IF ") Then
            ExecuteIfCommand(command)
        ElseIf command.ToUpper().StartsWith("FOR ") Then
            ExecuteForLoop(command)
        ElseIf command.ToUpper().StartsWith("WHILE ") Then
            ExecuteWhileLoop(command)
        ElseIf command.ToUpper().StartsWith("DEFINE FUNCTION ") Then
            DefineFunctionFromCommand(command)
        ElseIf command.ToUpper().StartsWith("CLASS ") Then
            DefineClass(command)
        ElseIf command.ToUpper().StartsWith("IMPORT ") Then
            ImportModule(command)
        ElseIf command.ToUpper().StartsWith("FILE ") Then
            ExecuteFileCommand(command)
        ElseIf command.ToUpper().StartsWith("THREAD ") Then
            ExecuteThreadCommand(command)
        ElseIf command.ToUpper().StartsWith("TRY ") Then
            ExecuteTryCatch(command)
        ElseIf command.ToUpper().StartsWith("GENERATOR ") Then
            ExecuteGenerator(command)
        ElseIf command.ToUpper().StartsWith("DECORATOR ") Then
            ExecuteDecorator(command)
        ElseIf command.ToUpper().StartsWith("ASYNC ") Then
            ExecuteAsync(command)
        ElseIf command.ToUpper().StartsWith("HTTP ") Then
            ExecuteHttpCommand(command)
        Else
            Throw New Exception("Неверная команда")
        End If
    End Sub

    Private Sub ExecuteDatabaseCommand(command As String)
        Throw New NotImplementedException()
    End Sub

    Private Sub ExecuteWhileLoop(command As String)
        Throw New NotImplementedException()
    End Sub

    Private Sub ExecuteForLoop(command As String)
        Throw New NotImplementedException()
    End Sub

    Private Sub ExecuteIfCommand(command As String)
        Throw New NotImplementedException()
    End Sub

    Sub DefineClass(command As String)
        Dim parts = command.Split()
        If parts.Length < 2 Then Throw New Exception("Недопустимая команда CLASS")

        Dim className = parts(1)
        Dim classDefinition = New ClassDefinition()
        classes(className) = classDefinition
        Console.WriteLine($"Класс {className} определен")
    End Sub

    Sub ImportModule(command As String)
        Dim moduleName = command.Substring(7).Trim()

        LoadModule(moduleName)
    End Sub

    Sub LoadModule(moduleName As String)
        Dim modulePath = $"{moduleName}.txt"
        If File.Exists(modulePath) Then
            Dim lines = File.ReadAllLines(modulePath)
            For Each line In lines
                ExecuteCommand(line)
            Next
        Else
            Throw New Exception($"Модуль {moduleName} не найден")
        End If
    End Sub

    Sub ExecuteFileCommand(command As String)
        Dim parts = command.Split()
        If parts.Length < 3 Then Throw New Exception("Недопустимая команда FILE")

        Dim operation = parts(1)
        Dim fileName = parts(2)

        Select Case operation.ToUpper()
            Case "READ"
                If File.Exists(fileName) Then
                    Dim content = File.ReadAllText(fileName)
                    variables("FILE_CONTENT") = content
                    Console.WriteLine($"Файл {fileName} прочитан успешно")
                Else
                    Throw New Exception($"Файл {fileName} не найден")
                End If
            Case "WRITE"
                Dim content = EvaluateExpression(command.Substring(command.IndexOf(fileName) + fileName.Length).Trim())
                File.WriteAllText(fileName, content.ToString())
                Console.WriteLine($"Файл {fileName} записан успешно")
            Case Else
                Throw New Exception("Недопустимая команда FILE")
        End Select
    End Sub

    Sub ExecuteThreadCommand(command As String)
        Dim threadCode = command.Substring(7).Trim()
        Dim thread As New Thread(Sub() ExecuteBlock(threadCode))
        thread.Start()
    End Sub

    Sub ExecuteTryCatch(command As String)
        Dim tryBlock = command.Substring(4).Trim()
        Dim catchBlock As String = Nothing

        If tryBlock.Contains(" CATCH ") Then
            Dim parts = tryBlock.Split(" CATCH ")
            tryBlock = parts(0)
            catchBlock = parts(1)
        End If

        Try
            ExecuteBlock(tryBlock)
        Catch ex As Exception
            If catchBlock IsNot Nothing Then
                ExecuteBlock(catchBlock)
            Else
                Throw
            End If
        End Try
    End Sub

    Sub ExecuteGenerator(command As String)
        Dim generatorCode = command.Substring(10).Trim()
        Dim generator = Function() EvaluateExpression(generatorCode)
        variables("GENERATOR") = generator
    End Sub

    Sub ExecuteDecorator(command As String)
        Dim parts = command.Split()
        If parts.Length < 3 Then Throw New Exception("Недопустимая команда DECORATOR")

        Dim decoratorName = parts(1)
        Dim functionName = parts(2)

        Dim originalFunction = functions(functionName)
        Dim decoratedFunction = Function(args)
                                    Console.WriteLine($"DECORATOR {decoratorName} запущен")
                                    Return originalFunction(args)
                                End Function

        functions(functionName) = decoratedFunction
    End Sub

    Async Sub ExecuteAsync(command As String)
        Dim asyncCode = command.Substring(6).Trim()
        Await Task.Run(Sub() ExecuteBlock(asyncCode))
    End Sub

    Async Sub ExecuteHttpCommand(command As String)
        Dim parts = command.Split()
        If parts.Length < 3 Then Throw New Exception("Недопустимая команда HTTP")

        Dim operation = parts(1)
        Dim url = parts(2)

        Using client As New HttpClient()
            Select Case operation.ToUpper()
                Case "GET"
                    Dim response = Await client.GetAsync(url)
                    Dim content = Await response.Content.ReadAsStringAsync()
                    Console.WriteLine(content)
                Case Else
                    Throw New Exception("Недопустимая команда HTTP")
            End Select
        End Using
    End Sub

    Sub DefineFunctionFromCommand(command As String)
        Dim parts = command.Split()
        If parts.Length < 5 OrElse parts(2).ToUpper() <> "AS" Then Throw New Exception("Недопустимая команда DEFINE FUNCTION")

        Dim funcName = parts(1)
        Dim funcBody = command.Substring(command.IndexOf("AS") + 3).Trim()

        DefineFunction(funcName, Function(args) EvaluateExpression(funcBody))
    End Sub

    Sub DefineFunction(name As String, func As Func(Of Object(), Object))
        functions(name) = func
    End Sub

    Function EvaluateExpression(expression As String) As Object
        Try
            For Each var In variables
                expression = expression.Replace(var.Key, var.Value.ToString())
            Next

            For Each func In functions
                If expression.Contains(func.Key) Then
                    Dim args = expression.Split("("c, ")"c)(1).Split(","c)
                    Dim argValues = args.Select(Function(arg) EvaluateExpression(arg.Trim())).ToArray()
                    Return func.Value(argValues)
                End If
            Next

            If expression.StartsWith("""") AndAlso expression.EndsWith("""") Then
                Return expression.Trim(""""c)
            End If

            Dim table As New DataTable()
            Dim column As New DataColumn("Eval", GetType(Double), expression)
            table.Columns.Add(column)
            table.Rows.Add(0)
            Return CDbl(table.Rows(0)("Eval"))
        Catch ex As Exception
            Throw New Exception("Ошибка вычисления выражения: " & ex.Message)
        End Try
    End Function

    Sub ExecuteBlock(block As String)
        Dim lines = block.Split(Environment.NewLine)
        For Each line In lines
            ExecuteCommand(line)
        Next
    End Sub
End Module
