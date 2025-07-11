Imports System
Imports System.IO
Imports System.Collections.Generic

Module DdsExtractor
    ' DDS 文件头标识
    Private ReadOnly DDS_HEADER As Byte() = {&H44, &H44, &H53, &H20} ' "DDS "
    Private ReadOnly POF_MARKER As String = "POF"

    Sub Main()
        Console.WriteLine("DDS 文件提取工具")
        Console.WriteLine("请拖放要处理的 .afb 或 .svo 文件到窗口，或输入文件路径(支持多个文件)")
        Console.WriteLine("输入 'exit' 退出程序")

        ' 持续处理循环
        While True
            Console.WriteLine()
            Console.Write("> ")
            Dim input As String = Console.ReadLine()

            ' 检查退出命令
            If input.Trim().ToLower() = "exit" Then
                Exit While
            End If

            ' 处理输入的文件
            ProcessInputFiles(input)
        End While

        Console.WriteLine("程序已退出")
    End Sub

    Private Sub ProcessInputFiles(input As String)
        ' 处理拖放的文件路径(Windows终端会用引号包裹带空格的文件路径)
        Dim filePaths As New List(Of String)()
        Dim inQuotes As Boolean = False
        Dim currentPath As New System.Text.StringBuilder()

        For Each c As Char In input
            If c = """"c Then
                If inQuotes Then
                    ' 结束引号包裹的路径
                    filePaths.Add(currentPath.ToString())
                    currentPath.Clear()
                    inQuotes = False
                Else
                    ' 开始引号包裹的路径
                    inQuotes = True
                End If
            ElseIf Not inQuotes AndAlso Char.IsWhiteSpace(c) Then
                ' 非引号包裹的空格分隔符
                If currentPath.Length > 0 Then
                    filePaths.Add(currentPath.ToString())
                    currentPath.Clear()
                End If
            Else
                ' 添加到当前路径
                currentPath.Append(c)
            End If
        Next

        ' 添加最后一个路径
        If currentPath.Length > 0 Then
            filePaths.Add(currentPath.ToString())
        End If

        ' 处理每个文件
        For Each filePath In filePaths
            If Not String.IsNullOrWhiteSpace(filePath) Then
                Try
                    ProcessFile(filePath)
                Catch ex As Exception
                    Console.WriteLine($"处理文件 {filePath} 时出错: {ex.Message}")
                End Try
            End If
        Next
    End Sub

    Private Sub ProcessFile(filePath As String)
        If Not File.Exists(filePath) Then
            Console.WriteLine($"文件不存在: {filePath}")
            Return
        End If

        Dim extension As String = Path.GetExtension(filePath).ToLower()
        If extension <> ".afb" AndAlso extension <> ".svo" Then
            Console.WriteLine($"不支持的文件类型: {filePath} (仅支持 .afb 和 .svo)")
            Return
        End If

        Console.WriteLine($"正在处理文件: {filePath}")

        Dim fileData As Byte() = File.ReadAllBytes(filePath)
        Dim ddsList As List(Of Byte()) = ExtractDdsFiles(fileData, extension = ".afb")

        Console.WriteLine($"找到 {ddsList.Count} 个 DDS 文件")

        ' 保存提取的 DDS 文件
        Dim baseName As String = Path.GetFileNameWithoutExtension(filePath)
        Dim outputDir As String = Path.Combine(Path.GetDirectoryName(filePath), $"{baseName}_extracted")

        Directory.CreateDirectory(outputDir)

        For i As Integer = 0 To ddsList.Count - 1
            Dim outputPath As String = Path.Combine(outputDir, $"{baseName}_{i + 1}.dds")
            File.WriteAllBytes(outputPath, ddsList(i))
            Console.WriteLine($"已保存: {outputPath}")
        Next
    End Sub

    Private Function ExtractDdsFiles(fileData As Byte(), isAfbFile As Boolean) As List(Of Byte())
        Dim ddsFiles As New List(Of Byte())()
        Dim position As Integer = 0

        While position < fileData.Length - 4
            ' 检查是否是 DDS 文件头
            If fileData(position) = DDS_HEADER(0) AndAlso
               fileData(position + 1) = DDS_HEADER(1) AndAlso
               fileData(position + 2) = DDS_HEADER(2) AndAlso
               fileData(position + 3) = DDS_HEADER(3) Then

                ' 查找下一个 DDS 文件头或结束标记
                Dim nextDdsPos As Integer = FindNextDdsHeader(fileData, position + 4)
                Dim endPos As Integer = If(nextDdsPos <> -1, nextDdsPos, fileData.Length)

                ' 对于 AFB 文件，检查是否有 POF 标记
                If isAfbFile AndAlso nextDdsPos = -1 Then
                    Dim pofPos As Integer = FindPofMarker(fileData, position + 4)
                    If pofPos <> -1 Then
                        endPos = pofPos
                    End If
                End If

                ' 提取 DDS 数据
                Dim ddsLength As Integer = endPos - position
                Dim ddsData(ddsLength - 1) As Byte
                Array.Copy(fileData, position, ddsData, 0, ddsLength)
                ddsFiles.Add(ddsData)

                position = endPos
            Else
                position += 1
            End If
        End While

        Return ddsFiles
    End Function

    Private Function FindNextDdsHeader(data As Byte(), startPos As Integer) As Integer
        For i As Integer = startPos To data.Length - 4
            If data(i) = DDS_HEADER(0) AndAlso
               data(i + 1) = DDS_HEADER(1) AndAlso
               data(i + 2) = DDS_HEADER(2) AndAlso
               data(i + 3) = DDS_HEADER(3) Then
                Return i
            End If
        Next
        Return -1
    End Function

    Private Function FindPofMarker(data As Byte(), startPos As Integer) As Integer
        ' POF 标记是 ASCII 字符串 "POF"
        For i As Integer = startPos To data.Length - 3
            If data(i) = AscW("P"c) AndAlso
               data(i + 1) = AscW("O"c) AndAlso
               data(i + 2) = AscW("F"c) Then
                Return i
            End If
        Next
        Return -1
    End Function
End Module