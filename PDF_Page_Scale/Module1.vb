Imports System.IO
Imports System.Text.RegularExpressions
Imports iTextSharp.text
Imports iTextSharp.text.pdf

Module Module1

    Dim inputfile As String = ""
    Dim outputfile As String = ""

    Dim myScale As Single = 1 'default no scale
    Dim myCustomPageSize As Rectangle = PageSize.A4

    Dim myExe As String = Path.GetFileName(System.Reflection.Assembly.GetExecutingAssembly().Location)

    Sub WriteError(s As String)
        Console.ForegroundColor = ConsoleColor.Red
        Console.Error.Write(s)
        Console.ResetColor()
    End Sub
    Sub WriteLineError(s As String)
        Console.ForegroundColor = ConsoleColor.Red
        Console.Error.WriteLine(s)
        Console.ResetColor()
    End Sub

    Function Main() As Integer

        'print usage
        If My.Application.CommandLineArgs.Count < 2 Then
            Console.WriteLine("Resize all pages to A4 in a pdf and save the output to a new file. Page rotation will not change." & vbNewLine)
            '            Console.WriteLine("NOTE: PDF Bookmars are currently not preserved.")
            Console.WriteLine("USAGE: " & myExe & " <input file> <output file> [content_scale_factor [custom_page_size] ]" & vbNewLine)
            Console.WriteLine("   content_scale_factor is expressed in percentage (without % mark), default is 100% (fit to page)" & vbNewLine)
            Console.WriteLine("   if using custom_page_size parameter, you must specify one of the known page sizes (https://afterlogic.com/mailbee-net/docs-itextsharp/html/d37ebb4c-4453-ad77-c842-ba4e5b252a78.htm) or <W>x<H>mm or <W>x<H>in or <W>x<H>pt. This parameter is not case sensitive." & vbNewLine)
            Console.WriteLine(" example: " & myExe & "myfancypdffile.pdf mytinypdffile 50 200x300mm   --> shrinks the content to half and uses 200x300mm paper" & vbNewLine)
            If Debugger.IsAttached Then
                Console.WriteLine("Press any key to exit.")
                Console.ReadKey()
            End If
            Return -1
        End If

        'parameter parsing
        inputfile = My.Application.CommandLineArgs(0)
        outputfile = My.Application.CommandLineArgs(1)
        If My.Application.CommandLineArgs.Count > 2 Then
            Try
                Dim percent As Single = My.Application.CommandLineArgs(2)
                myScale = myScale * (percent / 100)
                Console.WriteLine("content scale factor: " & percent & "%")
            Catch ex As Exception
                WriteLineError("Invalid parameter! Scale factor must be a number." & vbNewLine & ex.Message)
                If Debugger.IsAttached Then
                    Console.WriteLine("Press any key to exit.")
                    Console.ReadKey()
                End If
                Return -2
            End Try

            'custom page size possibility
            Dim exactsize As String = ""
            Try
                If My.Application.CommandLineArgs.Count > 3 Then

                    exactsize = My.Application.CommandLineArgs(3).ToUpper
                    Try
                        myCustomPageSize = PageSize.GetRectangle(exactsize)
                        Console.WriteLine("custom page size set: " & exactsize)
                    Catch ex As ArgumentException
                        'if not known size, then throws exception
                        ' try to parse as one of these: millimeter, inch, postscript point
                        Dim r_mm, r_in, r_pt As Regex
                        r_mm = New Regex("(?<W>\d+?)X(?<H>\d+?)MM")   ' eg:  200x100mm
                        r_in = New Regex("(?<W>\d+?)X(?<H>\d+?)IN")   ' eg:  100x200in
                        r_pt = New Regex("(?<W>\d+?)X(?<H>\d+?)PT")   ' eg:  300x400pt
                        If r_mm.IsMatch(exactsize) Then
                            myCustomPageSize = New Rectangle(Utilities.MillimetersToPoints(r_mm.Matches(exactsize)(0).Groups("W").Value), Utilities.MillimetersToPoints(r_mm.Matches(exactsize)(0).Groups("H").Value))
                            Console.WriteLine("custom page size set: " & Utilities.PointsToMillimeters(myCustomPageSize.Width) & " x " & Utilities.PointsToMillimeters(myCustomPageSize.Height) & " mm")
                        ElseIf r_in.IsMatch(exactsize) Then
                            myCustomPageSize = New Rectangle(Utilities.InchesToPoints(r_in.Matches(exactsize)(0).Groups("W").Value), Utilities.InchesToPoints(r_in.Matches(exactsize)(0).Groups("H").Value))
                            Console.WriteLine("custom page size set: " & Utilities.PointsToInches(myCustomPageSize.Width) & " x " & Utilities.PointsToInches(myCustomPageSize.Height) & " in")
                        ElseIf r_pt.IsMatch(exactsize) Then
                            myCustomPageSize = New Rectangle(r_pt.Matches(exactsize)(0).Groups("W").Value, r_pt.Matches(exactsize)(0).Groups("H").Value)
                            Console.WriteLine("custom page size set: " & myCustomPageSize.Width & " x " & myCustomPageSize.Height & " pt")
                        Else
                            Throw ex  'if cannot identify the page size, throw exception
                        End If
                    End Try

                End If
            Catch ex As Exception
                WriteLineError("Invalid parameter! Unidentifiable page size: " & exactsize & vbNewLine & ex.Message)
                If Debugger.IsAttached Then
                    Console.WriteLine("Press any key to exit.")
                    Console.ReadKey()
                End If
                Return -3
            End Try
        End If



        Dim doc As Document = New Document(myCustomPageSize)
        Dim ms As MemoryStream = New MemoryStream
        Dim writer As PdfWriter = PdfWriter.GetInstance(doc, ms)



        'eliminate password protected pdf issue
        PdfReader.unethicalreading = True

        Try
            Dim reader As New PdfReader(New MemoryStream(File.ReadAllBytes(inputfile)))
            reader.ConsolidateNamedDestinations()

            'save bookmarks - will be added back later
            Dim bookm As IList(Of Dictionary(Of String, Object)) = SimpleBookmark.GetBookmark(reader)

            Console.WriteLine("Opening file: " & inputfile)
            doc.Open()

            Dim cb As PdfContentByte = writer.DirectContent

            For pageNumber = 1 To reader.NumberOfPages
                Console.WriteLine("Processing page: " & pageNumber)
                Dim page As PdfImportedPage = writer.GetImportedPage(reader, pageNumber)
                'Console.WriteLine("page rotation: " & page.Rotation)
                If page.Width <= page.Height Then
                    doc.SetPageSize(myCustomPageSize)
                Else
                    doc.SetPageSize(myCustomPageSize.Rotate())
                End If
                doc.NewPage()

                Dim widthFactor = doc.PageSize.Width / page.Width
                Dim heightFactor = doc.PageSize.Height / page.Height
                Dim factor = Math.Min(widthFactor, heightFactor) * myScale
                Dim offsetX = (doc.PageSize.Width - (page.Width * factor)) / 2
                Dim offsetY = (doc.PageSize.Height - (page.Height * factor)) / 2
                cb.AddTemplate(page, factor, 0, 0, factor, offsetX, offsetY)


                'fix rotations (if VIEW is rotated in the original file instead of page WIDTH/HEIGHT and CONTENT):
                Select Case page.Rotation
                    Case 90
                        writer.AddPageDictEntry(PdfName.ROTATE, PdfPage.LANDSCAPE)
                    Case 180
                        writer.AddPageDictEntry(PdfName.ROTATE, PdfPage.INVERTEDPORTRAIT)
                    Case 270
                        writer.AddPageDictEntry(PdfName.ROTATE, PdfPage.SEASCAPE)
                End Select

            Next

            doc.Close()
            reader.Close()
            Console.WriteLine("Saving file: " & outputfile)
            File.WriteAllBytes(outputfile, ms.GetBuffer())

            'add back bookmarks
            If Not IsNothing(bookm) AndAlso bookm.Count > 0 Then
                reader = New PdfReader(New MemoryStream(File.ReadAllBytes(outputfile)))

                Dim stamper As PdfStamper = New PdfStamper(reader, New FileStream(outputfile & ".tmp.pdf", FileMode.OpenOrCreate))
                ' do stuff
                stamper.Outlines = bookm
                stamper.Close()
                reader.Close()

                'some playing with filenames and backup files, because itext cannot modify pdf file inplace
                My.Computer.FileSystem.RenameFile(outputfile, Path.GetFileName(outputfile) & ".bak")
                Threading.Thread.Sleep(100) 'delay to prevent file remeaining in place
                My.Computer.FileSystem.RenameFile(outputfile & ".tmp.pdf", Path.GetFileName(outputfile))
                Threading.Thread.Sleep(100) 'delay to prevent file remeaining in place
                My.Computer.FileSystem.DeleteFile(outputfile & ".bak")
                Console.WriteLine("Restored " & bookm.Count & " bookmarks.")
            End If


        Catch ex As Exception
            WriteLineError("Error: " & vbNewLine & ex.Message)
        If Debugger.IsAttached Then
            Console.WriteLine("Press any key to exit.")
            Console.ReadKey()
        End If
        Return 100
        End Try

        Console.WriteLine("All done.")
        If Debugger.IsAttached Then
            Console.WriteLine("Press any key to exit.")
            Console.ReadKey()
        End If

        Return 0
    End Function

End Module
