Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.Office.Interop.Excel
Imports System.Net.Mail

Module Module1

    Public Sub Main()

        '' GO !!!!!!

        Dim MSXApp As New Microsoft.Office.Interop.Excel.Application
        MSXApp.Visible = True

        Dim DaWorkbook As Workbook
        DaWorkbook = MSXApp.Workbooks.Add() ''Add "Book1"



        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Dis nou maar die lang metode                            '
        ' want dit lyk my nie die EXCELL Interop werk te goed nie '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim DaConnection As SqlConnection = New SqlConnection
        '' DaConnection.ConnectionString = "Password=M33Q1Serv!ce09;Persist Security Info=True;User ID=sa;Initial Catalog=VIP_EXPANDED;Data Source=10.131.0.105"

        ''''''''''''''''''''''''''
        '' Connection String changed 
        '' want al die data lê nou op die nuwe PRD server 
        '' SL 30 April 2008
        '''''''''''''''''''''''''''''''
        '' DaConnection.ConnectionString = "Password=M33Q1Serv!ce09;Persist Security Info=True;User ID=sa;Initial Catalog=VIP_EXPANDED;Data Source=10.123.45.180"

        '''''''''''''''''''''''''''''''''''''''''
        ' We Moded to the New Productuion Sever '
        ' 10.125.153.14 - Database (WebSQL)    Vul_WEB_SQL  '
        '''''''''''''''''''''''''''''''''''''''''
        DaConnection.ConnectionString = "Password=S@P@SSW0RD;Persist Security Info=True;User ID=sa;Initial Catalog=VIP_EXPANDED;Data Source=10.125.153.14"

        DaConnection.Open()

        '' Dim DaRange As Microsoft.Office.Interop.Excel.Range
        '' ProgressBar1.Maximum = 25000

        Using DaConnection

            Dim Command As SqlCommand = New SqlCommand("SELECT The_Month, GROUP_NAME,  APPLICATION_TYPE, APPLICATION_NAME, Username, Hits, DeptCode  FROM [Visitors for Unique Hits Departments]   where [The_Month] = cast(year(getdate()) as varchar(100))   +  right('00' + cast(month(getdate()) as varchar(100) ),2)", DaConnection)
            Command.CommandTimeout = 600
            Dim Reader As SqlDataReader = Command.ExecuteReader()

            MSXApp.Range("A1").Value2 = "The_Month"
            MSXApp.Range("B1").Value2 = "GROUP_NAME"
            MSXApp.Range("C1").Value2 = "APPLICATION_TYPE"
            MSXApp.Range("D1").Value2 = "APPLICATION_NAME"
            MSXApp.Range("E1").Value2 = "Username"
            MSXApp.Range("F1").Value2 = "Hits"
            MSXApp.Range("G1").Value2 = "DeptCode"

            '' Form1.ProgressBar1.Maximum = 25000

            '''''''''''''''''''''''''''''''''''
            ' Populate die Excell Spreadsheet '
            '''''''''''''''''''''''''''''''''''
            Dim i As Integer '' Col Count 
            Dim DaRow As Double = 2 '' Row Count, Starts at 2

            Console.Write(Now() + vbCrLf)

            Do While Reader.Read()
                For i = 0 To Reader.FieldCount - 1

                    If i = 2 Then
                        '' This could be resolved @ Database Level!
                        Select Case Reader.Item(i).ToString()
                            Case "D"
                                MSXApp.Range(Chr(i + 65).ToString + (DaRow).ToString).Value2 = "Download"
                            Case "H"
                                MSXApp.Range(Chr(i + 65).ToString + (DaRow).ToString).Value2 = "History"
                            Case "C"
                                MSXApp.Range(Chr(i + 65).ToString + (DaRow).ToString).Value2 = "Current"
                            Case Else
                                MSXApp.Range(Chr(i + 65).ToString + (DaRow).ToString).Value2 = Reader.Item(i).ToString()
                        End Select
                    Else
                        MSXApp.Range(Chr(i + 65).ToString + (DaRow).ToString).Value2 = Reader.Item(i).ToString()
                    End If



                Next  '' Next Column
                DaRow = DaRow + 1 '' Next Row Pointer
                '' Form1.ProgressBar1.Increment(1)

                '' MSXApp.Visible = True

                Console.Write("Importing Row:" + DaRow.ToString() + vbCrLf)

            Loop

            Console.Write(Now() + vbCrLf)

            Reader.Close()

        End Using

        '''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Open excell and create a pivot table of the data! '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''

        With MSXApp

            .Sheets.Add()


            '.ActiveWorkbook.PivotCaches.Create(SourceType:=XlPivotTableSourceType.xlDatabase, SourceData:= _
            '    "'Sheet1'!R1C1:R100000C7", Version:=XlPivotTableVersionList.xlPivotTableVersion11).CreatePivotTable( _
            '    TableDestination:="'Sheet4'!R3C1", TableName:="PivotTable1", DefaultVersion _
            '    :=XlPivotTableVersionList.xlPivotTableVersion11)

            .ActiveWorkbook.PivotCaches.Create(SourceType:=XlPivotTableSourceType.xlDatabase, SourceData:= _
                "'Sheet1'!R1C1:R100000C7", Version:=XlPivotTableVersionList.xlPivotTableVersion11).CreatePivotTable( _
                TableDestination:="R3C1", TableName:="PivotTable1", DefaultVersion _
                :=XlPivotTableVersionList.xlPivotTableVersion11)




            .Sheets("Sheet2").Select()
            .Cells(3, 1).Select()

        End With '' Application

        '' DISABLE SUBTOTALS
        Dim A1() = {False, False, False, False, False, False, False, False, False, False, False, False}
        MSXApp.ActiveSheet.PivotTables("PivotTable1").PivotFields("GROUP_NAME").Subtotals = A1
        MSXApp.ActiveSheet.PivotTables("PivotTable1").PivotFields("APPLICATION_TYPE").Subtotals = A1

        With MSXApp

            With .ActiveSheet.PivotTables("PivotTable1").PivotFields("The_Month")
                .Orientation = XlPivotFieldOrientation.xlColumnField
                .Position = 1
            End With
            .ActiveSheet.PivotTables("PivotTable1").AddDataField(.ActiveSheet.PivotTables( _
                "PivotTable1").PivotFields("Hits"), "Count of Hits", XlConsolidationFunction.xlCount)
            With .ActiveSheet.PivotTables("PivotTable1").PivotFields("GROUP_NAME")
                .Orientation = XlPivotFieldOrientation.xlRowField
                .Position = 1
            End With
            With .ActiveSheet.PivotTables("PivotTable1").PivotFields("APPLICATION_TYPE")
                .Orientation = XlPivotFieldOrientation.xlRowField
                .Position = 2
            End With
            With .ActiveSheet.PivotTables("PivotTable1").PivotFields("APPLICATION_NAME")
                .Orientation = XlPivotFieldOrientation.xlRowField
                .Position = 3
            End With

            With .ActiveSheet.PivotTables("PivotTable1")
                .ColumnGrand = True
                .NullString = "0"
                .RowGrand = True
                .FieldListSortAscending = False
            End With

        End With '' MSXApp

        DaWorkbook.ShowPivotTableFieldList = False

        '' Delete the Temp File Then
        Dim DaFileName As String = "C:\tst2\TheWorkbook MONTHLY " + DateTime.Today.Year.ToString() + "-" + DateTime.Today.Month.ToString() + " .xlsx"

        '' Try 1
        Try
            FileSystem.Kill(DaFileName)
        Catch ex As Exception
            '' the file does not exist
            Console.WriteLine(ex.Message)
        End Try

        '' Try 2 - Cleanout the directory!
        Try
            FileSystem.Kill("C:\tst2\*.*")
        Catch ex As Exception
            '' the file does not exist
        End Try

        ''''''''''''''''''''''''''''''''
        ' Then Save as an Excell sheet '
        ' IN 2003 Format (.xls)        ;
        ''''''''''''''''''''''''''''''''
        Dim TT = "Hello"

        Try
            MSXApp.Workbooks(1).SaveAs(DaFileName, XlFileFormat.xlOpenXMLWorkbook, , , , , , , , , , )

            Console.WriteLine("File saved OK : " + DaFileName)

        Catch ex As Exception

            Console.WriteLine("File saved FAILED : " + DaFileName + " " + ex.Message)

            '' Change Filename and save again
            DaFileName = "C:\tst2\TheWorkbook1.xlsx"
            MSXApp.Workbooks(1).SaveAs(DaFileName, XlFileFormat.xlOpenXMLWorkbook, , , , , , , , , , )

            '' 


        End Try

        '' This is just a test for now......
        ''  MSXApp.Workbooks(1).SendMail("stefan.labuschagne@treasury.gov.za", "Monthly Usage Report: " + Now())

        MSXApp.Workbooks.Close()
        MSXApp.Quit()

        MSXApp = Nothing


        ''''''''''''''''''''''''''
        ''Now Zip The File      ''
        '' Not Needed anymore ! ''
        ''''''''''''''''''''''''''
        ''Dim DaZipFile As String = "c:\TheWorkbook.zip"
        ''Dim DaDestZipFile As String = "c:\tst2\TheWorkbook.zip"

        ''Try
        ''    '' JUST IN CASE THAT IT ALREADY EXISTS
        ''    FileSystem.Kill(DaZipFile)
        ''Catch ex As Exception


        ''End Try

        '''' Create The Zip File
        ''System.IO.Compression.ZipFile.CreateFromDirectory("c:\tst2\", DaZipFile)

        '''' 
        ''System.IO.File.Move(DaZipFile, DaDestZipFile)


        '' Send the Mail with Excell directly!?

        Dim DaMessage As MailMessage = New MailMessage()
        Dim DaClient As SmtpClient = New SmtpClient()

        DaMessage.From = New MailAddress("Stefan.Labuschagne@treasury.gov.za")

        '' DaMessage.To.Add(New MailAddress("Anna-marie.Pienaar@treasury.gov.za"))
        '' DaMessage.To.Add(New MailAddress("Elmari.DeWitt@treasury.gov.za"))
        '' DaMessage.To.Add(New MailAddress("Elaine.Eybers@treasury.gov.za"))
        DaMessage.Bcc.Add(New MailAddress("Stefan.Labuschagne@treasury.gov.za"))
        DaMessage.Bcc.Add(New MailAddress("Wessel.Husselman@treasury.gov.za"))
        '' DaMessage.Bcc.Add(New MailAddress("Wessel.Husselman@treasury.gov.za"))
        '' DaMessage.Bcc.Add(New MailAddress("Wessel.Husselman@treasury.gov.za"))

        DaMessage.Subject = "Vulindlela Usage Report: " + DateAndTime.MonthName(DateTime.Now.Month) + " " + Now()
        DaMessage.Body = ""


        '' EMAIL THE FILE HERE!        
        DaMessage.Attachments.Add(New Attachment(DaFileName))

        '' DaMessage.Attachments.Add(New Attachment(DaDestZipFile))

        '' vul_smtp.vulindlelaprd.gov.za

        '' Hi Stefan
        '' Probeer eers die DNS naam
        '' vul_smtp.vulindlelaprd.gov.za
        '' andersins
        '' 10.125.153.25

        DaClient.Host = "10.125.153.25"


        '' DaClient.Host = "10.125.153.35"    '' "cenexc01.treasury.gov.za"
        '' DaClient.Host = "10.131.12.118"
        ''DaClient.Port = 25
        ''DaClient.Host = "localhost"

        DaClient.DeliveryMethod = SmtpDeliveryMethod.Network


        Try
            DaClient.Send(DaMessage)
        Catch ex As Exception
            Console.Write(ex.Message)
            ''DaClient.Send(DaMessage)
        End Try
        ''''''''''''''''''''''''''
        '' Delete the Temp File ''
        ''''''''''''''''''''''''''
        Try
            '' FileSystem.Kill(DaFileName)
        Catch ex As Exception
            '' the file does not exist
        End Try

        Console.Write("Press Enter to continue...")
        Console.ReadLine()

    End Sub

End Module
