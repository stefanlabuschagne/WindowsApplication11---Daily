Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.Office.Interop.Excel

Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim MSXApp As New Microsoft.Office.Interop.Excel.Application
        '' MSXApp.Visible = True

        Dim DaWorkbook As Workbook
        DaWorkbook = MSXApp.Workbooks.Add() ''Add "Book1"

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Dis nou maar die lang metode                            '
        ' want dit lyk my nie die EXCELL Interop werk te goed nie '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim DaConnection As SqlConnection = New SqlConnection
        DaConnection.ConnectionString = "Password=M33Q1Serv!ce09;Persist Security Info=True;User ID=sa;Initial Catalog=VIP_EXPANDED;Data Source=10.131.0.105"
        DaConnection.Open()

        '' Dim DaRange As Microsoft.Office.Interop.Excel.Range

        Using DaConnection

            Dim Command As SqlCommand = New SqlCommand("SELECT The_Month, GROUP_NAME,  APPLICATION_TYPE, APPLICATION_NAME, Username, Hits, DeptCode  FROM [Visitors for Unique Hits Departments]", DaConnection)
            Dim Reader As SqlDataReader = Command.ExecuteReader()

            MSXApp.Range("A1").Value2 = "The_Month"
            MSXApp.Range("B1").Value2 = "GROUP_NAME"
            MSXApp.Range("C1").Value2 = "APPLICATION_TYPE"
            MSXApp.Range("D1").Value2 = "APPLICATION_NAME"
            MSXApp.Range("E1").Value2 = "Username"
            MSXApp.Range("F1").Value2 = "Hits"
            MSXApp.Range("G1").Value2 = "DeptCode"

            '' ProgressBar1.Maximum 

            '''''''''''''''''''''''''''''''''''
            ' Populate die Excell Spreadsheet '
            '''''''''''''''''''''''''''''''''''
            Dim i As Integer '' Col Count 
            Dim DaRow As Double = 2 '' Row Count, Starts at 2

            Debug.Print(Now())

            Do While Reader.Read()
                For i = 0 To Reader.FieldCount - 1
                    MSXApp.Range(Chr(i + 65).ToString + (DaRow).ToString).Value2 = Reader.Item(i).ToString()
                Next  '' Next Column
                DaRow = DaRow + 1 '' Next Row Pointer
            Loop

            Debug.Print(Now())

            Reader.Close()

        End Using

        '''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Open excell and create a pivot table of the data! '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''

        '' Dim A1() As Object = {"OLEDB;Provider=SQLOLEDB.1;Persist Security Info=True;User ID=sa;password=M33Q1Serv!ce09;Data Source=10.131.0.105;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=95R5X1J;Use Encryption for Data=False;Tag with column collation when possible=False;Initial Catalog=VIP_EXPANDED"}

        ''                                "OLEDB;Provider=SQLOLEDB.1;Persist Security Info=True;User ID=sa;password=M33Q1Serv!ce09;Data Source=10.131.0.105;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=95R5X1J;Use Encryption for Data=False;Tag with column collation when possible=False;Initial Catalog=VIP_EXPANDED"

        '' Dim A2() As Object = {"""VIP_EXPANDED"".""dbo"".""Visitors for Unique Hits Departments"""}

        '' Impoprt the data into Worksheet1 As TABLE

        '' Dim DaRange As Range

        '' DaRange = MSXApp.Range("$A$1")




        With MSXApp



            ''With .ActiveSheet.ListObjects.Add(SourceType:=0, Source:=A1, Destination:=DaRange).QueryTable
            ''    .CommandType = XlCmdType.xlCmdTable
            ''    .CommandText = A2
            ''    .RowNumbers = False
            ''    .FillAdjacentFormulas = False
            ''    .PreserveFormatting = True
            ''    .RefreshOnFileOpen = False
            ''    .BackgroundQuery = True
            ''    .RefreshStyle = XlCellInsertionMode.xlInsertDeleteCells
            ''    .SavePassword = False
            ''    .SaveData = True
            ''    .AdjustColumnWidth = True
            ''    .RefreshPeriod = 0
            ''    .PreserveColumnInfo = True
            ''    .SourceConnectionFile = "C:\Documents and Settings\vul2214\My Documents\My Data Sources\10.131.0.105 VIP_EXPANDED Visitors for Unique Hits Departments.odc"
            ''    .ListObject.DisplayName = "Table__10.131.0.105_VIP_EXPANDED_Visitors_for_Unique_Hits_Departments"
            ''    .Refresh(BackgroundQuery:=False)
            ''End With
            ''.ActiveWorkbook.Connections("10.131.0.105 VIP_EXPANDED Visitors for Unique Hits Departments").Delete()

            '' Create the Pivottable on a new worksheet 

            ''    .Cells.Select()
            .Sheets.Add()
            .ActiveWorkbook.PivotCaches.Create(SourceType:=XlPivotTableSourceType.xlDatabase, SourceData:= _
                "Sheet1!R1C1:R100000C7", Version:=XlPivotTableVersionList.xlPivotTableVersion12).CreatePivotTable( _
                TableDestination:="Sheet4!R3C1", TableName:="PivotTable1", DefaultVersion _
                :=XlPivotTableVersionList.xlPivotTableVersion12)
            .Sheets("Sheet4").Select()
            .Cells(3, 1).Select()

            ''With .ActiveSheet.PivotTables("PivotTable1").PivotFields("The_Month")
            ''    .Orientation = XlPivotFieldOrientation.xlPageField
            ''    .Position = 1
            ''End With
            ''With .ActiveSheet.PivotTables("PivotTable1").PivotFields("APPLICATION_TYPE")
            ''    .Orientation = XlPivotFieldOrientation.xlColumnField
            ''    .Position = 1
            ''End With
            ''With .ActiveSheet.PivotTables("PivotTable1").PivotFields("GROUP_NAME")
            ''    .Orientation = XlPivotFieldOrientation.xlRowField
            ''    .Position = 1
            ''End With
            ''.ActiveSheet.PivotTables("PivotTable1").AddDataField(.ActiveSheet.PivotTables( _
            ''    "PivotTable1").PivotFields("GROUP_NAME"), "Count of GROUP_NAME", XlConsolidationFunction.xlCount)

        End With '' Application


        '' The Old STuff.....


        '' Dim AppData2() As String = {"""VIP_EXPANDED"".""dbo"".""Visitors for Unique Hits Departments"""}



        '        Dim wbc As WorkbookConnection
        '        wbc = MSXApp.ActiveWorkbook.Connections.Add("10.131.0.105 VIP_EXPANDED Visitors for Unique Hits Departments", "", A1, A2, 3)

        With MSXApp

            ''.ActiveWorkbook.PivotCaches.Create(SourceType:=XlPivotTableSourceType.xlExternal, SourceData:= _
            ''    .ActiveWorkbook.Connections( _
            ''    "10.131.0.105 VIP_EXPANDED Visitors for Unique Hits Departments"), Version:= _
            ''    XlPivotTableVersionList.xlPivotTableVersion12).CreatePivotTable(TableDestination:="Sheet1!R1C1", _
            ''    TableName:="PivotTable1", DefaultVersion:=XlPivotTableVersionList.xlPivotTableVersion12)
            ''.Cells(1, 1).Select()



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

            '' Remove the database connection 
            '' MSXApp.ActiveWorkbook.Connections(1).Delete()

        End With '' MSXApp

        '' Send as an Attachment to Me, Wessel and Elmari. 100% 
        DaWorkbook.SendMail("stefan.labuschagne@treasury.gov.za", "Monthly Usage Report")

        MSXApp.Workbooks.Close()
        MSXApp.Quit()


    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        '' GO !!!!!!

        Dim MSXApp As New Microsoft.Office.Interop.Excel.Application
        '' MSXApp.Visible = True

        Dim DaWorkbook As Workbook
        DaWorkbook = MSXApp.Workbooks.Add() ''Add "Book1"

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Dis nou maar die lang metode                            '
        ' want dit lyk my nie die EXCELL Interop werk te goed nie '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim DaConnection As SqlConnection = New SqlConnection
        DaConnection.ConnectionString = "Password=M33Q1Serv!ce09;Persist Security Info=True;User ID=sa;Initial Catalog=VIP_EXPANDED;Data Source=10.131.0.105"
        DaConnection.Open()

        '' Dim DaRange As Microsoft.Office.Interop.Excel.Range

        ProgressBar1.Maximum = 25000

        Using DaConnection

            Dim Command As SqlCommand = New SqlCommand("SELECT The_Month, GROUP_NAME,  APPLICATION_TYPE, APPLICATION_NAME, Username, Hits, DeptCode  FROM [Visitors for Unique Hits Departments]", DaConnection)
            Dim Reader As SqlDataReader = Command.ExecuteReader()

            MSXApp.Range("A1").Value2 = "The_Month"
            MSXApp.Range("B1").Value2 = "GROUP_NAME"
            MSXApp.Range("C1").Value2 = "APPLICATION_TYPE"
            MSXApp.Range("D1").Value2 = "APPLICATION_NAME"
            MSXApp.Range("E1").Value2 = "Username"
            MSXApp.Range("F1").Value2 = "Hits"
            MSXApp.Range("G1").Value2 = "DeptCode"

            ProgressBar1.Maximum = 25000

            '' ProgressBar1.Maximum 

            '''''''''''''''''''''''''''''''''''
            ' Populate die Excell Spreadsheet '
            '''''''''''''''''''''''''''''''''''
            Dim i As Integer '' Col Count 
            Dim DaRow As Double = 2 '' Row Count, Starts at 2

            Debug.Print(Now())

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
                ProgressBar1.Increment(1)

            Loop

            Debug.Print(Now())

            Reader.Close()

        End Using

        '''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Open excell and create a pivot table of the data! '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''

        With MSXApp

            .Sheets.Add()
            .ActiveWorkbook.PivotCaches.Create(SourceType:=XlPivotTableSourceType.xlDatabase, SourceData:= _
                "Sheet1!R1C1:R100000C7", Version:=XlPivotTableVersionList.xlPivotTableVersion11).CreatePivotTable( _
                TableDestination:="Sheet4!R3C1", TableName:="PivotTable1", DefaultVersion _
                :=XlPivotTableVersionList.xlPivotTableVersion11)
            .Sheets("Sheet4").Select()
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

        ''''''''''''''''''''''''''''''''
        ' Then Save as an Excell sheet '
        ' IN 2003 Format (.xls)        ;
        ''''''''''''''''''''''''''''''''
        MSXApp.Workbooks(1).SaveAs("C:\TheWorkbook.xls", XlFileFormat.xlWorkbookNormal, , , , , , , , , , )

        '' Send as an Attachment to Me, Wessel and Elmari. 100% 
        DaWorkbook.SendMail("stefan.labuschagne@treasury.gov.za;wessel.husselman@treasury.gov.za", "Monthly Usage Report: " + Now())

        MSXApp.Workbooks.Close()
        MSXApp.Quit()

        '' Delete the Temp File Then
        Try
            FileSystem.Kill("C:\TheWorkbook.xls")
        Catch ex As Exception
            '' the file does not exist
        End Try

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Form1_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Module1.Main()

    End Sub
End Class
