Imports Excel = Microsoft.Office.Interop.Excel

Public Class Form1
    Dim fileSelected As Boolean
    Dim xlApp As Excel.Application
    Dim xlWorkBook As Excel.Workbook
    Dim xlWorkSheet As Excel.Worksheet
    Dim loopCheker As Boolean = False
    Dim loopChecker2 As Boolean = False
    Dim run2Row As Integer = 1
    Dim run3Row As Integer = 1
    Dim run4Row As Integer = 1
    Dim endRow As Integer

    Dim run2Addresses As New List(Of String)
    Dim run2numbers As New List(Of Integer)
    Dim run3Addresses As New List(Of String)
    Dim run3numbers As New List(Of Integer)
    Dim run4Addresses As New List(Of String)
    Dim run4numbers As New List(Of Integer)

    Dim cellAddress As String
    Dim addressCounter As Integer = 0
    Dim counter As Integer = 0
    Dim firstLoop As Boolean = True
    Dim cellRun As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub RunButton_Click(sender As Object, e As EventArgs) Handles runButton.Click

        'set button colour to light green to show the user has clicked and the program is working
        runButton.BackColor = Color.LightSkyBlue
        If fileSelected = True Then

            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Open(OpenFileDialog.FileName)
            xlWorkSheet = xlWorkBook.Worksheets("Sheet")

            ResetVarsForLoops()
            run2Row = 0
            Do While loopCheker = False
                Dim run2Cell As String
                run2Row += 1
                run2Cell = ("A" & run2Row)

                If xlWorkSheet.Range(run2Cell).Value = "Run 2" Then
                    loopCheker = True

                ElseIf run2Row >= 1000 Then
                    loopCheker = True

                End If
            Loop

            ResetVarsForLoops()
            run3Row = 0
            Do While loopCheker = False
                Dim run3Cell As String
                run3Row += 1
                run3Cell = ("A" & run3Row)

                If xlWorkSheet.Range(run3Cell).Value = "Run 3" Then
                    loopCheker = True

                ElseIf run3Row >= 1000 Then
                    loopCheker = True

                End If
            Loop

            ResetVarsForLoops()
            run4Row = 0
            Do While loopCheker = False
                Dim run4Cell As String
                run4Row += 1
                run4Cell = ("A" & run4Row)

                If xlWorkSheet.Range(run4Cell).Value = "Run 4" Then
                    loopCheker = True

                ElseIf run4Row >= 1000 Then
                    loopCheker = True

                End If
            Loop

            ResetVarsForLoops()
            endRow = 0
            Do While loopCheker = False
                Dim endRowCell As String
                endRow += 1
                endRowCell = ("A" & endRow)

                If xlWorkSheet.Range(endRowCell).Value = "Powered by Fleetmatics WORK" Then
                    loopCheker = True

                ElseIf endRow >= 1000 Then
                    loopCheker = True

                End If
            Loop

            ResetVarsForLoops()
            Do Until loopCheker = True
                cellAddress = ("B" & run2Row + addressCounter)

                If xlWorkSheet.Range(cellAddress).Row >= run3Row Then
                    loopCheker = True

                ElseIf String.IsNullOrEmpty(xlWorkSheet.Range(cellAddress).Value) = False Then
                    If firstLoop = False Then
                        run2numbers.Add(counter)

                    End If
                    firstLoop = False
                    counter = -1
                    run2Addresses.Add(xlWorkSheet.Range(cellAddress).Value)

                ElseIf String.IsNullOrEmpty(xlWorkSheet.Range(cellAddress).Value) Then
                    counter += 1

                End If
                addressCounter += 1

            Loop

            ResetVarsForLoops()
            Do Until loopCheker = True
                cellAddress = ("B" & run3Row + addressCounter)

                If xlWorkSheet.Range(cellAddress).Row >= run4Row Then
                    loopCheker = True

                ElseIf String.IsNullOrEmpty(xlWorkSheet.Range(cellAddress).Value) = False Then
                    If firstLoop = False Then
                        run3numbers.Add(counter)

                    End If
                    firstLoop = False
                    counter = -1
                    run3Addresses.Add(xlWorkSheet.Range(cellAddress).Value)

                ElseIf String.IsNullOrEmpty(xlWorkSheet.Range(cellAddress).Value) Then
                    counter += 1

                End If
                addressCounter += 1

            Loop

            ResetVarsForLoops()
            Do Until loopCheker = True
                cellAddress = ("B" & run4Row + addressCounter)

                If xlWorkSheet.Range(cellAddress).Row >= endRow Then
                    loopCheker = True

                ElseIf String.IsNullOrEmpty(xlWorkSheet.Range(cellAddress).Value) = False Then
                    If firstLoop = False Then
                        run4numbers.Add(counter)

                    End If
                    firstLoop = False
                    counter = -1
                    run4Addresses.Add(xlWorkSheet.Range(cellAddress).Value)

                ElseIf String.IsNullOrEmpty(xlWorkSheet.Range(cellAddress).Value) Then
                    counter += 1

                End If
                addressCounter += 1

            Loop

            xlWorkBook.Close()
            xlApp.Quit()
            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheet)

            runButton.BackColor = Color.LightGreen

        Else
            runButton.BackColor = Color.Red

        End If

    End Sub


    Private Sub FileSelectButton_Click(sender As Object, e As EventArgs) Handles fileSelectButton.Click
        'set button colour to light blue to show the user has clicked and the program is running
        fileSelectButton.BackColor = Color.LightSkyBlue

        'setup windows file picker to show users downloads folder & only .xlxs files to prevent errors
        OpenFileDialog.Title = "Select the monthly report"
        Dim strUser As String
        strUser = Environ("username")
        OpenFileDialog.InitialDirectory = "C:\Users\" & strUser & "\downloads"
        ' OpenFileDialog.Filter = "Excel File|*.xlxs"
        OpenFileDialog.Multiselect = False

        'open the file picker, set label to the selected file, & set the button to green to show the success of the process
        If OpenFileDialog.ShowDialog = Windows.Forms.DialogResult.Cancel Then
            fileLabel.Text = "File selection unsucessful"
            fileSelected = False
            fileSelectButton.BackColor = Color.Red

        Else
            fileLabel.Text = ("File Retrived Sucessfully" & vbCrLf & OpenFileDialog.FileName)
            fileSelected = True
            fileSelectButton.BackColor = Color.LightGreen

        End If

    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' xlWorkBook.Close()
        'xlApp.Quit()
        MsgBox(" Run 2 lenghths " & run2Addresses.Count & " & " & run2numbers.Count &
               " Run 3 lenghths " & run3Addresses.Count & " & " & run3numbers.Count &
               " Run 4 lenghths " & run4Addresses.Count & " & " & run4numbers.Count)
        MsgBox(" R2 " & run2Row &
               " R3 " & run3Row &
               " R4 " & run4Row)

        'Dim s As String
        's = ""
        'For i = 1 To run2Addresses.Length - 1
        '    s = s & run2Addresses(i) & " "
        'Next
        'MsgBox(s)
        's = ""
        'For i = 1 To run2numbers.Length - 1
        '    s = s & run2numbers(i) & " "
        'Next
        'MsgBox(s)

    End Sub

    Private Sub ResetVarsForLoops()
        'MsgBox("loop done " & i & " times")
        loopCheker = False
        loopChecker2 = False
        cellRun = ""
        cellAddress = ""
        addressCounter = 1
        counter = -1
        firstLoop = True
    End Sub

End Class
