Imports Excel = Microsoft.Office.Interop.Excel

Public Class Form1
    Dim fileSelected As Boolean
    Dim xlApp As Excel.Application
    Dim xlWorkBook As Excel.Workbook
    Dim xlWorkSheet As Excel.Worksheet
    Dim loopCheker As Boolean = False
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

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub RunButton_Click(sender As Object, e As EventArgs) Handles runButton.Click

        'set button colour to light green to show the user has clicked and the program is working
        runButton.BackColor = Color.LightSkyBlue
        If fileSelected = True Then

            'asign the excel vars their values of the excel application
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Open(OpenFileDialog.FileName)
            xlWorkSheet = xlWorkBook.Worksheets("Sheet")

            'reset the vars used for loop tracking then find the row containing RUN 2
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

            'reset the vars used for loop tracking then find the row containing RUN 3
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

            'reset the vars used for loop tracking then find the row containing RUN 4
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

            'reset the vars used for loop tracking then find the row containing the copyright notice used to find the end of the page
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
                'update the cell var to be the next cell down then check if it is a new address or if it is a job under the previous address
                cellAddress = ("B" & run2Row + addressCounter)

                'if the loop has found the row containg the next run it has gone too far and will exit the loop
                If xlWorkSheet.Range(cellAddress).Row >= run3Row Then
                    loopCheker = True

                    'if the checked cell contains text it is a new address and will add it to the list
                ElseIf String.IsNullOrEmpty(xlWorkSheet.Range(cellAddress).Value) = False Then
                    'providing that this isnt the first address found in the loop the finding of a new address will mean the end of the old one and the program will write the amount of jobs found under the previous address to the list
                    If firstLoop = False Then
                        run2numbers.Add(counter)

                    End If
                    firstLoop = False
                    counter = -1
                    run2Addresses.Add(xlWorkSheet.Range(cellAddress).Value)

                ElseIf String.IsNullOrEmpty(xlWorkSheet.Range(cellAddress).Value) Then
                    'increments every time a job is found starting at -1 moving to 0 to accomidate the blank rows under each address
                    If firstLoop = False Then
                        counter += 1

                    End If
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
                    If firstLoop = False Then
                        counter += 1

                    End If
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
                    If firstLoop = False Then
                        counter += 1

                    End If
                End If
                addressCounter += 1

            Loop

            'cleanup by closign the excel application and sending all the use memory to garbage collection 
            xlWorkBook.Close()
            xlApp.Quit()
            ReleaseObject(xlApp)
            ReleaseObject(xlWorkBook)
            ReleaseObject(xlWorkSheet)

            'change button to green to indicate success
            runButton.BackColor = Color.LightGreen

        Else
            'change button to red to show a file wasnt selected 
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

    Private Sub ReleaseObject(ByVal obj As Object)
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
        ' magical testing button land
        'everything here can be removed or commented out as needed


        ' xlWorkBook.Close()
        'xlApp.Quit()
        MsgBox(" Run 2 lenghths " & run2Addresses.Count & " & " & run2numbers.Count &
               " Run 3 lenghths " & run3Addresses.Count & " & " & run3numbers.Count &
               " Run 4 lenghths " & run4Addresses.Count & " & " & run4numbers.Count)
        MsgBox(" R2 " & run2Row &
               " R3 " & run3Row &
               " R4 " & run4Row)
        Dim test As String
        For Each Str As String In run2Addresses
            test = test & " " & vbCrLf & Str
        Next
        MsgBox(test)

    End Sub

    Private Sub ResetVarsForLoops()
        'all the vars that need to be reset for loops

        loopCheker = False  'bool - checks if the loop is finished
        cellAddress = ""    'string - holds the range of the cell being checked
        addressCounter = 0  'int - counts the number of addresses found
        counter = -1        'int - counts the number of jobs found
        firstLoop = True    'bool - checks if this is the first itteration of the loop
    End Sub

End Class
