﻿Imports Excel = Microsoft.Office.Interop.Excel

Public Class Form1
    Dim fileSelected As Boolean
    Dim xlApp As Excel.Application
    Dim xlWorkBook As Excel.Workbook
    Dim xlWorkSheet As Excel.Worksheet
    Dim int As Integer
    Dim loopCheker As Boolean = False
    Dim loopChecker2 As Boolean = False
    Dim run1Row As Integer = 1
    Dim run1Cell As String
    Dim run1Addresses(0) As String
    Dim run1numbers(0) As Integer
    Dim run2Row As Integer = 1
    Dim run2Cell As String
    Dim run2Addresses(0) As String
    Dim run2numbers(0) As Integer
    Dim run3Row As Integer = 1
    Dim run3Cell As String
    Dim run3Addresses(0) As String
    Dim run3numbers(0) As Integer
    Dim run4Row As Integer = 1
    Dim run4Cell As String
    Dim run4Addresses(0) As String
    Dim run4numbers(0) As Integer

    Dim cell As String
    Dim i As Integer = 0
    Dim i2 As Integer = 0
    Dim cell2 As String

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
            Do While loopCheker = False
                run1Cell = ("A" & run1Row)

                If xlWorkSheet.Range(run1Cell).Value = "Run 1" Then
                    loopCheker = True
                ElseIf run1Row >= 1000 Then
                    loopCheker = True
                End If
                run1Row += 1
            Loop

            ResetVarsForLoops()
            Do While loopCheker = False
                run2Cell = ("A" & run2Row)

                If xlWorkSheet.Range(run2Cell).Value = "Run 2" Then
                    loopCheker = True
                ElseIf run2Row >= 1000 Then
                    loopCheker = True
                End If
                run2Row += 1
            Loop

            ResetVarsForLoops()
            Do While loopCheker = False
                run3Cell = ("A" & run3Row)

                If xlWorkSheet.Range(run3Cell).Value = "Run 3" Then
                    loopCheker = True
                ElseIf run3Row >= 1000 Then
                    loopCheker = True
                End If
                run3Row += 1
            Loop

            ResetVarsForLoops()
            Do While loopCheker = False
                run4Cell = ("A" & run4Row)

                If xlWorkSheet.Range(run4Cell).Value = "Run 4" Then
                    loopCheker = True
                ElseIf run4Row >= 1000 Then
                    loopCheker = True
                End If
                run4Row += 1
            Loop

            ResetVarsForLoops()
            Do While loopCheker = False
                cell = ("B" & run1Row + i)

                If run1Row + i >= run2Row Then
                    loopCheker = True
                ElseIf xlWorkSheet.Range(cell).ToString <> "" Then
                    ReDim Preserve run1Addresses(run1Addresses.Length + 1)
                    run1Addresses(i) = xlWorkSheet.Range(cell).ToString

                    Do While loopChecker2 = False
                        cell2 = ("C" & run1Row + i2)
                        If xlWorkSheet.Range(cell2).ToString = "Run 2" Or
                            xlWorkSheet.Range(cell2).ToString = "Run 3" Or
                            xlWorkSheet.Range(cell2).ToString = "Run 4" Or
                            xlWorkSheet.Range("B" & run1Row + i2).ToString <> "" Or
                            i2 + run1Row >= run2Row Then
                            loopChecker2 = True

                        ElseIf xlWorkSheet.Range(cell2).ToString <> "" Then
                            ReDim Preserve run1numbers(run1numbers.Length + 1)
                            run1numbers(i) += 1

                        End If
                        i2 += 1
                    Loop
                End If
                i += 1
            Loop

            ResetVarsForLoops()
            Do While loopCheker = False
                cell = ("B" & run2Row + i)

                If run2Row + i >= run3Row Then
                    loopCheker = True
                ElseIf xlWorkSheet.Range(cell).ToString <> "" Then
                    ReDim Preserve run2Addresses(run2Addresses.Length + 1)
                    run2Addresses(i) = xlWorkSheet.Range(cell).ToString

                    Do While loopChecker2 = False
                        cell2 = ("C" & run2Row + i2)
                        If xlWorkSheet.Range(cell2).ToString = "Run 3" Or
                            xlWorkSheet.Range(cell2).ToString = "Run 4" Or
                            xlWorkSheet.Range("B" & run2Row + i2).ToString <> "" Or
                            i2 + run2Row >= run3Row Then
                            loopChecker2 = True

                        ElseIf xlWorkSheet.Range(cell2).ToString <> "" Then
                            ReDim Preserve run2numbers(run2numbers.Length + 1)
                            run2numbers(i) += 1
                        End If
                        i2 += 1
                    Loop
                    i += 1
                End If
            Loop

            ResetVarsForLoops()
            Do While loopCheker = False
                cell = ("B" & run3Row + i)

                If run3Row + i >= run4Row Then
                    loopCheker = True
                ElseIf xlWorkSheet.Range(cell).ToString <> "" Then
                    ReDim Preserve run3Addresses(run3Addresses.Length + 1)
                    run3Addresses(i) = xlWorkSheet.Range(cell).ToString

                    Do While loopChecker2 = False
                        cell2 = ("C" & run3Row + i2)
                        If xlWorkSheet.Range(cell2).ToString = "Run 4" Or
                            xlWorkSheet.Range("B" & run3Row + i2).ToString <> "" Or
                            i2 + run3Row >= run4Row Then
                            loopChecker2 = True

                        ElseIf xlWorkSheet.Range(cell2).ToString <> "" Then
                            ReDim Preserve run3numbers(run3numbers.Length + 1)
                            run3numbers(i) += 1
                        End If
                        i2 += 1
                    Loop
                End If
                i += 1
            Loop

            ResetVarsForLoops()
            Do While loopCheker = False
                cell = ("B" & run4Row + i)

                If run4Row + i >= 1000 Then
                    loopCheker = True
                ElseIf xlWorkSheet.Range(cell).ToString <> "" Then
                    ReDim Preserve run4Addresses(run4Addresses.Length + 1)
                    run4Addresses(i) = xlWorkSheet.Range(cell).ToString

                    Do While loopChecker2 = False
                        cell2 = ("C" & run4Row + i2)
                        If xlWorkSheet.Range("B" & run4Row + i2).ToString <> "" Or
                            i2 + run4Row >= 1000 Then
                            loopChecker2 = True

                        ElseIf xlWorkSheet.Range(cell2).ToString <> "" Then
                            ReDim Preserve run4numbers(run4numbers.Length + 1)
                            run4numbers(i) += 1
                        End If
                        i2 += 1
                    Loop
                End If
                i += 1
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
        MsgBox("Run 1 lenghths " & run1Addresses.Length & " & " & run1numbers.Length &
               " Run 2 lenghths " & run2Addresses.Length & " & " & run2numbers.Length &
               " Run 3 lenghths " & run3Addresses.Length & " & " & run3numbers.Length &
               " Run 4 lenghths " & run4Addresses.Length & " & " & run4numbers.Length)
    End Sub

    Private Sub ResetVarsForLoops()
        'MsgBox("loop done " & i & " times")
        loopCheker = False
        loopChecker2 = False
        int = 0
        cell = ""
        cell2 = ""
        i = 0
        i2 = 0
    End Sub

End Class
