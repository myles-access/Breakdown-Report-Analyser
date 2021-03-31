Imports Microsoft.Office.Interop

Public Class Form1
    Dim fileSelected As Boolean

    Private Sub RunButton_Click(sender As Object, e As EventArgs) Handles runButton.Click
        Dim run1Row As Integer
        Dim run2Row As Integer
        Dim run3Row As Integer
        Dim run4Row As Integer

        'set button colour to light green to show the user has clicked and the program is working
        runButton.BackColor = Color.LightSkyBlue
        If fileSelected = True Then
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

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
