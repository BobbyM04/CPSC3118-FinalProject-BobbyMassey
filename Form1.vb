'Bobby Massey - CPSC 3118: Graphical User Interface Development - Final Project - 04/25/2024
Public Class Form1
    Public Property inputTextBox As TextBox
    Public Property inchestoMetersRadioButton As RadioButton
    Public Property meterstoInchesRadioButton As RadioButton
    Public Property resultLabel As Label
    Public Property conversionHistoryListBox As ListBox
    Public Property recordsWrittenLabel As Label
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Color.FromArgb(220, 198, 224)

        'Set background image
        Dim picBuilding As New PictureBox()
        picBuilding.SizeMode = PictureBoxSizeMode.AutoSize
        picBuilding.Location = New Point(10, 10)
        picBuilding.Height = 200
        Try
            picBuilding.Image = Image.FromFile("C:\Users\Bobby\building.png")
        Catch ex As Exception
            MessageBox.Show("Error loading payroll image: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Me.Controls.Add(picBuilding)

        'Title Label
        Dim titleLabel As New Label
        titleLabel.Text = "Converter App 2"
        titleLabel.Font = New Font("Arial", 16, FontStyle.Bold)
        titleLabel.AutoSize = True
        titleLabel.Location = New Point((Me.ClientSize.Width - titleLabel.Width) / 2, 20)
        Me.Controls.Add(titleLabel)

        'Instruction Label
        Dim instructionLabel As New Label
        instructionLabel.Text = "Enter a value and choose conversion: "
        instructionLabel.Font = New Font("Arial", 12, FontStyle.Bold)
        instructionLabel.AutoSize = True
        instructionLabel.Location = New Point((Me.ClientSize.Width - instructionLabel.Width) / 2, titleLabel.Bottom + 20)
        Me.Controls.Add(instructionLabel)

        'Input Textbox
        inputTextBox = New TextBox
        inputTextBox.BackColor = Color.Blue
        inputTextBox.ForeColor = Color.White
        inputTextBox.Location = New Point(instructionLabel.Right + 4, instructionLabel.Top)
        inputTextBox.Size = New Size(100, 20)
        Me.Controls.Add(inputTextBox)

        'Blue box for radio buttons
        Dim blueGroupBox As New GroupBox
        blueGroupBox.BackColor = Color.Blue
        blueGroupBox.ForeColor = Color.White
        blueGroupBox.Size = New Size(200, 80)
        blueGroupBox.Location = New Point(instructionLabel.Left, instructionLabel.Bottom + 30)
        Me.Controls.Add(blueGroupBox)

        'Radio button for Inches to Meters
        inchestoMetersRadioButton = New RadioButton
        inchestoMetersRadioButton.Text = "Inches to Meters"
        inchestoMetersRadioButton.AutoSize = True
        inchestoMetersRadioButton.Location = New Point(10, 20)
        blueGroupBox.Controls.Add(inchestoMetersRadioButton)


        ' Radio button Meters to Inches
        meterstoInchesRadioButton = New RadioButton
        meterstoInchesRadioButton.Text = "Meters to Inches"
        meterstoInchesRadioButton.AutoSize = True
        meterstoInchesRadioButton.Location = New Point(10, 45)
        blueGroupBox.Controls.Add(meterstoInchesRadioButton)

        'Display label for result
        resultLabel = New Label
        resultLabel.Text = ""
        resultLabel.Font = New Font("Arial", 12, FontStyle.Bold)
        resultLabel.AutoSize = True
        resultLabel.Location = New Point((Me.ClientSize.Width - resultLabel.Width) / 2, blueGroupBox.Bottom + 50)
        Me.Controls.Add(resultLabel)

        'ListBox for conversion history
        conversionHistoryListBox = New ListBox
        conversionHistoryListBox.BackColor = Color.White
        conversionHistoryListBox.Location = New Point(10, picBuilding.Bottom + 10)
        conversionHistoryListBox.Size = New Size(280, 150)
        Me.Controls.Add(conversionHistoryListBox)

        'Convert Button
        Dim convertButton As New Button
        convertButton.Text = "Convert"
        convertButton.BackColor = Color.White
        convertButton.AutoSize = True
        convertButton.Location = New Point(348, Me.ClientSize.Height - 240)
        AddHandler convertButton.Click, AddressOf ConvertButton_Click
        Me.Controls.Add(convertButton)

        'Save Results Button
        Dim saveResultsButton As New Button
        saveResultsButton.Text = "Save Results"
        saveResultsButton.BackColor = Color.White
        saveResultsButton.AutoSize = True
        saveResultsButton.Location = New Point(348, Me.ClientSize.Height - 160)
        AddHandler saveResultsButton.Click, AddressOf SaveResultsButton_Click
        Me.Controls.Add(saveResultsButton)

        'Clear Results Button
        Dim clearResultsButton As New Button
        clearResultsButton.Text = "Clear Results"
        clearResultsButton.BackColor = Color.White
        clearResultsButton.AutoSize = True
        clearResultsButton.Location = New Point(348, Me.ClientSize.Height - 130)
        AddHandler clearResultsButton.Click, AddressOf ClearResultsButton_Click
        Me.Controls.Add(clearResultsButton)

        'Clear List Button
        Dim clearListButton As New Button
        clearListButton.Text = "Clear List"
        clearListButton.BackColor = Color.White
        clearListButton.AutoSize = True
        clearListButton.Location = New Point(50, Me.ClientSize.Height - 60)
        AddHandler clearListButton.Click, AddressOf ClearListButton_Click
        Me.Controls.Add(clearListButton)

        'Save To File Button
        Dim saveToFileButton As New Button
        saveToFileButton.Text = "Save To File"
        saveToFileButton.BackColor = Color.White
        saveToFileButton.AutoSize = True
        saveToFileButton.Location = New Point(150, Me.ClientSize.Height - 60)
        AddHandler saveToFileButton.Click, AddressOf SaveToFileButton_Click
        Me.Controls.Add(saveToFileButton)

        'Label for displaying records written to file
        recordsWrittenLabel = New Label
        recordsWrittenLabel.Text = ""
        recordsWrittenLabel.Font = New Font("Arial", 12, FontStyle.Bold)
        recordsWrittenLabel.AutoSize = True
        recordsWrittenLabel.Location = New Point(348, clearResultsButton.Bottom + 10)
        Me.Controls.Add(recordsWrittenLabel)
    End Sub
    Private Sub ConvertButton_Click(sender As Object, e As EventArgs)
        Dim inputValue As String = inputTextBox.Text.Trim()
        Dim result As String = ""

        If Not IsNumeric(inputValue) Then
            MessageBox.Show("Please enter a valid number.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim value As Double = CDbl(inputValue)

        If value < 0 Then
            MessageBox.Show("Please enter a positive number.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If inchestoMetersRadioButton.Checked Then
            'Inches to Meters Conversion
            Dim meters As Double = value * 0.0254
            result = $"{inputValue} inches is {meters:F3} meters."
        ElseIf meterstoInchesRadioButton.Checked Then
            'Meters to Inches conversion
            Dim inches As Double = value / 0.0254
            result = $"{inputValue} meters is {inches:F3} inches."
        Else
            MessageBox.Show("Please select a conversion option.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        resultLabel.Text = result
    End Sub
    Private Sub ClearResultsButton_Click(sender As Object, e As EventArgs)
        inputTextBox.Clear()
        resultLabel.Text = ""
        inchestoMetersRadioButton.Checked = False
        meterstoInchesRadioButton.Checked = False
        inputTextBox.Focus()
        recordsWrittenLabel.Text = ""
    End Sub
    Private Sub SaveResultsButton_Click(sender As Object, e As EventArgs)
        If Not String.IsNullOrEmpty(resultLabel.Text) Then
            conversionHistoryListBox.Items.Add(resultLabel.Text)
        End If
    End Sub
    Private Sub ClearListButton_Click(sender As Object, e As EventArgs)
        conversionHistoryListBox.Items.Clear()
        resultLabel.Text = ""
    End Sub
    Private Sub SaveToFileButton_Click(sender As Object, e As EventArgs)
        If conversionHistoryListBox.Items.Count > 0 Then
            Try
                Using writer As New System.IO.StreamWriter("C:\Users\Bobby\measures.txt")
                    For Each item As String In conversionHistoryListBox.Items
                        writer.WriteLine(item)
                    Next
                End Using
                recordsWrittenLabel.Text = $"{conversionHistoryListBox.Items.Count} records written to file."
            Catch ex As Exception
                MessageBox.Show("Error saving to file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else
            MessageBox.Show("No data to save.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub
End Class
