Imports System.IO
Imports System.Net
Imports System.Text.RegularExpressions
Imports System.Uri
Imports Microsoft.Win32


Public Class Form1

    Dim apikey As String
    Dim apipath As String
    Dim xpath As String
    Dim pauseBtn As Boolean = False
    Dim pauseURLBtn As Boolean = False
    Dim cancel As Boolean = True


    Private headerBox As CheckBox
    Private headerbox2 As CheckBox

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TabPage1.Text = "Parse from Text File"
        TabPage2.Text = "Parse from WebPage"
        TabPage3.Text = "MailboxValidator API Key"
        validateBtn.Enabled = False
        saveBtn.Enabled = False
        Label10.Text = "No file chosen"
        AbortBtn.Visible = False
        AbortBtn.Enabled = False
        Button1.Enabled = False
        Button1.Visible = False


        urlParseBtn.Enabled = False
        urlValidateBtn.Enabled = False
        urlSaveAsBtn.Enabled = False
        apikey = My.Settings.API_Key
        apipath = My.Settings.API_Path
        APIkeyTxtBox.Text = apikey
        APIsaveBtn.Enabled = False
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill
        DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill
        DataGridView1.AllowUserToAddRows = False
        DataGridView2.AllowUserToAddRows = False
        show_checkBox()
        show_checkBox2()
        DataGridView1.RowHeadersVisible = False
        DataGridView2.RowHeadersVisible = False
        TextBox2.BackColor = Color.White
        TextBox2.ForeColor = Color.White
        ParseBtn.Enabled = False
        DataGridView1.DefaultCellStyle.SelectionBackColor = Color.White
        DataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black
        DataGridView2.DefaultCellStyle.SelectionBackColor = Color.White
        DataGridView2.DefaultCellStyle.SelectionForeColor = Color.Black



    End Sub
    ' To add checkbox which use to check all the checkbox
    Private Sub show_checkBox() '
        Dim checkboxHeader As CheckBox = New CheckBox()

        Dim rect As Rectangle = DataGridView1.GetCellDisplayRectangle(0, -1, True)
        rect.X = 30
        rect.Y = 5.3
        With checkboxHeader
            .BackColor = Color.Transparent
        End With
        checkboxHeader.Name = "checkboxHeader"
        checkboxHeader.Size = New Size(14, 14)
        checkboxHeader.Location = rect.Location
        AddHandler checkboxHeader.CheckedChanged, AddressOf checkboxHeader_CheckedChanged
        DataGridView1.Controls.Add(checkboxHeader)

    End Sub
    ' To add checkbox which use to check all the checkbox
    Private Sub show_checkBox2()
        Dim checkboxHeader2 As CheckBox = New CheckBox()
        Dim rect2 As Rectangle = DataGridView2.GetCellDisplayRectangle(0, -1, True)
        rect2.X = 30
        rect2.Y = 5.3
        With checkboxHeader2
            .BackColor = Color.Transparent
        End With
        checkboxHeader2.Name = "checkboxHeader2"
        checkboxHeader2.Size = New Size(14, 14)
        checkboxHeader2.Location = rect2.Location
        AddHandler checkboxHeader2.CheckedChanged, AddressOf checkboxHeader2_CheckedChanged
        DataGridView2.Controls.Add(checkboxHeader2)

    End Sub
    ' check all checkboxes (for datagridview 1)
    Private Sub checkboxHeader_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)

        If DataGridView1.CurrentCell Is Nothing Then
            headerBox = DirectCast(DataGridView1.Controls.Find("checkboxHeader", False)(0), CheckBox)
            headerBox.Checked = False


        Else
            headerBox = DirectCast(DataGridView1.Controls.Find("checkboxHeader", True)(0), CheckBox)
            DataGridView1.CurrentCell = DataGridView1.Rows(0).Cells(1)
            For Each row As DataGridViewRow In DataGridView1.Rows    ''Loop and check all the checkboxes
                row.Cells(0).Value = headerBox.Checked
            Next

        End If

    End Sub
    ' check all checkboxes (for datagridview2)
    Private Sub checkboxHeader2_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)

        If DataGridView2.CurrentCell Is Nothing Then
            headerbox2 = DirectCast(DataGridView2.Controls.Find("checkboxHeader2", True)(0), CheckBox)
            headerbox2.Checked = False


        Else
            headerbox2 = DirectCast(DataGridView2.Controls.Find("checkboxHeader2", True)(0), CheckBox)
            DataGridView2.CurrentCell = DataGridView2.Rows(0).Cells(1)
            For Each row As DataGridViewRow In DataGridView2.Rows  ''Loop and check all the checkboxes
                row.Cells(0).Value = headerbox2.Checked

            Next
        End If

    End Sub
    'function use to check email format by comparing the regex
    Public Function checkEmailFormat(emailAddress) As Boolean
        Dim result As Boolean
        Dim rgx As New Regex("^[a-zA-Z0-9_\-]*(\.[a-zA-Z][a-zA-Z0-9_\-]*)?@[a-z][a-zA-Z-0-9_\-]*\.[a-z]+(\.[a-z]+)?$") ' check if email match the regex
        If rgx.IsMatch(emailAddress) Then
            result = True
            Return result
        Else
            result = False
            Return result
        End If
    End Function
    'validate email button
    Private Sub validateBtn_Click(sender As Object, e As EventArgs) Handles validateBtn.Click

        Dim x As Integer = 0
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells(0).Value = True Then                       ' counter for checked index
                x = x + 1
            Else
            End If
        Next
        If x = 0 Then
            MessageBox.Show("Please select at least 1 email.", "Error")
        ElseIf x > 1 Then
            MessageBox.Show("Please select 1 email only.", "Error")
        ElseIf x = 1 Then
            For Each rows As DataGridViewRow In DataGridView1.Rows
                If rows.Cells(0).Value = True Then
                    Dim email As String = rows.Cells(1).Value.ToString
                    ValidateEmail(email)        'pass the email to validateEmail function
                End If
            Next

        End If

    End Sub
    'Save email button
    Private Sub saveBtn_Click(sender As Object, e As EventArgs) Handles saveBtn.Click
        SaveFileDialog1.Filter = "TEXT files (*.txt)|*.txt|All files (*.*)|*.*|CSV Files (*.csv)|*.csv"
        SaveFileDialog1.FileName = ""
        Dim checkedindex As Integer = 0
        Dim totalrow As Integer = 0
        For Each row As DataGridViewRow In DataGridView1.Rows   'counter for checked index
            If row.Cells(0).Value = True Then
                checkedindex = checkedindex + 1
            End If
        Next
        For Each row As DataGridViewRow In DataGridView1.Rows   'counter to check total row in the datagrid
            totalrow = totalrow + 1
        Next

        If checkedindex >= 1 And checkedindex < totalrow Then ' if user select some of the emails
            Dim result As DialogResult = MessageBox.Show("Save selected emails?", "Save", MessageBoxButtons.YesNo)
            If result = vbYes Then
                If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
                    Dim newArray As New ArrayList()
                    Dim path As String = SaveFileDialog1.FileName
                    If File.Exists(path) = True Then  ' if the file is exists

                        Using sr As StreamReader = New StreamReader(path, True) ' read the file
                            Do Until sr.EndOfStream
                                newArray.Add(sr.ReadLine)   ' put all the emails in the files into arraylist
                            Loop
                        End Using
                        System.IO.File.WriteAllText(path, "")   'clean the file
                        For Each email As DataGridViewRow In DataGridView1.Rows
                            If email.Cells(0).Value = True Then
                                newArray.Add(email.Cells(1).Value)    ' add selected the emails in the datagrid into the same arraylist
                            End If
                        Next
                        newArray.Sort()   ' sort the emails
                        Using sw As StreamWriter = New StreamWriter(path, True)
                            For Each item As String In newArray
                                sw.WriteLine(item)   ' rewrite the file
                            Next
                        End Using
                        MessageBox.Show("Save successfully.", "Save")
                    ElseIf File.Exists(path) = False Then ' if the file not exists
                        Dim newFile As IO.StreamWriter = System.IO.File.CreateText(path) ' create a new file
                        newFile.Close()
                        For Each email As DataGridViewRow In DataGridView1.Rows
                            If email.Cells(0).Value = True Then
                                newArray.Add(email.Cells(1).Value)  ' add all the selected emails into arraylist
                            End If
                        Next
                        Using sw As StreamWriter = New StreamWriter(path, True)
                            For Each item As String In newArray
                                sw.WriteLine(item)    'write all the emails in the arraylist to the textfiles
                            Next
                        End Using
                        MessageBox.Show("Save successfully.", "Save")

                    End If
                Else
                End If
            End If

        ElseIf checkedindex = totalrow Then
            If checkedindex = 0 Then ' no emails 
                MessageBox.Show("No email selected.", "Save")
            Else

                Dim result2 As DialogResult = MessageBox.Show("Save all email?", "Save", MessageBoxButtons.YesNo)
                If result2 = vbYes Then
                    If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
                        Dim newArray As New ArrayList()
                        Dim path As String = SaveFileDialog1.FileName
                        If File.Exists(path) = True Then       ' if file to be saved exists

                            Using sr As StreamReader = New StreamReader(path, True)
                                Do Until sr.EndOfStream   ' read file
                                    newArray.Add(sr.ReadLine) ' add the emails from the file to the text files
                                Loop
                            End Using
                            System.IO.File.WriteAllText(path, "")  ' clean the files
                            For Each email As DataGridViewRow In DataGridView1.Rows
                                If email.Cells(0).Value = True Then
                                    newArray.Add(email.Cells(1).Value)  ' add all the emails to the arraylist
                                End If
                            Next
                            newArray.Sort()
                            Using sw As StreamWriter = New StreamWriter(path, True)
                                For Each item As String In newArray
                                    sw.WriteLine(item)      'add all the emails in the list to the files
                                Next
                            End Using
                            MessageBox.Show("Save successfully.", "Save")
                        ElseIf File.Exists(path) = False Then
                            Dim newFile As IO.StreamWriter = System.IO.File.CreateText(path)
                            newFile.Close()
                            For Each email As DataGridViewRow In DataGridView1.Rows
                                If email.Cells(0).Value = True Then
                                    newArray.Add(email.Cells(1).Value)
                                End If
                            Next
                            Using sw As StreamWriter = New StreamWriter(path, True)
                                For Each item As String In newArray
                                    sw.WriteLine(item)
                                Next
                            End Using
                            MessageBox.Show("Save successfully.", "Save")
                        End If
                    Else
                    End If
                End If
            End If
        ElseIf checkedindex = 0 Then
            MessageBox.Show("Please select at least 1 email to save.", "Error")
        End If

    End Sub
    'Parse all the email from the text files
    Private Sub ParseBtn_Click(sender As Object, e As EventArgs) Handles ParseBtn.Click
        Me.UseWaitCursor = True
        Dim inprogress As Boolean = False
        Me.Refresh()
        pauseBtn = False
        Dim result As Boolean = True
        Dim filePath As String
        Dim fileNewArray As New ArrayList()
        Dim fileEmailList As New ArrayList()
        Dim fileConfirmList As New ArrayList()
        Dim x As Integer = 0
        Dim counter As Integer = 0
        DataGridView1.Rows.Clear()
        filePath = xpath
        ProgressBar1.Maximum = 0
        ProgressBar1.Value = 0
        ProgressBar1.Refresh()
        DataGridView1.Refresh()
        fileEmailList.Clear()
        fileConfirmList.Clear()
        taIndex = 0
        Try
            If xpath = "No file chosen." Then

            ElseIf xpath = "" Then

            Else
                Application.DoEvents()


                Using SR As StreamReader = New StreamReader(filePath, True)
                    Do Until SR.EndOfStream
                        fileNewArray.Add(SR.ReadLine)
                    Loop
                End Using

                For Each item As String In fileNewArray

                    If item.Contains("@") And item.Contains(".") Then
                        If checkEmailFormat(item) = True Then 'check if email exist
                            fileEmailList.Add(item)               'if yes, add email into an arraylist 
                            x = x + 1                        'counter for number of email
                        End If
                    End If
                Next
                If x = 0 Then                                   'if counter = 0 , email not exist
                    MessageBox.Show("No email found.")
                Else
                    fileEmailList.Sort()                               'if yes, sort email list

                    For Each email As String In fileEmailList
                        Dim isNew As Boolean = True                 'set all the email in email list as NEW
                        For Each newEmail As String In fileConfirmList
                            If (newEmail = email) Then isNew = False 'check if the email is existed, if yes email = false
                        Next
                        If (isNew) Then
                            fileConfirmList.Add(email) ' if not existed, add the email into another list
                            counter += 1
                        End If
                    Next
                End If
            End If
            inprogress = True

            If TabControl1.SelectedIndex = 0 Then   ' check if the selected tab is the 1st tab
                If counter > 0 Then
                    ProgressBar1.Maximum = counter
                    AbortBtn.Enabled = True
                    AbortBtn.Visible = True
                    ParseBtn.Visible = False
                    ParseBtn.Enabled = False

                    If ProgressBar1.Value <= ProgressBar1.Maximum Then
                        Do While inprogress = True                'prevent others to switch tab while is in progress
                            For Each email As String In fileConfirmList    'loop through the final list 

                                If TabControl1.SelectedIndex = 0 Then  'check if selectected tab is the 1st tab
                                    Do While pauseBtn = False          ' check if abort button clicked
                                        Application.DoEvents()
                                        If pauseBtn = False Then
                                            DataGridView1.Rows.Add(False, email)        ' add all the email in the list to the datagrid
                                            ProgressBar1.Value += 1
                                            result = True
                                            Exit Do

                                        Else

                                            MessageBox.Show("Progress aborted", "Aborted")
                                            result = False
                                            pauseBtn = True
                                            Exit Do
                                        End If
                                    Loop
                                Else
                                    ProgressBar1.Value += 1
                                    TabControl1.SelectedIndex = 0
                                    TabControl1.TabPages(TabControl1.SelectedIndex).Focus()
                                    TabControl1.TabPages(TabControl1.SelectedIndex).CausesValidation = True
                                    result = True
                                End If
                                If email Is fileConfirmList(fileConfirmList.Count - 1) Then  'check if the email in the list is the last
                                    inprogress = False                                     'if yes, exit the do loop
                                End If
                            Next
                        Loop


                    End If
                End If
            End If
            Me.UseWaitCursor = False
            If result = True Then

                validateBtn.Enabled = True
                saveBtn.Enabled = True
                AbortBtn.Enabled = False
                AbortBtn.Visible = False
                ParseBtn.Visible = True
                ParseBtn.Enabled = True

            Else
                DataGridView1.Refresh()
                saveBtn.Enabled = False
                AbortBtn.Enabled = False
                AbortBtn.Visible = False
                ParseBtn.Visible = True
                ParseBtn.Enabled = True
                validateBtn.Enabled = False
            End If

        Catch ex As Exception
            MessageBox.Show("Unknown Error", "Error")


        End Try





    End Sub
    'browse function which allow user to select files
    Private Sub BrowseBtn_Click(sender As Object, e As EventArgs) Handles BrowseBtn.Click
        OpenFileDialog1.Filter = "TEXT files (*.txt)|*.txt|All files (*.*)|*.*|CSV Files (*.csv)|*.csv"
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            DataGridView1.Rows.Clear()
            Label10.Text = OpenFileDialog1.FileName
            xpath = OpenFileDialog1.FileName
            If (Label10.Text).Length > 40 Then
                Label10.Text = Label10.Text.Substring(0, 40)
                Label10.Update()
            End If
            saveBtn.Enabled = False
            validateBtn.Enabled = False
            ParseBtn.Enabled = True
            ParseBtn.Visible = True
            AbortBtn.Enabled = False
            AbortBtn.Visible = False
        Else
            ParseBtn.Enabled = True
            ParseBtn.Visible = True
            AbortBtn.Enabled = False
            AbortBtn.Visible = False
            saveBtn.Enabled = False
            validateBtn.Enabled = False
        End If

    End Sub
    'parse email from the URL address that entered by the user
    Private Sub urlParseBtn_Click(sender As Object, e As EventArgs) Handles urlParseBtn.Click
        DataGridView2.Rows.Clear()
        ProgressBar2.Value = 0
        Dim emaillist As New List(Of String)
        Me.UseWaitCursor = True
        Application.DoEvents()
        extract(emaillist)
    End Sub
    Private Sub urlValidateBtn_Click(sender As Object, e As EventArgs) Handles urlValidateBtn.Click
        Dim x As Integer = 0
        For Each row As DataGridViewRow In DataGridView2.Rows
            If row.Cells(0).Value = True Then
                x = x + 1
            Else
            End If
        Next
        If x = 0 Then
            MessageBox.Show("Please select at least 1 email.", "Error")
        ElseIf x > 1 Then
            MessageBox.Show("Please select 1 email only.", "Error")
        ElseIf x = 1 Then
            For Each rows As DataGridViewRow In DataGridView2.Rows
                If rows.Cells(0).Value = True Then
                    Dim email As String = rows.Cells(1).Value.ToString
                    Me.UseWaitCursor = True
                    ValidateEmail(email)
                    Me.UseWaitCursor = False
                End If
            Next

        End If

    End Sub
    ' Save function that same as the previous saving method
    Private Sub urlSaveAsBtn_Click(sender As Object, e As EventArgs) Handles urlSaveAsBtn.Click
        SaveFileDialog1.Filter = "TEXT files (*.txt)|*.txt|All files (*.*)|*.*|CSV Files (*.csv)|*.csv"
        SaveFileDialog1.FileName = ""
        Dim checkedindex As Integer = 0
        Dim totalrow As Integer = 0
        For Each row As DataGridViewRow In DataGridView2.Rows
            If row.Cells(0).Value = True Then
                checkedindex = checkedindex + 1
            End If
        Next
        For Each row As DataGridViewRow In DataGridView2.Rows
            totalrow = totalrow + 1
        Next


        If checkedindex >= 1 And checkedindex < totalrow Then
            Dim result As DialogResult = MessageBox.Show("Save checked email?", "Save", MessageBoxButtons.YesNo)
            If result = vbYes Then

                If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
                    Dim newArray As New ArrayList()
                    Dim path As String = SaveFileDialog1.FileName
                    If File.Exists(path) = True Then

                        Using sr As StreamReader = New StreamReader(path, True)
                            Do Until sr.EndOfStream
                                newArray.Add(sr.ReadLine)
                            Loop
                        End Using
                        System.IO.File.WriteAllText(path, "")
                        For Each email As DataGridViewRow In DataGridView2.Rows
                            If email.Cells(0).Value = True Then
                                newArray.Add(email.Cells(1).Value)
                            End If
                        Next
                        newArray.Sort()
                        Using sw As StreamWriter = New StreamWriter(path, True)
                            For Each item As String In newArray
                                sw.WriteLine(item)
                            Next
                        End Using
                        MessageBox.Show("Save successfully.", "Save")
                    ElseIf File.Exists(path) = False Then
                        Dim newFile As IO.StreamWriter = System.IO.File.CreateText(path)
                        newFile.Close()
                        For Each email As DataGridViewRow In DataGridView2.Rows
                            If email.Cells(0).Value = True Then
                                newArray.Add(email.Cells(1).Value)
                            End If
                        Next
                        Using sw As StreamWriter = New StreamWriter(path, True)
                            For Each item As String In newArray
                                sw.WriteLine(item)
                            Next
                        End Using
                        MessageBox.Show("Save successfully.", "Save")

                    End If
                Else
                End If
            End If

        ElseIf checkedindex = totalrow Then
            If checkedindex = 0 Then
                MessageBox.Show("No email selected.", "Save")
            Else
                Dim result2 As DialogResult = MessageBox.Show("Save all email?", "Save", MessageBoxButtons.YesNo)
                If result2 = vbYes Then
                    If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
                        Dim newArray As New ArrayList()
                        Dim path As String = SaveFileDialog1.FileName
                        If File.Exists(path) = True Then

                            Using sr As StreamReader = New StreamReader(path, True)
                                Do Until sr.EndOfStream
                                    newArray.Add(sr.ReadLine)
                                Loop
                            End Using
                            System.IO.File.WriteAllText(path, "")
                            For Each email As DataGridViewRow In DataGridView2.Rows
                                If email.Cells(0).Value = True Then
                                    newArray.Add(email.Cells(1).Value)
                                End If
                            Next
                            newArray.Sort()
                            Using sw As StreamWriter = New StreamWriter(path, True)
                                For Each item As String In newArray
                                    sw.WriteLine(item)
                                Next
                            End Using
                            MessageBox.Show("Save successfully.", "Save")
                        ElseIf File.Exists(path) = False Then
                            Dim newFile As IO.StreamWriter = System.IO.File.CreateText(path)
                            newFile.Close()
                            For Each email As DataGridViewRow In DataGridView2.Rows
                                If email.Cells(0).Value = True Then
                                    newArray.Add(email.Cells(1).Value)
                                End If
                            Next
                            Using sw As StreamWriter = New StreamWriter(path, True)
                                For Each item As String In newArray
                                    sw.WriteLine(item)
                                Next
                            End Using
                            MessageBox.Show("Save successfully.", "Save")
                        End If

                    Else
                    End If

                End If
            End If

        ElseIf checkedindex = 0 Then
            MessageBox.Show("Please select at least 1 email to save.", "Error")
        End If



    End Sub
    'save API key in windows registry
    Private Sub APIsaveBtn_Click(sender As Object, e As EventArgs) Handles APIsaveBtn.Click

        SaveFileDialog1.Filter = "TEXT files (*.txt)|*.txt|All files (*.*)|*.*|CSV Files (*.csv)|*.csv"
        SaveFileDialog1.FileName = ""
        Dim newKey As RegistryKey = Registry.CurrentUser.CreateSubKey("Software\EmailParser")
        newKey.SetValue("MBV_APIKEY", APIkeyTxtBox.Text)

        Dim readvalue = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\EmailParser", "MBV_APIKEY", Nothing)

        Dim newArray As New ArrayList()
        My.Settings.API_Key = readvalue
        apikey = readvalue
        Me.UseWaitCursor = True
        MessageBox.Show("Successfully saved.", "Save")




    End Sub
    'extract email from the webpage HTML code and display in the datagridview
    Private Function extract(ByVal email As List(Of String))
        Dim inprogress As Boolean = False
        Me.UseWaitCursor = True
        urlValidateBtn.Enabled = False
        urlSaveAsBtn.Enabled = False
        Dim EmailList As New ArrayList()
        Dim confirmEmailList As New ArrayList()
        Dim x As Integer = 0
        Dim url As String = TextBox1.Text
        Dim z As Integer = 0
        Dim result As Boolean = False
        taIndex = 1

        Try

            Dim pattern As String = "(http|https):\/\/[\w\-_]+(\.[\w\-_]+)+([\w\-\.,@?^=%&amp;:/~\+#]*[\w\-\@?^=%&amp;/~\+#])?"
            If url = "" Then
                MessageBox.Show("Please enter url.", "Error")
            ElseIf url.StartsWith("www.") Then   ' if the URL typed by the user startswith www.
                url = "http://" + url  'add http:// for the url

                If Regex.IsMatch(url, pattern) Then
                    Dim r As HttpWebRequest = HttpWebRequest.Create(url)
                    r.KeepAlive = True
                    r.UserAgent = "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/29.0.1547.2 Safari/537.36"
                    Dim re As HttpWebResponse = r.GetResponse()
                    Dim src As String = New StreamReader(re.GetResponseStream()).ReadToEnd()
                    Dim words As String() = src.Split(" ")  'every word split by using " "

                    For Each word As String In words
                        If (word.Contains("@") And word.Contains(".")) Then
                            If (word.Contains("<") And word.Contains(">")) Then
                                Dim toAdd As New List(Of String)
                                Dim noTags As String() = GetBetweenAll(word, ">", "<")      ' pass word into GetBetweenAll function to get word between ">" and "<" and save all into noTags
                                For Each w As String In noTags
                                    If (w.Contains("@") And w.Contains(".") And Not w.Contains("=")) Then
                                        If (w.EndsWith(",") Or w.EndsWith(".")) Then
                                            toAdd.Add(w.Substring(0, w.Length - 1))      ' add the email by deleting the last character ";" or "." into list
                                        Else
                                            toAdd.Add(w)                            ' add the email into list
                                        End If
                                    End If
                                Next

                                For Each item In toAdd
                                    If (checkEmailFormat(item)) = True Then        'pass the email from the list to checkemailformat function
                                        If (toAdd.Count > 0) Then      'if valid format email exists
                                            If (toAdd.Count > 1) Then  ' if valid format email more than 1
                                                For Each t As String In toAdd
                                                    If checkEmailFormat(t) = True Then  'loop to check the emails format
                                                        EmailList.Add(t)   'add the valid format email into a list
                                                        x = x + 1
                                                    Else
                                                    End If
                                                Next
                                            Else
                                                EmailList.Add(toAdd(0))   ' if only 1 email with valid form add it into list
                                                x = x + 1
                                            End If
                                        Else
                                            EmailList.Add(word)     ' if no email exist, add to the email list
                                            x = x + 1
                                        End If
                                    End If

                                Next

                            Else

                            End If
                        End If
                    Next
                    If x = 0 Then
                        MessageBox.Show("No email found.", "Result")
                    End If
                Else
                    z = 1
                    MessageBox.Show("                   Please enter a valid URL." & vbCrLf & "Example: https://www.mailboxvalidator.com/", "Error")

                End If




            ElseIf url.StartsWith("https://") = False And url.StartsWith("http://") = False Then
                url = ("https://") + url ' if the url type by the user is facebook.com

                If Regex.IsMatch(url, pattern) Then
                    Dim r As HttpWebRequest = HttpWebRequest.Create(url)
                    r.KeepAlive = True
                    r.UserAgent = "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/29.0.1547.2 Safari/537.36"
                    Dim re As HttpWebResponse = r.GetResponse()
                    Dim src As String = New StreamReader(re.GetResponseStream()).ReadToEnd()
                    Dim words As String() = src.Split(" ")
                    For Each word As String In words
                        If (word.Contains("@") And word.Contains(".")) Then
                            If (word.Contains("<") And word.Contains(">")) Then
                                Dim toAdd As New List(Of String)
                                Dim noTags As String() = GetBetweenAll(word, ">", "<")
                                For Each w As String In noTags
                                    If (w.IndexOf("@") > 1) Then

                                        If (w.Contains("@") And w.Contains(".") And Not w.Contains("=")) Then
                                            If (w.EndsWith(",") Or w.EndsWith(".")) Then
                                                toAdd.Add(w.Substring(0, w.Length - 1))
                                            Else
                                                toAdd.Add(w)
                                            End If
                                        End If

                                    End If
                                Next

                                For Each item In toAdd
                                    If (checkEmailFormat(item)) = True Then
                                        If (toAdd.Count > 0) Then
                                            If (toAdd.Count > 1) Then
                                                For Each t As String In toAdd
                                                    If checkEmailFormat(t) = True Then
                                                        EmailList.Add(t)
                                                        x = x + 1
                                                    Else
                                                    End If
                                                Next
                                            Else
                                                EmailList.Add(toAdd(0))
                                                x = x + 1
                                            End If
                                        Else
                                            EmailList.Add(word)
                                            x = x + 1
                                        End If
                                    End If

                                Next
                            Else
                            End If
                        End If
                    Next
                    If x = 0 Then
                        MessageBox.Show("No email found.", "Result")
                    End If


                Else
                    z = 1
                    MessageBox.Show("                   Please enter a valid URL." & vbCrLf & "Example: https://www.mailboxvalidator.com/", "Error")
                End If
                Application.DoEvents()

            ElseIf url.StartsWith("https://") = True Or url.StartsWith("http://") = True Then
                'correct url format

                If Regex.IsMatch(url, pattern) Then
                    Dim r As HttpWebRequest = HttpWebRequest.Create(url)
                    r.KeepAlive = True
                    r.UserAgent = "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/29.0.1547.2 Safari/537.36"
                    Dim re As HttpWebResponse = r.GetResponse()
                    Dim src As String = New StreamReader(re.GetResponseStream()).ReadToEnd()
                    Dim words As String() = src.Split(" ")
                    For Each word As String In words
                        If (word.Contains("@") And word.Contains(".")) Then
                            If (word.Contains("<") And word.Contains(">")) Then
                                Dim toAdd As New List(Of String)
                                Dim noTags As String() = GetBetweenAll(word, ">", "<")
                                For Each w As String In noTags
                                    If (w.IndexOf("@") > 1) Then

                                        If (w.Contains("@") And w.Contains(".") And Not w.Contains("=")) Then
                                            If (w.EndsWith(",") Or w.EndsWith(".")) Then
                                                toAdd.Add(w.Substring(0, w.Length - 1))
                                            Else
                                                toAdd.Add(w)
                                            End If
                                        End If

                                    End If
                                Next

                                For Each item In toAdd
                                    If (checkEmailFormat(item)) = True Then
                                        If (toAdd.Count > 0) Then
                                            If (toAdd.Count > 1) Then
                                                For Each t As String In toAdd
                                                    If checkEmailFormat(t) = True Then
                                                        EmailList.Add(t)
                                                        x = x + 1
                                                    Else
                                                    End If
                                                Next
                                            Else
                                                EmailList.Add(toAdd(0))
                                                x = x + 1
                                            End If
                                        Else
                                            EmailList.Add(word)
                                            x = x + 1
                                        End If
                                    End If

                                Next



                            Else

                            End If


                        End If

                    Next
                    If x = 0 Then
                        MessageBox.Show("No email found.", "Result")
                    End If

                Else
                    z = 1
                    MessageBox.Show("                   Please enter a valid URL." & vbCrLf & "Example: https://www.mailboxvalidator.com/", "Error")

                End If
            End If

            EmailList.Sort()
            Button1.Visible = True
            Button1.Enabled = True
            urlParseBtn.Visible = False
            urlParseBtn.Enabled = False

            For Each sortedEmail As String In EmailList
                Dim isNew As Boolean = True
                For Each newEmail As String In confirmEmailList
                    If (newEmail = sortedEmail) Then isNew = False
                Next
                If (isNew) Then confirmEmailList.Add(sortedEmail)

            Next

            Dim counter As Integer = 0
            For Each y As String In confirmEmailList
                counter += 1
            Next

            Application.DoEvents()

            inprogress = True
            If TabControl1.SelectedIndex = 1 Then
                If counter > 0 Then
                    ProgressBar2.Maximum = counter


                    If ProgressBar2.Value <= ProgressBar2.Maximum Then
                        Do While inprogress = True                        'set inprogress as true to check below condition
                            For Each confirmedemail As String In confirmEmailList    'loop through the final list 

                                If TabControl1.SelectedIndex = 1 Then
                                    Do While pauseURLBtn = False      ' check is abort button clicked
                                        Application.DoEvents()
                                        If pauseURLBtn = False Then
                                            DataGridView2.Rows.Add(False, confirmedemail)     ' add all the email in the list to the datagrid
                                            result = True
                                            ProgressBar2.Value += 1
                                            Exit Do

                                        Else

                                            MessageBox.Show("Progress aborted", "Aborted")
                                            result = False
                                            pauseURLBtn = True
                                            Exit Do
                                        End If

                                    Loop
                                Else

                                    TabControl1.SelectedIndex = 1
                                    TabControl1.TabPages(TabControl1.SelectedIndex).Focus()
                                    TabControl1.TabPages(TabControl1.SelectedIndex).CausesValidation = True
                                    result = True
                                End If
                                If email Is confirmEmailList(confirmEmailList.Count - 1) Then    ' if the email is the last email in the list, exit for loop
                                    Exit For

                                End If
                            Next
                            inprogress = False                 ' set in progress to exit do loop
                        Loop


                    End If

                End If

            Else

            End If
            Me.UseWaitCursor = False

            If result = True Then
                If z = 1 Or x = 0 Then
                    urlValidateBtn.Enabled = False
                    urlSaveAsBtn.Enabled = False
                    Button1.Visible = False
                    Button1.Enabled = False
                    urlParseBtn.Visible = True
                    urlParseBtn.Enabled = True
                Else
                    Button1.Visible = False
                    Button1.Enabled = False
                    urlParseBtn.Visible = True
                    urlParseBtn.Enabled = True
                    urlValidateBtn.Enabled = True
                    urlSaveAsBtn.Enabled = True
                End If
            Else
                MessageBox.Show("Progress aborted", "Abort")
                Button1.Visible = False
                Button1.Enabled = False
                urlParseBtn.Visible = True
                urlParseBtn.Enabled = True
            End If
            inprogress = False






        Catch ex As Exception
            MessageBox.Show("Invalid URL.", "Error")
            Button1.Visible = False
            Button1.Enabled = False
            urlParseBtn.Visible = True
            urlParseBtn.Enabled = True
        End Try







    End Function
    'Removing tags from a string
    Private Function removeTags(ByVal w As String) 'Function to remove tags in the HTML code
        Dim toReturn As New List(Of String)
        Dim noTags As String() = GetBetweenAll(w, ">", "<")
        For Each word As String In noTags
            If (word.Contains("@") And word.Contains(".") And Not word.Contains("=")) Then
                toReturn.Add(word)
            End If
        Next
        Return toReturn
    End Function
    'function to get the string that between ">" and "<"
    Private Function GetBetweenAll(ByVal Source As String, ByVal Str1 As String, ByVal Str2 As String) As String()
        Dim Results, T As New List(Of String)
        T.AddRange(Regex.Split(Source, Str1))
        T.RemoveAt(0)
        For Each I As String In T
            Results.Add(Regex.Split(I, Str2)(0))
        Next
        Return Results.ToArray
    End Function
    'function that validate email by passing email to the function that provided by MailboxValidator to check whether is a valid email
    Private Sub ValidateEmail(email As String)
        Dim request As HttpWebRequest = Nothing
        Dim response As HttpWebResponse = Nothing

        Dim xapiKey As String = apikey
        Dim data As New Dictionary(Of String, String)

        data.Add("format", "html")
        data.Add("email", email)
        Dim datastr As String = String.Join("&", data.[Select](Function(x) x.Key & "=" & EscapeDataString(x.Value)).ToArray())

        request = Net.WebRequest.Create("https://api.mailboxvalidator.com/v1/validation/single?key=" & apikey & "&" & datastr)

        request.Method = "GET"
        response = request.GetResponse()



        Dim reader As String = New IO.StreamReader(response.GetResponseStream()).ReadToEnd()
        Dim result As String() = reader.Split(",")
        For Each item As String In result
            If item.Contains("status"":""False") Then
                MessageBox.Show("Email address is invalid.", "Result")
            ElseIf item.Contains("status"":""True") Then
                MessageBox.Show("Email address is valid.", "Result")
            ElseIf item.Contains("error_code"":""100") Then
                MessageBox.Show("Missing Parameter.", "Error_Code: 100")
            ElseIf item.Contains("error_code"":""101") Then
                MessageBox.Show("API Key not found.", "Error_Code: 101")
            ElseIf item.Contains("error_code"":""102") Then
                MessageBox.Show("API Key disabled.", "Error_Code: 102")
            ElseIf item.Contains("error_code"":""103") Then
                MessageBox.Show("API key expired.", "Error_Code: 103")
            ElseIf item.Contains("error_code"":""104") Then
                MessageBox.Show("Insufficient credits.", "Error_Code: 104")
            ElseIf item.Contains("error_code"":""105") Then
                MessageBox.Show("Unknown error.", "Error_Code: 105")


            End If

        Next




    End Sub
    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click
        System.Diagnostics.Process.Start("https://www.mailboxvalidator.com/plans#api")
    End Sub
    Private Sub APIkeyTxtBox_TextChanged(sender As Object, e As EventArgs) Handles APIkeyTxtBox.TextChanged
        APIsaveBtn.Enabled = True

    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        urlParseBtn.Enabled = True
    End Sub

    Private Sub AbortBtn_Click(sender As Object, e As EventArgs) Handles AbortBtn.Click
        pauseBtn = True

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        pauseURLbtn = True
    End Sub

End Class



