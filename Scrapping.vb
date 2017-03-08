Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Net.Mail
Imports System.Threading
Imports System.Drawing
Imports System.IO.FileStream
Imports System.Web




'***********************************************************************************
'******
'******  This class class file was created by Juan Jose Garcia 
'******
'******  The pourpose of this file is to speed up the process of creating software
'******
'******  If you want more information about this project please contact me
'******
'******   juangarciamtl@gmail.com   | Get my other Udemy courses at http://josegarcia.ca/udemy  
'******
'******   Like my Fanpage to get the latest information about my projects and updates about my courses
'******   https://www.facebook.com/pages/Juan-Jose-Garcia/869675799741084
'******   
'******  Learn more about myself at my website where i post video tutorials, and useful information
'******  My website:  http://josegarcia.ca
'******
'******  Version of this Class V 2.0.3     Date: February 19, 2017
'******  
'******  More than 100 Subroutines for your use
'*************************************************************************************




Public Class Scrapping


    ''' <summary>
    ''' This subroutine will allow you to save a texbox  into a txt file.
    ''' </summary>
    ''' <param name="DIRECTORYTOSAVE">This is the directory where the file will be saved</param>
    ''' <param name="FILENAME">This is the filename that you want to give to the file</param>
    ''' <param name="EXTENCION">This is the extencion that you want to use most of the time will be "txt"</param>
    ''' <param name="TEXTBOXTOSAVE">Here you select wich is the textbox that contains the text that you would like to save</param>
    ''' <remarks></remarks>
    Public Shared Sub SAVETXTFROMTEXTBOX(DIRECTORYTOSAVE As String, FILENAME As String, EXTENCION As String, TEXTBOXTOSAVE As TextBox)

        Try
            If Not Directory.Exists(DIRECTORYTOSAVE) Then
                Directory.CreateDirectory(DIRECTORYTOSAVE)
            End If
            'this is the path where the text file will be saved
            My.Computer.FileSystem.WriteAllText(DIRECTORYTOSAVE & "\" & FILENAME & "." & EXTENCION, TEXTBOXTOSAVE.Text, False, Encoding.UTF8)

        Catch ex As Exception

            MsgBox("There was an error the Text Document has not been saved" & Environment.NewLine & ex.Message)

        End Try
    End Sub

    Public Shared Sub SAVETXTFROMLABEL(DIRECTORYTOSAVE As String, FILENAME As String, EXTENCION As String, TEXTBOXTOSAVE As Label)

        Try
            If Not Directory.Exists(DIRECTORYTOSAVE) Then
                Directory.CreateDirectory(DIRECTORYTOSAVE)
            End If
            'this is the path where the text file will be saved
            My.Computer.FileSystem.WriteAllText(DIRECTORYTOSAVE & "\" & FILENAME & "." & EXTENCION, TEXTBOXTOSAVE.Text, False, Encoding.UTF8)

        Catch ex As Exception

            MsgBox("There was an error the Text Document has not been saved" & Environment.NewLine & ex.Message)

        End Try
    End Sub




    Public Shared Sub SAVETXTFROMstring(DIRECTORYTOSAVE As String, FILENAME As String, EXTENCION As String, TEXTBOXTOSAVE As String)

        Try
            If Not Directory.Exists(DIRECTORYTOSAVE) Then
                Directory.CreateDirectory(DIRECTORYTOSAVE)
            End If
            'this is the path where the text file will be saved
            My.Computer.FileSystem.WriteAllText(DIRECTORYTOSAVE & "\" & FILENAME & "." & EXTENCION, TEXTBOXTOSAVE, False, Encoding.ASCII)

            'My.Computer.FileSystem.WriteAllText(My.Settings.DefaultOutput, fileText, True, System.Text.Encoding.ASCII)

        Catch ex As Exception

            MsgBox("There was an error the Text Document has not been saved")

        End Try
    End Sub



    ''' <summary>
    ''' This subroutine will import all lines from a text file and will place them in a listbox
    ''' Every line in the textbox will become a line (item) in the listbox 
    ''' </summary>
    ''' <param name="FolderFileNameExt">
    ''' This should include the foldername the file name and the extention
    ''' Format similar to c:\cardscan\front.jpg
    ''' </param>
    ''' <param name="Listbox">This is the listbox that will show the result from the .txt file</param>
    ''' <remarks></remarks>
    Public Shared Sub TxtFileToListbox(FolderFileNameExt As String, Listbox As ListBox)

        Try
            Dim file_name As String = FolderFileNameExt
            Dim stream_reader As New IO.StreamReader(file_name, Encoding.GetEncoding("iso-8859-1"))
            Dim line As String

            ' Read the file one line at a time.
            line = stream_reader.ReadLine()
            Do While Not (line Is Nothing)
                ' Trim and make sure the line isn't blank.
                line = line.Trim()
                If line.Length > 0 Then _
                    Listbox.Items.Add(line)

                ' Get the next line.
                line = stream_reader.ReadLine()
            Loop
            Listbox.SelectedIndex = 0
            stream_reader.Close()
        Catch exc As Exception
            ' Report all errors.
            MsgBox(exc.Message, MsgBoxStyle.Exclamation, "Read " & "Error")
        End Try

    End Sub



    ''' <summary>
    ''' Do you want to color your datagridviews when you the value of a column is equal to a string. 
    ''' Then use this subroutine
    ''' </summary>
    ''' <param name="DTGRV">This is the Datagridview that you want to use</param>
    ''' <param name="Column">This is the title of the column where you will compare the value equal to</param>
    ''' <param name="Cell">This is the value that you compare, if you your value is equal to this string
    '''  then the row will change to the color you specified</param>
    ''' <param name="Color">This is the color that you want the row to change when the value is matched</param>
    ''' <remarks></remarks>

    Public Shared Sub COLORDATAGRID(DTGRV As DataGridView, Column As String, Cell As String, Color As Color)
        Try

            For i = 0 To DTGRV.RowCount - 2
                If DTGRV.Rows(i).Cells(Column).Value = Cell Then
                    DTGRV.Rows(i).DefaultCellStyle.BackColor = Color
                End If
            Next
        Catch ex As Exception
        End Try
    End Sub



    ''' <summary>
    ''' Do you want to color your datagridviews when you the value of a column is equal to a string. 
    ''' Then use this subroutine
    ''' </summary>
    ''' <param name="DTGRV">This is the Datagridview that you want to use</param>
    ''' <param name="Column">This is the title of the column where you will compare the value equal to</param>
    ''' <param name="Cell">This is the value that you compare, if you your value is equal to this string
    '''  then the row will change to the color you specified</param>
    ''' <param name="Color">This is the color that you want the row to change when the value is matched</param>
    ''' <remarks></remarks>

    Public Shared Sub COLORDATAGRIDIFCONTAIN(DTGRV As DataGridView, Column As String, Cell As String, Color As Color)
        Try

            For i = 0 To DTGRV.RowCount - 2

                Dim VALUEDATAGRID As String = DTGRV.Rows(i).Cells(Column).Value

                If VALUEDATAGRID.Contains(Cell) = True Then
                    DTGRV.Rows(i).DefaultCellStyle.BackColor = Color
                End If
            Next
        Catch ex As Exception
        End Try
    End Sub




    ''' <summary>
    ''' Do you want to color your datagridviews when you the value of a column is equal to a string. 
    ''' Then use this subroutine
    ''' </summary>
    ''' <param name="DTGRV">This is the Datagridview that you want to use</param>
    ''' <param name="Column">This is the title of the column where you will compare the value equal to</param>
    ''' <param name="Cell">This is the value that you compare, if you your value is equal to this string
    '''  then the row will change to the color you specified</param>
    ''' <param name="Color">This is the color that you want the row to change when the value is matched</param>
    ''' <remarks></remarks>

    Public Shared Sub COLORDATAGRIDIFNOTCONTAIN(DTGRV As DataGridView, Column As String, Cell As String, Color As Color)
        Try

            For i = 0 To DTGRV.RowCount - 2

                Dim VALUEDATAGRID As String = DTGRV.Rows(i).Cells(Column).Value


                If VALUEDATAGRID.Contains(Cell) = False Then
                    DTGRV.Rows(i).DefaultCellStyle.BackColor = Color
                End If
            Next
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' This subroutine will clear the value of the combobox and it will remove all the items attached to that combobox.
    ''' </summary>
    ''' <param name="ComboboxToReset">This is the name of the combobox that you want to be reseted</param>
    ''' <remarks></remarks>
    Public Shared Sub ResetCombobox(ComboboxToReset As ComboBox)

        ComboboxToReset.Text = ""
        ComboboxToReset.Items.Clear()

    End Sub


    ''' <summary>
    ''' This will convert the html entities to html codes
    ''' </summary>
    ''' <param name="TEXTBOX"></param>

    Public Shared Sub HtmlentitiestoHTML(TEXTBOX As TextBox)

        TEXTBOX.Text = TEXTBOX.Text.Replace("amp;", "&")

        TEXTBOX.Text = TEXTBOX.Text.Replace("%3F", "?")

        TEXTBOX.Text = TEXTBOX.Text.Replace("%2B", "+")
        TEXTBOX.Text = TEXTBOX.Text.Replace("%3D", "=")
        TEXTBOX.Text = TEXTBOX.Text.Replace("%26", "&")

        TEXTBOX.Text = TEXTBOX.Text.Replace("#39;", "'")
        TEXTBOX.Text = TEXTBOX.Text.Replace("%252B", "+")
        TEXTBOX.Text = TEXTBOX.Text.Replace("quot;", """")
        TEXTBOX.Text = TEXTBOX.Text.Replace("nbsp;", " ")

        'nbsp;

    End Sub


    ''' <summary>
    ''' Sometimes you have dupplicated values in a combobox. To remove the Dupplicated values use this Subroutine
    ''' </summary>
    ''' <param name="COMBOBOXWITHDUPLICATEDVALUE">The name of the Combobox with the Dupplicated Values.</param>
    ''' <remarks></remarks>

    Public Shared Sub RemoveDuplicatedValuesFromCombobox(COMBOBOXWITHDUPLICATEDVALUE As ComboBox)

        For i As Int16 = 0 To COMBOBOXWITHDUPLICATEDVALUE.Items.Count - 2
            For j As Int16 = COMBOBOXWITHDUPLICATEDVALUE.Items.Count - 1 To i + 1 Step -1
                If COMBOBOXWITHDUPLICATEDVALUE.Items(i).ToString = COMBOBOXWITHDUPLICATEDVALUE.Items(j).ToString Then
                    COMBOBOXWITHDUPLICATEDVALUE.Items.RemoveAt(j)
                End If
            Next
        Next

    End Sub


    ''' <summary>
    '''  Transfer the value to a label from the selected row in the datagridview
    '''  This will optimize the time since you will get the value of the column on the selected datagridview
    ''' </summary>
    ''' <param name="DATAGRIDVIEW">This is the datagridview that contains the row where you want to get the value</param>
    ''' <param name="LABEL">This is the label where the result will be shown</param>
    ''' <param name="COLUMNNUMBER">
    ''' This is the the column number of the datagridview where the value that you will get is located
    ''' Remember that the datagridview start with 0 so you will have to substract 1 from the column number 
    ''' </param>
    ''' <remarks></remarks>

    Public Shared Sub SelectedDGVToLabel(DATAGRIDVIEW As DataGridView, LABEL As Label, COLUMNNUMBER As Integer)

        For Each RW As DataGridViewRow In DATAGRIDVIEW.SelectedRows
            LABEL.Text = RW.Cells(COLUMNNUMBER).Value.ToString
        Next

    End Sub

    ''' <summary>
    '''  Transfer the value to a textbox from the selected row in the datagridview
    '''  This will optimize the time since you will get the of the column on the selected datagridview
    ''' </summary>
    ''' <param name="DATAGRIDVIEW">This is the datagridview that contains the row where you want to get the value</param>
    ''' <param name="TEXTBOX">This is the textbox where the result will be shown</param>
    ''' <param name="COLUMNNUMBER">
    ''' This is the the column number of the datagridview where the value that you will get is located
    ''' Remember that the datagridview start with 0 so you will have to substract 1 from the column number 
    ''' </param>
    ''' <remarks></remarks>
    Public Shared Sub SELECTEDDATAGRIDVIEWTOTEXTBOX(DATAGRIDVIEW As DataGridView, TEXTBOX As TextBox, COLUMNNUMBER As Integer)

        For Each RW As DataGridViewRow In DATAGRIDVIEW.SelectedRows
            TEXTBOX.Text = RW.Cells(COLUMNNUMBER).Value.ToString
        Next

    End Sub

    ''' <summary>
    '''  Transfer the value to a textbox from the selected row in the datagridview
    '''  This will optimize the time since you will get the of the column on the selected datagridview
    ''' </summary>
    ''' <param name="DATAGRIDVIEW">This is the datagridview that contains the row where you want to get the value</param>
    ''' <param name="TEXTBOX">This is the label where the result will be shown</param>
    ''' <param name="COLUMNNUMBER"></param>
    Public Shared Sub SELECTEDDATAGRIDVIEWTOLABEL(DATAGRIDVIEW As DataGridView, Label As Label, COLUMNNUMBER As Integer)

        For Each RW As DataGridViewRow In DATAGRIDVIEW.SelectedRows
            Label.Text = RW.Cells(COLUMNNUMBER).Value.ToString
        Next

    End Sub

    ''' <summary>
    '''  Transfer the value to a Combobox from the selected row in the datagridview
    '''  This will optimize the time since you will get the of the column on the selected datagridview
    ''' </summary>
    ''' <param name="DATAGRIDVIEW">This is the datagridview that contains the row where you want to get the value</param>
    ''' <param name="COMBOBOX">This is the combobox where the result will be shown</param>
    ''' <param name="COLUMNNUMBER">
    ''' This is the the column number of the datagridview where the value that you will get is located
    ''' Remember that the datagridview start with 0 so you will have to substract 1 from the column number 
    ''' </param>
    ''' <remarks></remarks>

    Public Shared Sub SELECTEDDATAGRIDVIEWTOCOMBOBOX(DATAGRIDVIEW As DataGridView, COMBOBOX As ComboBox, COLUMNNUMBER As Integer)

        For Each RW As DataGridViewRow In DATAGRIDVIEW.SelectedRows
            COMBOBOX.Text = RW.Cells(COLUMNNUMBER).Value.ToString
        Next

    End Sub

    ''' <summary>
    ''' This will return a string from the selected datagridview in the column of your choice
    ''' </summary>
    ''' <param name="DATAGRIDVIEW">This is the datagridview that contains the row where you want to get the value</param>
    ''' <param name="COLUMNNUMBER">
    ''' This is the the column number of the datagridview where the value that you will get is located
    ''' Remember that the datagridview start with 0 so you will have to substract 1 from the column number
    ''' </param>
    ''' <returns>This will return the value of the selected row in the specified column</returns>
    ''' <remarks></remarks>

    Public Shared Function FromDGVRetunSelectedAsString(DATAGRIDVIEW As DataGridView, COLUMNNUMBER As Integer) As String
        Try

            For Each RW As DataGridViewRow In DATAGRIDVIEW.SelectedRows
            Dim SelectedItem = RW.Cells(COLUMNNUMBER).Value.ToString

            Return SelectedItem
        Next


        Catch ex As Exception

            End Try
    End Function

    ''' <summary>
    ''' This will remove the row that is selected in the datagridview
    ''' </summary>
    ''' <param name="DATAGRIDVIEW1">This is the datagridview that contains the row that you want to delete</param>
    ''' <remarks></remarks>
    Public Shared Sub REMOVESELECTED(DATAGRIDVIEW1 As DataGridView)
        Try
            If DATAGRIDVIEW1.SelectedRows.Count > 0 AndAlso _
        Not DATAGRIDVIEW1.SelectedRows(0).Index = DATAGRIDVIEW1.Rows.Count - 1 Then
            DATAGRIDVIEW1.Rows.RemoveAt(DATAGRIDVIEW1.SelectedRows(0).Index)
        End If

        Catch ex As Exception

            End Try
    End Sub

    ''' <summary>
    ''' This will kill a process, so for example if you want to read a file that is already open, you have to close it before
    ''' so you can use this subroutine to kill that process
    ''' </summary>
    ''' <param name="Application">
    ''' This is the name of the application that will be closed.
    ''' for notepad you would type "notepad" without quotes
    ''' </param>
    ''' <remarks></remarks>
    Public Shared Sub KillProcess(Application As String)
        Try
            Dim PROCESS() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcessesByName(Application)
            For Each process1 As System.Diagnostics.Process In PROCESS
                process1.Kill()
            Next
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' This will delete a file from the desktop, you only need to specify the name and the extension
    ''' </summary>
    ''' <param name="FileName">This is the file name with extention for example  "filename.txt"</param>
    ''' <remarks></remarks>
    Public Shared Sub DeleteFileFromDesktop(FileName As String)
        Try
            Dim folder As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            FileName = folder + "\" + FileName
            My.Computer.FileSystem.DeleteFile(FileName)
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' This will delete a file from the computer you have to specify the full path, filename and extention
    ''' </summary>
    ''' <param name="PathFileNameExt">full path, filename and extention of the file that you want to delete</param>
    ''' <remarks></remarks>
    Public Shared Sub DeleteFileFromComputer(PathFileNameExt As String)
        Try
            My.Computer.FileSystem.DeleteFile(PathFileNameExt)
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' This will remove the first character in a string, It requires the textbox name and the number of character
    ''' that you want to remove. This are the first characters from the string
    ''' </summary>
    ''' <param name="textboxname">This is the textbox name. The source and result will be on the same textbox</param>
    ''' <param name="QytCharactersRemove">This is an integer telling you how many characteres will be removed</param>
    ''' <remarks></remarks>
    Public Shared Sub RemoveFirstCharacters(textboxname As TextBox, QytCharactersRemove As Integer)
        Try
            Dim str9 As String = textboxname.Text
        str9 = str9.Remove(0, QytCharactersRemove)
            textboxname.Text = str9

        Catch ex As Exception

            End Try
    End Sub

    Public Shared Sub RemoveFirstCharactersLB(textboxname As Label, QytCharactersRemove As Integer)
        Try
            Dim str9 As String = textboxname.Text
            str9 = str9.Remove(0, QytCharactersRemove)
            textboxname.Text = str9

        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' This will remove the last characters from a string, It requires the textbox name and the number of character
    ''' that you want to remove. This are the last characters from the string
    ''' </summary>
    ''' <param name="textboxname">This is the textbox name. The source and result will be on the same textbox</param>
    ''' <param name="QytCharactersRemove">This is an integer telling you how many characteres will be removed</param>
    ''' <remarks></remarks>
    ''' 
    Public Shared Sub RemoveLastCharacters(textboxname As TextBox, QytCharactersRemove As Integer)
        Dim s15 As String
        s15 = textboxname.Text
        textboxname.Text = textboxname.Text.Substring(0, s15.Length - QytCharactersRemove)
    End Sub


    Public Shared Sub RemoveLastCharactersLB(textboxname As Label, QytCharactersRemove As Integer)
        Dim s15 As String
        s15 = textboxname.Text
        textboxname.Text = textboxname.Text.Substring(0, s15.Length - QytCharactersRemove)
    End Sub
    ''' <summary>
    ''' This will run a file tha is located on the same folder where the application is running
    ''' </summary>
    ''' <param name="filename">
    ''' This is the filename it should contains the extention.
    ''' String Format:  Filename.exe  or Filename.txt or Filename.extension
    ''' </param>
    ''' <remarks></remarks>

    Public Shared Sub RunFileSameDirectoryThisApp(filename As String)
        Dim SourcePath As String = System.AppDomain.CurrentDomain.BaseDirectory & filename
        If System.IO.File.Exists(SourcePath) Then
            Process.Start(System.AppDomain.CurrentDomain.BaseDirectory & filename)
        Else
            MsgBox("The file " & filename & " doesn't exist")
        End If
    End Sub


    ''' <summary>
    ''' This will download the html source from a website and will place it to a Textbox
    ''' </summary>
    ''' <param name="url">This is the Url that will be download into a textbox</param>
    ''' <param name="textbox">This is the name of the textbox where you will save the html source code</param>
    ''' <remarks></remarks>

    Shared Sub DownloadHtmlPage(ByVal url As String, textbox As TextBox)
        Try
            Dim result As String
            Dim objResponse As WebResponse
            Dim objRequest As WebRequest = System.Net.HttpWebRequest.Create(url)

            DirectCast(objRequest, System.Net.HttpWebRequest).UserAgent = "Mozilla/5.0 (Windows NT 6.3; WOW64; rv:34.0) Gecko/20100101 Firefox/34.0"
            objResponse = objRequest.GetResponse()
            Using sr As New StreamReader(objResponse.GetResponseStream())
                result = sr.ReadToEnd()
                'Close and clean up the StreamReader
                sr.Close()
            End Using
            result = result.ToString

            textbox.Text = result
            textbox.Text = textbox.Text.Replace("""", "#")

        Catch ex As Exception
            'lblStatus.Text = ex.Message
            textbox.Text = ex.Message.ToString

        End Try
    End Sub




    ''' <summary>
    ''' This will download the html source from a website and will place it to a Textbox
    ''' </summary>
    ''' <param name="url">This is the Url that will be download into a textbox</param>
    ''' <param name="textbox">This is the name of the textbox where you will save the html source code</param>
    ''' <remarks></remarks>

    Shared Sub DownloadHtmlPageWithReffer(ByVal url As String, textbox As TextBox)
        Try

            Dim result As String
            ' Create a 'HttpWebRequest' object.
            Dim myHttpWebRequest As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
            ' Referer property is set to http://www.microsoft.com
            myHttpWebRequest.Referer = "http://video.bajaryoutube.com/"




            ' The response object of 'HttpWebRequest' is assigned to a 'HttpWebResponse' variable.
            Dim myHttpWebResponse As HttpWebResponse = CType(myHttpWebRequest.GetResponse(), HttpWebResponse)
            ' Displaying the contents of the page to the console
            '  Dim streamResponse As Stream = myHttpWebResponse.GetResponseStream()


            Using sr As New StreamReader(myHttpWebResponse.GetResponseStream())
                result = sr.ReadToEnd()
                'Close and clean up the StreamReader
                sr.Close()
            End Using
            result = result.ToString

            textbox.Text = result
            textbox.Text = textbox.Text.Replace("""", "#")

            '  Console.WriteLine("Referer to the site is:{0}", myHttpWebRequest.Referer)




        Catch ex As Exception
            'lblStatus.Text = ex.Message
            textbox.Text = ex.Message.ToString

        End Try
    End Sub


    Shared Sub DownloadHtmlPageWithCustomReffer(ByVal url As String, textbox As TextBox, refer As String)
        Try

            Dim result As String
            ' Create a 'HttpWebRequest' object.
            Dim myHttpWebRequest As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
            ' Referer property is set to http://www.microsoft.com
            myHttpWebRequest.Referer = refer
            ' The response object of 'HttpWebRequest' is assigned to a 'HttpWebResponse' variable.
            Dim myHttpWebResponse As HttpWebResponse = CType(myHttpWebRequest.GetResponse(), HttpWebResponse)
            ' Displaying the contents of the page to the console
            '  Dim streamResponse As Stream = myHttpWebResponse.GetResponseStream()


            Using sr As New StreamReader(myHttpWebResponse.GetResponseStream())
                result = sr.ReadToEnd()
                'Close and clean up the StreamReader
                sr.Close()
            End Using
            result = result.ToString

            textbox.Text = result
            textbox.Text = textbox.Text.Replace("""", "#")

            '  Console.WriteLine("Referer to the site is:{0}", myHttpWebRequest.Referer)




        Catch ex As Exception
            'lblStatus.Text = ex.Message
            textbox.Text = ex.Message.ToString

        End Try
    End Sub

    ''' <summary>
    ''' This will clear the html code so there are no break lines and you can modify the html code
    ''' This will also change the quotes for the symbol #
    ''' </summary>
    ''' <param name="HtmlCode">This is a richtextbox that contains the Html source code</param>
    ''' <remarks></remarks>

    Shared Sub ClearHtmlRichTB(HtmlCode As RichTextBox)

        ' THIS WILL CLEAR THE EXTRA INFORMATION

        ' Remove new lines since they are not visible in HTML
        HtmlCode.Text = HtmlCode.Text.Replace("\n", " ")

        ' Remove tab spaces
        HtmlCode.Text = HtmlCode.Text.Replace("\t", " ")

        ' Remove multiple white spaces from HTML
        HtmlCode.Text = Regex.Replace(HtmlCode.Text, "\\s+", " ")

        ' Remove HEAD tag
        ' HtmlCode.Text = Regex.Replace(HtmlCode.Text, "<head.*?</head>", "" _
        ' , RegexOptions.IgnoreCase Or RegexOptions.Singleline)

        ' Remove any JavaScript
        HtmlCode.Text = Regex.Replace(HtmlCode.Text, "<script.*?</script>", "" _
        , RegexOptions.IgnoreCase Or RegexOptions.Singleline)


        HtmlCode.Text = HtmlCode.Text.Replace(vbCr, "").Replace(vbLf, "")
        HtmlCode.Text = HtmlCode.Text.Replace("\n", "").Replace("\r", "")
        HtmlCode.Text = HtmlCode.Text.Replace(Chr(10), "").Replace(Chr(13), "")

        HtmlCode.Text = HtmlCode.Text.Replace("""", "#")
        HtmlCode.Text = HtmlCode.Text.Replace("&", "")
        HtmlCode.Text = HtmlCode.Text.Replace("'", "")
        HtmlCode.Text = HtmlCode.Text.Replace(Environment.NewLine, "")


    End Sub


    ''' <summary>
    ''' This will clear the html code so there are no break lines and you can modify the html code
    ''' This will also change the quotes for the symbol #
    ''' </summary>
    ''' <param name="HtmlCode">This is a Textbox that contains the Html source code</param>
    ''' <remarks></remarks>


    Shared Sub ClearHtmlTB(HtmlCode As TextBox)

        ' THIS WILL CLEAR THE EXTRA INFORMATION

        ' Remove new lines since they are not visible in HTML
        HtmlCode.Text = HtmlCode.Text.Replace("\n", " ")

        ' Remove tab spaces
        HtmlCode.Text = HtmlCode.Text.Replace("\t", " ")

        ' Remove multiple white spaces from HTML
        HtmlCode.Text = Regex.Replace(HtmlCode.Text, "\\s+", " ")


        ' Remove HEAD tag
        ' HtmlCode.Text = Regex.Replace(HtmlCode.Text, "<head.*?</head>", "" _
        ' , RegexOptions.IgnoreCase Or RegexOptions.Singleline)

        ' Remove any JavaScript
        HtmlCode.Text = Regex.Replace(HtmlCode.Text, "<script.*?</script>", "" _
        , RegexOptions.IgnoreCase Or RegexOptions.Singleline)

        HtmlCode.Text = HtmlCode.Text.Replace(vbCr, "").Replace(vbLf, "")
        HtmlCode.Text = HtmlCode.Text.Replace("\n", "").Replace("\r", "")
        HtmlCode.Text = HtmlCode.Text.Replace(Chr(10), "").Replace(Chr(13), "")

        HtmlCode.Text = HtmlCode.Text.Replace("""", "#")
        HtmlCode.Text = HtmlCode.Text.Replace("&", "")
        HtmlCode.Text = HtmlCode.Text.Replace("'", "")
        HtmlCode.Text = HtmlCode.Text.Replace(Environment.NewLine, "")


    End Sub


    Public Shared Sub RemoveTextBetweenTags(htmlcode As TextBox, starttag As String, endtag As String)

        htmlcode.Text = Regex.Replace(htmlcode.Text, starttag & ".*?" & endtag, "" _
      , RegexOptions.IgnoreCase Or RegexOptions.Singleline)

    End Sub


    Public Shared Sub CopyItemsComboboxToCombobox(OriginalCombobox As ComboBox, NewCombobox As ComboBox)

        For i = 0 To OriginalCombobox.Items.Count - 1
            NewCombobox.Items.Add(OriginalCombobox.Items(i))


        Next

    End Sub


    ''' <summary>
    ''' This will clear the html code so there are no break lines and you can modify the html code
    ''' This will also change the quotes for the symbol #
    ''' </summary>
    ''' <param name="HtmlCode">This is a Textbox that contains the Html source code</param>
    ''' <remarks></remarks>


    Shared Sub ClearHtmlTBLeaveAmp(HtmlCode As TextBox)

        Try
            ' THIS WILL CLEAR THE EXTRA INFORMATION

            ' Remove new lines since they are not visible in HTML
            HtmlCode.Text = HtmlCode.Text.Replace("\n", " ")

        ' Remove tab spaces
        HtmlCode.Text = HtmlCode.Text.Replace("\t", " ")

        ' Remove multiple white spaces from HTML
        HtmlCode.Text = Regex.Replace(HtmlCode.Text, "\\s+", " ")


        ' Remove HEAD tag
        ' HtmlCode.Text = Regex.Replace(HtmlCode.Text, "<head.*?</head>", "" _
        ' , RegexOptions.IgnoreCase Or RegexOptions.Singleline)

        ' Remove any JavaScript
        HtmlCode.Text = Regex.Replace(HtmlCode.Text, "<script.*?</script>", "" _
        , RegexOptions.IgnoreCase Or RegexOptions.Singleline)

        HtmlCode.Text = HtmlCode.Text.Replace(vbCr, "").Replace(vbLf, "")
        HtmlCode.Text = HtmlCode.Text.Replace("\n", "").Replace("\r", "")
        HtmlCode.Text = HtmlCode.Text.Replace(Chr(10), "").Replace(Chr(13), "")

        HtmlCode.Text = HtmlCode.Text.Replace("""", "#")
        '  HtmlCode.Text = HtmlCode.Text.Replace("&", "")
        HtmlCode.Text = HtmlCode.Text.Replace("'", "")
        HtmlCode.Text = HtmlCode.Text.Replace(Environment.NewLine, "")



        Catch ex As Exception

            End Try

    End Sub


    Shared Sub TextBetweenTagsToListboxSTRING(HtmlCode As String, StartTag1 As String, EndTag1 As String, listbox As ListBox)
        Try


            ' this will check for matches and will add them to a listbox
            Dim contenido As String = HtmlCode
            If contenido <> String.Empty Then
                With listbox
                    ' limpiar el control listbox
                    .DataSource = Nothing
                    ' Mostrar el resultado en el control ListBox
                    .DataSource = Obtener_TextBetweenTags(contenido.ToString, StartTag1, EndTag1)
                    ' MsgBox("Cantidad de Links : " & .Items.Count.ToString, MsgBoxStyle.Information)
                End With
            End If

        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' This will check for matches between tags and will add them to a listbox
    ''' </summary>
    ''' <param name="HtmlCode">This is the html source code. It is a textbox.</param>
    ''' <param name="StartTag1">This is the starting tag and it is a string</param>
    ''' <param name="EndTag1">This is the end tag and it is a string</param>
    ''' <param name="listbox">This is the Listbox where all the text between the tags that matches the tags will be shown</param>
    ''' <remarks></remarks>

    Shared Sub TextBetweenTagsToListbox(HtmlCode As TextBox, StartTag1 As String, EndTag1 As String, listbox As ListBox)
        Try


            ' this will check for matches and will add them to a listbox
            Dim contenido As String = HtmlCode.Text
            If contenido <> String.Empty Then
                With listbox
                    ' limpiar el control listbox
                    .DataSource = Nothing
                    ' Mostrar el resultado en el control ListBox
                    .DataSource = Obtener_TextBetweenTags(contenido.ToString, StartTag1, EndTag1)
                    ' MsgBox("Cantidad de Links : " & .Items.Count.ToString, MsgBoxStyle.Information)
                End With
            End If

        Catch ex As Exception

        End Try
    End Sub


    Shared Function Obtener_TextBetweenTags(ByVal fuente As String, StartTag1 As String, EndTag1 As String) As ArrayList

        Dim temp_arrayList As New ArrayList
        ' expresión regular
        Dim pattern As String = StartTag1 & "(.*?)" & EndTag1
        Try
            ' Colección para obtener los links
            Dim Links As MatchCollection = Regex.Matches(fuente, pattern)
            ' añadirlos
            For Each Link As Match In Links
                temp_arrayList.Add(Replace(Link.Value.ToString, Chr(34), String.Empty))
                ' TextBox6.Text = Link.Value.ToString
            Next
            ' retornar
            Return temp_arrayList
            ' errores
        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally

        End Try
        Return Nothing
    End Function











    Public Shared Function TextBetweenTagsAsString(TagStart As String, TagEnd As String, HtmlCode As String) As String

        Try
            ' once you have the listbox this will separate for specific tags

            Dim Result As String

            Dim Tags5 As List(Of String) = GetTextinTagsPattern(TagStart, TagEnd, HtmlCode)
            ' declaring the starting andend tag
            Dim TagStartSize = TagStart.Length
            Dim TagEndSize = TagEnd.Length
            'You can loop through the list to view all of the results
            For Each Tag As String In Tags5

                'this is the result textbox
                Result = Tag
                Try
                    'this remove the first characteres in the string and the length is equal to the length of the starting tag
                    Dim str9 As String = Result
                    str9 = str9.Remove(0, TagStartSize)
                    Result = str9



                Catch ex As Exception

                End Try
                Try
                    'this remove the last characteres in the string and the length is equal to the length of the end tag
                    Dim s15 As String
                    s15 = Result
                    Result = Result.Substring(0, s15.Length - TagEndSize)


                Catch ex As Exception

                End Try

                Return Result

            Next



        Catch ex As Exception

        End Try


    End Function



    ''' <summary>
    ''' This will check for the text between tags and will display the results in a textbox
    ''' </summary>
    ''' <param name="TagStart">This is the starting tag and it is a string</param>
    ''' <param name="TagEnd">This is the end tag and it is a string</param>
    ''' <param name="HtmlCode">This is the html source code. It is a textbox.</param>
    ''' <param name="Result">This is a textbox that will display the result</param>
    ''' <remarks></remarks>

    Public Shared Sub FindTextBetweenTags(TagStart As String, TagEnd As String, HtmlCode As String, Result As TextBox)

        Try
            ' once you have the listbox this will separate for specific tags

            Dim Tags5 As List(Of String) = GetTextinTagsPattern(TagStart, TagEnd, HtmlCode)
            ' declaring the starting andend tag
            Dim TagStartSize = TagStart.Length
            Dim TagEndSize = TagEnd.Length
            'You can loop through the list to view all of the results
            For Each Tag As String In Tags5

                'this is the result textbox
                Result.Text = Tag
                Try
                    'this remove the first characteres in the string and the length is equal to the length of the starting tag
                    Dim str9 As String = Result.Text
                    str9 = str9.Remove(0, TagStartSize)
                    Result.Text = str9



                Catch ex As Exception

                End Try
                Try
                    'this remove the last characteres in the string and the length is equal to the length of the end tag
                    Dim s15 As String
                    s15 = Result.Text
                    Result.Text = Result.Text.Substring(0, s15.Length - TagEndSize)


                Catch ex As Exception

                End Try
            Next



        Catch ex As Exception

        End Try

    End Sub




    Public Shared Sub FindTextBetweenTagsLB(TagStart As String, TagEnd As String, HtmlCode As String, Result As Label)

        Try
            ' once you have the listbox this will separate for specific tags

            Dim Tags5 As List(Of String) = GetTextinTagsPattern(TagStart, TagEnd, HtmlCode)
            ' declaring the starting andend tag
            Dim TagStartSize = TagStart.Length
            Dim TagEndSize = TagEnd.Length
            'You can loop through the list to view all of the results
            For Each Tag As String In Tags5

                'this is the result textbox
                Result.Text = Tag
                Try
                    'this remove the first characteres in the string and the length is equal to the length of the starting tag
                    Dim str9 As String = Result.Text
                    str9 = str9.Remove(0, TagStartSize)
                    Result.Text = str9



                Catch ex As Exception

                End Try
                Try
                    'this remove the last characteres in the string and the length is equal to the length of the end tag
                    Dim s15 As String
                    s15 = Result.Text
                    Result.Text = Result.Text.Substring(0, s15.Length - TagEndSize)


                Catch ex As Exception

                End Try
            Next



        Catch ex As Exception

        End Try

    End Sub




    Shared Function GetTextinTagsPattern(ByVal TagStart As String, ByVal TagEnd As String, ByVal HTML As String) As List(Of String)
        Dim lMatch As New List(Of String) 'Get the results in a List of strings

        'RegexOptions.IgnoreCase allows case mismatch e.g. if TagName="title" results can include "title", "Title", "TITLE" etc.
        'RegexOptions.Singleline allows .* to see past CarriageReturn characters
        Dim Tag As New Regex(TagStart & "(.*?)" & TagEnd, RegexOptions.IgnoreCase Or RegexOptions.Singleline)
        For Each rMatch As Match In Tag.Matches(HTML)
            lMatch.Add(rMatch.Value)
        Next

        Return lMatch
    End Function


    ''' <summary>
    ''' This will leave a number with 2 digits after the point
    ''' </summary>
    ''' <param name="Textbox">This is the textbox that will be converted into 2 digits after the point</param>
    ''' <remarks></remarks>
    Public Shared Sub FormatLikePriceTB(Textbox As TextBox)
        If Textbox.Text.Contains(".") Then
            Textbox.Text = Format(Convert.ToDouble(Textbox.Text), "####.##")
        Else
            Textbox.Text = Textbox.Text
        End If
    End Sub

    ''' <summary>
    ''' This will leave a number with 2 digits after the point
    ''' </summary>
    ''' <param name="Label">This is the label that will be converted into 2 digits after the point</param>
    ''' <remarks></remarks>
    Public Shared Sub FormatLikePriceLB(Label As Label)
        If Label.Text.Contains(".") Then
            Label.Text = Format(Convert.ToDouble(Label.Text), "####.##")
        Else
            Label.Text = Label.Text
        End If
    End Sub

    ''' <summary>
    '''  This will leave a number with 2 digits after the point
    ''' </summary>
    ''' <param name="Combobox">This is the Combobox that will be converted into 2 digits after the point</param>
    ''' <remarks></remarks>
    Public Shared Sub FormatLikePriceCB(Combobox As ComboBox)
        If Combobox.Text.Contains(".") Then
            Combobox.Text = Format(Convert.ToDouble(Combobox.Text), "####.##")
        Else
            Combobox.Text = Combobox.Text
        End If
    End Sub

    ''' <summary>
    ''' This will generate a random number and place it in a texbox
    ''' </summary>
    ''' <param name="MaxNum">This is the maximumnumber for the random number will be from minnum to maxnum</param>
    ''' <param name="MinNum">This is the minimumnumber for the random number will be from minnum to maxnum</param>
    ''' <param name="TextboxResult">The randum number will be displayed in this textbox</param>
    ''' <remarks></remarks>
    Public Shared Sub RandomNumber(MinNum As Integer, MaxNum As Integer, TextboxResult As TextBox)

        Dim MAXIMUM1 As Integer = MaxNum

        Dim rnd1 = New Random()
        Dim nextValue1 = rnd1.Next(MinNum, MAXIMUM1)
        TextboxResult.Text = nextValue1
    End Sub

    ''' <summary>
    ''' This will generate a random number
    ''' </summary>
    ''' <param name="MaxNum">The maximum number that the random number could be</param>
    ''' <returns></returns>

    Public Shared Function RndNumber(MaxNum As Integer) As String


        Dim MAXIMUM1 As Integer = MaxNum

        Dim rnd1 = New Random()
        Dim nextValue1 = rnd1.Next(0, MAXIMUM1)
        Return nextValue1


    End Function

    ''' <summary>
    ''' This will generate a random string that will with uppercase letters, numbers and lowercase letters
    ''' </summary>
    ''' <param name="CharacterN">The number of characters that you want your string to be</param>
    ''' <returns>Random string with uppercase, lower case letters and numbers</returns>
    Public Shared Function RndString(CharacterN As Integer) As String

        Dim s As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789abcdefghijklmnopqrstuwxyz"
        Dim r As New Random
        Dim sb As New StringBuilder
        For i As Integer = 1 To CharacterN
            Dim idx As Integer = r.Next(0, 60)
            sb.Append(s.Substring(idx, 1))
        Next
        Return sb.ToString()

    End Function





    ''' <summary>
    ''' This will remove the dupplicated items from a listbox
    ''' </summary>
    ''' <param name="listbox">The listbox that contains the dupplicated items that you want to clean</param>
    ''' <remarks></remarks>
    Public Shared Sub RemoveDuplicatedListbox(listbox As ListBox)
        Try
            listbox.Sorted = True
        listbox.Refresh()

        Dim index As Integer
        Dim itemcount As Integer = listbox.Items.Count

        If itemcount > 1 Then
            Dim lastitem As String = listbox.Items(itemcount - 1)

            For index = itemcount - 2 To 0 Step -1
                If listbox.Items(index) = lastitem Then
                    listbox.Items.RemoveAt(index)
                Else
                    lastitem = listbox.Items(index)
                End If
            Next
        End If


        Catch ex As Exception

            End Try
    End Sub


    Public Shared Sub RemoveAllLetters(textbox As TextBox)
        Try

            textbox.Text = textbox.Text.Replace("a", "")
            textbox.Text = textbox.Text.Replace("b", "")
            textbox.Text = textbox.Text.Replace("c", "")
            textbox.Text = textbox.Text.Replace("d", "")
            textbox.Text = textbox.Text.Replace("e", "")
            textbox.Text = textbox.Text.Replace("f", "")
            textbox.Text = textbox.Text.Replace("g", "")
            textbox.Text = textbox.Text.Replace("h", "")
            textbox.Text = textbox.Text.Replace("i", "")
            textbox.Text = textbox.Text.Replace("j", "")
            textbox.Text = textbox.Text.Replace("j", "")
            textbox.Text = textbox.Text.Replace("l", "")
            textbox.Text = textbox.Text.Replace("m", "")
            textbox.Text = textbox.Text.Replace("n", "")
            textbox.Text = textbox.Text.Replace("o", "")
            textbox.Text = textbox.Text.Replace("p", "")
            textbox.Text = textbox.Text.Replace("q", "")
            textbox.Text = textbox.Text.Replace("r", "")
            textbox.Text = textbox.Text.Replace("s", "")
            textbox.Text = textbox.Text.Replace("t", "")
            textbox.Text = textbox.Text.Replace("u", "")
            textbox.Text = textbox.Text.Replace("v", "")
            textbox.Text = textbox.Text.Replace("w", "")
            textbox.Text = textbox.Text.Replace("x", "")
            textbox.Text = textbox.Text.Replace("y", "")
            textbox.Text = textbox.Text.Replace("z", "")
            textbox.Text = textbox.Text.Replace("A", "")
            textbox.Text = textbox.Text.Replace("B", "")
            textbox.Text = textbox.Text.Replace("C", "")
            textbox.Text = textbox.Text.Replace("D", "")
            textbox.Text = textbox.Text.Replace("E", "")
            textbox.Text = textbox.Text.Replace("F", "")
            textbox.Text = textbox.Text.Replace("G", "")
            textbox.Text = textbox.Text.Replace("H", "")
            textbox.Text = textbox.Text.Replace("I", "")
            textbox.Text = textbox.Text.Replace("J", "")
            textbox.Text = textbox.Text.Replace("K", "")
            textbox.Text = textbox.Text.Replace("L", "")
            textbox.Text = textbox.Text.Replace("M", "")
            textbox.Text = textbox.Text.Replace("N", "")
            textbox.Text = textbox.Text.Replace("O", "")
            textbox.Text = textbox.Text.Replace("P", "")
            textbox.Text = textbox.Text.Replace("Q", "")
            textbox.Text = textbox.Text.Replace("R", "")
            textbox.Text = textbox.Text.Replace("S", "")
            textbox.Text = textbox.Text.Replace("T", "")
            textbox.Text = textbox.Text.Replace("U", "")
            textbox.Text = textbox.Text.Replace("V", "")
            textbox.Text = textbox.Text.Replace("X", "")
            textbox.Text = textbox.Text.Replace("Y", "")
            textbox.Text = textbox.Text.Replace("Z", "")

        Catch ex As Exception

        End Try
    End Sub


    ''' <summary>
    ''' This will extract the emails from a textbox andplace them in a listbox
    ''' </summary>
    ''' <param name="source">This is the textbox html source that contains the emails that you want to scrape</param>
    ''' <param name="listbox1">This is the listbox where all the emails will be added</param>
    ''' <remarks></remarks>

    Public Shared Sub ExtractEmailAddressesFromString(source As TextBox, listbox1 As ListBox)

        Dim mc As MatchCollection
        Dim i As Integer

        Dim htmlsourcecode As String = source.Text

        ' In this section i can change the patter of my regular expressions. whatever the match is, it will be added to a listbox
        mc = Regex.Matches(htmlsourcecode, "([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})")



        Dim results(mc.Count - 1) As String
        For i = 0 To results.Length - 1
            results(i) = mc(i).Value
            listbox1.Items.Add(mc(i).Value)
        Next

        ' Return results
    End Sub

    ''' <summary>
    ''' You have a custom regular expresion? then use this subroutine to add all the matches to a listbox
    ''' </summary>
    ''' <param name="source">
    ''' This is the texbox that contains the text source where you will run the regular expression
    '''  to search for the information with your your specific pattern
    ''' </param>
    ''' <param name="listbox1">This is the texbox where it will add all the matches in your regular expression pattern</param>
    ''' <param name="pattern">
    ''' This is the patter of your regular expression you can use diferent patters. It is a string
    ''' Some samples are 
    ''' Date: (0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])[- /.](19|20)[0-9]{2}
    ''' Returns values like  1999/12/31 or  1999.12.31
    ''' 
    ''' Email Address: (?i)([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3}) 
    ''' Return values like Dell@ireland.com
    '''
    ''' </param>
    ''' <remarks></remarks>

    Public Shared Sub RegexMatchestoListBox(source As TextBox, listbox1 As ListBox, pattern As String)

        Dim mc As MatchCollection
        Dim i As Integer
        Dim htmlsourcecode As String = source.Text
        ' In this section i can change the patter of my regular expressions. whatever the match is, it will be added to a listbox
        mc = Regex.Matches(htmlsourcecode, pattern)

        Dim results(mc.Count - 1) As String
        For i = 0 To results.Length - 1
            results(i) = mc(i).Value
            listbox1.Items.Add(mc(i).Value)
        Next


    End Sub

    ''' <summary>
    ''' This will save the Listbox Items to a text File
    ''' </summary>
    ''' <param name="SaveFileDialog">This is the save file dialog that will be use to get the path where the file will be saved</param>
    ''' <param name="listbox">This is the listbox that you will save into a text file</param>
    ''' <remarks></remarks>

    Public Shared Sub SaveListboxToTextFile(SaveFileDialog As SaveFileDialog, listbox As ListBox)

        SaveFileDialog.ShowDialog()

        Dim saveurl As String = SaveFileDialog.FileName

        Using SW As New IO.StreamWriter(saveurl & ".txt", True)
            For Each itm As String In listbox.Items
                SW.WriteLine(itm)
            Next
        End Using

    End Sub


    Public Shared Sub RemoveEmptyLinesTB(textbox As TextBox)

        textbox.Text = textbox.Text.Replace(vbLf & vbCr, Environment.NewLine)


    End Sub

    ''' <summary>
    ''' This will remove all the tags in a Html code. Leaving raw data without the quotes
    ''' </summary>
    ''' <param name="textbox">This is the textbox source and result where the quotes will be replaced</param>
    ''' <remarks></remarks>

    Public Shared Sub RemoveAllTagshtml(textbox As TextBox)

        ' this will remove the tags
        textbox.Text = Regex.Replace(textbox.Text, "<.*?>", String.Empty)

    End Sub


    ''' <summary>
    ''' This will remove everything after the symbol or string of your choice. It will also remove the symbol.
    ''' </summary>
    ''' <param name="textbox">This is the textbox that contains the string with the symbol</param>
    ''' <param name="symbol">This is the symbol that i will use as delimitator</param>
    ''' <remarks></remarks>

    Public Shared Sub RemoveEverythingAfterSymbol(textbox As TextBox, symbol As String)

        Try
            If textbox.Text.Contains(symbol) = True Then
            'DELETE EVERITING AFTERT POINT   ' INTEGERS
            textbox.Text = textbox.Text.Substring(0, textbox.Text.IndexOf(symbol))

        End If

        Catch ex As Exception

            End Try
    End Sub


    Public Shared Function RemoveEverythingAfterSymbolString(String1 As String, symbol As String) As String

        Try
            If String1.Contains(symbol) = True Then
                'DELETE EVERITING AFTERT POINT   ' INTEGERS
                Return String1.Substring(0, String1.IndexOf(symbol))

            End If

        Catch ex As Exception

        End Try
    End Function


    Public Shared Sub RemoveEverythingAfterSymbolCB(textbox As ComboBox, symbol As String)

        Try
            If textbox.Text.Contains(symbol) = True Then
                'DELETE EVERITING AFTERT POINT   ' INTEGERS
                textbox.Text = textbox.Text.Substring(0, textbox.Text.IndexOf(symbol))

            End If

        Catch ex As Exception

        End Try
    End Sub

    Public Shared Sub RemoveEverythingAfterSymbolLB(textbox As Label, symbol As String)

        Try
            If textbox.Text.Contains(symbol) = True Then
                'DELETE EVERITING AFTERT POINT   ' INTEGERS
                textbox.Text = textbox.Text.Substring(0, textbox.Text.IndexOf(symbol))

            End If

        Catch ex As Exception

        End Try
    End Sub


    ''' <summary>
    ''' This will remove everything before the symbol or string of your choice. It will also remove the symbol.
    ''' </summary>
    ''' <param name="textbox">This is the textbox that contains the string with the symbol</param>
    ''' <param name="symbol">This is the symbol that i will use as delimitator</param>
    ''' <remarks></remarks>

    Public Shared Sub RemoveEverythingBeforeSymbol(textbox As TextBox, symbol As String)
        Try
            'DELETE EVERYTHING BEFORE THE SYMBOL > 'this was added to remove the extra information
            Dim index As Integer = textbox.Text.IndexOf(symbol)
        Dim output As String = textbox.Text.Substring(index, textbox.Text.Length - index)
        textbox.Text = output
        Dim sizestring As Integer = symbol.Length
            Scrapping.RemoveFirstCharacters(textbox, sizestring) ' this will remove the > simbol

        Catch ex As Exception

            End Try
    End Sub


    ''' <summary>
    ''' This will remove the duplicated rows in a datagridview you can selet the column that you will compare to see if they are dupplicated
    ''' </summary>
    ''' <param name="Datagridview">This is teh datagridview that you will use to remove the dupplicated. Do not use it on bound datagridviews</param>
    ''' <param name="column">This is the column that contains the dupplicated data. Remove 1 from the number of the column</param>
    ''' <remarks></remarks>
    Public Shared Sub RemoveDuplicatedDGV(Datagridview As DataGridView, column As Integer)


        'this will remove the duplicated items

        Try
            For i As Integer = 0 To Datagridview.RowCount - 2
                For j As Integer = i + 1 To Datagridview.RowCount - 2
                    If Datagridview.Rows(i).Cells(column).Value = Datagridview.Rows(j).Cells(column).Value Then

                        Datagridview.Rows.Remove(Datagridview.Rows(i))
                        i -= 1
                        Debug.Print("duplicated value " & Datagridview.Rows(i).Cells(column).Value)
                        'DataGridView1.Rows(i).Cells(1).Value = "Duplicate"
                    End If

                Next



            Next


        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>
    ''' This will get the local ip from your computer in your network
    ''' </summary>
    ''' <param name="textbox">This is the textbox where the ip will be displayed</param>
    ''' <remarks></remarks>
    Public Shared Sub GetMyLocalIp(textbox As TextBox)

        Dim hostname As String = Dns.GetHostName()
        Dim ipaddress As String = CType(Dns.GetHostByName(hostname).AddressList.GetValue(0), IPAddress).ToString

        textbox.Text = ipaddress

    End Sub

    ''' <summary>
    ''' This will make the app to start with windows
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub MakeMyAppStartWithWindows()


        My.Computer.FileSystem.CopyFile(Application.ExecutablePath,
"C:\Documents and Settings\All Users\Start Menu\Programs\Startup\" & My.Application.Info.AssemblyName & ".exe")


    End Sub



    ' ====================== in this section i add for posting to webbrowser control========================

    ''' <summary>
    ''' This is subroutine will place the value of a texbox from the form into a textbox inside the webbrowser control 
    ''' </summary>
    ''' <param name="Webbrowser1">This is the webbrowser control that contains the textbox where you will place the textbox value from your form</param>
    ''' <param name="IDelement">This is the id name. It is used to select the specific textbox Check html to find it</param>
    ''' <param name="TextboxValue">This is the textbox that contains the value that will be placed in the webbrowser textbox </param>
    ''' <remarks></remarks>
    Public Shared Sub SetValueTBWebbrowser(Webbrowser1 As WebBrowser, IDelement As String, TextboxValue As TextBox)
        'this is for textboxes mainly
        Webbrowser1.Document.GetElementById(IDelement).SetAttribute("value", TextboxValue.Text)



    End Sub

    ''' <summary>
    ''' This will click the button with the value that you select in the parameter
    ''' </summary>
    ''' <param name="Webbrowser1">This is the webbrowser that contains the button</param>
    ''' <param name="Buttonvalue">This is the value that will allow you to select wich button i want to click</param>
    ''' <remarks></remarks>
    Public Shared Sub ClickBTNWithValue(Webbrowser1 As WebBrowser, Buttonvalue As String)

        'this will click the button with the value selected "Login"
        Dim allelements As HtmlElementCollection = Webbrowser1.Document.All
        For Each webpageelement As HtmlElement In allelements
            If webpageelement.GetAttribute("value") = Buttonvalue Then
                webpageelement.InvokeMember("click")
            End If
        Next

    End Sub

    ''' <summary>
    ''' This will click the button with the ID that you select in the parameter
    ''' </summary>
    ''' <param name="Webbrowser1">This is the webbrowser that contains the button</param>
    ''' <param name="ButtonID">This is the value of the id that will allow you to select wich button i want to click</param>
    ''' <remarks></remarks>

    Public Shared Sub ClickBTNWithID(Webbrowser1 As WebBrowser, ButtonID As String)

        'this will click the button with the value selected "Login"
        Dim allelements As HtmlElementCollection = Webbrowser1.Document.All
        For Each webpageelement As HtmlElement In allelements
            If webpageelement.GetAttribute("id") = ButtonID Then
                webpageelement.InvokeMember("click")
            End If
        Next

    End Sub

    ''' <summary>
    '''  This will click the button with the ID that you select in the parameter
    ''' </summary>
    ''' <param name="Webbrowser1">This is the webbrowser that contains the button</param>
    ''' <param name="ButtonNameValue">This is the value of the name that will allow you to select wich button i want to click</param>
    ''' <remarks></remarks>
    Public Shared Sub ClickBTNWithName(Webbrowser1 As WebBrowser, ButtonNameValue As String)

        'this will click the button with the value selected "Login"
        Dim allelements As HtmlElementCollection = Webbrowser1.Document.All
        For Each webpageelement As HtmlElement In allelements
            If webpageelement.GetAttribute("name") = ButtonNameValue Then
                webpageelement.InvokeMember("click")
            End If
        Next

    End Sub

    ''' <summary>
    ''' This will click the button with the class value that you select in the parameter
    ''' </summary>
    ''' <param name="Webbrowser1">This is the webbrowser that contains the button</param>
    ''' <param name="ClassName">This is the value of the class that will allow you to select wich button i want to click</param>
    ''' <remarks></remarks>
    Public Shared Sub ClickBTNWithClass(Webbrowser1 As WebBrowser, ClassName As String)

        'this will click the button with the value selected "Login"
        Dim allelements As HtmlElementCollection = Webbrowser1.Document.All
        For Each webpageelement As HtmlElement In allelements
            If webpageelement.GetAttribute("class") = ClassName Then
                webpageelement.InvokeMember("click")

                webpageelement.InvokeMember("click")

            End If
        Next

    End Sub

    ''' <summary>
    ''' This will click a button with the selected anchor text, this also works for links
    ''' </summary>
    ''' <param name="Webbrowser1">This is the webbrowser that contains the button</param>
    ''' <param name="AnchorText">This is the Anchor text that is present in the button</param>
    ''' <param name="MainTagName">
    ''' This is the name of the tag. it can be something like this.
    ''' 'tag name "BUTTON"
    ''' 'tag name "a"
    ''' 'tag name "input"
    ''' </param>
    ''' <remarks></remarks>
    Public Shared Sub ClickBTNByAnchorText(Webbrowser1 As WebBrowser, AnchorText As String, MainTagName As String)


        For Each elem As HtmlElement In Webbrowser1.Document.GetElementsByTagName(MainTagName) 'tag name "BUTTON"
            If elem.InnerText = AnchorText Then
                elem.InvokeMember("click")
            End If
        Next


    End Sub

    ''' <summary>
    ''' This will focus on the element that you selected with the help of the id value
    ''' </summary>
    ''' <param name="WebBrowser1">This is the webbrowser that contains the textbox that you want to focus on</param>
    ''' <param name="idvalue">This is the value of the id that will help you select the textbox</param>
    ''' <remarks></remarks>
    Public Shared Sub FocusTBwithID(WebBrowser1 As WebBrowser, idvalue As String)
        Dim allelements2 As HtmlElementCollection = WebBrowser1.Document.All
        For Each webpageelement4 As HtmlElement In allelements2
            If webpageelement4.GetAttribute("id") = idvalue Then
                webpageelement4.Focus()
            End If
        Next
    End Sub

    ''' <summary>
    ''' This will focus on the element that you selected with the help of the value
    ''' </summary>
    ''' <param name="WebBrowser1">This is the webbrowser that contains the textbox that you want to focus on</param>
    ''' <param name="idvalue">this is the value of the value parameter in the textbox</param>
    ''' <remarks></remarks>
    Public Shared Sub FocusTBwithValue(WebBrowser1 As WebBrowser, idvalue As String)
        Dim allelements2 As HtmlElementCollection = WebBrowser1.Document.All
        For Each webpageelement4 As HtmlElement In allelements2
            If webpageelement4.GetAttribute("value") = idvalue Then
                webpageelement4.Focus()
            End If
        Next
    End Sub
    ''' <summary>
    ''' This will focus on the element that you selected with the help of the value of the name
    ''' </summary>
    ''' <param name="WebBrowser1">This is the webbrowser that contains the textbox that you want to focus on</param>
    ''' <param name="idvalue">This is the value of the name.</param>
    ''' <remarks></remarks>
    Public Shared Sub FocusTBwithName(WebBrowser1 As WebBrowser, idvalue As String)
        Dim allelements2 As HtmlElementCollection = WebBrowser1.Document.All
        For Each webpageelement4 As HtmlElement In allelements2
            If webpageelement4.GetAttribute("name") = idvalue Then
                webpageelement4.Focus()
            End If
        Next
    End Sub

    ''' <summary>
    ''' This will focus on the text area. For example form with text area where you enter a multiline information
    ''' </summary>
    ''' <param name="WebBrowser1">This is the webbrowser that contains the textarea that you want to focus on</param>
    ''' <param name="NameValue">This is the value of the name.</param>
    ''' <remarks></remarks>
    Public Shared Sub FocusTextAreaWithName(WebBrowser1 As WebBrowser, NameValue As String)
        Dim htmlElements As HtmlElementCollection = WebBrowser1.Document.GetElementsByTagName("textarea")
        For Each el As HtmlElement In htmlElements
            If el.GetAttribute("name").Equals(NameValue) Then
                ' el.SetAttribute("Value", TextBox3.Text)
                el.Focus()
            End If
        Next
    End Sub

    ''' <summary>
    ''' This will focus on the text area. For example form with text area where you enter a multiline information
    ''' </summary>
    ''' <param name="WebBrowser1">This is the webbrowser that contains the textarea that you want to focus on</param>
    ''' <param name="NameValue">This is the value of the id.</param>
    ''' <remarks></remarks>
    Public Shared Sub FocusTextAreaWithID(WebBrowser1 As WebBrowser, NameValue As String)
        Dim htmlElements As HtmlElementCollection = WebBrowser1.Document.GetElementsByTagName("textarea")
        For Each el As HtmlElement In htmlElements
            If el.GetAttribute("id").Equals(NameValue) Then
                ' el.SetAttribute("Value", TextBox3.Text)
                el.Focus()
            End If
        Next
    End Sub

    ''' <summary>
    '''  This will focus on the text area. For example form with text area where you enter a multiline information
    ''' </summary>
    ''' <param name="WebBrowser1">This is the webbrowser that contains the textarea that you want to focus on</param>
    ''' <param name="NameValue">This is the value of the value parameter</param>
    ''' <remarks></remarks>
    Public Shared Sub FocusTextAreaWithValue(WebBrowser1 As WebBrowser, NameValue As String)
        Dim htmlElements As HtmlElementCollection = WebBrowser1.Document.GetElementsByTagName("textarea")
        For Each el As HtmlElement In htmlElements
            If el.GetAttribute("value").Equals(NameValue) Then
                ' el.SetAttribute("Value", TextBox3.Text)
                el.Focus()
            End If
        Next
    End Sub

    ''' <summary>
    ''' When you have a dropdown list or select. you can use this subroutine to set the value. the value that you set should exist in the select options
    ''' </summary>
    ''' <param name="WebBrowser1">This is the webbrowser that contains the the dropdown list control</param>
    ''' <param name="SETValue">This is the value that will be set. It should exist in the dropdownlist. It should be the value not the text</param>
    ''' <param name="NameControl">This is the value of the name atrribute</param>
    ''' <remarks></remarks>

    Public Shared Sub SetValueDropDownListWithName(WebBrowser1 As WebBrowser, SETValue As String, NameControl As String)

        Dim theElementCollection1 = WebBrowser1.Document.GetElementsByTagName("select")
        For Each curElement As HtmlElement In theElementCollection1
            Dim controlName As String = curElement.GetAttribute("name").ToString
            If controlName = NameControl Then
                curElement.SetAttribute("Value", SETValue)
            End If
        Next

    End Sub

    ''' <summary>
    '''  When you have a dropdown list or select. you can use this subroutine to set the value. the value that you set should exist in the select options
    '''
    ''' </summary>
    ''' <param name="WebBrowser1">This is the webbrowser that contains the the dropdown list control</param>
    ''' <param name="SETValue">This is the value that will be set. It should exist in the dropdownlist. It should be the value not the text</param>
    ''' <param name="IDValue">This is the value of the id atrribute</param>
    ''' <remarks></remarks>
    Public Shared Sub SetValueDropDownListWithID(WebBrowser1 As WebBrowser, SETValue As String, IDValue As String)
        Try
            Dim theElementCollection1 = WebBrowser1.Document.GetElementsByTagName("select")
        For Each curElement As HtmlElement In theElementCollection1
            Dim controlName As String = curElement.GetAttribute("id").ToString
            If controlName = IDValue Then
                curElement.SetAttribute("Value", SETValue)
            End If
        Next


        Catch ex As Exception

            End Try
    End Sub

    ''' <summary>
    ''' This will Paste the content inside the focused textbox or text area
    ''' </summary>
    ''' <param name="textbox">This is the textbox that contains the text that will be pasted</param>
    ''' <remarks></remarks>
    Public Shared Sub PasteContentOnFocusedTB(textbox As TextBox)
        Clipboard.SetText(textbox.Text)
        SendKeys.Send("^(v)")
    End Sub

    ''' <summary>
    ''' This will Paste the content inside the focused textbox or text area
    ''' </summary>
    ''' <param name="label">This is the label that contains the text that will be pasted</param>
    ''' <remarks></remarks>
    Public Shared Sub PasteContentOnFocusedLB(label As Label)
        Clipboard.SetText(label.Text)
        SendKeys.Send("^(v)")
    End Sub
    ''' <summary>
    ''' This will Paste the content inside the focused textbox or text area
    ''' </summary>
    ''' <param name="combobox">This is the combobox that contains the text that will be pasted</param>
    ''' <remarks></remarks>
    Public Shared Sub PasteContentOnFocusedCB(combobox As ComboBox)
        Clipboard.SetText(combobox.Text)
        SendKeys.Send("^(v)")
    End Sub

    ''' <summary>
    ''' This will Paste the content inside the focused textbox or text area
    ''' </summary>
    ''' <param name="stringvalue">This is the string that contains the text that will be pasted</param>
    ''' <remarks></remarks>
    Public Shared Sub PasteContentOnFocusedString(stringvalue As String)
        Clipboard.SetText(stringvalue)
        SendKeys.Send("^(v)")
    End Sub



    Public Shared Function ShowAsCurrency(textbox As String)

        Dim value As Decimal = textbox
        '   Dim currencyvalue As String = String.Format("{0:n}", value)

        Dim currencyvalue As String = value.ToString("C2")

        Return currencyvalue

    End Function


    ''' <summary>
    ''' This will send the enter key. Once you focused on a button you can use this to submit the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub HitEnterKey()

        SendKeys.Send("{ENTER}")

    End Sub


    ''' <summary>
    ''' This will click a link based on the Anchor text provided on the parameter
    ''' </summary>
    ''' <param name="webbrowser1">This is the webbrowser that contains the link that you want to click</param>
    ''' <param name="AnchorText">This is the anchor of the link that you want to click</param>
    ''' <remarks></remarks>
    Public Shared Sub ClickLinkWithAnchorText(webbrowser1 As WebBrowser, AnchorText As String)

        Dim a As HtmlElement = GetAnchorWithLabel(webbrowser1, AnchorText)
        a.InvokeMember("click")

    End Sub
    Shared Function GetAnchorWithLabel(Webbrowser1 As WebBrowser, ByVal sLabel As String) As HtmlElement

        Dim anchors As HtmlElementCollection = Webbrowser1.Document.GetElementsByTagName("a")
        For Each anc As HtmlElement In anchors
            If anc.InnerText = sLabel Then Return anc
        Next

        Return Nothing ''failed to locate anchor, return nothing

    End Function

    '--------------------------------------------------------
    ''' <summary>
    ''' This will spin the text for example {hello|hi|hey} will return one of the values inside the spin sintax
    ''' everytime you run this it will select a different value ramdomnly 
    ''' </summary>
    ''' <param name="TBSppinerSource">This is the textbox that will be the source It will contain all the spin sintaxis</param>
    ''' <param name="TBSppinedText">This will be the textbox that will receive the value of the sppined article</param>
    ''' <remarks></remarks>
    Public Shared Sub SpinTextNow(TBSppinerSource As TextBox, TBSppinedText As TextBox)

        TBSppinedText.Text = ""
        Dim rnd As New Random
        Dim text As String = TBSppinerSource.Text
        TBSppinedText.Text = spintext(rnd, [text])

    End Sub


    Private Shared Function spintext(ByVal rnd As Random, ByVal str As String) As String
        Dim pattern As String = "{[^{}]*}"
        Dim match As Match = Regex.Match(str, pattern)
        Do While match.Success
            Dim strArray As String() = str.Substring((match.Index + 1), (match.Length - 2)).Split(New Char() {"|"c})
            str = (str.Substring(0, match.Index) & strArray(rnd.Next(strArray.Length)) & str.Substring((match.Index + match.Length)))
            match = Regex.Match(str, pattern)
        Loop
        Return str
    End Function

    '--------------------------------------------------------------------

    ''' <summary>
    ''' This will read a csv file and it will display it in a datagrid view
    ''' </summary>
    ''' <param name="Datagridview1">This is the datagridview that will display the results from the csv file</param>
    ''' <param name="openfiledialog1">This is the open file dialog that will be used to locate the csv file that you want to open in your software</param>
    ''' <param name="tbfolder">This is a textbox that will hold the value of the folder</param>
    ''' <param name="tbfilename">This is a textbox that will hold the value of the folder. Should contain the .csv or .txt extention</param>
    ''' <remarks></remarks>
    ''' 
    Public Shared Sub ReadCSVtoDGV(Datagridview1 As DataGridView, openfiledialog1 As OpenFileDialog, tbfolder As TextBox, tbfilename As TextBox)

        Try
            openfiledialog1.ShowDialog()

            tbfolder.Text = (Path.GetDirectoryName(openfiledialog1.FileName)) & "\"

            tbfilename.Text = openfiledialog1.SafeFileName

            Dim csvFileFolder As String = tbfolder.Text
            Dim csvFileName As String = tbfilename.Text
            Dim connString As String = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" _
                & csvFileFolder & ";Extended Properties=""Text;HDR=No;FMT=Delimited"""
            Dim conn As New Odbc.OdbcConnection(connString)
            'Open a data adapter, specifying the file name to load
            Dim da As New Odbc.OdbcDataAdapter("SELECT * FROM [" & csvFileName & "]", conn)
            'Then fill a data table, which can be bound to a grid
            Dim dt As New DataTable
            da.Fill(dt)
            Datagridview1.DataSource = dt
            With Datagridview1
                '.AutoGenerateColumns = True
                .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End With

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub



    Public Shared Sub GetCsvData(Datagridview1 As DataGridView, openfiledialog1 As OpenFileDialog, tbfolder As TextBox, tbfilename As TextBox)


        Try

            openfiledialog1.ShowDialog()


            tbfolder.Text = (Path.GetDirectoryName(openfiledialog1.FileName)) & "\"

            tbfilename.Text = openfiledialog1.SafeFileName

            Dim csvFileFolder As String = tbfolder.Text
            Dim csvFileName As String = tbfilename.Text
            Dim connString As String = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" _
                & csvFileFolder & ";Extended Properties=""Text;HDR=No;FMT=Delimited"""
            Dim conn As New Odbc.OdbcConnection(connString)
            'Open a data adapter, specifying the file name to load
            Dim da As New Odbc.OdbcDataAdapter("SELECT * FROM [" & csvFileName & "]", conn)
            'Then fill a data table, which can be bound to a grid
            Dim dt As New DataTable
            da.Fill(dt)
            Datagridview1.DataSource = dt
            With Datagridview1
                '.AutoGenerateColumns = True
                .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End With




        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub





    ''' <summary>
    ''' This will read an Excel file and will display the columns into a datagridview 
    ''' </summary>
    ''' <param name="openFileDialog1">This is the open filedialog that will be used to locate the file</param>
    ''' <param name="DatagridView">This is the datagridview that will display the results from the Excel file</param>
    ''' <remarks></remarks>

    Public Shared Sub ReadExcelFile(openFileDialog1 As OpenFileDialog, DatagridView As DataGridView)

        openFileDialog1.ShowDialog()

        Dim FilenameLocation = openFileDialog1.FileName

        Try

            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim DtSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FilenameLocation & ";Extended Properties=""Excel 12.0 Xml;HDR=YES"";")
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection)
            MyCommand.TableMappings.Add("Table", "josegarcia.ca")
            DtSet = New System.Data.DataSet
            MyCommand.Fill(DtSet)
            DatagridView.DataSource = DtSet.Tables(0)
            MyConnection.Close()

            'the data will be bounded, remember to copy all data from datagridview into another datagridview So you can modify and remove rows

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub


    ''' <summary>
    ''' This will move the selected row 1 place upwards
    ''' </summary>
    ''' <param name="datagridviewtest">This is the datagridview that contains the row that you want to move up</param>
    ''' <remarks></remarks>
    Public Shared Sub MoveUpselectedDGV(datagridviewtest As DataGridView)

        Try
            If (datagridviewtest.SelectedCells.Count > 0) Then
                Dim curr_index As Integer = datagridviewtest.CurrentCell.RowIndex
                If curr_index <> 0 Then
                    Dim curr_col_index As Integer = datagridviewtest.CurrentCell.ColumnIndex
                    Dim curr_row As DataGridViewRow = datagridviewtest.CurrentRow
                    datagridviewtest.Rows.Remove(curr_row)
                    datagridviewtest.Rows.Insert(curr_index + (-1), curr_row)
                    datagridviewtest.CurrentCell = datagridviewtest(curr_col_index, curr_index + (-1))
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    '''  This will move the selected row 1 place downwards
    ''' </summary>
    ''' <param name="datagridviewtest">This is the datagridview that contains the row that you want to move down</param>
    ''' <remarks></remarks>
    Public Shared Sub MoveDownSelectedDGV(datagridviewtest As DataGridView)

        Try
            If (datagridviewtest.SelectedCells.Count > 0) Then
                Dim curr_index As Integer = datagridviewtest.CurrentCell.RowIndex
                Dim totalnumberofrows As Integer = datagridviewtest.Rows.Count - 1
                If curr_index <> totalnumberofrows - 1 Then
                    Dim curr_col_index As Integer = datagridviewtest.CurrentCell.ColumnIndex
                    Dim curr_row As DataGridViewRow = datagridviewtest.CurrentRow
                    datagridviewtest.Rows.Remove(curr_row)
                    datagridviewtest.Rows.Insert(curr_index + 1, curr_row)
                    datagridviewtest.CurrentCell = datagridviewtest(curr_col_index, curr_index + 1)
                End If
            End If
        Catch ex As Exception
        End Try

    End Sub


    Public Shared Sub RunJavaScriptCode(WebBrowser3 As WebBrowser, javascriptcode As String, javascriptfunctionname As String)


        ' sample string javascript   "function toggle(source) { checkboxes = document.querySelectorAll('input[type=checkbox]'); for(var i=0, n=checkboxes.length;i<n;i++) {   checkboxes[i].checked = source.checked;  }}"


        Dim head As HtmlElement = WebBrowser3.Document.GetElementsByTagName("head")(0)
        Dim script1 As HtmlElement = WebBrowser3.Document.CreateElement("script")
        script1.SetAttribute("text", javascriptcode)
        head.AppendChild(script1)
        WebBrowser3.Document.InvokeScript(javascriptfunctionname)


    End Sub

    ''' <summary>
    ''' This will select the row that you specified. 
    ''' </summary>
    ''' <param name="datagridview">This is the datagridview that contains the row that you want to select</param>
    ''' <param name="row">This is the row that you want to select. it start at 1</param>
    ''' <remarks></remarks>
    Public Shared Sub SelectRowDGV(datagridview As DataGridView, row As Integer)
        datagridview.Rows(row - 1).Selected = True
    End Sub


    Public Shared Sub SelectRowDGV0(datagridview As DataGridView, row As Integer)
        datagridview.Rows(row).Selected = True
    End Sub

    ''' <summary>
    ''' This will send an email with your gmail account.
    ''' You have to go to your gmail account settings and activate the realying pop3 option
    ''' All of the parameters are string. Then decalre your variables to pass them in this subroutine
    ''' </summary>
    ''' <param name="EmailSender">This is your gmail email. The one used to send the email.</param>
    ''' <param name="Password">This is the password of your email</param>
    ''' <param name="EmailReceiver">This is the person that will receive the email</param>
    ''' <param name="Subject">This is the subject of the email</param>
    ''' <param name="EmailBody">This is the email body.</param>
    ''' <remarks></remarks>

    Public Shared Sub SendEmailGMAIL(EmailSender As String, Password As String, EmailReceiver As String, Subject As String, EmailBody As String)


        Dim MyMailMessage As New MailMessage()
        Try
            MyMailMessage.From = New MailAddress(EmailSender)
            MyMailMessage.To.Add(EmailReceiver)
            MyMailMessage.Subject = Subject
            MyMailMessage.Body = EmailBody
            Dim SMTP As New SmtpClient("smtp.gmail.com")
            SMTP.Port = 587
            SMTP.EnableSsl = True
            SMTP.Credentials = New System.Net.NetworkCredential(EmailSender, Password)
            SMTP.Send(MyMailMessage)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub


    ''' <summary>
    ''' This will send an email with an attachemnt. Go to the class and change the information based on your server configuration
    ''' </summary>
    ''' <param name="AdressFrom">This is you email or the email that you will use to send the email</param>
    ''' <param name="AdressTo">This is the receiver email address</param>
    ''' <param name="Subject">This is the subject of the email</param>
    ''' <param name="Body">This is the body</param>
    ''' <param name="Attached">Here you type the file path.</param>
    ''' <remarks></remarks>

    Public Shared Sub SendMailWithAttach(AdressFrom As String, AdressTo As String, Subject As String, Body As String, Attached As String)
        Try
            Dim MyMailMessage As New MailMessage()
            MyMailMessage.From = New MailAddress(AdressFrom)
            MyMailMessage.To.Add(AdressTo)
            MyMailMessage.Subject = Subject
            MyMailMessage.Body = Body

            If Attached <> "" Then
                Dim MyAttachment As Attachment
                MyAttachment = New System.Net.Mail.Attachment(Attached)
                MyMailMessage.Attachments.Add(MyAttachment)
            End If

            Dim SMTP As New SmtpClient("your.stmp.server")
            SMTP.Credentials = New System.Net.NetworkCredential("username (email or user depending your server)", "password")
            SMTP.Port = 25  'this is your port
            SMTP.Send(MyMailMessage)
            MsgBox("Email sent successfully")
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' Sometimes you want to submit information into a PhP form.
    ''' This subroutine helps you to make a post request to the php form selected
    ''' This is useful because you do not need a webbrowser
    ''' </summary>
    ''' <param name="UrlToPost">This is the url that will process the information that you post into it</param>
    ''' <param name="PostDataString">
    ''' This are the parameters that you will pass to the form. They are considered post information
    ''' Example  Dim postData = "post_name=" ampersand posturl
    ''' </param>
    ''' <remarks></remarks>

    Public Shared Sub SendPostInfoPHP(UrlToPost As String, PostDataString As String)




        '  Dim postData =   'name1=value1&name2=value2
        '  Dim request As WebRequest = WebRequest.Create("http://domainname.com/deletebyURLPost.php")

        Dim postData = PostDataString
        Dim request As WebRequest = WebRequest.Create(UrlToPost)
        request.Method = "POST"

        Dim byteArray As Byte() = Encoding.UTF8.GetBytes(postData)
        request.ContentType = "application/x-www-form-urlencoded"
        request.ContentLength = byteArray.Length

        Dim dataStream As Stream = request.GetRequestStream()
        dataStream.Write(byteArray, 0, byteArray.Length)
        dataStream.Close()

    End Sub


    ''' <summary>
    ''' this Will capture a screenshot of the current form and it will place it inside a picturebox
    ''' </summary>
    ''' <param name="formname">This is the form name. Used to get the location on the screen, and the size of the form</param>
    ''' <param name="picturebox">This is the picturebox that will receive the screenshot</param>
    ''' <remarks></remarks>


    Public Shared Sub CaptureScreenShotForm(formname As Form, picturebox As PictureBox)
        Dim memoryImage As Bitmap

        Dim myGraphics As Graphics = formname.CreateGraphics()
        Dim s As Size = formname.Size
        memoryImage = New Bitmap(s.Width, s.Height, myGraphics)
        Dim memoryGraphics As Graphics = Graphics.FromImage(memoryImage)
        memoryGraphics.CopyFromScreen(formname.Location.X, formname.Location.Y, 0, 0, s)
        picturebox.SizeMode = PictureBoxSizeMode.StretchImage
        picturebox.Image = memoryImage
    End Sub

    ''' <summary>
    ''' This will take a screenshot of the full screen and it will save it inside the picturebox of your choice
    ''' </summary>
    ''' <param name="picturebox">This is the picturebox that will receive the screenshot</param>
    ''' <remarks></remarks>

    Public Shared Sub CaptureScreenShotAllScreen(picturebox As PictureBox)

        Dim bounds As Rectangle
        Dim screenshot As System.Drawing.Bitmap
        Dim graph As Graphics

        bounds = Screen.PrimaryScreen.Bounds
        screenshot = New System.Drawing.Bitmap(bounds.Width, bounds.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
        graph = Graphics.FromImage(screenshot)
        graph.CopyFromScreen(bounds.X, bounds.Y, 0, 0, bounds.Size, CopyPixelOperation.SourceCopy)
        picturebox.Image = screenshot

    End Sub

    ''' <summary>
    ''' This will Save the an image from a picturebox And will save it on the desired location
    ''' </summary>
    ''' <param name="Picturebox">This is the name of the picturebox control that contains the image that you want to save</param>
    ''' <param name="FilepathFilenameExt">Here you will put a string with the path filename and extention (.jpg)</param>
    ''' <remarks></remarks>
    Public Shared Sub SavePictureboxToComputer(Picturebox As PictureBox, FilepathFilenameExt As String)

        Dim bmp As New Bitmap(Picturebox.Image)
        bmp.Save(FilepathFilenameExt, System.Drawing.Imaging.ImageFormat.Jpeg)
        bmp.Dispose()

    End Sub

    ''' <summary>
    ''' This will rotate the picture in 90 degrees.
    ''' </summary>
    ''' <param name="Picturebox">This is the picturebox that contains the image that will be rotated</param>
    ''' <remarks></remarks>
    Public Shared Sub RotateImagePicturebox(Picturebox As PictureBox)

        Dim bitmap1 As Bitmap
        bitmap1 = Picturebox.Image

        If bitmap1 IsNot Nothing Then
            bitmap1.RotateFlip(RotateFlipType.Rotate90FlipNone)
            Picturebox.Image = bitmap1
        End If

    End Sub

    ''' <summary>
    ''' This will open an image from your computer and will place it inside a picturebox
    ''' This only opens jpg and png
    ''' </summary>
    ''' <param name="openfiledialog">This is the picturebox that will help you to select the file that you will open inside a picturebox</param>
    ''' <param name="picturebox">This is the picturebox where you want the image to be shown.</param>
    ''' <remarks></remarks>
    Public Shared Sub OpenImageToPicturebox(openfiledialog As OpenFileDialog, picturebox As PictureBox)

        Try
            openfiledialog.ShowDialog()
            Dim folderfileext = openfiledialog.FileName

            If folderfileext.Contains(".jpg") Or folderfileext.Contains(".JPG") Or folderfileext.Contains(".png") Or folderfileext.Contains(".PNG") Then
                picturebox.ImageLocation = folderfileext
            End If

        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>
    ''' This is a subroutine that will place an image inside a picurebox
    ''' </summary>
    ''' <param name="OpenFileDialog1">This is the picturebox that will help you to select the file that you will open inside a picturebox</param>
    ''' <param name="picturebox1">This is the picturebox where you want the image to be shown.</param>
    ''' <remarks></remarks>

    Public Shared Sub OpenImageToPicturebox2(OpenFileDialog1 As OpenFileDialog, picturebox1 As PictureBox)


        Try

            OpenFileDialog1.Filter = "Picture Files (*.jpg)|*.jpg|(*.bmp)|*.bmp|(*.png)|*.png"
            OpenFileDialog1.Title = "Select Picture Files"
            OpenFileDialog1.FileName = ""
            OpenFileDialog1.InitialDirectory = System.AppDomain.CurrentDomain.BaseDirectory()

            OpenFileDialog1.Multiselect = True
            OpenFileDialog1.FilterIndex = 1
            OpenFileDialog1.ShowDialog()

            Dim strFilePath As String
            strFilePath = OpenFileDialog1.FileName

            picturebox1.Load(strFilePath)

        Catch ex As Exception

        End Try

    End Sub




    ''' <summary>
    ''' This will upload a file into an ftp site. You have to go to your scrapping class and modify the setting
    ''' You have to modify the passwords and the folder where the file will be uploaded
    ''' </summary>
    ''' <param name="FilenameInFTP">This is the name that will have in the ftp server, you can rename it here</param>
    ''' <param name="FullPath">This is the full path filename and extention of the file you want to upload</param>
    ''' <param name="ShowResultURLTextBox">Once you upload the file, this will retrieve the url to access that file in the future</param>
    ''' <remarks></remarks>

    Public Shared Sub UploadFileFTP(FilenameInFTP As String, FullPath As String, ShowResultURLTextBox As TextBox)


        Dim username As String = "yourusername"

        Dim password As String = "yourpassword"

        Dim filenamefront As String = FilenameInFTP  '"filenameoffile.jpg" ' this is the name that will be on the server

        Dim pathfilefront As String = FullPath '"c:\cardscan\front.jpg"

        ' this is for the front part

        Dim request As System.Net.FtpWebRequest = DirectCast(System.Net.WebRequest.Create("ftp://domainname.com/public_html/cardscanner/" & filenamefront), System.Net.FtpWebRequest)
        request.Credentials = New System.Net.NetworkCredential(username, password)
        request.Method = System.Net.WebRequestMethods.Ftp.UploadFile

        Dim file() As Byte = System.IO.File.ReadAllBytes(pathfilefront)

        Dim strz As System.IO.Stream = request.GetRequestStream()
        strz.Write(file, 0, file.Length)
        strz.Close()
        strz.Dispose()

        ShowResultURLTextBox.Text = "http://domainname.com/cardscanner/" & filenamefront


    End Sub


    ''' <summary>
    ''' This is another subroutine to upload a file into your ftp. It will add a random value before the name so when you upload files they do not overlap
    ''' </summary>
    ''' <param name="pathfile">This is the path of the file that will uploaded</param>
    ''' <param name="file_name">This is the name that will have in the ftp server, you can rename it here</param>
    ''' <param name="progressbarname">This is the progressbar that will display how much of the file has been uploaded so far</param>
    ''' <param name="resulturl">This will retrieve the url where the file has been saved</param>
    ''' <param name="ButtonName">This is the button that will trigger the uploading. When uploading the button will be disbled and afterupload it will be enabled again</param>
    ''' <remarks></remarks>


    Public Shared Sub UploadFileShowProgressBar(pathfile As String, file_name As String, progressbarname As ProgressBar, resulturl As TextBox, ButtonName As Button)

        Try


            Dim username As String = "yourusername"

            Dim password As String = "yourpassword"


            Dim rnd = New Random()
            Dim nextValue = rnd.Next(999999)


            Dim pathfilefront As String = pathfile


            Dim FILENAME As String = nextValue & file_name

            Dim request As System.Net.FtpWebRequest = DirectCast(System.Net.WebRequest.Create("ftp://domain.com/public_html/123vip/" & FILENAME), System.Net.FtpWebRequest)
            request.Credentials = New System.Net.NetworkCredential(username, password)
            request.Method = System.Net.WebRequestMethods.Ftp.UploadFile


            Dim file() As Byte = System.IO.File.ReadAllBytes(pathfilefront)

            Dim strz As System.IO.Stream = request.GetRequestStream()
            strz.Write(file, 0, file.Length)



            For offset As Integer = 0 To file.Length Step 1024
                progressbarname.Value = CType(offset * progressbarname.Maximum / file.Length, Integer)
                Dim chunkSize As Integer = file.Length - offset - 1
                If chunkSize > 1024 Then chunkSize = 1024

                progressbarname.Value = progressbarname.Maximum
            Next



            strz.Close()
            strz.Dispose()



            resulturl.Text = "http://www.domain.com/123vip/" & FILENAME

            MsgBox("THE FILE HAS BEEN UPLOADED")
            progressbarname.Value = 0
            ButtonName.Text = "UPLOAD"
            ButtonName.Enabled = True




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try

    End Sub

    ''' <summary>
    ''' This will select the tab that you want to display. Useful to change between tabs when click a button or if you want to show certain information
    ''' inside another tab.
    ''' </summary>
    ''' <param name="tabcontrol">This is the tabcontrol that contains the tabpaged</param>
    ''' <param name="tabpage">This is the tabpage that you want to select or to be shown</param>
    ''' <remarks></remarks>
    Public Shared Sub SelectTabPage(tabcontrol As TabControl, tabpage As TabPage)

        tabpage.Enabled = True
        tabcontrol.SelectedTab = tabpage


    End Sub


    ''' <summary>
    ''' This will sum all the values from a column and return the result into a textbox
    ''' </summary>
    ''' <param name="DataGridView">This is the datagridview that contains the column that you want to sum.</param>
    ''' <param name="columnname">
    ''' This is the column name where the values that you want to add up are located. 
    ''' Make sure there are no empty rows and that there are no strings. all values should be numbers ( not strings)
    ''' </param>
    ''' <param name="textboxresult">The result of the addition of all the values in the column will be shown here</param>
    ''' <remarks></remarks>

    Public Shared Sub AddColumnGetTotal(DataGridView As DataGridView, columnname As String, textboxresult As TextBox)
        Try
            Dim tot As Double = 0
            Dim i As Integer = 0
            For i = 0 To DataGridView.Rows.Count - 2
                tot = tot + Convert.ToDouble(DataGridView.Rows(i).Cells(columnname).Value)
            Next i
            textboxresult.Text = tot

            FormatLikePriceTB(textboxresult)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    ''' <summary>
    '''  This get the actual cursor position and return it into a label.
    ''' Place this into a timer so whenever you move your mouse the timer will make this subroutine to display the cursor position
    ''' </summary>
    ''' <param name="PositionX">The label where you want to show the position x of the cursor in the screen</param>
    ''' <param name="positionY">The label where you want to show the position y of the cursor in the screen</param>
    ''' <remarks></remarks>
    Public Shared Sub GetCursorActualPosition(PositionX As Label, positionY As Label)
        'place this into a timer so you always get the current possition

        PositionX.Text = Cursor.Position.X
        positionY.Text = Cursor.Position.Y

    End Sub

    ''' <summary>
    ''' This will return the folder from where the application is running and will place it in a textbox
    ''' </summary>
    ''' <param name="textbox">This is the textbox that will displat the folder path from where your applications runs</param>
    ''' <remarks></remarks>
    Public Shared Sub GetFolderWhereAppRuns(textbox As TextBox)

        textbox.Text = Application.StartupPath
    End Sub

    ''' <summary>
    ''' This will read a txt file and will place the content of the txt file inside the textbox
    ''' </summary>
    ''' <param name="textbox">This is the textbox that will display the content from the txt file</param>
    ''' <param name="folderfileext">This is the full path + the name of file + extention (.txt .dll)</param>
    ''' <remarks></remarks>
    Public Shared Sub ReadTxtFiletoTextbox(textbox As TextBox, folderfileext As String)
        Try
            Dim LOCATION = folderfileext
        Dim Reader As StreamReader = File.OpenText(LOCATION)
        Dim FileText As String = Reader.ReadToEnd()
        textbox.Text = FileText
        Reader.Close()



        Catch ex As Exception
            '  MsgBox(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' When you have a login screen and you want to remember the username and password. Use this subroutine
    ''' This will create a file called login.txt and will contain the username and password.
    ''' So you can use the subroutine read username and it will place the values inside the username and password section
    ''' </summary>
    ''' <param name="cbUsername">This is a combobox because you give the option to choose the user from a combobox. This is the username combobox</param>
    ''' <param name="tbpassword">This is the password textbox. Remember to mask the chracter so it displays **** instead of the actual password</param>
    ''' <remarks></remarks>

    Public Shared Sub RECORDUSERNAME(cbUsername As ComboBox, tbpassword As TextBox)


        Dim FileToDelete As String
        FileToDelete = System.AppDomain.CurrentDomain.BaseDirectory & "login.txt"
        If System.IO.File.Exists(FileToDelete) = True Then
            System.IO.File.Delete(FileToDelete)
        End If


        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter(System.AppDomain.CurrentDomain.BaseDirectory & "login.txt", True)
        file.WriteLine("<USERNAME>" & cbUsername.Text & "</USERNAME>" & Environment.NewLine & "<PASSWORD>" & tbpassword.Text & "</PASSWORD>")
        file.Close()


    End Sub

    ''' <summary>
    ''' This will read the login.txt file and place the username and password in their respective field.
    ''' You saved the login file with recorusername subroutine.
    ''' This allows you to remember password and user information so you do not have to retype it everytime.
    ''' </summary>
    ''' <param name="cbUsername">This is the combobox for the username</param>
    ''' <param name="tbpassword">This is the password textbox</param>
    ''' <remarks></remarks>
    Public Shared Sub READUSERNAME(cbUsername As ComboBox, tbpassword As TextBox)


        Try

            Dim LOCATION = Application.StartupPath & "\login.txt"
            Dim Reader As StreamReader = File.OpenText(LOCATION)
            Dim FileText As String = Reader.ReadToEnd()
            Reader.Close()

            Dim LOGINFILE = FileText

            Dim Tags As List(Of String) = Get_HTMLTag("USERNAME", LOGINFILE)
            For Each Tag As String In Tags
                cbUsername.Text = Tag
            Next
            Dim Tags1 As List(Of String) = Get_HTMLTag("PASSWORD", LOGINFILE)
            For Each Tag As String In Tags1
                tbpassword.Text = Tag
            Next

        Catch ex As Exception
        End Try

    End Sub
    'You can give any TagName e.g. title, H1, div, head etc. etc.
    Shared Function Get_HTMLTag(ByVal TagName As String, ByVal HTML As String) As List(Of String)
        Dim lMatch As New List(Of String) 'Get the results in a List of strings

        'RegexOptions.IgnoreCase allows case mismatch e.g. if TagName="title" results can include "title", "Title", "TITLE" etc.
        'RegexOptions.Singleline allows .* to see past CarriageReturn characters 
        Dim Tag As New Regex("(?<=<" & TagName & ">).*(?=<\/" & TagName & ">)", RegexOptions.IgnoreCase Or RegexOptions.Singleline)
        For Each rMatch As Match In Tag.Matches(HTML)
            lMatch.Add(rMatch.Value)
        Next

        Return lMatch
    End Function


    ''' <summary>
    ''' This will get the exact text inside the tags with the same name open and ending tag must be the same
    ''' </summary>
    ''' <param name="Exacttag">This is the tag could be like title, div, password</param>
    ''' <param name="tbsource">This is the texbox that will be the source of the extraction</param>
    ''' <param name="TBResult">You will place the text inside the tags in this textbox</param>
    ''' <remarks></remarks>
    Public Shared Sub GetTextBetweenExactTag(Exacttag As String, tbsource As TextBox, TBResult As TextBox)
        Try
            Dim LOGINFILE = tbsource.Text

            Dim Tags1 As List(Of String) = Get_HTMLTag(Exacttag, LOGINFILE)
            For Each Tag As String In Tags1
                TBResult.Text = Tag
            Next



        Catch ex As Exception

        End Try
    End Sub

    Public Shared Sub GetTextBetweenExactTagLB(Exacttag As String, tbsource As TextBox, TBResult As Label)
        Try
            Dim LOGINFILE = tbsource.Text

            Dim Tags1 As List(Of String) = Get_HTMLTag(Exacttag, LOGINFILE)
            For Each Tag As String In Tags1
                TBResult.Text = Tag
            Next



        Catch ex As Exception

        End Try
    End Sub


    Public Shared Sub GetTextBetweenExactTagCB(Exacttag As String, tbsource As TextBox, TBResult As ComboBox)
        Try
            Dim LOGINFILE = tbsource.Text

            Dim Tags1 As List(Of String) = Get_HTMLTag(Exacttag, LOGINFILE)
            For Each Tag As String In Tags1
                TBResult.Text = Tag
            Next



        Catch ex As Exception

        End Try
    End Sub


    Public Shared Function GetTextBetweenExactTagString(Exacttag As String, tbsource As TextBox) As String
        Try
            Dim LOGINFILE = tbsource.Text

            Dim Tags1 As List(Of String) = Get_HTMLTag(Exacttag, LOGINFILE)
            For Each Tag As String In Tags1
                Return Tag
            Next



        Catch ex As Exception

        End Try
    End Function

    ''' <summary>
    ''' This will get all files from the specific directory into the combobox
    ''' </summary>
    ''' <param name="combobox">This is the combobox that will display the files</param>
    ''' <param name="tbdirectory">This is the directory where you will extract the files names</param>
    ''' <remarks></remarks>
    Public Shared Sub GetFilesFromDirectorytoCB(combobox As ComboBox, tbdirectory As TextBox)

        Try

            combobox.Items.AddRange(Directory.GetFiles(tbdirectory.Text))


        Catch ex As Exception

        End Try
    End Sub





    ''' <summary>
    ''' This will start the process mail to and will allow you to open outlook to send an email from outlook
    ''' by calling the mail to it will pass the email, subject and body to outlook new email form
    ''' </summary>
    ''' <param name="LBEmail">This is the email of the person that you want to email</param>
    ''' <param name="Subject">This is the subject of the email. </param>
    ''' <param name="Body">This is the body</param>
    ''' <remarks></remarks>

    Public Shared Sub SendEmailUsingOutlook(LBEmail As Label, Subject As TextBox, Body As TextBox)

        Process.Start("MAILTO:" & LBEmail.Text & "?subject=" & Subject.Text & "&body=" & Body.Text & "")

    End Sub


    ''' <summary>
    ''' If you want to check if a value exist inside a datagridview. Then use this subroutine
    ''' it will hightligth the row that contains the value that you are looking for (in yellow if you want another color modify this subroutine)
    ''' this is useful for searching information on datagridviews
    ''' </summary>
    ''' <param name="Textbox">This is the texbox that contains the value that you want to check</param>
    ''' <param name="datagridview">This is the datagridview where we will look to see if our string is present</param>
    ''' <remarks></remarks>
    Public Shared Sub CheckIFTextExistDGV(Textbox As TextBox, datagridview As DataGridView)

        Try

            Dim someText As String = Textbox.Text
            Dim gridRow As Integer = 0
            Dim gridColumn As Integer = 0
            For Each Row As DataGridViewRow In datagridview.Rows
                For Each column As DataGridViewColumn In datagridview.Columns
                    Dim cell As DataGridViewCell = (datagridview.Rows(gridRow).Cells(gridColumn))
                    If cell.Value.ToString.ToLower.Contains(someText.ToLower) Then
                        cell.Style.BackColor = Color.Yellow  ' te rows where the text exist will change to yellow

                        'textbox1.text = "Exist" if you want to get a confirmation that the text exist

                    End If
                    gridColumn += 1
                Next column
                gridColumn = 0
                gridRow += 1
            Next Row


        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>
    ''' This will change the column width in the datagridview selected
    ''' </summary>
    ''' <param name="datagridview">This is the datagridview that contains the column that you want to change the width</param>
    ''' <param name="ColumnNumber">This is the column number minus 1 because column count start with 0</param>
    ''' <param name="WidthValue">This is the width of the column by default you should put 200</param>
    ''' <remarks></remarks>
    Public Shared Sub ChangeWidthColumDGV(datagridview As DataGridView, ColumnNumber As Integer, WidthValue As Integer)
        'datagridview1.Columns(0).Width = 200 'Original as sample

        datagridview.Columns(ColumnNumber).Width = WidthValue

    End Sub

    ''' <summary>
    ''' This will order the the rows by ascending order so if the column selected contains numbers
    ''' The order will be like 1,2,3,4,5
    ''' </summary>
    ''' <param name="Datagridview1">This is the datagridview that contains the column that you want to sort</param>
    ''' <param name="Column">This is the column that you want to sort by ascending order</param>
    ''' <remarks></remarks>
    Public Shared Sub OrderByAscendingDGV(Datagridview1 As DataGridView, Column As Integer)

        Datagridview1.Sort(Datagridview1.Columns(Column), System.ComponentModel.ListSortDirection.Ascending)


    End Sub
    ''' <summary>
    ''' This will order the the rows by descending order so if the column selected contains numbers
    ''' The order will be like 5,4,3,2,1,0
    ''' </summary>
    ''' <param name="Datagridview1">This is the datagridview that contains the column that you want to sort</param>
    ''' <param name="Column">This is the column that you want to sort by descending order</param>
    ''' <remarks></remarks>
    Public Shared Sub OrderByDescendingDGV(Datagridview1 As DataGridView, Column As Integer)

        Datagridview1.Sort(Datagridview1.Columns(Column), System.ComponentModel.ListSortDirection.Descending)


    End Sub
    ''' <summary>
    ''' This will count the number of items inside a datagridview and display the result in a label
    ''' </summary>
    ''' <param name="Datagridview1">This is the datagridview that contains the rows that you want to count</param>
    ''' <param name="Result">This is the label where you want to show the result of how many rows you have</param>
    ''' <remarks></remarks>
    Public Shared Sub CountItemsDGVinLB(Datagridview1 As DataGridView, Result As Label)
        Result.Text = Datagridview1.Rows.Count - 1
    End Sub
    ''' <summary>
    ''' This will count the number of items inside a datagridview and display the result in a textbox
    ''' </summary>
    ''' <param name="Datagridview1">This is the datagridview that contains the rows that you want to count</param>
    ''' <param name="Result">This is the label where you want to show the result of how many rows you have</param>
    ''' <remarks></remarks>
    Public Shared Sub CountItemsDGVinTB(Datagridview1 As DataGridView, Result As TextBox)
        Result.Text = Datagridview1.Rows.Count - 1
    End Sub

    ''' <summary>
    ''' This will add padding zero to a label.Basically it will add 0 to the left until you get 6 digits total
    ''' This is useful when sorting data so it does it in the rigth order
    ''' so if you have  25  you will get 000025 which if you sort in a datagridview will give you the rigth order
    ''' </summary>
    ''' <param name="label">This is the label that contains the number to which you will add zeros to the left</param>
    ''' <remarks></remarks>
    Public Shared Sub PaddingZeroLB(label As Label)

        label.Text = label.Text.PadLeft(6, "0")

    End Sub
    ''' <summary>
    '''    This will add padding zero to a textbox.Basically it will add 0 to the left until you get 6 digits total
    ''' This is useful when sorting data so it does it in the rigth order
    ''' so if you have  25  you will get 000025 which if you sort in a datagridview will give you the rigth order
    ''' </summary>
    ''' <param name="textbox">This is the Textbox that contains the number to which you will add zeros to the left</param>
    ''' <remarks></remarks>
    Public Shared Sub PaddingZeroTB(textbox As TextBox)

        textbox.Text = textbox.Text.PadLeft(6, "0")

    End Sub


    ''' <summary>
    ''' This will delete the cookies in the webbrowser
    ''' This is so you can login again and the sites forget the information you submited previously
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub ClearCookiesWB()


        ''***********delete cokies code***********
        ''Temporary Internet Files
        'System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 8")

        ''Cookies()
        'System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 2")

        ''History()
        'System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 1")

        ''Form(Data)
        'System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 16")

        ''Passwords()
        'System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 32")

        ''Delete(All)
        'System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 255")

        'Delete All – Also delete files and settings stored by add-ons 
        System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 4351")


    End Sub


    Sub DeleteFilesTempFolder()


        For Each file As IO.FileInfo In New IO.DirectoryInfo(Path.GetTempPath()).GetFiles("*.*")
            'If (Now - file.CreationTime).Days > 1 Then
            Try
                file.Delete()
            Catch
                ' log exception or ignore '
            End Try
            'End If
        Next

    End Sub


    Public Shared Function RandomNumber(maxnum As Integer) As Integer

        Dim rn As New Random
        Dim NewRANDOMNUMBER = rn.Next(0, maxnum)
        Return NewRANDOMNUMBER

    End Function



    Public Shared Function ShortenBitly(longUri As String, login As String, apiKey As String, addHistory As Boolean) As String

        Try
            Const bitlyUrl As String = "http://api.bit.ly/shorten?longUrl={0}&apiKey={1}&login={2}&version=2.0.1&format=json&history={3}"
        Dim request = WebRequest.Create(String.Format(bitlyUrl, longUri, apiKey, login, If(addHistory, "1", "0")))
        Dim response = DirectCast(request.GetResponse(), HttpWebResponse)
        Dim bitlyResponse As String
        Using reader = New StreamReader(response.GetResponseStream())
            bitlyResponse = reader.ReadToEnd()
        End Using
        response.Close()
        If Not String.IsNullOrEmpty(bitlyResponse) Then
            Const options As RegexOptions = ((RegexOptions.IgnorePatternWhitespace Or RegexOptions.Multiline) Or RegexOptions.IgnoreCase)
            Const rx As String = """shortUrl"":\ ""(?<short>.*?)"""
            Dim reg As New Regex(rx, options)
            Dim tmp As String = reg.Match(bitlyResponse).Groups("short").Value
            Return tmp
        End If
            Return longUri


        Catch ex As Exception

            End Try
    End Function






    Public Shared Sub SAVETOEXCEL(DATAGRIDVIEWTOSAVE As DataGridView, DIRECTORYTOSAVE As TextBox, FILENAME As TextBox)



        Dim xlApp As Microsoft.Office.Interop.Excel.Application
        Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer
        Dim j As Integer

        xlApp = New Microsoft.Office.Interop.Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("Sheet1")


        For i = 0 To DATAGRIDVIEWTOSAVE.RowCount - 2
            For j = 0 To DATAGRIDVIEWTOSAVE.ColumnCount - 1
                For k As Integer = 1 To DATAGRIDVIEWTOSAVE.Columns.Count
                    xlWorkSheet.Cells(1, k) = DATAGRIDVIEWTOSAVE.Columns(k - 1).HeaderText
                    xlWorkSheet.Cells(i + 2, j + 1) = DATAGRIDVIEWTOSAVE(j, i).Value.ToString()
                Next
            Next
        Next



        Dim dir As New IO.DirectoryInfo(DIRECTORYTOSAVE.Text)
        If dir.Exists Then

            xlWorkSheet.SaveAs(DIRECTORYTOSAVE.Text & "\" & FILENAME.Text & ".xlsx")
            xlWorkBook.Close()
            xlApp.Quit()

            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheet)

            MsgBox("You can find the file" & DIRECTORYTOSAVE.Text & "\" & FILENAME.Text & ".xlsx")

        Else
            Directory.CreateDirectory(DIRECTORYTOSAVE.Text)

            xlWorkSheet.SaveAs(DIRECTORYTOSAVE.Text & "\" & FILENAME.Text & ".xlsx")
            xlWorkBook.Close()
            xlApp.Quit()

            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheet)

            MsgBox("You can find the file " & DIRECTORYTOSAVE.Text & "\" & FILENAME.Text & ".xlsx")
        End If



    End Sub


    Public Shared Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub





    Public Shared Sub SaveToCSV(DATAGRIDVIEWTOSAVE As DataGridView, DIRECTORYTOSAVE As TextBox, FILENAME As TextBox)



        Dim xlApp As Microsoft.Office.Interop.Excel.Application
        Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer
        Dim j As Integer

        xlApp = New Microsoft.Office.Interop.Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("Sheet1")


        For i = 0 To DATAGRIDVIEWTOSAVE.RowCount - 2
            For j = 0 To DATAGRIDVIEWTOSAVE.ColumnCount - 1
                For k As Integer = 1 To DATAGRIDVIEWTOSAVE.Columns.Count
                    xlWorkSheet.Cells(1, k) = DATAGRIDVIEWTOSAVE.Columns(k - 1).HeaderText
                    xlWorkSheet.Cells(i + 2, j + 1) = DATAGRIDVIEWTOSAVE(j, i).Value.ToString()
                Next
            Next
        Next



        Dim dir As New IO.DirectoryInfo(DIRECTORYTOSAVE.Text)
        If dir.Exists Then

            xlApp.DisplayAlerts = False
            xlWorkSheet.SaveAs(DIRECTORYTOSAVE.Text & "\" & FILENAME.Text & ".csv", FileFormat:=6)
            xlApp.DisplayAlerts = True

            xlWorkBook.Close(True)

            xlApp.Quit()

            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheet)

            MsgBox("You can find the file" & DIRECTORYTOSAVE.Text & "\" & FILENAME.Text & ".csv", )

        Else
            Directory.CreateDirectory(DIRECTORYTOSAVE.Text)

            xlApp.DisplayAlerts = False
            xlWorkSheet.SaveAs(DIRECTORYTOSAVE.Text & "\" & FILENAME.Text & ".csv", FileFormat:=6)
            xlApp.DisplayAlerts = True
            xlWorkBook.Close(True)

            xlApp.Quit()

            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheet)

            MsgBox("You can find the file " & DIRECTORYTOSAVE.Text & "\" & FILENAME.Text & ".csv")
        End If



    End Sub




    'this are useful subroutines that need to be placed inside the form that you are working on
    'Other wise they will not work

    '----------------------------------------------------------------------------------------------

    'this is to download a file showing the progress bar

    'Public Shared Sub DownloadFileFromInternet(ProgressBar As ProgressBar)


    '    Dim client As WebClient = New WebClient
    '    AddHandler client.DownloadProgressChanged, AddressOf client_ProgressChanged
    '    AddHandler client.DownloadFileCompleted, AddressOf client_DownloadCompleted
    '    client.DownloadFileAsync(New Uri("http://urlfile.com/filename.zip"), foldername" & filename.Text & ".zip")
    '    Button2.Text = "Download in Progress"
    '    Button2.Enabled = False



    'End Sub

    'Private Sub client_ProgressChanged(ByVal sender As Object, ByVal e As DownloadProgressChangedEventArgs)
    '    Dim bytesIn As Double = Double.Parse(e.BytesReceived.ToString())
    '    Dim totalBytes As Double = Double.Parse(e.TotalBytesToReceive.ToString())
    '    Dim percentage As Double = bytesIn / totalBytes * 100

    '    ProgressBar1.Value = Int32.Parse(Math.Truncate(percentage).ToString())
    'End Sub


    'Private Sub client_DownloadCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs)
    '    MessageBox.Show("Download Complete")
    '    Button2.Text = "Start Download"
    '    Button2.Enabled = True
    '    ProgressBar1.Value = 0
    'End Sub



    '------------------------------------------------------------------------------------------------------

    'if you open with a login section, then you hide it, this will kill it
    'Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
    '    LOGIN.Close()


    'End Sub


    '--------------------------------------------------------------------------------------------------------

    'Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing


    '    If LBLogStatus.Text = "LOG OUT" Then
    '    Else
    '        Dim choice As DialogResult = MessageBox.Show(" All unsaved progress will be lost if you close, do you want to close?", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)

    '        If choice = Windows.Forms.DialogResult.No Then
    '            e.Cancel = True
    '        Else

    '            login.Close()

    '        End If
    '    End If

    'End Sub

    '----------------------------------------------------------------------------------------------

    'Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
    '    Dim msg1
    '    msg1 = MsgBox("Quit?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Quit")
    '    If msg1 = vbNo Then e.Cancel = True

    '    If Me.Label4.Text = "ACTIVATED" Then

    '    Else
    '        If msg1 = vbYes Then Process.Start("http://josegarcia.com/redirect/trafic.php")
    '    End If


    'End Sub


    '-----------------------------------------------------------------------------------------------------


    ' This will check for the enter key press


    'Private Sub TBPassword_TextChanged(sender As Object, e As EventArgs) Handles TBPassword.KeyPress

    '    Dim tmp As System.Windows.Forms.KeyPressEventArgs = e
    '    If tmp.KeyChar = ChrW(Keys.Enter) Then

    '        LoginNow()  'this is the login subroutine that will be executed when someone hit the enter key

    '    Else

    '    End If
    'End Sub





End Class