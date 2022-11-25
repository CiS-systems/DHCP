Imports System.Management.Automation
Imports System.Threading
Imports System.Xml

Public Class Form1
    Public Class MyParameters
        Public Shared IPName As String
        Public Shared Updatedgv4 As Integer
        Public Shared vRemoveMAC As String
    End Class

    'Spouštění 
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim xmlSettings = XDocument.Load("Settings.xml")

        'Do dgvServer se funkcí Add přidávají řádku nahoru. Proto je potřeba nejprve zjistit počet elemenů a vytvořit příslušný počet řádků. Teprve poté se naplní daty
        Me.dgv1.Rows.Clear()
        Me.dgv12.Rows.Clear()

        Dim xmlSettingsCount As XmlDocument = New XmlDocument()
        xmlSettingsCount.Load("Settings.xml")
        Dim objTest As XmlElement

        'Načteme seznam serverů pro MAC
        For i = 0 To 1000
            objTest = xmlSettingsCount.SelectSingleNode("//Settings/Servers/List/" & "Server" & i + 1)
            If Not (objTest Is Nothing) Then
                'Element v XML souboru existuje, tak přidáme řádek
                Me.dgv1.Rows.Add()
            End If
        Next i

        'Doplníme servery do tabulky
        For i = 1 To Me.dgv1.RowCount - 1
            With Me.dgv1
                .Rows.Item(i - 1).Cells(1).Value = xmlSettings.Element("Settings").Element("Servers").Element("List").Element("Server" & i).Element("ServerIP").Value
                .Rows.Item(i - 1).Cells(2).Value = xmlSettings.Element("Settings").Element("Servers").Element("List").Element("Server" & i).Element("ServerName").Value
            End With
        Next i

        'Vybereme uložené řádky
        Dim selected As String = xmlSettings.Element("Settings").Element("Servers").Element("Selected").Value
        For i = 1 To Me.dgv1.RowCount - 1
            If selected.Contains(dgv1.Rows.Item(i - 1).Cells(1).Value) Then
                dgv1.Rows.Item(i - 1).Cells(0).Value = True
            End If
        Next i

        'Načteme seznam serverů pro filtry
        For i = 0 To 1000
            objTest = xmlSettingsCount.SelectSingleNode("//Settings/Filters/List/" & "Server" & i + 1)
            If Not (objTest Is Nothing) Then
                'Element v XML souboru existuje, tak přidáme řádek
                Me.dgv12.Rows.Add()
            End If
        Next i

        'Doplníme servery do tabulky
        For i = 1 To Me.dgv12.RowCount - 1
            With Me.dgv12
                .Rows.Item(i - 1).Cells(0).Value = xmlSettings.Element("Settings").Element("Filters").Element("List").Element("Server" & i).Element("ServerIP").Value
                .Rows.Item(i - 1).Cells(1).Value = xmlSettings.Element("Settings").Element("Filters").Element("List").Element("Server" & i).Element("ScopeID").Value
            End With
        Next i

        'Zkopírujeme obsah dgv1 do dgv2, dgv5, dgv8 a dgv9 - řádek 220 (+/-)
        CopyGrid1()
        CopyGrid2()
        CopyGrid3()
        CopyGrid4()
        CopyGrid5()

        SplitContainer1.SplitterDistance = Me.Width / 2
        SplitContainer2.SplitterDistance = Me.Width / 2
        SplitContainer3.SplitterDistance = Me.Width / 3
        SplitContainer4.SplitterDistance = Me.Width / 3
        SplitContainer5.SplitterDistance = Me.Width / 3

        TabControl1.TabPages.Remove(TabPage5)
    End Sub

    'Ukládání nastavení
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim selection As String = ""
        Dim xmlSettings As XmlDocument = New XmlDocument()
        xmlSettings.Load("Settings.xml")
        Dim objTest As XmlElement

        'MAC
        'Vymažeme starý seznam serverů
        Dim nodes As XmlNodeList
        nodes = xmlSettings.SelectNodes("/Settings/Servers/List")

        For Each node As XmlNode In nodes
            If node IsNot Nothing Then
                node.RemoveAll()
            End If
        Next

        'Servery zobrazené v tabulce zapíšeme do XML
        For i = 0 To dgv1.RowCount - 2
            objTest = xmlSettings.SelectSingleNode("//Settings/Servers/List/" & "Server" & i + 1)

            'Zapíšeme nové servery - IP a jméno
            If dgv1.Rows.Item(i).Cells(1).Value IsNot Nothing Then
                With xmlSettings.SelectSingleNode("Settings/Servers/List").CreateNavigator.AppendChild
                    .WriteStartElement("Server" & i + 1)
                    .WriteElementString("ServerIP", dgv1.Rows.Item(i).Cells(1).Value)
                    .WriteElementString("ServerName", dgv1.Rows.Item(i).Cells(2).Value)
                    .WriteEndElement()
                    .Flush()
                    .Close()
                    .Dispose()
                End With
            End If

            'Zapíšeme nové vybrané servery do proměnné "selection"
            If dgv1.Rows.Item(i).Cells(1).Value IsNot Nothing And dgv1.Rows.Item(i).Cells(0).Value = True Then
                selection = dgv1.Rows.Item(i).Cells(1).Value & ", " & selection
            End If
        Next i
        xmlSettings.SelectSingleNode("Settings/Servers/Selected").InnerText() = selection


        'Filters
        'Vymažeme starý seznam serverů
        nodes = xmlSettings.SelectNodes("/Settings/Filters/List")

        For Each node As XmlNode In nodes
            If node IsNot Nothing Then
                node.RemoveAll()
            End If
        Next

        'Servery zobrazené v tabulce zapíšeme do XML
        For i = 0 To dgv12.RowCount - 2
            objTest = xmlSettings.SelectSingleNode("//Settings/Filters/List/" & "Server" & i + 1)

            'Zapíšeme nové servery - IP a jméno
            If dgv12.Rows.Item(i).Cells(1).Value IsNot Nothing Then
                With xmlSettings.SelectSingleNode("Settings/Filters/List").CreateNavigator.AppendChild
                    .WriteStartElement("Server" & i + 1)
                    .WriteElementString("ServerIP", dgv12.Rows.Item(i).Cells(0).Value)
                    .WriteElementString("ScopeID", dgv12.Rows.Item(i).Cells(1).Value)
                    .WriteEndElement()
                    .Flush()
                    .Close()
                    .Dispose()
                End With
            End If
        Next i
        xmlSettings.Save("Settings.xml") 'Ulož XML

        Progress(96)

        'Zkopírujeme obsah dgv1 do dgv2 a dgv5
        CopyGrid1()
        CopyGrid2()
        CopyGrid3()
    End Sub

    'ProgressBar
    Public Sub Progress(i As Integer)
        If ProgressBar1.Value + i > 100 Then
            ProgressBar1.Value = 100
        Else
            ProgressBar1.Value = ProgressBar1.Value + i
        End If

        If ProgressBar1.Value >= 95 Then
            ProgressBar1.Value = 100
            TimerProgress.Enabled = True
        End If
    End Sub

    'Deaktivace tlačítek a dgv
    Public Sub DeactivateComponents()
        Button1.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False

        dgv3.Enabled = False
        dgv4.Enabled = False
        dgv5.Enabled = False
    End Sub
    'Aktivace tlačítek a dgv
    Public Sub ActivateComponents()
        Button1.Enabled = True
        Button4.Enabled = True
        Button5.Enabled = True
        Button6.Enabled = True

        dgv3.Enabled = True
        dgv4.Enabled = True
        dgv5.Enabled = True
    End Sub

    'Timer
    Private Sub TimerProgress_Tick(sender As Object, e As EventArgs) Handles TimerProgress.Tick
        ProgressBar1.Value = 0
        TimerProgress.Enabled = False

        ActivateComponents()

        If MyParameters.Updatedgv4 = 1 Then
            MyParameters.Updatedgv4 = 0
            ReadMAC()
        End If
    End Sub

    'Zapsání do LOGu
    Public Sub SetErrorsLog(log As String)
        ErrorLog.AppendText(DateAndTime.Now & ": " & log & vbNewLine)
        'Progress(100)
    End Sub

    'Kopírování tabulek a nastavování jejich vlastností
    Private Sub CopyGrid1()
        'References to source and target grid.
        Dim sourceGrid As DataGridView = Me.dgv1
        Dim targetGrid As DataGridView = Me.dgv2

        'Copy all rows and cells.
        Dim targetRows = New List(Of DataGridViewRow)

        For Each sourceRow As DataGridViewRow In sourceGrid.Rows
            If (Not sourceRow.IsNewRow) Then
                Dim targetRow = CType(sourceRow.Clone(), DataGridViewRow)
                'The Clone method do not copy the cell values so we must do this manually.
                'See: https://msdn.microsoft.com/en-us/library/system.windows.forms.DataGridViewRow.clone(v=vs.110).aspx
                For Each cell As DataGridViewCell In sourceRow.Cells
                    targetRow.Cells(cell.ColumnIndex).Value = cell.Value
                Next
                targetRows.Add(targetRow)
            End If
        Next

        'Clear target columns and then clone all source columns.
        targetGrid.Columns.Clear()
        For Each column As DataGridViewColumn In sourceGrid.Columns
            targetGrid.Columns.Add(CType(column.Clone(), DataGridViewColumn))
        Next

        'It's recommended to use the AddRange method (if available)
        'when adding multiple items to a collection.
        targetGrid.Rows.AddRange(targetRows.ToArray())
    End Sub
    Private Sub CopyGrid2()
        'References to source and target grid.
        Dim sourceGrid As DataGridView = Me.dgv1
        Dim targetGrid As DataGridView = Me.dgv5

        'Copy all rows and cells.
        Dim targetRows = New List(Of DataGridViewRow)

        For Each sourceRow As DataGridViewRow In sourceGrid.Rows
            If (Not sourceRow.IsNewRow) Then
                Dim targetRow = CType(sourceRow.Clone(), DataGridViewRow)
                'The Clone method do not copy the cell values so we must do this manually.
                'See: https://msdn.microsoft.com/en-us/library/system.windows.forms.DataGridViewRow.clone(v=vs.110).aspx
                For Each cell As DataGridViewCell In sourceRow.Cells
                    targetRow.Cells(cell.ColumnIndex).Value = cell.Value
                Next
                targetRows.Add(targetRow)
            End If
        Next

        'Clear target columns and then clone all source columns.
        targetGrid.Columns.Clear()
        For Each column As DataGridViewColumn In sourceGrid.Columns
            targetGrid.Columns.Add(CType(column.Clone(), DataGridViewColumn))
        Next

        'It's recommended to use the AddRange method (if available)
        'when adding multiple items to a collection.
        targetGrid.Rows.AddRange(targetRows.ToArray())
    End Sub
    Private Sub CopyGrid3()
        'References to source and target grid.
        Dim sourceGrid As DataGridView = Me.dgv1
        Dim targetGrid As DataGridView = Me.dgv6

        'Copy all rows and cells.
        Dim targetRows = New List(Of DataGridViewRow)

        For Each sourceRow As DataGridViewRow In sourceGrid.Rows
            If (Not sourceRow.IsNewRow) Then
                Dim targetRow = CType(sourceRow.Clone(), DataGridViewRow)
                'The Clone method do not copy the cell values so we must do this manually.
                'See: https://msdn.microsoft.com/en-us/library/system.windows.forms.DataGridViewRow.clone(v=vs.110).aspx
                For Each cell As DataGridViewCell In sourceRow.Cells
                    targetRow.Cells(cell.ColumnIndex).Value = cell.Value
                Next
                targetRows.Add(targetRow)
            End If
        Next

        'Clear target columns and then clone all source columns.
        targetGrid.Columns.Clear()
        For Each column As DataGridViewColumn In sourceGrid.Columns
            targetGrid.Columns.Add(CType(column.Clone(), DataGridViewColumn))
        Next

        'It's recommended to use the AddRange method (if available)
        'when adding multiple items to a collection.
        targetGrid.Rows.AddRange(targetRows.ToArray())
    End Sub
    Private Sub CopyGrid4()
        'References to source and target grid.
        Dim sourceGrid As DataGridView = Me.dgv1
        Dim targetGrid As DataGridView = Me.dgv8

        'Copy all rows and cells.
        Dim targetRows = New List(Of DataGridViewRow)

        For Each sourceRow As DataGridViewRow In sourceGrid.Rows
            If (Not sourceRow.IsNewRow) Then
                Dim targetRow = CType(sourceRow.Clone(), DataGridViewRow)
                'The Clone method do not copy the cell values so we must do this manually.
                'See: https://msdn.microsoft.com/en-us/library/system.windows.forms.DataGridViewRow.clone(v=vs.110).aspx
                For Each cell As DataGridViewCell In sourceRow.Cells
                    targetRow.Cells(cell.ColumnIndex).Value = cell.Value
                Next
                targetRows.Add(targetRow)
            End If
        Next

        'Clear target columns and then clone all source columns.
        targetGrid.Columns.Clear()
        For Each column As DataGridViewColumn In sourceGrid.Columns
            targetGrid.Columns.Add(CType(column.Clone(), DataGridViewColumn))
        Next

        'It's recommended to use the AddRange method (if available)
        'when adding multiple items to a collection.
        targetGrid.Rows.AddRange(targetRows.ToArray())
    End Sub
    Private Sub CopyGrid5()
        'References to source and target grid.
        Dim sourceGrid As DataGridView = Me.dgv1
        Dim targetGrid As DataGridView = Me.dgv9

        'Copy all rows and cells.
        Dim targetRows = New List(Of DataGridViewRow)

        For Each sourceRow As DataGridViewRow In sourceGrid.Rows
            If (Not sourceRow.IsNewRow) Then
                Dim targetRow = CType(sourceRow.Clone(), DataGridViewRow)
                'The Clone method do not copy the cell values so we must do this manually.
                'See: https://msdn.microsoft.com/en-us/library/system.windows.forms.DataGridViewRow.clone(v=vs.110).aspx
                For Each cell As DataGridViewCell In sourceRow.Cells
                    targetRow.Cells(cell.ColumnIndex).Value = cell.Value
                Next
                targetRows.Add(targetRow)
            End If
        Next

        'Clear target columns and then clone all source columns.
        targetGrid.Columns.Clear()
        For Each column As DataGridViewColumn In sourceGrid.Columns
            targetGrid.Columns.Add(CType(column.Clone(), DataGridViewColumn))
        Next

        'It's recommended to use the AddRange method (if available)
        'when adding multiple items to a collection.
        targetGrid.Rows.AddRange(targetRows.ToArray())
    End Sub
    Private Sub dgv3_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles dgv3.RowsAdded
        For rowIndex As Integer = 0 To dgv3.Rows.Count - 1
            If dgv3.Rows(rowIndex).Cells(0).Value = "" And dgv3.Rows(rowIndex).Cells(1).Value = "" Then
                dgv3.Rows(rowIndex).Cells(2).Value = "Allow"
            End If
        Next
    End Sub
    Private Sub dgv5_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles dgv5.EditingControlShowing ', dgv5.EditingControlShowing
        dgv5.Columns(1).ReadOnly = True
        dgv5.Columns(2).ReadOnly = True
    End Sub
    Private Sub dgv2_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles dgv2.EditingControlShowing
        dgv2.Columns(1).ReadOnly = True
        dgv2.Columns(2).ReadOnly = True
    End Sub

    'Získáme seznam serverů z AD
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Button2.Enabled = False
        SetErrorsLog("Erhalte Liste des Servers aus AD")

        dgv1.Rows.Clear()
        Progress(20)

        For i = 1 To 2
            Dim thread As New Thread(AddressOf ThreadServer)
            thread.Start(i)
        Next

    End Sub
    Private Sub ThreadServer(i As Integer)

        Dim command As New PSCommand()
        Dim powershell As Management.Automation.PowerShell = PowerShell.Create()

        If i = 1 Then
            'První vlákno pro načtení IP adres
            command.AddScript(TextBox4.Text)
            powershell.Commands = command
            Dim results = powershell.Invoke()

            For Each result As PSObject In results
                If (TextBox3.InvokeRequired) Then
                    Dim args() As String = {result.ToString}
                    Me.Invoke(New Action(Of String)(AddressOf SetIP), args)
                Else
                    TextBox3.Text = result.ToString
                End If
            Next

            If powershell.HadErrors Then
                For Each errorRecord As ErrorRecord In powershell.Streams.Error
                    If (ErrorLog.InvokeRequired) Then
                        Dim args() As String = {errorRecord.Exception.Message & vbNewLine}
                        Me.Invoke(New Action(Of String)(AddressOf SetErrorsLog), args)
                    End If

                    If (ProgressBar1.InvokeRequired) Then
                        Me.Invoke(Sub() Progress(80))
                    Else
                        Progress(80)
                    End If
                Next
                MsgBox("Während der Verarbeitung ist ein Fehler aufgetreten." & Environment.NewLine & "Überprüfen Sie das Log Protokoll.", MsgBoxStyle.Critical, "Fehler")
            End If

        Else
            'Druhé vlákno pro načtení DNS jmen
            command.AddScript(TextBox2.Text)
            powershell.Commands = command
            Dim results = powershell.Invoke()

            For Each result As PSObject In results
                If (TextBox5.InvokeRequired) Then
                    Dim args() As String = {result.ToString}
                    Me.Invoke(New Action(Of String)(AddressOf SetDNS), args)
                Else
                    TextBox5.Text = result.ToString
                End If
            Next

            If powershell.HadErrors Then
                For Each errorRecord As ErrorRecord In powershell.Streams.Error
                    If (ErrorLog.InvokeRequired) Then
                        Dim args() As String = {errorRecord.Exception.Message & vbNewLine}
                        Me.Invoke(New Action(Of String)(AddressOf SetErrorsLog), args)
                    End If
                Next
                MsgBox("Während der Verarbeitung ist ein Fehler aufgetreten." & Environment.NewLine & "Überprüfen Sie das Log Protokoll.", MsgBoxStyle.Critical, "Fehler")
            End If
        End If
    End Sub
    Public Sub SetIP(log As String)
        TextBox3.Text = log
        Progress(30)
        Filldgv1()
    End Sub
    Public Sub SetDNS(log As String)
        TextBox5.Text = log
        Progress(30)
        Filldgv1()
    End Sub
    Public Sub Filldgv1()
        If TextBox3.Text <> "" And TextBox5.Text <> "" Then
            'Byla spuštěna funkce načtení serverů z AD

            'Zkopírujeme do tabulky
            For j = 0 To TextBox3.Lines.Count - 1
                dgv1.Rows.Add(False, TextBox3.Lines(j), TextBox5.Lines(j))
            Next
            'Odstraníme první tři řádky
            For i = 0 To 2
                dgv1.Rows.RemoveAt(0)
            Next
            'Odstraním poslední prázdné řádky
            For i = dgv1.Rows.Count - 2 To 0 Step -1
                If dgv1.Rows.Count > 0 Then
                    If dgv1.Rows(i).Cells(2).Value = "" Then
                        dgv1.Rows.RemoveAt(i)
                    End If
                End If
            Next
            'Odstraníme mezery
            Dim value
            For i = 0 To dgv1.Rows.Count - 1
                value = dgv1.Rows(i).Cells(1).Value
                If value IsNot Nothing Then
                    dgv1.Rows(i).Cells(1).Value = value.ToString().Replace(" ", "")
                End If
            Next
            For i = 0 To dgv1.Rows.Count - 1
                value = dgv1.Rows(i).Cells(2).Value
                If value IsNot Nothing Then
                    dgv1.Rows(i).Cells(2).Value = value.ToString().Replace(" ", "")
                End If
            Next

            SetErrorsLog("Erhalte Liste aus server: " & TextBox3.Text)
            TextBox3.Clear()
            TextBox5.Clear()

            Progress(15)

            dgv1.ClearSelection()
        End If
    End Sub

    'Načtení MAC filtru ze serveru
    Public Sub ReadMAC()
        ProgressBar1.Value = 0
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""

        If dgv5.CurrentCell IsNot Nothing Then
            If dgv5.CurrentCell.Value IsNot Nothing Then
                'Při spuštění programu ještě není vykreslená tabulka dgv5, takže výše uvedenýma
                'podmínkama čekneme, jestli existuje a obsahuje data.

                DeactivateComponents()
                dgv4.Rows.Clear()
                SetErrorsLog("MAC Adressen wurden aus Server " & dgv5.Rows(dgv5.CurrentCell.RowIndex).Cells(1).Value & " eingelesen")

                'Vybereme si pouze IP a pošleme ji se spuštěním vlákna
                MyParameters.IPName = dgv5.Rows(dgv5.CurrentCell.RowIndex).Cells(1).Value

                Progress(10)
                For i = 1 To 3
                    '1 až 3 je pro TextBoxy: čtení MAC, Description a List
                    Dim thread As New Thread(AddressOf ThreadGet)
                    thread.Start(i)
                Next
            End If
        End If
    End Sub
    Private Sub dgv5_SelectionChanged(sender As Object, e As EventArgs) Handles dgv5.SelectionChanged
        ReadMAC()
    End Sub
    Private Sub ThreadGet(i As Integer)
        Dim command As New PSCommand()
        Dim powershell As Management.Automation.PowerShell = PowerShell.Create()

        If i = 1 Then
            command.AddScript(Replace(TextBox6.Text, "xServer", MyParameters.IPName))

            powershell.Commands = command
            Dim results = powershell.Invoke()

            For Each result As PSObject In results
                If (TextBox7.InvokeRequired) Then
                    Dim args() As String = {result.ToString}
                    Me.Invoke(New Action(Of String)(AddressOf FillGetMac), args)
                Else
                    TextBox7.Text = result.ToString
                End If
            Next

            If powershell.HadErrors Then
                For Each errorRecord As ErrorRecord In powershell.Streams.Error
                    If (ErrorLog.InvokeRequired) Then
                        Dim args() As String = {errorRecord.Exception.Message & vbNewLine}
                        Me.Invoke(New Action(Of String)(AddressOf SetErrorsLog), args)
                    End If

                    If (ProgressBar1.InvokeRequired) Then
                        Me.Invoke(Sub() Progress(90))
                    Else
                        Progress(90)
                    End If
                Next
                MsgBox("Während der Verarbeitung ist ein Fehler aufgetreten." & Environment.NewLine & "Überprüfen Sie das Log Protokoll.", MsgBoxStyle.Critical, "Fehler")
            End If
        ElseIf i = 2 Then
            command.AddScript(Replace(TextBox10.Text, "xServer", MyParameters.IPName))

            powershell.Commands = command
            Dim results = powershell.Invoke()

            For Each result As PSObject In results
                If (TextBox8.InvokeRequired) Then
                    Dim args() As String = {result.ToString}
                    Me.Invoke(New Action(Of String)(AddressOf FillGetList), args)
                Else
                    TextBox8.Text = result.ToString
                End If
            Next

            If powershell.HadErrors Then
                For Each errorRecord As ErrorRecord In powershell.Streams.Error
                    If (ErrorLog.InvokeRequired) Then
                        Dim args() As String = {errorRecord.Exception.Message & vbNewLine}
                        Me.Invoke(New Action(Of String)(AddressOf SetErrorsLog), args)
                    End If

                    If (ProgressBar1.InvokeRequired) Then
                        Me.Invoke(Sub() Progress(90))
                    Else
                        Progress(90)
                    End If
                Next
                MsgBox("Während der Verarbeitung ist ein Fehler aufgetreten." & Environment.NewLine & "Überprüfen Sie das Log Protokoll.", MsgBoxStyle.Critical, "Fehler")
            End If
        Else
            command.AddScript(Replace(TextBox11.Text, "xServer", MyParameters.IPName))

            powershell.Commands = command
            Dim results = powershell.Invoke()

            For Each result As PSObject In results
                If (TextBox9.InvokeRequired) Then
                    Dim args() As String = {result.ToString}
                    Me.Invoke(New Action(Of String)(AddressOf FillGetDesc), args)
                Else
                    TextBox9.Text = result.ToString
                End If
            Next

            If powershell.HadErrors Then
                For Each errorRecord As ErrorRecord In powershell.Streams.Error
                    If (ErrorLog.InvokeRequired) Then
                        Dim args() As String = {errorRecord.Exception.Message & vbNewLine}
                        Me.Invoke(New Action(Of String)(AddressOf SetErrorsLog), args)

                        If (ProgressBar1.InvokeRequired) Then
                            Me.Invoke(Sub() Progress(90))
                        Else
                            Progress(90)
                        End If
                    End If
                Next
                MsgBox("Während der Verarbeitung ist ein Fehler aufgetreten." & Environment.NewLine & "Überprüfen Sie das Log Protokoll.", MsgBoxStyle.Critical, "Fehler")
            End If
        End If

    End Sub
    Public Sub FillGetMac(log As String)
        TextBox7.Text = log
        Progress(20)
        Filldgv4()
    End Sub
    Public Sub FillGetList(log As String)
        TextBox8.Text = log
        Progress(20)
        Filldgv4()
    End Sub
    Public Sub FillGetDesc(log As String)
        TextBox9.Text = log
        Progress(20)
        Filldgv4()
    End Sub
    Public Sub Filldgv4()
        If TextBox7.Text <> "" And TextBox8.Text <> "" And TextBox9.Text <> "" Then
            'Byla spuštěna funkce načtení filtru ze serveru
            For j = 0 To TextBox7.Lines.Count - 1
                dgv4.Rows.Add(False, TextBox7.Lines(j), TextBox8.Lines(j), TextBox9.Lines(j))
            Next

            'Odstraníme první tři řádky
            For i = 0 To 2
                dgv4.Rows.RemoveAt(0)
            Next
            'Odstraním poslední prázdné řádky
            For i = dgv4.Rows.Count - 2 To 0 Step -1
                If dgv4.Rows.Count > 0 Then
                    If dgv4.Rows(i).Cells(2).Value = "" Then
                        dgv4.Rows.RemoveAt(i)
                    End If
                End If
            Next

            TextBox7.Clear()
            TextBox8.Clear()
            TextBox9.Clear()

            Progress(25)

            dgv4.ClearSelection()
            Button4.Enabled = False
            Button5.Enabled = False
            Button6.Enabled = False
        End If
    End Sub
    Private Sub dgv4_SelectionChanged(sender As Object, e As EventArgs) Handles dgv4.SelectionChanged
        If dgv4.CurrentRow.Index > -1 Then
            Button4.Enabled = True
            Button5.Enabled = True
            Button6.Enabled = True
        Else
            Button4.Enabled = False
            Button5.Enabled = False
            Button6.Enabled = False
        End If
    End Sub

    'Přidání nových MAC adres
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'Zkontrolujeme data zadaná v tabulce a upravíme formát MAC adres
        Dim err As Integer = 0
        Dim warn As Integer = 0

        'Jsou nějaké buňky prázdné?
        For i = dgv3.Rows.Count - 2 To 0 Step -1
            If dgv3.Rows.Item(i).Cells(0).Value Is Nothing And dgv3.Rows.Item(i).Cells(1).Value Is Nothing Then
                dgv3.Rows.Item(i).Cells(0).Style.BackColor = Color.White
                dgv3.Rows.Remove(dgv3.Rows(i))
            Else
                dgv3.Rows.Item(i).Cells(0).Style.BackColor = Color.White
            End If
        Next

        'Kontrola a převod zadanýc MAC adres
        Dim str As String
        Dim mac, state, desc, server As String
        Dim sum As Integer
        For i = 0 To dgv3.Rows.Count - 2
            If dgv3.Rows.Item(i).Cells(0).Value <> "" Then
                str = dgv3.Rows.Item(i).Cells(0).Value
                If str.Length <12 Then
                    'MAC adresa nemá dostatčný počet znaků
                    dgv3.Rows.Item(i).Cells(0).Style.BackColor = Color.Red
                    err = 1
                ElseIf str.Length > 17 Then
                    'MAC adresa obsahuje příliš mnoho znaků
                    dgv3.Rows.Item(i).Cells(0).Style.BackColor = Color.Red
                    err = 1
                ElseIf str.Length = 12 Then
                    'MAC adresa obsahuje správný počet znaků, ale bez rozdělovacích znamének. Doplníme je
                    str = str.Insert(10, "-").Insert(8, "-").Insert(6, "-").Insert(4, "-").Insert(2, "-")
                    dgv3.Rows.Item(i).Cells(0).Value = str
                ElseIf str.Length = 17 Then
                    'MAC adresa obsahuje správný počet znaků
                    dgv3.Rows.Item(i).Cells(0).Style.BackColor = Color.White
                    If str.Chars(2) = "-" And str.Chars(5) = "-" And str.Chars(8) = "-" And str.Chars(11) = "-" And str.Chars(14) = "-" Then
                        'MAC adresa je ve správném formátu
                    ElseIf str.Chars(2) = ":" And str.Chars(5) = ":" And str.Chars(8) = ":" And str.Chars(11) = ":" And str.Chars(14) = ":" Then
                        'MAC adresa je rozdělená dvojtečkami, nahradíme je mínusem
                        str = str.Replace(":", "-")
                    Else
                        'MAC adresa je rozdělená nějakým jiným symbolem, tak ho nahradíme mínusem. Uživatele upozorníme podbarvením
                        warn = 1
                        dgv3.Rows.Item(i).Cells(0).Style.BackColor = Color.Yellow
                        str = str.Replace(str.Chars(2), "-").Replace(str.Chars(5), "-").Replace(str.Chars(8), "-").Replace(str.Chars(11), "-").Replace(str.Chars(14), "-")
                    End If
                    dgv3.Rows.Item(i).Cells(0).Value = str
                End If
            End If
        Next

        If warn = 1 Then
            Dim Response
            Response = MsgBox("Bei eine oder mehr MAC Adresse war dei Trennung korigiert (gelb markiert)" & Environment.NewLine & "Wollen Sie weiter fortfahren?", MsgBoxStyle.YesNo, "Warnung")
            If Response = vbNo Then
                Exit Sub
            End If
        ElseIf err = 1 Then
            MsgBox("Eine oder mehr MAC Adresse ist in falschen Format", MsgBoxStyle.Critical, "Fehler")
            Exit Sub
        Else
            'Formáty jsou v pořádku, spustíme to
            MyParameters.Updatedgv4 = 1
            DeactivateComponents()

            sum = 0
            For Each selectedItem As DataGridViewRow In dgv2.Rows
                If selectedItem.Cells(0) IsNot Nothing AndAlso selectedItem.Cells(0).Value = True Then
                    sum = sum + 1
                End If
            Next selectedItem

            sum = 90 / sum
            Progress(10)

            For j = 0 To dgv2.Rows.Count - 1
                If dgv2.Rows(j).Cells(0).Value = True Then
                    server = dgv2.Rows(j).Cells(1).Value
                    For i = 0 To dgv3.Rows.Count - 2
                        mac = dgv3.Rows.Item(i).Cells(0).Value
                        desc = dgv3.Rows.Item(i).Cells(1).Value
                        state = dgv3.Rows.Item(i).Cells(2).Value

                        'Spustíme vlákno pro každý server
                        'Jako parametr předáme server, MAC, State, Desc a hodnotu Progressu
                        Dim thread = New Thread(Sub() Me.ThreadAdd(server, mac, desc, state, sum))
                        thread.Start()

                        SetErrorsLog("Folgende Eintrag wird auf Server " & server & " hinzugefügt: " & mac & "/" & state & "/" & desc)

                        Wait(500)
                    Next
                End If
            Next
        End If
    End Sub
    Public Sub ThreadAdd(server As String, mac As String, desc As String, state As String, progressval As Integer)
        Dim command As New PSCommand()
        Dim powershell As Management.Automation.PowerShell = PowerShell.Create()

        Dim comm As String
        comm = TextBox1.Text
        comm = Replace(comm, "xServer", server)
        comm = Replace(comm, "xMAC", mac)
        comm = Replace(comm, "xStatus", state)
        comm = Replace(comm, "xDesc", desc)

        command.AddScript(comm)

        If (ProgressBar1.InvokeRequired) Then
            Me.Invoke(Sub() Progress(progressval))
        Else
            Progress(progressval)
        End If

        powershell.Commands = command
        Dim results = powershell.Invoke()

        If powershell.HadErrors Then
            For Each errorRecord As ErrorRecord In powershell.Streams.Error
                If (ErrorLog.InvokeRequired) Then
                    Dim args() As String = {errorRecord.Exception.Message & vbNewLine}
                    Me.Invoke(New Action(Of String)(AddressOf SetErrorsLog), args)

                    If (ProgressBar1.InvokeRequired) Then
                        Me.Invoke(Sub() Progress(90))
                    Else
                        Progress(90)
                    End If
                End If
            Next
            MsgBox("Während der Verarbeitung (hinzufügen) ist ein Fehler aufgetreten." & Environment.NewLine & "Überprüfen Sie das Log Protokoll.", MsgBoxStyle.Critical, "Fehler")
        End If
    End Sub

    'Smazání MAC
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        RemoveMAC.Clear()
        MyParameters.Updatedgv4 = 1

        'Načtení MAC adres(y) do TextBoxu
        For i = 0 To dgv4.Rows.Count - 2
            If dgv4.Rows.Item(i).Cells(1).Value IsNot Nothing And dgv4.Rows.Item(i).Cells(0).Value = True Then
                If RemoveMAC.Text = "" Then
                    RemoveMAC.Text = dgv4.Rows.Item(i).Cells(1).Value
                Else
                    RemoveMAC.Text = RemoveMAC.Text & """, """ & dgv4.Rows.Item(i).Cells(1).Value
                End If
            End If
        Next

        If RemoveMAC.Text = "" Then
            MsgBox("Wählen Sie zuerst mindestens eine Zeile aus", MsgBoxStyle.Exclamation, "Keine Auswahl")
        Else
            DeactivateComponents()

            Dim sum As Integer
            sum = 0
            For Each selectedItem As DataGridViewRow In dgv5.Rows
                If selectedItem.Cells(0) IsNot Nothing AndAlso selectedItem.Cells(0).Value = True Then
                    sum = sum + 1
                    SetErrorsLog("MAC-Adresse " & RemoveMAC.Text & " wurde aus Server " & selectedItem.Cells(1).Value & " gelöscht.")
                End If
            Next selectedItem

            sum = 90 / sum
            Progress(10)

            'Spustíme vlákno pro každý server
            'Jako parametr předáme jméno serveru a hodnotu progressu
            For Each selectedItem As DataGridViewRow In dgv5.Rows
                If selectedItem.Cells(0) IsNot Nothing AndAlso selectedItem.Cells(0).Value = True Then
                    'Dim thread As New Thread(AddressOf ThreadDel)
                    Dim thread = New Thread(Sub() Me.ThreadDel(selectedItem.Cells(1).Value, sum))
                    thread.Start()
                End If
            Next selectedItem
        End If
    End Sub
    Public Sub ThreadDel(adress As String, progressval As Integer)
        Dim command As New PSCommand()
        Dim powershell As Management.Automation.PowerShell = PowerShell.Create()

        Dim comm As String
        If RemoveMAC.Text = "" Then
            'Komponenta RemoveMac není sdílená mezi Formy, pomůžeme si sdílenou proměnnou
            RemoveMAC.Text = MyParameters.vRemoveMAC
        End If
        comm = TextBox14.Text
        comm = Replace(comm, "xServer", adress)
        comm = Replace(comm, "xMAC", RemoveMAC.Text)
        command.AddScript(comm)

        If (ProgressBar1.InvokeRequired) Then
            Me.Invoke(Sub() Progress(progressval))
        Else
            Progress(progressval)
        End If

        powershell.Commands = command
        Dim results = powershell.Invoke()

        If powershell.HadErrors Then
            For Each errorRecord As ErrorRecord In powershell.Streams.Error
                If (ErrorLog.InvokeRequired) Then
                    Dim args() As String = {errorRecord.Exception.Message & vbNewLine}
                    Me.Invoke(New Action(Of String)(AddressOf SetErrorsLog), args)

                    If (ProgressBar1.InvokeRequired) Then
                        Me.Invoke(Sub() Progress(90))
                    Else
                        Progress(90)
                    End If
                End If
            Next
            MsgBox("Während der Verarbeitung (löschen) ist ein Fehler aufgetreten." & Environment.NewLine & "Überprüfen Sie das Log Protokoll.", MsgBoxStyle.Critical, "Fehler")
        End If
    End Sub

    'Změna stavu Allow/Deny
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim mac, state, desc As String
        Dim checked, sum As Integer
        checked = 0

        'Je něco vybráno?
        For i = 0 To dgv4.Rows.Count - 2
            If dgv4.Rows.Item(i).Cells(1).Value IsNot Nothing And dgv4.Rows.Item(i).Cells(0).Value = True Then
                checked = 1
            End If
        Next
        If checked = 0 Then
            MsgBox("Wählen Sie zuerst mindestens eine Zeile aus", MsgBoxStyle.Exclamation, "Keine Auswahl")
        Else
            DeactivateComponents()

            sum = 0
            For Each selectedItem As DataGridViewRow In dgv5.Rows
                If selectedItem.Cells(0) IsNot Nothing AndAlso selectedItem.Cells(0).Value = True Then
                    sum = sum + 1
                End If
            Next selectedItem

            Progress(5)
            sum = (90 / 2) / sum

            'Nejprve smažeme stávající
            RemoveMAC.Clear()
            For i = 0 To dgv4.Rows.Count - 2
                If dgv4.Rows.Item(i).Cells(1).Value IsNot Nothing And dgv4.Rows.Item(i).Cells(0).Value = True Then
                    mac = dgv4.Rows.Item(i).Cells(1).Value
                    If RemoveMAC.Text = "" Then
                        RemoveMAC.Text = mac
                    Else
                        RemoveMAC.Text = RemoveMAC.Text & """, """ & mac
                    End If
                End If
            Next

            'Spustíme vlákno pro každý server
            'Jako parametr předáme server a hodnotu Progressu
            For Each selectedItem As DataGridViewRow In dgv5.Rows
                If selectedItem.Cells(0) IsNot Nothing AndAlso selectedItem.Cells(0).Value = True Then
                    Dim thread = New Thread(Sub() Me.ThreadDel(selectedItem.Cells(1).Value, sum))
                    thread.Start()
                End If
            Next selectedItem

            Wait(1000)

            'Teď přidáme nové
            For i = 0 To dgv4.Rows.Count - 2
                If dgv4.Rows.Item(i).Cells(1).Value IsNot Nothing And dgv4.Rows.Item(i).Cells(0).Value = True Then
                    mac = dgv4.Rows.Item(i).Cells(1).Value
                    desc = dgv4.Rows.Item(i).Cells(3).Value

                    'Na konci jsou mezery,které doplňuje PS. Musíme je odstranit
                    desc = desc.Trim
                    mac = mac.Trim

                    If dgv4.Rows.Item(i).Cells(2).Value = "Allow" Then
                        state = "Deny"
                    Else
                        state = "Allow"
                    End If

                    Progress(5)
                    sum = (90 / 2) / sum

                    'Spustíme vlákno pro každý server
                    'Jako parametr předáme server, MAC, Desc, State a hodnotu Progressu
                    For Each selectedItem As DataGridViewRow In dgv5.Rows
                        If selectedItem.Cells(0) IsNot Nothing AndAlso selectedItem.Cells(0).Value = True Then
                            Dim thread = New Thread(Sub() Me.ThreadAdd(selectedItem.Cells(1).Value, mac, desc, state, sum))
                            thread.Start()
                            SetErrorsLog("Auf Server " & selectedItem.Cells(1).Value & " wird Status der MAC Adresse " & mac & " auf " & state & " geändert")
                            Wait(500)
                        End If
                    Next selectedItem
                End If
            Next
            MyParameters.Updatedgv4 = 1
        End If
    End Sub

    'Další úpravy stávající MAC
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Form2.dgv4.Rows.Clear()

        Dim val1, val2, val3 As String
        Dim checked As Integer
        checked = 0
        For i = 0 To dgv4.Rows.Count - 2
            If dgv4.Rows.Item(i).Cells(1).Value IsNot Nothing And dgv4.Rows.Item(i).Cells(0).Value = True Then
                checked = 1

                'Na konci jsou mezery,které doplňuje PS. Musíme je odstranit
                val1 = dgv4.Rows.Item(i).Cells(1).Value
                val2 = dgv4.Rows.Item(i).Cells(2).Value
                val3 = dgv4.Rows.Item(i).Cells(3).Value
                val1 = val1.Trim
                val2 = val2.Trim
                val3 = val3.Trim

                Form2.dgv4.Rows.Add(New String() {val1, val2, val3, val1, val2, val3})
            End If
        Next
        If checked = 0 Then
            MsgBox("Wählen Sie zuerst mindestens eine Zeile aus", MsgBoxStyle.Exclamation, "Keine Auswahl")
        Else
            Form2.ShowDialog()
        End If
    End Sub

    'Časovač
    Private Sub Wait(ByVal interval As Integer)
        Dim stopW As New Stopwatch
        stopW.Start()
        Do While stopW.ElapsedMilliseconds < interval
            ' Allows your UI to remain responsive
            Application.DoEvents()
        Loop
        stopW.Stop()
    End Sub

    'Hledání MAC
    Public Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        dgv7.Rows.Clear()

        If TextBox12.Text = "" Then
            MsgBox("Geben Sie gesuchte Text ein", MsgBoxStyle.Critical, "Fehler")
            Exit Sub
        End If


        Dim str As String = TextBox12.Text.ToUpper
        'Kontrola a převod zadanýc MAC adres
        If str.Length < 12 Then
            'MAC adresa nemá dostatčný počet znaků
            MsgBox("Eingegebene MAC Adresses ist im falschen Format", MsgBoxStyle.Critical, "Fehler")
            Exit Sub
        ElseIf str.Length > 17 Then
            'MAC adresa obsahuje příliš mnoho znaků
            MsgBox("Eingegebene MAC Adresses ist im falschen Format", MsgBoxStyle.Critical, "Fehler")
            Exit Sub
        ElseIf str.Length = 12 Then
            'MAC adresa obsahuje správný počet znaků, ale bez rozdělovacích znamének. Doplníme je
            str = str.Insert(10, "-").Insert(8, "-").Insert(6, "-").Insert(4, "-").Insert(2, "-")
        ElseIf str.Length = 17 Then
            'MAC adresa obsahuje správný počet znaků
            If str.Chars(2) = "-" And str.Chars(5) = "-" And str.Chars(8) = "-" And str.Chars(11) = "-" And str.Chars(14) = "-" Then
                'MAC adresa je ve správném formátu
            ElseIf str.Chars(2) = ":" And str.Chars(5) = ":" And str.Chars(8) = ":" And str.Chars(11) = ":" And str.Chars(14) = ":" Then
                'MAC adresa je rozdělená dvojtečkami, nahradíme je mínusem
                str = str.Replace(":", "-")
            Else
                'MAC adresa je rozdělená nějakým jiným symbolem, tak ho nahradíme mínusem. Uživatele upozorníme podbarvením
                str = str.Replace(str.Chars(2), "-").Replace(str.Chars(5), "-").Replace(str.Chars(8), "-").Replace(str.Chars(11), "-").Replace(str.Chars(14), "-")
            End If
        End If

        'Jdeme hledat
        Dim sum As Integer = 0

        For Each selectedItem As DataGridViewRow In dgv6.Rows
            If selectedItem.Cells(0) IsNot Nothing AndAlso selectedItem.Cells(0).Value = True Then
                sum = sum + 1
            End If
        Next selectedItem

        Progress(17)
        sum = 80 / sum

        'Spustíme vlákno pro každý server
        'Jako parametr předáme IP adresu, jméno serveru, MAC a hodnotu progressu
        For Each selectedItem As DataGridViewRow In dgv6.Rows
            If selectedItem.Cells(0) IsNot Nothing AndAlso selectedItem.Cells(0).Value = True Then
                Dim thread = New Thread(Sub() Me.ThreadFindMAC(selectedItem.Cells(1).Value, selectedItem.Cells(2).Value, str, sum))
                thread.Start()
            End If
        Next selectedItem
    End Sub
    Public Sub ThreadFindMAC(ip As String, name As String, mac As String, progressval As Integer)
        Dim command As New PSCommand()
        Dim powershell As Management.Automation.PowerShell = PowerShell.Create()

        command.AddScript(Replace(TextBox16.Text, "xServer", ip))

        powershell.Commands = command
        Dim results = powershell.Invoke()

        For Each result As PSObject In results
            If result.ToString.Contains(mac) Then
                'Nalezena MAC
                Me.Invoke(Sub() FindResult(ip, name, "true"))
                Me.Invoke(Sub() SetErrorsLog("MAC Adresse " & mac & " war auf Server " & ip & "/" & name & " gefunden."))
            Else
                'Nenalezena MAC
                Me.Invoke(Sub() FindResult(ip, name, "false"))
                Me.Invoke(Sub() SetErrorsLog("MAC Adresse " & mac & " war auf Server " & ip & "/" & name & " nicht gefunden."))
            End If
        Next

        If (ProgressBar1.InvokeRequired) Then
            Me.Invoke(Sub() Progress(progressval))
        Else
            Progress(progressval)
        End If

        If powershell.HadErrors Then
            For Each errorRecord As ErrorRecord In powershell.Streams.Error
                If (ErrorLog.InvokeRequired) Then
                    Dim args() As String = {errorRecord.Exception.Message & vbNewLine}
                    Me.Invoke(New Action(Of String)(AddressOf SetErrorsLog), args)

                    If (ProgressBar1.InvokeRequired) Then
                        Me.Invoke(Sub() Progress(90))
                    Else
                        Progress(90)
                    End If
                End If
            Next
            MsgBox("Während der Verarbeitung ist ein Fehler aufgetreten." & Environment.NewLine & "Überprüfen Sie das Log Protokoll.", MsgBoxStyle.Critical, "Fehler")
        End If
    End Sub
    Public Sub FindResult(ip As String, name As String, result As String)
        dgv7.Rows.Add(ip, name, result)
    End Sub

    'Hledání desc
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        dgv7.Rows.Clear()

        If TextBox13.Text = "" Then
            MsgBox("Geben Sie gesuchte Text ein", MsgBoxStyle.Critical, "Fehler")
            Exit Sub
        End If

        'Jdeme hledat
        Dim sum As Integer = 0
        For Each selectedItem As DataGridViewRow In dgv6.Rows
            If selectedItem.Cells(0) IsNot Nothing AndAlso selectedItem.Cells(0).Value = True Then
                sum = sum + 1
            End If
        Next selectedItem

        Progress(17)
        sum = 80 / sum

        'Spustíme vlákno pro každý server
        'Jako parametr předáme IP adresu, jméno serveru, Desc a hodnotu progressu
        For Each selectedItem As DataGridViewRow In dgv6.Rows
            If selectedItem.Cells(0) IsNot Nothing AndAlso selectedItem.Cells(0).Value = True Then
                Dim thread = New Thread(Sub() Me.ThreadFindDesc(selectedItem.Cells(1).Value, selectedItem.Cells(2).Value, "*" & TextBox13.Text & "*", sum))
                thread.Start()
            End If
        Next selectedItem
    End Sub
    Public Sub ThreadFindDesc(ip As String, name As String, desc As String, progressval As Integer)
        Dim command As New PSCommand()
        Dim powershell As Management.Automation.PowerShell = PowerShell.Create()

        command.AddScript(Replace(TextBox16.Text, "xServer", ip))

        powershell.Commands = command
        Dim results = powershell.Invoke()

        For Each result As PSObject In results
            If result.ToString.ToUpper Like desc.ToUpper Then
                'Nalezena MAC
                Me.Invoke(Sub() FindResult(ip, name, "true"))
                Me.Invoke(Sub() SetErrorsLog("MAC Adresse mit Beschreibung " & desc & " war auf Server " & ip & "/" & name & " gefunden."))
            Else
                'Nenalezena MAC
                Me.Invoke(Sub() FindResult(ip, name, "false"))
                Me.Invoke(Sub() SetErrorsLog("MAC Adresse mit Beschreibung " & desc & " war auf Server " & ip & "/" & name & " nicht gefunden."))
            End If
        Next

        If (ProgressBar1.InvokeRequired) Then
            Me.Invoke(Sub() Progress(progressval))
        Else
            Progress(progressval)
        End If

        If powershell.HadErrors Then
            For Each errorRecord As ErrorRecord In powershell.Streams.Error
                If (ErrorLog.InvokeRequired) Then
                    Dim args() As String = {errorRecord.Exception.Message & vbNewLine}
                    Me.Invoke(New Action(Of String)(AddressOf SetErrorsLog), args)

                    If (ProgressBar1.InvokeRequired) Then
                        Me.Invoke(Sub() Progress(90))
                    Else
                        Progress(90)
                    End If
                End If
            Next
            MsgBox("Während der Verarbeitung ist ein Fehler aufgetreten." & Environment.NewLine & "Überprüfen Sie das Log Protokoll.", MsgBoxStyle.Critical, "Fehler")
        End If
    End Sub
    Private Sub dgv7_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv7.CellContentClick
        Dim adr As String = dgv7.Rows(dgv7.CurrentCell.RowIndex).Cells(0).Value
        TabControl1.SelectedTab = TabPage2
        For i = 0 To dgv5.Rows.Count - 1
            If dgv5.Rows(i).Cells(1).Value = adr Then
                dgv5.ClearSelection()
                dgv5.Rows(i).Selected = True
                dgv5.CurrentCell = dgv5.Rows(i).Cells(1)
                ReadMAC()
            End If
        Next
    End Sub
    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged
        AcceptButton = Button8
    End Sub
    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged
        AcceptButton = Button7
    End Sub

    'LinkLabels
    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start(LinkLabel1.Text)
    End Sub
    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        System.Diagnostics.Process.Start(LinkLabel2.Text)
    End Sub
    Private Sub LinkLabel3_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        System.Diagnostics.Process.Start(LinkLabel3.Text)
    End Sub

    'Načtení filtru ze serveru
    Private Sub TabControl3_Selected(sender As Object, e As TabControlEventArgs) Handles TabControl3.Selected
        'Výběr prvního řádku, pokud ještě nebyl filtr načten
        If TabControl3.SelectedTab Is TabPage11 And dgv10.Rows.Count = 0 Then
            If dgv9.CurrentCell IsNot Nothing Then
                If dgv9.CurrentCell.Value IsNot Nothing Then
                    dgv9.CurrentCell = dgv9.Rows(0).Cells(1)
                End If
            End If
        End If
    End Sub
    Private Sub dgv9_SelectionChanged(sender As Object, e As EventArgs) Handles dgv9.SelectionChanged
        'Změna serveru
        Wait(500)
        If (Button11.InvokeRequired) Then
            Me.Invoke(Sub() Button11.Enabled = False)
        Else
            Button11.Enabled = False
        End If
        If (dgv9.InvokeRequired) Then
            Me.Invoke(Sub() dgv9.Enabled = False)
        Else
            dgv9.Enabled = False
        End If
        'Button11.Enabled = False
        'dgv9.Enabled = False

        'Pošleme IP se spuštěním vlákna
        If (ProgressBar1.InvokeRequired) Then
            Me.Invoke(Sub() Progress(45))
        Else
            Progress(45)
        End If
        'Progress(45)
        Dim thread = New Thread(Sub() Me.ThreadGetFilter(dgv9.Rows(dgv9.CurrentCell.RowIndex).Cells(1).Value))
        thread.Start()
    End Sub
    Public Sub ThreadGetFilter(ip As String)
        'Vlákno, které přes PS načte filtr ze serveru
        Dim command As New PSCommand()
        Dim powershell As Management.Automation.PowerShell = PowerShell.Create()

        command.AddScript(Replace(TextBox21.Text, "xServer", ip))
        powershell.Commands = command

        For Each result As PSObject In powershell.Invoke()
            If (TextBox22.InvokeRequired) Then
                Dim args() As String = {result.ToString}
                Me.Invoke(New Action(Of String)(AddressOf FillGetFilter), args)
            Else
                TextBox22.Text = result.ToString
            End If
        Next

        If powershell.HadErrors Then
            For Each errorRecord As ErrorRecord In powershell.Streams.Error
                If (ErrorLog.InvokeRequired) Then
                    Dim args() As String = {errorRecord.Exception.Message & vbNewLine}
                    Me.Invoke(New Action(Of String)(AddressOf SetErrorsLog), args)
                End If

                If (ProgressBar1.InvokeRequired) Then
                    Me.Invoke(Sub() Progress(90))
                Else
                    Progress(90)
                End If
            Next
            MsgBox("Während der Verarbeitung ist ein Fehler aufgetreten." & Environment.NewLine & "Überprüfen Sie das Log Protokoll.", MsgBoxStyle.Critical, "Fehler")
        End If
    End Sub
    Public Sub FillGetFilter(log As String)
        'If log IsNot Nothing Then
        TextBox22.Text = log
        ' End If
        Dim pole As String() = TextBox22.Lines 'Načteme textbox do pole
        TextBox22.Clear()
        dgv10.Rows.Clear()

        Dim c1, c2, c3, c4, c5, c6 As String

        For i = 0 To pole.Count - 1
            If pole(i) <> "" Then
                'Obsah pole vypíšeme do proměnných c?
                'Substring odstraňuje hlavičku řádků, kterou poslal powershell
                'Např.: IPAddress   : 172.30.5.240
                If Not String.IsNullOrEmpty(pole(i)) AndAlso pole(i).Contains("IPAddress") Then
                    c1 = pole(i).Substring(pole(i).IndexOf(": ") + 2)
                Else
                    c1 = ""
                End If
                If Not String.IsNullOrEmpty(pole(i + 1)) AndAlso pole(i + 1).Contains("ClientId") Then
                    c2 = pole(i + 1).Substring(pole(i + 1).IndexOf(": ") + 2)
                Else
                    c2 = ""
                End If
                If Not String.IsNullOrEmpty(pole(i + 2)) AndAlso pole(i + 2).Contains("ScopeId") Then
                    c3 = pole(i + 2).Substring(pole(i + 2).IndexOf(": ") + 2)
                Else
                    c3 = ""
                End If
                If Not String.IsNullOrEmpty(pole(i + 3)) AndAlso pole(i + 3).Contains("Name") Then
                    c4 = pole(i + 3).Substring(pole(i + 3).IndexOf(": ") + 2)
                Else
                    c4 = ""
                End If
                If Not String.IsNullOrEmpty(pole(i + 4)) AndAlso pole(i + 4).Contains("Typ") Then
                    c5 = pole(i + 4).Substring(pole(i + 4).IndexOf(": ") + 2)
                Else
                    c5 = ""
                End If
                If Not String.IsNullOrEmpty(pole(i + 5)) AndAlso pole(i + 5).Contains("Description") Then
                    c6 = pole(i + 5).Substring(pole(i + 5).IndexOf(": ") + 2)
                Else
                    c6 = ""
                End If

                If c1 > "" And c2 > "" And c3 > "" And c4 > "" And c5 > "" And c6 > "" Then
                    dgv10.Rows.Add(False, c1, c2, c3, c4, c5, c6)
                End If


                If pole(i).Contains("IPAddress") Then
                    i = i + 5
                Else
                    i = i + 1
                End If
            End If
        Next

        If (ProgressBar1.InvokeRequired) Then
            Me.Invoke(Sub() Progress(50))
        Else
            Progress(50)
        End If

        Button11.Enabled = True
        dgv9.Enabled = True
    End Sub

    'Smazání filtru
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        'Dim RemoveFilter As String
        Dim sum1, sum2 As Integer

        sum1 = 0
        For i = 0 To dgv9.Rows.Count - 1
            If dgv9.Rows(i).Cells(0).Value = True Then
                sum1 = sum1 + 1
            End If
        Next

        sum2 = 0
        For i = 0 To dgv10.Rows.Count - 1
            If dgv10.Rows.Item(i).Cells(2).Value IsNot Nothing AndAlso dgv10.Rows.Item(i).Cells(0).Value = True Then
                sum2 = sum2 + 1
            End If
        Next

        sum1 = 90 / (sum1 * sum2)
        Progress(10)

        'Spustíme vlákno pro každý server
        'Jako parametr předáme jméno serveru, MAC adresu filtru a hodnotu progressu
        For Each selectedItem As DataGridViewRow In dgv9.Rows
            If selectedItem.Cells(0) IsNot Nothing AndAlso selectedItem.Cells(0).Value = True Then
                For i = 0 To dgv10.Rows.Count - 1
                    If dgv10.Rows.Item(i).Cells(2).Value IsNot Nothing And dgv10.Rows.Item(i).Cells(0).Value = True Then
                        Dim thread = New Thread(Sub() Me.ThreadDelFilter(selectedItem.Cells(1).Value, dgv10.Rows.Item(i).Cells(2).Value, dgv10.Rows.Item(i).Cells(3).Value, sum1))
                        thread.Start()
                        SetErrorsLog("Filter mit MAC-Adresse " & dgv10.Rows.Item(i).Cells(2).Value & " wurde aus Server " & selectedItem.Cells(1).Value & " gelöscht.")
                        Wait(500)
                    End If
                Next
            End If
        Next selectedItem
    End Sub
    Public Sub ThreadDelFilter(server As String, mac As String, scope As String, progressval As Integer)
        Dim command As New PSCommand()
        Dim powershell As Management.Automation.PowerShell = PowerShell.Create()

        Dim comm As String
        comm = TextBox17.Text
        comm = Replace(comm, "xServer", server)
        comm = Replace(comm, "xMAC", mac)
        comm = Replace(comm, "xScope", scope)
        command.AddScript(comm)

        If (ProgressBar1.InvokeRequired) Then
            Me.Invoke(Sub() Progress(progressval))
        Else
            Progress(progressval)
        End If

        powershell.Commands = command
        Dim results = powershell.Invoke()

        If powershell.HadErrors Then
            For Each errorRecord As ErrorRecord In powershell.Streams.Error
                If (ErrorLog.InvokeRequired) Then
                    Dim args() As String = {errorRecord.Exception.Message & vbNewLine}
                    Me.Invoke(New Action(Of String)(AddressOf SetErrorsLog), args)

                    If (ProgressBar1.InvokeRequired) Then
                        Me.Invoke(Sub() Progress(90))
                    Else
                        Progress(90)
                    End If
                End If
            Next
            MsgBox("Während der Verarbeitung (löschen) ist ein Fehler aufgetreten." & Environment.NewLine & "Überprüfen Sie das Log Protokoll.", MsgBoxStyle.Critical, "Fehler")
        End If

        Wait(200)
        If ProgressBar1.Value > 94 Or ProgressBar1.Value = 0 Then
            dgv9_SelectionChanged(Nothing, EventArgs.Empty)
        End If
    End Sub

    'Přidání filtru
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        'Zkontrolujeme data zadaná v tabulce
        Dim err As Integer = 0
        Dim warn As Integer = 0

        'Jsou nějaké buňky prázdné? Pokud ano, smažeme je.
        For i = dgv11.Rows.Count - 2 To 0 Step -1
            If dgv11.Rows.Item(i).Cells(1).Value Is Nothing And dgv11.Rows.Item(i).Cells(2).Value Is Nothing Then
                dgv11.Rows.Item(i).Cells(2).Style.BackColor = Color.White
                dgv11.Rows.Remove(dgv11.Rows(i))
            Else
                dgv11.Rows.Item(i).Cells(2).Style.BackColor = Color.White
            End If
        Next

        'Kontrola a převod zadanýc MAC adres
        Dim str As String
        Dim mac, scope, ip, fName, desc, server As String
        Dim sum1, sum2 As Integer
        For i = 0 To dgv11.Rows.Count - 2
            If dgv11.Rows.Item(i).Cells(2).Value <> "" Then
                str = dgv11.Rows.Item(i).Cells(2).Value
                If str.Length < 12 Then
                    'MAC adresa nemá dostatčný počet znaků
                    dgv11.Rows.Item(i).Cells(2).Style.BackColor = Color.Red
                    err = 1
                ElseIf str.Length > 17 Then
                    'MAC adresa obsahuje příliš mnoho znaků
                    dgv11.Rows.Item(i).Cells(2).Style.BackColor = Color.Red
                    err = 1
                ElseIf str.Length = 12 Then
                    'MAC adresa obsahuje správný počet znaků, ale bez rozdělovacích znamének. Doplníme je
                    str = str.Insert(10, "-").Insert(8, "-").Insert(6, "-").Insert(4, "-").Insert(2, "-")
                    dgv11.Rows.Item(i).Cells(2).Value = str
                ElseIf str.Length = 17 Then
                    'MAC adresa obsahuje správný počet znaků
                    dgv11.Rows.Item(i).Cells(0).Style.BackColor = Color.White
                    If str.Chars(2) = "-" And str.Chars(5) = "-" And str.Chars(8) = "-" And str.Chars(11) = "-" And str.Chars(14) = "-" Then
                        'MAC adresa je ve správném formátu
                    ElseIf str.Chars(2) = ":" And str.Chars(5) = ":" And str.Chars(8) = ":" And str.Chars(11) = ":" And str.Chars(14) = ":" Then
                        'MAC adresa je rozdělená dvojtečkami, nahradíme je mínusem
                        str = str.Replace(":", "-")
                    Else
                        'MAC adresa je rozdělená nějakým jiným symbolem, tak ho nahradíme mínusem. Uživatele upozorníme podbarvením
                        warn = 1
                        dgv11.Rows.Item(i).Cells(2).Style.BackColor = Color.Yellow
                        str = str.Replace(str.Chars(2), "-").Replace(str.Chars(5), "-").Replace(str.Chars(8), "-").Replace(str.Chars(11), "-").Replace(str.Chars(14), "-")
                    End If
                    dgv11.Rows.Item(i).Cells(2).Value = str
                End If
            End If
        Next

        If warn = 1 Then
            Dim Response
            Response = MsgBox("Bei eine oder mehr MAC Adresse war dei Trennung korigiert (gelb markiert)" & Environment.NewLine & "Wollen Sie weiter fortfahren?", MsgBoxStyle.YesNo, "Warnung")
            If Response = vbNo Then
                Exit Sub
            End If
        ElseIf err = 1 Then
            MsgBox("Eine oder mehr MAC Adresse ist in falschen Format", MsgBoxStyle.Critical, "Fehler")
            Exit Sub
        Else
            'Formáty jsou v pořádku, spustíme to
            sum1 = 0
            For i = 0 To dgv8.Rows.Count - 1
                If dgv8.Rows(i).Cells(0).Value = True Then
                    sum1 = sum1 + 1
                End If
            Next

            sum2 = 0
            For j = 0 To dgv11.Rows.Count - 1
                If dgv11.Rows(j).Cells(1).Value IsNot Nothing And dgv11.Rows(j).Cells(2).Value IsNot Nothing And dgv11.Rows(j).Cells(3).Value IsNot Nothing And dgv11.Rows(j).Cells(4).Value IsNot Nothing Then
                    sum2 = sum2 + 1
                End If
            Next

            sum1 = 90 / (sum1 * sum2)
            Progress(10)

            For j = 0 To dgv8.Rows.Count - 1
                If dgv8.Rows(j).Cells(0).Value = True Then
                    server = dgv8.Rows(j).Cells(1).Value
                    For i = 0 To dgv11.Rows.Count - 2
                        scope = dgv11.Rows.Item(i).Cells(0).Value
                        ip = dgv11.Rows.Item(i).Cells(1).Value
                        mac = dgv11.Rows.Item(i).Cells(2).Value
                        fName = dgv11.Rows.Item(i).Cells(3).Value
                        desc = dgv11.Rows.Item(i).Cells(4).Value

                        'Dosadíme Scope ID podle nastavení, pokud obsahuje text "Nach Einstellung"
                        If scope = "Nach Einstellung" Then
                            For k = 0 To dgv12.Rows.Count - 1
                                If dgv12.Rows(k).Cells(0).Value = server Then
                                    scope = dgv12.Rows(k).Cells(1).Value
                                End If
                            Next
                        End If

                        'Spustíme vlákno pro každý server
                        'Jako parametr předáme server, scope, ip, MAC, Desc a hodnotu Progressu
                        Dim thread = New Thread(Sub() Me.ThreadFilterAdd(server, scope, ip, mac, fName, desc, sum1))
                        thread.Start()

                        SetErrorsLog("Folgende Eintrag wird auf Server " & server & " hinzugefügt: " & scope & "/" & ip & "/" & mac & "/" & desc)

                        Wait(500)
                    Next
                End If
            Next
        End If
    End Sub
    Public Sub ThreadFilterAdd(server As String, scope As String, ip As String, mac As String, fName As String, desc As String, progressval As Integer)
        Dim command As New PSCommand()
        Dim powershell As Management.Automation.PowerShell = PowerShell.Create()

        Dim comm As String
        comm = TextBox15.Text
        comm = Replace(comm, "xServer", server)
        comm = Replace(comm, "xMAC", mac)
        comm = Replace(comm, "xScope", scope)
        comm = Replace(comm, "xIP", ip)
        comm = Replace(comm, "xName", fName)
        comm = Replace(comm, "xDesc", desc)

        command.AddScript(comm)

        If (ProgressBar1.InvokeRequired) Then
            Me.Invoke(Sub() Progress(progressval))
        Else
            Progress(progressval)
        End If

        powershell.Commands = command
        Dim results = powershell.Invoke()

        If powershell.HadErrors Then
            For Each errorRecord As ErrorRecord In powershell.Streams.Error
                If (ErrorLog.InvokeRequired) Then
                    Dim args() As String = {errorRecord.Exception.Message & vbNewLine}
                    Me.Invoke(New Action(Of String)(AddressOf SetErrorsLog), args)

                    If (ProgressBar1.InvokeRequired) Then
                        Me.Invoke(Sub() Progress(100))
                    Else
                        Progress(100)
                    End If
                End If
            Next
            MsgBox("Während der Verarbeitung (hinzufügen) ist ein Fehler aufgetreten." & Environment.NewLine & "Überprüfen Sie das Log Protokoll.", MsgBoxStyle.Critical, "Fehler")
        End If

        Wait(200)
        If ProgressBar1.Value > 94 Or ProgressBar1.Value = 0 Then
            dgv9_SelectionChanged(Nothing, EventArgs.Empty)
        End If
    End Sub
    Private Sub dgv11_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles dgv11.RowsAdded
        dgv11.Rows(dgv11.Rows.Count - 1).Cells(0).Value = "Nach Einstellung"
    End Sub
End Class