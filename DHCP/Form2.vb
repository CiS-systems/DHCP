Imports System.Threading

Public Class Form2
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        dgv4.Rows().Clear()
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim mac, state, desc As String
        Dim checked, sum As Integer
        checked = 0

        Form1.DeactivateComponents()

        sum = 0
        For Each selectedItem As DataGridViewRow In Form1.dgv5.Rows
            If selectedItem.Cells(0) IsNot Nothing AndAlso selectedItem.Cells(0).Value = True Then
                sum = sum + 1
            End If
        Next selectedItem

        Form1.Progress(5)
        sum = (90 / 2) / sum

        'Nejprve smažeme stávající
        Form1.MyParameters.vRemoveMAC = ""
        For i = 0 To dgv4.Rows.Count - 1
            mac = dgv4.Rows.Item(i).Cells(0).Value
            If Form1.MyParameters.vRemoveMAC = "" Then
                Form1.MyParameters.vRemoveMAC = mac
            Else
                Form1.MyParameters.vRemoveMAC = Form1.MyParameters.vRemoveMAC & """, """ & mac
            End If
        Next

        'Spustíme vlákno pro každý server
        'Jako parametr předáme server a hodnotu Progressu
        For Each selectedItem As DataGridViewRow In Form1.dgv5.Rows
                If selectedItem.Cells(0) IsNot Nothing AndAlso selectedItem.Cells(0).Value = True Then
                    Dim thread = New Thread(Sub() Form1.ThreadDel(selectedItem.Cells(1).Value, sum))
                thread.Start()
                Form1.SetErrorsLog("Aus Server " & selectedItem.Cells(1).Value & " werden folgende MAC Adressen gelöscht: " & Form1.MyParameters.vRemoveMAC)
            End If
            Next selectedItem

            Wait(1000)

        'Teď přidáme nové
        For i = 0 To dgv4.Rows.Count - 1
            mac = dgv4.Rows.Item(i).Cells(3).Value
            state = dgv4.Rows.Item(i).Cells(4).Value
            desc = dgv4.Rows.Item(i).Cells(5).Value

            'Na konci jsou mezery,které doplňuje PS. Musíme je odstranit
            desc = desc.Trim
            mac = mac.Trim

            'Spustíme vlákno pro každý server
            'Jako parametr předáme server, MAC, Desc, State a hodnotu Progressu
            For Each selectedItem As DataGridViewRow In Form1.dgv5.Rows
                If selectedItem.Cells(0) IsNot Nothing AndAlso selectedItem.Cells(0).Value = True Then
                    Dim thread = New Thread(Sub() Form1.ThreadAdd(selectedItem.Cells(1).Value, mac, desc, state, sum))
                    thread.Start()
                    Wait(500)
                    Form1.SetErrorsLog("Auf Server " & selectedItem.Cells(1).Value & " wird folgende Eintrag hinzugefügt: " & mac & "/" & desc & "/" & state)
                End If
            Next selectedItem
        Next
        Form1.MyParameters.Updatedgv4 = 1
        Form1.Progress(75)
        Wait(1000)
        Form1.Progress(20)
        Me.Close()
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
End Class