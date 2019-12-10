Imports System.IO
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.DAL

Public Class frmMain

#Region "Common Delegates"
    Delegate Sub SetObjectEnableDisable_Delegate(ByVal [obj] As Object, ByVal [value] As Boolean)
    Public Sub SetObjectEnableDisable_ThreadSafe(ByVal [obj] As Object, ByVal [value] As Boolean)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [obj].InvokeRequired Then
            Dim MyDelegate As New SetObjectEnableDisable_Delegate(AddressOf SetObjectEnableDisable_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[obj], [value]})
        Else
            [obj].Enabled = [value]
        End If
    End Sub

    Delegate Sub SetLabelText_Delegate(ByVal [label] As Label, ByVal [text] As String)
    Public Sub SetLabelText_ThreadSafe(ByVal [label] As Label, ByVal [text] As String)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [label].InvokeRequired Then
            Dim MyDelegate As New SetLabelText_Delegate(AddressOf SetLabelText_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[label], [text]})
        Else
            [label].Text = [text]
        End If
    End Sub

    Delegate Function GetLabelText_Delegate(ByVal [label] As Label) As String
    Public Function GetLabelText_ThreadSafe(ByVal [label] As Label) As String
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [label].InvokeRequired Then
            Dim MyDelegate As New GetLabelText_Delegate(AddressOf GetLabelText_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[label]})
        Else
            Return [label].Text
        End If
    End Function

    Delegate Sub SetLabelTag_Delegate(ByVal [label] As Label, ByVal [tag] As String)
    Public Sub SetLabelTag_ThreadSafe(ByVal [label] As Label, ByVal [tag] As String)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [label].InvokeRequired Then
            Dim MyDelegate As New SetLabelTag_Delegate(AddressOf SetLabelTag_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[label], [tag]})
        Else
            [label].Tag = [tag]
        End If
    End Sub

    Delegate Function GetLabelTag_Delegate(ByVal [label] As Label) As String
    Public Function GetLabelTag_ThreadSafe(ByVal [label] As Label) As String
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [label].InvokeRequired Then
            Dim MyDelegate As New GetLabelTag_Delegate(AddressOf GetLabelTag_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[label]})
        Else
            Return [label].Tag
        End If
    End Function
    Delegate Sub SetToolStripLabel_Delegate(ByVal [toolStrip] As StatusStrip, ByVal [label] As ToolStripStatusLabel, ByVal [text] As String)
    Public Sub SetToolStripLabel_ThreadSafe(ByVal [toolStrip] As StatusStrip, ByVal [label] As ToolStripStatusLabel, ByVal [text] As String)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [toolStrip].InvokeRequired Then
            Dim MyDelegate As New SetToolStripLabel_Delegate(AddressOf SetToolStripLabel_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[toolStrip], [label], [text]})
        Else
            [label].Text = [text]
        End If
    End Sub

    Delegate Function GetToolStripLabel_Delegate(ByVal [toolStrip] As StatusStrip, ByVal [label] As ToolStripLabel) As String
    Public Function GetToolStripLabel_ThreadSafe(ByVal [toolStrip] As StatusStrip, ByVal [label] As ToolStripLabel) As String
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [toolStrip].InvokeRequired Then
            Dim MyDelegate As New GetToolStripLabel_Delegate(AddressOf GetToolStripLabel_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[toolStrip], [label]})
        Else
            Return [label].Text
        End If
    End Function

    Delegate Function GetDateTimePickerValue_Delegate(ByVal [dateTimePicker] As DateTimePicker) As Date
    Public Function GetDateTimePickerValue_ThreadSafe(ByVal [dateTimePicker] As DateTimePicker) As Date
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [dateTimePicker].InvokeRequired Then
            Dim MyDelegate As New GetDateTimePickerValue_Delegate(AddressOf GetDateTimePickerValue_ThreadSafe)
            Return Me.Invoke(MyDelegate, New DateTimePicker() {[dateTimePicker]})
        Else
            Return [dateTimePicker].Value
        End If
    End Function

    Delegate Function GetNumericUpDownValue_Delegate(ByVal [numericUpDown] As NumericUpDown) As Integer
    Public Function GetNumericUpDownValue_ThreadSafe(ByVal [numericUpDown] As NumericUpDown) As Integer
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [numericUpDown].InvokeRequired Then
            Dim MyDelegate As New GetNumericUpDownValue_Delegate(AddressOf GetNumericUpDownValue_ThreadSafe)
            Return Me.Invoke(MyDelegate, New NumericUpDown() {[numericUpDown]})
        Else
            Return [numericUpDown].Value
        End If
    End Function

    Delegate Function GetComboBoxIndex_Delegate(ByVal [combobox] As ComboBox) As Integer
    Public Function GetComboBoxIndex_ThreadSafe(ByVal [combobox] As ComboBox) As Integer
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [combobox].InvokeRequired Then
            Dim MyDelegate As New GetComboBoxIndex_Delegate(AddressOf GetComboBoxIndex_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[combobox]})
        Else
            Return [combobox].SelectedIndex
        End If
    End Function

    Delegate Function GetComboBoxItem_Delegate(ByVal [ComboBox] As ComboBox) As String
    Public Function GetComboBoxItem_ThreadSafe(ByVal [ComboBox] As ComboBox) As String
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [ComboBox].InvokeRequired Then
            Dim MyDelegate As New GetComboBoxItem_Delegate(AddressOf GetComboBoxItem_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[ComboBox]})
        Else
            Return [ComboBox].SelectedItem.ToString
        End If
    End Function

    Delegate Function GetTextBoxText_Delegate(ByVal [textBox] As TextBox) As String
    Public Function GetTextBoxText_ThreadSafe(ByVal [textBox] As TextBox) As String
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [textBox].InvokeRequired Then
            Dim MyDelegate As New GetTextBoxText_Delegate(AddressOf GetTextBoxText_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[textBox]})
        Else
            Return [textBox].Text
        End If
    End Function

    Delegate Function GetCheckBoxChecked_Delegate(ByVal [checkBox] As CheckBox) As Boolean
    Public Function GetCheckBoxChecked_ThreadSafe(ByVal [checkBox] As CheckBox) As Boolean
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [checkBox].InvokeRequired Then
            Dim MyDelegate As New GetCheckBoxChecked_Delegate(AddressOf GetCheckBoxChecked_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[checkBox]})
        Else
            Return [checkBox].Checked
        End If
    End Function

    Delegate Sub SetDatagridBindDatatable_Delegate(ByVal [datagrid] As DataGridView, ByVal [table] As DataTable)
    Public Sub SetDatagridBindDatatable_ThreadSafe(ByVal [datagrid] As DataGridView, ByVal [table] As DataTable)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [datagrid].InvokeRequired Then
            Dim MyDelegate As New SetDatagridBindDatatable_Delegate(AddressOf SetDatagridBindDatatable_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[datagrid], [table]})
        Else
            [datagrid].DataSource = [table]
        End If
    End Sub
#End Region

#Region "Event Handlers"
    Private Sub OnHeartbeat(message As String)
        SetLabelText_ThreadSafe(lblProgress, message)
    End Sub
    Private Sub OnDocumentDownloadComplete()
        'OnHeartbeat("Document download compelete")
    End Sub
    Private Sub OnDocumentRetryStatus(currentTry As Integer, totalTries As Integer)
        OnHeartbeat(String.Format("Try #{0}/{1}: Connecting...", currentTry, totalTries))
    End Sub
    Public Sub OnWaitingFor(ByVal elapsedSecs As Integer, ByVal totalSecs As Integer, ByVal msg As String)
        OnHeartbeat(String.Format("{0}, waiting {1}/{2} secs", msg, elapsedSecs, totalSecs))
    End Sub
#End Region

    Private _canceller As CancellationTokenSource

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetObjectEnableDisable_ThreadSafe(btnStop, False)
        txtFilePath.Text = My.Settings.FilePath
    End Sub

    Private Async Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        _canceller = New CancellationTokenSource
        My.Settings.FilePath = txtFilePath.Text
        My.Settings.Save()
        SetObjectEnableDisable_ThreadSafe(btnStart, False)
        SetObjectEnableDisable_ThreadSafe(btnStop, True)

        Await Task.Run(AddressOf StartProcessing).ConfigureAwait(False)
    End Sub
    Private Async Function StartProcessing() As Task
        Try
            Dim cmn As Common = New Common(_canceller)
            AddHandler cmn.DocumentDownloadComplete, AddressOf OnDocumentDownloadComplete
            AddHandler cmn.Heartbeat, AddressOf OnHeartbeat
            AddHandler cmn.WaitingFor, AddressOf OnWaitingFor
            AddHandler cmn.DocumentRetryStatus, AddressOf OnDocumentRetryStatus

            Dim templateFile As String = GetTextBoxText_ThreadSafe(txtFilePath)
            Dim outputFilename As String = Path.Combine(My.Application.Info.DirectoryPath, "Output", "Output File.xlsx")
            OnHeartbeat("File Copy in progress")
            If File.Exists(outputFilename) Then File.Delete(outputFilename)
            File.Copy(templateFile, outputFilename)
            Dim instrumentList As List(Of String) = cmn.GetAllStockList(Common.DataBaseTable.EOD_Futures, Now.Date.AddDays(-1))

            If instrumentList IsNot Nothing AndAlso instrumentList.Count > 0 Then
                Dim optionStocks As Dictionary(Of String, PairInstrumentDetails) = Nothing
                Dim lastTradingDate As Date = Date.MinValue
                Dim ctr As Integer = 0
                For Each runningStock In instrumentList.Take(10)
                    ctr += 1
                    SetLabelText_ThreadSafe(lblMainProgress, String.Format("Processing Option Chain Data for {0} ({1}/{2})", runningStock, ctr, instrumentList.Count))
                    _canceller.Token.ThrowIfCancellationRequested()
                    Dim stockPayload As Dictionary(Of Date, Payload) = Await cmn.GetHistoricalDataAsync(Common.DataBaseTable.EOD_Futures, runningStock, Now.Date.AddDays(-10), Now.Date).ConfigureAwait(False)
                    _canceller.Token.ThrowIfCancellationRequested()
                    If stockPayload IsNot Nothing AndAlso stockPayload.Count > 0 Then
                        Dim openPrice As Decimal = stockPayload.LastOrDefault.Value.Open
                        Dim tradingSymbol As String = stockPayload.LastOrDefault.Value.TradingSymbol
                        lastTradingDate = stockPayload.LastOrDefault.Value.PayloadDate
                        Dim strikePriceList As List(Of Decimal) = Nothing
                        Using optnChnHlpr As New OptionChainHelper(_canceller, runningStock)
                            AddHandler optnChnHlpr.DocumentDownloadComplete, AddressOf OnDocumentDownloadComplete
                            AddHandler optnChnHlpr.Heartbeat, AddressOf OnHeartbeat
                            AddHandler optnChnHlpr.WaitingFor, AddressOf OnWaitingFor
                            AddHandler optnChnHlpr.DocumentRetryStatus, AddressOf OnDocumentRetryStatus

                            strikePriceList = Await optnChnHlpr.GetStrikePriceList().ConfigureAwait(False)

                            RemoveHandler optnChnHlpr.DocumentDownloadComplete, AddressOf OnDocumentDownloadComplete
                            RemoveHandler optnChnHlpr.Heartbeat, AddressOf OnHeartbeat
                            RemoveHandler optnChnHlpr.WaitingFor, AddressOf OnWaitingFor
                            RemoveHandler optnChnHlpr.DocumentRetryStatus, AddressOf OnDocumentRetryStatus
                        End Using
                        _canceller.Token.ThrowIfCancellationRequested()
                        If strikePriceList IsNot Nothing AndAlso strikePriceList.Count > 0 Then
                            Dim peStrikePrice As Decimal = strikePriceList.FindAll(Function(x)
                                                                                       Return x <= openPrice
                                                                                   End Function).LastOrDefault
                            _canceller.Token.ThrowIfCancellationRequested()
                            Dim ceStrikePrice As Decimal = strikePriceList.FindAll(Function(x)
                                                                                       Return x > openPrice
                                                                                   End Function).FirstOrDefault

                            _canceller.Token.ThrowIfCancellationRequested()
                            Dim peStockName As String = tradingSymbol.Replace("FUT", String.Format("{0}PE", CInt(peStrikePrice)))
                            Dim ceStockName As String = tradingSymbol.Replace("FUT", String.Format("{0}CE", CInt(ceStrikePrice)))
                            Dim lotsize As Integer = cmn.GetLotSize(Common.DataBaseTable.EOD_Futures, tradingSymbol, lastTradingDate)
                            If optionStocks Is Nothing Then optionStocks = New Dictionary(Of String, PairInstrumentDetails)
                            optionStocks.Add(runningStock, New PairInstrumentDetails With {.Instrument1 = peStockName, .Instrument2 = ceStockName, .LotSize = lotsize})
                        End If
                    End If
                Next
                SetLabelText_ThreadSafe(lblMainProgress, "")
                If optionStocks IsNot Nothing AndAlso optionStocks.Count > 0 AndAlso lastTradingDate <> Date.MinValue Then
                    OnHeartbeat("Opening Excel")
                    Using excelWriter As New ExcelHelper(outputFilename, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _canceller)
                        AddHandler excelWriter.Heartbeat, AddressOf OnHeartbeat
                        AddHandler excelWriter.WaitingFor, AddressOf OnWaitingFor
                        Dim stockCounter As Integer = 0
                        For Each runningStock In optionStocks.Keys
                            stockCounter += 1
                            SetLabelText_ThreadSafe(lblMainProgress, String.Format("Processing for {0} ({1}/{2})", runningStock, stockCounter, optionStocks.Count))
                            If optionStocks(runningStock) IsNot Nothing Then
                                Dim optionPayloads As Dictionary(Of Date, PairPayload) = Nothing

                                Dim instrument1Payloads As Dictionary(Of Date, Payload) = Await cmn.GetHistoricalDataForSpecificTradingSymbolAsync(Common.DataBaseTable.Intraday_Futures, optionStocks(runningStock).Instrument1, lastTradingDate.Date, lastTradingDate.Date).ConfigureAwait(False)
                                If instrument1Payloads IsNot Nothing AndAlso instrument1Payloads.Count > 0 Then
                                    For Each runningPayload In instrument1Payloads.Values
                                        If optionPayloads Is Nothing Then optionPayloads = New Dictionary(Of Date, PairPayload)
                                        optionPayloads.Add(runningPayload.PayloadDate, New PairPayload With {.Instrument1Payload = runningPayload})
                                    Next

                                    Dim instrument2Payloads As Dictionary(Of Date, Payload) = Await cmn.GetHistoricalDataForSpecificTradingSymbolAsync(Common.DataBaseTable.Intraday_Futures, optionStocks(runningStock).Instrument2, lastTradingDate.Date, lastTradingDate.Date).ConfigureAwait(False)
                                    If instrument2Payloads IsNot Nothing AndAlso instrument2Payloads.Count > 0 Then
                                        For Each runningPayload In optionPayloads
                                            If instrument2Payloads.ContainsKey(runningPayload.Key) Then
                                                optionPayloads(runningPayload.Key).Instrument2Payload = instrument2Payloads(runningPayload.Key)
                                            End If
                                        Next
                                    End If
                                End If
                                If optionPayloads IsNot Nothing AndAlso optionPayloads.Count > 0 Then
                                    OnHeartbeat("Writing Excel")
                                    Await WriteToExcel(excelWriter, optionPayloads, GetSheetName(runningStock), optionStocks(runningStock).LotSize).ConfigureAwait(False)
                                End If
                            End If
                        Next
                    End Using
                End If
            End If
            If MessageBox.Show("Do you want to open file?", "Algo2Trade Option Data Tester", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Process.Start(outputFilename)
            End If
        Catch cex As OperationCanceledException
            MsgBox(cex.Message)
        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            OnHeartbeat("Process Complete")
            SetLabelText_ThreadSafe(lblMainProgress, "")
            SetObjectEnableDisable_ThreadSafe(btnStart, True)
            SetObjectEnableDisable_ThreadSafe(btnStop, False)
        End Try
    End Function

    Private Sub btnStop_Click(sender As Object, e As EventArgs) Handles btnStop.Click
        _canceller.Cancel()
    End Sub

    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        OpenFileDialog.ShowDialog()
    End Sub

    Private Sub OpenFileDialog_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog.FileOk
        txtFilePath.Text = OpenFileDialog.FileName
    End Sub

    Private Async Function WriteToExcel(ByVal excelWriter As ExcelHelper, ByVal dataToWrite As Dictionary(Of Date, PairPayload), ByVal sheetName As String, ByVal lotSize As Integer) As Task
        Await Task.Delay(0, _canceller.Token).ConfigureAwait(False)

        excelWriter.CopyExcelSheet("Template", sheetName)
        excelWriter.SetActiveSheet(sheetName)

        Dim instrument1StartingColumn As Integer = 0
        Dim instrument2StartingColumn As Integer = 8
        Dim rowCounter As Integer = 0
        Dim columnCounter As Integer = 0

        If dataToWrite IsNot Nothing AndAlso dataToWrite.Count > 0 Then
            Dim mainRawData(dataToWrite.Count - 1, 13) As Object
            Dim dateCtr As Integer = 0
            Dim lastInstrument1Data As Payload = Nothing
            Dim lastInstrument2Data As Payload = Nothing
            For Each runningData In dataToWrite
                dateCtr += 1
                OnHeartbeat(String.Format("Excel printing for Date: {0} [{1} of {2}]", runningData.Key.Date.ToShortDateString, dateCtr, dataToWrite.Count))

                If runningData.Value.Instrument1Payload IsNot Nothing Then
                    columnCounter = instrument1StartingColumn
                    mainRawData(rowCounter, columnCounter) = runningData.Value.Instrument1Payload.TradingSymbol
                    columnCounter += 1
                    mainRawData(rowCounter, columnCounter) = runningData.Value.Instrument1Payload.PayloadDate
                    columnCounter += 1
                    mainRawData(rowCounter, columnCounter) = runningData.Value.Instrument1Payload.Open
                    columnCounter += 1
                    mainRawData(rowCounter, columnCounter) = runningData.Value.Instrument1Payload.Low
                    columnCounter += 1
                    mainRawData(rowCounter, columnCounter) = runningData.Value.Instrument1Payload.High
                    columnCounter += 1
                    mainRawData(rowCounter, columnCounter) = runningData.Value.Instrument1Payload.Close
                    columnCounter += 1
                    mainRawData(rowCounter, columnCounter) = runningData.Value.Instrument1Payload.Volume

                    lastInstrument1Data = runningData.Value.Instrument1Payload
                Else
                    columnCounter = instrument1StartingColumn
                    mainRawData(rowCounter, columnCounter) = lastInstrument1Data.TradingSymbol
                    columnCounter += 1
                    mainRawData(rowCounter, columnCounter) = runningData.Key
                    columnCounter += 1
                    mainRawData(rowCounter, columnCounter) = ""
                    columnCounter += 1
                    mainRawData(rowCounter, columnCounter) = ""
                    columnCounter += 1
                    mainRawData(rowCounter, columnCounter) = ""
                    columnCounter += 1
                    mainRawData(rowCounter, columnCounter) = ""
                    columnCounter += 1
                    mainRawData(rowCounter, columnCounter) = ""
                End If

                If runningData.Value.Instrument2Payload IsNot Nothing Then
                    columnCounter = instrument2StartingColumn
                    mainRawData(rowCounter, columnCounter) = runningData.Value.Instrument2Payload.TradingSymbol
                    columnCounter += 1
                    mainRawData(rowCounter, columnCounter) = runningData.Value.Instrument2Payload.PayloadDate
                    columnCounter += 1
                    mainRawData(rowCounter, columnCounter) = runningData.Value.Instrument2Payload.Open
                    columnCounter += 1
                    mainRawData(rowCounter, columnCounter) = runningData.Value.Instrument2Payload.Low
                    columnCounter += 1
                    mainRawData(rowCounter, columnCounter) = runningData.Value.Instrument2Payload.High
                    columnCounter += 1
                    mainRawData(rowCounter, columnCounter) = runningData.Value.Instrument2Payload.Close
                    columnCounter += 1
                    mainRawData(rowCounter, columnCounter) = runningData.Value.Instrument2Payload.Volume

                    lastInstrument2Data = runningData.Value.Instrument2Payload
                Else
                    columnCounter = instrument2StartingColumn
                    mainRawData(rowCounter, columnCounter) = lastInstrument2Data.TradingSymbol
                    columnCounter += 1
                    mainRawData(rowCounter, columnCounter) = runningData.Key
                    columnCounter += 1
                    mainRawData(rowCounter, columnCounter) = ""
                    columnCounter += 1
                    mainRawData(rowCounter, columnCounter) = ""
                    columnCounter += 1
                    mainRawData(rowCounter, columnCounter) = ""
                    columnCounter += 1
                    mainRawData(rowCounter, columnCounter) = ""
                    columnCounter += 1
                    mainRawData(rowCounter, columnCounter) = ""
                End If
                rowCounter += 1
            Next

            Dim range As String = excelWriter.GetNamedRange(2, rowCounter - 1, 1, 15)
            excelWriter.WriteArrayToExcel(mainRawData, range)
            Erase mainRawData
            mainRawData = Nothing

            excelWriter.SetData(2, 31, lotSize, "##,##,##0", ExcelHelper.XLAlign.Right)
        End If
        excelWriter.SaveExcel()
    End Function

    Private Function GetSheetName(ByVal name As String) As String
        Dim ret As String = name
        If name.Contains(":") Then
            ret = name.Replace(":", " ")
        ElseIf name.Contains("\") Then
            ret = name.Replace("\", " ")
        ElseIf name.Contains("/") Then
            ret = name.Replace("/", " ")
        ElseIf name.Contains("?") Then
            ret = name.Replace("?", " ")
        ElseIf name.Contains("*") Then
            ret = name.Replace("*", " ")
        ElseIf name.Contains("[") Then
            ret = name.Replace("[", " ")
        ElseIf name.Contains("]") Then
            ret = name.Replace("]", " ")
        Else
            ret = name
        End If
        If ret.Length > 30 Then ret = ret.Substring(0, 25)
        Return ret
    End Function
End Class
