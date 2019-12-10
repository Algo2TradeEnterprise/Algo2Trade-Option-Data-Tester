Imports System.IO
Imports HtmlAgilityPack
Imports System.Net.Http
Imports System.Threading
Imports Utilities.Network
Imports Utilities.DAL

Public Class OptionChainHelper
    Implements IDisposable

#Region "Events/Event handlers"
    Public Event DocumentDownloadComplete()
    Public Event DocumentRetryStatus(ByVal currentTry As Integer, ByVal totalTries As Integer)
    Public Event Heartbeat(ByVal msg As String)
    Public Event WaitingFor(ByVal elapsedSecs As Integer, ByVal totalSecs As Integer, ByVal msg As String)
    'The below functions are needed to allow the derived classes to raise the above two events
    Protected Overridable Sub OnDocumentDownloadComplete()
        RaiseEvent DocumentDownloadComplete()
    End Sub
    Protected Overridable Sub OnDocumentRetryStatus(ByVal currentTry As Integer, ByVal totalTries As Integer)
        RaiseEvent DocumentRetryStatus(currentTry, totalTries)
    End Sub
    Protected Overridable Sub OnHeartbeat(ByVal msg As String)
        RaiseEvent Heartbeat(msg)
    End Sub
    Protected Overridable Sub OnWaitingFor(ByVal elapsedSecs As Integer, ByVal totalSecs As Integer, ByVal msg As String)
        RaiseEvent WaitingFor(elapsedSecs, totalSecs, msg)
    End Sub
#End Region

    Private ReadOnly _cts As CancellationTokenSource
    Private ReadOnly _NSEOpenChainURL As String = "https://www.nseindia.com/live_market/dynaContent/live_watch/option_chain/optionKeys.jsp?symbolCode=0&symbol={0}&symbol={0}&instrument=OPTSTK&date=-&segmentLink=17&segmentLink=17"
    Private ReadOnly _instrumentName As String
    Private ReadOnly _filename As String
    Public Sub New(ByVal canceller As CancellationTokenSource, ByVal instrumentName As String)
        _cts = canceller
        _instrumentName = instrumentName
        _filename = Path.Combine(My.Application.Info.DirectoryPath, "Option Chain", String.Format("{0}_{1}.csv", _instrumentName, Now.ToString("yyyyMMdd")))
    End Sub

    Private Async Function GetOptionChainDataAsync() As Task(Of DataTable)
        Dim ret As DataTable = Nothing
        Dim openPositionDataURL As String = String.Format(_NSEOpenChainURL, _instrumentName)
        Dim outputResponse As HtmlDocument = Nothing
        Dim proxyToBeUsed As HttpProxy = Nothing
        Using browser As New HttpBrowser(proxyToBeUsed, Net.DecompressionMethods.GZip, New TimeSpan(0, 1, 0), _cts)
            AddHandler browser.DocumentDownloadComplete, AddressOf OnDocumentDownloadComplete
            AddHandler browser.Heartbeat, AddressOf OnHeartbeat
            AddHandler browser.WaitingFor, AddressOf OnWaitingFor
            AddHandler browser.DocumentRetryStatus, AddressOf OnDocumentRetryStatus
            'Get to the landing page first
            Dim l As Tuple(Of Uri, Object) = Await browser.NonPOSTRequestAsync(openPositionDataURL,
                                                                                HttpMethod.Get,
                                                                                Nothing,
                                                                                True,
                                                                                Nothing,
                                                                                True,
                                                                                "text/html").ConfigureAwait(False)
            If l IsNot Nothing AndAlso l.Item2 IsNot Nothing Then
                outputResponse = l.Item2
            End If
            RemoveHandler browser.DocumentDownloadComplete, AddressOf OnDocumentDownloadComplete
            RemoveHandler browser.Heartbeat, AddressOf OnHeartbeat
            RemoveHandler browser.WaitingFor, AddressOf OnWaitingFor
            RemoveHandler browser.DocumentRetryStatus, AddressOf OnDocumentRetryStatus
        End Using
        If outputResponse IsNot Nothing AndAlso outputResponse.DocumentNode IsNot Nothing Then
            OnHeartbeat("Extracting Option Chain from HTML")
            Dim calls As List(Of OptionChain) = Nothing
            Dim puts As List(Of OptionChain) = Nothing
            If outputResponse.DocumentNode.SelectNodes("//table[@id='octable']") IsNot Nothing Then
                For Each table As HtmlNode In outputResponse.DocumentNode.SelectNodes("//table[@id='octable']")
                    _cts.Token.ThrowIfCancellationRequested()
                    If table IsNot Nothing AndAlso table.SelectNodes("tr") IsNot Nothing Then
                        For Each row As HtmlNode In table.SelectNodes("tr")
                            _cts.Token.ThrowIfCancellationRequested()
                            If row IsNot Nothing AndAlso row.SelectNodes("td") IsNot Nothing Then
                                Dim callData As OptionChain = Nothing
                                Dim putData As OptionChain = Nothing
                                Dim counter As Integer = 0
                                For Each cell As HtmlNode In row.SelectNodes("td")
                                    _cts.Token.ThrowIfCancellationRequested()
                                    If cell IsNot Nothing AndAlso cell.InnerText IsNot Nothing AndAlso cell.InnerText <> "" Then
                                        If cell.InnerText.Trim = "Total" Then
                                            Exit For
                                        End If
                                        If callData Is Nothing Then callData = New OptionChain
                                        If putData Is Nothing Then putData = New OptionChain
                                        counter += 1
                                        If counter = 1 Then
                                            callData.OI = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                        ElseIf counter = 2 Then
                                            callData.ChangeInOI = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                        ElseIf counter = 3 Then
                                            callData.Volume = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                        ElseIf counter = 4 Then
                                            callData.IV = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                        ElseIf counter = 5 Then
                                            callData.LTP = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                        ElseIf counter = 6 Then
                                            callData.NetChange = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                        ElseIf counter = 7 Then
                                            callData.BidQuantity = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                        ElseIf counter = 8 Then
                                            callData.BidPrice = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                        ElseIf counter = 9 Then
                                            callData.AskPrice = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                        ElseIf counter = 10 Then
                                            callData.AskQuantity = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                        ElseIf counter = 11 Then
                                            callData.StrikePrice = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                            putData.StrikePrice = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                        ElseIf counter = 12 Then
                                            putData.BidQuantity = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                        ElseIf counter = 13 Then
                                            putData.BidPrice = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                        ElseIf counter = 14 Then
                                            putData.AskPrice = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                        ElseIf counter = 15 Then
                                            putData.AskQuantity = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                        ElseIf counter = 16 Then
                                            putData.NetChange = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                        ElseIf counter = 17 Then
                                            putData.LTP = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                        ElseIf counter = 18 Then
                                            putData.IV = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                        ElseIf counter = 19 Then
                                            putData.Volume = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                        ElseIf counter = 20 Then
                                            putData.ChangeInOI = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                        ElseIf counter = 21 Then
                                            putData.OI = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                        End If
                                    End If
                                Next
                                If callData IsNot Nothing Then
                                    If calls Is Nothing Then calls = New List(Of OptionChain)
                                    calls.Add(callData)
                                End If
                                If putData IsNot Nothing Then
                                    If puts Is Nothing Then puts = New List(Of OptionChain)
                                    puts.Add(putData)
                                End If
                            End If
                        Next
                    End If
                Next
            End If
            _cts.Token.ThrowIfCancellationRequested()
            If calls IsNot Nothing AndAlso calls.Count > 0 AndAlso puts IsNot Nothing AndAlso puts.Count > 0 AndAlso calls.Count = puts.Count Then
                _cts.Token.ThrowIfCancellationRequested()
                OnHeartbeat("Creating table of option chain data")
                ret = New DataTable
                ret.Columns.Add("Calls OI")
                ret.Columns.Add("Calls ChangeInOI")
                ret.Columns.Add("Calls Volume")
                ret.Columns.Add("Calls IV")
                ret.Columns.Add("Calls LTP")
                ret.Columns.Add("Calls NetChange")
                ret.Columns.Add("Calls BidQuantity")
                ret.Columns.Add("Calls BidPrice")
                ret.Columns.Add("Calls AskPrice")
                ret.Columns.Add("Calls AskQuantity")
                ret.Columns.Add("StrikePrice")
                ret.Columns.Add("Puts BidQuantity")
                ret.Columns.Add("Puts BidPrice")
                ret.Columns.Add("Puts AskPrice")
                ret.Columns.Add("Puts AskQuantity")
                ret.Columns.Add("Puts NetChange")
                ret.Columns.Add("Puts LTP")
                ret.Columns.Add("Puts IV")
                ret.Columns.Add("Puts Volume")
                ret.Columns.Add("Puts ChangeInOI")
                ret.Columns.Add("Puts OI")
                _cts.Token.ThrowIfCancellationRequested()
                For runningItem As Integer = 0 To calls.Count - 1
                    _cts.Token.ThrowIfCancellationRequested()
                    Dim row As DataRow = ret.NewRow
                    row("Calls OI") = If(calls(runningItem).OI <> Decimal.MinValue, calls(runningItem).OI, "-")
                    row("Calls ChangeInOI") = If(calls(runningItem).ChangeInOI <> Decimal.MinValue, calls(runningItem).ChangeInOI, "-")
                    row("Calls Volume") = If(calls(runningItem).Volume <> Decimal.MinValue, calls(runningItem).Volume, "-")
                    row("Calls IV") = If(calls(runningItem).IV <> Decimal.MinValue, calls(runningItem).IV, "-")
                    row("Calls LTP") = If(calls(runningItem).LTP <> Decimal.MinValue, calls(runningItem).LTP, "-")
                    row("Calls NetChange") = If(calls(runningItem).NetChange <> Decimal.MinValue, calls(runningItem).NetChange, "-")
                    row("Calls BidQuantity") = If(calls(runningItem).BidQuantity <> Decimal.MinValue, calls(runningItem).BidQuantity, "-")
                    row("Calls BidPrice") = If(calls(runningItem).BidPrice <> Decimal.MinValue, calls(runningItem).BidPrice, "-")
                    row("Calls AskPrice") = If(calls(runningItem).AskPrice <> Decimal.MinValue, calls(runningItem).AskPrice, "-")
                    row("Calls AskQuantity") = If(calls(runningItem).AskQuantity <> Decimal.MinValue, calls(runningItem).AskQuantity, "-")
                    row("StrikePrice") = If(calls(runningItem).StrikePrice <> Decimal.MinValue, calls(runningItem).StrikePrice, "-")
                    row("Puts BidQuantity") = If(puts(runningItem).BidQuantity <> Decimal.MinValue, puts(runningItem).BidQuantity, "-")
                    row("Puts BidPrice") = If(puts(runningItem).BidPrice <> Decimal.MinValue, puts(runningItem).BidPrice, "-")
                    row("Puts AskPrice") = If(puts(runningItem).AskPrice <> Decimal.MinValue, puts(runningItem).AskPrice, "-")
                    row("Puts AskQuantity") = If(puts(runningItem).AskQuantity <> Decimal.MinValue, puts(runningItem).AskQuantity, "-")
                    row("Puts NetChange") = If(puts(runningItem).NetChange <> Decimal.MinValue, puts(runningItem).NetChange, "-")
                    row("Puts LTP") = If(puts(runningItem).LTP <> Decimal.MinValue, puts(runningItem).LTP, "-")
                    row("Puts IV") = If(puts(runningItem).IV <> Decimal.MinValue, puts(runningItem).IV, "-")
                    row("Puts Volume") = If(puts(runningItem).Volume <> Decimal.MinValue, puts(runningItem).Volume, "-")
                    row("Puts ChangeInOI") = If(puts(runningItem).ChangeInOI <> Decimal.MinValue, puts(runningItem).ChangeInOI, "-")
                    row("Puts OI") = If(puts(runningItem).OI <> Decimal.MinValue, puts(runningItem).OI, "-")
                    ret.Rows.Add(row)
                Next
            End If
            _cts.Token.ThrowIfCancellationRequested()
        End If
        Return ret
    End Function

    Public Async Function GetDataAsync() As Task(Of List(Of OptionChain))
        Dim ret As List(Of OptionChain) = Nothing
        Dim optionChainData As DataTable = Nothing
        If File.Exists(_filename) Then
            Using csvhlpr As New CSVHelper(_filename, ",", _cts)
                optionChainData = csvhlpr.GetDataTableFromCSV(1)
            End Using
        Else
            optionChainData = Await GetOptionChainDataAsync().ConfigureAwait(False)
            If optionChainData IsNot Nothing Then
                Using csvhlpr As New CSVHelper(_filename, ",", _cts)
                    csvhlpr.GetCSVFromDataTable(optionChainData)
                End Using
            End If
        End If
        If optionChainData IsNot Nothing Then

        End If
        Return ret
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class
