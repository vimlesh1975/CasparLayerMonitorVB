Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms

Public Class MainForm
    Inherits Form

    Private ReadOnly titleLabel As New Label()
    Private ReadOnly fileInfoPanel As New Panel()
    Private ReadOnly fileNameLabel As New Label()
    Private ReadOnly remainingLargeLabel As New Label()
    Private ReadOnly sourceHostLabel As New Label()
    Private ReadOnly sourceHostTextBox As New TextBox()
    Private ReadOnly listenPortLabel As New Label()
    Private ReadOnly listenPortUpDown As New NumericUpDown()
    Private ReadOnly channelLabel As New Label()
    Private ReadOnly channelUpDown As New NumericUpDown()
    Private ReadOnly layerLabel As New Label()
    Private ReadOnly layerUpDown As New NumericUpDown()
    Private ReadOnly restartButton As New Button()
    Private ReadOnly producerCaptionLabel As New Label()
    Private ReadOnly producerValueLabel As New Label()
    Private ReadOnly remainingCaptionLabel As New Label()
    Private ReadOnly remainingValueLabel As New Label()
    Private ReadOnly audioCaptionLabel As New Label()
    Private ReadOnly leftAudioMeterPanel As New Panel()
    Private ReadOnly rightAudioMeterPanel As New Panel()
    Private ReadOnly audioValueLabel As New Label()
    Private ReadOnly addressCaptionLabel As New Label()
    Private ReadOnly addressValueLabel As New Label()
    Private ReadOnly statusLabel As New Label()
    Private ReadOnly uiRefreshTimer As New Timer()

    Private listener As UdpClient
    Private listenerTask As Task
    Private currentFilePath As String = String.Empty
    Private currentProducer As String = "-"
    Private currentOscAddress As String = "-"
    Private currentTimeSeconds As Double = 0.0R
    Private currentDurationSeconds As Double = 0.0R
    Private currentPlayedFrames As Double = 0.0R
    Private currentTotalFrames As Double = 0.0R
    Private currentFps As Double = 0.0R
    Private currentLeftAudioLevel As Double = 0.0R
    Private currentRightAudioLevel As Double = 0.0R
    Private currentAudioDetails As String = "-"
    Private currentLeftAudioText As String = "-"
    Private currentRightAudioText As String = "-"
    Private pendingLeftAudioLevel As Double = 0.0R
    Private pendingRightAudioLevel As Double = 0.0R
    Private pendingLeftAudioLevelSum As Double = 0.0R
    Private pendingRightAudioLevelSum As Double = 0.0R
    Private pendingLeftAudioLevelCount As Integer = 0
    Private pendingRightAudioLevelCount As Integer = 0
    Private pendingLeftAudioText As String = "-"
    Private pendingRightAudioText As String = "-"
    Private hasPendingAudioSample As Boolean = False
    Private lastDbfsAudioAt As DateTime = DateTime.MinValue
    Private hasDbfsAudioSample As Boolean = False
    Private selectedSourceIp As String = "127.0.0.1"
    Private selectedChannel As Integer = 1
    Private selectedLayer As Integer = 1
    Private pendingUiRefresh As Boolean = False

    Public Sub New()
        Text = "CasparCG OSC Running File Monitor"
        StartPosition = FormStartPosition.CenterScreen
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False
        ClientSize = New Size(520, 420)

        leftAudioMeterPanel.Location = New Point(6, 18)
        leftAudioMeterPanel.Size = New Size(24, 360)
        leftAudioMeterPanel.BorderStyle = BorderStyle.FixedSingle
        leftAudioMeterPanel.BackColor = Color.FromArgb(18, 18, 18)
        AddHandler leftAudioMeterPanel.Paint, AddressOf LeftAudioMeterPanel_Paint

        rightAudioMeterPanel.Location = New Point(490, 18)
        rightAudioMeterPanel.Size = New Size(24, 360)
        rightAudioMeterPanel.BorderStyle = BorderStyle.FixedSingle
        rightAudioMeterPanel.BackColor = Color.FromArgb(18, 18, 18)
        AddHandler rightAudioMeterPanel.Paint, AddressOf RightAudioMeterPanel_Paint

        titleLabel.AutoSize = True
        titleLabel.Font = New Font("Segoe UI Semibold", 15.0F, FontStyle.Bold)
        titleLabel.Location = New Point(46, 18)
        titleLabel.Text = "CasparCG OSC Layer Monitor"

        fileInfoPanel.Location = New Point(46, 58)
        fileInfoPanel.Size = New Size(430, 112)
        fileInfoPanel.BorderStyle = BorderStyle.FixedSingle
        fileInfoPanel.BackColor = Color.FromArgb(245, 245, 245)

        fileNameLabel.Font = New Font("Segoe UI", 21.0F, FontStyle.Bold)
        fileNameLabel.Location = New Point(12, 10)
        fileNameLabel.Size = New Size(404, 48)
        fileNameLabel.TextAlign = ContentAlignment.MiddleLeft
        fileNameLabel.Text = "Waiting for OSC data..."

        remainingLargeLabel.Font = New Font("Segoe UI", 26.0F, FontStyle.Bold)
        remainingLargeLabel.ForeColor = Color.DarkRed
        remainingLargeLabel.Location = New Point(12, 60)
        remainingLargeLabel.Size = New Size(404, 40)
        remainingLargeLabel.TextAlign = ContentAlignment.MiddleLeft
        remainingLargeLabel.Text = "--:--:--"

        fileInfoPanel.Controls.Add(fileNameLabel)
        fileInfoPanel.Controls.Add(remainingLargeLabel)

        sourceHostLabel.AutoSize = True
        sourceHostLabel.Location = New Point(46, 188)
        sourceHostLabel.Text = "Source IP"

        sourceHostTextBox.Location = New Point(46, 208)
        sourceHostTextBox.Size = New Size(210, 27)
        sourceHostTextBox.Text = "127.0.0.1"

        listenPortLabel.AutoSize = True
        listenPortLabel.Location = New Point(46, 242)
        listenPortLabel.Text = "OSC Port"

        listenPortUpDown.Location = New Point(46, 262)
        listenPortUpDown.Maximum = 65535D
        listenPortUpDown.Value = 6250D
        listenPortUpDown.Size = New Size(120, 27)

        channelLabel.AutoSize = True
        channelLabel.Location = New Point(190, 242)
        channelLabel.Text = "Channel"

        channelUpDown.Location = New Point(190, 262)
        channelUpDown.Minimum = 1D
        channelUpDown.Value = 1D
        channelUpDown.Size = New Size(90, 27)

        layerLabel.AutoSize = True
        layerLabel.Location = New Point(304, 242)
        layerLabel.Text = "Layer"

        layerUpDown.Location = New Point(304, 262)
        layerUpDown.Minimum = 1D
        layerUpDown.Value = 1D
        layerUpDown.Size = New Size(90, 27)

        restartButton.Location = New Point(46, 314)
        restartButton.Size = New Size(348, 36)
        restartButton.Text = "Restart OSC Listener"
        AddHandler restartButton.Click, AddressOf RestartButton_Click

        producerCaptionLabel.AutoSize = True
        producerCaptionLabel.Location = New Point(50, 296)
        producerCaptionLabel.Text = "Producer:"
        producerCaptionLabel.Visible = False

        producerValueLabel.AutoSize = True
        producerValueLabel.Font = New Font("Segoe UI Semibold", 9.0F, FontStyle.Bold)
        producerValueLabel.Location = New Point(125, 296)
        producerValueLabel.Text = "-"
        producerValueLabel.Visible = False

        remainingCaptionLabel.Visible = False
        remainingValueLabel.Visible = False

        audioCaptionLabel.AutoSize = True
        audioCaptionLabel.Location = New Point(50, 322)
        audioCaptionLabel.Text = "Audio:"
        audioCaptionLabel.Visible = False

        audioValueLabel.AutoSize = True
        audioValueLabel.Font = New Font("Segoe UI Semibold", 9.0F, FontStyle.Bold)
        audioValueLabel.Location = New Point(106, 322)
        audioValueLabel.Text = "L: - dBFS   R: - dBFS"
        audioValueLabel.Visible = False

        addressCaptionLabel.AutoSize = True
        addressCaptionLabel.Location = New Point(50, 355)
        addressCaptionLabel.Text = "OSC Address:"
        addressCaptionLabel.Visible = False

        addressValueLabel.AutoSize = True
        addressValueLabel.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular)
        addressValueLabel.Location = New Point(144, 355)
        addressValueLabel.Text = "-"
        addressValueLabel.Visible = False

        statusLabel.AutoSize = True
        statusLabel.Location = New Point(50, 390)
        statusLabel.Text = "Listener not started."
        statusLabel.Visible = False

        uiRefreshTimer.Interval = 40
        AddHandler uiRefreshTimer.Tick, AddressOf UiRefreshTimer_Tick
        uiRefreshTimer.Start()

        Controls.Add(leftAudioMeterPanel)
        Controls.Add(rightAudioMeterPanel)
        Controls.Add(titleLabel)
        Controls.Add(fileInfoPanel)
        Controls.Add(sourceHostLabel)
        Controls.Add(sourceHostTextBox)
        Controls.Add(listenPortLabel)
        Controls.Add(listenPortUpDown)
        Controls.Add(channelLabel)
        Controls.Add(channelUpDown)
        Controls.Add(layerLabel)
        Controls.Add(layerUpDown)
        Controls.Add(restartButton)
        Controls.Add(producerCaptionLabel)
        Controls.Add(producerValueLabel)
        Controls.Add(remainingCaptionLabel)
        Controls.Add(remainingValueLabel)
        Controls.Add(audioCaptionLabel)
        Controls.Add(audioValueLabel)
        Controls.Add(addressCaptionLabel)
        Controls.Add(addressValueLabel)
        Controls.Add(statusLabel)

        AddHandler Shown, AddressOf MainForm_Shown
    End Sub

    Private Async Sub MainForm_Shown(sender As Object, e As EventArgs)
        Await RestartListenerAsync()
    End Sub

    Private Async Sub RestartButton_Click(sender As Object, e As EventArgs)
        Await RestartListenerAsync()
    End Sub

    Private Async Function RestartListenerAsync() As Task
        restartButton.Enabled = False
        statusLabel.Text = "Starting OSC listener..."

        Try
            StopListener()
            ResetDisplay()

            Dim port = Decimal.ToInt32(listenPortUpDown.Value)
            selectedSourceIp = sourceHostTextBox.Text.Trim()
            selectedChannel = Decimal.ToInt32(channelUpDown.Value)
            selectedLayer = Decimal.ToInt32(layerUpDown.Value)
            listener = New UdpClient(port)
            listenerTask = Task.Run(Function() ListenLoopAsync(listener))

            statusLabel.Text = String.Format("Listening for CasparCG OSC on UDP {0}.", port)
        Catch ex As Exception
            statusLabel.Text = "Listener error: " & ex.Message
        Finally
            restartButton.Enabled = True
        End Try

        Await Task.CompletedTask
    End Function

    Private Async Function ListenLoopAsync(activeListener As UdpClient) As Task
        While True
            Dim result As UdpReceiveResult

            Try
                result = Await activeListener.ReceiveAsync()
            Catch ex As ObjectDisposedException
                Exit While
            Catch ex As SocketException
                Exit While
            End Try

            Dim allowedSource = selectedSourceIp
            If allowedSource.Length > 0 AndAlso Not String.Equals(allowedSource, "0.0.0.0", StringComparison.Ordinal) Then
                If Not IsAllowedSource(result.RemoteEndPoint.Address, allowedSource) Then
                    Continue While
                End If
            End If

            ProcessOscPacket(result.Buffer)
        End While
    End Function

    Private Shared Function IsAllowedSource(remoteAddress As IPAddress, allowedSource As String) As Boolean
        Dim normalizedAllowedSource = allowedSource.Trim()
        If normalizedAllowedSource.Length = 0 Then
            Return True
        End If

        Dim normalizedRemoteAddress = NormalizeAddress(remoteAddress)

        If String.Equals(normalizedAllowedSource, "localhost", StringComparison.OrdinalIgnoreCase) Then
            Return IPAddress.IsLoopback(normalizedRemoteAddress)
        End If

        Dim parsedAllowedAddress As IPAddress = Nothing
        If IPAddress.TryParse(normalizedAllowedSource, parsedAllowedAddress) Then
            Dim normalizedAllowedAddress = NormalizeAddress(parsedAllowedAddress)

            If IPAddress.IsLoopback(normalizedAllowedAddress) Then
                Return IPAddress.IsLoopback(normalizedRemoteAddress)
            End If

            If normalizedRemoteAddress.Equals(normalizedAllowedAddress) Then
                Return True
            End If

            If IsLocalMachineAddress(normalizedAllowedAddress) Then
                Return IPAddress.IsLoopback(normalizedRemoteAddress) OrElse IsLocalMachineAddress(normalizedRemoteAddress)
            End If

            Return False
        End If

        Return String.Equals(normalizedRemoteAddress.ToString(), normalizedAllowedSource, StringComparison.OrdinalIgnoreCase)
    End Function

    Private Shared Function NormalizeAddress(address As IPAddress) As IPAddress
        If address Is Nothing Then
            Return IPAddress.None
        End If

        If address.IsIPv4MappedToIPv6 Then
            Return address.MapToIPv4()
        End If

        Return address
    End Function

    Private Shared Function IsLocalMachineAddress(address As IPAddress) As Boolean
        Dim normalizedAddress = NormalizeAddress(address)
        If IPAddress.IsLoopback(normalizedAddress) Then
            Return True
        End If

        Try
            For Each localAddress In Dns.GetHostAddresses(Dns.GetHostName())
                If NormalizeAddress(localAddress).Equals(normalizedAddress) Then
                    Return True
                End If
            Next
        Catch
            Return False
        End Try

        Return False
    End Function

    Private Sub ProcessOscPacket(buffer As Byte())
        If buffer Is Nothing OrElse buffer.Length = 0 Then
            Return
        End If

        Dim offset = 0
        Dim rootAddress = ReadOscString(buffer, offset)

        If String.Equals(rootAddress, "#bundle", StringComparison.Ordinal) Then
            offset += 8 ' skip timetag

            While offset + 4 <= buffer.Length
                Dim elementSize = ReadInt32BigEndian(buffer, offset)
                offset += 4

                If elementSize <= 0 OrElse offset + elementSize > buffer.Length Then
                    Exit While
                End If

                Dim element(elementSize - 1) As Byte
                System.Buffer.BlockCopy(buffer, offset, element, 0, elementSize)
                offset += elementSize
                ProcessOscPacket(element)
            End While

            Return
        End If

        Dim typeTag = ReadOscString(buffer, offset)
        If String.IsNullOrEmpty(rootAddress) OrElse String.IsNullOrEmpty(typeTag) Then
            Return
        End If

        Dim arguments = ReadOscArguments(buffer, offset, typeTag)
        HandleOscMessage(rootAddress, arguments)
    End Sub

    Private Sub HandleOscMessage(address As String, arguments As List(Of Object))
        Dim newPrefix = String.Format("/channel/{0}/stage/layer/{1}/foreground/", selectedChannel, selectedLayer)
        Dim legacyPrefix = String.Format("/channel/{0}/stage/layer/{1}/", selectedChannel, selectedLayer)
        Dim channelPrefix = String.Format("/channel/{0}/", selectedChannel)

        Dim updated = False

        If address.StartsWith(newPrefix, StringComparison.OrdinalIgnoreCase) OrElse
           address.StartsWith(legacyPrefix, StringComparison.OrdinalIgnoreCase) Then

            Dim layerSuffix = GetLayerSuffix(address, newPrefix, legacyPrefix)

            If IsOneOf(layerSuffix, "file/path", "foreground/file/path", "host/path") Then
                Dim pathValue = GetFirstArgumentAsString(arguments)
                currentFilePath = pathValue
                currentOscAddress = address
                updated = True
            ElseIf IsOneOf(layerSuffix, "producer", "foreground/producer", "type", "foreground/type") Then
                currentProducer = GetFirstArgumentAsString(arguments)
                currentOscAddress = address
                updated = True
            ElseIf IsOneOf(layerSuffix, "file/name", "foreground/file/name") Then
                Dim nameValue = GetFirstArgumentAsString(arguments)
                If nameValue.Length > 0 Then
                    currentFilePath = nameValue
                    currentOscAddress = address
                    updated = True
                End If
            ElseIf IsOneOf(layerSuffix, "file/time", "foreground/file/time", "time", "foreground/time") Then
                updated = TryUpdateTimeFromSeconds(arguments, address)
            ElseIf IsOneOf(layerSuffix, "file/frame", "frame") Then
                updated = TryUpdateTimeFromFrames(arguments, address)
            ElseIf IsOneOf(layerSuffix, "file/fps", "foreground/file/streams/0/fps", "host/fps") Then
                updated = TryUpdateFps(arguments, address)
            End If
        ElseIf address.StartsWith(channelPrefix, StringComparison.OrdinalIgnoreCase) Then
            updated = TryHandleMixerAudioMessage(address, arguments, channelPrefix)
        End If

        If updated Then
            RequestDisplayRefresh()
        End If
    End Sub

    Private Sub RequestDisplayRefresh()
        pendingUiRefresh = True
    End Sub

    Private Sub UiRefreshTimer_Tick(sender As Object, e As EventArgs)
        If hasPendingAudioSample Then
            If pendingLeftAudioLevelCount > 0 Then
                currentLeftAudioLevel = pendingLeftAudioLevelSum / pendingLeftAudioLevelCount
            Else
                currentLeftAudioLevel = pendingLeftAudioLevel
            End If

            If pendingRightAudioLevelCount > 0 Then
                currentRightAudioLevel = pendingRightAudioLevelSum / pendingRightAudioLevelCount
            Else
                currentRightAudioLevel = pendingRightAudioLevel
            End If

            currentLeftAudioText = pendingLeftAudioText
            currentRightAudioText = pendingRightAudioText
            currentAudioDetails = FormatCurrentAudioDetails()

            pendingLeftAudioLevel = 0.0R
            pendingRightAudioLevel = 0.0R
            pendingLeftAudioLevelSum = 0.0R
            pendingRightAudioLevelSum = 0.0R
            pendingLeftAudioLevelCount = 0
            pendingRightAudioLevelCount = 0
            hasPendingAudioSample = False
        Else
            currentLeftAudioLevel *= 0.75R
            currentRightAudioLevel *= 0.75R
        End If

        If Not pendingUiRefresh Then
            Return
        End If

        pendingUiRefresh = False
        UpdateDisplay()
    End Sub

    Private Sub UpdateDisplay()
        Dim displayName = currentFilePath.Trim()
        If displayName.Length = 0 Then
            displayName = "No file received yet"
        ElseIf displayName.Contains("/") OrElse displayName.Contains("\") Then
            displayName = Path.GetFileName(displayName.Replace("/"c, Path.DirectorySeparatorChar))
        End If

        Dim producerDisplay = currentProducer.Trim()
        If producerDisplay.Length = 0 Then
            producerDisplay = "-"
        End If

        Dim remainingDisplay = FormatRemainingTime()
        Dim audioDetails = currentAudioDetails.Trim()
        If audioDetails.Length = 0 Then
            audioDetails = "-"
        End If

        Dim addressDisplay = currentOscAddress.Trim()
        If addressDisplay.Length = 0 Then
            addressDisplay = "-"
        End If

        If InvokeRequired Then
            BeginInvoke(New Action(
                Sub()
                    fileNameLabel.Text = displayName
                    remainingLargeLabel.Text = remainingDisplay
                    producerValueLabel.Text = producerDisplay
                    audioValueLabel.Text = audioDetails
                    leftAudioMeterPanel.Invalidate()
                    rightAudioMeterPanel.Invalidate()
                    addressValueLabel.Text = addressDisplay
                    statusLabel.Text = String.Format("Live OSC update received for {0}-{1}.", selectedChannel, selectedLayer)
                End Sub))
            Return
        End If

        fileNameLabel.Text = displayName
        remainingLargeLabel.Text = remainingDisplay
        producerValueLabel.Text = producerDisplay
        audioValueLabel.Text = audioDetails
        leftAudioMeterPanel.Invalidate()
        rightAudioMeterPanel.Invalidate()
        addressValueLabel.Text = addressDisplay
        statusLabel.Text = String.Format("Live OSC update received for {0}-{1}.", selectedChannel, selectedLayer)
    End Sub

    Private Sub ResetDisplay()
        currentFilePath = String.Empty
        currentProducer = "-"
        currentOscAddress = "-"
        currentTimeSeconds = 0.0R
        currentDurationSeconds = 0.0R
        currentPlayedFrames = 0.0R
        currentTotalFrames = 0.0R
        currentFps = 0.0R
        currentLeftAudioLevel = 0.0R
        currentRightAudioLevel = 0.0R
        currentAudioDetails = "-"
        currentLeftAudioText = "-"
        currentRightAudioText = "-"
        pendingLeftAudioLevel = 0.0R
        pendingRightAudioLevel = 0.0R
        pendingLeftAudioLevelSum = 0.0R
        pendingRightAudioLevelSum = 0.0R
        pendingLeftAudioLevelCount = 0
        pendingRightAudioLevelCount = 0
        pendingLeftAudioText = "-"
        pendingRightAudioText = "-"
        hasPendingAudioSample = False
        lastDbfsAudioAt = DateTime.MinValue
        hasDbfsAudioSample = False
        fileNameLabel.Text = "Waiting for OSC data..."
        remainingLargeLabel.Text = "--:--:--"
        producerValueLabel.Text = "-"
        audioValueLabel.Text = "L: - dBFS   R: - dBFS"
        leftAudioMeterPanel.Invalidate()
        rightAudioMeterPanel.Invalidate()
        addressValueLabel.Text = "-"
    End Sub

    Private Sub LeftAudioMeterPanel_Paint(sender As Object, e As PaintEventArgs)
        PaintVerticalAudioMeter(e, leftAudioMeterPanel, currentLeftAudioLevel)
    End Sub

    Private Sub RightAudioMeterPanel_Paint(sender As Object, e As PaintEventArgs)
        PaintVerticalAudioMeter(e, rightAudioMeterPanel, currentRightAudioLevel)
    End Sub

    Private Shared Sub PaintVerticalAudioMeter(e As PaintEventArgs, panel As Panel, level As Double)
        Dim meterBounds = panel.ClientRectangle
        e.Graphics.Clear(panel.BackColor)

        If meterBounds.Width <= 2 OrElse meterBounds.Height <= 2 Then
            Return
        End If

        meterBounds.Inflate(-1, -1)
        Using backgroundBrush As New SolidBrush(Color.FromArgb(24, 24, 24))
            e.Graphics.FillRectangle(backgroundBrush, meterBounds)
        End Using

        PaintVerticalMeterZone(e.Graphics, meterBounds, 0.6R, 0.0R, Color.FromArgb(28, 90, 28))
        PaintVerticalMeterZone(e.Graphics, meterBounds, 0.85R, 0.6R, Color.FromArgb(120, 105, 18))
        PaintVerticalMeterZone(e.Graphics, meterBounds, 1.0R, 0.85R, Color.FromArgb(120, 28, 28))

        Dim clampedLevel = Clamp(level, 0.0R, 1.0R)
        If clampedLevel <= 0.0R Then
            DrawVerticalMeterGuides(e.Graphics, meterBounds)
            Return
        End If

        PaintVerticalMeterZone(e.Graphics, meterBounds, Math.Min(clampedLevel, 0.6R), 0.0R, Color.LimeGreen)
        PaintVerticalMeterZone(e.Graphics, meterBounds, Math.Min(clampedLevel, 0.85R), 0.6R, Color.Gold)
        PaintVerticalMeterZone(e.Graphics, meterBounds, clampedLevel, 0.85R, Color.OrangeRed)
        DrawVerticalMeterGuides(e.Graphics, meterBounds)
    End Sub

    Private Shared Function GetFirstArgumentAsString(arguments As List(Of Object)) As String
        If arguments.Count = 0 OrElse arguments(0) Is Nothing Then
            Return String.Empty
        End If

        Return Convert.ToString(arguments(0), Globalization.CultureInfo.InvariantCulture)
    End Function

    Private Shared Function GetArgumentAsDouble(value As Object) As Double
        If value Is Nothing Then
            Return 0.0R
        End If

        Return Convert.ToDouble(value, Globalization.CultureInfo.InvariantCulture)
    End Function

    Private Shared Function GetAudioLevelData(arguments As List(Of Object)) As List(Of Double)
        Dim values As New List(Of Double)()

        For Each argument As Object In arguments
            If TypeOf argument Is Integer OrElse
               TypeOf argument Is Single OrElse
               TypeOf argument Is Double OrElse
               TypeOf argument Is Decimal Then
                values.Add(NormalizeAudioValue(GetArgumentAsDouble(argument)))
            End If
        Next

        Return values
    End Function

    Private Shared Function GetLayerSuffix(address As String, newPrefix As String, legacyPrefix As String) As String
        If address.StartsWith(newPrefix, StringComparison.OrdinalIgnoreCase) Then
            Return address.Substring(newPrefix.Length)
        End If

        If address.StartsWith(legacyPrefix, StringComparison.OrdinalIgnoreCase) Then
            Return address.Substring(legacyPrefix.Length)
        End If

        Return String.Empty
    End Function

    Private Shared Function IsOneOf(value As String, ParamArray candidates As String()) As Boolean
        For Each candidate In candidates
            If String.Equals(value, candidate, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        Next

        Return False
    End Function

    Private Shared Function FormatAudioLevels(values As List(Of Double)) As String
        If values.Count = 0 Then
            Return "L: -   R: -"
        End If

        Dim leftValue = Clamp(values(0), 0.0R, 1.0R)
        Dim rightValue = Clamp(If(values.Count > 1, values(1), values(0)), 0.0R, 1.0R)
        Return String.Format(Globalization.CultureInfo.InvariantCulture, "L: {0:0.00}   R: {1:0.00}", leftValue, rightValue)
    End Function

    Private Function TryHandleMixerAudioMessage(address As String, arguments As List(Of Object), channelPrefix As String) As Boolean
        Dim mixerPrefix = channelPrefix & "mixer/audio/"

        If String.Equals(address, mixerPrefix & "volume", StringComparison.OrdinalIgnoreCase) Then
            If hasDbfsAudioSample OrElse DateTime.UtcNow.Subtract(lastDbfsAudioAt).TotalSeconds < 3.0R Then
                Return False
            End If

            Dim audioData = GetAudioLevelData(arguments)
            If audioData.Count = 0 Then
                Return False
            End If

            Dim leftRaw = GetArgumentAsDouble(arguments(0))
            Dim rightRaw = If(arguments.Count > 1, GetArgumentAsDouble(arguments(1)), leftRaw)
            UpdateAudioLevels(
                audioData(0),
                If(audioData.Count > 1, audioData(1), audioData(0)),
                String.Format(Globalization.CultureInfo.InvariantCulture, "VOL {0:0.000}", leftRaw),
                String.Format(Globalization.CultureInfo.InvariantCulture, "VOL {0:0.000}", rightRaw))

            currentOscAddress = address
            Return True
        End If

            Dim audioChannelIndex As Integer
        If TryParseMixerDbfsAddress(address, mixerPrefix, audioChannelIndex) AndAlso arguments.Count > 0 Then
            Dim rawDbfs = GetArgumentAsDouble(arguments(0))
            Dim normalized = NormalizeDbfsValue(rawDbfs)
            Dim clampedDbfs = Clamp(rawDbfs, -40.0R, 0.0R)
            Dim displayText = String.Format(Globalization.CultureInfo.InvariantCulture, "{0:0.0} dBFS", clampedDbfs)
            lastDbfsAudioAt = DateTime.UtcNow
            hasDbfsAudioSample = True

            If audioChannelIndex Mod 2 = 1 Then
                UpdateAudioLevels(normalized, Nothing, displayText, Nothing)
            Else
                UpdateAudioLevels(Nothing, normalized, Nothing, displayText)
            End If

            currentOscAddress = address
            Return True
        End If

        Return False
    End Function

    Private Function TryParseMixerDbfsAddress(address As String, mixerPrefix As String, ByRef audioChannelIndex As Integer) As Boolean
        audioChannelIndex = 0

        If Not address.StartsWith(mixerPrefix, StringComparison.OrdinalIgnoreCase) Then
            Return False
        End If

        If Not address.EndsWith("/dBFS", StringComparison.OrdinalIgnoreCase) Then
            Return False
        End If

        Dim suffix = address.Substring(mixerPrefix.Length)
        Dim separatorIndex = suffix.IndexOf("/"c)
        If separatorIndex <= 0 Then
            Return False
        End If

        Return Integer.TryParse(suffix.Substring(0, separatorIndex), audioChannelIndex)
    End Function

    Private Sub UpdateAudioLevels(leftLevel As Double?, rightLevel As Double?, leftText As String, rightText As String)
        If leftLevel.HasValue Then
            Dim clampedLeft = Clamp(leftLevel.Value, 0.0R, 1.0R)
            pendingLeftAudioLevel = Math.Max(pendingLeftAudioLevel, clampedLeft)
            pendingLeftAudioLevelSum += clampedLeft
            pendingLeftAudioLevelCount += 1
        End If

        If rightLevel.HasValue Then
            Dim clampedRight = Clamp(rightLevel.Value, 0.0R, 1.0R)
            pendingRightAudioLevel = Math.Max(pendingRightAudioLevel, clampedRight)
            pendingRightAudioLevelSum += clampedRight
            pendingRightAudioLevelCount += 1
        End If

        If leftText IsNot Nothing Then
            pendingLeftAudioText = leftText
        End If

        If rightText IsNot Nothing Then
            pendingRightAudioText = rightText
        End If

        hasPendingAudioSample = True
    End Sub

    Private Function FormatCurrentAudioDetails() As String
        If Not hasDbfsAudioSample Then
            Return "L: - dBFS   R: - dBFS"
        End If

        Return String.Format("L: {0}   R: {1}", currentLeftAudioText, currentRightAudioText)
    End Function

    Private Function TryUpdateTimeFromSeconds(arguments As List(Of Object), address As String) As Boolean
        If arguments.Count < 2 Then
            Return False
        End If

        currentTimeSeconds = GetArgumentAsDouble(arguments(0))
        currentDurationSeconds = GetArgumentAsDouble(arguments(1))
        currentOscAddress = address
        Return True
    End Function

    Private Function TryUpdateTimeFromFrames(arguments As List(Of Object), address As String) As Boolean
        If arguments.Count < 2 Then
            Return False
        End If

        currentPlayedFrames = GetArgumentAsDouble(arguments(0))
        currentTotalFrames = GetArgumentAsDouble(arguments(1))

        If currentFps > 0.0R Then
            currentTimeSeconds = currentPlayedFrames / currentFps
            currentDurationSeconds = currentTotalFrames / currentFps
        End If

        currentOscAddress = address
        Return True
    End Function

    Private Function TryUpdateFps(arguments As List(Of Object), address As String) As Boolean
        If arguments.Count = 0 Then
            Return False
        End If

        currentFps = GetArgumentAsDouble(arguments(0))
        If currentFps > 0.0R AndAlso currentTotalFrames > 0.0R Then
            currentTimeSeconds = currentPlayedFrames / currentFps
            currentDurationSeconds = currentTotalFrames / currentFps
        End If

        currentOscAddress = address
        Return True
    End Function

    Private Shared Sub PaintVerticalMeterZone(graphics As Graphics, meterBounds As Rectangle, fromRatio As Double, toRatio As Double, zoneColor As Color)
        Dim topRatio = Clamp(fromRatio, 0.0R, 1.0R)
        Dim bottomRatio = Clamp(toRatio, 0.0R, 1.0R)
        If topRatio <= bottomRatio Then
            Return
        End If

        Dim zoneBottom = meterBounds.Bottom - CInt(Math.Round(meterBounds.Height * bottomRatio, MidpointRounding.AwayFromZero))
        Dim zoneTop = meterBounds.Bottom - CInt(Math.Round(meterBounds.Height * topRatio, MidpointRounding.AwayFromZero))
        Dim zoneHeight = Math.Max(1, zoneBottom - zoneTop)

        Using zoneBrush As New SolidBrush(zoneColor)
            graphics.FillRectangle(zoneBrush, New Rectangle(meterBounds.X + 2, zoneTop, Math.Max(1, meterBounds.Width - 4), zoneHeight))
        End Using
    End Sub

    Private Shared Sub DrawVerticalMeterGuides(graphics As Graphics, meterBounds As Rectangle)
        For guideIndex = 1 To 4
            Dim y = meterBounds.Top + CInt((meterBounds.Height / 5.0R) * guideIndex)
            Using guidePen As New Pen(Color.FromArgb(55, 255, 255, 255))
                graphics.DrawLine(guidePen, meterBounds.Left, y, meterBounds.Right, y)
            End Using
        Next
    End Sub

    Private Shared Function NormalizeAudioValue(rawValue As Double) As Double
        If Double.IsNaN(rawValue) OrElse Double.IsInfinity(rawValue) Then
            Return 0.0R
        End If

        If rawValue < 0.0R Then
            Return Clamp((rawValue + 80.0R) / 80.0R, 0.0R, 1.0R)
        End If

        Dim normalized = rawValue
        If normalized <= 1.0R Then
            normalized = Math.Pow(normalized, 0.25R)
        ElseIf normalized <= 255.0R Then
            normalized = normalized / 255.0R
        ElseIf normalized <= 10.0R Then
            normalized = normalized / 10.0R
        Else
            normalized = normalized / 100.0R
        End If

        If normalized > 0.0R AndAlso normalized < 0.08R Then
            normalized = 0.08R
        End If

        If normalized < 0.02R Then
            normalized = 0.0R
        End If

        Return Clamp(normalized, 0.0R, 1.0R)
    End Function

    Private Shared Function NormalizeDbfsValue(rawDbfs As Double) As Double
        If Double.IsNaN(rawDbfs) OrElse Double.IsInfinity(rawDbfs) Then
            Return 0.0R
        End If

        Dim progressValue = Clamp(40.0R + rawDbfs, 0.0R, 40.0R)
        Dim normalized = progressValue / 40.0R

        ' Compress the visual response so normal program audio does not sit near full scale.
        Return Math.Pow(normalized, 1.8R)
    End Function

    Private Shared Function Clamp(value As Double, minimum As Double, maximum As Double) As Double
        Return Math.Max(minimum, Math.Min(maximum, value))
    End Function

    Private Function FormatRemainingTime() As String
        If currentDurationSeconds <= 0.0R Then
            Return "--:--:--"
        End If

        Dim remaining = Math.Max(0.0R, currentDurationSeconds - currentTimeSeconds)
        Dim span = TimeSpan.FromSeconds(remaining)
        Return span.ToString("hh\:mm\:ss", Globalization.CultureInfo.InvariantCulture)
    End Function

    Private Shared Function ReadOscArguments(buffer As Byte(), ByRef offset As Integer, typeTag As String) As List(Of Object)
        Dim values As New List(Of Object)()

        For Each typeCode As Char In typeTag
            If typeCode = ","c Then
                Continue For
            End If

            Select Case typeCode
                Case "s"c
                    values.Add(ReadOscString(buffer, offset))
                Case "i"c
                    values.Add(ReadInt32BigEndian(buffer, offset))
                    offset += 4
                Case "f"c
                    values.Add(ReadSingleBigEndian(buffer, offset))
                    offset += 4
                Case "T"c
                    values.Add(True)
                Case "F"c
                    values.Add(False)
                Case Else
                    Exit For
            End Select
        Next

        Return values
    End Function

    Private Shared Function ReadOscString(buffer As Byte(), ByRef offset As Integer) As String
        If offset >= buffer.Length Then
            Return String.Empty
        End If

        Dim start = offset
        While offset < buffer.Length AndAlso buffer(offset) <> 0
            offset += 1
        End While

        Dim value = Encoding.UTF8.GetString(buffer, start, offset - start)

        While offset < buffer.Length AndAlso buffer(offset) = 0
            offset += 1
            If offset Mod 4 = 0 Then
                Exit While
            End If
        End While

        Return value
    End Function

    Private Shared Function ReadInt32BigEndian(buffer As Byte(), offset As Integer) As Integer
        If offset + 4 > buffer.Length Then
            Return 0
        End If

        Return (buffer(offset) << 24) Or (buffer(offset + 1) << 16) Or (buffer(offset + 2) << 8) Or buffer(offset + 3)
    End Function

    Private Shared Function ReadSingleBigEndian(buffer As Byte(), offset As Integer) As Single
        If offset + 4 > buffer.Length Then
            Return 0.0F
        End If

        Dim bytes = New Byte() {buffer(offset + 3), buffer(offset + 2), buffer(offset + 1), buffer(offset)}
        Return BitConverter.ToSingle(bytes, 0)
    End Function

    Private Sub StopListener()
        Dim active = listener
        listener = Nothing

        If active IsNot Nothing Then
            active.Close()
            active.Dispose()
        End If
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        uiRefreshTimer.Stop()
        StopListener()
        MyBase.OnFormClosing(e)
    End Sub
End Class
