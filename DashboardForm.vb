Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Windows.Forms

Public Class MainForm
    Inherits Form

    Private ReadOnly newsControl As New StudioMonitorControl("NEWS")
    Private ReadOnly commercialControl As New StudioMonitorControl("COMMERCIAL")
    Private ReadOnly promoControl As New StudioMonitorControl("PROMO")

    Public Sub New()
        Text = "Caspar Layer Monitor"
        StartPosition = FormStartPosition.CenterScreen
        FormBorderStyle = FormBorderStyle.Sizable
        MaximizeBox = True
        MinimumSize = New Size(220, 120)
        ClientSize = New Size(1290, 470)

        Controls.Add(newsControl)
        Controls.Add(commercialControl)
        Controls.Add(promoControl)

        AddHandler Resize, AddressOf MainForm_Resize
        LayoutStudioControls()
    End Sub

    Private Sub MainForm_Resize(sender As Object, e As EventArgs)
        LayoutStudioControls()
    End Sub

    Private Sub LayoutStudioControls()
        Const margin As Integer = 10
        Const spacing As Integer = 10
        Const singlePanelThreshold As Integer = 760
        Const dualPanelThreshold As Integer = 1140

        Dim availableHeight = Math.Max(80, ClientSize.Height - (margin * 2))
        Dim availableWidth = Math.Max(80, ClientSize.Width - (margin * 2))

        If ClientSize.Width < singlePanelThreshold Then
            ShowOnePanel(newsControl, availableWidth, availableHeight, margin)
            commercialControl.Visible = False
            promoControl.Visible = False
            Return
        End If

        If ClientSize.Width < dualPanelThreshold Then
            Dim dualWidth = Math.Max(80, (ClientSize.Width - (margin * 2) - spacing) \ 2)

            ShowPanel(newsControl, margin, margin, dualWidth, availableHeight)
            ShowPanel(commercialControl, newsControl.Right + spacing, margin, dualWidth, availableHeight)
            promoControl.Visible = False
            Return
        End If

        Dim tripleWidth = Math.Max(80, (ClientSize.Width - (margin * 2) - (spacing * 2)) \ 3)

        ShowPanel(newsControl, margin, margin, tripleWidth, availableHeight)
        ShowPanel(commercialControl, newsControl.Right + spacing, margin, tripleWidth, availableHeight)
        ShowPanel(promoControl, commercialControl.Right + spacing, margin, tripleWidth, availableHeight)
    End Sub

    Private Shared Sub ShowOnePanel(panel As Control, width As Integer, height As Integer, margin As Integer)
        ShowPanel(panel, margin, margin, width, height)
    End Sub

    Private Shared Sub ShowPanel(panel As Control, x As Integer, y As Integer, width As Integer, height As Integer)
        panel.Location = New Point(x, y)
        panel.Size = New Size(width, height)
        panel.Visible = True
    End Sub
End Class
