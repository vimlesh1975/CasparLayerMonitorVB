Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Windows.Forms

Public Class MainForm
    Inherits Form

    Private ReadOnly studioAControl As New StudioMonitorControl("Studio A")
    Private ReadOnly studioBControl As New StudioMonitorControl("Studio B")

    Public Sub New()
        Text = "Caspar Layer Monitor"
        StartPosition = FormStartPosition.CenterScreen
        FormBorderStyle = FormBorderStyle.Sizable
        MaximizeBox = True
        MinimumSize = New Size(220, 120)
        ClientSize = New Size(860, 470)

        Controls.Add(studioAControl)
        Controls.Add(studioBControl)

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

        Dim availableHeight = Math.Max(80, ClientSize.Height - (margin * 2))
        Dim availableWidth = Math.Max(80, ClientSize.Width - (margin * 2))

        If ClientSize.Width < singlePanelThreshold Then
            studioAControl.Location = New Point(margin, margin)
            studioAControl.Size = New Size(availableWidth, availableHeight)
            studioAControl.Visible = True

            studioBControl.Visible = False
            Return
        End If

        Dim dualWidth = Math.Max(80, (ClientSize.Width - (margin * 2) - spacing) \ 2)

        studioAControl.Location = New Point(margin, margin)
        studioAControl.Size = New Size(dualWidth, availableHeight)
        studioAControl.Visible = True

        studioBControl.Location = New Point(studioAControl.Right + spacing, margin)
        studioBControl.Size = New Size(dualWidth, availableHeight)
        studioBControl.Visible = True
    End Sub
End Class
