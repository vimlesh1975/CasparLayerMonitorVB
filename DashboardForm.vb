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
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False
        ClientSize = New Size(860, 470)

        studioAControl.Location = New Point(10, 10)
        studioAControl.Size = New Size(420, 430)

        studioBControl.Location = New Point(430, 10)
        studioBControl.Size = New Size(420, 430)

        Controls.Add(studioAControl)
        Controls.Add(studioBControl)
    End Sub
End Class
