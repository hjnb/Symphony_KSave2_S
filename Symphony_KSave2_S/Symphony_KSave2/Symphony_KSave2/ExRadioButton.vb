﻿Public Class ExRadioButton
    Inherits RadioButton

    Private Sub ExRadioButton_CheckedChanged(sender As Object, e As System.EventArgs) Handles Me.CheckedChanged
        If Me.Checked = True Then
            Me.BackColor = Color.FromArgb(255, 192, 255)
        Else
            Me.BackColor = Color.FromKnownColor(KnownColor.Control)
        End If
    End Sub
End Class
