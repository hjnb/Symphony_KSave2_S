Imports System.Text

Public Class ExTextBox
    Inherits TextBox

    Private mouseDownFlg As Boolean = False
    Private _limitLengthByte As Integer = 100
    Private _inputType As Integer = 0

    Public Property LimitLengthByte() As Integer
        Get
            Return _limitLengthByte
        End Get
        Set(value As Integer)
            _limitLengthByte = value
        End Set
    End Property

    Public Property InputType() As Integer
        Get
            Return _inputType
        End Get
        Set(value As Integer)
            If value = 1 Then
                _inputType = 1
            Else
                _inputType = 0
            End If
        End Set
    End Property

    Private Sub ExTextBox_Enter(sender As Object, e As System.EventArgs) Handles Me.Enter
        Me.SelectionStart = Me.TextLength
        mouseDownFlg = True
    End Sub

    Private Sub ExTextBox_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        If mouseDownFlg = True Then
            Me.SelectionStart = Me.TextLength
            mouseDownFlg = False
        End If
    End Sub

    Private Sub ExTextBox_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        Dim text As String = CType(sender, ExTextBox).Text
        Dim lengthByte As Integer = Encoding.GetEncoding("Shift_JIS").GetByteCount(text)
        If _inputType = 1 Then
            If lengthByte >= _limitLengthByte Then '設定されているバイト数以上の時
                If e.KeyChar = ChrW(Keys.Back) Then
                    'Backspaceは入力可能
                    e.Handled = False
                Else
                    '入力できなくする
                    e.Handled = True
                End If
            Else
                If (ChrW(Keys.D0) <= e.KeyChar AndAlso e.KeyChar <= ChrW(Keys.D9)) OrElse e.KeyChar = ChrW(Keys.Back) Then
                    e.Handled = False
                Else
                    e.Handled = True
                End If
            End If
        Else
            If lengthByte >= _limitLengthByte Then '設定されているバイト数以上の時
                If e.KeyChar = ChrW(Keys.Back) Then
                    'Backspaceは入力可能
                    e.Handled = False
                Else
                    '入力できなくする
                    e.Handled = True
                End If
            End If
        End If
    End Sub
End Class
