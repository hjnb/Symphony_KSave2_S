Public Class ExDataGridView
    Inherits DataGridView

    Private editingColumnIndex As Integer = 0

    Private _targetTextBox As ymdBox.ymdBox

    Public Property targetTextBox() As ymdBox.ymdBox
        Get
            Return _targetTextBox
        End Get

        Set(ByVal value As ymdBox.ymdBox)
            _targetTextBox = value
        End Set
    End Property

    Protected Overrides Function ProcessDialogKey(keyData As System.Windows.Forms.Keys) As Boolean
        If keyData = Keys.Enter OrElse keyData = Keys.Tab Then
            If editingColumnIndex = 5 Then
                Me.CurrentCell = Me(11, 0)
            End If
            Return Me.ProcessTabKey(keyData)
        Else
            Return MyBase.ProcessDialogKey(keyData)
        End If
    End Function

    Protected Overrides Function ProcessDataGridViewKey(e As System.Windows.Forms.KeyEventArgs) As Boolean
        Dim tb As DataGridViewTextBoxEditingControl = CType(Me.EditingControl, DataGridViewTextBoxEditingControl)
        If Not IsNothing(tb) AndAlso ((e.KeyCode = Keys.Left AndAlso tb.SelectionStart = 0) OrElse (e.KeyCode = Keys.Right AndAlso tb.SelectionStart = tb.TextLength)) Then
            Return False
        Else
            If editingColumnIndex = 21 Then
                If Not IsNothing(_targetTextBox) Then
                    _targetTextBox.Focus()
                End If
            End If
            Return MyBase.ProcessDataGridViewKey(e)
        End If
    End Function

    Private Sub ExDataGridView_CellEnter(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles Me.CellEnter
        If Me(e.ColumnIndex, e.RowIndex).ReadOnly = False Then
            Me.BeginEdit(False)
            editingColumnIndex = e.ColumnIndex
            Me(e.ColumnIndex, e.RowIndex).Style.Alignment = DataGridViewContentAlignment.TopCenter
        End If
    End Sub

    Private Sub ExDataGridView_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles Me.CellLeave
        Me(e.ColumnIndex, e.RowIndex).Style.Alignment = DataGridViewContentAlignment.BottomCenter
    End Sub

    Private Sub ExDataGridView_EditingControlShowing(sender As Object, e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles Me.EditingControlShowing
        If TypeOf e.Control Is DataGridViewTextBoxEditingControl Then
            Dim dgv As DataGridView = CType(sender, DataGridView)
            Dim tb As DataGridViewTextBoxEditingControl = CType(e.Control, DataGridViewTextBoxEditingControl)
            tb.ImeMode = Windows.Forms.ImeMode.Disable
            tb.MaxLength = 1

            'イベントハンドラを削除
            RemoveHandler tb.KeyDown, AddressOf dataGridViewTextBox_KeyDown

            'イベントハンドラを追加
            AddHandler tb.KeyDown, AddressOf dataGridViewTextBox_KeyDown
        End If
    End Sub

    Private Sub dataGridViewTextBox_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
        If Not ((Keys.D0 <= e.KeyCode AndAlso e.KeyCode <= Keys.D9) OrElse (Keys.NumPad0 <= e.KeyCode AndAlso e.KeyCode <= Keys.NumPad9) OrElse e.KeyCode = Keys.Back OrElse e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Left OrElse e.KeyCode = Keys.Right) Then
            e.SuppressKeyPress = True
        End If
    End Sub
End Class
