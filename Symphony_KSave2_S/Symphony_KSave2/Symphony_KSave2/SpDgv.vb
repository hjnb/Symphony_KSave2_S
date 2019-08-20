Imports System.Text

Public Class SpDgv
    Inherits DataGridView

    Private cellEnterFlg As Boolean = False
    Private selectedRowIndex As Integer = 0
    Private CRR_LIMIT_LENGTHBYTE As Integer = 6
    Private TXT_LIMIT_LENGTHBYTE As Integer = 100

    Public dt As New DataTable()

    Protected Overrides Function ProcessDialogKey(keyData As System.Windows.Forms.Keys) As Boolean
        If keyData = Keys.Enter Then
            If Me.Columns(CurrentCell.ColumnIndex).Name = "Crr" Then
                Me.EndEdit()
                Dim inputStr As String = StrConv(CurrentCell.Value.ToString(), VbStrConv.Narrow) '半角に変換

                If System.Text.RegularExpressions.Regex.IsMatch(inputStr, "^\d+(,\d+|\-\d+|\.\d+)*$") Then
                    CurrentCell.Value = inputStr
                Else
                    CurrentCell.Value = ""
                End If
            End If
            Return Me.ProcessTabKey(keyData)
        Else
            Return MyBase.ProcessDialogKey(keyData)
        End If
    End Function

    Protected Overrides Function ProcessDataGridViewKey(e As System.Windows.Forms.KeyEventArgs) As Boolean
        If e.KeyCode = Keys.Enter Then
            Return Me.ProcessTabKey(e.KeyCode)
        End If

        Dim tb As DataGridViewTextBoxEditingControl = CType(Me.EditingControl, DataGridViewTextBoxEditingControl)
        If Not IsNothing(tb) AndAlso ((e.KeyCode = Keys.Left AndAlso tb.SelectionStart = 0) OrElse (e.KeyCode = Keys.Right AndAlso tb.SelectionStart = tb.TextLength)) Then
            Return False
        Else
            Return MyBase.ProcessDataGridViewKey(e)
        End If
    End Function

    Public Sub clearText()
        For Each row As DataGridViewRow In Me.Rows
            row.Cells(0).Value = ""
            row.Cells(1).Value = ""
        Next
    End Sub

    Public Sub rowInsert()
        Dim row As DataRow = dt.NewRow()
        row(0) = ""
        row(1) = ""
        dt.Rows.InsertAt(row, selectedRowIndex)
        dt.Rows.RemoveAt(dt.Rows.Count - 1)
    End Sub

    Public Sub rowDelete()
        dt.Rows.RemoveAt(selectedRowIndex)
        Dim row As DataRow = dt.NewRow()
        row(0) = ""
        row(1) = ""
        dt.Rows.Add(row)
    End Sub

    Private Sub SpDgv_CellEnter(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles Me.CellEnter
        If cellEnterFlg Then
            Me.BeginEdit(False)
            selectedRowIndex = e.RowIndex
        End If
    End Sub

    Private Sub SpDgv_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles Me.CellMouseClick
        cellEnterFlg = True
        Me.BeginEdit(False)
        selectedRowIndex = e.RowIndex
    End Sub

    Private Sub SpDgvTextBox_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs)
        Dim text As String = CType(sender, DataGridViewTextBoxEditingControl).Text
        Dim lengthByte As Integer = Encoding.GetEncoding("Shift_JIS").GetByteCount(text)
        Dim limitLengthByte As Integer = If(Me.Columns(Me.CurrentCell.ColumnIndex).Name = "Crr", CRR_LIMIT_LENGTHBYTE, TXT_LIMIT_LENGTHBYTE)

        If lengthByte >= limitLengthByte Then '設定されているバイト数以上の時
            If e.KeyChar = ChrW(Keys.Back) Then
                'Backspaceは入力可能
                e.Handled = False
            Else
                '入力できなくする
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub SpDgv_EditingControlShowing(sender As Object, e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles Me.EditingControlShowing
        Dim editTextBox As DataGridViewTextBoxEditingControl = CType(e.Control, DataGridViewTextBoxEditingControl)
        editTextBox.ImeMode = If(Me.Columns(Me.CurrentCell.ColumnIndex).Name = "Crr", Windows.Forms.ImeMode.NoControl, Windows.Forms.ImeMode.Hiragana)

        'イベントハンドラを削除、追加
        RemoveHandler editTextBox.KeyPress, AddressOf SpDgvTextBox_KeyPress
        AddHandler editTextBox.KeyPress, AddressOf SpDgvTextBox_KeyPress
    End Sub
End Class
