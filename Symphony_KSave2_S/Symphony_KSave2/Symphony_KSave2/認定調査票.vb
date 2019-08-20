Imports System.Data.OleDb
Imports System.Text
Imports ymdBox.ymdBox
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Microsoft.Reporting.WinForms

Public Class 認定調査票

    Private Const INPUT_NUMBER As Integer = 1

    Public Sub New()
        InitializeComponent()
        Me.WindowState = FormWindowState.Maximized
        Me.KeyPreview = True
    End Sub

    Private Sub 認定調査票_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            If e.Control = False Then
                Me.SelectNextControl(Me.ActiveControl, Not e.Shift, True, True, True)
            End If
        End If
    End Sub

    Private Sub 認定調査票_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        '利用者リスト表示
        displayUserList()

        'dgv初期設定
        initDgvNumInput()
        initDgvSp(SpDgv1)
        initDgvSp(SpDgv2)
        initDgvSp(SpDgv3)
        initDgvSp(SpDgv4)
        initDgvSp(SpDgv5)
        initDgvSp(SpDgv6)
        initDgvSp(SpDgv7)

        '初期フォーカス
        dgvNumInput.Focus()
        SendKeys.Send("{ESC}")
        SendKeys.Send("{F2}")

        '入力ボックス設定
        settingInputBox()
        clearOverviewPageInputBox()

        'エンターキーでの処理用設定
        dgvNumInput.targetTextBox = dateYmdBox

    End Sub

    Private Sub settingUserList()
        'DoubleBufferedプロパティをTrue
        Util.EnableDoubleBuffering(userList)

        'dgv設定
        With userList
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect 'クリック時に行選択
            .MultiSelect = False
            .ReadOnly = True
            .ColumnHeadersVisible = False
            .RowHeadersVisible = False
            .RowTemplate.Height = 14
            .CellBorderStyle = DataGridViewCellBorderStyle.None
            .ShowCellToolTips = False
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
            .DefaultCellStyle.BackColor = Color.FromKnownColor(KnownColor.Control)
        End With
    End Sub

    Private Sub displayUserList()
        settingUserList()

        Dim cnn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim sql As String = "select Nam, Kana from UsrM order by Kana"
        cnn.Open(topForm.DB_KSave2)
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        Dim da As OleDbDataAdapter = New OleDbDataAdapter()
        Dim ds As DataSet = New DataSet()
        da.Fill(ds, rs, "UsrM")
        userList.DataSource = ds.Tables(0)
        cnn.Close()

        userList.Columns("Kana").Visible = False
        userList.Columns("Nam").Width = 89
        userList.CurrentCell.Selected = False
    End Sub

    Private Sub displayRecordList(userNam As String)
        Dim cnn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim sql As String = "select distinct Ymd1 from Auth1 where Nam='" & userNam & "' order by Ymd1 Desc"
        cnn.Open(topForm.DB_KSave2)
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        recordList.Items.Clear()
        While Not rs.EOF
            recordList.Items.Add(convADStrToWarekiStr(rs.Fields("Ymd1").Value))
            rs.MoveNext()
        End While
        rs.Close()
        cnn.Close()
    End Sub

    Private Sub userList_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles userList.CellMouseClick
        clearAllInputData()
        Dim userNam As String = userList("Nam", e.RowIndex).Value
        Dim userKana As String = userList("Kana", e.RowIndex).Value
        kanaLabel.Text = userKana
        userLabel.Text = userNam
        displayRecordList(userNam)
    End Sub

    Private Sub initDgvNumInput()
        Util.EnableDoubleBuffering(dgvNumInput)

        With dgvNumInput
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False '行削除禁止
            .MultiSelect = False
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersVisible = False
            .RowHeadersVisible = False
            .RowTemplate.Height = 29
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            .DefaultCellStyle.Font = New Font("MS UI Gothic", 14, FontStyle.Bold)
            .DefaultCellStyle.BackColor = Color.FromArgb(145, 172, 244)
            .DefaultCellStyle.SelectionBackColor = Color.FromArgb(145, 172, 244)
            .DefaultCellStyle.SelectionForeColor = Color.Black
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
            .ShowCellToolTips = False
            .BorderStyle = System.Windows.Forms.BorderStyle.None
            .GridColor = Color.FromArgb(236, 233, 216)
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        End With

        '空セル作成
        Dim dt As New DataTable()
        For i As Integer = 1 To 6
            dt.Columns.Add("GDay" & i, Type.GetType("System.String"))
        Next
        dt.Columns.Add("GSpace1", Type.GetType("System.String"))
        For i As Integer = 1 To 4
            dt.Columns.Add("GAuto" & i, Type.GetType("System.String"))
        Next
        dt.Columns.Add("GSpace2", Type.GetType("System.String"))
        For i As Integer = 1 To 10
            dt.Columns.Add("GNum" & i, Type.GetType("System.String"))
        Next
        Dim row As DataRow = dt.NewRow()
        row("GAuto1") = "0"
        row("GAuto2") = "1"
        row("GAuto3") = "1"
        row("GAuto4") = "3"
        dt.Rows.Add(row)
        dgvNumInput.DataSource = dt

        With dgvNumInput
            For i = 1 To 6
                With .Columns("GDay" & i)
                    .Width = 23
                End With
            Next
            For i = 1 To 4
                With .Columns("GAuto" & i)
                    .Width = 23
                    .ReadOnly = True
                    .DefaultCellStyle.SelectionBackColor = Color.FromArgb(145, 172, 244)
                    .DefaultCellStyle.SelectionForeColor = Color.Black
                End With
            Next
            For i = 1 To 10
                With .Columns("GNum" & i)
                    .Width = 23
                End With
            Next
            For i = 1 To 2
                With .Columns("GSpace" & i)
                    .Width = 12
                    .ReadOnly = True
                    .DefaultCellStyle.BackColor = Color.FromArgb(236, 233, 216)
                    .DefaultCellStyle.SelectionBackColor = Color.FromArgb(236, 233, 216)
                End With
            Next
        End With
    End Sub

    Private Sub initDgvSp(dgv As SpDgv)
        Util.EnableDoubleBuffering(dgv)

        With dgv
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False '行削除禁止
            .MultiSelect = False
            .RowHeadersVisible = False
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersHeight = 19
            .RowTemplate.Height = 17
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
            .DefaultCellStyle.SelectionBackColor = Color.White
            .DefaultCellStyle.SelectionForeColor = Color.Black
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .ShowCellToolTips = False
            .EnableHeadersVisualStyles = False
        End With

        '列追加、空の行追加
        dgv.dt.Columns.Add("Crr", Type.GetType("System.String"))
        dgv.dt.Columns.Add("Txt", Type.GetType("System.String"))
        Dim row As DataRow
        For i = 0 To 59
            row = dgv.dt.NewRow()
            row(0) = ""
            row(1) = ""
            dgv.dt.Rows.Add(row)
        Next

        dgv.DataSource = dgv.dt

        With dgv
            With .Columns("Crr")
                .HeaderText = "項目"
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Width = 47
            End With
            With .Columns("Txt")
                .HeaderText = "内容"
                .Width = 530
            End With
        End With

        '並び替えができないようにする
        For Each c As DataGridViewColumn In dgv.Columns
            c.SortMode = DataGridViewColumnSortMode.NotSortable
        Next

    End Sub

    Private Sub settingInputBox()
        '実施者ボックス
        Dim cnn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim sql As String = "select Nam from EtcM order by Num"
        cnn.Open(topForm.DB_KSave2)
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        etcBox.Items.Clear()
        While Not rs.EOF
            etcBox.Items.Add(rs.Fields("Nam").Value)
            rs.MoveNext()
        End While
        rs.Close()
        cnn.Close()
        etcBox.ImeMode = Windows.Forms.ImeMode.Hiragana

        '所属機関ボックス
        companyBox.Items.AddRange({"特別養護老人ホーム シンフォニー", "居宅介護支援事業所 シンフォニー"})
        companyBox.ImeMode = Windows.Forms.ImeMode.Hiragana

        '実施場所自宅外ボックス
        houseTextBox.LimitLengthByte = 34 '全角17文字
        houseTextBox.ImeMode = Windows.Forms.ImeMode.Hiragana

        '前回認定結果ボックス
        certifiedResultBox.Items.AddRange({"非該当", "要支援1", "要支援2", "要介護1", "要介護2", "要介護3", "要介護4", "要介護5"})
        certifiedResultBox.ImeMode = Windows.Forms.ImeMode.Hiragana

        '現在所
        With currentPostCode1
            .InputType = INPUT_NUMBER
            .LimitLengthByte = 3
            .ImeMode = Windows.Forms.ImeMode.Disable
            .TextAlign = HorizontalAlignment.Center
        End With
        With currentPostCode2
            .InputType = INPUT_NUMBER
            .LimitLengthByte = 4
            .ImeMode = Windows.Forms.ImeMode.Disable
            .TextAlign = HorizontalAlignment.Center
        End With
        With currentAddress
            .LimitLengthByte = 60
            .ImeMode = Windows.Forms.ImeMode.Hiragana
        End With
        With currentTel1
            .InputType = INPUT_NUMBER
            .LimitLengthByte = 4
            .ImeMode = Windows.Forms.ImeMode.Disable
            .TextAlign = HorizontalAlignment.Center
        End With
        With currentTel2
            .InputType = INPUT_NUMBER
            .LimitLengthByte = 4
            .ImeMode = Windows.Forms.ImeMode.Disable
            .TextAlign = HorizontalAlignment.Center
        End With
        With currentTel3
            .InputType = INPUT_NUMBER
            .LimitLengthByte = 4
            .ImeMode = Windows.Forms.ImeMode.Disable
            .TextAlign = HorizontalAlignment.Center
        End With

        '家族等
        With familyPostCode1
            .InputType = INPUT_NUMBER
            .LimitLengthByte = 3
            .ImeMode = Windows.Forms.ImeMode.Disable
            .TextAlign = HorizontalAlignment.Center
        End With
        With familyPostCode2
            .InputType = INPUT_NUMBER
            .LimitLengthByte = 4
            .ImeMode = Windows.Forms.ImeMode.Disable
            .TextAlign = HorizontalAlignment.Center
        End With
        With familyAddress
            .LimitLengthByte = 60
            .ImeMode = Windows.Forms.ImeMode.Hiragana
        End With
        With familyTel1
            .InputType = INPUT_NUMBER
            .LimitLengthByte = 4
            .ImeMode = Windows.Forms.ImeMode.Disable
            .TextAlign = HorizontalAlignment.Center
        End With
        With familyTel2
            .InputType = INPUT_NUMBER
            .LimitLengthByte = 4
            .ImeMode = Windows.Forms.ImeMode.Disable
            .TextAlign = HorizontalAlignment.Center
        End With
        With familyTel3
            .InputType = INPUT_NUMBER
            .LimitLengthByte = 4
            .ImeMode = Windows.Forms.ImeMode.Disable
            .TextAlign = HorizontalAlignment.Center
        End With

        '氏名ボックス
        With namBox
            .LimitLengthByte = 16
            .ImeMode = Windows.Forms.ImeMode.Hiragana
        End With

        '調査対象者との関係ボックス
        relationBox.Items.AddRange({"夫", "妻", "息子", "娘", "長男", "二男", "三男", "四男", "長女", "二女", "三女", "四女", "五女", "子の嫁", "子の夫", "兄", "弟", "姉", "妹", "父", "母", "孫", "伯父", "叔父", "伯母", "叔母", "知人", "その他", "姪", "甥"})
        relationBox.MaxDropDownItems = 8
        relationBox.IntegralHeight = False
        relationBox.ImeMode = Windows.Forms.ImeMode.Hiragana

        'txtNum1～txtNum21ボックス
        For i = 1 To 21
            If i = 13 Then
                Continue For
            End If
            With CType(overview3Panel.Controls("txtNum" & i), ExTextBox)
                .InputType = INPUT_NUMBER
                .LimitLengthByte = 4
                .ImeMode = Windows.Forms.ImeMode.Disable
                .TextAlign = HorizontalAlignment.Center
            End With
        Next

        '市町村特別給付
        With txtGentxt1
            .LimitLengthByte = 90
            .ImeMode = Windows.Forms.ImeMode.Hiragana
        End With

        '介護保険給付外の在宅サービス
        With txtGentxt2
            .LimitLengthByte = 76
            .ImeMode = Windows.Forms.ImeMode.Hiragana
        End With

        '施設連絡先
        With facilityNameBox
            .LimitLengthByte = 40
            .ImeMode = Windows.Forms.ImeMode.Hiragana
        End With
        With facilityPostCode1
            .InputType = INPUT_NUMBER
            .LimitLengthByte = 3
            .ImeMode = Windows.Forms.ImeMode.Disable
            .TextAlign = HorizontalAlignment.Center
        End With
        With facilityPostCode2
            .InputType = INPUT_NUMBER
            .LimitLengthByte = 4
            .ImeMode = Windows.Forms.ImeMode.Disable
            .TextAlign = HorizontalAlignment.Center
        End With
        With facilityAddress
            .LimitLengthByte = 60
            .ImeMode = Windows.Forms.ImeMode.Hiragana
        End With
        With facilityTel1
            .InputType = INPUT_NUMBER
            .LimitLengthByte = 4
            .ImeMode = Windows.Forms.ImeMode.Disable
            .TextAlign = HorizontalAlignment.Center
        End With
        With facilityTel2
            .InputType = INPUT_NUMBER
            .LimitLengthByte = 4
            .ImeMode = Windows.Forms.ImeMode.Disable
            .TextAlign = HorizontalAlignment.Center
        End With
        With facilityTel3
            .InputType = INPUT_NUMBER
            .LimitLengthByte = 4
            .ImeMode = Windows.Forms.ImeMode.Disable
            .TextAlign = HorizontalAlignment.Center
        End With

        '特記テキスト
        With spText1
            .Font = New Font("MS UI Gothic", 9.4)
            .LimitLengthByte = 128
            .ImeMode = Windows.Forms.ImeMode.Hiragana
        End With
        With spText2
            .Font = New Font("MS UI Gothic", 9.4)
            .LimitLengthByte = 128
            .ImeMode = Windows.Forms.ImeMode.Hiragana
        End With
        With spText3
            .Font = New Font("MS UI Gothic", 9.4)
            .LimitLengthByte = 128
            .ImeMode = Windows.Forms.ImeMode.Hiragana
        End With
        With spText4
            .Font = New Font("MS UI Gothic", 9.4)
            .LimitLengthByte = 128
            .ImeMode = Windows.Forms.ImeMode.Hiragana
        End With

    End Sub

    Private Sub clearOverviewPageInputBox()
        Dim todayStr As String = Today.ToString("yyyy/MM/dd")
        '番号
        For Each cell As DataGridViewCell In dgvNumInput.Rows(0).Cells
            If cell.ReadOnly = False Then
                cell.Value = ""
            End If
        Next
        '実施日
        dateYmdBox.setADStr(todayStr)
        '実施者
        etcBox.Text = ""
        '所属機関
        companyBox.Text = ""
        '実施場所
        rbtnHouseIn.Checked = False
        rbtnHouseOut.Checked = False
        houseTextBox.Text = ""
        '過去の認定
        rbtnFirstCount.Checked = False
        rbtnSecondCount.Checked = False
        lastCertifiedCheckBox.Checked = False
        lastCertifiedYmdBox.setADStr(todayStr)
        '前回認定結果
        certifiedResultBox.Text = ""
        '性別
        rbtnMan.Checked = False
        rbtnWoman.Checked = False
        '生年月日
        birthYmdBox.setADStr(todayStr)
        ageLabel.Text = ""
        '現在所
        currentPostCode1.Text = ""
        currentPostCode2.Text = ""
        currentAddress.Text = ""
        currentTel1.Text = ""
        currentTel2.Text = ""
        currentTel3.Text = ""
        '家族等
        familyPostCode1.Text = ""
        familyPostCode2.Text = ""
        familyAddress.Text = ""
        familyTel1.Text = ""
        familyTel2.Text = ""
        familyTel3.Text = ""
        '氏名
        namBox.Text = ""
        '調査対象者との関係
        relationBox.Text = ""
        '（介護予防）訪問介護（ホームヘルプサービス）
        checkGen1.Checked = False
        txtNum1.Text = ""
        '（介護予防）訪問入浴介護
        checkGen2.Checked = False
        txtNum2.Text = ""
        '（介護予防）訪問看護
        checkGen3.Checked = False
        txtNum3.Text = ""
        '（介護予防）訪問リハビリテーション
        checkGen4.Checked = False
        txtNum4.Text = ""
        '（介護予防）居宅療養管理指導
        checkGen5.Checked = False
        txtNum5.Text = ""
        '（介護予防）通所介護（デイサービス）
        checkGen6.Checked = False
        txtNum6.Text = ""
        '（介護予防）通所リハビリテーション（デイケア）
        checkGen7.Checked = False
        txtNum7.Text = ""
        '（介護予防）短期入所生活介護（特養等）
        checkGen8.Checked = False
        txtNum8.Text = ""
        '（介護予防）短期入所療養介護（老健・診療所）
        checkGen9.Checked = False
        txtNum9.Text = ""
        '（介護予防）特定施設入居者生活介護
        checkGen10.Checked = False
        txtNum10.Text = ""
        '（介護予防）福祉用具貸与
        checkGen11.Checked = False
        txtNum11.Text = ""
        '特定（介護予防）福祉用具販売
        checkGen12.Checked = False
        txtNum12.Text = ""
        '住宅改修
        checkGen13.Checked = False
        CheckNum13Exists.Checked = False
        CheckNum13None.Checked = False
        '夜間対応型訪問介護
        checkGen14.Checked = False
        txtNum14.Text = ""
        '（介護予防）認知症対応型通所介護
        checkGen15.Checked = False
        txtNum15.Text = ""
        '（介護予防）小規模多機能型居宅介護
        checkGen16.Checked = False
        txtNum16.Text = ""
        '（介護予防）認知症対応型共同生活介護
        checkGen17.Checked = False
        txtNum17.Text = ""
        '地域密着型特定施設入居者生活介護
        checkGen18.Checked = False
        txtNum18.Text = ""
        '地域密着型介護老人福祉施設入所者生活介護
        checkGen19.Checked = False
        txtNum19.Text = ""
        '定期巡回・随時対応型訪問介護看護
        checkGen20.Checked = False
        txtNum20.Text = ""
        '複合型サービス
        checkGen23.Checked = False
        txtNum21.Text = ""
        '市町村特別給付
        checkGen21.Checked = False
        txtGentxt1.Text = ""
        '介護保険給付外の在宅サービス
        checkGen22.Checked = False
        txtGentxt2.Text = ""
        '利用施設
        '介護老人福祉施設
        checkStay1.Checked = False
        '介護老人保健施設
        checkStay2.Checked = False
        '介護療養型医療施設
        checkStay3.Checked = False
        '認知症対応型共同生活介護適用施設（ｸﾞﾙｰﾌﾟﾎｰﾑ）
        checkStay4.Checked = False
        '特定施設入所者生活介護適用施設（ｹｱﾊｳｽ等）
        checkStay5.Checked = False
        '医療機関（医療保険適用療養病床）
        checkStay6.Checked = False
        '医療機関（療養病床以外）
        checkStay7.Checked = False
        'その他の施設
        checkStay8.Checked = False
        '施設連絡先
        facilityNameBox.Text = ""
        facilityPostCode1.Text = ""
        facilityPostCode2.Text = ""
        facilityAddress.Text = ""
        facilityTel1.Text = ""
        facilityTel2.Text = ""
        facilityTel3.Text = ""
        '特記テキスト
        spText1.Text = ""
        spText2.Text = ""
        spText3.Text = ""
        spText4.Text = ""

    End Sub

    Private Sub clearAllInputData()
        '概況調査タブの内容クリア
        clearOverviewPageInputBox()

        '特記事項タブの内容クリア
        SpDgv1.clearText()
        SpDgv2.clearText()
        SpDgv3.clearText()
        SpDgv4.clearText()
        SpDgv5.clearText()
        SpDgv6.clearText()
        SpDgv7.clearText()

        '基本調査タブの内容クリア
        For Each tp As TabPage In bsTab.TabPages
            For Each c As Control In tp.Controls
                If TypeOf c Is GroupBox Then
                    For Each ex As Control In c.Controls
                        If TypeOf ex Is ExCheckBox Then
                            DirectCast(ex, ExCheckBox).Checked = False
                        ElseIf TypeOf ex Is ExRadioButton Then
                            DirectCast(ex, ExRadioButton).Checked = False
                        End If
                    Next
                End If
            Next
        Next

    End Sub

    Private Sub displayUserData(nam As String, kana As String, ymd1 As String)
        clearAllInputData()
        Dim cnn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim sql As String = "select * from Auth1 where Nam='" & nam & "' and Ymd1='" & ymd1 & "'"
        cnn.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        cnn.Open(topForm.DB_KSave2)
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

        '概況調査タブの表示処理
        rs.Filter = "Gyo=61"
        '調査日
        For i = 1 To 6
            dgvNumInput("GDay" & i, 0).Value = Util.checkDBNullValue(rs.Fields("GDay" & i).Value)
        Next
        '被保険者番号
        For i = 1 To 10
            dgvNumInput("GNum" & i, 0).Value = Util.checkDBNullValue(rs.Fields("GNum" & i).Value)
        Next
        dateYmdBox.setADStr(Util.checkDBNullValue(rs.Fields("Ymd1").Value)) '実施日
        etcBox.Text = Util.checkDBNullValue(rs.Fields("Tanto").Value) '実施者
        companyBox.Text = Util.checkDBNullValue(rs.Fields("Kikan").Value) '所属機関
        '実施場所
        If Util.checkDBNullValue(rs.Fields("Home").Value) = "0" Then
            rbtnHouseIn.Checked = True '自宅内
        ElseIf Util.checkDBNullValue(rs.Fields("Home").Value) = "1" Then
            rbtnHouseOut.Checked = True '自宅外
        End If
        houseTextBox.Text = Util.checkDBNullValue(rs.Fields("Nonhm").Value) '自宅外の詳細
        '過去の認定
        If Util.checkDBNullValue(rs.Fields("Kako").Value) = "0" Then
            rbtnFirstCount.Checked = True '初回
        ElseIf Util.checkDBNullValue(rs.Fields("Kako").Value) = "1" Then
            rbtnSecondCount.Checked = True '2回目以降
        End If
        '前回認定
        If Util.checkDBNullValue(rs.Fields("Ymd2").Value) <> "" Then
            lastCertifiedCheckBox.Checked = True
            lastCertifiedYmdBox.setADStr(Util.checkDBNullValue(rs.Fields("Ymd2").Value))
        End If
        certifiedResultBox.Text = If(Util.checkDBNullValue(rs.Fields("Kai").Value) = "", "", certifiedResultBox.Items.Item(rs.Fields("Kai").Value)) '前回認定結果
        If Util.checkDBNullValue(rs.Fields("Sex").Value) = "0" Then
            rbtnMan.Checked = True '男
        ElseIf Util.checkDBNullValue(rs.Fields("Sex").Value) = "1" Then
            rbtnWoman.Checked = True '女
        End If
        birthYmdBox.setADStr(Util.checkDBNullValue(rs.Fields("Ymd3").Value)) '生年月日
        ageLabel.Text = Util.checkDBNullValue(rs.Fields("Age").Value) '年齢
        '現在所
        currentPostCode1.Text = Util.checkDBNullValue(rs.Fields("Pn11").Value)
        currentPostCode2.Text = Util.checkDBNullValue(rs.Fields("Pn12").Value)
        currentAddress.Text = Util.checkDBNullValue(rs.Fields("Ad1").Value)
        currentTel1.Text = Util.checkDBNullValue(rs.Fields("Tel11").Value)
        currentTel2.Text = Util.checkDBNullValue(rs.Fields("Tel12").Value)
        currentTel3.Text = Util.checkDBNullValue(rs.Fields("Tel13").Value)
        '家族等
        familyPostCode1.Text = Util.checkDBNullValue(rs.Fields("Pn21").Value)
        familyPostCode2.Text = Util.checkDBNullValue(rs.Fields("Pn22").Value)
        familyAddress.Text = Util.checkDBNullValue(rs.Fields("Ad2").Value)
        familyTel1.Text = Util.checkDBNullValue(rs.Fields("Tel21").Value)
        familyTel2.Text = Util.checkDBNullValue(rs.Fields("Tel22").Value)
        familyTel3.Text = Util.checkDBNullValue(rs.Fields("Tel23").Value)
        'Ⅲ
        namBox.Text = Util.checkDBNullValue(rs.Fields("Fa").Value) '氏名
        relationBox.Text = Util.checkDBNullValue(rs.Fields("Far").Value) '調査対象者との関係
        For i = 1 To 20
            'Gen1～20,Num1～20部分
            If Util.checkDBNullValue(rs.Fields("Gen" & i).Value) = "1" Then
                CType(overview3Panel.Controls("checkGen" & i), CheckBox).Checked = True
            End If
            If i <> 13 Then
                CType(overview3Panel.Controls("txtNum" & i), ExTextBox).Text = Util.checkDBNullValue(rs.Fields("Num" & i).Value)
            Else
                If Util.checkDBNullValue(rs.Fields("Num" & i).Value) = "1" Then
                    CheckNum13Exists.Checked = True
                ElseIf Util.checkDBNullValue(rs.Fields("Num" & i).Value) = "2" Then
                    CheckNum13None.Checked = True
                End If
            End If
        Next
        '複合型サービス
        If Util.checkDBNullValue(rs.Fields("Gen23").Value) = "1" Then
            checkGen23.Checked = True
        End If
        txtNum21.Text = Util.checkDBNullValue(rs.Fields("Num21").Value)
        '市町村特別給付
        If Util.checkDBNullValue(rs.Fields("Gen21").Value) = "1" Then
            checkGen21.Checked = True
        End If
        txtGentxt1.Text = Util.checkDBNullValue(rs.Fields("Gentxt1").Value)
        '介護保険給付外の在宅サービス
        If Util.checkDBNullValue(rs.Fields("Gen22").Value) = "1" Then
            checkGen22.Checked = True
        End If
        txtGentxt2.Text = Util.checkDBNullValue(rs.Fields("Gentxt2").Value)
        '利用施設
        For i = 1 To 8
            If Util.checkDBNullValue(rs.Fields("Stay" & i).Value) = "1" Then
                CType(facilityPanel.Controls("checkStay" & i), CheckBox).Checked = True
            End If
        Next
        '施設連絡先
        facilityNameBox.Text = Util.checkDBNullValue(rs.Fields("Name").Value) '連絡先
        facilityPostCode1.Text = Util.checkDBNullValue(rs.Fields("Pn31").Value)
        facilityPostCode2.Text = Util.checkDBNullValue(rs.Fields("Pn32").Value)
        facilityAddress.Text = Util.checkDBNullValue(rs.Fields("Ad3").Value)
        facilityTel1.Text = Util.checkDBNullValue(rs.Fields("Tel31").Value)
        facilityTel2.Text = Util.checkDBNullValue(rs.Fields("Tel32").Value)
        facilityTel3.Text = Util.checkDBNullValue(rs.Fields("Tel33").Value)
        'Ⅳ
        spText1.Text = Util.checkDBNullValue(rs.Fields("GTokki1").Value)
        spText2.Text = Util.checkDBNullValue(rs.Fields("GTokki2").Value)
        spText3.Text = Util.checkDBNullValue(rs.Fields("GTokki3").Value)
        spText4.Text = Util.checkDBNullValue(rs.Fields("GTokki4").Value)

        '特記事項タブの表示処理
        '1.身体機能・起居動作
        rs.Filter = "Sp=0 and Gyo>=4 and Gyo<>61"
        rs.Sort = "Gyo ASC"
        If rs.RecordCount >= 1 Then
            rs.MoveFirst()
            Dim i As Integer = 0
            While Not rs.EOF
                SpDgv1("Crr", i).Value = Util.checkDBNullValue(rs.Fields("Crr").Value)
                SpDgv1("Txt", i).Value = Util.checkDBNullValue(rs.Fields("Txt").Value)
                i += 1
                rs.MoveNext()
            End While
        End If
        '2.生活機能
        rs.Filter = "Sp=1 and Gyo>=5"
        rs.Sort = "Gyo ASC"
        If rs.RecordCount >= 1 Then
            rs.MoveFirst()
            Dim i As Integer = 0
            While Not rs.EOF
                SpDgv2("Crr", i).Value = Util.checkDBNullValue(rs.Fields("Crr").Value)
                SpDgv2("Txt", i).Value = Util.checkDBNullValue(rs.Fields("Txt").Value)
                i += 1
                rs.MoveNext()
            End While
        End If
        '3.認知機能
        rs.Filter = "Sp=2 and Gyo>=5"
        rs.Sort = "Gyo ASC"
        If rs.RecordCount >= 1 Then
            rs.MoveFirst()
            Dim i As Integer = 0
            While Not rs.EOF
                SpDgv3("Crr", i).Value = Util.checkDBNullValue(rs.Fields("Crr").Value)
                SpDgv3("Txt", i).Value = Util.checkDBNullValue(rs.Fields("Txt").Value)
                i += 1
                rs.MoveNext()
            End While
        End If
        '4.精神・行動障害
        rs.Filter = "Sp=3 and Gyo>=6"
        rs.Sort = "Gyo ASC"
        If rs.RecordCount >= 1 Then
            rs.MoveFirst()
            Dim i As Integer = 0
            While Not rs.EOF
                SpDgv4("Crr", i).Value = Util.checkDBNullValue(rs.Fields("Crr").Value)
                SpDgv4("Txt", i).Value = Util.checkDBNullValue(rs.Fields("Txt").Value)
                i += 1
                rs.MoveNext()
            End While
        End If
        '5.社会生活への適応
        rs.Filter = "Sp=4 and Gyo>=5"
        rs.Sort = "Gyo ASC"
        If rs.RecordCount >= 1 Then
            rs.MoveFirst()
            Dim i As Integer = 0
            While Not rs.EOF
                SpDgv5("Crr", i).Value = Util.checkDBNullValue(rs.Fields("Crr").Value)
                SpDgv5("Txt", i).Value = Util.checkDBNullValue(rs.Fields("Txt").Value)
                i += 1
                rs.MoveNext()
            End While
        End If
        '6.特別な医療
        rs.Filter = "Sp=5 and Gyo>=4"
        rs.Sort = "Gyo ASC"
        If rs.RecordCount >= 1 Then
            rs.MoveFirst()
            Dim i As Integer = 0
            While Not rs.EOF
                SpDgv6("Crr", i).Value = Util.checkDBNullValue(rs.Fields("Crr").Value)
                SpDgv6("Txt", i).Value = Util.checkDBNullValue(rs.Fields("Txt").Value)
                i += 1
                rs.MoveNext()
            End While
        End If
        '7.日常生活自立度
        rs.Filter = "Sp=6 and Gyo>=4"
        rs.Sort = "Gyo ASC"
        If rs.RecordCount >= 1 Then
            rs.MoveFirst()
            Dim i As Integer = 0
            While Not rs.EOF
                SpDgv7("Crr", i).Value = Util.checkDBNullValue(rs.Fields("Crr").Value)
                SpDgv7("Txt", i).Value = Util.checkDBNullValue(rs.Fields("Txt").Value)
                i += 1
                rs.MoveNext()
            End While
        End If
        rs.Close()

        '基本調査タブの表示処理
        sql = "select * from Auth2 where Nam='" & nam & "' and Ymd='" & ymd1 & "' order by Gyo"
        rs = New ADODB.Recordset
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

        '1.身体機能・起居動作
        '1-1 麻痺等の有無
        rs.Find("Gyo=0")
        Ch2_1.Checked = If(Util.checkDBNullValue(rs.Fields("Ch2").Value) = 1, True, False)
        rs.Find("Gyo=1")
        Ch2_2.Checked = If(Util.checkDBNullValue(rs.Fields("Ch2").Value) = 1, True, False)
        rs.Find("Gyo=2")
        Ch2_3.Checked = If(Util.checkDBNullValue(rs.Fields("Ch2").Value) = 1, True, False)
        rs.Find("Gyo=3")
        Ch2_4.Checked = If(Util.checkDBNullValue(rs.Fields("Ch2").Value) = 1, True, False)
        rs.Find("Gyo=4")
        Ch2_5.Checked = If(Util.checkDBNullValue(rs.Fields("Ch2").Value) = 1, True, False)
        rs.Find("Gyo=5")
        Ch2_6.Checked = If(Util.checkDBNullValue(rs.Fields("Ch2").Value) = 1, True, False)
        '1-2 拘縮の有無
        rs.Find("Gyo=0", , ADODB.SearchDirectionEnum.adSearchBackward)
        Ch3_1.Checked = If(Util.checkDBNullValue(rs.Fields("Ch3").Value) = 1, True, False)
        rs.Find("Gyo=1")
        Ch3_2.Checked = If(Util.checkDBNullValue(rs.Fields("Ch3").Value) = 1, True, False)
        rs.Find("Gyo=2")
        Ch3_3.Checked = If(Util.checkDBNullValue(rs.Fields("Ch3").Value) = 1, True, False)
        rs.Find("Gyo=3")
        Ch3_4.Checked = If(Util.checkDBNullValue(rs.Fields("Ch3").Value) = 1, True, False)
        rs.Find("Gyo=4")
        Ch3_5.Checked = If(Util.checkDBNullValue(rs.Fields("Ch3").Value) = 1, True, False)
        '1-3 寝返り
        rs.Find("Gyo=0", , ADODB.SearchDirectionEnum.adSearchBackward)
        rb1_3_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=1")
        rb1_3_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=2")
        rb1_3_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '1-4 起き上がり
        rs.Find("Gyo=3")
        rb1_4_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=4")
        rb1_4_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=5")
        rb1_4_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '1-5 座位保持
        rs.Find("Gyo=6")
        rb1_5_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=7")
        rb1_5_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=8")
        rb1_5_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=9")
        rb1_5_4.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '1-6 両足での立位保持
        rs.Find("Gyo=10")
        rb1_6_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=11")
        rb1_6_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=12")
        rb1_6_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '1-7 歩行
        rs.Find("Gyo=13")
        rb1_7_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=14")
        rb1_7_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=15")
        rb1_7_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '1-8 立ち上がり
        rs.Find("Gyo=16")
        rb1_8_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=17")
        rb1_8_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=18")
        rb1_8_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '1-9 片足での立位保持
        rs.Find("Gyo=19")
        rb1_9_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=20")
        rb1_9_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=21")
        rb1_9_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '1-10 洗身
        rs.Find("Gyo=22")
        rb1_10_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=23")
        rb1_10_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=24")
        rb1_10_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=25")
        rb1_10_4.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '1-11 つめ切り
        rs.Find("Gyo=26")
        rb1_11_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=27")
        rb1_11_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=28")
        rb1_11_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '1-12 視力について
        rs.Find("Gyo=29")
        rb1_12_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=30")
        rb1_12_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=31")
        rb1_12_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=32")
        rb1_12_4.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=33")
        rb1_12_5.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '1-13 聴力について
        rs.Find("Gyo=34")
        rb1_13_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=35")
        rb1_13_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=36")
        rb1_13_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=37")
        rb1_13_4.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=38")
        rb1_13_5.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)

        '2.生活機能
        '2-1 移乗
        rs.Find("Gyo=39")
        rb2_1_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=40")
        rb2_1_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=41")
        rb2_1_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=42")
        rb2_1_4.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '2-2 移動
        rs.Find("Gyo=43")
        rb2_2_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=44")
        rb2_2_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=45")
        rb2_2_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=46")
        rb2_2_4.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '2-3 えん下
        rs.Find("Gyo=47")
        rb2_3_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=48")
        rb2_3_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=49")
        rb2_3_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '2-4 食事摂取
        rs.Find("Gyo=50")
        rb2_4_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=51")
        rb2_4_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=52")
        rb2_4_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=53")
        rb2_4_4.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '2-5 排尿
        rs.Find("Gyo=54")
        rb2_5_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=55")
        rb2_5_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=56")
        rb2_5_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=57")
        rb2_5_4.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '2-6 排便
        rs.Find("Gyo=58")
        rb2_6_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=59")
        rb2_6_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=60")
        rb2_6_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=61")
        rb2_6_4.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '2-7 口腔清潔
        rs.Find("Gyo=62")
        rb2_7_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=63")
        rb2_7_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=64")
        rb2_7_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '2-8 洗顔
        rs.Find("Gyo=65")
        rb2_8_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=66")
        rb2_8_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=67")
        rb2_8_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '2-9 整髪
        rs.Find("Gyo=68")
        rb2_9_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=69")
        rb2_9_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=70")
        rb2_9_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '2-10 上衣の着脱
        rs.Find("Gyo=71")
        rb2_10_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=72")
        rb2_10_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=73")
        rb2_10_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=74")
        rb2_10_4.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '2-11 ズボン等の着脱
        rs.Find("Gyo=75")
        rb2_11_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=76")
        rb2_11_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=77")
        rb2_11_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=78")
        rb2_11_4.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '2-12 外出頻度
        rs.Find("Gyo=79")
        rb2_12_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=80")
        rb2_12_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=81")
        rb2_12_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)

        '3.認知機能
        '3-1 意思の伝達
        rs.Find("Gyo=82")
        rb3_1_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=83")
        rb3_1_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=84")
        rb3_1_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=85")
        rb3_1_4.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '3-2 毎日の日課を理解
        rs.Find("Gyo=86")
        rb3_2_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=87")
        rb3_2_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '3-3 生年月日や年齢を言う
        rs.Find("Gyo=88")
        rb3_3_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=89")
        rb3_3_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '3-4 短期記憶（面接調査の直前に何をしていたのか思い出す）
        rs.Find("Gyo=90")
        rb3_4_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=91")
        rb3_4_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '3-5 自分の名前を言う
        rs.Find("Gyo=92")
        rb3_5_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=93")
        rb3_5_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '3-6 今の季節を理解
        rs.Find("Gyo=94")
        rb3_6_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=95")
        rb3_6_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '3-7 場所の理解（自分がいる場所を答える）
        rs.Find("Gyo=96")
        rb3_7_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=97")
        rb3_7_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '3-8 徘徊
        rs.Find("Gyo=98")
        rb3_8_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=99")
        rb3_8_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=100")
        rb3_8_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '3-9 外出すると戻れない
        rs.Find("Gyo=101")
        rb3_9_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=102")
        rb3_9_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=103")
        rb3_9_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)

        '4.精神・行動障害
        '4-1 物を盗られたなどと被害的になる
        rs.Find("Gyo=104")
        rb4_1_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=105")
        rb4_1_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=106")
        rb4_1_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '4-2 作話をすること
        rs.Find("Gyo=107")
        rb4_2_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=108")
        rb4_2_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=109")
        rb4_2_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '4-3 泣いたり、笑ったりして感情が不安定になる
        rs.Find("Gyo=110")
        rb4_3_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=111")
        rb4_3_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=112")
        rb4_3_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '4-4 昼夜の逆転
        rs.Find("Gyo=113")
        rb4_4_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=114")
        rb4_4_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=115")
        rb4_4_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '4-5 しつこく同じ話をする
        rs.Find("Gyo=116")
        rb4_5_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=117")
        rb4_5_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=118")
        rb4_5_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '4-6 大声を出す
        rs.Find("Gyo=119")
        rb4_6_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=120")
        rb4_6_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=121")
        rb4_6_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '4-7 介護に抵抗する
        rs.Find("Gyo=122")
        rb4_7_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=123")
        rb4_7_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=124")
        rb4_7_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '4-8 「家に帰る」等と言い落ち着きがない
        rs.Find("Gyo=125")
        rb4_8_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=126")
        rb4_8_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=127")
        rb4_8_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '4-9 一人で外に出たがり目が離せない
        rs.Find("Gyo=128")
        rb4_9_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=129")
        rb4_9_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=130")
        rb4_9_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '4-10 いろいろなものを集めたり、無断でもってくる
        rs.Find("Gyo=131")
        rb4_10_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=132")
        rb4_10_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=133")
        rb4_10_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '4-11 物を壊したり、衣類を破いたりする
        rs.Find("Gyo=134")
        rb4_11_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=135")
        rb4_11_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=136")
        rb4_11_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '4-12 ひどい物忘れ
        rs.Find("Gyo=137")
        rb4_12_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=138")
        rb4_12_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=139")
        rb4_12_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '4-13 意味もなく独り言や独り笑いをする
        rs.Find("Gyo=140")
        rb4_13_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=141")
        rb4_13_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=142")
        rb4_13_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '4-14 自分勝手に行動する
        rs.Find("Gyo=143")
        rb4_14_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=144")
        rb4_14_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=145")
        rb4_14_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '4-15 話がまとまらず、会話にならない
        rs.Find("Gyo=146")
        rb4_15_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=147")
        rb4_15_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=148")
        rb4_15_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)

        '5.社会生活への適応
        '5-1 薬の内服
        rs.Find("Gyo=149")
        rb5_1_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=150")
        rb5_1_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=151")
        rb5_1_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '5-2 金銭の管理
        rs.Find("Gyo=152")
        rb5_2_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=153")
        rb5_2_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=154")
        rb5_2_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '5-3 日常の意思決定
        rs.Find("Gyo=155")
        rb5_3_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=156")
        rb5_3_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=157")
        rb5_3_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=158")
        rb5_3_4.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '5-4 集団への不適応
        rs.Find("Gyo=159")
        rb5_4_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=160")
        rb5_4_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=161")
        rb5_4_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '5-5 買い物
        rs.Find("Gyo=162")
        rb5_5_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=163")
        rb5_5_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=164")
        rb5_5_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=165")
        rb5_5_4.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '5-6 簡単な調理
        rs.Find("Gyo=166")
        rb5_6_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=167")
        rb5_6_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=168")
        rb5_6_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=169")
        rb5_6_4.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)

        '6.特別な医療
        '点滴の管理
        rs.Find("Gyo=0", , ADODB.SearchDirectionEnum.adSearchBackward)
        Ch4_1.Checked = If(Util.checkDBNullValue(rs.Fields("Ch4").Value) = 1, True, False)
        '中心静脈栄養
        rs.Find("Gyo=1")
        Ch4_2.Checked = If(Util.checkDBNullValue(rs.Fields("Ch4").Value) = 1, True, False)
        '透析
        rs.Find("Gyo=2")
        Ch4_3.Checked = If(Util.checkDBNullValue(rs.Fields("Ch4").Value) = 1, True, False)
        'ストーマ（人工肛門）の処置
        rs.Find("Gyo=3")
        Ch4_4.Checked = If(Util.checkDBNullValue(rs.Fields("Ch4").Value) = 1, True, False)
        '酸素療法
        rs.Find("Gyo=4")
        Ch4_5.Checked = If(Util.checkDBNullValue(rs.Fields("Ch4").Value) = 1, True, False)
        'レスピレーター（人工呼吸器）
        rs.Find("Gyo=5")
        Ch4_6.Checked = If(Util.checkDBNullValue(rs.Fields("Ch4").Value) = 1, True, False)
        '気管切開の処置
        rs.Find("Gyo=6")
        Ch4_7.Checked = If(Util.checkDBNullValue(rs.Fields("Ch4").Value) = 1, True, False)
        '疼痛の看護
        rs.Find("Gyo=7")
        Ch4_8.Checked = If(Util.checkDBNullValue(rs.Fields("Ch4").Value) = 1, True, False)
        '経管栄養
        rs.Find("Gyo=8")
        Ch4_9.Checked = If(Util.checkDBNullValue(rs.Fields("Ch4").Value) = 1, True, False)
        'モニター測定（血圧、心拍、酸素飽和度等）
        rs.Find("Gyo=9")
        Ch4_10.Checked = If(Util.checkDBNullValue(rs.Fields("Ch4").Value) = 1, True, False)
        'じょくそうの処置
        rs.Find("Gyo=10")
        Ch4_11.Checked = If(Util.checkDBNullValue(rs.Fields("Ch4").Value) = 1, True, False)
        'カテーテル（コンドームカテーテル、留置カテーテル、ウロストーマ等）
        rs.Find("Gyo=11")
        Ch4_12.Checked = If(Util.checkDBNullValue(rs.Fields("Ch4").Value) = 1, True, False)

        '7.日常生活自立度
        '障害高齢者の日常生活自立度（寝たきり度）
        rs.Find("Gyo=170")
        rb7_1_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=171")
        rb7_1_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=172")
        rb7_1_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=173")
        rb7_1_4.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=174")
        rb7_1_5.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=175")
        rb7_1_6.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=176")
        rb7_1_7.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=177")
        rb7_1_8.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=178")
        rb7_1_9.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        '認知症高齢者の日常生活自立度
        rs.Find("Gyo=179")
        rb7_2_1.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=180")
        rb7_2_2.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=181")
        rb7_2_3.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=182")
        rb7_2_4.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=183")
        rb7_2_5.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=184")
        rb7_2_6.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=185")
        rb7_2_7.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)
        rs.Find("Gyo=186")
        rb7_2_8.Checked = If(Util.checkDBNullValue(rs.Fields("Opt4").Value) = 1, True, False)

        rs.Close()
        cnn.Close()
    End Sub

    Private Sub lastCertifiedCheckBox_CheckedChanged(sender As Object, e As System.EventArgs) Handles lastCertifiedCheckBox.CheckedChanged
        If lastCertifiedCheckBox.Checked = True Then
            lastCertifiedYmdBox.Visible = True
            lastCertifiedYmdBox.setADStr(Today.ToString("yyyy/MM/dd"))
        Else
            lastCertifiedYmdBox.Visible = False
        End If
    End Sub

    Private Sub btnClear_Click(sender As System.Object, e As System.EventArgs) Handles btnClear.Click
        clearOverviewPageInputBox()
    End Sub

    Private Sub btnCalcAge_Click(sender As System.Object, e As System.EventArgs) Handles btnCalcAge.Click
        Dim doDate As DateTime = New DateTime(CInt(dateYmdBox.getADStr().Substring(0, 4)), CInt(dateYmdBox.getADStr().Substring(5, 2)), CInt(dateYmdBox.getADStr().Substring(8, 2)))
        Dim birthDate As DateTime = New DateTime(CInt(birthYmdBox.getADStr().Substring(0, 4)), CInt(birthYmdBox.getADStr().Substring(5, 2)), CInt(birthYmdBox.getADStr().Substring(8, 2)))
        Dim age As Integer = doDate.Year - birthDate.Year
        '誕生日がまだ来ていなければ、1引く
        If doDate.Month < birthDate.Month OrElse (doDate.Month = birthDate.Month AndAlso doDate.Day < birthDate.Day) Then
            age -= 1
        End If
        ageLabel.Text = age
    End Sub

    Private Sub txtGentxt1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtGentxt1.KeyDown
        If e.KeyCode = Keys.Enter Then
            txtGentxt2.Focus()
        End If
    End Sub

    Private Sub txtGentxt2_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtGentxt2.KeyDown
        If e.KeyCode = Keys.Enter Then
            facilityNameBox.Focus()
        End If
    End Sub

    Private Sub spText1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles spText1.KeyDown
        If e.KeyCode = Keys.Down Then
            spText2.Focus()
        End If
    End Sub

    Private Sub spText2_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles spText2.KeyDown
        If e.KeyCode = Keys.Down Then
            spText3.Focus()
        ElseIf e.KeyCode = Keys.Up Then
            spText1.Focus()
        End If
    End Sub

    Private Sub spText3_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles spText3.KeyDown
        If e.KeyCode = Keys.Down Then
            spText4.Focus()
        ElseIf e.KeyCode = Keys.Up Then
            spText2.Focus()
        End If
    End Sub

    Private Sub spText4_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles spText4.KeyDown
        If e.KeyCode = Keys.Up Then
            spText3.Focus()
        ElseIf e.KeyCode = Keys.Enter Then
            btnRegist.Focus()
        End If
    End Sub

    Private Sub spTabBtnClear_Click(sender As System.Object, e As System.EventArgs) Handles btnClear1.Click, btnClear2.Click, btnClear3.Click, btnClear4.Click, btnClear5.Click, btnClear6.Click, btnClear7.Click
        Dim b As Button = CType(sender, Button)
        Dim tp As TabPage = b.Parent
        Dim num As String = b.Name.Substring(b.Name.Length - 1)
        CType(tp.Controls("SpDgv" & num), SpDgv).clearText()
    End Sub

    Private Sub spTabBtnRowInsert_Click(sender As System.Object, e As System.EventArgs) Handles btnRowInsert1.Click, btnRowInsert2.Click, btnRowInsert3.Click, btnRowInsert4.Click, btnRowInsert5.Click, btnRowInsert6.Click, btnRowInsert7.Click
        Dim b As Button = CType(sender, Button)
        Dim tp As TabPage = b.Parent
        Dim num As String = b.Name.Substring(b.Name.Length - 1)
        CType(tp.Controls("SpDgv" & num), SpDgv).rowInsert()
    End Sub

    Private Sub spTabBtnRowDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnRowDelete1.Click, btnRowDelete2.Click, btnRowDelete3.Click, btnRowDelete4.Click, btnRowDelete5.Click, btnRowDelete6.Click, btnRowDelete7.Click
        Dim b As Button = CType(sender, Button)
        Dim tp As TabPage = b.Parent
        Dim num As String = b.Name.Substring(b.Name.Length - 1)
        CType(tp.Controls("SpDgv" & num), SpDgv).rowDelete()
    End Sub

    Private Sub recordList_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles recordList.SelectedIndexChanged
        If Not IsNothing(recordList.SelectedItem) Then
            displayUserData(userLabel.Text, kanaLabel.Text, convWarekiStrToADStr(recordList.SelectedItem.ToString()))
        End If
    End Sub

    Private Sub createSpDgvTabRecord(rs As ADODB.Recordset, tabIndex As Integer, dgv As SpDgv, txtArray As String(), userName As String, userKana As String, registYmd As String)
        Dim gyoIndex As Integer = 0
        For gyoIndex = 0 To txtArray.Length - 1
            rs.AddNew()
            rs.Fields("Nam").Value = userName
            rs.Fields("Kana").Value = userKana
            rs.Fields("Gyo").Value = 1 + gyoIndex
            rs.Fields("Ymd1").Value = registYmd
            rs.Fields("Sp").Value = tabIndex
            rs.Fields("Txt2").Value = txtArray(gyoIndex)
        Next
        For i = 0 To dgv.Rows.Count - 1
            If Util.checkDBNullValue(dgv("Txt", i).Value) = "" Then
                Exit For
            Else
                rs.AddNew()
                rs.Fields("Nam").Value = userName
                rs.Fields("Kana").Value = userKana
                rs.Fields("Gyo").Value = 2 + gyoIndex + i
                rs.Fields("Ymd1").Value = registYmd
                rs.Fields("Sp").Value = tabIndex
                rs.Fields("Crr").Value = dgv("Crr", i).Value
                rs.Fields("Txt").Value = dgv("Txt", i).Value
            End If
        Next
    End Sub

    Private Sub createBSTabRecord(rs As ADODB.Recordset, userName As String, registYmd As String, gyo As Integer, rbOpt4 As ExRadioButton, Optional ch2 As ExCheckBox = Nothing, Optional ch3 As ExCheckBox = Nothing, Optional ch4 As ExCheckBox = Nothing)
        rs.AddNew()
        rs.Fields("Nam").Value = userName
        rs.Fields("Ymd").Value = registYmd
        rs.Fields("Gyo").Value = gyo
        rs.Fields("Opt4").Value = If(rbOpt4.Checked, 1, 0)
        If Not IsNothing(ch2) Then
            rs.Fields("Ch2").Value = If(ch2.Checked, "1", "0")
        End If
        If Not IsNothing(ch3) Then
            rs.Fields("Ch3").Value = If(ch3.Checked, "1", "0")
        End If
        If Not IsNothing(ch4) Then
            rs.Fields("Ch4").Value = If(ch4.Checked, "1", "0")
        End If
    End Sub

    Private Sub btnRegist_Click(sender As System.Object, e As System.EventArgs) Handles btnRegist.Click
        Dim userName As String = userLabel.Text '利用者名漢字
        Dim userKana As String = kanaLabel.Text '利用者名カナ
        If userName = "" OrElse userKana = "" Then
            MsgBox("利用者を選択して下さい。")
            Return
        End If

        '概況調査タブ
        '調査日(GDay)
        Dim gDay1 As String = Util.checkDBNullValue(dgvNumInput("GDay1", 0).Value)
        Dim gDay2 As String = Util.checkDBNullValue(dgvNumInput("GDay2", 0).Value)
        Dim gDay3 As String = Util.checkDBNullValue(dgvNumInput("GDay3", 0).Value)
        Dim gDay4 As String = Util.checkDBNullValue(dgvNumInput("GDay4", 0).Value)
        Dim gDay5 As String = Util.checkDBNullValue(dgvNumInput("GDay5", 0).Value)
        Dim gDay6 As String = Util.checkDBNullValue(dgvNumInput("GDay6", 0).Value)
        '被保険者番号(GNum)
        Dim gNum1 As String = Util.checkDBNullValue(dgvNumInput("GNum1", 0).Value)
        Dim gNum2 As String = Util.checkDBNullValue(dgvNumInput("GNum2", 0).Value)
        Dim gNum3 As String = Util.checkDBNullValue(dgvNumInput("GNum3", 0).Value)
        Dim gNum4 As String = Util.checkDBNullValue(dgvNumInput("GNum4", 0).Value)
        Dim gNum5 As String = Util.checkDBNullValue(dgvNumInput("GNum5", 0).Value)
        Dim gNum6 As String = Util.checkDBNullValue(dgvNumInput("GNum6", 0).Value)
        Dim gNum7 As String = Util.checkDBNullValue(dgvNumInput("GNum7", 0).Value)
        Dim gNum8 As String = Util.checkDBNullValue(dgvNumInput("GNum8", 0).Value)
        Dim gNum9 As String = Util.checkDBNullValue(dgvNumInput("GNum9", 0).Value)
        Dim gNum10 As String = Util.checkDBNullValue(dgvNumInput("GNum10", 0).Value)
        '実施日
        Dim ymd1 As String = dateYmdBox.getADStr()
        '実施者
        Dim tanto As String = etcBox.Text
        '所属機関
        Dim kikan As String = companyBox.Text
        '実施場所
        Dim home As String = If(rbtnHouseIn.Checked, "0", If(rbtnHouseOut.Checked, "1", ""))
        Dim nonHm As String = houseTextBox.Text

        '過去の認定
        Dim kako As String = If(rbtnFirstCount.Checked, "0", If(rbtnSecondCount.Checked, "1", ""))
        Dim ymd2 As String = If(lastCertifiedCheckBox.Checked, lastCertifiedYmdBox.getADStr(), "")
        '前回認定結果
        Dim kai As String = If(certifiedResultBox.FindStringExact(certifiedResultBox.Text) <> -1, certifiedResultBox.FindString(certifiedResultBox.Text), "")
        '性別
        Dim sex As String = If(rbtnMan.Checked, "0", If(rbtnWoman.Checked, "1", ""))
        '生年月日
        Dim ymd3 As String = birthYmdBox.getADStr()
        Dim age As String = ageLabel.Text
        '現在所
        Dim pn11 As String = currentPostCode1.Text '〒
        Dim pn12 As String = currentPostCode2.Text '〒
        Dim ad1 As String = currentAddress.Text '住所
        Dim tel11 As String = currentTel1.Text '電話
        Dim tel12 As String = currentTel2.Text '電話
        Dim tel13 As String = currentTel3.Text '電話
        '家族等
        Dim pn21 As String = familyPostCode1.Text '〒
        Dim pn22 As String = familyPostCode2.Text '〒
        Dim ad2 As String = familyAddress.Text '住所
        Dim tel21 As String = familyTel1.Text '電話
        Dim tel22 As String = familyTel2.Text '電話
        Dim tel23 As String = familyTel3.Text '電話
        '氏名
        Dim fa As String = namBox.Text
        '調査対象者との関係
        Dim far As String = relationBox.Text

        'Ⅲ
        '(介護予防)訪問介護(ﾎｰﾑﾍﾙﾌﾟｻｰﾋﾞｽ)
        Dim gen1 As String = If(checkGen1.Checked, "1", "0")
        Dim num1 As String = txtNum1.Text
        '(介護予防)訪問入浴介護
        Dim gen2 As String = If(checkGen2.Checked, "1", "0")
        Dim num2 As String = txtNum2.Text
        '(介護予防)訪問看護
        Dim gen3 As String = If(checkGen3.Checked, "1", "0")
        Dim num3 As String = txtNum3.Text
        '(介護予防)訪問ﾘﾊﾋﾞﾘﾃｰｼｮﾝ
        Dim gen4 As String = If(checkGen4.Checked, "1", "0")
        Dim num4 As String = txtNum4.Text
        '(介護予防)居宅療養管理指導
        Dim gen5 As String = If(checkGen5.Checked, "1", "0")
        Dim num5 As String = txtNum5.Text
        '(介護予防)通所介護(ﾃﾞｲｻｰﾋﾞｽ)
        Dim gen6 As String = If(checkGen6.Checked, "1", "0")
        Dim num6 As String = txtNum6.Text
        '(介護予防)通所ﾘﾊﾋﾞﾘﾃｰｼｮﾝ(ﾃﾞｲｹｱ)
        Dim gen7 As String = If(checkGen7.Checked, "1", "0")
        Dim num7 As String = txtNum7.Text
        '(介護予防)短期入所生活介護(特養等)
        Dim gen8 As String = If(checkGen8.Checked, "1", "0")
        Dim num8 As String = txtNum8.Text
        '(介護予防)短期入所療養介護(老健・診療所)
        Dim gen9 As String = If(checkGen9.Checked, "1", "0")
        Dim num9 As String = txtNum9.Text
        '(介護予防)特定施設入居者生活介護
        Dim gen10 As String = If(checkGen10.Checked, "1", "0")
        Dim num10 As String = txtNum10.Text
        '(介護予防)福祉用具貸与
        Dim gen11 As String = If(checkGen11.Checked, "1", "0")
        Dim num11 As String = txtNum11.Text
        '特定(介護予防)福祉用具販売
        Dim gen12 As String = If(checkGen12.Checked, "1", "0")
        Dim num12 As String = txtNum12.Text
        '住宅改修
        Dim gen13 As String = If(checkGen13.Checked, "1", "0")
        Dim num13 As String = If(CheckNum13Exists.Checked, "1", If(CheckNum13None.Checked, "2", ""))
        '夜間対応型訪問介護
        Dim gen14 As String = If(checkGen14.Checked, "1", "0")
        Dim num14 As String = txtNum14.Text
        '(介護予防)認知症対応型通所介護
        Dim gen15 As String = If(checkGen15.Checked, "1", "0")
        Dim num15 As String = txtNum15.Text
        '(介護予防)小規模多機能型居宅介護
        Dim gen16 As String = If(checkGen16.Checked, "1", "0")
        Dim num16 As String = txtNum16.Text
        '(介護予防)認知症対応型共同生活介護
        Dim gen17 As String = If(checkGen17.Checked, "1", "0")
        Dim num17 As String = txtNum17.Text
        '地域密着型特定施設入居者生活介護
        Dim gen18 As String = If(checkGen18.Checked, "1", "0")
        Dim num18 As String = txtNum18.Text
        '地域密着型介護老人福祉施設入所者生活介護
        Dim gen19 As String = If(checkGen19.Checked, "1", "0")
        Dim num19 As String = txtNum19.Text
        '定期巡回・随時対応型訪問介護看護
        Dim gen20 As String = If(checkGen20.Checked, "1", "0")
        Dim num20 As String = txtNum20.Text
        '複合型サービス
        Dim gen23 As String = If(checkGen23.Checked, "1", "0")
        Dim num21 As String = txtNum21.Text
        '市町村特別給付
        Dim gen21 As String = If(checkGen21.Checked, "1", "0")
        Dim gentxt1 As String = txtGentxt1.Text
        '介護保険給付外の在宅サービス
        Dim gen22 As String = If(checkGen22.Checked, "1", "0")
        Dim gentxt2 As String = txtGentxt2.Text

        '利用施設
        '介護老人福祉施設
        Dim stay1 As String = If(checkStay1.Checked, "1", "0")
        '介護老人保健施設
        Dim stay2 As String = If(checkStay2.Checked, "1", "0")
        '介護療養型医療施設
        Dim stay3 As String = If(checkStay3.Checked, "1", "0")
        '認知症対応型共同生活介護適用施設(ｸﾞﾙｰﾌﾟﾎｰﾑ)
        Dim stay4 As String = If(checkStay4.Checked, "1", "0")
        '特定施設入所者生活介護適用施設(ｹｱﾊｳｽ等)
        Dim stay5 As String = If(checkStay5.Checked, "1", "0")
        '医療機関(医療保険適用療養病床)
        Dim stay6 As String = If(checkStay6.Checked, "1", "0")
        '医療機関(療養病床以外)
        Dim stay7 As String = If(checkStay7.Checked, "1", "0")
        'その他の施設
        Dim stay8 As String = If(checkStay8.Checked, "1", "0")

        '施設連絡先
        Dim name As String = facilityNameBox.Text '施設名
        Dim pn31 As String = facilityPostCode1.Text '〒
        Dim pn32 As String = facilityPostCode2.Text '〒
        Dim ad3 As String = facilityAddress.Text '住所
        Dim tel31 As String = facilityTel1.Text '電話
        Dim tel32 As String = facilityTel2.Text '電話
        Dim tel33 As String = facilityTel3.Text '電話

        'Ⅳ
        Dim gTokki1 As String = spText1.Text
        Dim gTokki2 As String = spText2.Text
        Dim gTokki3 As String = spText3.Text
        Dim gTokki4 As String = spText4.Text

        '特記事項タブ
        '各タブのタイトル、項目名をAuth1sbテーブルから取得
        Dim cnn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        cnn.Open(topForm.DB_KSave2)
        Dim sql As String = "select Sp, Gyo, Txt from Auth1sb order by Sp, Gyo"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        '1タブの項目
        Dim txt0_1 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)
        rs.MoveNext()
        Dim txt0_2 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)
        rs.MoveNext()
        Dim txt0_3 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)
        rs.MoveNext()
        '2タブの項目
        Dim txt1_1 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)
        rs.MoveNext()
        Dim txt1_2 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)
        rs.MoveNext()
        Dim txt1_3 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)
        rs.MoveNext()
        Dim txt1_4 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)
        rs.MoveNext()
        '3タブの項目
        Dim txt2_1 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)
        rs.MoveNext()
        Dim txt2_2 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)
        rs.MoveNext()
        Dim txt2_3 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)
        rs.MoveNext()
        Dim txt2_4 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)
        rs.MoveNext()
        '4タブの項目
        Dim txt3_1 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)
        rs.MoveNext()
        Dim txt3_2 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)
        rs.MoveNext()
        Dim txt3_3 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)
        rs.MoveNext()
        Dim txt3_4 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)
        rs.MoveNext()
        Dim txt3_5 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)
        rs.MoveNext()
        '5タブの項目
        Dim txt4_1 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)
        rs.MoveNext()
        Dim txt4_2 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)
        rs.MoveNext()
        Dim txt4_3 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)
        rs.MoveNext()
        Dim txt4_4 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)
        rs.MoveNext()
        '6タブの項目
        Dim txt5_1 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)
        rs.MoveNext()
        Dim txt5_2 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)
        rs.MoveNext()
        Dim txt5_3 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)
        rs.MoveNext()
        '7タブの項目
        Dim txt6_1 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)
        rs.MoveNext()
        Dim txt6_2 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)
        rs.MoveNext()
        Dim txt6_3 As String = Util.checkDBNullValue(rs.Fields("Txt").Value)

        rs.Close()

        '登録
        rs = New ADODB.Recordset
        sql = "select * from Auth1 where Nam='" & userName & "' and Ymd1='" & ymd1 & "'"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
        If rs.RecordCount <> 0 Then
            Dim result As DialogResult = MessageBox.Show("変更登録してよろしいですか？", "登録", MessageBoxButtons.YesNo)
            If result = Windows.Forms.DialogResult.Yes Then
                '既存データ削除
                Dim cmd As New ADODB.Command()
                cmd.ActiveConnection = cnn
                cmd.CommandText = "delete from Auth1 where Nam='" & userName & "' and Ymd1='" & ymd1 & "'"
                cmd.Execute()
            Else
                '終了
                rs.Close()
                cnn.Close()
                Return
            End If
        End If

        '追加処理
        rs.AddNew()
        '概況調査タブ部分のレコード作成
        rs.Fields("Nam").Value = userName
        rs.Fields("Kana").Value = userKana
        rs.Fields("Gyo").Value = 61
        rs.Fields("Ymd1").Value = ymd1
        rs.Fields("Sp").Value = 0
        rs.Fields("GDay1").Value = gDay1
        rs.Fields("GDay2").Value = gDay2
        rs.Fields("GDay3").Value = gDay3
        rs.Fields("GDay4").Value = gDay4
        rs.Fields("GDay5").Value = gDay5
        rs.Fields("GDay6").Value = gDay6
        rs.Fields("GNum1").Value = gNum1
        rs.Fields("GNum2").Value = gNum2
        rs.Fields("GNum3").Value = gNum3
        rs.Fields("GNum4").Value = gNum4
        rs.Fields("GNum5").Value = gNum5
        rs.Fields("GNum6").Value = gNum6
        rs.Fields("GNum7").Value = gNum7
        rs.Fields("GNum8").Value = gNum8
        rs.Fields("GNum9").Value = gNum9
        rs.Fields("GNum10").Value = gNum10
        rs.Fields("Tanto").Value = tanto
        rs.Fields("Kikan").Value = kikan
        rs.Fields("Home").Value = home
        rs.Fields("Nonhm").Value = nonHm
        rs.Fields("Kako").Value = kako
        rs.Fields("Ymd2").Value = ymd2
        rs.Fields("Kai").Value = kai
        rs.Fields("Sex").Value = sex
        rs.Fields("Ymd3").Value = ymd3
        rs.Fields("Age").Value = age
        rs.Fields("Pn11").Value = pn11
        rs.Fields("Pn12").Value = pn12
        rs.Fields("Ad1").Value = ad1
        rs.Fields("Tel11").Value = tel11
        rs.Fields("Tel12").Value = tel12
        rs.Fields("Tel13").Value = tel13
        rs.Fields("Pn21").Value = pn21
        rs.Fields("Pn22").Value = pn22
        rs.Fields("Ad2").Value = ad2
        rs.Fields("Tel21").Value = tel21
        rs.Fields("Tel22").Value = tel22
        rs.Fields("Tel23").Value = tel23
        rs.Fields("Fa").Value = fa
        rs.Fields("Far").Value = far
        rs.Fields("Gen1").Value = gen1
        rs.Fields("Gen2").Value = gen2
        rs.Fields("Gen3").Value = gen3
        rs.Fields("Gen4").Value = gen4
        rs.Fields("Gen5").Value = gen5
        rs.Fields("Gen6").Value = gen6
        rs.Fields("Gen7").Value = gen7
        rs.Fields("Gen8").Value = gen8
        rs.Fields("Gen9").Value = gen9
        rs.Fields("Gen10").Value = gen10
        rs.Fields("Gen11").Value = gen11
        rs.Fields("Gen12").Value = gen12
        rs.Fields("Gen13").Value = gen13
        rs.Fields("Gen14").Value = gen14
        rs.Fields("Gen15").Value = gen15
        rs.Fields("Gen16").Value = gen16
        rs.Fields("Gen17").Value = gen17
        rs.Fields("Gen18").Value = gen18
        rs.Fields("Gen19").Value = gen19
        rs.Fields("Gen20").Value = gen20
        rs.Fields("Gen21").Value = gen21
        rs.Fields("Gen22").Value = gen22
        rs.Fields("Gen23").Value = gen23
        rs.Fields("Num1").Value = num1
        rs.Fields("Num2").Value = num2
        rs.Fields("Num3").Value = num3
        rs.Fields("Num4").Value = num4
        rs.Fields("Num5").Value = num5
        rs.Fields("Num6").Value = num6
        rs.Fields("Num7").Value = num7
        rs.Fields("Num8").Value = num8
        rs.Fields("Num9").Value = num9
        rs.Fields("Num10").Value = num10
        rs.Fields("Num11").Value = num11
        rs.Fields("Num12").Value = num12
        rs.Fields("Num13").Value = num13
        rs.Fields("Num14").Value = num14
        rs.Fields("Num15").Value = num15
        rs.Fields("Num16").Value = num16
        rs.Fields("Num17").Value = num17
        rs.Fields("Num18").Value = num18
        rs.Fields("Num19").Value = num19
        rs.Fields("Num20").Value = num20
        rs.Fields("Num21").Value = num21
        rs.Fields("GenTxt1").Value = gentxt1
        rs.Fields("GenTxt2").Value = gentxt2
        rs.Fields("Stay1").Value = stay1
        rs.Fields("Stay2").Value = stay2
        rs.Fields("Stay3").Value = stay3
        rs.Fields("Stay4").Value = stay4
        rs.Fields("Stay5").Value = stay5
        rs.Fields("Stay6").Value = stay6
        rs.Fields("Stay7").Value = stay7
        rs.Fields("Stay8").Value = stay8
        rs.Fields("Name").Value = name
        rs.Fields("Pn31").Value = pn31
        rs.Fields("Pn32").Value = pn32
        rs.Fields("Ad3").Value = ad3
        rs.Fields("Tel31").Value = tel31
        rs.Fields("Tel32").Value = tel32
        rs.Fields("Tel33").Value = tel33
        rs.Fields("GTokki1").Value = gTokki1
        rs.Fields("GTokki2").Value = gTokki2
        rs.Fields("GTokki3").Value = gTokki3
        rs.Fields("GTokki4").Value = gTokki4

        '特記事項タブのレコード作成
        '1.身体機能・起居動作タブ
        createSpDgvTabRecord(rs, 0, SpDgv1, {txt0_1, txt0_2, txt0_3}, userName, userKana, ymd1)
        '2.生活機能タブ
        createSpDgvTabRecord(rs, 1, SpDgv2, {txt1_1, txt1_2, txt1_3, txt1_4}, userName, userKana, ymd1)
        '3.認知機能タブ
        createSpDgvTabRecord(rs, 2, SpDgv3, {txt2_1, txt2_2, txt2_3, txt2_4}, userName, userKana, ymd1)
        '4.精神・行動障害タブ
        createSpDgvTabRecord(rs, 3, SpDgv4, {txt3_1, txt3_2, txt3_3, txt3_4, txt3_5}, userName, userKana, ymd1)
        '5.社会生活への適応
        createSpDgvTabRecord(rs, 4, SpDgv5, {txt4_1, txt4_2, txt4_3, txt4_4}, userName, userKana, ymd1)
        '6.特別な医療
        createSpDgvTabRecord(rs, 5, SpDgv6, {txt5_1, txt5_2, txt5_3}, userName, userKana, ymd1)
        '7.日常生活自立度
        createSpDgvTabRecord(rs, 6, SpDgv7, {txt6_1, txt6_2, txt6_3}, userName, userKana, ymd1)

        rs.Update()
        rs.Close()

        '基本調査タブのレコード作成
        rs = New ADODB.Recordset
        sql = "select Nam, Ymd, Gyo, Opt4, Ch2, Ch3, Ch4 from Auth2 where Nam='" & userName & "' and Ymd='" & ymd1 & "'"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
        If rs.RecordCount <> 0 Then
            '既存データ削除
            Dim cmd As New ADODB.Command()
            cmd.ActiveConnection = cnn
            cmd.CommandText = "delete from Auth2 where Nam='" & userName & "' and Ymd='" & ymd1 & "'"
            cmd.Execute()
        End If

        '1.身体機能・起居動作
        createBSTabRecord(rs, userName, ymd1, 0, rb1_3_1, Ch2_1, Ch3_1, Ch4_1) '1-3-1
        createBSTabRecord(rs, userName, ymd1, 1, rb1_3_2, Ch2_2, Ch3_2, Ch4_2) '1-3-2
        createBSTabRecord(rs, userName, ymd1, 2, rb1_3_3, Ch2_3, Ch3_3, Ch4_3) '1-3-3
        createBSTabRecord(rs, userName, ymd1, 3, rb1_4_1, Ch2_4, Ch3_4, Ch4_4) '1-4-1
        createBSTabRecord(rs, userName, ymd1, 4, rb1_4_2, Ch2_5, Ch3_5, Ch4_5) '1-4-2
        createBSTabRecord(rs, userName, ymd1, 5, rb1_4_3, Ch2_6, , Ch4_6) '1-4-3
        createBSTabRecord(rs, userName, ymd1, 6, rb1_5_1, , , Ch4_7) '1-5-1
        createBSTabRecord(rs, userName, ymd1, 7, rb1_5_2, , , Ch4_8) '1-5-2
        createBSTabRecord(rs, userName, ymd1, 8, rb1_5_3, , , Ch4_9) '1-5-3
        createBSTabRecord(rs, userName, ymd1, 9, rb1_5_4, , , Ch4_10) '1-5-4
        createBSTabRecord(rs, userName, ymd1, 10, rb1_6_1, , , Ch4_11) '1-6-1
        createBSTabRecord(rs, userName, ymd1, 11, rb1_6_2, , , Ch4_12) '1-6-2
        createBSTabRecord(rs, userName, ymd1, 12, rb1_6_3) '1-6-3
        createBSTabRecord(rs, userName, ymd1, 13, rb1_7_1) '1-7-1
        createBSTabRecord(rs, userName, ymd1, 14, rb1_7_2) '1-7-2
        createBSTabRecord(rs, userName, ymd1, 15, rb1_7_3) '1-7-3
        createBSTabRecord(rs, userName, ymd1, 16, rb1_8_1) '1-8-1
        createBSTabRecord(rs, userName, ymd1, 17, rb1_8_2) '1-8-2
        createBSTabRecord(rs, userName, ymd1, 18, rb1_8_3) '1-8-3
        createBSTabRecord(rs, userName, ymd1, 19, rb1_9_1) '1-9-1
        createBSTabRecord(rs, userName, ymd1, 20, rb1_9_2) '1-9-2
        createBSTabRecord(rs, userName, ymd1, 21, rb1_9_3) '1-9-3
        createBSTabRecord(rs, userName, ymd1, 22, rb1_10_1) '1-10-1
        createBSTabRecord(rs, userName, ymd1, 23, rb1_10_2) '1-10-2
        createBSTabRecord(rs, userName, ymd1, 24, rb1_10_3) '1-10-3
        createBSTabRecord(rs, userName, ymd1, 25, rb1_10_4) '1-10-4
        createBSTabRecord(rs, userName, ymd1, 26, rb1_11_1) '1-11-1
        createBSTabRecord(rs, userName, ymd1, 27, rb1_11_2) '1-11-2
        createBSTabRecord(rs, userName, ymd1, 28, rb1_11_3) '1-11-3
        createBSTabRecord(rs, userName, ymd1, 29, rb1_12_1) '1-12-1
        createBSTabRecord(rs, userName, ymd1, 30, rb1_12_2) '1-12-2
        createBSTabRecord(rs, userName, ymd1, 31, rb1_12_3) '1-12-3
        createBSTabRecord(rs, userName, ymd1, 32, rb1_12_4) '1-12-4
        createBSTabRecord(rs, userName, ymd1, 33, rb1_12_5) '1-12-5
        createBSTabRecord(rs, userName, ymd1, 34, rb1_13_1) '1-13-1
        createBSTabRecord(rs, userName, ymd1, 35, rb1_13_2) '1-13-2
        createBSTabRecord(rs, userName, ymd1, 36, rb1_13_3) '1-13-3
        createBSTabRecord(rs, userName, ymd1, 37, rb1_13_4) '1-13-4
        createBSTabRecord(rs, userName, ymd1, 38, rb1_13_5) '1-13-5
        '2.生活機能
        createBSTabRecord(rs, userName, ymd1, 39, rb2_1_1) '2-1-1
        createBSTabRecord(rs, userName, ymd1, 40, rb2_1_2) '2-1-2
        createBSTabRecord(rs, userName, ymd1, 41, rb2_1_3) '2-1-3
        createBSTabRecord(rs, userName, ymd1, 42, rb2_1_4) '2-1-4
        createBSTabRecord(rs, userName, ymd1, 43, rb2_2_1) '2-2-1
        createBSTabRecord(rs, userName, ymd1, 44, rb2_2_2) '2-2-2
        createBSTabRecord(rs, userName, ymd1, 45, rb2_2_3) '2-2-3
        createBSTabRecord(rs, userName, ymd1, 46, rb2_2_4) '2-2-4
        createBSTabRecord(rs, userName, ymd1, 47, rb2_3_1) '2-3-1
        createBSTabRecord(rs, userName, ymd1, 48, rb2_3_2) '2-3-2
        createBSTabRecord(rs, userName, ymd1, 49, rb2_3_3) '2-3-3
        createBSTabRecord(rs, userName, ymd1, 50, rb2_4_1) '2-4-1
        createBSTabRecord(rs, userName, ymd1, 51, rb2_4_2) '2-4-2
        createBSTabRecord(rs, userName, ymd1, 52, rb2_4_3) '2-4-3
        createBSTabRecord(rs, userName, ymd1, 53, rb2_4_4) '2-4-4
        createBSTabRecord(rs, userName, ymd1, 54, rb2_5_1) '2-5-1
        createBSTabRecord(rs, userName, ymd1, 55, rb2_5_2) '2-5-2
        createBSTabRecord(rs, userName, ymd1, 56, rb2_5_3) '2-5-3
        createBSTabRecord(rs, userName, ymd1, 57, rb2_5_4) '2-5-4
        createBSTabRecord(rs, userName, ymd1, 58, rb2_6_1) '2-6-1
        createBSTabRecord(rs, userName, ymd1, 59, rb2_6_2) '2-6-2
        createBSTabRecord(rs, userName, ymd1, 60, rb2_6_3) '2-6-3
        createBSTabRecord(rs, userName, ymd1, 61, rb2_6_4) '2-6-4
        createBSTabRecord(rs, userName, ymd1, 62, rb2_7_1) '2-7-1
        createBSTabRecord(rs, userName, ymd1, 63, rb2_7_2) '2-7-2
        createBSTabRecord(rs, userName, ymd1, 64, rb2_7_3) '2-7-3
        createBSTabRecord(rs, userName, ymd1, 65, rb2_8_1) '2-8-1
        createBSTabRecord(rs, userName, ymd1, 66, rb2_8_2) '2-8-2
        createBSTabRecord(rs, userName, ymd1, 67, rb2_8_3) '2-8-3
        createBSTabRecord(rs, userName, ymd1, 68, rb2_9_1) '2-9-1
        createBSTabRecord(rs, userName, ymd1, 69, rb2_9_2) '2-9-2
        createBSTabRecord(rs, userName, ymd1, 70, rb2_9_3) '2-9-3
        createBSTabRecord(rs, userName, ymd1, 71, rb2_10_1) '2-10-1
        createBSTabRecord(rs, userName, ymd1, 72, rb2_10_2) '2-10-2
        createBSTabRecord(rs, userName, ymd1, 73, rb2_10_3) '2-10-3
        createBSTabRecord(rs, userName, ymd1, 74, rb2_10_4) '2-10-4
        createBSTabRecord(rs, userName, ymd1, 75, rb2_11_1) '2-11-1
        createBSTabRecord(rs, userName, ymd1, 76, rb2_11_2) '2-11-2
        createBSTabRecord(rs, userName, ymd1, 77, rb2_11_3) '2-11-3
        createBSTabRecord(rs, userName, ymd1, 78, rb2_11_4) '2-11-4
        createBSTabRecord(rs, userName, ymd1, 79, rb2_12_1) '2-12-1
        createBSTabRecord(rs, userName, ymd1, 80, rb2_12_2) '2-12-2
        createBSTabRecord(rs, userName, ymd1, 81, rb2_12_3) '2-12-3
        '3.認知機能
        createBSTabRecord(rs, userName, ymd1, 82, rb3_1_1) '3-1-1
        createBSTabRecord(rs, userName, ymd1, 83, rb3_1_2) '3-1-2
        createBSTabRecord(rs, userName, ymd1, 84, rb3_1_3) '3-1-3
        createBSTabRecord(rs, userName, ymd1, 85, rb3_1_4) '3-1-4
        createBSTabRecord(rs, userName, ymd1, 86, rb3_2_1) '3-2-1
        createBSTabRecord(rs, userName, ymd1, 87, rb3_2_2) '3-2-2
        createBSTabRecord(rs, userName, ymd1, 88, rb3_3_1) '3-3-1
        createBSTabRecord(rs, userName, ymd1, 89, rb3_3_2) '3-3-2
        createBSTabRecord(rs, userName, ymd1, 90, rb3_4_1) '3-4-1
        createBSTabRecord(rs, userName, ymd1, 91, rb3_4_2) '3-4-2
        createBSTabRecord(rs, userName, ymd1, 92, rb3_5_1) '3-5-1
        createBSTabRecord(rs, userName, ymd1, 93, rb3_5_2) '3-5-2
        createBSTabRecord(rs, userName, ymd1, 94, rb3_6_1) '3-6-1
        createBSTabRecord(rs, userName, ymd1, 95, rb3_6_2) '3-6-2
        createBSTabRecord(rs, userName, ymd1, 96, rb3_7_1) '3-7-1
        createBSTabRecord(rs, userName, ymd1, 97, rb3_7_2) '3-7-2
        createBSTabRecord(rs, userName, ymd1, 98, rb3_8_1) '3-8-1
        createBSTabRecord(rs, userName, ymd1, 99, rb3_8_2) '3-8-2
        createBSTabRecord(rs, userName, ymd1, 100, rb3_8_3) '3-8-3
        createBSTabRecord(rs, userName, ymd1, 101, rb3_9_1) '3-9-1
        createBSTabRecord(rs, userName, ymd1, 102, rb3_9_2) '3-9-2
        createBSTabRecord(rs, userName, ymd1, 103, rb3_9_3) '3-9-3
        '4.精神・行動障害
        createBSTabRecord(rs, userName, ymd1, 104, rb4_1_1) '4-1-1
        createBSTabRecord(rs, userName, ymd1, 105, rb4_1_2) '4-1-2
        createBSTabRecord(rs, userName, ymd1, 106, rb4_1_3) '4-1-3
        createBSTabRecord(rs, userName, ymd1, 107, rb4_2_1) '4-2-1
        createBSTabRecord(rs, userName, ymd1, 108, rb4_2_2) '4-2-2
        createBSTabRecord(rs, userName, ymd1, 109, rb4_2_3) '4-2-3
        createBSTabRecord(rs, userName, ymd1, 110, rb4_3_1) '4-3-1
        createBSTabRecord(rs, userName, ymd1, 111, rb4_3_2) '4-3-2
        createBSTabRecord(rs, userName, ymd1, 112, rb4_3_3) '4-3-3
        createBSTabRecord(rs, userName, ymd1, 113, rb4_4_1) '4-4-1
        createBSTabRecord(rs, userName, ymd1, 114, rb4_4_2) '4-4-2
        createBSTabRecord(rs, userName, ymd1, 115, rb4_4_3) '4-4-3
        createBSTabRecord(rs, userName, ymd1, 116, rb4_5_1) '4-5-1
        createBSTabRecord(rs, userName, ymd1, 117, rb4_5_2) '4-5-2
        createBSTabRecord(rs, userName, ymd1, 118, rb4_5_3) '4-5-3
        createBSTabRecord(rs, userName, ymd1, 119, rb4_6_1) '4-6-1
        createBSTabRecord(rs, userName, ymd1, 120, rb4_6_2) '4-6-2
        createBSTabRecord(rs, userName, ymd1, 121, rb4_6_3) '4-6-3
        createBSTabRecord(rs, userName, ymd1, 122, rb4_7_1) '4-7-1
        createBSTabRecord(rs, userName, ymd1, 123, rb4_7_2) '4-7-2
        createBSTabRecord(rs, userName, ymd1, 124, rb4_7_3) '4-7-3
        createBSTabRecord(rs, userName, ymd1, 125, rb4_8_1) '4-8-1
        createBSTabRecord(rs, userName, ymd1, 126, rb4_8_2) '4-8-2
        createBSTabRecord(rs, userName, ymd1, 127, rb4_8_3) '4-8-3
        createBSTabRecord(rs, userName, ymd1, 128, rb4_9_1) '4-9-1
        createBSTabRecord(rs, userName, ymd1, 129, rb4_9_2) '4-9-2
        createBSTabRecord(rs, userName, ymd1, 130, rb4_9_3) '4-9-3
        createBSTabRecord(rs, userName, ymd1, 131, rb4_10_1) '4-10-1
        createBSTabRecord(rs, userName, ymd1, 132, rb4_10_2) '4-10-2
        createBSTabRecord(rs, userName, ymd1, 133, rb4_10_3) '4-10-3
        createBSTabRecord(rs, userName, ymd1, 134, rb4_11_1) '4-11-1
        createBSTabRecord(rs, userName, ymd1, 135, rb4_11_2) '4-11-2
        createBSTabRecord(rs, userName, ymd1, 136, rb4_11_3) '4-11-3
        createBSTabRecord(rs, userName, ymd1, 137, rb4_12_1) '4-12-1
        createBSTabRecord(rs, userName, ymd1, 138, rb4_12_2) '4-12-2
        createBSTabRecord(rs, userName, ymd1, 139, rb4_12_3) '4-12-3
        createBSTabRecord(rs, userName, ymd1, 140, rb4_13_1) '4-13-1
        createBSTabRecord(rs, userName, ymd1, 141, rb4_13_2) '4-13-2
        createBSTabRecord(rs, userName, ymd1, 142, rb4_13_3) '4-13-3
        createBSTabRecord(rs, userName, ymd1, 143, rb4_14_1) '4-14-1
        createBSTabRecord(rs, userName, ymd1, 144, rb4_14_2) '4-14-2
        createBSTabRecord(rs, userName, ymd1, 145, rb4_14_3) '4-14-3
        createBSTabRecord(rs, userName, ymd1, 146, rb4_15_1) '4-15-1
        createBSTabRecord(rs, userName, ymd1, 147, rb4_15_2) '4-15-2
        createBSTabRecord(rs, userName, ymd1, 148, rb4_15_3) '4-15-3
        '5.社会生活への適応
        createBSTabRecord(rs, userName, ymd1, 149, rb5_1_1) '5-1-1
        createBSTabRecord(rs, userName, ymd1, 150, rb5_1_2) '5-1-2
        createBSTabRecord(rs, userName, ymd1, 151, rb5_1_3) '5-1-3
        createBSTabRecord(rs, userName, ymd1, 152, rb5_2_1) '5-2-1
        createBSTabRecord(rs, userName, ymd1, 153, rb5_2_2) '5-2-2
        createBSTabRecord(rs, userName, ymd1, 154, rb5_2_3) '5-2-3
        createBSTabRecord(rs, userName, ymd1, 155, rb5_3_1) '5-3-1
        createBSTabRecord(rs, userName, ymd1, 156, rb5_3_2) '5-3-2
        createBSTabRecord(rs, userName, ymd1, 157, rb5_3_3) '5-3-3
        createBSTabRecord(rs, userName, ymd1, 158, rb5_3_4) '5-3-4
        createBSTabRecord(rs, userName, ymd1, 159, rb5_4_1) '5-4-1
        createBSTabRecord(rs, userName, ymd1, 160, rb5_4_2) '5-4-2
        createBSTabRecord(rs, userName, ymd1, 161, rb5_4_3) '5-4-3
        createBSTabRecord(rs, userName, ymd1, 162, rb5_5_1) '5-5-1
        createBSTabRecord(rs, userName, ymd1, 163, rb5_5_2) '5-5-2
        createBSTabRecord(rs, userName, ymd1, 164, rb5_5_3) '5-5-3
        createBSTabRecord(rs, userName, ymd1, 165, rb5_5_4) '5-5-4
        createBSTabRecord(rs, userName, ymd1, 166, rb5_6_1) '5-6-1
        createBSTabRecord(rs, userName, ymd1, 167, rb5_6_2) '5-6-2
        createBSTabRecord(rs, userName, ymd1, 168, rb5_6_3) '5-6-3
        createBSTabRecord(rs, userName, ymd1, 169, rb5_6_4) '5-6-4
        '7.日常生活自立度
        createBSTabRecord(rs, userName, ymd1, 170, rb7_1_1) '7-1-1
        createBSTabRecord(rs, userName, ymd1, 171, rb7_1_2) '7-1-2
        createBSTabRecord(rs, userName, ymd1, 172, rb7_1_3) '7-1-3
        createBSTabRecord(rs, userName, ymd1, 173, rb7_1_4) '7-1-4
        createBSTabRecord(rs, userName, ymd1, 174, rb7_1_5) '7-1-5
        createBSTabRecord(rs, userName, ymd1, 175, rb7_1_6) '7-1-6
        createBSTabRecord(rs, userName, ymd1, 176, rb7_1_7) '7-1-7
        createBSTabRecord(rs, userName, ymd1, 177, rb7_1_8) '7-1-8
        createBSTabRecord(rs, userName, ymd1, 178, rb7_1_9) '7-1-9
        createBSTabRecord(rs, userName, ymd1, 179, rb7_2_1) '7-2-1
        createBSTabRecord(rs, userName, ymd1, 180, rb7_2_2) '7-2-2
        createBSTabRecord(rs, userName, ymd1, 181, rb7_2_3) '7-2-3
        createBSTabRecord(rs, userName, ymd1, 182, rb7_2_4) '7-2-4
        createBSTabRecord(rs, userName, ymd1, 183, rb7_2_5) '7-2-5
        createBSTabRecord(rs, userName, ymd1, 184, rb7_2_6) '7-2-6
        createBSTabRecord(rs, userName, ymd1, 185, rb7_2_7) '7-2-7
        createBSTabRecord(rs, userName, ymd1, 186, rb7_2_8) '7-2-8

        rs.Update()
        rs.Close()
        cnn.Close()

        '再表示
        displayRecordList(userName)
        displayUserData(userName, userKana, ymd1)
    End Sub

    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click
        Dim userName As String = userLabel.Text
        Dim registYmd As String = dateYmdBox.getADStr()

        Dim rs As New ADODB.Recordset
        Dim cnn As New ADODB.Connection
        cnn.Open(topForm.DB_KSave2)
        Dim sql As String = "select * from Auth1 where Nam='" & userName & "' and Ymd1='" & registYmd & "'"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
        If rs.RecordCount = 0 Then
            MsgBox("削除対象のデータが存在しません。")
            rs.Close()
            cnn.Close()
            Return
        Else
            Dim result As DialogResult = MessageBox.Show("削除してよろしいですか？", "削除", MessageBoxButtons.YesNo)
            If result = Windows.Forms.DialogResult.Yes Then
                Dim cmd As New ADODB.Command()
                cmd.ActiveConnection = cnn

                'Auth1テーブルの削除(概況調査、特記事項タブ情報)
                cmd.CommandText = "delete from Auth1 where Nam='" & userName & "' and Ymd1='" & registYmd & "'"
                cmd.Execute()

                'Auth2テーブルの削除(基本調査タブ情報)
                cmd.CommandText = "delete from Auth2 where Nam='" & userName & "' and Ymd='" & registYmd & "'"
                cmd.Execute()

                rs.Close()
                cnn.Close()

                '再表示
                displayRecordList(userName)
                clearAllInputData()

            Else
                rs.Close()
                cnn.Close()
            End If
        End If
    End Sub

    Private Sub btnPrint_Click(sender As System.Object, e As System.EventArgs) Handles btnPrint.Click
        Dim userName As String = userLabel.Text '利用者名漢字
        Dim userKana As String = kanaLabel.Text '利用者名カナ
        If userName = "" OrElse userKana = "" Then
            MsgBox("利用者を選択して下さい。")
            Return
        End If
        Dim ymd1 As String = dateYmdBox.getADStr() '実施日

        Dim cnn As New ADODB.Connection
        cnn.Open(topForm.DB_KSave2)
        Dim rs As New ADODB.Recordset
        Dim sql = "select * from Auth1 where Nam='" & userName & "' and Ymd1='" & ymd1 & "' and Gyo=61"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
        If rs.RecordCount <= 0 Then
            MsgBox("対象の日付のデータが存在しません。")
            Return
        End If

        '日付変換(西暦→和暦)
        Dim ymd1Wareki As String = convADStrToWarekiStr(Util.checkDBNullValue(rs.Fields("Ymd1").Value))
        Dim ymd1Kanji As String = getKanji(ymd1Wareki)
        Dim ymd1Era As String = CInt(ymd1Wareki.Substring(1, 2)).ToString
        Dim ymd1Month As String = CInt(ymd1Wareki.Substring(4, 2)).ToString
        Dim ymd1Day As String = CInt(ymd1Wareki.Substring(7, 2)).ToString
        Dim ymd2Wareki As String = convADStrToWarekiStr(Util.checkDBNullValue(rs.Fields("Ymd2").Value))
        Dim ymd2Kanji As String = getKanji(ymd2Wareki)
        Dim ymd2Char As String = If(ymd2Wareki <> "", ymd2Wareki.Substring(0, 1), "")
        Dim ymd2Era As String = If(ymd2Wareki <> "", CInt(ymd2Wareki.Substring(1, 2)).ToString, "")
        Dim ymd2Month As String = If(ymd2Wareki <> "", CInt(ymd2Wareki.Substring(4, 2)).ToString, "")
        Dim ymd2Day As String = If(ymd2Wareki <> "", CInt(ymd2Wareki.Substring(7, 2)).ToString, "")
        Dim ymd3Wareki As String = convADStrToWarekiStr(Util.checkDBNullValue(rs.Fields("Ymd3").Value))
        Dim ymd3Kanji As String = getKanji(ymd3Wareki)
        Dim ymd3Era As String = If(ymd3Wareki <> "", CInt(ymd3Wareki.Substring(1, 2)).ToString, "")
        Dim ymd3Month As String = If(ymd3Wareki <> "", CInt(ymd3Wareki.Substring(4, 2)).ToString, "")
        Dim ymd3Day As String = If(ymd3Wareki <> "", CInt(ymd3Wareki.Substring(7, 2)).ToString, "")

        Dim gDay(5) As String
        Dim gNum(9) As String

        Dim objExcel As Object
        Dim objWorkBooks As Object
        Dim objWorkBook As Object
        Dim oSheet As Object
        Dim border As Object

        objExcel = CreateObject("Excel.Application")
        objWorkBooks = objExcel.Workbooks
        objWorkBook = objWorkBooks.Open(topForm.excelFilePass)

        '概況調査シート
        oSheet = objWorkBook.Worksheets("概況調査改")
        '調査日番号
        For i = 0 To 5
            gDay(i) = Util.checkDBNullValue(rs.Fields("GDay" & (i + 1)).Value)
        Next
        oSheet.Range("B4").value = gDay(0)
        oSheet.Range("D4").value = gDay(1)
        oSheet.Range("F4").value = gDay(2)
        oSheet.Range("G4").value = gDay(3)
        oSheet.Range("I4").value = gDay(4)
        oSheet.Range("L4").value = gDay(5)
        '被保険者番号
        For i = 0 To 9
            gNum(i) = Util.checkDBNullValue(rs.Fields("GNum" & (i + 1)).Value)
        Next
        oSheet.Range("AJ4").value = gNum(0)
        oSheet.Range("AO4").value = gNum(1)
        oSheet.Range("AR4").value = gNum(2)
        oSheet.Range("AV4").value = gNum(3)
        oSheet.Range("AX4").value = gNum(4)
        oSheet.Range("BB4").value = gNum(5)
        oSheet.Range("BG4").value = gNum(6)
        oSheet.Range("BL4").value = gNum(7)
        oSheet.Range("BP4").value = gNum(8)
        oSheet.Range("BV4").value = gNum(9)
        '実施日時
        oSheet.Range("H10").value = ymd1Kanji
        oSheet.Range("J10").value = ymd1Era
        oSheet.Range("N10").value = ymd1Month
        oSheet.Range("T10").value = ymd1Day
        '実施場所
        oSheet.Range("AF11").value = "自宅内"
        oSheet.Range("AO11").value = "自宅外"
        If Util.checkDBNullValue(rs.Fields("Home").Value) = "0" Then
            border = oSheet.Range("AF11", "AK11").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AF11", "AK11").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AF11").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AL11").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        ElseIf Util.checkDBNullValue(rs.Fields("Home").Value) = "1" Then
            border = oSheet.Range("AO11", "AS11").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AO11", "AS11").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AO11").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AT11").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        End If
        oSheet.Range("AU11").value = Util.checkDBNullValue(rs.Fields("Nonhm").Value) '自宅外テキスト
        oSheet.Range("I13").value = If(Util.checkDBNullValue(rs.Fields("Tanto").Value) = "河合　哲也", "ｶﾜｲ ﾃﾂﾔ", "") '記入者フリガナ
        oSheet.Range("I14").value = Util.checkDBNullValue(rs.Fields("Tanto").Value) '記入者氏名
        oSheet.Range("AR13").value = Util.checkDBNullValue(rs.Fields("Kikan").Value) '所属機関
        '過去の認定
        oSheet.Range("K19").value = "初回"
        oSheet.Range("Q19").value = "2回め以降"
        If Util.checkDBNullValue(rs.Fields("Kako").Value) = "0" Then
            border = oSheet.Range("K19", "N19").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("K20", "N20").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("K19", "K20").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("O19", "O20").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        ElseIf Util.checkDBNullValue(rs.Fields("Kako").Value) = "1" Then
            border = oSheet.Range("Q19", "Y19").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("Q20", "Y20").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("Q19", "Q20").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("Z19", "Z20").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        End If
        '前回認定
        oSheet.Range("P21").value = ymd2Char & ymd2Era
        oSheet.Range("U21").value = ymd2Month
        oSheet.Range("Y21").value = ymd2Day
        '前回認定結果
        oSheet.Range("AT20").value = "非該当"
        oSheet.Range("AX20").value = "要支援"
        oSheet.Range("BK20").value = "要介護"
        If Util.checkDBNullValue(rs.Fields("Kai").Value) = "" Then
            oSheet.Range("BF20").value = ""
            oSheet.Range("BR20").value = ""
        ElseIf Util.checkDBNullValue(rs.Fields("Kai").Value) = "0" Then
            border = oSheet.Range("AT20", "AV20").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AT21", "AV21").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AT20", "AT21").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AW20", "AW21").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            oSheet.Range("BF20").value = ""
            oSheet.Range("BR20").value = ""
        ElseIf (Util.checkDBNullValue(rs.Fields("Kai").Value) = "1") OrElse (Util.checkDBNullValue(rs.Fields("Kai").Value) = "2") Then
            border = oSheet.Range("AX20", "BB20").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AX21", "BB21").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AX20", "AX21").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BC20", "BC21").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            oSheet.Range("BF20").value = rs.Fields("Kai").Value
            oSheet.Range("BR20").value = ""
        ElseIf Util.checkDBNullValue(rs.Fields("Kai").Value) <= "7" Then
            border = oSheet.Range("BK20", "BN20").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BK21", "BN21").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BK20", "BK21").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BO20", "BO21").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            oSheet.Range("BF20").value = ""
            oSheet.Range("BR20").value = (rs.Fields("Kai").Value - 2)
        End If
        oSheet.Range("I23").value = Util.checkDBNullValue(rs.Fields("Kana").Value) 'ふりがな
        oSheet.Range("I26").value = Util.checkDBNullValue(rs.Fields("Nam").Value) '対象者氏名
        '性別
        oSheet.Range("AK25").value = "男"
        oSheet.Range("AP25").value = "女"
        If Util.checkDBNullValue(rs.Fields("Sex").Value) = "0" Then
            border = oSheet.Range("AK25", "AM25").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AK27", "AM27").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AK25", "AK27").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AN25", "AN27").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        ElseIf Util.checkDBNullValue(rs.Fields("Sex").Value) = "1" Then
            border = oSheet.Range("AP25", "AP25").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AP27", "AP27").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AP25", "AP27").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AQ25", "AQ27").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        End If
        '生年月日
        oSheet.Range("AY24").value = "明治"
        oSheet.Range("BC24").value = "大正"
        oSheet.Range("BH24").value = "昭和"
        oSheet.Range("BL24").value = "平成"
        If ymd3Kanji = "明治" Then
            border = oSheet.Range("AY24", "AZ24").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AY26", "AZ26").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AY24", "AY26").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BA24", "BA26").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        ElseIf ymd3Kanji = "大正" Then
            border = oSheet.Range("BC24", "BE24").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BC26", "BE26").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BC24", "BC26").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BF24", "BF26").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        ElseIf ymd3Kanji = "昭和" Then
            border = oSheet.Range("BH24", "BJ24").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BH26", "BJ26").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BH24", "BH26").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BK24", "BK26").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        ElseIf ymd3Kanji = "平成" Then
            border = oSheet.Range("BL24", "BM24").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BL26", "BM26").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BL24", "BL26").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BN24", "BN26").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        End If
        oSheet.Range("AY27").value = ymd3Era
        oSheet.Range("BE27").value = ymd3Month
        oSheet.Range("BJ27").value = ymd3Day
        oSheet.Range("BP27").value = Util.checkDBNullValue(rs.Fields("Age").Value)
        '現在所
        oSheet.Range("J29").value = Util.checkDBNullValue(rs.Fields("Pn11").Value)
        oSheet.Range("P29").value = Util.checkDBNullValue(rs.Fields("Pn12").Value)
        oSheet.Range("I30").value = Util.checkDBNullValue(rs.Fields("Ad1").Value)
        oSheet.Range("AY29").value = Util.checkDBNullValue(rs.Fields("Tel11").Value)
        oSheet.Range("BE29").value = Util.checkDBNullValue(rs.Fields("Tel12").Value)
        oSheet.Range("BN29").value = Util.checkDBNullValue(rs.Fields("Tel13").Value)
        '家族等連絡先
        oSheet.Range("J31").value = Util.checkDBNullValue(rs.Fields("Pn21").Value)
        oSheet.Range("P31").value = Util.checkDBNullValue(rs.Fields("Pn22").Value)
        oSheet.Range("I32").value = Util.checkDBNullValue(rs.Fields("Ad2").Value)
        oSheet.Range("AY31").value = Util.checkDBNullValue(rs.Fields("Tel21").Value)
        oSheet.Range("BE31").value = Util.checkDBNullValue(rs.Fields("Tel22").Value)
        oSheet.Range("BN31").value = Util.checkDBNullValue(rs.Fields("Tel23").Value)
        oSheet.Range("K33").value = Util.checkDBNullValue(rs.Fields("Fa").Value)
        oSheet.Range("AM33").value = Util.checkDBNullValue(rs.Fields("Far").Value)
        '在宅利用
        'チェックボックス部分
        oSheet.Range("D41").value = If(Util.checkDBNullValue(rs.Fields("Gen1").Value) = "1", "レ", "")
        oSheet.Range("D44").value = If(Util.checkDBNullValue(rs.Fields("Gen2").Value) = "1", "レ", "")
        oSheet.Range("D47").value = If(Util.checkDBNullValue(rs.Fields("Gen3").Value) = "1", "レ", "")
        oSheet.Range("D50").value = If(Util.checkDBNullValue(rs.Fields("Gen4").Value) = "1", "レ", "")
        oSheet.Range("D53").value = If(Util.checkDBNullValue(rs.Fields("Gen5").Value) = "1", "レ", "")
        oSheet.Range("D56").value = If(Util.checkDBNullValue(rs.Fields("Gen6").Value) = "1", "レ", "")
        oSheet.Range("D59").value = If(Util.checkDBNullValue(rs.Fields("Gen7").Value) = "1", "レ", "")
        oSheet.Range("D62").value = If(Util.checkDBNullValue(rs.Fields("Gen8").Value) = "1", "レ", "")
        oSheet.Range("D65").value = If(Util.checkDBNullValue(rs.Fields("Gen9").Value) = "1", "レ", "")
        oSheet.Range("D68").value = If(Util.checkDBNullValue(rs.Fields("Gen10").Value) = "1", "レ", "")
        oSheet.Range("AI41").value = If(Util.checkDBNullValue(rs.Fields("Gen11").Value) = "1", "レ", "")
        oSheet.Range("AI44").value = If(Util.checkDBNullValue(rs.Fields("Gen12").Value) = "1", "レ", "")
        oSheet.Range("AI47").value = If(Util.checkDBNullValue(rs.Fields("Gen13").Value) = "1", "レ", "")
        oSheet.Range("AI50").value = If(Util.checkDBNullValue(rs.Fields("Gen14").Value) = "1", "レ", "")
        oSheet.Range("AI53").value = If(Util.checkDBNullValue(rs.Fields("Gen15").Value) = "1", "レ", "")
        oSheet.Range("AI56").value = If(Util.checkDBNullValue(rs.Fields("Gen16").Value) = "1", "レ", "")
        oSheet.Range("AI59").value = If(Util.checkDBNullValue(rs.Fields("Gen17").Value) = "1", "レ", "")
        oSheet.Range("AI62").value = If(Util.checkDBNullValue(rs.Fields("Gen18").Value) = "1", "レ", "")
        oSheet.Range("AI65").value = If(Util.checkDBNullValue(rs.Fields("Gen19").Value) = "1", "レ", "")
        oSheet.Range("AI68").value = If(Util.checkDBNullValue(rs.Fields("Gen20").Value) = "1", "レ", "")
        oSheet.Range("D74").value = If(Util.checkDBNullValue(rs.Fields("Gen21").Value) = "1", "レ", "")
        oSheet.Range("D77").value = If(Util.checkDBNullValue(rs.Fields("Gen22").Value) = "1", "レ", "")
        oSheet.Range("D71").value = If(Util.checkDBNullValue(rs.Fields("Gen23").Value) = "1", "レ", "")
        '回数部分
        oSheet.Range("AB41").value = Util.checkDBNullValue(rs.Fields("Num1").Value)
        oSheet.Range("AB44").value = Util.checkDBNullValue(rs.Fields("Num2").Value)
        oSheet.Range("AB47").value = Util.checkDBNullValue(rs.Fields("Num3").Value)
        oSheet.Range("AB50").value = Util.checkDBNullValue(rs.Fields("Num4").Value)
        oSheet.Range("AB53").value = Util.checkDBNullValue(rs.Fields("Num5").Value)
        oSheet.Range("AB56").value = Util.checkDBNullValue(rs.Fields("Num6").Value)
        oSheet.Range("AB59").value = Util.checkDBNullValue(rs.Fields("Num7").Value)
        oSheet.Range("AB62").value = Util.checkDBNullValue(rs.Fields("Num8").Value)
        oSheet.Range("AB65").value = Util.checkDBNullValue(rs.Fields("Num9").Value)
        oSheet.Range("AB68").value = Util.checkDBNullValue(rs.Fields("Num10").Value)
        oSheet.Range("BN41").value = Util.checkDBNullValue(rs.Fields("Num11").Value)
        oSheet.Range("BN44").value = Util.checkDBNullValue(rs.Fields("Num12").Value)
        oSheet.Range("BO47").value = "あり"
        oSheet.Range("BU47").value = "なし"
        If Util.checkDBNullValue(rs.Fields("Num13").Value) = "1" Then
            border = oSheet.Range("BO47", "BR47").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BO47", "BR47").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BO47").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BS47").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        ElseIf Util.checkDBNullValue(rs.Fields("Num13").Value) = "2" Then
            border = oSheet.Range("BU47", "BV47").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BU47", "BV47").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BU47").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BW47").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        End If
        oSheet.Range("BR50").value = Util.checkDBNullValue(rs.Fields("Num14").Value)
        oSheet.Range("BR53").value = Util.checkDBNullValue(rs.Fields("Num15").Value)
        oSheet.Range("BR56").value = Util.checkDBNullValue(rs.Fields("Num16").Value)
        oSheet.Range("BR59").value = Util.checkDBNullValue(rs.Fields("Num17").Value)
        oSheet.Range("BR62").value = Util.checkDBNullValue(rs.Fields("Num18").Value)
        oSheet.Range("BR65").value = Util.checkDBNullValue(rs.Fields("Num19").Value)
        oSheet.Range("BR68").value = Util.checkDBNullValue(rs.Fields("Num20").Value)
        oSheet.Range("AB71").value = Util.checkDBNullValue(rs.Fields("Num21").Value)
        oSheet.Range("J74").value = Util.checkDBNullValue(rs.Fields("Gentxt1").Value)
        oSheet.Range("T77").value = Util.checkDBNullValue(rs.Fields("Gentxt2").Value)
        '利用施設
        'チェックボックス部分
        oSheet.Range("D83").value = If(Util.checkDBNullValue(rs.Fields("Stay1").Value) = "1", "レ", "")
        oSheet.Range("D86").value = If(Util.checkDBNullValue(rs.Fields("Stay2").Value) = "1", "レ", "")
        oSheet.Range("D89").value = If(Util.checkDBNullValue(rs.Fields("Stay3").Value) = "1", "レ", "")
        oSheet.Range("D92").value = If(Util.checkDBNullValue(rs.Fields("Stay4").Value) = "1", "レ", "")
        oSheet.Range("D95").value = If(Util.checkDBNullValue(rs.Fields("Stay5").Value) = "1", "レ", "")
        oSheet.Range("D98").value = If(Util.checkDBNullValue(rs.Fields("Stay6").Value) = "1", "レ", "")
        oSheet.Range("D101").value = If(Util.checkDBNullValue(rs.Fields("Stay7").Value) = "1", "レ", "")
        oSheet.Range("D104").value = If(Util.checkDBNullValue(rs.Fields("Stay8").Value) = "1", "レ", "")
        '施設連絡先
        oSheet.Range("AQ84").value = Util.checkDBNullValue(rs.Fields("Name").Value) '施設名
        oSheet.Range("AQ92").value = Util.checkDBNullValue(rs.Fields("Pn31").Value) '〒
        oSheet.Range("AV92").value = Util.checkDBNullValue(rs.Fields("Pn32").Value) '〒
        oSheet.Range("AQ94").value = Util.checkDBNullValue(rs.Fields("Ad3").Value) '住所
        oSheet.Range("AW104").value = Util.checkDBNullValue(rs.Fields("Tel31").Value) '電話
        oSheet.Range("BC104").value = Util.checkDBNullValue(rs.Fields("Tel32").Value) '電話
        oSheet.Range("BL104").value = Util.checkDBNullValue(rs.Fields("Tel33").Value) '電話
        'Ⅳ
        oSheet.Range("D109").value = Util.checkDBNullValue(rs.Fields("GTokki1").Value)
        oSheet.Range("D110").value = Util.checkDBNullValue(rs.Fields("GTokki2").Value)
        oSheet.Range("D111").value = Util.checkDBNullValue(rs.Fields("GTokki3").Value)
        oSheet.Range("D112").value = Util.checkDBNullValue(rs.Fields("GTokki4").Value)

        '基本調査タブ用値取得
        rs.Close()
        sql = "select * from Auth2 where Nam='" & userName & "' and Ymd='" & ymd1 & "' order by Gyo"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
        Dim opt4(186) As String
        Dim ch2(5) As String
        Dim ch3(4) As String
        Dim ch4(11) As String
        For i = 0 To 186
            If i <= 4 Then
                opt4(i) = Util.checkDBNullValue(rs.Fields("Opt4").Value)
                ch2(i) = Util.checkDBNullValue(rs.Fields("Ch2").Value)
                ch3(i) = Util.checkDBNullValue(rs.Fields("Ch3").Value)
                ch4(i) = Util.checkDBNullValue(rs.Fields("Ch4").Value)
            ElseIf i <= 5 Then
                opt4(i) = Util.checkDBNullValue(rs.Fields("Opt4").Value)
                ch2(i) = Util.checkDBNullValue(rs.Fields("Ch2").Value)
                ch4(i) = Util.checkDBNullValue(rs.Fields("Ch4").Value)
            ElseIf i <= 11 Then
                opt4(i) = Util.checkDBNullValue(rs.Fields("Opt4").Value)
                ch4(i) = Util.checkDBNullValue(rs.Fields("Ch4").Value)
            Else
                opt4(i) = Util.checkDBNullValue(rs.Fields("Opt4").Value)
            End If
            rs.MoveNext()
        Next
        rs.Close()

        '基本調査1
        '調査日番号
        oSheet.Range("B118").value = gDay(0)
        oSheet.Range("D118").value = gDay(1)
        oSheet.Range("F118").value = gDay(2)
        oSheet.Range("G118").value = gDay(3)
        oSheet.Range("I118").value = gDay(4)
        oSheet.Range("L118").value = gDay(5)
        '被保険者番号
        oSheet.Range("AJ118").value = gNum(0)
        oSheet.Range("AO118").value = gNum(1)
        oSheet.Range("AR118").value = gNum(2)
        oSheet.Range("AV118").value = gNum(3)
        oSheet.Range("AX118").value = gNum(4)
        oSheet.Range("BB118").value = gNum(5)
        oSheet.Range("BG118").value = gNum(6)
        oSheet.Range("BL118").value = gNum(7)
        oSheet.Range("BP118").value = gNum(8)
        oSheet.Range("BV118").value = gNum(9)
        '1-1
        oSheet.Range("C124").value = If(ch2(0) = "1", "①", "1.")
        oSheet.Range("G124").value = If(ch2(1) = "1", "②", "2.")
        oSheet.Range("P124").value = If(ch2(2) = "1", "③", "3.")
        oSheet.Range("AB124").value = If(ch2(3) = "1", "④", "4.")
        oSheet.Range("AP124").value = If(ch2(4) = "1", "⑤", "5.")
        oSheet.Range("AZ124").value = If(ch2(5) = "1", "⑥", "6.")
        '1-2
        oSheet.Range("C128").value = If(ch3(0) = "1", "①", "1.")
        oSheet.Range("H128").value = If(ch3(1) = "1", "②", "2.")
        oSheet.Range("T128").value = If(ch3(2) = "1", "③", "3.")
        oSheet.Range("AG128").value = If(ch3(3) = "1", "④", "4.")
        oSheet.Range("AV128").value = If(ch3(4) = "1", "⑤", "5.")
        '1-3
        oSheet.Range("C132").value = If(opt4(0) = "1", "①", "1.")
        oSheet.Range("W132").value = If(opt4(1) = "1", "②", "2.")
        oSheet.Range("BD132").value = If(opt4(2) = "1", "③", "3.")
        '1-4
        oSheet.Range("C136").value = If(opt4(3) = "1", "①", "1.")
        oSheet.Range("W136").value = If(opt4(4) = "1", "②", "2.")
        oSheet.Range("BD136").value = If(opt4(5) = "1", "③", "3.")
        '1-5
        oSheet.Range("C140").value = If(opt4(6) = "1", "①", "1.")
        oSheet.Range("I140").value = If(opt4(7) = "1", "②", "2.")
        oSheet.Range("AD140").value = If(opt4(8) = "1", "③", "3.")
        oSheet.Range("BD140").value = If(opt4(9) = "1", "④", "4.")
        '1-6
        oSheet.Range("C144").value = If(opt4(10) = "1", "①", "1.")
        oSheet.Range("W144").value = If(opt4(11) = "1", "②", "2.")
        oSheet.Range("BD144").value = If(opt4(12) = "1", "③", "3.")
        '1-7
        oSheet.Range("C148").value = If(opt4(13) = "1", "①", "1.")
        oSheet.Range("W148").value = If(opt4(14) = "1", "②", "2.")
        oSheet.Range("BD148").value = If(opt4(15) = "1", "③", "3.")
        '1-8
        oSheet.Range("C152").value = If(opt4(16) = "1", "①", "1.")
        oSheet.Range("W152").value = If(opt4(17) = "1", "②", "2.")
        oSheet.Range("BD152").value = If(opt4(18) = "1", "③", "3.")
        '1-9
        oSheet.Range("C156").value = If(opt4(19) = "1", "①", "1.")
        oSheet.Range("W156").value = If(opt4(20) = "1", "②", "2.")
        oSheet.Range("BD156").value = If(opt4(21) = "1", "③", "3.")
        '1-10
        oSheet.Range("C160").value = If(opt4(22) = "1", "①", "1.")
        oSheet.Range("R160").value = If(opt4(23) = "1", "②", "2.")
        oSheet.Range("AK160").value = If(opt4(24) = "1", "③", "3.")
        oSheet.Range("BD160").value = If(opt4(25) = "1", "④", "4.")
        '1-11
        oSheet.Range("C164").value = If(opt4(26) = "1", "①", "1.")
        oSheet.Range("W164").value = If(opt4(27) = "1", "②", "2.")
        oSheet.Range("BD164").value = If(opt4(28) = "1", "③", "3.")

        '基本調査2
        '調査日番号
        oSheet.Range("B174").value = gDay(0)
        oSheet.Range("D174").value = gDay(1)
        oSheet.Range("F174").value = gDay(2)
        oSheet.Range("G174").value = gDay(3)
        oSheet.Range("I174").value = gDay(4)
        oSheet.Range("L174").value = gDay(5)
        '被保険者番号
        oSheet.Range("AJ174").value = gNum(0)
        oSheet.Range("AO174").value = gNum(1)
        oSheet.Range("AR174").value = gNum(2)
        oSheet.Range("AV174").value = gNum(3)
        oSheet.Range("AX174").value = gNum(4)
        oSheet.Range("BB174").value = gNum(5)
        oSheet.Range("BG174").value = gNum(6)
        oSheet.Range("BL174").value = gNum(7)
        oSheet.Range("BP174").value = gNum(8)
        oSheet.Range("BV174").value = gNum(9)
        '1-12
        oSheet.Range("C179").value = If(opt4(29) = "1", "①", "1.")
        oSheet.Range("C180").value = If(opt4(30) = "1", "②", "2.")
        oSheet.Range("C181").value = If(opt4(31) = "1", "③", "3.")
        oSheet.Range("C182").value = If(opt4(32) = "1", "④", "4.")
        oSheet.Range("C183").value = If(opt4(33) = "1", "⑤", "5.")
        '1-13
        oSheet.Range("C188").value = If(opt4(34) = "1", "①", "1.")
        oSheet.Range("C189").value = If(opt4(35) = "1", "②", "2.")
        oSheet.Range("C190").value = If(opt4(36) = "1", "③", "3.")
        oSheet.Range("C191").value = If(opt4(37) = "1", "④", "4.")
        oSheet.Range("C192").value = If(opt4(38) = "1", "⑤", "5.")
        '2-1
        oSheet.Range("C197").value = If(opt4(39) = "1", "①", "1.")
        oSheet.Range("T197").value = If(opt4(40) = "1", "②", "2.")
        oSheet.Range("AJ197").value = If(opt4(41) = "1", "③", "3.")
        oSheet.Range("BD197").value = If(opt4(42) = "1", "④", "4.")
        '2-2
        oSheet.Range("C201").value = If(opt4(43) = "1", "①", "1.")
        oSheet.Range("T201").value = If(opt4(44) = "1", "②", "2.")
        oSheet.Range("AJ201").value = If(opt4(45) = "1", "③", "3.")
        oSheet.Range("BD201").value = If(opt4(46) = "1", "④", "4.")
        '2-3
        oSheet.Range("C205").value = If(opt4(47) = "1", "①", "1.")
        oSheet.Range("V205").value = If(opt4(48) = "1", "②", "2.")
        oSheet.Range("BD205").value = If(opt4(49) = "1", "③", "3.")
        '2-4
        oSheet.Range("C209").value = If(opt4(50) = "1", "①", "1.")
        oSheet.Range("T209").value = If(opt4(51) = "1", "②", "2.")
        oSheet.Range("AJ209").value = If(opt4(52) = "1", "③", "3.")
        oSheet.Range("BD209").value = If(opt4(53) = "1", "④", "4.")
        '2-5
        oSheet.Range("C213").value = If(opt4(54) = "1", "①", "1.")
        oSheet.Range("T213").value = If(opt4(55) = "1", "②", "2.")
        oSheet.Range("AJ213").value = If(opt4(56) = "1", "③", "3.")
        oSheet.Range("BD213").value = If(opt4(57) = "1", "④", "4.")
        '2-6
        oSheet.Range("C217").value = If(opt4(58) = "1", "①", "1.")
        oSheet.Range("T217").value = If(opt4(59) = "1", "②", "2.")
        oSheet.Range("AJ217").value = If(opt4(60) = "1", "③", "3.")
        oSheet.Range("BD217").value = If(opt4(61) = "1", "④", "4.")
        '2-7
        oSheet.Range("C221").value = If(opt4(62) = "1", "①", "1.")
        oSheet.Range("Z221").value = If(opt4(63) = "1", "②", "2.")
        oSheet.Range("BD221").value = If(opt4(64) = "1", "③", "3.")
        '2-8
        oSheet.Range("C225").value = If(opt4(65) = "1", "①", "1.")
        oSheet.Range("Z225").value = If(opt4(66) = "1", "②", "2.")
        oSheet.Range("BD225").value = If(opt4(67) = "1", "③", "3.")
        '2-9
        oSheet.Range("C229").value = If(opt4(68) = "1", "①", "1.")
        oSheet.Range("Z229").value = If(opt4(69) = "1", "②", "2.")
        oSheet.Range("BD229").value = If(opt4(70) = "1", "③", "3.")

        '基本調査3
        '調査日番号
        oSheet.Range("B233").value = gDay(0)
        oSheet.Range("D233").value = gDay(1)
        oSheet.Range("F233").value = gDay(2)
        oSheet.Range("G233").value = gDay(3)
        oSheet.Range("I233").value = gDay(4)
        oSheet.Range("L233").value = gDay(5)
        '被保険者番号
        oSheet.Range("AJ233").value = gNum(0)
        oSheet.Range("AO233").value = gNum(1)
        oSheet.Range("AR233").value = gNum(2)
        oSheet.Range("AV233").value = gNum(3)
        oSheet.Range("AX233").value = gNum(4)
        oSheet.Range("BB233").value = gNum(5)
        oSheet.Range("BG233").value = gNum(6)
        oSheet.Range("BL233").value = gNum(7)
        oSheet.Range("BP233").value = gNum(8)
        oSheet.Range("BV233").value = gNum(9)
        '2-10
        oSheet.Range("C238").value = If(opt4(71) = "1", "①", "1.")
        oSheet.Range("S238").value = If(opt4(72) = "1", "②", "2.")
        oSheet.Range("AJ238").value = If(opt4(73) = "1", "③", "3.")
        oSheet.Range("BD238").value = If(opt4(74) = "1", "④", "4.")
        '2-11
        oSheet.Range("C242").value = If(opt4(75) = "1", "①", "1.")
        oSheet.Range("S242").value = If(opt4(76) = "1", "②", "2.")
        oSheet.Range("AJ242").value = If(opt4(77) = "1", "③", "3.")
        oSheet.Range("BD242").value = If(opt4(78) = "1", "④", "4.")
        '2-12
        oSheet.Range("C246").value = If(opt4(79) = "1", "①", "1.")
        oSheet.Range("W246").value = If(opt4(80) = "1", "②", "2.")
        oSheet.Range("BD246").value = If(opt4(81) = "1", "③", "3.")
        '3-1
        oSheet.Range("C250").value = If(opt4(82) = "1", "①", "1.")
        oSheet.Range("C251").value = If(opt4(83) = "1", "②", "2.")
        oSheet.Range("C252").value = If(opt4(84) = "1", "③", "3.")
        oSheet.Range("C253").value = If(opt4(85) = "1", "④", "4.")
        '3-2
        oSheet.Range("C258").value = If(opt4(86) = "1", "①", "1.")
        oSheet.Range("W258").value = If(opt4(87) = "1", "②", "2.")
        '3-3
        oSheet.Range("C262").value = If(opt4(88) = "1", "①", "1.")
        oSheet.Range("W262").value = If(opt4(89) = "1", "②", "2.")
        '3-4
        oSheet.Range("C266").value = If(opt4(90) = "1", "①", "1.")
        oSheet.Range("W266").value = If(opt4(91) = "1", "②", "2.")
        '3-5
        oSheet.Range("C270").value = If(opt4(92) = "1", "①", "1.")
        oSheet.Range("W270").value = If(opt4(93) = "1", "②", "2.")
        '3-6
        oSheet.Range("C274").value = If(opt4(94) = "1", "①", "1.")
        oSheet.Range("W274").value = If(opt4(95) = "1", "②", "2.")
        '3-7
        oSheet.Range("C278").value = If(opt4(96) = "1", "①", "1.")
        oSheet.Range("W278").value = If(opt4(97) = "1", "②", "2.")
        '3-8
        oSheet.Range("C282").value = If(opt4(98) = "1", "①", "1.")
        oSheet.Range("Z282").value = If(opt4(99) = "1", "②", "2.")
        oSheet.Range("BD282").value = If(opt4(100) = "1", "③", "3.")
        '3-9
        oSheet.Range("C286").value = If(opt4(101) = "1", "①", "1.")
        oSheet.Range("Z286").value = If(opt4(102) = "1", "②", "2.")
        oSheet.Range("BD286").value = If(opt4(103) = "1", "③", "3.")

        '基本調査4
        '調査日番号
        oSheet.Range("B292").value = gDay(0)
        oSheet.Range("D292").value = gDay(1)
        oSheet.Range("F292").value = gDay(2)
        oSheet.Range("G292").value = gDay(3)
        oSheet.Range("I292").value = gDay(4)
        oSheet.Range("L292").value = gDay(5)
        '被保険者番号
        oSheet.Range("AJ292").value = gNum(0)
        oSheet.Range("AO292").value = gNum(1)
        oSheet.Range("AR292").value = gNum(2)
        oSheet.Range("AV292").value = gNum(3)
        oSheet.Range("AX292").value = gNum(4)
        oSheet.Range("BB292").value = gNum(5)
        oSheet.Range("BG292").value = gNum(6)
        oSheet.Range("BL292").value = gNum(7)
        oSheet.Range("BP292").value = gNum(8)
        oSheet.Range("BV292").value = gNum(9)
        '4-1
        oSheet.Range("C297").value = If(opt4(104) = "1", "①", "1.")
        oSheet.Range("Z297").value = If(opt4(105) = "1", "②", "2.")
        oSheet.Range("BD297").value = If(opt4(106) = "1", "③", "3.")
        '4-2
        oSheet.Range("C301").value = If(opt4(107) = "1", "①", "1.")
        oSheet.Range("Z301").value = If(opt4(108) = "1", "②", "2.")
        oSheet.Range("BD301").value = If(opt4(109) = "1", "③", "3.")
        '4-3
        oSheet.Range("C305").value = If(opt4(110) = "1", "①", "1.")
        oSheet.Range("Z305").value = If(opt4(111) = "1", "②", "2.")
        oSheet.Range("BD305").value = If(opt4(112) = "1", "③", "3.")
        '4-4
        oSheet.Range("C309").value = If(opt4(113) = "1", "①", "1.")
        oSheet.Range("Z309").value = If(opt4(114) = "1", "②", "2.")
        oSheet.Range("BD309").value = If(opt4(115) = "1", "③", "3.")
        '4-5
        oSheet.Range("C313").value = If(opt4(116) = "1", "①", "1.")
        oSheet.Range("Z313").value = If(opt4(117) = "1", "②", "2.")
        oSheet.Range("BD313").value = If(opt4(118) = "1", "③", "3.")
        '4-6
        oSheet.Range("C317").value = If(opt4(119) = "1", "①", "1.")
        oSheet.Range("Z317").value = If(opt4(120) = "1", "②", "2.")
        oSheet.Range("BD317").value = If(opt4(121) = "1", "③", "3.")
        '4-7
        oSheet.Range("C321").value = If(opt4(122) = "1", "①", "1.")
        oSheet.Range("Z321").value = If(opt4(123) = "1", "②", "2.")
        oSheet.Range("BD321").value = If(opt4(124) = "1", "③", "3.")
        '4-8
        oSheet.Range("C325").value = If(opt4(125) = "1", "①", "1.")
        oSheet.Range("Z325").value = If(opt4(126) = "1", "②", "2.")
        oSheet.Range("BD325").value = If(opt4(127) = "1", "③", "3.")
        '4-9
        oSheet.Range("C329").value = If(opt4(128) = "1", "①", "1.")
        oSheet.Range("Z329").value = If(opt4(129) = "1", "②", "2.")
        oSheet.Range("BD329").value = If(opt4(130) = "1", "③", "3.")
        '4-10
        oSheet.Range("C333").value = If(opt4(131) = "1", "①", "1.")
        oSheet.Range("Z333").value = If(opt4(132) = "1", "②", "2.")
        oSheet.Range("BD333").value = If(opt4(133) = "1", "③", "3.")
        '4-11
        oSheet.Range("C337").value = If(opt4(134) = "1", "①", "1.")
        oSheet.Range("Z337").value = If(opt4(135) = "1", "②", "2.")
        oSheet.Range("BD337").value = If(opt4(136) = "1", "③", "3.")
        '4-12
        oSheet.Range("C341").value = If(opt4(137) = "1", "①", "1.")
        oSheet.Range("Z341").value = If(opt4(138) = "1", "②", "2.")
        oSheet.Range("BD341").value = If(opt4(139) = "1", "③", "3.")
        '4-13
        oSheet.Range("C345").value = If(opt4(140) = "1", "①", "1.")
        oSheet.Range("Z345").value = If(opt4(141) = "1", "②", "2.")
        oSheet.Range("BD345").value = If(opt4(142) = "1", "③", "3.")

        '基本調査5
        '調査日番号
        oSheet.Range("B351").value = gDay(0)
        oSheet.Range("D351").value = gDay(1)
        oSheet.Range("F351").value = gDay(2)
        oSheet.Range("G351").value = gDay(3)
        oSheet.Range("I351").value = gDay(4)
        oSheet.Range("L351").value = gDay(5)
        '被保険者番号
        oSheet.Range("AJ351").value = gNum(0)
        oSheet.Range("AO351").value = gNum(1)
        oSheet.Range("AR351").value = gNum(2)
        oSheet.Range("AV351").value = gNum(3)
        oSheet.Range("AX351").value = gNum(4)
        oSheet.Range("BB351").value = gNum(5)
        oSheet.Range("BG351").value = gNum(6)
        oSheet.Range("BL351").value = gNum(7)
        oSheet.Range("BP351").value = gNum(8)
        oSheet.Range("BV351").value = gNum(9)
        '4-14
        oSheet.Range("C356").value = If(opt4(143) = "1", "①", "1.")
        oSheet.Range("Z356").value = If(opt4(144) = "1", "②", "2.")
        oSheet.Range("BD356").value = If(opt4(145) = "1", "③", "3.")
        '4-15
        oSheet.Range("C360").value = If(opt4(146) = "1", "①", "1.")
        oSheet.Range("Z360").value = If(opt4(147) = "1", "②", "2.")
        oSheet.Range("BD360").value = If(opt4(148) = "1", "③", "3.")
        '5-1
        oSheet.Range("C364").value = If(opt4(149) = "1", "①", "1.")
        oSheet.Range("Z364").value = If(opt4(150) = "1", "②", "2.")
        oSheet.Range("BD364").value = If(opt4(151) = "1", "③", "3.")
        '5-2
        oSheet.Range("C368").value = If(opt4(152) = "1", "①", "1.")
        oSheet.Range("Z368").value = If(opt4(153) = "1", "②", "2.")
        oSheet.Range("BD368").value = If(opt4(154) = "1", "③", "3.")
        '5-3
        oSheet.Range("C372").value = If(opt4(155) = "1", "①", "1.")
        oSheet.Range("S372").value = If(opt4(156) = "1", "②", "2.")
        oSheet.Range("AP372").value = If(opt4(157) = "1", "③", "3.")
        oSheet.Range("BD372").value = If(opt4(158) = "1", "④", "4.")
        '5-4
        oSheet.Range("C376").value = If(opt4(159) = "1", "①", "1.")
        oSheet.Range("Z376").value = If(opt4(160) = "1", "②", "2.")
        oSheet.Range("BD376").value = If(opt4(161) = "1", "③", "3.")
        '5-5
        oSheet.Range("C380").value = If(opt4(162) = "1", "①", "1.")
        oSheet.Range("S380").value = If(opt4(163) = "1", "②", "2.")
        oSheet.Range("AJ380").value = If(opt4(164) = "1", "③", "3.")
        oSheet.Range("BD380").value = If(opt4(165) = "1", "④", "4.")
        '5-6
        oSheet.Range("C384").value = If(opt4(166) = "1", "①", "1.")
        oSheet.Range("S384").value = If(opt4(167) = "1", "②", "2.")
        oSheet.Range("AJ384").value = If(opt4(168) = "1", "③", "3.")
        oSheet.Range("BD384").value = If(opt4(169) = "1", "④", "4.")
        '6
        oSheet.Range("I390").value = If(ch4(0) = "1", "①", "1.")
        oSheet.Range("W390").value = If(ch4(1) = "1", "②", "2.")
        oSheet.Range("AO390").value = If(ch4(2) = "1", "③", "3.")
        oSheet.Range("AX390").value = If(ch4(3) = "1", "④", "4.")
        oSheet.Range("I391").value = If(ch4(4) = "1", "⑤", "5.")
        oSheet.Range("W391").value = If(ch4(5) = "1", "⑥", "6.")
        oSheet.Range("AX391").value = If(ch4(6) = "1", "⑦", "7.")
        oSheet.Range("I392").value = If(ch4(7) = "1", "⑧", "8.")
        oSheet.Range("W392").value = If(ch4(8) = "1", "⑨", "9.")
        oSheet.Range("I393").value = If(ch4(9) = "1", "⑩", "10.")
        oSheet.Range("AX393").value = If(ch4(10) = "1", "⑪", "11.")
        oSheet.Range("I394").value = If(ch4(11) = "1", "⑫", "12.")
        '7
        oSheet.Range("Y400").value = "自立"
        oSheet.Range("AC400").value = "J1"
        oSheet.Range("AI400").value = "J2"
        oSheet.Range("AP400").value = "A1"
        oSheet.Range("AU400").value = "A2"
        oSheet.Range("AY400").value = "B1"
        oSheet.Range("BE400").value = "B2"
        oSheet.Range("BL400").value = "C1"
        oSheet.Range("BR400").value = "C2"
        oSheet.Range("Y403").value = "自立"
        oSheet.Range("AC403").value = "Ⅰ"
        oSheet.Range("AI403").value = "Ⅱa"
        oSheet.Range("AP403").value = "Ⅱb"
        oSheet.Range("AU403").value = "Ⅲa"
        oSheet.Range("AY403").value = "Ⅲb"
        oSheet.Range("BE403").value = "Ⅳ"
        oSheet.Range("BL403").value = "M"
        If opt4(170) = "1" Then
            border = oSheet.Range("Y400", "AA400").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("Y400", "AA400").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("Y400").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AB400").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        ElseIf opt4(171) = "1" Then
            border = oSheet.Range("AC400", "AF400").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AC400", "AF400").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AC400").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AG400").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        ElseIf opt4(172) = "1" Then
            border = oSheet.Range("AI400", "AM400").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AI400", "AM400").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AI400").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AN400").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        ElseIf opt4(173) = "1" Then
            border = oSheet.Range("AP400", "AR400").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AP400", "AR400").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AP400").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AS400").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        ElseIf opt4(174) = "1" Then
            border = oSheet.Range("AU400", "AV400").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AU400", "AV400").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AU400").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AW400").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        ElseIf opt4(175) = "1" Then
            border = oSheet.Range("AY400", "BB400").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AY400", "BB400").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AY400").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BC400").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        ElseIf opt4(176) = "1" Then
            border = oSheet.Range("BE400", "BI400").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BE400", "BI400").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BE400").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BJ400").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        ElseIf opt4(177) = "1" Then
            border = oSheet.Range("BL400", "BO400").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BL400", "BO400").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BL400").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BP400").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        ElseIf opt4(178) = "1" Then
            border = oSheet.Range("BR400", "BV400").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BR400", "BV400").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BR400").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BW400").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        End If

        If opt4(179) = "1" Then
            border = oSheet.Range("Y403", "AA403").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("Y403", "AA403").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("Y403").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AB403").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        ElseIf opt4(180) = "1" Then
            border = oSheet.Range("AC403", "AF403").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AC403", "AF403").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AC403").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AG403").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        ElseIf opt4(181) = "1" Then
            border = oSheet.Range("AI403", "AM403").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AI403", "AM403").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AI403").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AN403").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        ElseIf opt4(182) = "1" Then
            border = oSheet.Range("AP403", "AR403").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AP403", "AR403").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AP403").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AS403").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        ElseIf opt4(183) = "1" Then
            border = oSheet.Range("AU403", "AV403").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AU403", "AV403").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AU403").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AW403").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        ElseIf opt4(184) = "1" Then
            border = oSheet.Range("AY403", "BB403").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AY403", "BB403").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("AY403").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BC403").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        ElseIf opt4(185) = "1" Then
            border = oSheet.Range("BE403", "BI403").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BE403", "BI403").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BE403").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BJ403").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        ElseIf opt4(186) = "1" Then
            border = oSheet.Range("BL403", "BO403").Borders(Excel.XlBordersIndex.xlEdgeTop)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BL403", "BO403").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BL403").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
            border = oSheet.Range("BP403").Borders(Excel.XlBordersIndex.xlEdgeLeft)
            border.LineStyle = Excel.XlLineStyle.xlDot
            border.Weight = Excel.XlBorderWeight.xlHairline
        End If

        '改ページ
        oSheet.HpageBreaks.add(oSheet.Range("A116"))
        oSheet.HpageBreaks.add(oSheet.Range("A172"))
        oSheet.HpageBreaks.add(oSheet.Range("A231"))
        oSheet.HpageBreaks.add(oSheet.Range("A290"))
        oSheet.HpageBreaks.add(oSheet.Range("A349"))

        '特記事項
        oSheet = objWorkBook.Worksheets("特記事項改")
        '調査日番号
        oSheet.Range("B5").value = gDay(0)
        oSheet.Range("E5").value = gDay(1)
        oSheet.Range("H5").value = gDay(2)
        oSheet.Range("I5").value = gDay(3)
        oSheet.Range("J5").value = gDay(4)
        oSheet.Range("K5").value = gDay(5)
        '真ん中の数字の
        oSheet.Range("P5").value = "1"
        '被保険者番号
        oSheet.Range("R5").value = gNum(0)
        oSheet.Range("S5").value = gNum(1)
        oSheet.Range("T5").value = gNum(2)
        oSheet.Range("U5").value = gNum(3)
        oSheet.Range("V5").value = gNum(4)
        oSheet.Range("W5").value = gNum(5)
        oSheet.Range("X5").value = gNum(6)
        oSheet.Range("Y5").value = gNum(7)
        oSheet.Range("Z5").value = gNum(8)
        oSheet.Range("AA5").value = gNum(9)

        Dim sp As Integer = 0
        Dim insertCount As Integer = 0
        Dim startIndexArray() As Integer = {13, 17, 21, 26, 30, 33, 36}
        Dim index As Integer = startIndexArray(0)
        oSheet.Range("7:7").rows.hidden = False
        sql = "select * from Auth1 where Nam='" & userName & "' and Ymd1='" & ymd1 & "' order by Sp, Gyo"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
        While Not rs.EOF
            If Util.checkDBNullValue(rs.Fields("Sp").Value) <> sp Then
                sp = Util.checkDBNullValue(rs.Fields("Sp").Value)
                index = startIndexArray(sp)
            End If
            If Util.checkDBNullValue(rs.Fields("Txt").Value) <> "" Then
                oSheet.Range("C7").value = Util.checkDBNullValue(rs.Fields("Crr").Value)
                oSheet.Range("H7").value = Util.checkDBNullValue(rs.Fields("Txt").Value)

                '行追加
                oSheet.Range((index + insertCount) & ":" & (index + insertCount)).insert()

                'コピペ
                Dim xlRange As Excel.Range = oSheet.Range("7:7")
                xlRange.Copy()
                Dim xlPasteRange As Excel.Range = oSheet.Range((index + insertCount) & ":" & (index + insertCount))
                oSheet.Paste(xlPasteRange)

                insertCount += 1
            End If
            rs.MoveNext()
        End While
        oSheet.Range("7:7").rows.hidden = True

        If 22 <= insertCount AndAlso insertCount <= 68 Then '枚数が2枚になるときの処理
            '行追加
            oSheet.Range("57:64").insert()

            '行のコピペ処理
            '上の番号部分
            Dim xlRange As Excel.Range = oSheet.Range("1:6")
            xlRange.Copy()
            Dim xlPasteRange As Excel.Range = oSheet.Range("57:63")
            oSheet.Paste(xlPasteRange)
            oSheet.Range("P61").value = "2"
            '頭の空白行分
            xlRange = oSheet.Range("8:9")
            xlRange.Copy()
            xlPasteRange = oSheet.Range("63:64")
            oSheet.Paste(xlPasteRange)
            oSheet.Range("B63").value = ""

            '改ページ
            oSheet.HpageBreaks.add(oSheet.Range("A57"))

        ElseIf 69 <= insertCount Then '枚数が3枚になるときの処理
            '2枚目部分
            oSheet.Range("57:64").insert() '行追加
            '行のコピペ処理
            '上の番号部分
            Dim xlRange As Excel.Range = oSheet.Range("1:6")
            xlRange.Copy()
            Dim xlPasteRange As Excel.Range = oSheet.Range("57:63")
            oSheet.Paste(xlPasteRange)
            oSheet.Range("P61").value = "2"
            '頭の空白行分
            xlRange = oSheet.Range("8:9")
            xlRange.Copy()
            xlPasteRange = oSheet.Range("63:64")
            oSheet.Paste(xlPasteRange)
            oSheet.Range("B63").value = ""
            '改ページ
            oSheet.HpageBreaks.add(oSheet.Range("A57"))

            '3枚目部分
            oSheet.Range("112:119").insert() '行追加
            '行のコピペ処理
            '上の番号部分
            xlRange = oSheet.Range("1:6")
            xlRange.Copy()
            xlPasteRange = oSheet.Range("112:118")
            oSheet.Paste(xlPasteRange)
            oSheet.Range("P116").value = "3"
            '頭の空白行分
            xlRange = oSheet.Range("8:9")
            xlRange.Copy()
            xlPasteRange = oSheet.Range("118:119")
            oSheet.Paste(xlPasteRange)
            oSheet.Range("B118").value = ""
            '改ページ
            oSheet.HpageBreaks.add(oSheet.Range("A112"))
        End If

        '変更保存確認ダイアログ非表示
        objExcel.DisplayAlerts = False

        '印刷1
        If topForm.rbtnPrint.Checked = True Then
            objWorkBook.Worksheets({"概況調査改"}).printOut()
        ElseIf topForm.rbtnPreview.Checked = True Then
            objExcel.Visible = True
            objWorkBook.Worksheets({"概況調査改"}).PrintPreview(1)
        End If

        '印刷2
        If topForm.rbtnPrint.Checked = True Then
            objWorkBook.Worksheets({"特記事項改"}).printOut()
        ElseIf topForm.rbtnPreview.Checked = True Then
            objExcel.Visible = True
            objWorkBook.Worksheets({"特記事項改"}).PrintPreview(1)
        End If

        ' EXCEL解放
        objExcel.Quit()
        Marshal.ReleaseComObject(objWorkBook)
        Marshal.ReleaseComObject(objExcel)
        oSheet = Nothing
        objWorkBook = Nothing
        objExcel = Nothing
    End Sub

    Private Sub txtNum_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtNum1.KeyDown, txtNum2.KeyDown, txtNum3.KeyDown, txtNum4.KeyDown, txtNum5.KeyDown, txtNum6.KeyDown, txtNum7.KeyDown, txtNum8.KeyDown, txtNum9.KeyDown, txtNum10.KeyDown, txtNum11.KeyDown, txtNum12.KeyDown, txtNum14.KeyDown, txtNum15.KeyDown, txtNum16.KeyDown, txtNum17.KeyDown, txtNum18.KeyDown, txtNum19.KeyDown, txtNum20.KeyDown, txtNum21.KeyDown
        Dim tb As TextBox = CType(sender, TextBox)
        Dim index As Integer = CInt(tb.Name.Substring(6, tb.Name.Length - 6))

        If e.KeyCode = Keys.Down Then '↓キー処理
            If index < 20 Then
                If index = 10 Then
                    txtNum21.Focus()
                ElseIf index = 12 Then
                    txtNum14.Focus()
                Else
                    overview3Panel.Controls("txtNum" & (index + 1)).Focus()
                End If
            End If
        ElseIf e.KeyCode = Keys.Up Then '↑キー処理
            If index <> 1 AndAlso index <> 11 Then
                If index = 21 Then
                    txtNum10.Focus()
                ElseIf index = 14 Then
                    txtNum12.Focus()
                Else
                    overview3Panel.Controls("txtNum" & (index - 1)).Focus()
                End If
            End If
        ElseIf e.KeyCode = Keys.Enter AndAlso index = 20 Then 'txtNum20のエンターキー処理
            txtGentxt1.Focus()
        End If
    End Sub
End Class