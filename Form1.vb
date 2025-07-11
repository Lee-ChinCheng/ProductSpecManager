Imports System.Drawing.Printing
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports ClosedXML.Excel
Imports DocumentFormat.OpenXml.Wordprocessing
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Collections.Generic
Imports System.Linq
Imports DocumentFormat.OpenXml.Office2016.Drawing.Charts
Imports System.Runtime.InteropServices.ComTypes
Imports System.Runtime.InteropServices
Imports System.Windows.Forms.VisualStyles


Public Class Form1

    'history filename
    Dim js_DB_Path As String = "C:\Users\user1\source\repos\WFA_TM\historyfile\AllDB.json"
    Dim js_DBdisc_Path As String = "C:\Users\user1\source\repos\WFA_TM\history_disc"

    Dim partSpec As New PartSpecification()

    '--- DataSet 來存儲所有零件資訊
    Dim partDataSet As New DataSet()

    '--- 規格儲存變數
    Dim Now_SPEC As Now_SPEC_C = New Now_SPEC_C()

    '--- for checking valid file name
    Dim HS_class As New HS_Clss()

    '--- Excel 相關路徑宣告
    Dim folderPath As String = AppDomain.CurrentDomain.BaseDirectory

    'name for output excel report
    Dim xlsxName As String = "Input_Model_name"

    'prepare template.xlsx first for defining output excel format 
    Dim template_excel As String = "template"
    Dim templatePath As String = IO.Path.Combine(folderPath, template_excel & ".xlsx")
    'C:\Users\\source\repos\WFAdemo1\bin\Debug



    '儲存目前使用的
    Dim NowTable As DataTable

    '數據規格表路徑
    Private PartDataFilePath As String = Path.Combine(Application.StartupPath, "PartData.json") '

    '---------

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        LoadData()

    End Sub

    Private Sub LoadData()

        ' 如果数据文件存在PartDataFilePath
        If File.Exists(PartDataFilePath) Then

            Dim json As String = File.ReadAllText(PartDataFilePath)
            NowTable = JsonConvert.DeserializeObject(Of DataTable)(json)
            NowTable.TableName = "PartTable"

            DataGridView1.DataSource = NowTable
            ' 凍結首列
            DataGridView1.Columns(0).Frozen = True
            ' 凍結首行
            DataGridView1.Rows(0).Frozen = True

            Debug.Print("Now Table Name = " & NowTable.TableName)

            ' 列出 DataTable 中所有欄位名稱
            For Each column As DataColumn In NowTable.Columns
                Debug.Print(column.ColumnName)
            Next

            ' 將 CoolingCapacity 加入 ComboBox
            Combobox_CoCap.DataSource = GetDataTables(NowTable, "Cooling Capacity")
            ComboBox_Pump.DataSource = GetDataTables(NowTable, "Pump Redundancy")
            ' in tabpage 3
            Cbx_h_CoCap.DataSource = GetDataTables(NowTable, "Cooling Capacity")
            Cbx_h_ApproachT.DataSource = GetDataTables(NowTable, "Approach Temp")

        End If

        '--- create HashSet for filename, when App active
        If File.Exists(js_DBdisc_Path) Then
            ' Get all JSON files in the folder
            Dim js_Files As String() = Directory.GetFiles(js_DBdisc_Path, "*.json")

            ' Add each filename (without extension) to the HashSet
            For Each js_element_Path In js_Files
                Dim fName_Only As String = Path.GetFileNameWithoutExtension(js_element_Path)
                HS_class.Name_HS.Add(fName_Only)
            Next
            'For Each set_element In HS_class.Name_HS
            'Console.WriteLine(set_element)
            'Next
        End If
        '-----------------------------------


        '---------------------

        'Textbox代入預設檔案儲存名稱
        TextBox_ModelName.Text = xlsxName
        Dim currentDate As String = DateTime.Now.ToString("yyyy-MM-dd")
        TextBox_DateName.Text = currentDate
        Tbx_DateEnd.Text = currentDate

    End Sub

    ''' <summary>
    ''' 取得指定DataTable中某一欄的所有數值，去掉重複值，輸出成DataTable給Combobox使用
    ''' </summary>
    ''' <param name="tableName">要找的table名稱</param>
    ''' <param name="ColumnName">要輸出的指定欄位</param>
    ''' <returns></returns>
    Private Function GetDataTables(table As DataTable, ColumnName As String) As List(Of String)

        ' 遍歷 DataTable 的所有行
        For Each row As DataRow In table.Rows
            ' 遍歷該行的所有欄位並列印欄位名稱和值
            Dim rowData As String = ""
            For Each column As DataColumn In table.Columns
                rowData &= column.ColumnName & ": " & row(column).ToString() & "   "
            Next
            ' 使用 Debug.Print 輸出行的資料
            'Debug.Print(rowData)
        Next

        Dim uniqueName As List(Of String) = New List(Of String)
        uniqueName = table _
            .AsEnumerable() _
            .Select(Function(row)
                        Dim fieldValue As Object = row(ColumnName)

                        If fieldValue IsNot DBNull.Value Then
                            ' 檢查是否為 List(Of String) 類型
                            If TypeOf fieldValue Is List(Of String) Then
                                ' 如果是 List(Of String)，則返回 List 中的值
                                Return String.Join(",", CType(fieldValue, List(Of String)))
                            Else
                                ' 如果是其他類型，直接轉換為字串
                                Return fieldValue.ToString()
                            End If
                        Else
                            ' 如果是 DBNull，則返回空字串
                            Return String.Empty
                        End If
                    End Function) _
            .Distinct() _
            .ToList()

        '加入空的一欄，讓Combobox一開始選擇空值
        uniqueName.Insert(0, "")

        Return uniqueName
    End Function

    ''' <summary>
    ''' Cooling變更事件
    ''' </summary>
    Private Sub CoolingSelectResult()

        '檢查是否有選取
        If Combobox_CoCap.SelectedIndex < 0 Or Combobox_CoCap.Text = "" Then
            ClearAll()
            Return
        End If

        Debug.Print("Combobox_CoCap.SelectedText = " & Combobox_CoCap.Text)

        Dim selectedClloingRow As DataRow
        selectedClloingRow = NowTable.AsEnumerable() _
            .FirstOrDefault(Function(row) Convert.ToDouble(row.Field(Of Object)("Cooling Capacity")) = Convert.ToDouble(Combobox_CoCap.Text))

        Dim inputValue As Integer = Combobox_CoCap.SelectedValue
        Dim parameter As Double

        Dim mocp As Double = selectedClloingRow.Field(Of Double)("MOCP")
        Dim fla As Double = selectedClloingRow.Field(Of Double)("FLA")

        '輸入錯誤
        If Not Double.TryParse(Textbox_Para.Text, parameter) AndAlso Textbox_Para.Text <> "" Then
            Textbox_CoFlow.Text = "Invalid Input"
            Return
        End If

        Dim result As Double = Math.Round(inputValue * parameter)   '四捨五入
        Now_SPEC.CoolingCapacity = Combobox_CoCap.SelectedValue
        Textbox_CoFlow.Text = result
        Try
            Now_SPEC.IndexFactor = Convert.ToDouble(Textbox_Para.Text)
            Now_SPEC.SecCoolantFlow = result
        Catch ex As Exception
            '忽略空值的計算錯誤
        End Try
        Now_SPEC.MOCP = mocp
        Now_SPEC.FLA = fla

        Dim filteredRows As DataRow()
        Dim uniqueNames As List(Of Double)
        ' 從 DataTable中的CoolingCapacity篩選出與所選相同的值，取出資料行
        filteredRows = NowTable.Select("[Cooling Capacity] = " & Now_SPEC.CoolingCapacity.ToString())

        '---------------------ApproachTemp處理-----------------------

        ' 將FrameDimension轉成Combobox可用的陣列
        uniqueNames = filteredRows _
        .Select(Function(row) Convert.ToDouble(row("Approach Temp"))) _
        .Distinct() _
        .ToList()

        ComboBox_ApproachT.DataSource = uniqueNames

        '如果ApproachTemp唯一則自動填入
        '如果ApproachTemp有多個則需要讓使用者選Pump
        If uniqueNames.Count() = 1 Then
            Now_SPEC.ApproachTemp = uniqueNames(0)

            Dim rows As Integer
            rows = filteredRows.Select(Function(row) row("Prim. Coolant Flow")).First()
            Now_SPEC.PrimCoolantFlow = rows
            TextBox_PrimCF.Text = Now_SPEC.PrimCoolantFlow & "LPM"
            rows = filteredRows.Select(Function(row) row("Operating DP")).First()
            TextBox_OperatingDP.Text = Now_SPEC.OperatingDP & "psi"
        Else
            TextBox_PrimCF.Text = ""
        End If
        '------------------------------------------------------------

        '---------------------ExternalDP處理-----------------------
        Dim sec = Now_SPEC.SecCoolantFlow
        Dim cc = Now_SPEC.CoolingCapacity
        'for future thermal calculation

        ' 將FrameDimension轉成Combobox可用的陣列
        uniqueNames = filteredRows _
        .Select(Function(row) Convert.ToDouble(row("External DP"))) _
        .Distinct() _
        .ToList()

        ComboBox_ExternalDP.DataSource = uniqueNames

        '如果ApproachTemp唯一則自動填入
        '如果ApproachTemp有多個則讓使用者選Pump
        ComboBox_ExternalDP.SelectedIndex = 0
        Now_SPEC.ExternalDP = uniqueNames(0)

        'Dim exdp As Integer
        'Now_SPEC.ExternalDP = exdp
        ComboBox_ExternalDP.DataSource = uniqueNames
        '------------------------------------------------------------

        '---------------------其他---------------------
        ' ComboBox_SecCoolantType.DataSource = GetDataTables(NowTable, "Sec. Coolant Type")
        'ComboBox_Display.DataSource = GetDataTables(NowTable, "Display")
        'ComboBox_Protocols.DataSource = GetDataTables(NowTable, "Protocols")
        'ComboBox_PrimConnection.DataSource = GetDataTables(NowTable, "Prim. Connection")
        'ComboBox_SecConnection.DataSource = GetDataTables(NowTable, "Sec. Connection")
        'ComboBox_ConnectionLoaction.Items.Clear() ' 清空現有的項目
        'ComboBox_ConnectionLoaction.Items.AddRange(New String() {"Top", "Down"})
    End Sub

    Private Sub Button_export_Click(sender As Object, e As EventArgs) Handles Button_export.Click

        Now_SPEC.PrintNowSpecValues(Now_SPEC)

        ' store user-type-in message
        Now_SPEC.ModelName = TextBox_ModelName.Text
        Now_SPEC.OwnerName = TextBox_OwnerName.Text
        Now_SPEC.DateName = TextBox_DateName.Text
        Label65.Text = "" 'clear this 

        Try
            'read the format of template.xlsx
            Dim workbook = New XLWorkbook(file:=templatePath)
            Dim worksheet As IXLWorksheet = workbook.Worksheet(1)

            Dim um_unit As String = ChrW(&HB5) & "m" 'alphabet µm
            'Console.WriteLine(um_unit)  ' Output: µm

            worksheet.Cell("B3").Value = Now_SPEC.ModelName
            worksheet.Cell("D3").Value = Now_SPEC.OwnerName
            worksheet.Cell("F3").Value = Now_SPEC.DateName

            worksheet.Cell("B4").Value = Now_SPEC.CoolingCapacity & "kW"
            worksheet.Cell("B6").Value = Convert.ToDouble(Now_SPEC.IndexFactor)
            worksheet.Cell("B7").Value = Now_SPEC.PrimaryWater
            worksheet.Cell("B8").Value = Now_SPEC.PrimCoolantType
            worksheet.Cell("B9").Value = Now_SPEC.PrimCoolantFlow & um_unit
            worksheet.Cell("B10").Value = Now_SPEC.OperatingDP & "psi"
            worksheet.Cell("B11").Value = Now_SPEC.PrimCoolantFilter & um_unit
            worksheet.Cell("B13").Value = Now_SPEC.SecCoolantType
            worksheet.Cell("B14").Value = Now_SPEC.SecCoolantFlow & "LPM"
            worksheet.Cell("B15").Value = Now_SPEC.ApproachTemp & "°C "
            worksheet.Cell("B16").Value = Now_SPEC.SecCoolantFilter & um_unit
            worksheet.Cell("B17").Value = Now_SPEC.ExternalDP & "psi"
            If Now_SPEC.PH60Hz Then worksheet.Cell("B19").Value = "v" Else worksheet.Cell("B19").Value = ""
            If Now_SPEC.PH50Hz Then worksheet.Cell("B20").Value = "v" Else worksheet.Cell("B20").Value = ""
            worksheet.Cell("B21").Value = Now_SPEC.MOCP '& "A"
            worksheet.Cell("B22").Value = Now_SPEC.FLA '& "A"
            If Now_SPEC.DualPowerFeed Then worksheet.Cell("B23").Value = "v" Else worksheet.Cell("B23").Value = ""
            If Now_SPEC.AutomaticTransferSwitch Then worksheet.Cell("B24").Value = "v" Else worksheet.Cell("B24").Value = ""
            If Now_SPEC.UPSbackuppower Then worksheet.Cell("B25").Value = "v" Else worksheet.Cell("B25").Value = ""
            If Now_SPEC.MCUControl Then worksheet.Cell("B26").Value = "v" Else worksheet.Cell("B26").Value = ""
            worksheet.Cell("B27").Value = Now_SPEC.PowerFeedLocation
            worksheet.Cell("B29").Value = Now_SPEC.PrimConnection
            worksheet.Cell("B30").Value = Now_SPEC.SecConnection
            worksheet.Cell("B31").Value = Now_SPEC.ConnectionLocation

            worksheet.Cell("B33").Value = Now_SPEC.Dimension
            worksheet.Cell("B34").Value = Now_SPEC.NetWeight & "kg"
            worksheet.Cell("B35").Value = Now_SPEC.OpWeight & "kg"
            worksheet.Cell("B37").Value = Now_SPEC.Display
            worksheet.Cell("B38").Value = Now_SPEC.Protocols

            Dim safety_li As (String, String, String) = (Now_SPEC.Safety_CE, Now_SPEC.Safety_UL, Now_SPEC.Safety_IEC)
            worksheet.Cell("B40").Value = Safety_Comma(safety_li)
            worksheet.Cell("B42").Value = Now_SPEC.SoundPressureLevel

            If Now_SPEC.LeakDetection Then worksheet.Cell("B43").Value = "v" Else worksheet.Cell("B43").Value = ""
            If Now_SPEC.DewPointMonitor Then worksheet.Cell("B44").Value = "v" Else worksheet.Cell("B44").Value = ""
            If Now_SPEC.ControlSensorRed Then worksheet.Cell("B45").Value = "v" Else worksheet.Cell("B45").Value = ""
            If Now_SPEC.VFD Then worksheet.Cell("B46").Value = "v" Else worksheet.Cell("B46").Value = ""
            If Now_SPEC.ExpantionVessel Then worksheet.Cell("B47").Value = "v" Else worksheet.Cell("B47").Value = ""
            If Now_SPEC.FillingTank Then worksheet.Cell("B48").Value = "v" Else worksheet.Cell("B48").Value = ""
            If Now_SPEC.AutoRestart Then worksheet.Cell("B49").Value = "v" Else worksheet.Cell("B49").Value = ""
            If Now_SPEC.PHSensor Then worksheet.Cell("B50").Value = "v" Else worksheet.Cell("B50").Value = ""
            If Now_SPEC.WaterLeaverSensor Then worksheet.Cell("B51").Value = "v" Else worksheet.Cell("B51").Value = ""
            worksheet.Cell("B52").Value = Now_SPEC.Pump
            worksheet.Cell("B53").Value = Now_SPEC.GroupControl

            'Auto-fit columns in ClosedXML
            worksheet.Columns().AdjustToContents()
            'Set page size to A4
            'worksheet.PrinterSettings.PaperSize = ePaperSize.A4

            ' Validate filename length
            If xlsxName.Length < 3 OrElse xlsxName.Length > 40 Then
                MessageBox.Show("Filename must be between 3 and 40 characters.", "Invalid Filename", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return ' Stop execution if filename is invalid
            End If

            xlsxName = Incremental_Fname(xlsxName)
            TextBox_ModelName.Text = xlsxName

            While HS_class.Name_HS.Contains(xlsxName)
                xlsxName = Incremental_Fname(xlsxName)
                TextBox_ModelName.Text = xlsxName
            End While

            'TextBox_ModelName.textchanged -> do Now_SPEC.ModelName = xlsxName
            worksheet.Cell("B3").Value = Now_SPEC.ModelName 'update again


            Dim saveFileDialog As New SaveFileDialog()
            ' 設定對話框標題
            saveFileDialog.Title = "選擇存檔位置"
            ' 限制存檔格式為 .xlsx
            saveFileDialog.Filter = "Excel 檔案 (*.xlsx)|*.xlsx"
            ' 預設檔名
            saveFileDialog.FileName = xlsxName & ".xlsx"
            ' 顯示對話框

            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                Dim excelPath As String = saveFileDialog.FileName
                ' 確保路徑有效
                If Not String.IsNullOrEmpty(excelPath) Then
                    workbook.SaveAs(excelPath)
                    MessageBox.Show("Excel 檔案已儲存：" & excelPath, "存檔成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If

                Dim result As DialogResult = MessageBox.Show("是否也要儲存於歷史資料庫?" & vbCrLf &
                                                             "(請確認檔案內容無誤)" & vbCrLf &
                                                             "Would you also like to save in history database ?" & vbCrLf &
                                                             "(make sure the content is valid, without further modification)", "Save as JSON", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If result = DialogResult.Yes Then

                    ' Read existing JSON file
                    Dim js_Text As String = File.ReadAllText(js_DB_Path)
                    ' Deserialize JSON into a list of Now_SPEC_C objects
                    Dim data_List As List(Of Now_SPEC_C) = JsonConvert.DeserializeObject(Of List(Of Now_SPEC_C))(js_Text)

                    data_List.Add(Now_SPEC)
                    ' Sort the list by DateName in descending order
                    data_List = data_List.OrderByDescending(Function(item)
                                                                Dim parsedDate As DateTime
                                                                If DateTime.TryParse(item.DateName, parsedDate) Then
                                                                    Return parsedDate
                                                                Else
                                                                    Return DateTime.MinValue ' Assign minimum date for invalid cases
                                                                End If
                                                            End Function).ToList()

                    Dim updated_Js_DB As String = JsonConvert.SerializeObject(data_List, Formatting.Indented)
                    File.WriteAllText(js_DB_Path, updated_Js_DB)


                    '=== write individual JSON in path js_DBdisc_Path ===

                    Dim js_ind_content As String = JsonConvert.SerializeObject(Now_SPEC, Formatting.Indented)
                    Dim js_ind_Path As String = js_DBdisc_Path & "\" & Now_SPEC.ModelName & ".json"
                    File.WriteAllText(js_ind_Path, js_ind_content)

                    MessageBox.Show("已儲存於歷史資料庫" & vbCrLf &
                                    "history database updated successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Output Excel Error : " & ex.Message)
        End Try
    End Sub

    Private Sub Combobox_CoCap_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combobox_CoCap.SelectedIndexChanged

        CoolingSelectResult()
        Checked60Hz50Hz()
    End Sub

    Private Sub Textbox_Para_TextChanged(sender As Object, e As EventArgs) Handles Textbox_Para.TextChanged
        CoolingSelectResult()
        Now_SPEC.IndexFactor = Textbox_Para.Text
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_Pump.SelectedIndexChanged

    End Sub


    Private Sub CheckBox_60Hz_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_60Hz.CheckedChanged
        Checked60Hz50Hz()
        Now_SPEC.PH60Hz = CheckBox_60Hz.Checked()
    End Sub

    Private Sub CheckBox_50Hz_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_50Hz.CheckedChanged
        Checked60Hz50Hz()
        Now_SPEC.PH50Hz = CheckBox_50Hz.Checked()
    End Sub

    ''' <summary>
    ''' 處理Checkbox事件
    ''' </summary>
    Private Sub Checked60Hz50Hz()
        If CheckBox_50Hz.Checked = True Or CheckBox_60Hz.Checked = True Then
            '判斷MOCP,FLA是否有值 & Cooling是否被選擇(正常情況Hz有值則Cooling一定會被選擇)
            If Now_SPEC.MOCP <> 0 And Now_SPEC.FLA <> 0 Then
                TextBox_MOCP.Text = Now_SPEC.MOCP & "A"
                TextBox_FLA.Text = Now_SPEC.FLA & "A"
            Else
                '規格參數沒有值
                TextBox_MOCP.Text = ""
                TextBox_FLA.Text = ""
            End If
        Else
            TextBox_MOCP.Text = ""
            TextBox_FLA.Text = ""
        End If
    End Sub

    Private Sub Combobox_Dimensions_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combobox_Dimensions.SelectedIndexChanged

        Dim dimension_list As List(Of String) = Combobox_Dimensions.DataSource

        If Combobox_Dimensions.SelectedIndex >= 0 Then

            '用Pump和FrameDimension篩選出唯一值
            If dimension_list.Count > 1 Then
                Dim pp As String = Now_SPEC.Pump
                Dim fd As String = Combobox_Dimensions.SelectedValue.ToString()

                Debug.Print("pp = " & pp)
                Debug.Print("fd = " & fd)

                ' 使用 LINQ 查詢符合條件的資料
                Dim selectedRow As DataRow = NowTable _
                    .AsEnumerable() _
                    .FirstOrDefault(Function(row) row("Pump Redundancy").ToString() = pp AndAlso row("Dimensions").ToString() = fd)

                If selectedRow IsNot Nothing Then
                    ' 更新 UI 或處理資料
                    Now_SPEC.Dimension = fd
                    Now_SPEC.NetWeight = selectedRow("Net Weight").ToString() + "kg"
                    Now_SPEC.OpWeight = selectedRow("Operation Weight").ToString() + "kg"
                    Textbox_NWeight.Text = Now_SPEC.NetWeight
                    Textbox_OWeight.Text = Now_SPEC.OpWeight
                End If
            End If

        End If

    End Sub

    Private Sub CheckBox_DualPF_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_DualPF.CheckedChanged
        Now_SPEC.DualPowerFeed = CheckBox_DualPF.Checked
    End Sub

    Private Sub CheckBox_AutomaticTS_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_AutomaticTS.CheckedChanged
        Now_SPEC.AutomaticTransferSwitch = CheckBox_AutomaticTS.Checked
    End Sub

    Private Sub CheckBox_MCUControl_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_MCUControl.CheckedChanged
        Now_SPEC.MCUControl = CheckBox_MCUControl.Checked
    End Sub

    Private Sub TextBox_ModelName_TextChanged(sender As Object, e As EventArgs) Handles TextBox_ModelName.TextChanged
        xlsxName = TextBox_ModelName.Text
        Now_SPEC.ModelName = xlsxName
    End Sub 'anchor




    Private Sub ComboBox_ApproachT_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_ApproachT.SelectedIndexChanged
        Dim ApproachTemp_list As List(Of Double) = ComboBox_ApproachT.DataSource
        Debug.Print("ApproachTemp_list.Count = " & ApproachTemp_list.Count)

        If ApproachTemp_list.Count >= 1 And ComboBox_ApproachT.SelectedIndex >= 0 Then

            Dim cc As String = Now_SPEC.CoolingCapacity
            Dim at As String = ComboBox_ApproachT.Text
            Dim scf As Double = Now_SPEC.SecCoolantFlow

            Dim selectedRow As DataRow

            Debug.Print("sfc = " & scf)
            '判斷sec. coolantFlow是否被輸入
            If scf <> 0 And ApproachTemp_list.Count > 1 Then
                ' 使用 LINQ 查詢符合條件的資料
                selectedRow = NowTable _
                    .AsEnumerable() _
                    .FirstOrDefault(Function(row) row("Cooling Capacity").ToString() = cc AndAlso row("Sec. Coolant Flow") = scf AndAlso row("Approach Temp").ToString() = at)
                If selectedRow IsNot Nothing Then
                    Now_SPEC.PrimCoolantFlow = selectedRow("Prim. Coolant Flow")
                Else
                    Return
                End If

            Else
                selectedRow = NowTable _
                    .AsEnumerable() _
                    .FirstOrDefault(Function(row) row("Cooling Capacity").ToString() = cc AndAlso row("Approach Temp").ToString() = at)
                Now_SPEC.PrimCoolantFlow = selectedRow("Prim. Coolant Flow")
            End If

            Now_SPEC.ApproachTemp = ComboBox_ApproachT.SelectedValue
            TextBox_PrimCF.Text = Now_SPEC.PrimCoolantFlow.ToString() & "LPM"

            Now_SPEC.OperatingDP = selectedRow("Operating DP")
            TextBox_OperatingDP.Text = Now_SPEC.OperatingDP.ToString() & "psi"

            Now_SPEC.ApproachTemp = ComboBox_ApproachT.SelectedValue
            Now_SPEC.PrimCoolantFlow = Now_SPEC.PrimCoolantFlow

            TextBox_PrimCF.Text = Now_SPEC.PrimCoolantFlow.ToString() & "LPM"

            '帶入Pump
            Now_SPEC.Pump = selectedRow("Pump Redundancy")
            ComboBox_Pump.SelectedText = Now_SPEC.Pump


            '---------------處理Pump--------------------
            ' 獲取所選的 Pump 值
            Dim selectedPump As String = Now_SPEC.Pump
            ' 從 DataTable 中篩選出與所選 Pump 中相同的 FrameDimension名稱
            Dim filteredRows As DataRow() = NowTable _
                .Select("[Pump Redundancy] = '" & selectedPump & "'")

            ' 將FrameDimension轉成Combobox可用的陣列
            Dim uniqueNames As List(Of String) = filteredRows _
            .Select(Function(row) row("Dimensions").ToString()) _
            .Distinct() _
            .ToList()

            '如果FrameDimension_WxDxH唯一則自動填入
            '如果FrameDimension_WxDxH有多個則讓使用者選Pump
            Now_SPEC.Pump = ComboBox_Pump.Text.ToString()

            If uniqueNames.Count = 1 Then
                Now_SPEC.Dimension = uniqueNames(0)
                ' 取得對應的資料列
                selectedRow = filteredRows(0)
                Now_SPEC.NetWeight = selectedRow("Net Weight")
                Now_SPEC.OpWeight = selectedRow("Operation Weight")

                Combobox_Dimensions.DataSource = uniqueNames
                Combobox_Dimensions.SelectedIndex = 0   '自動選擇首項
                Textbox_NWeight.Text = Now_SPEC.NetWeight.ToString() & "kg"
                Textbox_OWeight.Text = Now_SPEC.OpWeight.ToString() & "kg"
            Else
                Combobox_Dimensions.SelectedIndex = -1
                Combobox_Dimensions.DataSource = uniqueNames
            End If


            '---------------處理之後的Checkbox √--------------------
            If selectedRow("380-480V/3PH/60Hz") = "√" Then
                CheckBox_60Hz.Checked = True
            Else
                CheckBox_60Hz.Checked = False
            End If
            If selectedRow("380-480V/3PH/50Hz") = "√" Then
                CheckBox_50Hz.Checked = True
            Else
                CheckBox_50Hz.Checked = False
            End If
            If selectedRow("Dual Power Feed") = "√" Then
                CheckBox_DualPF.Checked = True
            Else
                CheckBox_DualPF.Checked = False
            End If
            If selectedRow("UPS Backup Power") = "√" Then
                CheckBox_UPS.Checked = True
            Else
                CheckBox_UPS.Checked = False
            End If
            If selectedRow("Automatic Transfer Switch") = "√" Then
                CheckBox_AutomaticTS.Checked = True
            Else
                CheckBox_AutomaticTS.Checked = False
            End If
            If selectedRow("MCU Controls") = "√" Then
                CheckBox_MCUControl.Checked = True
            Else
                CheckBox_MCUControl.Checked = False
            End If
            Dim safetyStr As String
            safetyStr = selectedRow("Safety")

            CheckBox_CE.Checked = (InStr(safetyStr, "CE") > 0)
            CheckBox_UL.Checked = (InStr(safetyStr, "UL/CSA 60335") > 0)
            CheckBox_IEC.Checked = (InStr(safetyStr, "IEC62368") > 0)

            If selectedRow("Leak Detection") = "√" Then
                CheckBox_LeakDetection.Checked = True
            Else
                CheckBox_LeakDetection.Checked = False
            End If
            If selectedRow("Filling Tank") = "√" Then
                CheckBox_FillingTank.Checked = True
            Else
                CheckBox_FillingTank.Checked = False
            End If
            If selectedRow("Dew Point Monitor") = "√" Then
                CheckBox_DewPointMonitor.Checked = True
            Else
                CheckBox_DewPointMonitor.Checked = False
            End If
            If selectedRow("Auto-restart") = "√" Then
                CheckBox_AutoRestart.Checked = True
            Else
                CheckBox_AutoRestart.Checked = False
            End If
            If selectedRow("Control Sensor Redundancy") = "√" Then
                CheckBox_ControlSensorRed.Checked = True
            Else
                CheckBox_ControlSensorRed.Checked = False
            End If
            If selectedRow("PH Sensor") = "√" Then
                CheckBox_PHSensor.Checked = True
            Else
                CheckBox_PHSensor.Checked = False
            End If
            If selectedRow("VFD") = "√" Then
                CheckBox_VFD.Checked = True
            Else
                CheckBox_VFD.Checked = False
            End If
            If selectedRow("Water Level Sensor") = "√" Then
                CheckBox_WaterLeverSensor.Checked = True
            Else
                CheckBox_WaterLeverSensor.Checked = False
            End If
            If selectedRow("Expansion Vessel") = "√" Then
                CheckBox_ExpantionVessel.Checked = True
            Else
                CheckBox_ExpantionVessel.Checked = False
            End If

            '---------------處理之後的Textbox--------------------
            Now_SPEC.SecCoolantType = selectedRow("Sec. Coolant Type")
            TextBox_SecCoolantType.Text = Now_SPEC.SecCoolantType

            Now_SPEC.Display = selectedRow("Display")
            TextBox_Display.Text = Now_SPEC.Display

            Now_SPEC.Protocols = selectedRow("Protocols")
            TextBox_Protocols.Text = Now_SPEC.Protocols

            Now_SPEC.PrimConnection = selectedRow("Prim. Connection")
            TextBox_PrimConnection.Text = Now_SPEC.PrimConnection

            Now_SPEC.SecConnection = selectedRow("Sec. Connection")
            TextBox_SecConnection.Text = Now_SPEC.SecConnection
        Else
            MessageBox.Show("Not found")
        End If
    End Sub


    ''' <summary>
    ''' External DP 計算公式
    ''' </summary>
    ''' <param name="flowRateLPM"></param>
    ''' <param name="diameterMM"></param>
    ''' <param name="k"></param>
    ''' <returns></returns>
    Public Function CalculateExternalDP(flowRateLPM As Double, diameterMM As Double, k As Double) As Double
        ' 將 LPM 轉換為 m³/s
        Dim flowRateM3S As Double = flowRateLPM / (1000 * 60)

        ' 將直徑 mm 轉換為 m
        Dim diameterM As Double = diameterMM / 1000

        ' 計算 External DP（壓力降）
        Dim externalDP As Double = k * (flowRateM3S ^ 2) / (diameterM ^ 4)

        ' 返回 External DP，單位 kPa
        Return externalDP / 1000 ' 轉換為 kPa
    End Function

    '--- Function for proper comma input for safety checkbox
    'ex CE,,IEC -> willbe CE,IEC
    Public Function Safety_Comma(inputTuple As (String, String, String)) As String
        ' Filter out empty or whitespace strings
        Dim filteredList As List(Of String) = New List(Of String) From {
        inputTuple.Item1.Trim(),
        inputTuple.Item2.Trim(),
        inputTuple.Item3.Trim()
    }
        ' Remove empty elements from the list
        filteredList.RemoveAll(Function(s) String.IsNullOrEmpty(s))
        ' Join remaining elements with commas
        Return String.Join(",", filteredList)
    End Function




    '=== incremental filename 進版號機制 
    'ex default -> default_v1, fname_v1 -> fname_v2, fname_v10 -> fname_v11
    Public Function Incremental_Fname(filename As String) As String
        ' Ensure the filename has at least 3 characters before checking
        If filename.Length < 3 OrElse (Not filename.Substring(filename.Length - 3, 2).Equals("_v", StringComparison.OrdinalIgnoreCase)) Then
            'ex default -> default_v1
            MessageBox.Show("為符合進版號規範，自動更新檔名為" & filename & "_v1")
            Return filename & "_v1"
        End If

        ' Define the pattern for _v followed by digits
        Dim versionPattern As String = "_v(\d+)$"
        Dim match As Match = Regex.Match(filename, versionPattern)

        If match.Success Then
            Dim currentVersion As Integer = Integer.Parse(match.Groups(1).Value)
            Dim newVersion As Integer = currentVersion + 1
            MessageBox.Show("檔名" & filename & "已存在歷史資料庫" & vbCrLf & "故檔名自動進版為" & "_v" & newVersion.ToString())
            Return Regex.Replace(filename, versionPattern, "_v" & newVersion.ToString())
        Else
            MessageBox.Show("為符合進版號規範，自動更新檔名為" & filename & "_v1")
            Return filename & "_v1"
        End If
    End Function
    '============


    Private Sub ComboBox_ExternalDP_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_ExternalDP.SelectedIndexChanged
        Now_SPEC.ExternalDP = ComboBox_ExternalDP.SelectedValue
    End Sub



    Private Sub Button_BuildDefault_Click(sender As Object, e As EventArgs) Handles Button_BuildDefault.Click

        NowTable = partSpec.New_Default_PartDatatable
        DataGridView1.DataSource = NowTable

        ' 遍歷 DataTable 的所有行
        For Each row As DataRow In NowTable.Rows
            ' 遍歷該行的所有欄位並列印欄位名稱和值
            Dim rowData As String = ""
            For Each column As DataColumn In NowTable.Columns
                rowData &= column.ColumnName & ": " & row(column).ToString() & "   "
            Next
            ' 使用 Debug.Print 輸出行的資料
            Debug.Print(rowData)
        Next

        SaveData(NowTable)
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub SaveData(dt As DataTable)
        ' 将DataTable保存为JSON文件
        Dim json As String = JsonConvert.SerializeObject(dt)
        File.WriteAllText(PartDataFilePath, json)
    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged

        Dim dt As DataTable = DataGridView1.DataSource
        NowTable = dt
        SaveData(dt)

    End Sub

    Private Sub CheckBox_UPS_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_UPS.CheckedChanged
        Now_SPEC.UPSbackuppower = CheckBox_UPS.Checked
    End Sub

    'Private Sub ComboBox_Display_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_Display.SelectedIndexChanged
    '    Now_SPEC.Display = ComboBox_Display.Text
    'End Sub

    'Private Sub ComboBox_Protocols_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_Protocols.SelectedIndexChanged
    '   Now_SPEC.Protocols = ComboBox_Protocols.Text
    'End Sub

    Private Sub CheckBox_CE_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_CE.CheckedChanged
        If CheckBox_CE.Checked Then
            Now_SPEC.Safety_CE = "CE"
        Else
            Now_SPEC.Safety_CE = ""
        End If
    End Sub

    Private Sub CheckBox_UL_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_UL.CheckedChanged
        If CheckBox_UL.Checked Then
            Now_SPEC.Safety_UL = "UL/CSA 60335"
        Else
            Now_SPEC.Safety_CE = ""
        End If
    End Sub

    Private Sub CheckBox_IEC_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_IEC.CheckedChanged
        If CheckBox_IEC.Checked Then
            Now_SPEC.Safety_IEC = "IEC62368"
        Else
            Now_SPEC.Safety_IEC = ""
        End If
    End Sub

    Private Sub CheckBox_LeakDetection_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_LeakDetection.CheckedChanged
        Now_SPEC.LeakDetection = CheckBox_LeakDetection.Checked
    End Sub

    Private Sub CheckBox_DewPointMonitor_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_DewPointMonitor.CheckedChanged
        Now_SPEC.DewPointMonitor = CheckBox_DewPointMonitor.Checked
    End Sub

    Private Sub CheckBox_ControlSensorRed_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_ControlSensorRed.CheckedChanged
        Now_SPEC.ControlSensorRed = CheckBox_ControlSensorRed.Checked
    End Sub

    Private Sub CheckBox_VFD_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_VFD.CheckedChanged
        Now_SPEC.VFD = CheckBox_VFD.Checked
    End Sub

    Private Sub CheckBox_ExpantionVessel_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_ExpantionVessel.CheckedChanged
        Now_SPEC.ExpantionVessel = CheckBox_ExpantionVessel.Checked
    End Sub

    Private Sub CheckBox_FillingTank_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_FillingTank.CheckedChanged
        Now_SPEC.FillingTank = CheckBox_FillingTank.Checked
    End Sub

    Private Sub CheckBoxAutoRestart_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_AutoRestart.CheckedChanged
        Now_SPEC.AutoRestart = CheckBox_AutoRestart.Checked
    End Sub

    Private Sub CheckBox_PHSensor_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_PHSensor.CheckedChanged
        Now_SPEC.PHSensor = CheckBox_PHSensor.Checked
    End Sub

    Private Sub CheckBox_WaterLeverSensor_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_WaterLeverSensor.CheckedChanged
        Now_SPEC.WaterLeaverSensor = CheckBox_WaterLeverSensor.Checked
    End Sub

    Private Sub GroupBox7_Enter(sender As Object, e As EventArgs) Handles GroupBox7.Enter
    End Sub

    'Private Sub ComboBox_PrimConnection_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_PrimConnection.SelectedIndexChanged
    '   Now_SPEC.PrimConnection = ComboBox_PrimConnection.Text
    'End Sub

    'Private Sub ComboBox_SecConnection_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_SecConnection.SelectedIndexChanged
    '   Now_SPEC.SecConnection = ComboBox_SecConnection.Text
    'End Sub

    Private Sub ComboBox_ConnectionLoaction_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_ConnectionLoaction.SelectedIndexChanged
        Now_SPEC.ConnectionLocation = ComboBox_ConnectionLoaction.Text
    End Sub

    'Private Sub ComboBox_SecCoolantType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_SecCoolantType.SelectedIndexChanged
    '   Now_SPEC.SecCoolantType = ComboBox_SecCoolantType.Text 'dont confuse .SelectedText
    'End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ClearAll()
    End Sub

    Private Sub ClearAll()
        Now_SPEC = New Now_SPEC_C()

        Dim cbx_group = {Combobox_CoCap, ComboBox_ApproachT, ComboBox_Pump, Combobox_Dimensions,
                        ComboBox_ExternalDP}
        For Each cbx In cbx_group
            cbx.SelectedIndex = -1
        Next

        Dim chbx_group = {CheckBox_60Hz, CheckBox_50Hz, CheckBox_DualPF, CheckBox_UPS, CheckBox_AutomaticTS,
                            CheckBox_MCUControl, CheckBox_CE, CheckBox_LeakDetection, CheckBox_DewPointMonitor, CheckBox_ControlSensorRed}
        For Each chbx In chbx_group
            chbx.Checked = False
        Next

        'TextBox_SecCoolantType, TextBox_Display, TextBox_Protocols, TextBox_PrimConnection

        Label65.Text = ""
        Label47.Text = ""

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim dt As DataTable = DataGridView1.DataSource
        NowTable = dt
        SaveData(dt)
    End Sub



    '=== TabPage3, import & search history report ===

    Private Sub Button_Search_Click(sender As Object, e As EventArgs) Handles Button_Search.Click

        FilterData()
    End Sub


    Private Sub FilterData()
        ' Get user input
        Dim startDate As DateTime
        Dim endDate As DateTime
        Dim data_List_2 As List(Of Now_SPEC_C)

        ' Load the JSON data
        If File.Exists(js_DB_Path) Then
            Dim js_Text_2 As String = File.ReadAllText(js_DB_Path)
            data_List_2 = JsonConvert.DeserializeObject(Of List(Of Now_SPEC_C))(js_Text_2)
        Else
            data_List_2 = New List(Of Now_SPEC_C)()
        End If

        ' Validate input
        If String.IsNullOrWhiteSpace(Tbx_DateStart.Text) OrElse String.IsNullOrWhiteSpace(Tbx_DateEnd.Text) Then
            MessageBox.Show("Both DateStart and DateEnd are required.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        If Not Date.TryParse(Tbx_DateStart.Text, startDate) OrElse Not Date.TryParse(Tbx_DateEnd.Text, endDate) Then
            MessageBox.Show("Invalid date format. Use YYYY-MM-DD.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        If Tbx_h_ModelName.Text.Length > 40 Then
            MessageBox.Show("ModelName must be within 40 characters.", "Invalid ModelName", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        If Tbx_h_OwnerName.Text.Length > 40 Then
            MessageBox.Show("OwnerName must be within 40 characters.", "Invalid OwnerName", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim h_ApproachT As Double
        Dim is_h_ApproachT As Boolean = Double.TryParse(Cbx_h_ApproachT.Text, h_ApproachT)

        ' Get  input (empty means ignore it)
        Dim h_ModelName As String = Tbx_h_ModelName.Text
        Dim h_OwnerName As String = Tbx_h_OwnerName.Text
        Dim h_SCFlow As String = Tbx_h_SCFlow.Text
        Dim h_CoCap As String = Cbx_h_CoCap.Text

        ' Perform filtering
        Dim filteredData = data_List_2.Where(Function(item)
                                                 ' Filter by date range
                                                 Dim itemDate As Date
                                                 If Not Date.TryParse(item.DateName, itemDate) Then Return False
                                                 Dim isDateInRange As Boolean = (itemDate >= startDate AndAlso itemDate <= endDate)

                                                 ' Filter by ModelName if entered
                                                 Dim isModelMatch As Boolean = (String.IsNullOrEmpty(h_ModelName) OrElse item.ModelName.Equals(h_ModelName, StringComparison.OrdinalIgnoreCase))
                                                 Dim isOwnerMatch As Boolean = (String.IsNullOrEmpty(h_OwnerName) OrElse item.OwnerName.Equals(h_OwnerName, StringComparison.OrdinalIgnoreCase))
                                                 Dim isSCFlowMatch As Boolean = (String.IsNullOrEmpty(h_SCFlow) OrElse item.SecCoolantFlow.Equals(h_SCFlow, StringComparison.OrdinalIgnoreCase))
                                                 Dim isCoCapMatch As Boolean = (String.IsNullOrEmpty(h_CoCap) OrElse item.CoolingCapacity.Equals(h_CoCap, StringComparison.OrdinalIgnoreCase))



                                                 ' Filter by Power if entered
                                                 Dim isTemplMatch As Boolean = (Not is_h_ApproachT OrElse item.ApproachTemp = h_ApproachT)
                                                 'Dim isTemplMatch As Boolean = (String.IsNullOrEmpty(h_ApproachT) OrElse item.ApproachTemp.Equals(h_ApproachT, StringComparison.OrdinalIgnoreCase))

                                                 Return isDateInRange AndAlso isModelMatch AndAlso isTemplMatch AndAlso isOwnerMatch AndAlso isSCFlowMatch AndAlso isCoCapMatch
                                             End Function).ToList()

        ' Display filtered data
        DisplayData(filteredData)
    End Sub

    Private Sub DisplayData(data_form As List(Of Now_SPEC_C))

        ' Clear previous data
        DataGridView2.DataSource = Nothing
        DataGridView2.Columns.Clear()

        DataGridView2.DataSource = data_form
        ' Define columns to keep and their header names
        Dim allowedColumns As Dictionary(Of String, String) = New Dictionary(Of String, String) From {
        {"ModelName", "ModelName (w/ ver)"},
        {"DateName", "Date"},
        {"OwnerName", "Owner"},
        {"CoolingCapacity", "CoolingCapacity(kW)"},
        {"SecCoolantFlow", "Sec. Coolant Flow(LPM)"},
        {"ApproachTemp", "Approach Temp(°C)"}
        }
        ' Loop through all columns in DataGridView
        For Each col As DataGridViewColumn In DataGridView2.Columns
            If allowedColumns.ContainsKey(col.Name) Then
                ' Set the header text for allowed columns
                col.HeaderText = allowedColumns(col.Name)
            Else
                ' Hide unwanted columns
                col.Visible = False
            End If
        Next

        ' Adjust column width for better display
        DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
    End Sub

    Private Sub Btn_h_recall_Click(sender As Object, e As EventArgs) Handles Btn_h_recall.Click

        ' Validate input
        If String.IsNullOrWhiteSpace(Tbx_h_singlefile.Text) Then
            MessageBox.Show("ModelName can't be empty.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        If Tbx_h_singlefile.Text.Length > 40 Then
            MessageBox.Show("ModelName must be within 40 characters.", "Invalid ModelName", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub ' Stop execution if filename is invalid         
        End If

        Dim recall_js_Path As String = js_DBdisc_Path & "\" & Tbx_h_singlefile.Text & ".json"

        ' Check if the file exists
        If File.Exists(recall_js_Path) Then
            Try
                Label65.Text = "Note: You're reviewing the setting of " & Tbx_h_singlefile.Text
                'Label65.ForeColor = Color.HotTrack
                Label47.Text = "Note: Please back to Page 1 interface "

                Dim recall_Data As String = File.ReadAllText(recall_js_Path)

                ' Deserialize JSON into a Dictionary
                Dim innerDict As Dictionary(Of String, String) = JsonConvert.DeserializeObject(Of Dictionary(Of String, String))(recall_Data)

                TextBox_ModelName.Text = innerDict("ModelName").ToString()
                TextBox_OwnerName.Text = innerDict("OwnerName").ToString()
                TextBox_DateName.Text = innerDict("DateName").ToString()

                Combobox_CoCap.Text = If(innerDict("CoolingCapacity") IsNot Nothing, innerDict("CoolingCapacity").ToString(), "")
                Textbox_Para.Text = If(innerDict("IndexFactor") IsNot Nothing, innerDict("IndexFactor").ToString(), "")
                TextBox_PrimCF.Text = innerDict("PrimCoolantFlow").ToString()
                '"OperatingDP" "19.8",
                TextBox_SecCoolantType.Text = If(innerDict("SecCoolantType") IsNot Nothing, innerDict("SecCoolantType").ToString(), "")
                ' "SecCoolantFlow": "1500kw",
                '"SecCoolantFilter": "500",
                ' "ExternalDP": "27",

                ComboBox_ApproachT.Text = If(innerDict("ApproachTemp") IsNot Nothing, innerDict("ApproachTemp").ToString(), "")
                ComboBox_Pump.Text = If(innerDict("Pump") IsNot Nothing, innerDict("Pump").ToString(), "")
                Combobox_Dimensions.Text = If(innerDict("Dimension") IsNot Nothing, innerDict("Dimension").ToString(), "")
                Textbox_NWeight.Text = If(innerDict("NetWeight") IsNot Nothing, innerDict("NetWeight").ToString(), "")
                Textbox_OWeight.Text = If(innerDict("OpWeight") IsNot Nothing, innerDict("OpWeight").ToString(), "")

                CheckBox_60Hz.Checked = innerDict("PH60Hz").ToString()
                CheckBox_50Hz.Checked = innerDict("PH50Hz").ToString()
                TextBox_MOCP.Text = If(innerDict("MOCP") IsNot Nothing, innerDict("MOCP").ToString(), "")
                TextBox_FLA.Text = If(innerDict("FLA") IsNot Nothing, innerDict("FLA").ToString(), "")
                CheckBox_DualPF.Checked = innerDict("DualPowerFeed").ToString()
                CheckBox_AutomaticTS.Checked = innerDict("AutomaticTransferSwitch").ToString()
                CheckBox_UPS.Checked = innerDict("UPSbackuppower").ToString()
                CheckBox_MCUControl.Checked = innerDict("MCUControl").ToString()
                '  "PowerFeedLocation": "Top",
                CheckBox_CE.Checked = Not String.IsNullOrEmpty(innerDict("Safety_CE").ToString())
                CheckBox_UL.Checked = Not String.IsNullOrEmpty(innerDict("Safety_UL").ToString())
                CheckBox_IEC.Checked = Not String.IsNullOrEmpty(innerDict("Safety_IEC").ToString())
                '"SoundPressureLevel": "76dBA",
                TextBox_Display.Text = If(innerDict("Display") IsNot Nothing, innerDict("Display").ToString(), "")
                TextBox_Protocols.Text = If(innerDict("Protocols") IsNot Nothing, innerDict("Protocols").ToString(), "")
                CheckBox_LeakDetection.Checked = innerDict("LeakDetection").ToString()
                CheckBox_DewPointMonitor.Checked = innerDict("DewPointMonitor").ToString()
                CheckBox_ControlSensorRed.Checked = innerDict("ControlSensorRed").ToString()
                CheckBox_VFD.Checked = innerDict("VFD").ToString()
                CheckBox_ExpantionVessel.Checked = innerDict("ExpantionVessel").ToString()
                CheckBox_PHSensor.Checked = innerDict("PHSensor").ToString()
                CheckBox_AutoRestart.Checked = innerDict("AutoRestart").ToString()
                CheckBox_FillingTank.Checked = innerDict("FillingTank").ToString()
                CheckBox_WaterLeverSensor.Checked = innerDict("WaterLeaverSensor").ToString()

                TextBox_PrimConnection.Text = If(innerDict("PrimConnection") IsNot Nothing, innerDict("PrimConnection").ToString(), "")
                TextBox_SecConnection.Text = If(innerDict("SecConnection") IsNot Nothing, innerDict("SecConnection").ToString(), "")
                ComboBox_ConnectionLoaction.Text = If(innerDict("ConnectionLocation") IsNot Nothing, innerDict("ConnectionLocation").ToString(), "")
                '"GroupControl" 8.0,

            Catch ex As Exception
                MessageBox.Show("Error reading JSON file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else
            MessageBox.Show(Tbx_h_singlefile.Text & " not found in history database!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

    End Sub

    Private Sub TabPage3_Click(sender As Object, e As EventArgs) Handles TabPage3.Click

    End Sub
    '======
End Class

Public Class PartSpecification

    Function New_Default_PartDatatable() As DataTable
        ' 创建DataTable
        Dim dt As New DataTable("CDU_Technical_Specifications")

        dt.Columns.Add("Series", GetType(String))
        dt.Columns.Add("Cooling Capacity", GetType(Double))
        dt.Columns.Add("Model", GetType(String))
        dt.Columns.Add("Prim. Coolant Type", GetType(String))
        dt.Columns.Add("Prim. Coolant Flow", GetType(Double))
        dt.Columns.Add("Operating DP", GetType(Double))
        dt.Columns.Add("Prim Coolant Filter", GetType(String))
        dt.Columns.Add("Sec. Coolant Type", GetType(String))
        dt.Columns.Add("Sec. Coolant Flow", GetType(Double))
        dt.Columns.Add("Approach Temp", GetType(Double))
        dt.Columns.Add("Sec. Coolant Filter", GetType(String))
        dt.Columns.Add("External DP", GetType(Double))
        dt.Columns.Add("380-480V/3PH/60Hz", GetType(String))
        dt.Columns.Add("380-480V/3PH/50Hz", GetType(String))
        dt.Columns.Add("MOCP", GetType(Double))
        dt.Columns.Add("FLA", GetType(Double))
        dt.Columns.Add("Dual Power Feed", GetType(String))
        dt.Columns.Add("Automatic Transfer Switch", GetType(String))
        dt.Columns.Add("UPS Backup Power", GetType(String))
        dt.Columns.Add("MCU Controls", GetType(String))
        dt.Columns.Add("Power Feed Location", GetType(String))
        dt.Columns.Add("Prim. Connection", GetType(String))
        dt.Columns.Add("Sec. Connection", GetType(String))
        dt.Columns.Add("Dimensions", GetType(String))
        dt.Columns.Add("Net Weight", GetType(Double))
        dt.Columns.Add("Operation Weight", GetType(Double))
        dt.Columns.Add("Display", GetType(String))
        dt.Columns.Add("Protocols", GetType(String))
        dt.Columns.Add("Safety", GetType(String))
        dt.Columns.Add("Sound Pressure Level", GetType(Double))
        dt.Columns.Add("Leak Detection", GetType(String))
        dt.Columns.Add("Dew Point Monitor", GetType(String))
        dt.Columns.Add("Control Sensor Redundancy", GetType(String))
        dt.Columns.Add("VFD", GetType(String))
        dt.Columns.Add("Expansion Vessel", GetType(String))
        dt.Columns.Add("Filling Tank", GetType(String))
        dt.Columns.Add("Auto-restart", GetType(String))
        dt.Columns.Add("PH Sensor", GetType(String))
        dt.Columns.Add("Water Level Sensor", GetType(String))
        dt.Columns.Add("Pump Redundancy", GetType(String))
        dt.Columns.Add("Group Control", GetType(Double))

        dt.Rows.Add("SuperCDU-500", 500, "RDF400eSVS19070", "Water", 660, 22.0, "500μm", "25XPG/DI Water", 750, 5.6, "50μm (Option:25μm)", 18.0, "√", "√", 17, 12.5, "√", "√", "√", "√", "Top", "2.5 in. Victaulic Coupling", "3.0 in. Victaulic Coupling", "600x1350x2100mm", 675, 800, "10"" touch panel", "Modbus RTU, Modbus TCP, SNMP, BACnet", "CE", 72, "√", "√", "√", "√", "√", "√", "√", "N/A", "√", "1x pump run modes", 8)
        dt.Rows.Add("SuperCDU-1000", 800, "CLSC0554A7000-0", "Water", 1200, 19.6, "500μm", "25XPG/DI Water", 1200, 4.5, "50μm (Option:25μm)", 40.5, "√", "√", 60, 40, "√", "√", "√", "√", "Top", "4 in. Victaulic Coupling", "4 in. Victaulic Coupling", "900x1200x2300 mm", 1400, 1800, "10"" touch panel", "Modbus RTU, Modbus TCP, SNMP, BACnet", "CE", 74, "√", "√", "√", "√", "√", "√", "√", "N/A", "√", "1x pump (N+1) 2x pump (N+1)", 8)
        dt.Rows.Add("SuperCDU-1000", 1000, "RDF0516D7519035", "Water", 1200, 19.6, "500μm", "25XPG", 1500, 6, "50μm (Option:25μm)", 27.0, "√", "√", 60, 42, "√", "√", "N/A", "√", "Top", "4 in. sanitary ferrule", "4 in. sanitary ferrule", "900x1200x2300 mm", 1400, 1800, "10"" touch panel", "Modbus RTU, Modbus TCP, SNMP, BACnet", "CE, UL/CSA 60335", 75, "√", "√", "√", "√", "√", "√", "√", "√", "√", "1x pump (N+1) 2x pump (N+1)", 8)
        dt.Rows.Add("SuperCDU-1200", 1000, "RDF0516D7519035", "Water", 1350, 19.6, "500μm", "25XPG", 1500, 4, "50μm (Option:25μm)", 27.0, "√", "√", 60, 42, "√", "√", "N/A", "√", "Top", "4 in. sanitary ferrule", "4 in. sanitary ferrule", "900x1200x2300 mm", 1200, 1450, "10"" touch panel", "Modbus RTU, Modbus TCP, SNMP, BACnet", "CE, UL/CSA 60335", 75, "√", "√", "√", "√", "√", "√", "√", "√", "√", "2x pump (N+1) 3x pump run modes", 8)
        dt.Rows.Add("SuperCDU-1200", 1200, "RDF1036D1519044", "Water", 1160, 26.2, "500μm", "25XPG", 1200, 5, "50μm (Option:25μm)", 30.0, "√", "√", 70, 54, "√", "√", "N/A", "√", "Top", "4 in. sanitary ferrule", "4 in. sanitary ferrule", "900x1200x2300 mm", 1200, 1450, "10"" touch panel", "Modbus RTU, Modbus TCP, SNMP, BACnet", "CE, IEC62368, UL/CSA 60335", 76, "√", "√", "√", "√", "√", "√", "√", "√", "√", "2x pump (N+1) 3x pump run modes", 8)
        dt.Rows.Add("SuperCDU-1500", 1035, "RDF1036D1519044", "Water", 1200, 17.6, "500μm", "25XPG", 1500, 4, "50μm (Option:25μm)", 46.0, "√", "√", 70, 54, "√", "√", "N/A", "√", "Top", "4 in. sanitary ferrule", "4 in. sanitary ferrule", "1200x1200x2300 mm", 1600, 1900, "10"" touch panel", "Modbus RTU, Modbus TCP, SNMP, BACnet", "CE, IEC62368, UL/CSA 60335", 76, "√", "√", "√", "√", "√", "√", "√", "√", "√", "3x pump (N+1) 4x pump run mode", 8)
        dt.Rows.Add("SuperCDU-1500", 1377, "RDF1036D1519044", "Water", 1200, 13.5, "500μm", "25XPG", 1200, 4, "50μm (Option:25μm)", 49.0, "√", "√", 70, 54, "√", "√", "N/A", "√", "Top", "4 in. sanitary ferrule", "4 in. sanitary ferrule", "1200x1200x2300 mm", 1600, 1900, "10"" touch panel", "Modbus RTU, Modbus TCP, SNMP, BACnet", "CE, IEC62368, UL/CSA 60335", 76, "√", "√", "√", "√", "√", "√", "√", "√", "√", "3x pump (N+1) 4x pump run mode", 8)
        dt.Rows.Add("SuperCDU-1500", 1500, "RDF1036D1519044", "Water", 1350, 19.4, "500μm", "25XPG", 1500, 5, "50μm (Option:25μm)", 49.0, "√", "√", 70, 54, "√", "√", "N/A", "√", "Top", "4 in. sanitary ferrule", "4 in. sanitary ferrule", "1200x1200x2300 mm", 1600, 1900, "10"" touch panel", "Modbus RTU, Modbus TCP, SNMP, BACnet", "CE, IEC62368, UL/CSA 60335", 76, "√", "√", "√", "√", "√", "√", "√", "√", "√", "3x pump (N+1) 4x pump run mode", 8)

        ' 输出DataTable内容
        For Each row As DataRow In dt.Rows
            For Each col As DataColumn In dt.Columns
                Console.Write(row(col).ToString() & vbTab)
            Next
            Console.WriteLine()
        Next
        Return dt
    End Function

End Class

Public Class Now_SPEC_C
    Public Property ModelName As String = ""
    Public Property OwnerName As String = ""
    Public Property DateName As String = ""
    Public Property CoolingCapacity As String = ""
    '---Primary Side---
    Public Property IndexFactor As String = ""
    Public Property PrimaryWater As String = "Water"
    Public Property PrimCoolantType As String = "Water"
    Public Property PrimCoolantFlow As Integer
    Public Property OperatingDP As String = ""
    Public Property PrimCoolantFilter As String = "50"
    '---Secondary Side---
    Public Property SecCoolantType As String = ""
    Public Property SecCoolantFlow As String 'Double
    Public Property ApproachTemp As Double
    Public Property SecCoolantFilter As String = "500"
    Public Property ExternalDP As String = ""


    'power supply
    Public Property PH60Hz As Boolean = False
    Public Property PH50Hz As Boolean = False
    Public Property MOCP As String = ""
    Public Property FLA As String = ""
    Public Property DualPowerFeed As Boolean = False
    Public Property AutomaticTransferSwitch As Boolean = False
    Public Property UPSbackuppower As Boolean = False
    Public Property MCUControl As Boolean = False
    Public Property PowerFeedLocation As String = "Top"

    'deployment
    Public Property PrimConnection As String = ""
    Public Property SecConnection As String = ""
    Public Property ConnectionLocation As String = ""

    'physical
    Public Property Dimension As String = ""
    Public Property NetWeight As String = ""
    Public Property OpWeight As String = ""


    Public Property Display As String = ""
    Public Property Protocols As String = ""
    Public Property Safety_CE As String = ""
    Public Property Safety_UL As String = ""
    Public Property Safety_IEC As String = ""

    'features
    Public Property SoundPressureLevel As String = "<76dBA"
    Public Property LeakDetection As Boolean = False
    Public Property DewPointMonitor As Boolean = False
    Public Property ControlSensorRed As Boolean = False
    Public Property VFD As Boolean = False
    Public Property ExpantionVessel As Boolean = False
    Public Property FillingTank As Boolean = False
    Public Property AutoRestart As Boolean = False
    Public Property PHSensor As Boolean = False
    Public Property WaterLeaverSensor As Boolean = False
    Public Property Pump As String = ""
    Public Property GroupControl As Double = 8


    '測試用
    Public Sub PrintNowSpecValues(obj As Now_SPEC_C)
        Dim properties = obj.GetType().GetProperties()
        For Each prop In properties
            Dim value = prop.GetValue(obj, Nothing)
            Debug.Print($"{prop.Name}: {If(value Is Nothing, "Nothing", value.ToString())}")
        Next
    End Sub

End Class



'all history report name will be in Name_HS,
'for quickly checking non-repeated new filename
Public Class HS_Clss

    Public Property Name_HS As HashSet(Of String)

    ' Constructor to initialize the HashSet (optional)
    Public Sub New()
        Name_HS = New HashSet(Of String)()
    End Sub
End Class

