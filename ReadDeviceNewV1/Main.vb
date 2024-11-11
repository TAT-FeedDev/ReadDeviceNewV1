
Imports System.Data.SqlClient
Imports System.IO
Imports OfficeOpenXml
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.Configuration

Public Class Main
    Dim connString As String
    Dim MotorConfig_ As MotorConfig
    Dim ScadaForAnalog_ As ScadaForAnalog
    Dim BinParameter_ As BinParameter
    Dim ScadaForAlarm_ As ScadaForAlarm
    Dim ScaleParameter_ As ScaleParameter
    Dim AnalogCurrentConfig_ As AnalogCurrentConfig
    Dim AnalogCalibrationConfig_ As AnalogCalibrationConfig
    Dim nPlcNo As Int16

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadConfig()
        Dim dbManager As New DatabaseManager(connString)
        dbManager.LoadDatabases(connString, cmdDatabaseBatching)
        dbManager.LoadDatabases(connString, cmdDatabaseRoute)
    End Sub
    Private Sub LoadConfig()
        ' อ่านค่าการเชื่อมต่อจาก app.config
        Dim connectionStringBatching As String = ConfigurationManager.ConnectionStrings("MyDatabaseConnectionBatching").ConnectionString
        Dim connectionStringRoute As String = ConfigurationManager.ConnectionStrings("MyDatabaseConnectionRoute").ConnectionString

        ' แยกข้อมูลใน connection string
        Dim builderBatching As New System.Data.SqlClient.SqlConnectionStringBuilder(connectionStringBatching)
        Dim builderRoute As New System.Data.SqlClient.SqlConnectionStringBuilder(connectionStringRoute)

        ' นำค่ามาแสดงใน TextBox ต่าง ๆ
        txtServer.Text = builderBatching.DataSource
        txtUsername.Text = builderBatching.UserID
        txtPassword.Text = builderBatching.Password ' อาจไม่ต้องแสดง Password ใน TextBox
        cmdDatabaseBatching.Text = builderBatching.InitialCatalog
        cmdDatabaseRoute.Text = builderRoute.InitialCatalog

        connString = $"Data Source={builderBatching.DataSource};User ID={builderBatching.UserID};Password={ builderBatching.Password}"
    End Sub

    Private Sub btnSaveConfigDb_Click(sender As Object, e As EventArgs) Handles btnSaveConfigDb.Click
        Dim server As String = txtServer.Text
        Dim databaseBatching As String = cmdDatabaseBatching.Text
        Dim databaseRoute As String = cmdDatabaseRoute.Text
        Dim username As String = txtUsername.Text
        Dim password As String = txtPassword.Text

        ' สร้าง connection string
        Dim connStringBatching As String = $"Data Source={server};Initial Catalog={databaseBatching};User ID={username};Password={password}"
        Dim connStringRoute As String = $"Data Source={server};Initial Catalog={databaseRoute};User ID={username};Password={password}"

        ' เขียนค่าลง app.config
        Dim configuration As Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
        configuration.ConnectionStrings.ConnectionStrings("MyDatabaseConnectionBatching").ConnectionString = connStringBatching
        configuration.ConnectionStrings.ConnectionStrings("MyDatabaseConnectionRoute").ConnectionString = connStringRoute
        configuration.Save(ConfigurationSaveMode.Modified)
        ConfigurationManager.RefreshSection("connectionStrings")
        LoadConfig()
        UpdateStatus("Configuration saved successfully.")
    End Sub

    Private Sub cmdDatabases_MouseUp(sender As Object, e As MouseEventArgs) Handles cmdDatabaseBatching.MouseUp, cmdDatabaseRoute.MouseUp
        btnSaveConfigDb_Click(sender, e)
        Dim dbManager As New DatabaseManager(connString)
        dbManager.LoadDatabases(connString, cmdDatabaseBatching)
        dbManager.LoadDatabases(connString, cmdDatabaseRoute)
    End Sub

    Private Async Sub btnLoadExcel_Click(sender As Object, e As EventArgs) Handles btnLoadExcel.Click
        ' Clear existing tabs
        TabControl1.TabPages.Clear()

        ' Open a file dialog to select an Excel file
        Using openFileDialog As New OpenFileDialog()
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xls;*.xlsm"
            If openFileDialog.ShowDialog() = DialogResult.OK Then
                Dim filePath As String = openFileDialog.FileName
                UpdateStatus("Loading data, please wait...")
                tsProgressBar.Value = 0 ' เริ่มต้น ProgressBar
                tsProgressBar.Visible = True
                Await LoadExcelDataAsync(filePath)
                UpdateStatus("Data loaded successfully!")
            End If
        End Using
    End Sub

    Private Async Function LoadExcelDataAsync(filePath As String) As Task
        Try
            Dim worksheetsData As New List(Of (Name As String, Data As DataTable))()
            Await Task.Run(Sub()
                               ' Load the Excel file in a background task
                               Dim fileInfo As New FileInfo(filePath)
                               Using package As New ExcelPackage(fileInfo)
                                   ' Loop through all worksheets in the package
                                   Dim totalWorksheets As Integer = package.Workbook.Worksheets.Count
                                   For i As Integer = 0 To totalWorksheets - 1
                                       Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(i)

                                       If worksheet IsNot Nothing Then
                                           Try
                                               ' Create a new DataTable for the current worksheet
                                               Dim dataTable As DataTable = ReadWorksheetToDataTable(worksheet)
                                               worksheetsData.Add((worksheet.Name, dataTable))  ' Add worksheet name and DataTable to the list
                                           Catch ex As Exception
                                               ' Log the error or inform the user that the worksheet is empty
                                               LogError(ex)
                                               UpdateStatus($"Worksheet '{worksheet.Name}' is empty or not defined.")
                                           End Try
                                       End If

                                       ' Update status on UI thread
                                       Dim percentage As Integer = CInt((i + 1) * 100 / totalWorksheets)
                                       UpdateStatus($"Loading worksheet {worksheet.Name}... {percentage}%")
                                       ' ตรวจสอบให้แน่ใจว่าค่าอยู่ในช่วง 0 - 100
                                       UpdateProgressBar(Math.Min(percentage, 100)) ' ไม่ให้ค่าเกิน 100
                                   Next
                               End Using
                           End Sub)

            ' Update UI on the UI thread
            For Each ws In worksheetsData
                AddTabForWorksheet(ws.Name, ws.Data)
            Next
        Catch ex As Exception
            ' หากเกิด Exception ให้บันทึกข้อผิดพลาด
            LogError(ex)
            UpdateStatus("Error occurred while loading data. Check log for details.")
        End Try
    End Function

    Private Sub UpdateStatus(statusMessage As String)
        If Me.InvokeRequired Then
            Me.Invoke(Sub() UpdateStatus(statusMessage))
        Else
            tsStatus.Text = statusMessage
        End If
    End Sub

    Private Sub UpdateProgressBar(value As Integer)
        If Me.InvokeRequired Then
            Me.Invoke(Sub() UpdateProgressBar(value))
        Else
            ' ตรวจสอบให้แน่ใจค่าที่ตั้งอยู่ในช่วงที่ถูกต้อง
            If value < tsProgressBar.Minimum Then value = tsProgressBar.Minimum
            If value > tsProgressBar.Maximum Then value = tsProgressBar.Maximum

            tsProgressBar.Value = value
        End If
    End Sub
    Private Sub AddTabForWorksheet(worksheetName As String, dataTable As DataTable)
        If TabControl1.InvokeRequired Then
            TabControl1.Invoke(Sub() AddTabForWorksheet(worksheetName, dataTable))
            Return
        End If

        ' Create a new TabPage for the current worksheet
        Dim tabPage As New TabPage(worksheetName)
        Dim dataGridView As New DataGridView With {
           .DataSource = dataTable,
           .Dock = DockStyle.Fill,
           .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
           .ScrollBars = ScrollBars.Both ' Ensure scrollbars are enabled
        }

        ' Add the DataGridView to the TabPage
        tabPage.Controls.Add(dataGridView)

        ' Add the TabPage to the TabControl
        TabControl1.TabPages.Add(tabPage)
    End Sub

    Private Function ReadWorksheetToDataTable(worksheet As ExcelWorksheet) As DataTable
        Dim dataTable As New DataTable()

        ' ตรวจสอบว่า Dimension ไม่เป็น Nothing
        If worksheet.Dimension Is Nothing Then
            Throw New InvalidOperationException("The worksheet is empty or not properly defined.")
        End If

        ' ชุดข้อมูลเพื่อเก็บคอลัมน์ที่มีอยู่
        Dim columnNames As New HashSet(Of String)()

        ' Add the header row to the DataTable
        For col As Integer = 1 To worksheet.Dimension.End.Column
            Dim columnName As String = worksheet.Cells(1, col).Text

            ' ตรวจสอบว่า columnName มีอยู่ใน HashSet หรือไม่
            Dim uniqueColumnName As String = columnName
            Dim counter As Integer = 1

            ' สร้างชื่อที่ไม่ซ้ำกันสำหรับคอลัมน์
            While columnNames.Contains(uniqueColumnName)
                uniqueColumnName = $"{columnName}_{counter}"
                counter += 1
            End While

            ' เพิ่มชื่อคอลัมน์ที่ไม่ซ้ำลงใน DataTable และ HashSet
            dataTable.Columns.Add(uniqueColumnName)
            columnNames.Add(uniqueColumnName)
        Next

        ' Add the data rows to the DataTable
        For row As Integer = 2 To worksheet.Dimension.End.Row
            Dim dataRow As DataRow = dataTable.NewRow()
            For col As Integer = 1 To worksheet.Dimension.End.Column
                dataRow(col - 1) = worksheet.Cells(row, col).Text
            Next
            dataTable.Rows.Add(dataRow)
        Next

        Return dataTable
    End Function

    Private Sub BtnUpdateDatabase_Click(sender As Object, e As EventArgs) Handles btnUpdateDatabase.Click
        If String.IsNullOrEmpty(cboStation.Text) Then
            MessageBox.Show("Please select PLC No.")
            cboStation.Focus()
            Exit Sub
        End If
        nPlcNo = Convert.ToInt16(cboStation.Text)
        ' ตรวจสอบว่ามี TabPage ที่เลือกอยู่
        If TabControl1.SelectedTab IsNot Nothing Then
            Select Case TabControl1.SelectedTab.Text
                Case "NEW SCADA RELAY"
                    NewScadaRelay()

                Case "SCADA REG FOR ANALOG", "SCADA REG FOR ANALOG (R5)"
                    ScadaRegForAnalog()

                Case "BIN PARAMETER"
                    BinParameter()

                Case "SCADA REG FOR ALARM", "SCADA REG FOR ALARM (UDP)"
                    ScadaRegForAlarm(5)

                Case "SCADA REG FOR ALARM LD&PK (UDP)"
                    ScadaRegForAlarm(2)

                Case "ANALOG BUCKET PARAMITER"
                    AnalogCurrentConfig("BUCKET", 48)

                Case "ANALOG GRINDING PARAMITER"
                    AnalogCurrentConfig("GRINDING", 23)
                    Threading.Thread.Sleep(2000)
                    AnalogCurrentConfig("PID", 23)

                Case "ANALOG MIXER PARAMITER"
                    AnalogCurrentConfig("MIXER", 14)

                Case "ANALOG LIQUID PARAMITER"
                    AnalogCurrentConfig("LIQUID", 34)

                Case "ANALOG PELLET PARAMITER"
                    AnalogCurrentConfig("PELLET", 14)

                Case "ANALOG COOLER PARAMITER"
                    AnalogCurrentConfig("COOLER", 14)

                Case "ANALOG BLOWER PARAMITER"
                    AnalogCurrentConfig("BLOWER", 80)

                Case "ANALOG GENERAL PARAMITER "
                    AnalogCurrentConfig("ANALOGSENSOR", 48)

                    'Case "Scale Parameter"
                    '    ScaleParameter()

                    'Case "Mixer Parameter"
                    '    MixerParameter()

                Case Else
                    MessageBox.Show("Please select a tab ?")
            End Select
        Else
            MessageBox.Show("Please select a tab to save data.")
        End If
    End Sub
    Private Function StartsCheck(value As String) As Boolean
        ' ตรวจสอบ 3 เงื่อนไข และคืนค่า True หากข้อมูลเริ่มต้นด้วย 'R', 'ZR', หรือ 'M'
        Return value.StartsWith("R") OrElse value.StartsWith("ZR") OrElse value.StartsWith("M")
    End Function
    Private Async Sub NewScadaRelay()
        Dim searchWords() As String = {"MOTOR", "SLIDE", "KNOCKER", "FLAP", "DEVICE"}
        Dim selectedTab As TabPage = TabControl1.SelectedTab
        ' หา DataGridView ที่อยู่ใน TabPage ที่เลือก
        Dim dataGridView As DataGridView = CType(selectedTab.Controls(0), DataGridView)

        ' เชื่อมต่อกับ SQL Server
        Dim connectionString As String = ConfigurationManager.ConnectionStrings("MyDatabaseConnectionRoute").ConnectionString
        Dim dbManager As New DatabaseManager(connectionString)
        UpdateStatus("Saving data, please wait...")
        tsProgressBar.Value = 0
        tsProgressBar.Visible = True
        ' ทำการบันทึกข้อมูลใน background task
        Await Task.Run(Sub()
                           Dim totalRows As Integer = dataGridView.Rows.Count - 1  ' Exclude new row
                           Dim currentRow As Integer = 0
                           Dim jsonData As String
                           For Each row As DataGridViewRow In dataGridView.Rows
                               If Not row.IsNewRow Then
                                   For Each searchWord As String In searchWords
                                       If InStr(row.Cells(1).Value.ToString, searchWord, vbTextCompare) > 0 Then
                                           If ContainsNumber(row.Cells(1).Value.ToString()) Then
                                               With MotorConfig_
                                                   .motor_name = row.Cells(1).Value.ToString().Replace(" ", "_").Trim
                                                   .c_tat_code = row.Cells(2).Value.ToString().Trim
                                                   .motor_code = row.Cells(5).Value.ToString().Trim
                                                   .c_description = row.Cells(3).Value.ToString().Trim
                                                   .mservice = CheckNullOrEmpty(row.Cells(7).Value.ToString().Trim, "M0")
                                                   .mout = CheckNullOrEmpty(row.Cells(8).Value.ToString().Trim, "M0")
                                                   .mauto = CheckNullOrEmpty(row.Cells(9).Value.ToString().Trim, "M0")
                                                   .mrun = CheckNullOrEmpty(row.Cells(10).Value.ToString().Trim, "M0")
                                                   .mrun2 = CheckNullOrEmpty(row.Cells(11).Value.ToString().Trim, "M0")
                                                   .merr = CheckNullOrEmpty(row.Cells(12).Value.ToString().Trim, "M0")
                                                   .merr1 = CheckNullOrEmpty(row.Cells(13).Value.ToString().Trim, "M0")
                                                   .mcoverlock = CheckNullOrEmpty(row.Cells(14).Value.ToString().Trim, "M0")
                                                   .int_srt = CheckNullOrEmpty(row.Cells(16).Value.ToString().Trim, "M0")
                                                   .int_stp = CheckNullOrEmpty(row.Cells(17).Value.ToString().Trim, "M0")
                                                   .delay_open = CheckNullOrEmpty(row.Cells(19).Value.ToString().Trim, "R0")
                                                   .delay_close = CheckNullOrEmpty(row.Cells(20).Value.ToString().Trim, "R0")
                                                   .delay_step = CheckNullOrEmpty(row.Cells(21).Value.ToString().Trim, "R0")
                                                   .delay_on = ""
                                                   .delay_off = ""
                                                   .flag_status_run = "0"
                                                   .c_flag_enable_1 = CheckNullOrEmpty(row.Cells(22).Value.ToString().Trim, "R0")
                                                   .c_flag_enable_2 = CheckNullOrEmpty(row.Cells(23).Value.ToString().Trim, "R0")
                                                   .n_plc_station = nPlcNo
                                                   Select Case searchWord
                                                       Case "MOTOR", "SLIDE", "KNOCKER"
                                                           .motor_type_id = "Motor"
                                                           .n_first_motor = 1
                                                           .n_time_step = 30
                                                           .n_priority = 1
                                                       Case "FLAP"
                                                           .motor_type_id = "Wayer"
                                                           .n_first_motor = 2
                                                           .n_time_step = 0
                                                           .n_priority = 1
                                                       Case Else
                                                           .motor_type_id = "Device"
                                                           .n_first_motor = 0
                                                           .n_time_step = 0
                                                           .n_priority = 0
                                                   End Select
                                                   If StartsCheck(row.Cells(54).Value.ToString().Trim) Then
                                                       .c_run_hour = CheckNullOrEmpty(row.Cells(54).Value.ToString().Trim, "ZR0")
                                                   Else
                                                       .c_run_hour = "ZR0"
                                                   End If
                                                   If StartsCheck(row.Cells(55).Value.ToString().Trim) Then
                                                       .c_counter_times = CheckNullOrEmpty(row.Cells(55).Value.ToString().Trim, "ZR0")
                                                   Else
                                                       .c_counter_times = "ZR0"
                                                   End If
                                                   'If StartsCheck(row.Cells(55).Value.ToString().Trim) Then
                                                   '    .c_target_run_hour = CheckNullOrEmpty(row.Cells(55).Value.ToString().Trim, "ZR0")
                                                   'Else
                                                   '    .c_target_run_hour = "ZR0"
                                                   'End If
                                                   'If StartsCheck(row.Cells(57).Value.ToString().Trim) Then
                                                   '    .c_target_counter_times = CheckNullOrEmpty(row.Cells(57).Value.ToString().Trim, "R0")
                                                   'Else
                                                   '    .c_target_counter_times = "R0"
                                                   'End If
                                                   .c_target_run_hour = "ZR0"
                                                   .c_target_counter_times = "ZR0"
                                                   If StartsCheck(row.Cells(56).Value.ToString().Trim) Then
                                                       .c_pm_alarm = CheckNullOrEmpty(row.Cells(56).Value.ToString().Trim, "M0")
                                                   Else
                                                       .c_pm_alarm = "M0"
                                                   End If
                                                   '.c_run_hour = CheckNullOrEmpty(row.Cells(54).Value.ToString().Trim, "R0")
                                                   '.c_counter_times = CheckNullOrEmpty(row.Cells(55).Value.ToString().Trim, "R0")
                                                   '.c_target_run_hour = CheckNullOrEmpty(row.Cells(56).Value.ToString().Trim, "R0")
                                                   '.c_target_counter_times = CheckNullOrEmpty(row.Cells(57).Value.ToString().Trim, "R0")
                                                   '.c_pm_alarm = CheckNullOrEmpty(row.Cells(58).Value.ToString().Trim, "M0")
                                                   jsonData = "{""ALM1"":""" & CheckNullOrEmpty(row.Cells(58).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM2"":""" & CheckNullOrEmpty(row.Cells(59).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM3"":""" & CheckNullOrEmpty(row.Cells(60).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM4"":""" & CheckNullOrEmpty(row.Cells(61).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM5"":""" & CheckNullOrEmpty(row.Cells(62).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM6"":""" & CheckNullOrEmpty(row.Cells(63).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM7"":""" & CheckNullOrEmpty(row.Cells(64).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM8"":""" & CheckNullOrEmpty(row.Cells(65).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM9"":""" & CheckNullOrEmpty(row.Cells(66).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM10"":""" & CheckNullOrEmpty(row.Cells(67).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM11"":""" & CheckNullOrEmpty(row.Cells(68).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM12"":""" & CheckNullOrEmpty(row.Cells(69).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM13"":""" & CheckNullOrEmpty(row.Cells(70).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM14"":""" & CheckNullOrEmpty(row.Cells(71).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM15"":""" & CheckNullOrEmpty(row.Cells(72).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM16"":""" & CheckNullOrEmpty(row.Cells(73).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM17"":""" & CheckNullOrEmpty(row.Cells(74).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM18"":""" & CheckNullOrEmpty(row.Cells(75).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM19"":""" & CheckNullOrEmpty(row.Cells(76).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM20"":""" & CheckNullOrEmpty(row.Cells(77).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM21"":""" & CheckNullOrEmpty(row.Cells(78).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM22"":""" & CheckNullOrEmpty(row.Cells(79).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM23"":""" & CheckNullOrEmpty(row.Cells(80).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM24"":""" & CheckNullOrEmpty(row.Cells(81).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM25"":""" & CheckNullOrEmpty(row.Cells(82).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM26"":""" & CheckNullOrEmpty(row.Cells(83).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM27"":""" & CheckNullOrEmpty(row.Cells(84).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM28"":""" & CheckNullOrEmpty(row.Cells(85).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM29"":""" & CheckNullOrEmpty(row.Cells(86).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM30"":""" & CheckNullOrEmpty(row.Cells(87).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM31"":""" & CheckNullOrEmpty(row.Cells(88).Value.ToString().Trim, "M0") & ""","
                                                   jsonData &= """ALM32"":""" & CheckNullOrEmpty(row.Cells(89).Value.ToString().Trim, "M0") & """}"
                                                   .c_alarm_message = jsonData
                                                   .c_code_ref = row.Cells(4).Value.ToString().Trim
                                               End With
                                               dbManager.SaveDataScadaRelay(MotorConfig_)
                                           End If
                                       End If
                                   Next
                                   currentRow += 1
                                   Dim progressPercentage As Integer = CInt((currentRow * 100) / totalRows) ' Calculate progress
                                   UpdateStatus($"Saving data... {progressPercentage}%")  ' Update status in status strip
                                   ' ตรวจสอบให้แน่ใจว่าค่าอยู่ในช่วง 0 - 100
                                   UpdateProgressBar(Math.Min(progressPercentage, 100)) ' ไม่ให้ค่าเกิน 100
                               End If
                           Next
                       End Sub)
        'MessageBox.Show("Data saved successfully!")
        UpdateStatus("Data saved successfully!") ' Update status message after saving
    End Sub

    Private Async Sub ScadaRegForAnalog()
        Dim selectedTab As TabPage = TabControl1.SelectedTab
        ' หา DataGridView ที่อยู่ใน TabPage ที่เลือก
        Dim dataGridView As DataGridView = CType(selectedTab.Controls(0), DataGridView)

        ' เชื่อมต่อกับ SQL Server
        Dim connectionString As String = ConfigurationManager.ConnectionStrings("MyDatabaseConnectionBatching").ConnectionString
        Dim dbManager As New DatabaseManager(connectionString)
        UpdateStatus("Saving data, please wait...")
        tsProgressBar.Value = 0
        tsProgressBar.Visible = True
        ' ทำการบันทึกข้อมูลใน background task
        Await Task.Run(Sub()
                           Dim totalRows As Integer = dataGridView.Rows.Count - 1  ' Exclude new row
                           Dim currentRow As Integer = 0
                           Dim Multi As Double = 1
                           Dim tmpNumber As Int16 = 0
                           Dim tmpName As String
                           For Each row As DataGridViewRow In dataGridView.Rows
                               If row.Index >= 63 Then
                                   If Not row.IsNewRow Then
                                       If InStr(row.Cells(15).Value.ToString, "ZR4", vbTextCompare) > 0 Then
                                           If ContainsNumber(row.Cells(15).Value.ToString()) Then
                                               With ScadaForAnalog_
                                                   tmpName = CheckNullOrEmpty(row.Cells(2).Value.ToString().Trim, "")
                                                   If tmpName <> "" Then .motor_name = tmpName
                                                   .motor_code = CheckNullOrEmpty(row.Cells(13).Value.ToString().Trim, "")
                                                   tmpNumber = CheckNullOrEmpty(row.Cells(0).Value.ToString().Trim, "0")
                                                   If tmpNumber > 0 Then .n_number = tmpNumber
                                                   .r_address = CheckNullOrEmpty(row.Cells(15).Value.ToString().Trim, "ZR0")
                                                   .n_plc_station = nPlcNo
                                                   .c_description = CheckNullOrEmpty(row.Cells(5).Value.ToString().Trim, "")
                                                   .c_short_description = CheckNullOrEmpty(row.Cells(6).Value.ToString().Trim, "")
                                                   Multi = Convert.ToDouble(CheckNullOrEmpty(row.Cells(7).Value.ToString().Trim, "0")) / Convert.ToDouble(CheckNullOrEmpty(row.Cells(8).Value.ToString().Trim, "0"))
                                                   If Multi > 0.999 Then
                                                       .n_multiply = Multi
                                                   Else
                                                       .n_multiply = 1
                                                   End If
                                                   If InStr(row.Cells(6).Value.ToString().Trim, "CURRENT") > 0 Or InStr(row.Cells(6).Value.ToString().Trim, "(PV)") > 0 Or InStr(row.Cells(6).Value.ToString().Trim, "(SV)") > 0 Then
                                                       .n_priority = 1
                                                       .c_record = "Y"
                                                   ElseIf InStr(row.Cells(6).Value.ToString().Trim, "RPM") > 0 Then
                                                       .n_priority = 1
                                                       .c_record = "Y"
                                                   ElseIf InStr(row.Cells(6).Value.ToString().Trim, "TEMP") > 0 Then
                                                       .n_priority = 1
                                                       .c_record = "Y"
                                                   ElseIf InStr(row.Cells(6).Value.ToString().Trim, "(MV)") > 0 Then
                                                       .n_priority = 1
                                                       .c_record = "Y"
                                                   Else
                                                       .n_priority = 0
                                                       .c_record = "N"
                                                   End If
                                                   .c_format = CheckNullOrEmpty(row.Cells(10).Value.ToString().Trim, "N")
                                                   .c_tat_code = CheckNullOrEmpty(row.Cells(4).Value.ToString().Trim, "")
                                                   If .c_tat_code = "" Then
                                                       .c_tat_code = .motor_code
                                                   End If
                                                   .c_type = CheckNullOrEmpty(row.Cells(11).Value.ToString().Trim, "Int")
                                               End With
                                               dbManager.SaveDataAnalog(ScadaForAnalog_)
                                           End If
                                       End If
                                       currentRow += 1
                                       Dim progressPercentage As Integer = CInt((currentRow * 100) / totalRows) ' Calculate progress
                                       UpdateStatus($"Saving data... {progressPercentage}%")  ' Update status in status strip
                                       ' ตรวจสอบให้แน่ใจว่าค่าอยู่ในช่วง 0 - 100
                                       UpdateProgressBar(Math.Min(progressPercentage, 100)) ' ไม่ให้ค่าเกิน 100
                                   End If
                               End If
                           Next
                       End Sub)
        'MessageBox.Show("Data saved successfully!")
        UpdateStatus("Data saved successfully!") ' Update status message after saving
    End Sub

    Private Async Sub BinParameter()
        Dim selectedTab As TabPage = TabControl1.SelectedTab
        ' หา DataGridView ที่อยู่ใน TabPage ที่เลือก
        Dim dataGridView As DataGridView = CType(selectedTab.Controls(0), DataGridView)

        ' เชื่อมต่อกับ SQL Server
        Dim connectionString As String = ConfigurationManager.ConnectionStrings("MyDatabaseConnectionBatching").ConnectionString
        Dim dbManager As New DatabaseManager(connectionString)
        UpdateStatus("Saving data, please wait...")
        tsProgressBar.Value = 0
        tsProgressBar.Visible = True
        ' ทำการบันทึกข้อมูลใน background task
        Await Task.Run(Sub()
                           Dim totalCols As Integer = 250
                           Dim currentCol As Integer = 0
                           Dim tmpScaleNo As Int16 = 0
                           Dim delimiter As Char() = {" "c, "/"c}
                           Dim tmpBinName As String()
                           For i As Int16 = 5 To 255
                               If InStr(dataGridView.Rows(0).Cells(i).OwningColumn.DataPropertyName.ToString, "BIN", vbTextCompare) > 0 Then

                                   With BinParameter_
                                       .bin_name = CheckNullOrEmpty(dataGridView.Rows(0).Cells(i).OwningColumn.DataPropertyName.ToString, "")
                                       tmpBinName = .bin_name.Split(delimiter, StringSplitOptions.RemoveEmptyEntries)
                                       If tmpBinName.Length > 0 Then
                                           .bin_name = tmpBinName(tmpBinName.Length - 1)
                                       End If
                                       .bin_code = CheckNullOrEmpty(dataGridView.Rows(2).Cells(i).Value.ToString, "X")
                                       .d_address = CheckNullOrEmpty(dataGridView.Rows(4).Cells(i).Value.ToString, "D0")
                                       .d_address_extend = CheckNullOrEmpty(dataGridView.Rows(39).Cells(i).Value.ToString, "D0")
                                       .bin_index = CheckNullOrEmpty(ExtractNumbers(dataGridView.Rows(1).Cells(i).Value.ToString), "0")
                                       .bin_no = CheckNullOrEmpty(ExtractNumbers(dataGridView.Rows(3).Cells(i).Value.ToString), "0")
                                       .scale_no = CheckNullOrEmpty(ExtractNumbers(dataGridView.Rows(0).Cells(i).Value.ToString), "0")
                                       If .scale_no = 0 Then
                                           .scale_no = tmpScaleNo
                                       Else
                                           tmpScaleNo = .scale_no
                                       End If
                                       .c_location = ""
                                       .n_plc_station = nPlcNo
                                   End With
                                   dbManager.SaveDataBinParameter(BinParameter_)
                               End If
                               currentCol += 1
                               Dim progressPercentage As Integer = CInt((currentCol * 100) / totalCols) ' Calculate progress
                               UpdateStatus($"Saving data... {progressPercentage}%")  ' Update status in status strip
                               ' ตรวจสอบให้แน่ใจว่าค่าอยู่ในช่วง 0 - 100
                               UpdateProgressBar(Math.Min(progressPercentage, 100)) ' ไม่ให้ค่าเกิน 100
                           Next

                       End Sub)
        'MessageBox.Show("Data saved successfully!")
        UpdateStatus("Data saved successfully!") ' Update status message after saving
    End Sub

    Private Async Sub ScadaRegForAlarm(iCount As Int16)
        Dim searchWords() As String = {"SCALE", "MIXER", "HAND", "SURG", "LIQUID", "LOAD", "PACKING"}
        Dim selectedTab As TabPage = TabControl1.SelectedTab
        ' หา DataGridView ที่อยู่ใน TabPage ที่เลือก
        Dim dataGridView As DataGridView = CType(selectedTab.Controls(0), DataGridView)

        UpdateStatus("Saving data, please wait...")
        tsProgressBar.Value = 0
        tsProgressBar.Visible = True
        ' ทำการบันทึกข้อมูลใน background task
        Await Task.Run(Sub()
                           Dim totalRows As Integer = dataGridView.Rows.Count - 1  ' Exclude new row
                           Dim currentRow As Integer = 0
                           For i As Int16 = 0 To iCount
                               For Each row As DataGridViewRow In dataGridView.Rows
                                   If row.Index >= 22 AndAlso Not row.IsNewRow Then
                                       ProcessRow(row, i, searchWords.ToList())
                                   End If
                                   currentRow += 1
                               Next
                               Dim progressPercentage As Integer = CInt((currentRow * 100) / totalRows) ' Calculate progress
                               UpdateStatus($"Saving data... {progressPercentage}%")  ' Update status in status strip
                               ' ตรวจสอบให้แน่ใจว่าค่าอยู่ในช่วง 0 - 100
                               UpdateProgressBar(Math.Min(progressPercentage, 100)) ' ไม่ให้ค่าเกิน 100
                           Next
                       End Sub)
        UpdateStatus("Data saved successfully!") ' Update status message after saving
    End Sub

    Private Sub ProcessRow(row As DataGridViewRow, index As Int16, searchWords As List(Of String))
        For Each searchWord As String In searchWords
            If InStr(row.Cells(3 + (index * 9)).Value?.ToString(), searchWord, vbTextCompare) > 0 Then
                If InStr(row.Cells(6 + (index * 9)).Value?.ToString(), "ZR", vbTextCompare) > 0 Then
                    If InStr(row.Cells(4 + (index * 9)).Value?.ToString(), "ALARM CODE", vbTextCompare) > 0 Then
                        If ContainsNumber(row.Cells(3 + (index * 9)).Value.ToString()) Then
                            CreateAlarm(row, index, searchWord)
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub CreateAlarm(row As DataGridViewRow, index As Int16, searchWord As String)
        ' เชื่อมต่อกับ SQL Server
        Dim connectionString As String = ConfigurationManager.ConnectionStrings("MyDatabaseConnectionBatching").ConnectionString
        Dim dbManager As New DatabaseManager(connectionString)
        With ScadaForAlarm_
            Dim extractedNumber As String = CheckNullOrEmpty(ExtractNumbers(row.Cells(3 + (index * 9)).Value.ToString().Trim), "0")
            Select Case searchWord
                Case "SCALE"
                    .alarm_name = "SCALE_" & extractedNumber
                    .c_alarm_name_temp = "SCALE " & extractedNumber

                Case "MIXER"
                    .alarm_name = "MIXER_" & extractedNumber
                    .c_alarm_name_temp = "MIXER " & extractedNumber

                Case "HAND"
                    .alarm_name = "HANDADD_" & extractedNumber
                    .c_alarm_name_temp = "HANDADD " & extractedNumber

                Case "SURG"
                    .alarm_name = "SURGEBIN_" & extractedNumber
                    .c_alarm_name_temp = "SURGEBIN " & extractedNumber

                Case "LIQUID"
                    .alarm_name = "LIQUID_" & extractedNumber
                    .c_alarm_name_temp = "LIQUID " & extractedNumber

                Case "LOAD"
                    .alarm_name = "LOADOUT_" & extractedNumber
                    .c_alarm_name_temp = "LOADOUT " & extractedNumber

                Case "PACKING"
                    .alarm_name = "PACKING_" & extractedNumber
                    .c_alarm_name_temp = "PACKING " & extractedNumber

                Case Else
                    ' หากไม่ต้องการจัดการกับ case อื่น ๆ สามารถทำอะไรได้ที่นี่
            End Select
            .zr_address = CheckNullOrEmpty(row.Cells(6 + (index * 9)).Value.ToString().Trim, "R0")
            .c_location = "BATCHING 1"
            .n_plc_station = nPlcNo
        End With
        dbManager.SaveDataForAlarm(ScadaForAlarm_)
    End Sub

    Private Async Sub ScaleParameter()
        Dim selectedTab As TabPage = TabControl1.SelectedTab
        ' หา DataGridView ที่อยู่ใน TabPage ที่เลือก
        Dim dataGridView As DataGridView = CType(selectedTab.Controls(0), DataGridView)
        Dim SystemTimeDataGridView As DataGridView = Nothing
        ' ค้นหาที่ TabPage ด้วยชื่อที่ระบุ
        For Each tabPage As TabPage In TabControl1.TabPages
            If tabPage.Text = "PLC System Timer" Then
                ' รับ DataGridView ที่อยู่ภายใน TabPage นั้น
                SystemTimeDataGridView = CType(tabPage.Controls(0), DataGridView) ' แทนที่ 0 ด้วยตำแหน่งหาก DataGridView ไม่อยู่ตำแหน่งแรก
                Exit For
            End If
        Next
        ' เชื่อมต่อกับ SQL Server
        Dim connectionString As String = ConfigurationManager.ConnectionStrings("MyDatabaseConnectionBatching").ConnectionString
        Dim dbManager As New DatabaseManager(connectionString)
        UpdateStatus("Saving data, please wait...")
        tsProgressBar.Value = 0
        tsProgressBar.Visible = True
        ' ทำการบันทึกข้อมูลใน background task
        Await Task.Run(Sub()
                           Dim totalCols As Integer = 64
                           Dim currentCol As Integer = 10
                           For i As Int16 = 10 To totalCols
                               Dim scaleColumnName As String = dataGridView.Rows(0).Cells(i).OwningColumn.DataPropertyName.ToString
                               If InStr(scaleColumnName, "SCALE", vbTextCompare) > 0 Then
                                   With ScaleParameter_
                                       .scale_name = CheckNullOrEmpty(scaleColumnName.Replace("-", "_").Trim, "")
                                       .scale_code = CheckNullOrEmpty(dataGridView.Rows(1).Cells(i).Value.ToString.Trim, "D0")
                                       .d_address = CheckNullOrEmpty(dataGridView.Rows(3).Cells(i).Value.ToString.Trim, "D0")
                                       .scale_no = CheckNullOrEmpty(ExtractNumbers(dataGridView.Rows(1).Cells(i).Value.ToString.Trim), "0")

                                       ' กำหนดเวลา Time jogging ตามเงื่อนไข
                                       .time_joging = CheckNullOrEmpty(GetTimeJogging(SystemTimeDataGridView, i), "T0")

                                       .d_jog_time = CheckNullOrEmpty(dataGridView.Rows(189).Cells(i).Value.ToString, "D0")
                                       .d_extend = CheckNullOrEmpty(dataGridView.Rows(29).Cells(i).Value.ToString, "D0")
                                       .d_operate = CheckNullOrEmpty(dataGridView.Rows(137).Cells(i).Value.ToString, "D0")
                                       .c_location = ""
                                       .n_plc_station = nPlcNo
                                   End With
                                   dbManager.SaveDataScaleParameter(ScaleParameter_)
                               End If
                               currentCol += 1
                               Dim progressPercentage As Integer = CInt((currentCol * 100) / totalCols) ' Calculate progress
                               UpdateStatus($"Saving data... {progressPercentage}%")  ' Update status in status strip
                               ' ตรวจสอบให้แน่ใจว่าค่าอยู่ในช่วง 0 - 100
                               UpdateProgressBar(Math.Min(progressPercentage, 100)) ' ไม่ให้ค่าเกิน 100
                           Next
                       End Sub)
        'MessageBox.Show("Data saved successfully!")
        UpdateStatus("Data saved successfully!") ' Update status message after saving
    End Sub

    Function GetTimeJogging(dataGridView As DataGridView, index As Int16) As String
        For j As Int16 = 6 To 35
            Dim cellValue As String = dataGridView.Rows(178).Cells(j).Value.ToString.Trim

            ' ตรวจสอบว่ามีคีย์เวิร์ด LOAD หรือ PACK ในแถว
            If IsLoadOrPackType(cellValue) AndAlso
           ContainsNumber(cellValue) AndAlso
           ScaleParameter_.scale_name = NormalizeScaleName(cellValue) Then
                Return dataGridView.Rows(196).Cells(j).Value.ToString.Trim
            End If

            ' ตรวจสอบว่ามีคีย์เวิร์ด SCALE ในอีกแถวหนึ่ง
            cellValue = dataGridView.Rows(22).Cells(j).Value.ToString.Trim

            If InStr(cellValue, "SCALE", vbTextCompare) > 0 AndAlso
           ContainsNumber(cellValue) AndAlso
           ScaleParameter_.scale_name = NormalizeScaleName(cellValue) Then
                Return dataGridView.Rows(40).Cells(j).Value.ToString.Trim
            End If
        Next
        Return String.Empty
    End Function

    Function IsLoadOrPackType(cellValue As String) As Boolean
        Return InStr(cellValue, "LOAD", vbTextCompare) > 0 Or
           InStr(cellValue, "PACK", vbTextCompare) > 0
    End Function

    Function NormalizeScaleName(scaleName As String) As String
        Return scaleName.Replace("#", "_").Replace(" ", "").Replace("OUT", "").Trim
    End Function

    Private Async Sub MixerParameter()
        Dim selectedTab As TabPage = TabControl1.SelectedTab
        ' หา DataGridView ที่อยู่ใน TabPage ที่เลือก
        Dim dataGridView As DataGridView = CType(selectedTab.Controls(0), DataGridView)
        ' เชื่อมต่อกับ SQL Server
        Dim connectionString As String = ConfigurationManager.ConnectionStrings("MyDatabaseConnectionBatching").ConnectionString
        Dim dbManager As New DatabaseManager(connectionString)
        UpdateStatus("Saving data, please wait...")
        tsProgressBar.Value = 0
        tsProgressBar.Visible = True
        ' ทำการบันทึกข้อมูลใน background task
        Await Task.Run(Sub()
                           Dim totalCols As Integer = 64
                           Dim currentCol As Integer = 10
                           For i As Int16 = 10 To totalCols
                               Dim scaleColumnName As String = dataGridView.Rows(0).Cells(i).OwningColumn.DataPropertyName.ToString
                               If InStr(scaleColumnName, "SCALE", vbTextCompare) > 0 Then
                                   With ScaleParameter_
                                       .scale_name = CheckNullOrEmpty(scaleColumnName.Replace("-", "_").Trim, "")
                                       .scale_code = CheckNullOrEmpty(dataGridView.Rows(1).Cells(i).Value.ToString.Trim, "D0")
                                       .d_address = CheckNullOrEmpty(dataGridView.Rows(3).Cells(i).Value.ToString.Trim, "D0")
                                       .scale_no = CheckNullOrEmpty(ExtractNumbers(dataGridView.Rows(1).Cells(i).Value.ToString.Trim), "0")
                                       .d_jog_time = CheckNullOrEmpty(dataGridView.Rows(189).Cells(i).Value.ToString, "D0")
                                       .d_extend = CheckNullOrEmpty(dataGridView.Rows(29).Cells(i).Value.ToString, "D0")
                                       .d_operate = CheckNullOrEmpty(dataGridView.Rows(137).Cells(i).Value.ToString, "D0")
                                       .c_location = ""
                                       .n_plc_station = nPlcNo
                                   End With
                                   dbManager.SaveDataScaleParameter(ScaleParameter_)
                               End If
                               currentCol += 1
                               Dim progressPercentage As Integer = CInt((currentCol * 100) / totalCols) ' Calculate progress
                               UpdateStatus($"Saving data... {progressPercentage}%")  ' Update status in status strip
                               ' ตรวจสอบให้แน่ใจว่าค่าอยู่ในช่วง 0 - 100
                               UpdateProgressBar(Math.Min(progressPercentage, 100)) ' ไม่ให้ค่าเกิน 100
                           Next
                       End Sub)
        'MessageBox.Show("Data saved successfully!")
        UpdateStatus("Data saved successfully!") ' Update status message after saving
    End Sub

    Private Async Sub AnalogCurrentConfig(sName As String, nTotalNo As Int16)
        Dim selectedTab As TabPage = TabControl1.SelectedTab
        ' หา DataGridView ที่อยู่ใน TabPage ที่เลือก
        Dim dataGridView As DataGridView = CType(selectedTab.Controls(0), DataGridView)
        ' เชื่อมต่อกับ SQL Server
        Dim connectionString As String = ConfigurationManager.ConnectionStrings("MyDatabaseConnectionBatching").ConnectionString
        Dim dbManager As New DatabaseManager(connectionString)
        UpdateStatus("Saving data, please wait...")
        tsProgressBar.Value = 0
        tsProgressBar.Visible = True
        ' ทำการบันทึกข้อมูลใน background task
        Await Task.Run(Sub()
                           Dim totalCols As Integer = nTotalNo
                           Dim currentCol As Integer = 9
                           Dim iNumber As Int16 = 1
                           For i As Int16 = 9 To totalCols
                               Dim BuckColumnName As String = dataGridView.Rows(0).Cells(i).OwningColumn.DataPropertyName.ToString
                               If InStr(BuckColumnName, sName, vbTextCompare) > 0 Or sName = "PID" Then
                                   With AnalogCurrentConfig_
                                       .motor_name = CheckNullOrEmpty(BuckColumnName.Replace("#", "_").Replace(" ", ""), "N")
                                       Select Case sName
                                           Case "BUCKET"
                                               .c_tat_code = "BK" & iNumber.ToString("D2")

                                           Case "GRINDING"
                                               .c_tat_code = "HM-GD" & iNumber.ToString("D1")

                                           Case "PID"
                                               .motor_name = "PID" & iNumber.ToString("D1")
                                               .c_tat_code = "HM-GD" & iNumber.ToString("D1")

                                           Case "MIXER"
                                               .c_tat_code = "MIX" & iNumber.ToString("D1") & "-1"

                                           Case "LIQUID"
                                               .c_tat_code = "PUMP-LQ" & iNumber.ToString("D1")

                                           Case "PELLET"
                                               .c_tat_code = "MAIN-PL" & iNumber.ToString("D1")

                                           Case "COOLER"
                                               .c_tat_code = "MAIN-PL" & iNumber.ToString("D1")

                                           Case "BLOWER"
                                               .c_tat_code = "BW" & iNumber.ToString("D1")

                                           Case "ANALOGSENSOR"
                                               .c_tat_code = "INPUT" & iNumber.ToString("D1")

                                           Case Else
                                               .c_tat_code = "N"
                                       End Select
                                       .motor_code = CheckNullOrEmpty(dataGridView.Rows(1).Cells(i).Value.ToString.Trim, "N")
                                       If sName = "PID" Then
                                           .d_address = CheckNullOrEmpty(dataGridView.Rows(208).Cells(i).Value.ToString.Trim, "D0")
                                       Else
                                           .d_address = CheckNullOrEmpty(dataGridView.Rows(3).Cells(i).Value.ToString.Trim, "D0")
                                       End If
                                       .n_plc_station = nPlcNo
                                       .dt_lastupdate = Now
                                   End With
                                   dbManager.SaveDataAnalogConfig(AnalogCurrentConfig_)
                                   Select Case sName
                                       Case "BUCKET"
                                           AnalogCalibrationConfig(dataGridView, AnalogCurrentConfig_.motor_name, AnalogCurrentConfig_.c_tat_code, iNumber, 109, 12)
                                       Case "GRINDING"
                                           AnalogCalibrationConfig(dataGridView, AnalogCurrentConfig_.motor_name, AnalogCurrentConfig_.c_tat_code, iNumber, 81, 8)
                                       Case "MIXER"
                                           AnalogCalibrationConfig(dataGridView, AnalogCurrentConfig_.motor_name, AnalogCurrentConfig_.c_tat_code, iNumber, 55, 2)
                                       Case "LIQUID"
                                           AnalogCalibrationConfig(dataGridView, AnalogCurrentConfig_.motor_name, AnalogCurrentConfig_.c_tat_code, iNumber, 81, 4)
                                       Case "PELLET"
                                           AnalogCalibrationConfig(dataGridView, AnalogCurrentConfig_.motor_name, AnalogCurrentConfig_.c_tat_code, iNumber, 200, 14)
                                       Case "COOLER"
                                           AnalogCalibrationConfig(dataGridView, AnalogCurrentConfig_.motor_name, AnalogCurrentConfig_.c_tat_code, iNumber, 127, 9)
                                       Case "BLOWER"
                                           AnalogCalibrationConfig(dataGridView, AnalogCurrentConfig_.motor_name, AnalogCurrentConfig_.c_tat_code, iNumber, 26, 1)
                                       Case "ANALOGSENSOR"
                                           AnalogCalibrationConfig(dataGridView, AnalogCurrentConfig_.motor_name, AnalogCurrentConfig_.c_tat_code, iNumber, 38, 1)
                                   End Select
                                   iNumber += 1
                               End If
                               currentCol += 1
                               Dim progressPercentage As Integer = CInt((currentCol * 100) / totalCols) ' Calculate progress
                               UpdateStatus($"Saving data... {progressPercentage}%")  ' Update status in status strip
                               ' ตรวจสอบให้แน่ใจว่าค่าอยู่ในช่วง 0 - 100
                               UpdateProgressBar(Math.Min(progressPercentage, 100)) ' ไม่ให้ค่าเกิน 100
                           Next
                       End Sub)
        'MessageBox.Show("Data saved successfully!")
        UpdateStatus("Data saved successfully!") ' Update status message after saving
    End Sub

    Private Async Sub AnalogCalibrationConfig(tmpdataGridView As DataGridView, sParameterName As String, sTatCode As String, iNumber As Int16, firstRows As Int16, nTotalNo As Int16)
        ' เชื่อมต่อกับ SQL Server
        Dim connectionString As String = ConfigurationManager.ConnectionStrings("MyDatabaseConnectionBatching").ConnectionString
        Dim dbManager As New DatabaseManager(connectionString)

        Await Task.Run(Sub()
                           Dim tmpFirst As Int16 = firstRows
                           For i As Int16 = 0 To nTotalNo - 1
                               tmpFirst = tmpFirst
                               Dim BuckColumnName As String = tmpdataGridView.Rows(firstRows + (i * 12)).Cells(1).Value.ToString.Trim
                               If InStr(BuckColumnName, "Calibration", vbTextCompare) > 0 Then
                                   With AnalogCalibrationConfig_
                                       .motor_name = sParameterName
                                       .c_tat_code = sTatCode
                                       .parameter_no = i + 1
                                       .n_plc_station = nPlcNo
                                       .calibration_name = CheckNullOrEmpty(BuckColumnName.Replace("#", "_").Replace(" ", ""), "N")
                                       .d_address = CheckNullOrEmpty(tmpdataGridView.Rows(firstRows + (i * 12) + 1).Cells(8 + iNumber).Value.ToString.Trim, "D0")
                                       .dt_lastupdate = Now
                                   End With
                                   dbManager.SaveDataCalibrationConfig(AnalogCalibrationConfig_)
                               End If
                           Next
                       End Sub)
    End Sub

    Function CheckNullOrEmpty(value As Object, sDefault As String) As String
        If IsNothing(value) Then
            Return sDefault ' ค่าคือ null
        ElseIf TypeOf value Is String Then
            Dim strValue As String = CType(value, String)
            If String.IsNullOrWhiteSpace(strValue) Then
                Return sDefault ' ค่าว่าง
            Else
                Return strValue ' ค่าไม่เป็น Nothing หรือค่าว่าง
            End If
        End If
        Return sDefault ' หากค่าคือข้อมูลประเภทอื่น
    End Function

    Function ContainsNumber(input As String) As Boolean
        Dim pattern As String = "\d"
        Return Regex.IsMatch(input, pattern)
    End Function

    Function ExtractNumbers(input As String) As String
        Dim result As String = ""
        For Each ch As Char In input
            If Char.IsDigit(ch) Then
                result &= ch
            End If
        Next
        Return result
    End Function

    Private Sub LogError(ex As Exception)
        Try
            ' สร้างโฟลเดอร์สำหรับเก็บ log ตามปี เดือน วัน
            Dim logDirectory As String = Path.Combine("D:\Logs", DateTime.Now.ToString("yyyy"), DateTime.Now.ToString("MM"), DateTime.Now.ToString("dd"))
            Directory.CreateDirectory(logDirectory)

            ' สร้างชื่อไฟล์ log
            Dim logFilePath As String = Path.Combine(logDirectory, "log.txt")

            ' เขียนข้อความ log เข้าไปในไฟล์
            Using writer As New StreamWriter(logFilePath, True) ' True สำหรับ Append
                writer.WriteLine($"{DateTime.Now}: {ex.Message}")
                writer.WriteLine($"{DateTime.Now}: {ex.StackTrace}")
                writer.WriteLine("----------------------------------------------------")
            End Using
        Catch logEx As Exception
            ' หากการเขียน log เกิดข้อผิดพลาด ให้แสดงข้อผิดพลาดที่ Console (หรือสามารถจัดการตามต้องการ)
            Console.WriteLine($"Failed to log error: {logEx.Message}")
        End Try
    End Sub


End Class
