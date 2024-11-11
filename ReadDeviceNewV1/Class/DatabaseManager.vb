Imports System.Data.SqlClient
Imports System.IO

Public Class DatabaseManager
    Private connectionString As String

    ' Constructor to initialize the connection string
    Public Sub New(connectionString As String)
        Me.connectionString = connectionString
    End Sub
    Public Sub LoadDatabases(connString As String, ComBo As ComboBox)
        Try
            Using conn As New SqlConnection(connString)
                conn.Open()

                ' คำสั่ง SQL เพื่อดึงชื่อฐานข้อมูล
                Dim cmd As New SqlCommand("SELECT name FROM sys.databases", conn)
                Dim reader As SqlDataReader = cmd.ExecuteReader()

                ' ล้างข้อมูลเดิมใน ComboBox
                ComBo.Items.Clear()
                'cmdDatabaseRoute.Items.Clear()

                ' อ่านชื่อฐานข้อมูลและเพิ่มเข้า ComboBox
                While reader.Read()
                    ComBo.Items.Add(reader("name").ToString())
                    'cmdDatabaseRoute.Items.Add(reader("name").ToString())
                End While

                reader.Close()
                'MessageBox.Show("Databases loaded successfully.")
            End Using
        Catch ex As Exception
            'MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub
    ' Overloaded method for saving data using a Dictionary
    Public Sub SaveDataScadaRelay(data As Object)
        Dim strSql As String
        Dim maxRetries As Integer = 3 ' Maximum retry attempts
        Dim currentRetry As Integer = 0
        Dim success As Boolean = False
        While Not success And currentRetry < maxRetries
            Try
                Using connection As New SqlConnection(connectionString)
                    connection.Open()

                    ' ตรวจสอบว่ามีข้อมูลอยู่ในฐานข้อมูลหรือไม่
                    With data
                        Dim checkCmd As New SqlCommand("SELECT COUNT(*) FROM thaisia.motor_config WHERE motor_name = @MotorName AND n_plc_station = @PlcStation", connection)
                        checkCmd.Parameters.AddWithValue("@MotorName", .motor_name)
                        checkCmd.Parameters.AddWithValue("@PlcStation", .n_plc_station)
                        Dim userCount As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())
                        If userCount > 0 Then
                            ' หากพบข้อมูล ให้อัปเดต
                            strSql = "UPDATE thaisia.motor_config SET c_tat_code = @TATCode, "
                            strSql &= " motor_code = @MotorCode, motor_type_id = @MotorType, c_description = @Description, "
                            strSql &= " mservice = @MService, mout = @MOut, mauto = @MAuto, mrun = @MRun, "
                            strSql &= " merr = @MErr, merr1 = @MErr1, mcoverlock = @MCoverlock, delay_open = @DelayOpen, "
                            strSql &= " delay_close = @DelayClose, delay_step = @DelayStep, n_time_step = @TimeStep, "
                            strSql &= " mrun2 = @MRun2, c_run_hour = @RunHour, c_counter_times = @CountTime, c_target_run_hour = @TargetRunHour, "
                            strSql &= " c_target_counter_times = @TargetCountTime, c_pm_alarm = @PmAlarm, n_priority = @Priority, delay_on = @DelayOn, "
                            strSql &= " delay_off = @DelayOff, flag_status_run = @FlagStatusRun, int_srt = @IntSrt, int_stp = @IntStp, "
                            strSql &= " c_alarm_message = @AlarmMessage, c_flag_enable_1 = @FlagEnable1, c_flag_enable_2 = @FlagEnable2, c_code_ref = @CodeRef "
                            strSql &= " WHERE motor_name = @MotorName AND n_plc_station = @PlcStation"
                            Dim updateCmd As New SqlCommand(strSql, connection)
                            updateCmd.Parameters.AddWithValue("@MotorName", .motor_name)
                            updateCmd.Parameters.AddWithValue("@TATCode", .c_tat_code)
                            updateCmd.Parameters.AddWithValue("@MotorCode", .motor_code)
                            updateCmd.Parameters.AddWithValue("@MotorType", .motor_type_id)
                            updateCmd.Parameters.AddWithValue("@Description", .c_description)

                            updateCmd.Parameters.AddWithValue("@MService", .mservice)
                            updateCmd.Parameters.AddWithValue("@MOut", .mout)
                            updateCmd.Parameters.AddWithValue("@MAuto", .mauto)
                            updateCmd.Parameters.AddWithValue("@MRun", .mrun)

                            updateCmd.Parameters.AddWithValue("@MErr", .merr)
                            updateCmd.Parameters.AddWithValue("@MErr1", .merr1)
                            updateCmd.Parameters.AddWithValue("@MCoverlock", .mcoverlock)
                            updateCmd.Parameters.AddWithValue("@DelayOpen", .delay_open)

                            updateCmd.Parameters.AddWithValue("@DelayClose", .delay_close)
                            updateCmd.Parameters.AddWithValue("@DelayStep", .delay_step)
                            updateCmd.Parameters.AddWithValue("@PlcStation", .n_plc_station)
                            updateCmd.Parameters.AddWithValue("@TimeStep", .n_time_step)

                            updateCmd.Parameters.AddWithValue("@MRun2", .mrun2)
                            updateCmd.Parameters.AddWithValue("@RunHour", .c_run_hour)
                            updateCmd.Parameters.AddWithValue("@CountTime", .c_counter_times)
                            updateCmd.Parameters.AddWithValue("@TargetRunHour", .c_target_run_hour)

                            updateCmd.Parameters.AddWithValue("@TargetCountTime", .c_target_counter_times)
                            updateCmd.Parameters.AddWithValue("@PmAlarm", .c_pm_alarm)
                            updateCmd.Parameters.AddWithValue("@Priority", .n_priority)
                            updateCmd.Parameters.AddWithValue("@DelayOn", .delay_on)

                            updateCmd.Parameters.AddWithValue("@DelayOff", .delay_off)
                            updateCmd.Parameters.AddWithValue("@FlagStatusRun", .flag_status_run)
                            updateCmd.Parameters.AddWithValue("@IntSrt", .int_srt)
                            updateCmd.Parameters.AddWithValue("@IntStp", .int_stp)

                            updateCmd.Parameters.AddWithValue("@AlarmMessage", .c_alarm_message)
                            updateCmd.Parameters.AddWithValue("@FlagEnable1", .c_flag_enable_1)
                            updateCmd.Parameters.AddWithValue("@FlagEnable2", .c_flag_enable_2)
                            updateCmd.Parameters.AddWithValue("@CodeRef", .c_code_ref)

                            updateCmd.ExecuteNonQuery()
                            'MessageBox.Show("Data updated successfully!")
                        Else
                            ' หากไม่พบข้อมูล ให้เพิ่ม
                            strSql = "INSERT INTO thaisia.motor_config ( "
                            strSql &= " motor_name, c_tat_code, motor_code, motor_type_id, c_description, "
                            strSql &= " mservice, mout, mauto, mrun, "
                            strSql &= " merr, merr1, mcoverlock, delay_open, "
                            strSql &= " delay_close, delay_step, n_plc_station, n_time_step, "
                            strSql &= " mrun2, c_run_hour, c_counter_times, c_target_run_hour, "
                            strSql &= " c_target_counter_times, c_pm_alarm, n_priority, delay_on, "
                            strSql &= " delay_off, flag_status_run, int_srt, int_stp, "
                            strSql &= " c_alarm_message, c_flag_enable_1, c_flag_enable_2, c_code_ref) "
                            strSql &= " VALUES (@MotorName, @TATCode, @MotorCode, @MotorType, @Description, "
                            strSql &= " @MService, @MOut, @MAuto, @MRun, "
                            strSql &= " @MErr, @MErr1, @MCoverlock, @DelayOpen, "
                            strSql &= " @DelayClose, @DelayStep, @PlcStation, @TimeStep, "
                            strSql &= " @MRun2, @RunHour, @CountTime, @TargetRunHour, "
                            strSql &= " @TargetCountTime, @PmAlarm, @Priority, @DelayOn, "
                            strSql &= " @DelayOff, @FlagStatusRun, @IntSrt, @IntStp, "
                            strSql &= " @AlarmMessage, @FlagEnable1, @FlagEnable2, @CodeRef)"
                            Dim insertCmd As New SqlCommand(strSql, connection)

                            insertCmd.Parameters.AddWithValue("@MotorName", .motor_name)
                            insertCmd.Parameters.AddWithValue("@TATCode", .c_tat_code)
                            insertCmd.Parameters.AddWithValue("@MotorCode", .motor_code)
                            insertCmd.Parameters.AddWithValue("@MotorType", .motor_type_id)
                            insertCmd.Parameters.AddWithValue("@Description", .c_description)

                            insertCmd.Parameters.AddWithValue("@MService", .mservice)
                            insertCmd.Parameters.AddWithValue("@MOut", .mout)
                            insertCmd.Parameters.AddWithValue("@MAuto", .mauto)
                            insertCmd.Parameters.AddWithValue("@MRun", .mrun)

                            insertCmd.Parameters.AddWithValue("@MErr", .merr)
                            insertCmd.Parameters.AddWithValue("@MErr1", .merr1)
                            insertCmd.Parameters.AddWithValue("@MCoverlock", .mcoverlock)
                            insertCmd.Parameters.AddWithValue("@DelayOpen", .delay_open)

                            insertCmd.Parameters.AddWithValue("@DelayClose", .delay_close)
                            insertCmd.Parameters.AddWithValue("@DelayStep", .delay_step)
                            insertCmd.Parameters.AddWithValue("@PlcStation", .n_plc_station)
                            insertCmd.Parameters.AddWithValue("@TimeStep", .n_time_step)

                            insertCmd.Parameters.AddWithValue("@MRun2", .mrun2)
                            insertCmd.Parameters.AddWithValue("@RunHour", .c_run_hour)
                            insertCmd.Parameters.AddWithValue("@CountTime", .c_counter_times)
                            insertCmd.Parameters.AddWithValue("@TargetRunHour", .c_target_run_hour)

                            insertCmd.Parameters.AddWithValue("@TargetCountTime", .c_target_counter_times)
                            insertCmd.Parameters.AddWithValue("@PmAlarm", .c_pm_alarm)
                            insertCmd.Parameters.AddWithValue("@Priority", .n_priority)
                            insertCmd.Parameters.AddWithValue("@DelayOn", .delay_on)

                            insertCmd.Parameters.AddWithValue("@DelayOff", .delay_off)
                            insertCmd.Parameters.AddWithValue("@FlagStatusRun", .flag_status_run)
                            insertCmd.Parameters.AddWithValue("@IntSrt", .int_srt)
                            insertCmd.Parameters.AddWithValue("@IntStp", .int_stp)

                            insertCmd.Parameters.AddWithValue("@AlarmMessage", .c_alarm_message)
                            insertCmd.Parameters.AddWithValue("@FlagEnable1", .c_flag_enable_1)
                            insertCmd.Parameters.AddWithValue("@FlagEnable2", .c_flag_enable_2)
                            insertCmd.Parameters.AddWithValue("@CodeRef", .c_code_ref)
                            insertCmd.ExecuteNonQuery()
                            'MessageBox.Show("Data inserted successfully!")
                        End If
                    End With
                End Using
                success = True ' Success if data is saved without exception
            Catch ex As Exception
                currentRetry += 1
                If currentRetry < maxRetries Then
                    ' Log and wait before retrying
                    LogError(ex)
                    Threading.Thread.Sleep(2000) ' Wait before retrying (2 seconds)
                Else
                    ' Log and throw exception if max retries are reached
                    LogError(ex)
                    Throw New Exception("Failed to save data after multiple attempts. Please check the log for details.", ex)
                End If
            End Try
        End While
    End Sub

    Public Sub SaveDataAnalog(data As Object)
        Dim strSql As String
        Dim maxRetries As Integer = 3 ' Maximum retry attempts
        Dim currentRetry As Integer = 0
        Dim success As Boolean = False
        While Not success And currentRetry < maxRetries
            Try
                Using connection As New SqlConnection(connectionString)
                    connection.Open()

                    ' ตรวจสอบว่ามีข้อมูลอยู่ในฐานข้อมูลหรือไม่
                    With data
                        Dim checkCmd As New SqlCommand("SELECT COUNT(*) FROM thaisia.scada_reg_for_analog WHERE r_address = @RAddress AND n_plc_station = @PlcStation", connection)
                        checkCmd.Parameters.AddWithValue("@RAddress", .r_address)
                        checkCmd.Parameters.AddWithValue("@PlcStation", .n_plc_station)
                        Dim userCount As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())
                        If userCount > 0 Then
                            ' หากพบข้อมูล ให้อัปเดต
                            strSql = "UPDATE thaisia.scada_reg_for_analog SET motor_name = @MotorName, "
                            strSql &= " motor_code = @MotorCode, "
                            strSql &= " n_number = @NNumber, "
                            strSql &= " c_description = @Description, "
                            strSql &= " n_multiply = @NMultiply, "
                            strSql &= " n_priority = @Priority, "
                            strSql &= " c_format = @CFormate, "
                            strSql &= " c_record = @Record, "
                            strSql &= " c_tat_code = @TATCode, "
                            strSql &= " c_type = @Type, "
                            strSql &= " c_short_description = @ShortDescription "
                            strSql &= " WHERE r_address = @RAddress AND n_plc_station = @PlcStation"
                            Dim updateCmd As New SqlCommand(strSql, connection)
                            updateCmd.Parameters.AddWithValue("@MotorName", .motor_name)
                            updateCmd.Parameters.AddWithValue("@MotorCode", .motor_code)
                            updateCmd.Parameters.AddWithValue("@NNumber", .n_number)
                            updateCmd.Parameters.AddWithValue("@RAddress", .r_address)
                            updateCmd.Parameters.AddWithValue("@PlcStation", .n_plc_station)
                            updateCmd.Parameters.AddWithValue("@Description", .c_description)
                            updateCmd.Parameters.AddWithValue("@NMultiply", .n_multiply)
                            updateCmd.Parameters.AddWithValue("@Priority", .n_priority)
                            updateCmd.Parameters.AddWithValue("@CFormate", .c_format)
                            updateCmd.Parameters.AddWithValue("@Record", .c_record)
                            updateCmd.Parameters.AddWithValue("@TATCode", .c_tat_code)
                            updateCmd.Parameters.AddWithValue("@Type", .c_type)
                            updateCmd.Parameters.AddWithValue("@ShortDescription", .c_short_description)
                            updateCmd.ExecuteNonQuery()
                            'MessageBox.Show("Data updated successfully!")
                        Else
                            ' หากไม่พบข้อมูล ให้เพิ่ม
                            strSql = "INSERT INTO thaisia.scada_reg_for_analog ("
                            strSql &= " motor_name, "
                            strSql &= " motor_code, "
                            strSql &= " n_number, "
                            strSql &= " r_address, "
                            strSql &= " n_plc_station, "
                            strSql &= " c_description, "
                            strSql &= " n_multiply, "
                            strSql &= " n_priority, "
                            strSql &= " c_format, "
                            strSql &= " c_record, "
                            strSql &= " c_tat_code, "
                            strSql &= " c_type, "
                            strSql &= " c_short_description) "
                            strSql &= " VALUES ("
                            strSql &= " @MotorName, "
                            strSql &= " @MotorCode, "
                            strSql &= " @NNumber, "
                            strSql &= " @RAddress, "
                            strSql &= " @PlcStation, "
                            strSql &= " @Description, "
                            strSql &= " @NMultiply, "
                            strSql &= " @Priority, "
                            strSql &= " @CFormate, "
                            strSql &= " @Record, "
                            strSql &= " @TATCode, "
                            strSql &= " @Type, "
                            strSql &= " @ShortDescription)"
                            Dim insertCmd As New SqlCommand(strSql, connection)
                            insertCmd.Parameters.AddWithValue("@MotorName", .motor_name)
                            insertCmd.Parameters.AddWithValue("@MotorCode", .motor_code)
                            insertCmd.Parameters.AddWithValue("@NNumber", .n_number)
                            insertCmd.Parameters.AddWithValue("@RAddress", .r_address)
                            insertCmd.Parameters.AddWithValue("@PlcStation", .n_plc_station)
                            insertCmd.Parameters.AddWithValue("@Description", .c_description)
                            insertCmd.Parameters.AddWithValue("@NMultiply", .n_multiply)
                            insertCmd.Parameters.AddWithValue("@Priority", .n_priority)
                            insertCmd.Parameters.AddWithValue("@CFormate", .c_format)
                            insertCmd.Parameters.AddWithValue("@Record", .c_record)
                            insertCmd.Parameters.AddWithValue("@TATCode", .c_tat_code)
                            insertCmd.Parameters.AddWithValue("@Type", .c_type)
                            insertCmd.Parameters.AddWithValue("@ShortDescription", .c_short_description)
                            insertCmd.ExecuteNonQuery()
                            'MessageBox.Show("Data inserted successfully!")
                        End If
                    End With
                End Using
                success = True ' Success if data is saved without exception
            Catch ex As Exception
                currentRetry += 1
                If currentRetry < maxRetries Then
                    ' Log and wait before retrying
                    LogError(ex)
                    Threading.Thread.Sleep(2000) ' Wait before retrying (2 seconds)
                Else
                    ' Log and throw exception if max retries are reached
                    LogError(ex)
                    Throw New Exception("Failed to save data after multiple attempts. Please check the log for details.", ex)
                End If
            End Try
        End While
    End Sub

    Public Sub SaveDataBinParameter(data As Object)
        Dim strSql As String
        Dim maxRetries As Integer = 3 ' Maximum retry attempts
        Dim currentRetry As Integer = 0
        Dim success As Boolean = False
        While Not success And currentRetry < maxRetries
            Try
                Using connection As New SqlConnection(connectionString)
                    connection.Open()

                    ' ตรวจสอบว่ามีข้อมูลอยู่ในฐานข้อมูลหรือไม่
                    With data
                        Dim checkCmd As New SqlCommand("SELECT COUNT(*) FROM thaisia.bin_parameter_ WHERE bin_no = @BinNo AND n_plc_station = @PlcStation", connection)
                        checkCmd.Parameters.AddWithValue("@BinNo", .bin_no)
                        checkCmd.Parameters.AddWithValue("@PlcStation", .n_plc_station)
                        Dim userCount As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())
                        If userCount > 0 Then
                            ' หากพบข้อมูล ให้อัปเดต
                            strSql = "UPDATE thaisia.bin_parameter_ SET bin_name = @BinName, "
                            strSql &= " bin_code = @BinCode, "
                            strSql &= " d_address = @DAddress, "
                            strSql &= " d_address_extend = @DExtenAddress, "
                            strSql &= " bin_index = @BinIndex, "
                            strSql &= " scale_no = @ScaleNo "
                            strSql &= " WHERE bin_no = @BinNo AND n_plc_station = @PlcStation"
                            Dim updateCmd As New SqlCommand(strSql, connection)
                            updateCmd.Parameters.AddWithValue("@BinName", .bin_name)
                            updateCmd.Parameters.AddWithValue("@BinCode", .bin_code)
                            updateCmd.Parameters.AddWithValue("@DAddress", .d_address)
                            updateCmd.Parameters.AddWithValue("@DExtenAddress", .d_address_extend)
                            updateCmd.Parameters.AddWithValue("@BinIndex", .bin_index)
                            updateCmd.Parameters.AddWithValue("@BinNo", .bin_no)
                            updateCmd.Parameters.AddWithValue("@ScaleNo", .scale_no)
                            'updateCmd.Parameters.AddWithValue("@Location", .c_location)
                            updateCmd.Parameters.AddWithValue("@PlcStation", .n_plc_station)
                            updateCmd.ExecuteNonQuery()
                            'MessageBox.Show("Data updated successfully!")
                        Else
                            ' หากไม่พบข้อมูล ให้เพิ่ม
                            strSql = "INSERT INTO thaisia.bin_parameter_ ("
                            strSql &= " bin_name, "
                            strSql &= " bin_code, "
                            strSql &= " d_address, "
                            strSql &= " d_address_extend, "
                            strSql &= " bin_index, "
                            strSql &= " bin_no, "
                            strSql &= " scale_no, "
                            strSql &= " n_plc_station) "
                            strSql &= " VALUES ("
                            strSql &= " @BinName, "
                            strSql &= " @BinCode, "
                            strSql &= " @DAddress, "
                            strSql &= " @DExtenAddress, "
                            strSql &= " @BinIndex, "
                            strSql &= " @BinNo, "
                            strSql &= " @ScaleNo, "
                            strSql &= " @PlcStation)"
                            Dim insertCmd As New SqlCommand(strSql, connection)
                            insertCmd.Parameters.AddWithValue("@BinName", .bin_name)
                            insertCmd.Parameters.AddWithValue("@BinCode", .bin_code)
                            insertCmd.Parameters.AddWithValue("@DAddress", .d_address)
                            insertCmd.Parameters.AddWithValue("@DExtenAddress", .d_address_extend)
                            insertCmd.Parameters.AddWithValue("@BinIndex", .bin_index)
                            insertCmd.Parameters.AddWithValue("@BinNo", .bin_no)
                            insertCmd.Parameters.AddWithValue("@ScaleNo", .scale_no)
                            'insertCmd.Parameters.AddWithValue("@Location", .c_location)
                            insertCmd.Parameters.AddWithValue("@PlcStation", .n_plc_station)
                            insertCmd.ExecuteNonQuery()
                            'MessageBox.Show("Data inserted successfully!")
                        End If
                    End With
                End Using
                success = True ' Success if data is saved without exception
            Catch ex As Exception
                currentRetry += 1
                If currentRetry < maxRetries Then
                    ' Log and wait before retrying
                    LogError(ex)
                    Threading.Thread.Sleep(2000) ' Wait before retrying (2 seconds)
                Else
                    ' Log and throw exception if max retries are reached
                    LogError(ex)
                    Throw New Exception("Failed to save data after multiple attempts. Please check the log for details.", ex)
                End If
            End Try
        End While
    End Sub

    Public Sub SaveDataForAlarm(data As Object)
        Dim strSql As String
        Dim maxRetries As Integer = 3 ' Maximum retry attempts
        Dim currentRetry As Integer = 0
        Dim success As Boolean = False
        While Not success And currentRetry < maxRetries
            Try
                Using connection As New SqlConnection(connectionString)
                    connection.Open()

                    ' ตรวจสอบว่ามีข้อมูลอยู่ในฐานข้อมูลหรือไม่
                    With data
                        Dim checkCmd As New SqlCommand("SELECT COUNT(*) FROM thaisia.scada_reg_for_alarm WHERE alarm_name = @AlarmName AND zr_address = @RAddress AND n_plc_station = @PlcStation", connection)
                        checkCmd.Parameters.AddWithValue("@AlarmName", .alarm_name)
                        checkCmd.Parameters.AddWithValue("@RAddress", .zr_address)
                        checkCmd.Parameters.AddWithValue("@PlcStation", .n_plc_station)
                        Dim userCount As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())
                        If userCount > 0 Then
                            ' หากพบข้อมูล ให้อัปเดต
                            strSql = "UPDATE thaisia.scada_reg_for_alarm SET c_alarm_name_temp = @AlarmNameTemp "
                            strSql &= " WHERE alarm_name = @AlarmName AND zr_address = @RAddress AND n_plc_station = @PlcStation"
                            Dim updateCmd As New SqlCommand(strSql, connection)
                            updateCmd.Parameters.AddWithValue("@AlarmName", .alarm_name)
                            updateCmd.Parameters.AddWithValue("@RAddress", .zr_address)
                            updateCmd.Parameters.AddWithValue("@AlarmNameTemp", .c_alarm_name_temp)
                            updateCmd.Parameters.AddWithValue("@PlcStation", .n_plc_station)
                            updateCmd.ExecuteNonQuery()
                            'MessageBox.Show("Data updated successfully!")
                        Else
                            ' หากไม่พบข้อมูล ให้เพิ่ม
                            strSql = "INSERT INTO thaisia.scada_reg_for_alarm ("
                            strSql &= " alarm_name, "
                            strSql &= " zr_address, "
                            strSql &= " c_location, "
                            strSql &= " c_alarm_name_temp, "
                            strSql &= " n_plc_station) "
                            strSql &= " VALUES ("
                            strSql &= " @AlarmName, "
                            strSql &= " @RAddress, "
                            strSql &= " @Location, "
                            strSql &= " @AlarmNameTemp, "
                            strSql &= " @PlcStation)"
                            Dim insertCmd As New SqlCommand(strSql, connection)
                            insertCmd.Parameters.AddWithValue("@AlarmName", .alarm_name)
                            insertCmd.Parameters.AddWithValue("@RAddress", .zr_address)
                            insertCmd.Parameters.AddWithValue("@Location", .c_location)
                            insertCmd.Parameters.AddWithValue("@AlarmNameTemp", .c_alarm_name_temp)
                            insertCmd.Parameters.AddWithValue("@PlcStation", .n_plc_station)
                            insertCmd.ExecuteNonQuery()
                            'MessageBox.Show("Data inserted successfully!")
                        End If
                    End With
                End Using
                success = True ' Success if data is saved without exception
            Catch ex As Exception
                currentRetry += 1
                If currentRetry < maxRetries Then
                    ' Log and wait before retrying
                    LogError(ex)
                    Threading.Thread.Sleep(2000) ' Wait before retrying (2 seconds)
                Else
                    ' Log and throw exception if max retries are reached
                    LogError(ex)
                    Throw New Exception("Failed to save data after multiple attempts. Please check the log for details.", ex)
                End If
            End Try
        End While
    End Sub

    Public Sub SaveDataScaleParameter(data As Object)
        Dim strSql As String
        Dim maxRetries As Integer = 3 ' Maximum retry attempts
        Dim currentRetry As Integer = 0
        Dim success As Boolean = False
        While Not success And currentRetry < maxRetries
            Try
                Using connection As New SqlConnection(connectionString)
                    connection.Open()

                    ' ตรวจสอบว่ามีข้อมูลอยู่ในฐานข้อมูลหรือไม่
                    With data
                        Dim checkCmd As New SqlCommand("SELECT COUNT(*) FROM thaisia.scale_parameter_ WHERE d_address = @DAddress AND n_plc_station = @PlcStation", connection)
                        checkCmd.Parameters.AddWithValue("@DAddress", .d_address)
                        checkCmd.Parameters.AddWithValue("@PlcStation", .n_plc_station)
                        Dim userCount As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())
                        If userCount > 0 Then
                            ' หากพบข้อมูล ให้อัปเดต
                            strSql = "UPDATE thaisia.scale_parameter_ SET scale_name = @ScaleName, "
                            strSql &= " scale_code = @ScaleCode, "
                            strSql &= " scale_no = @ScaleNo, "
                            strSql &= " time_joging = @TimeJoging, "
                            strSql &= " d_jog_time = @DJogTime, "
                            strSql &= " d_extend = @DExtend, "
                            strSql &= " d_operate = @DOperate "
                            strSql &= " WHERE d_address = @DAddress AND n_plc_station = @PlcStation"
                            Dim updateCmd As New SqlCommand(strSql, connection)
                            updateCmd.Parameters.AddWithValue("@ScaleName", .scale_name)
                            updateCmd.Parameters.AddWithValue("@ScaleCode", .scale_code)
                            updateCmd.Parameters.AddWithValue("@DAddress", .d_address)
                            updateCmd.Parameters.AddWithValue("@PlcStation", .n_plc_station)
                            'updateCmd.Parameters.AddWithValue("@Location", .c_location)
                            updateCmd.Parameters.AddWithValue("@ScaleNo", .scale_no)
                            updateCmd.Parameters.AddWithValue("@TimeJoging", .time_joging)
                            updateCmd.Parameters.AddWithValue("@DJogTime", .d_jog_time)
                            updateCmd.Parameters.AddWithValue("@DExtend", .d_extend)
                            updateCmd.Parameters.AddWithValue("@DOperate", .d_operate)
                            updateCmd.ExecuteNonQuery()
                            'MessageBox.Show("Data updated successfully!")
                        Else
                            ' หากไม่พบข้อมูล ให้เพิ่ม
                            strSql = "INSERT INTO thaisia.scale_parameter_ ("
                            strSql &= " scale_name, "
                            strSql &= " scale_code, "
                            strSql &= " d_address, "
                            strSql &= " n_plc_station, "
                            strSql &= " scale_no, "
                            strSql &= " time_joging, "
                            strSql &= " d_jog_time, "
                            strSql &= " d_extend, "
                            strSql &= " d_operate) "
                            strSql &= " VALUES ("
                            strSql &= " @ScaleName, "
                            strSql &= " @ScaleCode, "
                            strSql &= " @DAddress, "
                            strSql &= " @PlcStation, "
                            strSql &= " @ScaleNo, "
                            strSql &= " @TimeJoging, "
                            strSql &= " @DJogTime, "
                            strSql &= " @DExtend, "
                            strSql &= " @DOperate)"
                            Dim insertCmd As New SqlCommand(strSql, connection)
                            insertCmd.Parameters.AddWithValue("@ScaleName", .scale_name)
                            insertCmd.Parameters.AddWithValue("@ScaleCode", .scale_code)
                            insertCmd.Parameters.AddWithValue("@DAddress", .d_address)
                            insertCmd.Parameters.AddWithValue("@PlcStation", .n_plc_station)
                            'insertCmd.Parameters.AddWithValue("@Location", .c_location)
                            insertCmd.Parameters.AddWithValue("@ScaleNo", .scale_no)
                            insertCmd.Parameters.AddWithValue("@TimeJoging", .time_joging)
                            insertCmd.Parameters.AddWithValue("@DJogTime", .d_jog_time)
                            insertCmd.Parameters.AddWithValue("@DExtend", .d_extend)
                            insertCmd.Parameters.AddWithValue("@DOperate", .d_operate)
                            insertCmd.ExecuteNonQuery()
                            'MessageBox.Show("Data inserted successfully!")
                        End If
                    End With
                End Using
                success = True ' Success if data is saved without exception
            Catch ex As Exception
                currentRetry += 1
                If currentRetry < maxRetries Then
                    ' Log and wait before retrying
                    LogError(ex)
                    Threading.Thread.Sleep(2000) ' Wait before retrying (2 seconds)
                Else
                    ' Log and throw exception if max retries are reached
                    LogError(ex)
                    Throw New Exception("Failed to save data after multiple attempts. Please check the log for details.", ex)
                End If
            End Try
        End While
    End Sub

    Public Sub SaveDataAnalogConfig(data As Object)
        Dim strSql As String
        Dim maxRetries As Integer = 3 ' Maximum retry attempts
        Dim currentRetry As Integer = 0
        Dim success As Boolean = False
        While Not success And currentRetry < maxRetries
            Try
                Using connection As New SqlConnection(connectionString)
                    connection.Open()

                    ' ตรวจสอบว่ามีข้อมูลอยู่ในฐานข้อมูลหรือไม่
                    With data
                        Dim checkCmd As New SqlCommand("SELECT COUNT(*) FROM thaisia.analog_current_config WHERE motor_name = @MotorName AND n_plc_station = @PlcStation", connection)
                        checkCmd.Parameters.AddWithValue("@MotorName", .motor_name)
                        checkCmd.Parameters.AddWithValue("@PlcStation", .n_plc_station)
                        Dim userCount As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())
                        If userCount > 0 Then
                            ' หากพบข้อมูล ให้อัปเดต
                            strSql = "UPDATE thaisia.analog_current_config SET motor_code = @MotorCode, "
                            strSql &= " c_tat_code = @TATCode, "
                            strSql &= " d_address = @DAddress, "
                            strSql &= " dt_lastupdate = @LastUpdate "
                            strSql &= " WHERE motor_name = @MotorName AND n_plc_station = @PlcStation"
                            Dim updateCmd As New SqlCommand(strSql, connection)
                            updateCmd.Parameters.AddWithValue("@MotorName", .motor_name)
                            updateCmd.Parameters.AddWithValue("@TATCode", .c_tat_code)
                            updateCmd.Parameters.AddWithValue("@MotorCode", .motor_code)
                            updateCmd.Parameters.AddWithValue("@DAddress", .d_address)
                            updateCmd.Parameters.AddWithValue("@PlcStation", .n_plc_station)
                            updateCmd.Parameters.AddWithValue("@LastUpdate", .dt_lastupdate)
                            updateCmd.ExecuteNonQuery()
                            'MessageBox.Show("Data updated successfully!")
                        Else
                            ' หากไม่พบข้อมูล ให้เพิ่ม
                            strSql = "INSERT INTO thaisia.analog_current_config ("
                            strSql &= " motor_name, "
                            strSql &= " c_tat_code, "
                            strSql &= " motor_code, "
                            strSql &= " d_address, "
                            strSql &= " n_plc_station, "
                            strSql &= " dt_lastupdate) "
                            strSql &= " VALUES ("
                            strSql &= " @MotorName, "
                            strSql &= " @TATCode, "
                            strSql &= " @MotorCode, "
                            strSql &= " @DAddress, "
                            strSql &= " @PlcStation,"
                            strSql &= " @LastUpdate)"
                            Dim insertCmd As New SqlCommand(strSql, connection)
                            insertCmd.Parameters.AddWithValue("@MotorName", .motor_name)
                            insertCmd.Parameters.AddWithValue("@TATCode", .c_tat_code)
                            insertCmd.Parameters.AddWithValue("@MotorCode", .motor_code)
                            insertCmd.Parameters.AddWithValue("@DAddress", .d_address)
                            insertCmd.Parameters.AddWithValue("@PlcStation", .n_plc_station)
                            insertCmd.Parameters.AddWithValue("@LastUpdate", .dt_lastupdate)
                            insertCmd.ExecuteNonQuery()
                            'MessageBox.Show("Data inserted successfully!")
                        End If
                    End With
                End Using
                success = True ' Success if data is saved without exception
            Catch ex As Exception
                currentRetry += 1
                If currentRetry < maxRetries Then
                    ' Log and wait before retrying
                    LogError(ex)
                    Threading.Thread.Sleep(2000) ' Wait before retrying (2 seconds)
                Else
                    ' Log and throw exception if max retries are reached
                    LogError(ex)
                    Throw New Exception("Failed to save data after multiple attempts. Please check the log for details.", ex)
                End If
            End Try
        End While
    End Sub

    Public Sub SaveDataCalibrationConfig(data As Object)
        Dim strSql As String
        Dim maxRetries As Integer = 3 ' Maximum retry attempts
        Dim currentRetry As Integer = 0
        Dim success As Boolean = False
        While Not success And currentRetry < maxRetries
            Try
                Using connection As New SqlConnection(connectionString)
                    connection.Open()

                    ' ตรวจสอบว่ามีข้อมูลอยู่ในฐานข้อมูลหรือไม่
                    With data
                        Dim checkCmd As New SqlCommand("SELECT COUNT(*) FROM thaisia.analog_calibration_config WHERE motor_name = @MotorName AND parameter_no = @ParameterNo AND n_plc_station = @PlcStation", connection)
                        checkCmd.Parameters.AddWithValue("@MotorName", .motor_name)
                        checkCmd.Parameters.AddWithValue("@ParameterNo", .parameter_no)
                        checkCmd.Parameters.AddWithValue("@PlcStation", .n_plc_station)
                        Dim userCount As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())
                        If userCount > 0 Then
                            ' หากพบข้อมูล ให้อัปเดต
                            strSql = "UPDATE thaisia.analog_calibration_config SET calibration_name = @CalibrationName, "
                            strSql &= " c_tat_code = @TATCode, "
                            strSql &= " d_address = @DAddress, "
                            strSql &= " dt_lastupdate = @LastUpdate "
                            strSql &= " WHERE motor_name = @MotorName AND parameter_no = @ParameterNo AND n_plc_station = @PlcStation"
                            Dim updateCmd As New SqlCommand(strSql, connection)
                            updateCmd.Parameters.AddWithValue("@MotorName", .motor_name)
                            updateCmd.Parameters.AddWithValue("@TATCode", .c_tat_code)
                            updateCmd.Parameters.AddWithValue("@ParameterNo", .parameter_no)
                            updateCmd.Parameters.AddWithValue("@CalibrationName", .calibration_name)
                            updateCmd.Parameters.AddWithValue("@DAddress", .d_address)
                            updateCmd.Parameters.AddWithValue("@PlcStation", .n_plc_station)
                            updateCmd.Parameters.AddWithValue("@LastUpdate", .dt_lastupdate)
                            updateCmd.ExecuteNonQuery()
                            'MessageBox.Show("Data updated successfully!")
                        Else
                            ' หากไม่พบข้อมูล ให้เพิ่ม
                            strSql = "INSERT INTO thaisia.analog_calibration_config ("
                            strSql &= " motor_name, "
                            strSql &= " c_tat_code, "
                            strSql &= " parameter_no, "
                            strSql &= " calibration_name, "
                            strSql &= " d_address, "
                            strSql &= " n_plc_station, "
                            strSql &= " dt_lastupdate) "
                            strSql &= " VALUES ("
                            strSql &= " @MotorName, "
                            strSql &= " @TATCode, "
                            strSql &= " @ParameterNo, "
                            strSql &= " @CalibrationName, "
                            strSql &= " @DAddress, "
                            strSql &= " @PlcStation,"
                            strSql &= " @LastUpdate)"
                            Dim insertCmd As New SqlCommand(strSql, connection)
                            insertCmd.Parameters.AddWithValue("@MotorName", .motor_name)
                            insertCmd.Parameters.AddWithValue("@TATCode", .c_tat_code)
                            insertCmd.Parameters.AddWithValue("@ParameterNo", .parameter_no)
                            insertCmd.Parameters.AddWithValue("@CalibrationName", .calibration_name)
                            insertCmd.Parameters.AddWithValue("@DAddress", .d_address)
                            insertCmd.Parameters.AddWithValue("@PlcStation", .n_plc_station)
                            insertCmd.Parameters.AddWithValue("@LastUpdate", .dt_lastupdate)
                            insertCmd.ExecuteNonQuery()
                            'MessageBox.Show("Data inserted successfully!")
                        End If
                    End With
                End Using
                success = True ' Success if data is saved without exception
            Catch ex As Exception
                currentRetry += 1
                If currentRetry < maxRetries Then
                    ' Log and wait before retrying
                    LogError(ex)
                    Threading.Thread.Sleep(2000) ' Wait before retrying (2 seconds)
                Else
                    ' Log and throw exception if max retries are reached
                    LogError(ex)
                    Throw New Exception("Failed to save data after multiple attempts. Please check the log for details.", ex)
                End If
            End Try
        End While
    End Sub

    ' Function to log errors
    Public Sub LogError(ex As Exception)
        Try
            ' กำหนดที่อยู่โฟลเดอร์สำหรับเก็บ log
            Dim logDirectory As String = Path.Combine("D:\ReadDevice\Logs", DateTime.Now.ToString("yyyy"), DateTime.Now.ToString("MM"), DateTime.Now.ToString("dd"))
            Directory.CreateDirectory(logDirectory)

            ' กำหนดชื่อไฟล์ log
            Dim logFilePath As String = Path.Combine(logDirectory, "log.txt")

            ' เขียนข้อผิดพลาดลงไฟล์ log
            Using writer As New StreamWriter(logFilePath, True) ' Append to file
                writer.WriteLine($"{DateTime.Now}: {ex.Message}")
                writer.WriteLine($"{DateTime.Now}: {ex.StackTrace}")
                writer.WriteLine("----------------------------------------------------")
            End Using
        Catch logEx As Exception
            ' ในกรณีที่การเขียน log มีข้อผิดพลาด ให้แสดงข้อความข้อผิดพลาดที่ console
            Console.WriteLine($"Failed to log error: {logEx.Message}")
        End Try
    End Sub

End Class
