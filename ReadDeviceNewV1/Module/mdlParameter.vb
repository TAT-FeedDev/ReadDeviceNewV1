Module mdlParameter
    Public Structure MotorConfig
        Dim motor_id As Int16
        Dim motor_name As String
        Dim c_tat_code As String
        Dim motor_code As String
        Dim motor_type_id As String
        Dim c_description As String
        Dim c_flow_chat As String
        Dim mservice As String
        Dim mout As String
        Dim mauto As String
        Dim mrun As String
        Dim merr As String
        Dim merr1 As String
        Dim mcoverlock As String
        Dim delay_open As String
        Dim delay_close As String
        Dim delay_step As String
        Dim mpm As Int16
        Dim flag_stop As Int16
        Dim n_plc_station As Int16
        Dim n_dest_motor As Int16
        Dim n_first_motor As Int16
        Dim n_flag_nocheckrun As Int16
        Dim n_motor_link As Int16
        Dim n_bin_no As Int16
        Dim c_run_hour_tar As String
        Dim c_run_hour_act As String
        Dim c_start_count_tar As String
        Dim c_start_count_act As String
        Dim n_time_step As Int16
        Dim c_to_bin As String
        Dim c_from_bin As String
        Dim c_location As String
        Dim c_lock_menu As String
        Dim mrun2 As String
        Dim c_run_hour As String
        Dim c_counter_times As String
        Dim c_target_run_hour As String
        Dim c_target_counter_times As String
        Dim c_pm_alarm As String
        Dim C_input_op As String
        Dim C_input_cl As String
        Dim C_output_op As String
        Dim C_output_cl As String
        Dim c_formula_code As String
        Dim n_login_run As Int16
        Dim n_priority As Int16
        Dim n_error_code As Int32
        Dim c_Error_text As String
        Dim c_error_no As Int32
        Dim c_add_running_hours As String
        Dim delay_on As String
        Dim delay_off As String
        Dim flag_status_run As String
        Dim n_totalton As Double
        Dim int_srt As String
        Dim int_stp As String
        Dim c_alarm_message As String
        Dim c_flag_enable_1 As String
        Dim c_flag_enable_2 As String
        Dim c_code_ref As String
    End Structure

    Public Structure ScadaForAnalog
        Dim motor_name As String
        Dim motor_code As String
        Dim n_number As Int16
        Dim r_address As String
        Dim n_plc_station As Int16
        Dim c_description As String
        Dim n_multiply As Int16
        Dim n_priority As Int16
        Dim c_format As String
        Dim c_record As String
        Dim c_tat_code As String
        Dim c_type As String
        Dim c_short_description As String
    End Structure

    Public Structure BinParameter
        Dim bin_name As String
        Dim bin_code As String
        Dim d_address As String
        Dim d_address_extend As String
        Dim bin_index As Int16
        Dim bin_no As Int16
        Dim scale_no As Int16
        Dim c_location As String
        Dim n_plc_station As Int16
    End Structure

    Public Structure ScadaForAlarm
        Dim alarm_name As String
        Dim zr_address As String
        Dim c_location As String
        Dim c_alarm_name_temp As String
        Dim n_plc_station As Int16
        Dim zr_auto_dosing As String
    End Structure

    Public Structure ScaleParameter
        Dim scale_name As String
        Dim scale_code As String
        Dim d_address As String
        Dim n_plc_station As Int16
        Dim c_location As String
        Dim scale_no As Int16
        Dim time_joging As String
        Dim d_jog_time As String
        Dim d_extend As String
        Dim d_operate As String
    End Structure

    Public Structure AnalogCurrentConfig
        Dim motor_name As String
        Dim c_tat_code As String
        Dim motor_code As String
        Dim d_address As String
        Dim n_plc_station As Int16
        Dim dt_lastupdate As Date
    End Structure

    Public Structure AnalogCalibrationConfig
        Dim motor_name As String
        Dim c_tat_code As String
        Dim parameter_no As Int16
        Dim n_plc_station As Int16
        Dim calibration_name As String
        Dim d_address As String
        Dim dt_lastupdate As Date
    End Structure

End Module
