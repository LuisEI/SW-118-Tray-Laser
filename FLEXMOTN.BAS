Attribute VB_Name = "FLEXMOTN"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FLEXMOTN.BAS File
' Include File for Visual Basica
' Copyright © National Instruments 1996-1998.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'A few notes about the structure of commands
'Commands have many arguments in common, for instance:
'
' boardID: While calling the FlexMotion Dynamic link library (FLEX16.DLL/FlexMotion32.DLL),
' this will be a number from 1 to 32. In 16 bit environment this will
' tell the FlexMotion DLL which record in the <nimotion.cfg> file to search, for
' the address of the board, as well as any the board type.
' In 32 bit environment this information is obtained from the registry (stored during
' Hardware installation).
'
' Resource/axis/encoder/vectorSpace: This tells the board exactly which resource (axis, encoder...) you are
' trying to control or read. The ResourceIDs are listed in the manual.
'
' inputVector: The Input Vector argument tells the board where you want the
' input data values for the command to come from.  For instance, the load_target_pos()
' function could have an inpvect of 0xFF (if data is being sent from the host computer)
' or 0x03 (if data is to be read from general purpose variable 3).
' See the FlexMotion user manual for more information on input
' vectors and general purpose variables.
'
' returnVector: The return vector tells the board where you would like the results of
' the operation to be sent, (ie. returned to the host or sent to a general
' purpose variable. This is only used on read commands and the flex_wait()
' function.
'
' All other arguments are specific to each command, and should be referred
' to in the function reference manual.
'
' The read commands are given in two different styles: with or without
' the _rtn suffix. The functions without the _rtn suffix do not return
' any value to the application that calls them, in fact they don't return data
' to the host application at all unless given the retvect of 0xFF. The value
' must be read explicitly by a separate execution of the flex_read_rdb command.
' The read commands that have the _rtn suffix are expecting a valid pointer
' to where you would like the data returned. These routines do in fact wait
' for the board to respond with data before returning to the calling
' application.
'
' Errors: There are two forms of errors returned:
'  1) Non Modal:   These are returned by the DLL in the return value of
'                  each function. These are usually serious errors like
'                  communication errors.
'  2) Modal:   These errors are returned by the board after the it gets the command.
'              These errors can be read using the flex_read_err_msg_rtn command. It
'              returns the command ID which caused the error, the Resource  and the error
'              code.
'  For Error Codes and Command IDs refer to the online Function Reference
'KEY
'% = short
'& = long
'# = double
'! = single
'$ = string
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Data structures used by FlexMotion Commands
' The registry information structure for objects stored on the board
Type registryRecord
    device As Integer
    type As Integer
    pStart As Long
    size As Long
End Type

'The data structure for loading PID values to an axis (servo axis)
Type PidVals
    kp As Integer
    ki As Integer
    ilim As Integer
    kd As Integer
    td As Integer
    kv As Integer
    aff As Integer
    vff As Integer
End Type

Declare Function flex_initialize_controller Lib "FlexMotion32.dll" (ByVal boardID%, ByVal settingsName As String) As Long
Declare Function flex_config_axis Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal primaryFeedback%, ByVal secondaryFeedback%, ByVal primaryOutput%, ByVal secondaryOutput%) As Long
Declare Function flex_config_mc_criteria Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal criteria%, ByVal deadband%, ByVal delay%, ByVal minPulse%) As Long
Declare Function flex_configure_stepper_output Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrStepperOutput%, ByVal outputMode%, ByVal polarity%, ByVal driveMode%) As Long
Declare Function flex_config_vect_spc Lib "FlexMotion32.dll" (ByVal boardID%, ByVal vectorSpace%, ByVal xAxis%, ByVal yAxis%, ByVal zAxis%) As Long
Declare Function flex_enable_axis Lib "FlexMotion32.dll" (ByVal boardID%, ByVal reserved%, ByVal PIDRate%, ByVal axismap%) As Long
Declare Function flex_load_advanced_control_parameter Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal parameterType%, ByVal value&, ByVal inputVector%) As Long
Declare Function flex_load_counts_steps_rev Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal unitType%, ByVal countsOrSteps&) As Long
Declare Function flex_load_pid_parameters Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, PIDValues As PidVals, ByVal inputVector%) As Long
Declare Function flex_load_single_pid_parameter Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal parameterType%, ByVal PIDValue%, ByVal inputVector%) As Long
Declare Function flex_set_stepper_loop_mode Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal loopMode%) As Long
Declare Function flex_load_velocity_filter_parameter Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal filterDistance%, ByVal filterTime%, ByVal inputVector%) As Long
Declare Function flex_load_commutation_parameter Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal attr%, ByVal value#) As Long
Declare Function flex_load_axis_configuration_parameter Lib "FlexMotion32.dll" (ByVal boardID&, ByVal axisID&, ByVal attr%, data%) As Long
Declare Function flex_load_acceleration Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal accelerationType%, ByVal acceleration&, ByVal inputVector%) As Long
Declare Function flex_load_follow_err Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal followingError%, ByVal inputVector%) As Long
Declare Function flex_load_rpm Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal RPM#, ByVal inputVector%) As Long
Declare Function flex_load_rpsps Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal accelerationType%, ByVal rpsps#, ByVal inputVector%) As Long
Declare Function flex_load_target_pos Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal targetPosition&, ByVal inputVector%) As Long
Declare Function flex_load_velocity Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal velocity&, ByVal inputVector%) As Long
Declare Function flex_load_vs_pos Lib "FlexMotion32.dll" (ByVal boardID%, ByVal vectorSpace%, ByVal xPosition&, ByVal yPosition&, ByVal zPosition&, ByVal inputVector%) As Long
Declare Function flex_load_move_constraint Lib "FlexMotion32.dll" (ByVal boardID&, ByVal axisID&, ByVal attr%, value%) As Long
Declare Function flex_read_axis_status Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal returnVector%) As Long
Declare Function flex_read_axis_status_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, axisStatus%) As Long
Declare Function flex_read_blend_status Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal returnVector%) As Long
Declare Function flex_read_blend_status_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, blendStatus%) As Long
Declare Function flex_read_follow_err Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal returnVector%) As Long
Declare Function flex_read_follow_err_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, followingError%) As Long
Declare Function flex_read_mcs_rtn Lib "FlexMotion32.dll" (ByVal boardID%, moveCompleteStatus%) As Long
Declare Function flex_read_pos Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal returnVector%) As Long
Declare Function flex_read_pos_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, position&) As Long
Declare Function flex_read_rpm Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal returnVector%) As Long
Declare Function flex_read_rpm_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, RPM#) As Long
Declare Function flex_read_trajectory_status Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal statusType%, ByVal returnVector%) As Long
Declare Function flex_read_trajectory_status_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal statusType%, status%) As Long
Declare Function flex_read_velocity Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal returnVector%) As Long
Declare Function flex_read_velocity_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, velocity&) As Long
Declare Function flex_read_vs_pos Lib "FlexMotion32.dll" (ByVal boardID%, ByVal vectorSpace%, ByVal returnVector%) As Long
Declare Function flex_read_vs_pos_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal vectorSpace%, xPosition&, yPosition&, zPosition&) As Long
Declare Function flex_reset_pos Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal position1&, ByVal position2&, ByVal inputVector%) As Long
Declare Function flex_set_op_mode Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal operationMode%) As Long
Declare Function flex_check_move_complete_status Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal axisOrVSMap%, moveComplete%) As Long
Declare Function flex_check_blend_complete_status Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal axisOrVSMap%, blendComplete%) As Long
Declare Function flex_wait_for_move_complete Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal axisOrVSMap%, ByVal timeout&, ByVal pollInterval&, moveComplete%) As Long
Declare Function flex_wait_for_blend_complete Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal axisOrVSMap%, ByVal timeout&, ByVal pollInterval&, blendComplete%) As Long
Declare Function flex_load_circular_arc Lib "FlexMotion32.dll" (ByVal boardID%, ByVal vectorSpace%, ByVal radius&, ByVal startangle#, ByVal travelangle#, ByVal inputVector%) As Long
Declare Function flex_load_helical_arc Lib "FlexMotion32.dll" (ByVal boardID%, ByVal vectorSpace%, ByVal radius&, ByVal startangle#, ByVal travelangle#, ByVal lineartravel&, ByVal inputVector%) As Long
Declare Function flex_load_spherical_arc Lib "FlexMotion32.dll" (ByVal boardID%, ByVal vectorSpace%, ByVal radius&, ByVal planepitch#, ByVal planeyaw#, ByVal startangle#, ByVal travelangle#, ByVal inputVector%) As Long
Declare Function flex_config_gear_master Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal masterAxisOrEncoderOrADC%) As Long
Declare Function flex_enable_gearing Lib "FlexMotion32.dll" (ByVal boardID%, ByVal gearMap%) As Long
Declare Function flex_enable_gearing_single_axis Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal enable%) As Long
Declare Function flex_load_gear_ratio Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal absoluteOrRelative%, ByVal ratioNumerator%, ByVal ratioDenominator%, ByVal inputVector%) As Long
Declare Function flex_configure_camming_master Lib "FlexMotion32.dll" (ByVal boardID&, ByVal axisID&, ByVal masterResource&, ByVal masterCycle#) As Long
Declare Function flex_load_camming_parameter Lib "FlexMotion32.dll" (ByVal boardID&, ByVal axisID&, ByVal attr%, data%) As Long
Declare Function flex_enable_camming Lib "FlexMotion32.dll" (ByVal boardID&, ByVal arraySize&, dataArray%) As Long
Declare Function flex_enable_camming_single_axis Lib "FlexMotion32.dll" (ByVal boardID&, ByVal axisID&, ByVal enable%, ByVal position#) As Long
Declare Function flex_acquire_trajectory_data Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axismap%, ByVal numberOfSamples%, ByVal timeperiod%) As Long
Declare Function flex_load_base_vel Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal baseVelocity%, ByVal inputVector%) As Long
Declare Function flex_load_blend_fact Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal blendFactor%, ByVal inputVector%) As Long
Declare Function flex_load_pos_modulus Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal positionModulus&, ByVal inputVector%) As Long
Declare Function flex_load_run_stop_threshold Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal runStopThreshold%, ByVal inputVector%) As Long
Declare Function flex_load_velocity_threshold Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal threshold&, ByVal inputVector%) As Long
Declare Function flex_load_rpm_thresh Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal threshold#, ByVal inputVector%) As Long
Declare Function flex_load_scurve_time Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal sCurveTime%, ByVal inputVector%) As Long
Declare Function flex_load_torque_lim Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal primaryPositiveLimit%, ByVal primaryNegativeLimit%, ByVal secondaryPositiveLimit%, ByVal secondaryNegativeLimit%, ByVal inputVector%) As Long
Declare Function flex_load_torque_offset Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal primaryOffset%, ByVal secondaryOffset%, ByVal inputVector%) As Long
Declare Function flex_load_velocity_override Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal overridePercentage%, ByVal inputVector%) As Long
Declare Function flex_read_dac Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrDAC%, ByVal returnVector%) As Long
Declare Function flex_read_dac_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrDAC%, DACValue%) As Long
Declare Function flex_read_dac_output_limit_status Lib "FlexMotion32.dll" (ByVal boardID%, ByVal returnVector%) As Long
Declare Function flex_read_dac_output_limit_status_rtn Lib "FlexMotion32.dll" (ByVal boardID%, positiveStatus%, negativeStatus%) As Long
Declare Function flex_read_steps_gen Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrStepperOutput%, ByVal returnVector%) As Long
Declare Function flex_read_steps_gen_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrStepperOutput%, steps&) As Long
Declare Function flex_read_target_pos Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal returnVector%) As Long
Declare Function flex_read_target_pos_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, targetPosition&) As Long
Declare Function flex_read_trajectory_data Lib "FlexMotion32.dll" (ByVal boardID%, ByVal returnVector%) As Long
Declare Function flex_read_trajectory_data_rtn Lib "FlexMotion32.dll" (ByVal boardID%, returnData&) As Long
Declare Function flex_configure_buffer Lib "FlexMotion32.dll" (ByVal boardID%, ByVal buffer%, ByVal resource%, ByVal bufferType%, ByVal bufferSize&, ByVal totalPoints&, ByVal oldDataStop%, ByVal requestedInterval#, actualInterval#) As Long
Declare Function flex_write_buffer Lib "FlexMotion32.dll" (ByVal boardID%, ByVal buffer%, ByVal numberOfPoints&, ByVal regenerationMode%, data&, ByVal inputVector%) As Long
Declare Function flex_read_buffer Lib "FlexMotion32.dll" (ByVal boardID%, ByVal buffer%, ByVal numberOfPoints&, ByVal returnVector%) As Long
Declare Function flex_read_buffer_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal buffer%, ByVal numberOfPoints&, data&) As Long
Declare Function flex_check_buffer Lib "FlexMotion32.dll" (ByVal boardID%, ByVal buffer%, ByVal returnVector%) As Long
Declare Function flex_check_buffer_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal buffer%, backlog&, bufferState%, pointsDone&) As Long
Declare Function flex_clear_buffer Lib "FlexMotion32.dll" (ByVal boardID%, ByVal buffer%) As Long
Declare Function flex_blend Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal axisOrVSMap%) As Long
Declare Function flex_start Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal axisOrVSMap%) As Long
Declare Function flex_stop_motion Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal stopType%, ByVal axisOrVSMap%) As Long
Declare Function flex_enable_home_inputs Lib "FlexMotion32.dll" (ByVal boardID%, ByVal homemap%) As Long
Declare Function flex_enable_axis_limit Lib "FlexMotion32.dll" (ByVal boardID%, ByVal limitType%, ByVal forwardLimitMap%, ByVal reverseLimitMap%) As Long
Declare Function flex_load_sw_lim_pos Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal forwardLimit&, ByVal reverseLimit&, ByVal inputVector%) As Long
Declare Function flex_read_home_input_status Lib "FlexMotion32.dll" (ByVal boardID%, ByVal returnVector%) As Long
Declare Function flex_read_home_input_status_rtn Lib "FlexMotion32.dll" (ByVal boardID%, homestatus%) As Long
Declare Function flex_read_axis_limit_status Lib "FlexMotion32.dll" (ByVal boardID%, ByVal limitType%, ByVal returnVector%) As Long
Declare Function flex_read_axis_limit_status_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal limitType%, forwardLimitStatus%, reverseLimitStatus%) As Long
Declare Function flex_set_home_polarity Lib "FlexMotion32.dll" (ByVal boardID%, ByVal homePolarityMap%) As Long
Declare Function flex_set_inhibit_output_momo Lib "FlexMotion32.dll" (ByVal boardID%, ByVal mustOn%, ByVal mustOff%) As Long
Declare Function flex_set_limit_input_polarity Lib "FlexMotion32.dll" (ByVal boardID%, ByVal forwardPolarityMap%, ByVal reversePolarityMap%) As Long
Declare Function flex_config_inhibit_output Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal enableInhibit%, ByVal polarity%, ByVal driveMode%) As Long
Declare Function flex_configure_breakpoint Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrEncoder%, ByVal enableMode%, ByVal actionOnBreakpoint%, ByVal operation%) As Long
Declare Function flex_configure_breakpoint_output Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrEncoder%, ByVal polarity%, ByVal driveMode%) As Long
Declare Function flex_enable_breakpoint Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrEncoder%, ByVal enable%) As Long
Declare Function flex_load_bp_modulus Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrEncoder%, ByVal breakpointModulus&, ByVal inputVector%) As Long
Declare Function flex_load_pos_bp Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrEncoder%, ByVal breakpointPosition&, ByVal inputVector%) As Long
Declare Function flex_read_breakpoint_status Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrEncoder%, ByVal breakpointType%, ByVal returnVector%) As Long
Declare Function flex_read_breakpoint_status_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrEncoder%, ByVal breakpointType%, breakpointStatus%) As Long
Declare Function flex_set_breakpoint_output_momo Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrEncoder%, ByVal mustOn%, ByVal mustOff%, ByVal inputVector%) As Long
Declare Function flex_enable_hs_capture Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrEncoder%, ByVal enable%) As Long
Declare Function flex_read_cap_pos Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrEncoder%, ByVal returnVector%) As Long
Declare Function flex_read_cap_pos_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrEncoder%, capturedPosition&) As Long
Declare Function flex_read_hs_cap_status Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrEncoder%, ByVal returnVector%) As Long
Declare Function flex_read_hs_cap_status_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrEncoder%, highSpeedCaptureStatus%) As Long
Declare Function flex_configure_hs_capture Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrEncoder%, ByVal captureMode%, ByVal operation%) As Long
Declare Function flex_configure_drive_signal Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal port%, ByVal pin%, ByVal mode%, ByVal polarity%) As Long
Declare Function flex_read_drive_signal_status Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal returnVector%) As Long
Declare Function flex_read_drive_signal_status_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, driveStatus%) As Long
Declare Function flex_load_motion_io_parameter Lib "FlexMotion32.dll" (ByVal boardID&, ByVal axisID&, ByVal attr%, data%) As Long
Declare Function flex_read_motion_io_execution_data Lib "FlexMotion32.dll" (ByVal boardID&, ByVal axisID&, ByVal attr%, data%) As Long
Declare Function flex_find_reference Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal axisOrVSMap%, ByVal findType%) As Long
Declare Function flex_wait_reference Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal axisOrVSMap%, ByVal timeout&, ByVal pollingInterval&, found%) As Long
Declare Function flex_read_reference_status Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal axisOrVSMap%, ByVal attr%, ByVal returnVector%) As Long
Declare Function flex_read_reference_status_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal axisOrVSMap%, ByVal attr%, value%) As Long
Declare Function flex_check_reference Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal axisOrVSMap%, found%, finding%) As Long
Declare Function flex_load_reference_parameter Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal findType%, ByVal attr%, ByVal value#) As Long
Declare Function flex_get_reference_parameter Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal findType%, ByVal attr%, value#) As Long
Declare Function flex_find_home Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal directionMap%) As Long
Declare Function flex_find_index Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal direction%, ByVal offset%) As Long
Declare Function flex_configure_pwm_output Lib "FlexMotion32.dll" (ByVal boardID%, ByVal PWMOutput%, ByVal enable%, ByVal clock%) As Long
Declare Function flex_enable_adcs Lib "FlexMotion32.dll" (ByVal boardID%, ByVal reserved%, ByVal ADCMap%) As Long
Declare Function flex_enable_encoders Lib "FlexMotion32.dll" (ByVal boardID%, ByVal encoderMap%) As Long
Declare Function flex_load_dac Lib "FlexMotion32.dll" (ByVal boardID%, ByVal DAC%, ByVal outputValue%, ByVal inputVector%) As Long
Declare Function flex_load_pwm_duty Lib "FlexMotion32.dll" (ByVal boardID%, ByVal PWMOutput%, ByVal dutycycle%, ByVal inputVector%) As Long
Declare Function flex_read_encoder Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrEncoder%, ByVal returnVector%) As Long
Declare Function flex_read_encoder_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrEncoder%, encoderCounts&) As Long
Declare Function flex_read_port Lib "FlexMotion32.dll" (ByVal boardID%, ByVal port%, ByVal returnVector%) As Long
Declare Function flex_read_port_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal port%, portData%) As Long
Declare Function flex_reset_encoder Lib "FlexMotion32.dll" (ByVal boardID%, ByVal encoder%, ByVal position&, ByVal inputVector%) As Long
Declare Function flex_select_signal Lib "FlexMotion32.dll" (ByVal boardID%, ByVal destination%, ByVal source%) As Long
Declare Function flex_read_adc16 Lib "FlexMotion32.dll" (ByVal boardID%, ByVal ADC%, ByVal returnVector%) As Long
Declare Function flex_read_adc16_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal ADC%, ADCValue&) As Long
Declare Function flex_set_adc_range Lib "FlexMotion32.dll" (ByVal boardID%, ByVal ADC%, ByVal range%) As Long
Declare Function flex_configure_encoder_filter Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrEncoder%, ByVal frequency%) As Long
Declare Function flex_configure_encoder_polarity Lib "FlexMotion32.dll" (ByVal boardID%, ByVal indexPolarity%, ByVal phaseAPolarity%, ByVal phaseBPolarity%) As Long
Declare Function flex_set_port_direction Lib "FlexMotion32.dll" (ByVal boardID%, ByVal port%, ByVal directionMap%) As Long
Declare Function flex_set_port Lib "FlexMotion32.dll" (ByVal boardID%, ByVal port%, ByVal mustOn%, ByVal mustOff%, ByVal inputVector%) As Long
Declare Function flex_set_port_pol Lib "FlexMotion32.dll" (ByVal boardID%, ByVal port%, ByVal portPolarityMap%) As Long
Declare Function flex_get_error_description Lib "FlexMotion32.dll" (ByVal descriptionType%, ByVal errorCode&, ByVal commandID%, ByVal resourceID%, ByVal charArray As String, sizeOfArray&) As Long
Declare Function flex_get_motion_board_info Lib "FlexMotion32.dll" (ByVal boardID%, ByVal informationType&, informationValue&) As Long
Declare Function flex_get_motion_board_name Lib "FlexMotion32.dll" (ByVal boardID%, ByVal charArray As String, sizeOfArray&) As Long
Declare Function flex_read_error_msg_rtn Lib "FlexMotion32.dll" (ByVal boardID%, commandID%, resourceID%, errorCode&) As Long
Declare Function flex_read_error_msg_detail_rtn Lib "FlexMotion32.dll" (ByVal boardID%, commandID%, resourceID%, errorCode&, lineNumber%, fileNumber%) As Long
Declare Function flex_get_last_error Lib "FlexMotion32.dll" (ByVal boardID%, commandID%, resourceID%, errorCode&) As Long
Declare Function flex_getu32 Lib "FlexMotion32.dll" (ByVal boardID%, ByVal resource%, ByVal attr%, value&) As Long
Declare Function flex_setu32 Lib "FlexMotion32.dll" (ByVal boardID%, ByVal resource%, ByVal attr%, ByVal value&) As Long
Declare Function flex_getf64 Lib "FlexMotion32.dll" (ByVal boardID%, ByVal resource%, ByVal attr%, value#) As Long
Declare Function flex_setf64 Lib "FlexMotion32.dll" (ByVal boardID%, ByVal resource%, ByVal attr%, ByVal value#) As Long
Declare Function flex_begin_store Lib "FlexMotion32.dll" (ByVal boardID%, ByVal program%) As Long
Declare Function flex_end_store Lib "FlexMotion32.dll" (ByVal boardID%, ByVal program%) As Long
Declare Function flex_insert_program_label Lib "FlexMotion32.dll" (ByVal boardID%, ByVal labelNumber%) As Long
Declare Function flex_jump_on_event Lib "FlexMotion32.dll" (ByVal boardID%, ByVal resource%, ByVal condition%, ByVal mustOn%, ByVal mustOff%, ByVal matchType%, ByVal labelNumber%) As Long
Declare Function flex_load_delay Lib "FlexMotion32.dll" (ByVal boardID%, ByVal delayTime&) As Long
Declare Function flex_pause_prog Lib "FlexMotion32.dll" (ByVal boardID%, ByVal program%) As Long
Declare Function flex_read_program_status Lib "FlexMotion32.dll" (ByVal boardID%, ByVal program%, ByVal returnVector%) As Long
Declare Function flex_read_program_status_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal program%, programStatus%) As Long
Declare Function flex_run_prog Lib "FlexMotion32.dll" (ByVal boardID%, ByVal program%) As Long
Declare Function flex_set_status_momo Lib "FlexMotion32.dll" (ByVal boardID%, ByVal mustOn%, ByVal mustOff%) As Long
Declare Function flex_stop_prog Lib "FlexMotion32.dll" (ByVal boardID%, ByVal program%) As Long
Declare Function flex_wait_on_event Lib "FlexMotion32.dll" (ByVal boardID%, ByVal resource%, ByVal waitType%, ByVal condition%, ByVal mustOn%, ByVal mustOff%, ByVal matchType%, ByVal timeout%, ByVal returnVector%) As Long
Declare Function flex_load_program_time_slice Lib "FlexMotion32.dll" (ByVal boardID%, ByVal program%, ByVal timeSlice%, ByVal inputVector%) As Long
Declare Function flex_load_description Lib "FlexMotion32.dll" (ByVal boardID%, ByVal object%, ByVal description As String) As Long
Declare Function flex_object_mem_manage Lib "FlexMotion32.dll" (ByVal boardID%, ByVal object%, ByVal operation%) As Long
Declare Function flex_read_description_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal object%, ByVal description As String) As Long
Declare Function flex_read_registry_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal Index%, rec As registryRecord) As Long
Declare Function flex_add_vars Lib "FlexMotion32.dll" (ByVal boardID%, ByVal variable1%, ByVal variable2%, ByVal returnVector%) As Long
Declare Function flex_and_vars Lib "FlexMotion32.dll" (ByVal boardID%, ByVal variable1%, ByVal variable2%, ByVal returnVector%) As Long
Declare Function flex_div_vars Lib "FlexMotion32.dll" (ByVal boardID%, ByVal variable1%, ByVal variable2%, ByVal returnVector%) As Long
Declare Function flex_load_var Lib "FlexMotion32.dll" (ByVal boardID%, ByVal value&, ByVal variable1%) As Long
Declare Function flex_lshift_var Lib "FlexMotion32.dll" (ByVal boardID%, ByVal variable1%, ByVal logicalShift%, ByVal returnVector%) As Long
Declare Function flex_mult_vars Lib "FlexMotion32.dll" (ByVal boardID%, ByVal variable1%, ByVal variable2%, ByVal returnVector%) As Long
Declare Function flex_not_var Lib "FlexMotion32.dll" (ByVal boardID%, ByVal variable1%, ByVal returnVector%) As Long
Declare Function flex_or_vars Lib "FlexMotion32.dll" (ByVal boardID%, ByVal variable1%, ByVal variable2%, ByVal returnVector%) As Long
Declare Function flex_read_var Lib "FlexMotion32.dll" (ByVal boardID%, ByVal variable1%, ByVal returnVector%) As Long
Declare Function flex_read_var_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal variable1%, value&) As Long
Declare Function flex_sub_vars Lib "FlexMotion32.dll" (ByVal boardID%, ByVal variable1%, ByVal variable2%, ByVal returnVector%) As Long
Declare Function flex_xor_vars Lib "FlexMotion32.dll" (ByVal boardID%, ByVal variable1%, ByVal variable2%, ByVal returnVector%) As Long
Declare Function flex_clear_pu_status Lib "FlexMotion32.dll" (ByVal boardID%) As Long
Declare Function flex_communicate Lib "FlexMotion32.dll" (ByVal boardID%, ByVal mode%, ByVal wordCount%, resource%, command%, data%, ByVal vector%) As Long
Declare Function flex_flush_rdb Lib "FlexMotion32.dll" (ByVal boardID%) As Long
Declare Function flex_enable_auto_start Lib "FlexMotion32.dll" (ByVal boardID%, ByVal enableOrDisable%, ByVal programToExecute%) As Long
Declare Function flex_enable_shutdown Lib "FlexMotion32.dll" (ByVal boardID%) As Long
Declare Function flex_enable_1394_watchdog Lib "FlexMotion32.dll" (ByVal boardID%, ByVal enableOrDisable%) As Long
Declare Function flex_read_csr_rtn Lib "FlexMotion32.dll" (ByVal boardID%, csr%) As Long
Declare Function flex_reset_defaults Lib "FlexMotion32.dll" (ByVal boardID%) As Long
Declare Function flex_save_defaults Lib "FlexMotion32.dll" (ByVal boardID%) As Long
Declare Function flex_set_irq_mask Lib "FlexMotion32.dll" (ByVal boardID%, ByVal mask%) As Long
Declare Function flex_read_board_temperature Lib "FlexMotion32.dll" (ByVal boardID%, temperature#) As Long
Declare Function flex_read_rdb Lib "FlexMotion32.dll" (ByVal boardID%, number%, wordCount%, resource%, command%, commandData%) As Long
Declare Function flex_configure_inhibit_output Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axismap%, ByVal polarityMap%) As Long
Declare Function flex_config_step_mode_pol Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrStepperOutput%, ByVal modeAndPolarityMap%) As Long
Declare Function flex_enable_axes Lib "FlexMotion32.dll" (ByVal boardID%, ByVal reserved%, ByVal PIDRate%, ByVal axismap%) As Long
Declare Function flex_configure_inhibits Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axismap%, ByVal polarityMap%) As Long
Declare Function flex_enable_limits Lib "FlexMotion32.dll" (ByVal boardID%, ByVal limitType%, ByVal forwardLimitMap%, ByVal reverseLimitMap%) As Long
Declare Function flex_read_limit_status Lib "FlexMotion32.dll" (ByVal boardID%, ByVal limitType%, ByVal returnVector%) As Long
Declare Function flex_read_limit_status_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal limitType%, forwardLimitStatus%, reverseLimitStatus%) As Long
Declare Function flex_set_limit_polarity Lib "FlexMotion32.dll" (ByVal boardID%, ByVal forwardPolarityMap%, ByVal reversePolarityMap%) As Long
Declare Function flex_read_dac_limit_status Lib "FlexMotion32.dll" (ByVal boardID%, ByVal returnVector%) As Long
Declare Function flex_read_dac_limit_status_rtn Lib "FlexMotion32.dll" (ByVal boardID%, positiveStatus%, negativeStatus%) As Long
Declare Function flex_set_breakpoint_momo Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrEncoder%, ByVal mustOn%, ByVal mustOff%, ByVal inputVector%) As Long
Declare Function flex_set_inhibit_momo Lib "FlexMotion32.dll" (ByVal boardID%, ByVal mustOn%, ByVal mustOff%) As Long
Declare Function flex_wait_on_condition Lib "FlexMotion32.dll" (ByVal boardID%, ByVal resource%, ByVal waitType%, ByVal condition%, ByVal mustOn%, ByVal mustOff%, ByVal matchType%, ByVal timeout%, ByVal returnVector%) As Long
Declare Function flex_jump_label_on_condition Lib "FlexMotion32.dll" (ByVal boardID%, ByVal resource%, ByVal condition%, ByVal mustOn%, ByVal mustOff%, ByVal matchType%, ByVal labelNumber%) As Long
Declare Function flex_read_adc Lib "FlexMotion32.dll" (ByVal boardID%, ByVal ADC%, ByVal returnVector%) As Long
Declare Function flex_read_adc_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal ADC%, ADCValue%) As Long
Declare Function flex_enable_encs Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal encmap%) As Long
Declare Function flex_load_2d_arc Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal radius&, ByVal startangle&, ByVal travelangle&, ByVal inpvect%) As Long
Declare Function flex_load_3d_arc Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal radius&, ByVal planepitch&, ByVal planeyaw&, ByVal startangle&, ByVal travelangle&, ByVal inpvect%) As Long
Declare Function flex_load_helix Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal radius&, ByVal startangle&, ByVal travelangle&, ByVal lineartravel&, ByVal inpvect%) As Long
Declare Function flex_enable_gear Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal axismap%) As Long
Declare Function flex_acq_traj_data Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal axismap%, ByVal numsamps%, ByVal timeperiod%) As Long
Declare Function flex_read_traj_data Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal retvect%) As Long
Declare Function flex_read_traj_data_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, retdata%) As Long
Declare Function flex_read_rs_status Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal retvect%) As Long
Declare Function flex_read_rs_status_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, status%) As Long
Declare Function flex_read_motoff_status Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal retvect%) As Long
Declare Function flex_read_motoff_status_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, status%) As Long
Declare Function flex_read_dac_status Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal retvect%) As Long
Declare Function flex_read_dac_status_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, status%) As Long
Declare Function flex_load_loop_params Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, PIDValues As PidVals, ByVal inputVector%) As Long
Declare Function flex_load_kp Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal kp%, ByVal inpvect%) As Long
Declare Function flex_load_ki Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal ki%, ByVal inpvect%) As Long
Declare Function flex_load_ilim Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal ilim%, ByVal inpvect%) As Long
Declare Function flex_load_kd Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal kd%, ByVal inpvect%) As Long
Declare Function flex_load_td Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal td%, ByVal inpvect%) As Long
Declare Function flex_load_kv Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal kv%, ByVal inpvect%) As Long
Declare Function flex_load_aff Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal aff%, ByVal inpvect%) As Long
Declare Function flex_load_vff Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal vff%, ByVal inpvect%) As Long
Declare Function flex_load_accel Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal accel&, ByVal inpvect%) As Long
Declare Function flex_load_decel Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal decel&, ByVal inpvect%) As Long
Declare Function flex_load_accel_rpsps Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal rpsps#, ByVal inpvect%) As Long
Declare Function flex_load_decel_rpsps Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal rpsps#, ByVal inpvect%) As Long
Declare Function flex_load_vel_override Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal overridescale%, ByVal inpvect%) As Long
Declare Function flex_stop Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal axismap%) As Long
Declare Function flex_halt Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal axismap%) As Long
Declare Function flex_kill Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal axismap%) As Long
Declare Function flex_enable_lims Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal limmap%) As Long
Declare Function flex_enable_sw_lims Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal swlimmap%) As Long
Declare Function flex_read_lim_status Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal retvect%) As Long
Declare Function flex_read_lim_status_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, limstatus%) As Long
Declare Function flex_read_sw_lim_status Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal retvect%) As Long
Declare Function flex_read_sw_lim_status_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, swlimstatus%) As Long
Declare Function flex_read_pos_bp_status Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal retvect%) As Long
Declare Function flex_read_pos_bp_status_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, bpstatus%) As Long
Declare Function flex_read_at_bp_status Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal retvect%) As Long
Declare Function flex_read_at_bp_status_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, status%) As Long
Declare Function flex_save_object Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%) As Long
Declare Function flex_delete_object Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%) As Long
Declare Function flex_free_object Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%) As Long
Declare Function flex_wait Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal cc%, ByVal mustOn%, ByVal mustOff%, ByVal timeout%, ByVal retvect%) As Long
Declare Function flex_wait_or Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal cc%, ByVal mustOn%, ByVal mustOff%, ByVal timeout%, ByVal retvect%) As Long
Declare Function flex_jump_label Lib "FlexMotion32.dll" (ByVal boardID%, ByVal resource%, ByVal condition%, ByVal mustOn%, ByVal mustOff%, ByVal labelNumber%) As Long
Declare Function flex_read_boot_ver_rtn Lib "FlexMotion32.dll" (ByVal boardID%, bootver&) As Long
Declare Function flex_read_firm_ver_rtn Lib "FlexMotion32.dll" (ByVal boardID%, firmver&) As Long
Declare Function flex_read_dsp_ver_rtn Lib "FlexMotion32.dll" (ByVal boardID%, dspver&) As Long
Declare Function flex_read_fpga_ver_rtn Lib "FlexMotion32.dll" (ByVal boardID%, fpgaver&) As Long
Declare Function flex_set_home_pol Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal homepol%) As Long
Declare Function flex_enable_homes Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal homemap%) As Long
Declare Function flex_set_lim_pol Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal polarity%) As Long
Declare Function flex_set_inhib_momo Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal mustOn%, ByVal mustOff%) As Long
Declare Function flex_config_inhibs Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal axismap%, ByVal polmap%) As Long
Declare Function flex_read_home_status Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal retvect%) As Long
Declare Function flex_read_home_status_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, homestatus%) As Long
Declare Function flex_read_vel_status Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal retvect%) As Long
Declare Function flex_read_vel_status_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, status%) As Long
Declare Function flex_insert_label Lib "FlexMotion32.dll" (ByVal boardID%, ByVal device%, ByVal labelnum%) As Long
Declare Function flex_load_vel Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal velocity&, ByVal inputVector%) As Long
Declare Function flex_read_vel Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal returnVector%) As Long
Declare Function flex_read_vel_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, velocity&) As Long
Declare Function flex_load_vel_thresh Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal threshold&, ByVal inputVector%) As Long
Declare Function flex_set_port_dir Lib "FlexMotion32.dll" (ByVal boardID%, ByVal port%, ByVal direction%) As Long
Declare Function flex_load_counts_rev Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal counts%) As Long
Declare Function flex_load_steps_rev Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal steps%) As Long
Declare Function flex_set_bp_momo Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrEncoder%, ByVal mustOn%, ByVal mustOff%) As Long
Declare Function flex_set_hs_cap_pol Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrEncoder%, ByVal highSpeedCapturePolarity%) As Long
Declare Function flex_enable_hs_caps Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrEncoder%, ByVal captureMap%) As Long
Declare Function flex_enable_bp Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisOrEncoder%, ByVal enableMode%, ByVal actionOnBreakpoint%) As Long
Declare Function flex_set_port_momo Lib "FlexMotion32.dll" (ByVal boardID%, ByVal port%, ByVal mustOn%, ByVal mustOff%) As Long
Declare Function flex_read_err_msg_rtn Lib "FlexMotion32.dll" (ByVal boardID%, commandID%, resourceID%, errorCode%) As Long
Declare Function flex_load_vel_tc_rs Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal filterTime%, ByVal runStopThreshold%, ByVal inputVector%) As Long
Declare Function flex_read_raw_vel Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, ByVal retvect%) As Long
Declare Function flex_read_raw_vel_rtn Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axisorvectorspace%, rawvel&) As Long
Declare Function flex_set_axis_mode Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal axismode%) As Long
