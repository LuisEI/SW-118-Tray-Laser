Attribute VB_Name = "MotnErr"
'*****************************************************************************************
'
'   MotnErr.bas
'
'   This file contains standard definitions used by the NI-Motion motion controller board families.
'*****************************************************************************************/

Global Const NIMC_noError = 0
Global Const NIMC_readyToReceiveTimeoutError = -70001
Global Const NIMC_currentPacketError = -70002
Global Const NIMC_noReturnDataBufferError = -70003
Global Const NIMC_halfReturnDataBufferError = -70004
Global Const NIMC_boardFailureError = -70005
Global Const NIMC_badResourceIDOrAxisError = -70006
Global Const NIMC_CIPBitError = -70007
Global Const NIMC_previousPacketError = -70008
Global Const NIMC_packetErrBitNotClearedError = -70009
Global Const NIMC_badCommandError = -70010
Global Const NIMC_badReturnDataBufferPacketError = -70011
Global Const NIMC_badBoardIDError = -70012
Global Const NIMC_packetLengthError = -70013
Global Const NIMC_closedLoopOnlyError = -70014
Global Const NIMC_returnDataBufferFlushError = -70015
Global Const NIMC_servoOnlyError = -70016
Global Const NIMC_stepperOnlyError = -70017
Global Const NIMC_closedLoopStepperOnlyError = -70018
Global Const NIMC_noBoardConfigInfoError = -70019
Global Const NIMC_countsNotConfiguredError = -70020
Global Const NIMC_systemResetError = -70021
Global Const NIMC_functionSupportError = -70022
Global Const NIMC_parameterValueError = -70023
Global Const NIMC_motionOnlyError = -70024
Global Const NIMC_returnDataBufferNotEmptyError = -70025
Global Const NIMC_modalErrorsReadError = -70026
Global Const NIMC_processTimeoutError = -70027
Global Const NIMC_insufficientSizeError = -70028
Global Const NIMC_reserved4Error = -70029
Global Const NIMC_reserved5Error = -70030
Global Const NIMC_reserved6Error = -70031
Global Const NIMC_reserved7Error = -70032
Global Const NIMC_badPointerError = -70033
Global Const NIMC_wrongReturnDataError = -70034
Global Const NIMC_watchdogTimeoutError = -70035
Global Const NIMC_invalidRatioError = -70036
Global Const NIMC_irrelevantAttributeError = -70037
Global Const NIMC_internalSoftwareError = -70038
Global Const NIMC_1394WatchdogEnableError = -70039
Global Const NIMC_reservedOnBoardProgramError = -70040
Global Const NIMC_boardIDInUseError = -70041
Global Const NIMC_RemoteConnectionFailureError = -70042
Global Const NIMC_calibrationOutOfRangeError = -70043
Global Const NIMC_calibrationStepError = -70044
Global Const NIMC_axesNotKilledError = -70045
Global Const NIMC_invalidCalibrationDataError = -70046
Global Const NIMC_modeNotSupportedError = -70047
Global Const NIMC_invalidBreakpointWindowError = -70048
Global Const NIMC_downloadChecksumError = -70049
Global Const NIMC_reserved50Error = -70050
Global Const NIMC_firmwareDownloadError = -70051
Global Const NIMC_FPGAProgramError = -70052
Global Const NIMC_DSPInitializationError = -70053
Global Const NIMC_corrupt68331FirmwareError = -70054
Global Const NIMC_corruptDSPFirmwareError = -70055
Global Const NIMC_corruptFPGAFirmwareError = -70056
Global Const NIMC_interruptConfigurationError = -70057
Global Const NIMC_IOInitializationError = -70058
Global Const NIMC_flashromCopyError = -70059
Global Const NIMC_corruptObjectSectorError = -70060
Global Const NIMC_bufferInUseError = -70061
Global Const NIMC_oldDataStopError = -70062
Global Const NIMC_bufferConfigurationError = -70063
Global Const NIMC_illegalBufferOperation = -70064
Global Const NIMC_illegalContouringError = -70065
Global Const NIMC_virtualBoardError = -70066
Global Const NIMC_maxBreakpointFrequencyError = -70067
Global Const NIMC_maxHSCaptureFrequencyError = -70068
Global Const NIMC_invalidHallSensorStateError = -70069
Global Const NIMC_commutationModeError = -70070
Global Const NIMC_DIOReservedForHallSensors = -70071
Global Const NIMC_boardInPowerUpResetStateError = -70072
Global Const NIMC_boardInShutDownStateError = -70073
Global Const NIMC_shutDownFailedError = -70074
Global Const NIMC_hostFIFOBufferFullError = -70075
Global Const NIMC_noHostDataError = -70076
Global Const NIMC_corruptHostDataError = -70077
Global Const NIMC_invalidFunctionDataError = -70078
Global Const NIMC_autoStartFailedError = -70079
Global Const NIMC_returnDataBufferFullError = -70080
Global Const NIMC_reserved81Error = -70081
Global Const NIMC_reserved82Error = -70082
Global Const NIMC_DSPXmitBufferFullError = -70083
Global Const NIMC_DSPInvalidCommandError = -70084
Global Const NIMC_DSPInvalidDeviceError = -70085
Global Const NIMC_invalidFeedbackResetPositionError = -70086
Global Const NIMC_blendNotCompleteError = -70087
Global Const NIMC_invalidFeedbackDeviceError = -70088
Global Const NIMC_axisFindingReferenceError = -70089
Global Const NIMC_onboardProgramSupportError = -70090
Global Const NIMC_availableForUse91 = -70091
Global Const NIMC_DSPXmitDataError = -70092
Global Const NIMC_DSPCommunicationsError = -70093
Global Const NIMC_DSPMessageBufferEmptyError = -70094
Global Const NIMC_DSPCommunicationsTimeoutError = -70095
Global Const NIMC_passwordError = -70096
Global Const NIMC_mustOnMustOffConflictError = -70097
Global Const NIMC_reserved98Error = -70098
Global Const NIMC_reserved99Error = -70099
Global Const NIMC_IOEventCounterError = -70100
Global Const NIMC_reserved101Error = -70101
Global Const NIMC_wrongIODirectionError = -70102
Global Const NIMC_wrongIOConfigurationError = -70103
Global Const NIMC_outOfEventsError = -70104
Global Const NIMC_IOReservedError = -70105
Global Const NIMC_outputDeviceNotAssignedError = -70106
Global Const NIMC_splineUnderflowError = -70107
Global Const NIMC_PIDUpdateRateError = -70108
Global Const NIMC_feedbackDeviceNotAssignedError = -70109
Global Const NIMC_reserved110Error = -70110
Global Const NIMC_axisConfigurationSwitchError = -70111
Global Const NIMC_axisConfigurationClLoopError = -70112
Global Const NIMC_noMoreRAMError = -70113
Global Const NIMC_reserved114Error = -70114
Global Const NIMC_jumpToInvalidLabelError = -70115
Global Const NIMC_invalidConditionCodeError = -70116
Global Const NIMC_homeLimitNotEnabledError = -70117
Global Const NIMC_findHomeError = -70118
Global Const NIMC_limitSwitchActiveError = -70119
Global Const NIMC_softwareUpdateRequiredError = -70120
Global Const NIMC_positionRangeError = -70121
Global Const NIMC_encoderDisabledError = -70122
Global Const NIMC_moduloBreakpointError = -70123
Global Const NIMC_findIndexError = -70124
Global Const NIMC_wrongModeError = -70125
Global Const NIMC_axisConfigurationError = -70126
Global Const NIMC_pointsTableFullError = -70127
Global Const NIMC_available128Error = -70128
Global Const NIMC_axisDisabledError = -70129
Global Const NIMC_memoryRangeError = -70130
Global Const NIMC_inPositionUpdateError = -70131
Global Const NIMC_targetPositionUpdateError = -70132
Global Const NIMC_pointRequestMissingError = -70133
Global Const NIMC_internalSamplesMissingError = -70134
Global Const NIMC_reserved135Error = -70135
Global Const NIMC_eventTimeoutError = -70136
Global Const NIMC_objectReferenceError = -70137
Global Const NIMC_outOfMemoryError = -70138
Global Const NIMC_registryFullError = -70139
Global Const NIMC_noMoreProgramPlayerError = -70140
Global Const NIMC_programOverruleError = -70141
Global Const NIMC_followingErrorOverruleError = -70142
Global Const NIMC_reserved143Error = -70143
Global Const NIMC_illegalVariableError = -70144
Global Const NIMC_illegalVectorSpaceError = -70145
Global Const NIMC_noMoreSamplesError = -70146
Global Const NIMC_slaveAxisKilledError = -70147
Global Const NIMC_ADCDisabledError = -70148
Global Const NIMC_operationModeError = -70149
Global Const NIMC_followingErrorOnFindHomeError = -70150
Global Const NIMC_invalidVelocityError = -70151
Global Const NIMC_invalidAccelerationError = -70152
Global Const NIMC_samplesBufferFullError = -70153
Global Const NIMC_illegalVectorError = -70154
Global Const NIMC_QSPIFailedError = -70155
Global Const NIMC_reserved156Error = -70156
Global Const NIMC_pointsBufferFullError = -70157
Global Const NIMC_axisInitializationError = -70158
Global Const NIMC_encoderInitializationError = -70159
Global Const NIMC_stepChannelInitializationError = -70160
Global Const NIMC_blendFactorConflictError = -70161
Global Const NIMC_torqueOffsetError = -70162
Global Const NIMC_invalidLimitRangeError = -70163
Global Const NIMC_ADCConfigurationError = -70164
Global Const NIMC_findReferenceError = -70165
Global Const NIMC_followingErrorOnFindReference = -70166
Global Const NIMC_initializationInProgress = -70167
Global Const NIMC_invalidMotionIDError = -70168
Global Const NIMC_invalidPointerError = -70169
Global Const NIMC_interfaceNotSupportedError = -70170
Global Const NIMC_breakpointBufferFullError = -70171
Global Const NIMC_hsCaptureBufferFullError = -70172
Global Const NIMC_internalBreakpointMissingError = -70173
Global Const NIMC_internalHSCaptureMissingError = -70174
Global Const NIMC_arcPointsBufferFullError = -70175
Global Const NIMC_timeGuaranteeError = -70176
Global Const NIMC_invalidTimeSliceError = -70177
Global Const NIMC_onlyInAProgramError = -70178
Global Const NIMC_invalidMasterAxisError = -70179
Global Const NIMC_invalidMasterEnabledError = -70180
Global Const NIMC_remoteBoardMismatchError = -70181
Global Const NIMC_deviceNotActivatedError = -70182
Global Const NIMC_dataTransmissionError = -70183
Global Const NIMC_deviceTypeNotSupported = -70184
Global Const NIMC_serializationFailedError = -70185
Global Const NIMC_homeSwitchActiveError = -70186
Global Const NIMC_incorrectBufferSizeError = -70187
Global Const NIMC_invalidBufferHandleSpecifiedError = -70188
Global Const NIMC_blendNotAllowedInThisModeError = -70189
Global Const NIMC_invalidLoopRateError = -70190
Global Const NIMC_axisCommunicationWatchdogError = -70191
Global Const NIMC_communicationInterfaceNotFoundError = -70192
Global Const NIMC_axisAlreadyAddedError = -70194
Global Const NIMC_invalidAxisScaleError = -70195
Global Const NIMC_axisNotPresentError = -70196
Global Const NIMC_startBlockedDueToFollowingError = -70197
Global Const NIMC_configurationFileNotFoundError = -70198
Global Const NIMC_gearingMasterNotActiveError = -70199
Global Const NIMC_controllerNotInPowerUpResetStateError = -70200
Global Const NIMC_vectorSpaceCannotBeConfiguredError = -70201
Global Const NIMC_failedToAddDeviceError = -70202
Global Const NIMC_noMoreBufferError = -70203
Global Const NIMC_badDeviceOrAxisError = NIMC_badResourceIDOrAxisError
