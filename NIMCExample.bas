Attribute VB_Name = "NIMCExample"
'////////////////////////////////////////////////////////////////////////////////
'
'      NIMCExample.bas
'      General implemenation file for all NI Motion Control Examples
'
'////////////////////////////////////////////////////////////////////////////////
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'////////////////////////////////////////////////
'   CheckError - Tests the result and raises
'       the error flag
Public Function CheckError(ByVal result As Long)
    If result = 0 Then
        Exit Function
    Else
        nimcDisplaySimpleError result
    End If
End Function

'////////////////////////////////////////////////
'   Function nimcDisplayError- Displays the error
'       message in a msgbox

Public Function nimcDisplayError(ByVal errorCode As Long, ByVal commandID As Integer, ByVal resourceID As Integer) As Integer
    Dim errorDescription As String  'String -  to get error description
    Dim sizeOfArray As Long         'Size of error description
    Dim descriptionType As Integer  'The type of description to be printed
    Dim status As Long              'Error returned by function

    'If no commandID is passed the only the error description needs to
    'be displayed - else the combined description of the error and the
    'function (and resource) in which it occurred should be displayed
    If (commandID = 0) Then
        descriptionType = NIMC_ERROR_ONLY
    Else
        descriptionType = NIMC_COMBINED_DESCRIPTION
    End If
    
    'First get the size for the error description
    sizeOfArray = 0
    errorDescription = "" 'Setting this to NULL returns the size required
    status = flex_get_error_description(descriptionType, errorCode, commandID, resourceID, errorDescription, sizeOfArray)

    'Allocate memory on the heap for the description
    sizeOfArray = sizeOfArray + 1 'So that the sizeOfArray is size of description + NULL character
    errorDescription = Space$(sizeOfArray)
    
    'Get Error Description
    status = flex_get_error_description(descriptionType, errorCode, commandID, resourceID, errorDescription, sizeOfArray)

    If Not (errorDescription = "") Then
        status = MsgBox(errorDescription, vbOKOnly + vbCritical + vbApplicationModal + vbDefaultButton2, "Error")
        End
   End If
End Function

'////////////////////////////////////////////////
'   Function nimcDisplaySimpleError- Displays the error description
'       in a message box.
Public Function nimcDisplaySimpleError(ByVal errorCode As Long) As Integer
    Dim errorDescription As String  'String -  to get error description
    Dim sizeOfArray As Long         'Size of error description
    Dim descriptionType As Integer  'The type of description to be printed
    Dim status As Long              'Error returned by function
 
    'Just get the error description for now
    descriptionType = NIMC_ERROR_ONLY
    
    'First get the size for the error description
    sizeOfArray = 0
    errorDescription = "" 'Setting this to NULL returns the size required
    status = flex_get_error_description(descriptionType, errorCode, 0, 0, errorDescription, sizeOfArray)

    'Allocate memory on the heap for the description
    sizeOfArray = sizeOfArray + 1 'So that the sizeOfArray is size of description + NULL character
    errorDescription = Space$(sizeOfArray)
    
    'Get Error Description
    status = flex_get_error_description(descriptionType, errorCode, commandID, resourceID, errorDescription, sizeOfArray)

    If Not (errorDescription = "") Then
        
        status = MsgBox(errorDescription, vbOKOnly + vbCritical + vbApplicationModal + vbDefaultButton2, "Error")
        End
                        
   End If
   
End Function



