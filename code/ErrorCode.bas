Attribute VB_Name = "ErrorCode"
'@IgnoreModule ConstantNotUsed
'@Folder("Library")

' Reference:
' - https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/raise-method
' - https://onlinelibrary.wiley.com/doi/pdf/10.1002/9781118257616.app3

' Built-In (0-512, and 1004)
' 0-9
Public Const RETURN_WITHOUT_GOSUB As Long = vbObjectError + 3
Public Const INVALID_PROCEDURE_CALL As Long = vbObjectError + 5
Public Const INVALID_ARGUMENT As Long = vbObjectError + 5
Public Const OVERFLOW As Long = vbObjectError + 6
Public Const OUT_OF_MEMORY As Long = vbObjectError + 7
Public Const SUBSCRIPT_OUT_OF_RANGE As Long = vbObjectError + 9

' 10-19
Public Const ARRAY_FIXED As Long = vbObjectError + 10
Public Const ARRAY_TEMPORARILY_LOCKED As Long = vbObjectError + 10
Public Const DIVISION_BY_ZERO As Long = vbObjectError + 11
Public Const TYPE_MISMATCH As Long = vbObjectError + 13
Public Const OUT_OF_STRING_SPACE As Long = vbObjectError + 14
Public Const EXPRESSION_TOO_COMPLEX As Long = vbObjectError + 16
Public Const CANNOT_PERFORM_REQUESTED_OPERATION As Long = vbObjectError + 17
Public Const USER_INTERUPT As Long = vbObjectError + 18

' 20-29
Public Const RESUME_WITHOUT_ERROR As Long = vbObjectError + 20
Public Const OUT_OF_STACK_SPACE As Long = vbObjectError + 28

' 30-39
Public Const SUB_NOT_DEFINED As Long = vbObjectError + 35
Public Const FUNCTION_NOT_DEFINED As Long = vbObjectError + 35

' 40-49
Public Const TOO_MANY_DLL_APP_CLIENTS As Long = vbObjectError + 47
Public Const ERROR_LOADING_DLL As Long = vbObjectError + 48
Public Const BAD_DLL_CALLING_CONVENTION As Long = vbObjectError + 49

' 50-59
Public Const INTERNAL_ERROR As Long = vbObjectError + 51
Public Const BAD_FILENAME As Long = vbObjectError + 52
Public Const BAD_NUMBER As Long = vbObjectError + 52
Public Const FILE_NOT_FOUND As Long = vbObjectError + 53
Public Const BAD_FILE_MODE As Long = vbObjectError + 54
Public Const FILE_ALREADY_OPEN As Long = vbObjectError + 55
Public Const DEVICE_IO_ERROR As Long = vbObjectError + 57
Public Const FILE_ALREADY_EXISTS As Long = vbObjectError + 58
Public Const BAD_RECORD_LENGTH As Long = vbObjectError + 59

' 60-69
Public Const DISK_FULL As Long = vbObjectError + 61
Public Const INPUT_PAST_EOF As Long = vbObjectError + 62
Public Const BAD_RECORD_NUMBER As Long = vbObjectError + 63
Public Const TOO_MANY_FILES As Long = vbObjectError + 67
Public Const DEVICE_UNAVAILABLE As Long = vbObjectError + 68

' 70-79
Public Const PERMISSION_DENIED As Long = vbObjectError + 70
Public Const DISK_NOT_READY As Long = vbObjectError + 71
Public Const CANNOT_RENAME_WITH_DIFFERENT_DRIVE As Long = vbObjectError + 74
Public Const PATH_ACCESS_ERROR As Long = vbObjectError + 75
Public Const FILE_ACCESS_ERROR As Long = vbObjectError + 75
Public Const PATH_NOT_FOUND As Long = vbObjectError + 76

' 90-99
Public Const OBJECT_VARIABLE_NOT_SET As Long = vbObjectError + 91
Public Const OBJECT_WITH_BLOCK_NOT_SET As Long = vbObjectError + 91
Public Const FOR_LOOP_NOT_INITIALIZED As Long = vbObjectError + 92
Public Const INVALID_PATTERN_STRING As Long = vbObjectError + 93
Public Const INVALID_USE_OF_NULL As Long = vbObjectError + 94
Public Const UNABLE_TO_SINK_EVENTS_OBJECT As Long = vbObjectError + 96 ' is this a typo in the reference? should it be sync?
Public Const INVALID_FRIEND_FUNCTION_CALL As Long = vbObjectError + 97
Public Const PROPERTY_REFERENCES_PRIVATE_OBJECT As Long = vbObjectError + 98
Public Const METHOD_REFERENCES_PRIVATE_OBJECT As Long = vbObjectError + 98

' 320-329
Public Const INVALID_FILE_FORMAT As Long = vbObjectError + 321
Public Const CANNOT_CREATE_NECESSARY_TEMP_FILE As Long = vbObjectError + 322
Public Const INVALID_FORMAT_IN_RESOURCE_FILE As Long = vbObjectError + 325

' 380-389
Public Const INVALID_PROPERTY_VALUE As Long = vbObjectError + 380
Public Const INVALID_PROPERTY_ARRAY_INDEX As Long = vbObjectError + 381
Public Const SET_NOT_SUPPORTED_AT_RUNTIME As Long = vbObjectError + 382
Public Const SET_NOT_SUPPORTED_READ_ONLY As Long = vbObjectError + 383
Public Const NEED_PROPERTY_ARRAY_INDEX As Long = vbObjectError + 385
Public Const SET_NOT_PERMITTED As Long = vbObjectError + 387

' 390-399
Public Const GET_NOT_SUPPORTED_AT_RUNTIME As Long = vbObjectError + 393
Public Const GET_NOT_SUPPORTED_WRITE_ONLY As Long = vbObjectError + 394

' 420-429
Public Const PROPERTY_NOT_FOUND As Long = vbObjectError + 422
Public Const PROPERTY_OR_METHOD_NOT_FOUND As Long = vbObjectError + 423
Public Const OBJECT_REQUIRED As Long = vbObjectError + 424
Public Const ACTIVEX_COMPONENT_CANNOT_CREATE_OBJECT As Long = vbObjectError + 429

' 430-439
Public Const AUTOMATION_NOT_SUPPORTED_BY_CLASS As Long = vbObjectError + 430
Public Const INTERFACE_NOT_SUPPORTED_BY_CLASS As Long = vbObjectError + 430
Public Const FILENAME_NOT_FOUND_DURING_AUTOMATION As Long = vbObjectError + 432
Public Const CLASS_NAME_NOT_FOUND_DURING_AUTOMATION As Long = vbObjectError + 432
Public Const PROPERTY_NOT_SUPPORTED_BY_OBJECT As Long = vbObjectError + 438
Public Const METHOD_NOT_SUPPORTED_BY_OBJECT As Long = vbObjectError + 438

' 440-449
Public Const AUTOMATION_ERROR As Long = vbObjectError + 440
Public Const CONNECTION_TO_LIBRARY_LOST As Long = vbObjectError + 442
Public Const AUTOMATION_OBJECT_WITHOUT_DEFAULT_VALUE As Long = vbObjectError + 443
Public Const ACTION_NOT_SUPPORTED_BY_OBJECT As Long = vbObjectError + 445
Public Const NAMED_ARGUMENTS_NOT_SUPPORTED_BY_OBJECT As Long = vbObjectError + 446
Public Const CURRENT_LOCALE_SETTING_NOT_SUPPORTED_BY_OBJECT As Long = vbObjectError + 447
Public Const NAMED_ARGUMENT_NOT_FOUND As Long = vbObjectError + 448
Public Const ARGUMENT_NOT_OPTIONAL As Long = vbObjectError + 449

' 450-459
Public Const WRONG_NUMBER_OF_ARGUMENTS As Long = vbObjectError + 450
Public Const INVALID_PROPERTY_ASSIGNMENT As Long = vbObjectError + 450
Public Const PROPERTY_LET_NOT_DEFINED_AND_PROPERTY_GET_DID_NOT_RETURN_OBJECT As Long = vbObjectError + 451
Public Const INVALID_ORDINAL As Long = vbObjectError + 452
Public Const SPECIFIED_DLL_FUNCTION_NOT_FOUND As Long = vbObjectError + 453
Public Const CODE_RESOURCE_NOT_FOUND As Long = vbObjectError + 454
Public Const CODE_RESOURCE_LOCK_ERROR As Long = vbObjectError + 455
Public Const KEY_ALREADY_IN_USE As Long = vbObjectError + 457
Public Const AUTOMATION_TYPE_NOT_SUPPORTED_IN_VB As Long = vbObjectError + 458
Public Const SET_OF_EVENTS_NOT_SUPPORTED_BY_OBJECT_OR_CLASS As Long = vbObjectError + 459

' 460-469
Public Const INVALID_CLIPBOARD_FORMAT As Long = vbObjectError + 460
Public Const METHOD_NOT_FOUND As Long = vbObjectError + 461
Public Const DATA_MEMBER_NOT_FOUND As Long = vbObjectError + 461
Public Const REMOTE_SERVER_NOT_FOUND As Long = vbObjectError + 462 ' not available or doesn't exist
Public Const CLASS_NOT_REGISTERED_ON_LOCAL_MACHINE As Long = vbObjectError + 463

' 480-489
Public Const INVALID_PICTURE As Long = vbObjectError + 481
Public Const PRINTER_ERROR As Long = vbObjectError + 482

' 730-739
Public Const CANNOT_SAVE_FILE_TO_TEMP As Long = vbObjectError + 735

' 740-749
Public Const SEARCH_TEXT_NOT_FOUND As Long = vbObjectError + 744
Public Const REPLACEMENTS_TOO_LONG As Long = vbObjectError + 746

' 1004
Public Const APPLICATION_DEFINED_ERROR As Long = vbObjectError + 1004
Public Const OBJECT_DEFINED_ERROR As Long = vbObjectError + 1004

' Custom (513-65535, excluding 1004)
Public Const LET_READ_ONLY_PROPERTY As Long = vbObjectError + 513

