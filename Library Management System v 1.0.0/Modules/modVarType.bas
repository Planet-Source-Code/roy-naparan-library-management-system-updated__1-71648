Attribute VB_Name = "modVarType"
''Variable type for GetCursorPos API function
Public Type POINTAPI
        x As Long
        Y As Long
End Type
''Variable type for tracking current users
''information
Public Type USER_INFO
        user_id     As Long
        user_name   As String
        time_login  As Date
End Type


''Varible type for paging
Public Type PAGE_INFO
        PAGE_CURRENT        As Long
        PAGE_NEXT           As Long
        PAGE_PREVIOUS       As Long
        PAGE_TOTAL          As Long
End Type



