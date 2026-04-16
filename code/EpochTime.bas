Attribute VB_Name = "EpochTime"
'@Folder("Library")
Option Explicit

'@Description("Return the number of seconds since Midnight, January 1, 1970.")
Public Function TimestampNow() As LongLong
Attribute TimestampNow.VB_Description = "Return the number of seconds since Midnight, January 1, 1970."
    TimestampNow = DateDiff("s", "1/1/1970 00:00:00", Now)
End Function
