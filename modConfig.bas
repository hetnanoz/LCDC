Option Explicit

Private Const CLASS_NAME As String = "modConfig"

' Set to True only during development/debugging.
' Production value must remain False.
Public Const DEV_MODE As Boolean = False

' Microsoft Information Protection - Internal label.
Public Const MIP_INTERNAL_LABEL_ID As String = _
    "8ffbc0b8-e97b-47d1-beac-cb0955d66f3b"

Public Const MIP_SITE_ID As String = _
    "614f9c25-bffa-42c7-86d8-964101f55fa2"

Public Const MIP_ASSIGNMENT_METHOD As Long = 1
Public Const MIP_CONTENT_BITS As Long = 2
