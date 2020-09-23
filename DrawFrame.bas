Attribute VB_Name = "DrawFrame"
Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'############################### Un1
Public Const CheckOrOption_Buton = &H4
    '######################### Un2
    Public Const SleectedOptionNormal = &H404
    Public Const OptionEmpty = &H24
    Public Const GrayedOutOptionChecked = &H504 '604
    Public Const GrayedOutOptionNormal = &H104
    Public Const CheckedCheckBoxNormal = &H400
    Public Const GrayedOutCheckedCheck = &H600
    Public Const GrayedOutCheckBoxEmpty = &H128
    Public Const CheckBoxEmpty = &H20
    
'############################## Un1
Public Const DropButtonsOrArrows = &H3
    '######################### Un2
    Public Const DropDown = &H1
    Public Const DropLeft = &H2
    Public Const DropRight = &H3
    Public Const DropUp = &H4
        Public Const ResizeWindowRight = &H8
        Public Const ResizeWindowLeft = &H10
    Public Const DisabledDropUp = &H100
    Public Const DisabledDropDown = &H101
    Public Const DisabledDropLeft = &H102
    Public Const DisabledDropRight = &H103
    Public Const DepresedDropUp = &H200
    Public Const DepresedDropDown = &H201
    Public Const DeprsedDropLeft = &H202
    Public Const DeprsedDropRight = &H203
    Public Const DepresedDisabledDropUp = &H300
    Public Const DeprsedDisabledDropDown = &H301
    Public Const DeprsedDisabledDropLeft = &H302
    Public Const DeprsedDisabledDropRight = &H303
    ' 400 - 503 Semi Disabled States
    
'######################## Un1
Public Const CaptionButtons = &H1
    '#################### Un2
    Public Const Max = &H2
    Public Const Min = &H1
    Public Const Restore = &H3
    Public Const Question = &H4
    Public Const CloseB = &H10
    Public Const DisabledClose = &H100
    Public Const DisabledMin = &H101
    Public Const DisabledMax = &H102
    Public Const DisabledRestore = &H103
    Public Const DisabledQuestion = &H104
    Public Const DeprsedClose = &H200
    Public Const DeprsedMin = &H201
    Public Const DeprsedMax = &H202
    Public Const DeprsedRestore = &H203
    Public Const DeprsedQuestion = &H204
    Public Const SemiGreyClose = &H400
    Public Const SemiGreyMin = &H401
    Public Const SemiGreyMax = &H402
    Public Const SemiGreyRestore = &H403
    Public Const SemiGreyQuestion = &H404
