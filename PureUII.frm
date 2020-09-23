VERSION 5.00
Begin VB.Form PureUII 
   Caption         =   "UserInterfaceExample"
   ClientHeight    =   435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   ScaleHeight     =   435
   ScaleWidth      =   3360
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "PureUII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Paint()
Dim R As RECT
    'Clear the form
    Me.Cls
    'API uses pixels
    Me.ScaleMode = vbPixels
    SetRect R, 0, 0, 15, 15
    DrawFrameControl Me.hdc, R, CaptionButtons, Min
    OffsetRect R, 15, 0
    DrawFrameControl Me.hdc, R, CaptionButtons, Max
    OffsetRect R, 15, 0
    DrawFrameControl Me.hdc, R, CaptionButtons, Restore
    OffsetRect R, 15, 0
    DrawFrameControl Me.hdc, R, CaptionButtons, Question
     OffsetRect R, 15, 0
    DrawFrameControl Me.hdc, R, CaptionButtons, CloseB
     OffsetRect R, 15, 0
    DrawFrameControl Me.hdc, R, CaptionButtons, DisabledClose
     OffsetRect R, 15, 0
    DrawFrameControl Me.hdc, R, CaptionButtons, DisabledMin
     OffsetRect R, 15, 0
    DrawFrameControl Me.hdc, R, CaptionButtons, DisabledMax
     OffsetRect R, 15, 0
    DrawFrameControl Me.hdc, R, CaptionButtons, DisabledRestore
     OffsetRect R, 15, 0
    DrawFrameControl Me.hdc, R, CaptionButtons, DisabledQuestion
     OffsetRect R, 15, 0
    DrawFrameControl Me.hdc, R, CaptionButtons, DeprsedClose
     OffsetRect R, 15, 0
    DrawFrameControl Me.hdc, R, CaptionButtons, DeprsedMin
     OffsetRect R, 15, 0
    DrawFrameControl Me.hdc, R, CaptionButtons, DeprsedMax
     OffsetRect R, 15, 0
    DrawFrameControl Me.hdc, R, CaptionButtons, DeprsedRestore
     OffsetRect R, 15, 0
    DrawFrameControl Me.hdc, R, CaptionButtons, DeprsedQuestion
    SetRect R, 0, 0, 15, 15
    OffsetRect R, 0, 15
    DrawFrameControl Me.hdc, R, CaptionButtons, SemiGreyClose
    OffsetRect R, 15, 0
    DrawFrameControl Me.hdc, R, CaptionButtons, SemiGreyMin
    OffsetRect R, 15, 0
    DrawFrameControl Me.hdc, R, CaptionButtons, SemiGreyMax
    OffsetRect R, 15, 0
    DrawFrameControl Me.hdc, R, CaptionButtons, SemiGreyRestore
    OffsetRect R, 15, 0
    DrawFrameControl Me.hdc, R, CaptionButtons, SemiGreyQuestion
End Sub
