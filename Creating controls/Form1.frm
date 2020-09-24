VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Using Controls without putting them on the form"
   ClientHeight    =   3105
   ClientLeft      =   720
   ClientTop       =   435
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' By James/Webmonster
' Contact: James@Websofts.org
' Or my Msn: Webmonster123@hotmail.com
' For more source codes like this visit www.WebSofts.org

Option Explicit

' Tmr1 = Timer1
' Tmr2 = Timer2
' Cmd1 = CommandButton1
' Lbl1 = Label1

Private WithEvents Tmr1 As Timer ' Defines variable
Attribute Tmr1.VB_VarHelpID = -1
Private WithEvents Tmr2 As Timer ' Defines variable
Attribute Tmr2.VB_VarHelpID = -1
Private WithEvents Cmd1 As CommandButton ' Defines variable
Attribute Cmd1.VB_VarHelpID = -1
Private WithEvents Lbl1 As Label ' Defines variable
Attribute Lbl1.VB_VarHelpID = -1

Private Sub Cmd1_Click()
'On Error Resume Next

If Cmd1.Caption = "Start!" Then
    
    MsgBox "Tm1 and 2 are now enabled! Lbl1 Will now blink red and black.", vbInformation, "Enabled"
    
    With Tmr1 ' Setting properties for Tmr1
        
        .Interval = 250
        .Enabled = True
        
    End With
    
    With Tmr2 ' Setting properties for Tmr2
    
        .Interval = 300
        .Enabled = True
        
    End With
    
    Cmd1.Caption = "Stop!"
    
    Else
    
    With Tmr1
    
    .Enabled = False
    
    End With
    
    With Tmr2
    
    .Enabled = False
        
    End With
    
    Cmd1.Caption = "Start!"
    
    MsgBox "Tm1 and 2 are now disabled! Lbl1 has stopped blinking red and black.", vbInformation, "Disabled"

    End If

End Sub

Private Sub Form_Load()
    Set Cmd1 = Me.Controls.Add("VB.commandbutton", "Cmd1") ' Defines variable
    Set Lbl1 = Me.Controls.Add("VB.label", "Lbl1") ' Defines variable
    Set Tmr1 = Me.Controls.Add("VB.Timer", "Tmr1") ' Defines variable
    Set Tmr2 = Me.Controls.Add("VB.Timer", "Tmr2") ' Defines variable
    
    With Cmd1 ' Setting properties for Cmd1
    
    .Width = 2655
    .Height = 735
    .Left = 1680
    .Top = 1080
    .Visible = True
    .Caption = "Start!"
    
    End With
    
    With Lbl1 ' Setting properties for Tmr2
    
    .Caption = "This program will show you how to use controls witout putting them on the form (creating controls during runtime)"
    .Height = 495
    .Left = 120
    .Top = 120
    .Width = 5895
    .Visible = True
    
    End With
End Sub

Private Sub Tmr1_Timer()

    With Lbl1
    
    .ForeColor = vbRed
    
    End With
End Sub

Private Sub Tmr2_Timer()

    With Lbl1

    .ForeColor = vbBlack
    
    End With
End Sub
