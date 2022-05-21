Attribute VB_Name = "ModProgress"
'           .---.        .-----------
'          /     \  __  /    ------
'         / /     \(..)/    -----
'        //////   ' \/ `   ---
'       //// / // :    : ---
'      // /   /  /`    '--
'     // /        //..\\
'   o===|========UU====UU=====-  -==========================o
'                '//||\\`
'                       DEVELOPED BY JGH
'
'   -=====================|===o  o===|======================-
Option Explicit

Public Sub UpdateP1(ByVal str_msg, ByVal myvalue As String, ByVal mylimit As Variant)

Dim myprogress As Variant

FormProgress.L_Progress_1 = str_msg

myprogress = myvalue / mylimit

FormProgress.Progress_1.Width = myprogress * 306

End Sub

Public Sub UpdateP2(ByVal str_msg, ByVal myvalue As String, ByVal mylimit As Variant)

Dim myprogress As Variant

FormProgress.L_Progress_2 = str_msg

myprogress = myvalue / mylimit

FormProgress.Progress_2.Width = myprogress * 306

End Sub

Public Sub UpdateP3(ByVal str_msg, ByVal myvalue As String, ByVal mylimit As Variant)

Dim myprogress As Variant

FormProgress.L_Progress_3 = str_msg

myprogress = myvalue / mylimit

FormProgress.Progress_3.Width = myprogress * 306

End Sub
