Attribute VB_Name = "DisplayChanges"

' MODULE NAME: DisplayChanges.BAS
' ===============================
'
' Module for display modes management, hide/show cursor...etc

Option Explicit

Global DispMode As DEVMODE
Sub CursorOFF()

 ' SUB : CursorOFF
 ' ===============
 '
 ' RETURNED VALUES: None
 '
 ' Hide the cursor.

 ShowCursor 0

End Sub
Sub CursorON()

 ' SUB : CursorON
 ' ==============
 '
 ' RETURNED VALUES: None
 '
 ' Show the cursor.

 ShowCursor 1

End Sub
Sub RestoreDisplayMode()

 ' SUB : RestoreDisplayMode
 ' ========================
 '
 ' RETURNED VALUES: None
 '
 ' Restore the old display mode.

 ChangeDisplaySettings DispMode, 0

End Sub
Private Sub RememberDisplayMode()

 ' SUB : RememberDisplayMode (private)
 ' ==================================
 '
 ' RETURNED VALUES: None
 '
 ' Remember the starting display mode settings,
 '  and store thems in a temporary storage.

 EnumDisplaySettings 0&, 0&, DispMode

 With DispMode
  .dmBitsPerPel = GetDeviceCaps(0, 12)
  .dmDisplayFrequency = GetDeviceCaps(0, 116)
  .dmPelsWidth = GetDeviceCaps(0, 8)
  .dmPelsHeight = GetDeviceCaps(0, 10)
 End With

End Sub
Sub ChangeDisplayMode(W&, H&, BPP&)

 ' SUB : ChangeDisplayMode
 ' =======================
 '
 ' RETURNED VALUES: None
 '
 ' Change the display mode with a supported specified
 '  width, height, and the BitPerPixel (BPP).
 '   The refresh rate (frequancy) is set's as default.

 RememberDisplayMode

 Dim DM As DEVMODE
 DM = DispMode

 With DM
  .dmPelsWidth = W
  .dmPelsHeight = H
  .dmBitsPerPel = BPP
 End With

 ChangeDisplaySettings DM, 0

End Sub
