Attribute VB_Name = "Custom_UI_Callbacks"
Option Explicit
Sub New_Spur_Fully_Taped()
Attribute New_Spur_Fully_Taped.VB_Description = "Inserts code for a new spur that is fully taped"
Attribute New_Spur_Fully_Taped.VB_ProcData.VB_Invoke_Func = "N\n14"

Module1.Start_New_Spur_Fully_Taped

End Sub

Sub New_Spur_Space_Taped()
Attribute New_Spur_Space_Taped.VB_Description = "Inserts Code for a new spur that is space taped"
Attribute New_Spur_Space_Taped.VB_ProcData.VB_Invoke_Func = "n\n14"

Module1.Start_New_Spur_Space_Taped

End Sub

Sub Taped_Length()
Attribute Taped_Length.VB_Description = "Inserts code for a taped length that is fully taped."
Attribute Taped_Length.VB_ProcData.VB_Invoke_Func = "T\n14"

Module1.Insert_Taped_Length

End Sub

Sub Space_Taped_Length()
Attribute Space_Taped_Length.VB_Description = "Inserts code for a taped length that is space taped"
Attribute Space_Taped_Length.VB_ProcData.VB_Invoke_Func = "t\n14"

Module1.Insert_Space_Taped_Length

End Sub

Sub Initial_Spur_Fully_Taped()
Attribute Initial_Spur_Fully_Taped.VB_Description = "Inserts default code for the first spur of a harness, fully taped."
Attribute Initial_Spur_Fully_Taped.VB_ProcData.VB_Invoke_Func = "I\n14"

Module1.Insert_Initial_Spur_Fully_Taped

End Sub

Sub Initial_Spur_Space_Taped()
Attribute Initial_Spur_Space_Taped.VB_Description = "Inserts default code for the first spur in a harness, space taped."
Attribute Initial_Spur_Space_Taped.VB_ProcData.VB_Invoke_Func = "i\n14"

Module1.Insert_Initial_Spur_Space_Taped

End Sub

Sub feed_without_wind()
Attribute feed_without_wind.VB_Description = "Inserts Code for feed without wind length"
Attribute feed_without_wind.VB_ProcData.VB_Invoke_Func = "F\n14"

Module1.Insert_Feed_Without_Wind

End Sub

Sub adapter()
Attribute adapter.VB_Description = "Inserts the Adapter line of code for use on lengths under 300mm"
Attribute adapter.VB_ProcData.VB_Invoke_Func = "A\n14"

Module1.Insert_Adapter_Line

End Sub

Sub Final_Cut()
Attribute Final_Cut.VB_Description = "Inserts Code for the last cut of a harness.  To be used at the end of every program"
Attribute Final_Cut.VB_ProcData.VB_Invoke_Func = "f\n14"

Module1.Insert_Final_Cut

End Sub

Sub Export_To_CLC()
Attribute Export_To_CLC.VB_Description = "Exports the File as a .clc file for the C16 machines"
Attribute Export_To_CLC.VB_ProcData.VB_Invoke_Func = "E\n14"

Module1.Export_As_CLC

End Sub

Sub Import_From_CLC()
Attribute Import_From_CLC.VB_Description = "Opens the .clc files into the excel spreadsheet"
Attribute Import_From_CLC.VB_ProcData.VB_Invoke_Func = "O\n14"

Module1.Import_CLC_To_Excel

End Sub

Sub Clear()
Attribute Clear.VB_Description = "Clears selected cells of text and formatting"
Attribute Clear.VB_ProcData.VB_Invoke_Func = "X\n14"

Module1.Clear_Cells_of_Formatting

End Sub

'Sub Clear()
'
'Module1.Clear_Cells_of_Formatting
'
'End Sub



