Attribute VB_Name = "Folder"
'-------------------------------------------------------------------------------------------------
' Copyright (C) 2002-2012 Hiroaki Inaba
'
' Rev 1.4
' - Support two folder name formats by switching "Folder_Format" document property.
'-------------------------------------------------------------------------------------------------

' Captions for shortcut menu
Const CAPTION_OPEN_FOLDER As String = "Open Folder"
Const CAPTION_ADJUST_ROW_HIGHT As String = "Adjust Height"
Const CAPTION_TOGGLE_FOLDER As String = "Toggle FormulaBar"


'----------------------------------------------------------------------------------------------------
' Public Functions
'----------------------------------------------------------------------------------------------------

' Register shortcut menus
Sub Folder_Add_Menues()
    Application.CommandBars("Cell").Controls.Item(1).BeginGroup = True
    
    Dim NewMenu As Variant
    Set NewMenu = Application.CommandBars("Cell").Controls.Add(Type:=msoControlButton, Before:=1)
    With NewMenu
        .caption = CAPTION_OPEN_FOLDER
        .OnAction = "Open_Folder"
        .FaceId = 7
    End With
    Set NewMenu = Application.CommandBars("Cell").Controls.Add(Type:=msoControlButton, Before:=2)
    With NewMenu
        .caption = CAPTION_ADJUST_ROW_HIGHT
        .OnAction = "Adjust_Raw_Height"
        .FaceId = 16
    End With
    
    
    Set NewMenu = Application.CommandBars("Cell").Controls.Add(Type:=msoControlButton, Before:=3)
    With NewMenu
        .caption = CAPTION_TOGGLE_FOLDER
        .OnAction = "Toggle_FormulaBar"
        .FaceId = 25
    End With
    Application.CommandBars("Cell").Controls.Item(5).BeginGroup = True
    
End Sub

' Unregister shortcuts menus
Sub Folder_Remove_Menues()
    Remove_MenuItem CAPTION_OPEN_FOLDER
    Remove_MenuItem CAPTION_ADJUST_ROW_HIGHT
    Remove_MenuItem CAPTION_TOGGLE_FOLDER
    Remove_MenuItem "Adjust Raw Height"
End Sub



'----------------------------------------------------------------------------------------------------
' Helper Functions
'----------------------------------------------------------------------------------------------------

' Remove a menu that specified the argument
Private Sub Remove_MenuItem(caption As String)
    Dim m As Variant
    On Error GoTo on_error
    While 1
        Set m = Application.CommandBars("Cell").Controls(caption)
        m.Delete
    Wend

on_error:
End Sub

Private Sub Toggle_FormulaBar()
    Application.DisplayFormulaBar = Not Application.DisplayFormulaBar
End Sub

Private Sub Adjust_Raw_Height()
    Dim old_address As String, a_row As Range, new_height As Integer
    
    ' stop updating screen to speed up
    Application.ScreenUpdating = False
    old_address = ActiveCell.Address

    Cells.EntireRow.AutoFit
    
    For Each a_row In ActiveSheet.UsedRange.Rows
        new_height = a_row.RowHeight + 10
        ' clip to 409
        If new_height > 409 Then
            new_height = 409
        End If
        a_row.RowHeight = new_height
    Next
    
    ' store cell position
    Range(old_address).Activate
    Application.ScreenUpdating = True
End Sub

Private Sub Open_Folder()
    Dim dir_path As String, cmd_line As String
    
    dir_path = get_dir()
    If dir_path = "" Then
        Exit Sub
    End If
    
    cmd_line = "explorer.exe  /n," & "" & dir_path & ""
    Shell cmd_line, vbNormalNoFocus
End Sub


'----------------------------------------------------------------------------------------------------
' Helper Functions for folder operation
'----------------------------------------------------------------------------------------------------

Private Function get_folder_format() As String
    On Error GoTo No_Custom_Property
    get_folder_format = ActiveWorkbook.CustomDocumentProperties("Folder_Format")
    Exit Function
    
No_Custom_Property:
    get_folder_format = "2"
End Function

' Determine folder name from current ActiveCell.  Will create the folder if it is not exist yet.
Private Function get_dir() As String
    On Error GoTo on_error
    Dim sheet As String, v1 As String, v2 As String
    Dim dir_name As String, dir_path As String
    Dim the_row As Integer
    Dim base_dir As String

    sheet = ActiveSheet.Name
    the_row = ActiveCell.Row
    
    ' If a document contains "Folder_Format" property with "1", folder name will be Column(1).
    ' Oterwise folder name will be "Column(1) - Column(2)".
    If get_folder_format() = "1" Then
        dir_name = translate_to_folder_name(Cells(the_row, 1).Value)
    Else
        dir_name = translate_to_folder_name(Cells(the_row, 1).Value & " - " & Cells(the_row, 2).Value)
    End If
    
    If dir_name = " - " Then
        MsgBox "Invalid line was selected", vbCritical
        get_dir = ""
        Exit Function
    End If
            
    ' dir_path is full path for the target directory
    base_dir = ActiveWorkbook.Path & "\" & sheet
    dir_path = base_dir & "\" & dir_name
    
    ' create base dir and the target dir if needed
    If Dir(base_dir, vbDirectory) = "" Then
        MkDir base_dir
    End If
    If Dir(dir_path, vbDirectory) = "" Then
        MkDir dir_path
    End If
    
    get_dir = dir_path
    Exit Function
    
on_error:
    MsgBox "Failed to create a folder" & dir_path, vbCritical

End Function

' repalce characters to underscore that can not be used as a file/dir name
Private Function translate_to_folder_name(target_string) As String
    Dim target As String
    target = Replace(target_string, "\", "_")
    target = Replace(target, "/", "_")
    target = Replace(target, ":", "_")
    target = Replace(target, "*", "_")
    target = Replace(target, "?", "_")
    target = Replace(target, """", "_")
    target = Replace(target, "<", "_")
    target = Replace(target, ">", "_")
    target = Replace(target, "|", "_")
    target = Replace(target, ",", "_")
    translate_to_folder_name = target
End Function

