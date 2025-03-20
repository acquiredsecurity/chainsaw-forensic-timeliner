Sub ColorRowsByArtifactName()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim colorMap As Object
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Find the last used row in column B
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    ' Create a dictionary for artifact name color mapping
    Set colorMap = CreateObject("Scripting.Dictionary")
    
    ' Assign colors to artifact names
    colorMap.Add "account_tampering", RGB(0, 0, 255) ' Blue
    colorMap.Add "antivirus", RGB(0, 128, 0) ' Green
    colorMap.Add "indicator_removal", RGB(255, 0, 0) ' Red
    colorMap.Add "lateral_movement", RGB(255, 165, 0) ' Orange
    colorMap.Add "login_attacks", RGB(255, 255, 0) ' Yellow
    colorMap.Add "MFT - FileNameCreated0x30", RGB(128, 0, 128) ' Purple
    colorMap.Add "microsoft_rds_events_-_user_profile_disk", RGB(0, 255, 255) ' Cyan
    colorMap.Add "persistence", RGB(128, 128, 0) ' Olive
    colorMap.Add "powershell_engine_state", RGB(255, 192, 203) ' Pink
    colorMap.Add "powershell_script", RGB(165, 42, 42) ' Brown
    colorMap.Add "rdp_events", RGB(0, 255, 0) ' Lime
    colorMap.Add "service_installation", RGB(0, 128, 128) ' Teal
    colorMap.Add "sigma", RGB(153, 50, 204) ' Dark Orchid (Purple)

    ' Loop through column B and apply colors
    For Each cell In ws.Range("B2:B" & lastRow)
        If colorMap.Exists(cell.Value) Then
            cell.EntireRow.Interior.Color = colorMap(cell.Value)
        End If
    Next cell

    ' Cleanup
    Set colorMap = Nothing
    MsgBox "Row coloring complete!", vbInformation
End Sub
