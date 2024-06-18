Attribute VB_Name = "Module9"
Public Sub AutoUpdate()
    filePath = "https://solaredge0.sharepoint.com/:x:/r/sites/SystemProductionTeam/Shared%20Documents/System%20Production%20Team/Z_DB/DB/_queriesdatabase.xlsx?d=wce952650a9d34baf8486c1f1ccf29755&csf=1&web=1&e=ktvtJF"
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
        .EnableEvents = False
        .AskToUpdateLinks = False
    End With
    Workbooks.Open filePath
    ActiveWorkbook.UpdateLink Name:=ActiveWorkbook.LinkSources
    ActiveWorkbook.Close True
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
        .EnableEvents = True
        .AskToUpdateLinks = True
    End With
End Sub
