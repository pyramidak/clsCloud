#Region " Cloud "

Public Enum Cloud
    Documents = 0
    OneDrive = 1
    DropBox = 2
    GoogleDisk = 3
    Sync = 4
End Enum

Public Class clsCloud
    Public DropBoxExist As Boolean
    Public DropBoxFolder As String = ""
    Public GoogleDriveExist As Boolean
    Public GoogleDriveFolder As String = ""
    Public OneDriveExist As Boolean
    Public OneDriveFolder As String = ""
    Public SyncExist As Boolean
    Public SyncFolder As String = ""

    Sub New()
        CheckClouds()
    End Sub

    Public Sub CheckClouds()
        'DropBox
        Dim dbPath As String = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Dropbox\host.db")
        If myFile.Exist(dbPath) = False Then dbPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Dropbox\host.db")
        If myFile.Exist(dbPath) Then
            Try
                Dim lines As String() = System.IO.File.ReadAllLines(dbPath)
                Dim dbBase64Text As Byte() = Convert.FromBase64String(lines(1))
                Dim dFolder As String = System.Text.ASCIIEncoding.ASCII.GetString(dbBase64Text)
                DropBoxExist = myFolder.Exist(dFolder)
                If DropBoxExist Then DropBoxFolder = dFolder
            Catch ex As Exception
                DropBoxExist = False
            End Try
        End If

        'GoogleDrive
        dbPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\Google\Drive\user_default\sync_config.db"
        If myFile.Exist(dbPath) = False Then dbPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\Google\Drive\sync_config.db"
        Try
            Dim gFolder As String = GetFolderFromFile(dbPath)
            GoogleDriveExist = myFolder.Exist(gFolder)
            If GoogleDriveExist Then GoogleDriveFolder = gFolder
        Catch ex As Exception
            GoogleDriveExist = False
        End Try

        'OneDrive
        Dim sFolder As String = myRegister.GetValue(HKEY.CURRENT_USER, "Software\Microsoft\OneDrive", "UserFolder", "")
        If sFolder = "" Then myRegister.GetValue(HKEY.CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\SkyDrive", "UserFolder", "")
        If sFolder = "" Then myRegister.GetValue(HKEY.CURRENT_USER, "Software\Microsoft\SkyDrive", "UserFolder", "")
        If sFolder = "" Then myRegister.GetValue(HKEY.CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\OneDrive", "UserFolder", "")
        OneDriveExist = myFolder.Exist(sFolder)
        If OneDriveExist Then OneDriveFolder = sFolder

        'Sync
        Try
            sFolder = GetFolderFromFile(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\Sync.Config\1.1\cfg.db")
            SyncExist = myFolder.Exist(sFolder)
            If SyncExist Then SyncFolder = sFolder
        Catch ex As Exception
            SyncExist = False
        End Try

    End Sub

    Private Function GetFolderFromFile(dbPath As String) As String
        GetFolderFromFile = ""
        If myFile.Exist(dbPath) Then
            If myFile.Copy(dbPath, dbPath + ".bak") Then
                dbPath = dbPath + ".bak"
                Dim SR As New IO.StreamReader(dbPath)
                Dim sText As String = SR.ReadToEnd
                SR.Dispose()

                Dim disk As String = dbPath.Substring(0, 3)
                Dim iStart As Integer = sText.IndexOf(disk, 1)
                If iStart = -1 Then iStart = sText.IndexOf(disk.Replace("\", "/"), 1)
                If iStart = -1 Then Return ""
                Dim iEnd As Integer = sText.IndexOf("Sync", iStart)
                If iEnd = -1 Then
                    iEnd = sText.IndexOf("Disk Google", iStart)
                    If iEnd = -1 Then
                        iEnd = sText.IndexOf("Google Drive", iStart)
                        If iEnd = -1 Then
                            Return ""
                        Else
                            iEnd += 12
                        End If
                    Else
                        iEnd += 11
                    End If
                Else
                    iEnd += 4
                End If

                GetFolderFromFile = sText.Substring(iStart, iEnd - iStart)
                If GetFolderFromFile.Contains("/") Then GetFolderFromFile = GetFolderFromFile.Replace("/", "\")
            End If
        End If
    End Function

    Public Function NewAppPath(newCloud As Cloud, Optional filename As String = "") As String
        If filename = "" Then filename = "pyramidak\" & Application.ExeName & ".xml"

        Select Case newCloud
            Case Cloud.OneDrive
                If OneDriveExist = False Then newCloud = Cloud.Documents
            Case Cloud.DropBox
                If DropBoxExist = False Then newCloud = Cloud.Documents
            Case Cloud.GoogleDisk
                If GoogleDriveExist = False Then newCloud = Cloud.Documents
            Case Cloud.Sync
                If SyncExist = False Then newCloud = Cloud.Documents
        End Select

        Select Case newCloud
            Case Cloud.OneDrive
                Return myFile.Join(OneDriveFolder, filename)
            Case Cloud.DropBox
                Return myFile.Join(DropBoxFolder, filename)
            Case Cloud.GoogleDisk
                Return myFile.Join(GoogleDriveFolder, filename)
            Case Cloud.Sync
                Return myFile.Join(SyncFolder, filename)
            Case Else
                Return myFile.Join(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), filename)
        End Select
    End Function

End Class

#End Region