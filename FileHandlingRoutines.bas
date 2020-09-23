Attribute VB_Name = "FileHandlingRoutines"
Public totalBytes
Public picpath As String
Public Function GetFileSize(FileName) As String
    On Error GoTo Gfserror
    Dim bytes As String
    
    bytes = FileLen(FileName)

totalBytes = bytes
    If bytes >= "1024" Then
        'KB
        bytes = CCur(bytes / 1024) & " Kb"
    Else


        If bytes >= "1048576" Then
            'MB
            bytes = CCur(bytes / (1024 * 1024)) & " Mb"
        Else
            bytes = CCur(bytes) & " B"
        End If
    End If
    GetFileSize = bytes
    Exit Function
Gfserror:
    GetFileSize = "0 B"
    Resume
End Function
Public Function GetFileExtension(FileName As String)
    On Error Resume Next
    Dim TempStr As String
    TempStr = Right(FileName, 2)


    If Left(TempStr, 1) = "." Then
        GetFileExtension = Right(FileName, 1)
        Exit Function
    Else
        TempStr = Right(FileName, 3)


        If Left(TempStr, 1) = "." Then
            GetFileExtension = Right(FileName, 2)
            Exit Function
        Else
            TempStr = Right(FileName, 4)


            If Left(TempStr, 1) = "." Then
                GetFileExtension = Right(FileName, 3)
                Exit Function
            Else
                TempStr = Right(FileName, 5)


                If Left(TempStr, 1) = "." Then
                    GetFileExtension = Right(FileName, 4)
                    Exit Function
                Else
                    GetFileExtension = "Unknown"
                End If
            End If
        End If
    End If
End Function

