Sub Macro1()

    
    myfolder = ActiveDocument.FilePath
    
    Dim currTime As Date
    currTime = Time()
    
    Dim xDim(0 To 5) As Integer
    xDim(0) = 48
    xDim(1) = 72
    xDim(2) = 96
    xDim(3) = 144
    xDim(4) = 196
    xDim(5) = 512
    Dim xName(0 To 5) As String
    xName(0) = "mdpi"
    xName(1) = "hdpi"
    xName(2) = "xhdpi"
    xName(3) = "xxhdpi"
    xName(4) = "xxxhdpi"
    xName(5) = "512"   
     
    Dim xFirst As Integer
    Dim xLast As Integer
    xFirst = LBound(xDim)
    xLast = UBound(xDim)
    
    Dim qPages
    qPages = ActiveDocument.Pages.Count
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Dim finalFolder As String
    For x = xFirst To xLast
        finalFolder = myfolder & xName(x)
        If fs.FolderExists(finalFolder) = False Then
            fs.createfolder (finalFolder)
        End If
    Next
    
    Dim thisPage
    Dim expOpts As New StructExportOptions
    expOpts.ImageType = cdrCMYKColorImage
    expOpts.AntiAliasingType = cdrNormalAntiAliasing
    expOpts.ResolutionX = 72
    expOpts.ResolutionY = 72
    
    Dim thisPng As String

    For i = 0 To qPages
        thisPage = ActiveDocument.Pages.Item(i).Name
        
         If thisPage = "rounded" Or thisPage = "square" Then
            ActiveDocument.Pages(i).Activate
            For x = xFirst To xLast
                expOpts.SizeX = xDim(x)
                expOpts.SizeY = xDim(x)
                If thisPage = "square" Then
                    thisPng = myfolder & "\" & xName(x) & "\ic_launcher.png"
                Else
                    thisPng = myfolder & "\" & xName(x) & "\ic_launcher_round.png"
                End If
                ActiveDocument.Export thisPng, cdrPNG, cdrCurrentPage, expOpts
            Next
        End If   
    Next
    
    Dim msgFinal As String
    msgFinal = "fin del proceso: " & currTime
    Debug.Print (msgFinal)
    MsgBox (msgFinal)
    
End Sub
