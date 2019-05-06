Attribute VB_Name = "modFormIcon"
' // modFormIcon.bas - set icon to form
' // © Krivous Anatoly Anatolevich (The trick), 2016

Option Explicit

Public Function SetIconToForm( _
                ByVal hWnd As Long, _
                ByRef ResId As Variant) As Boolean
    Dim hIcon   As Long:    Dim dat()   As Byte
    Dim index   As Long:    Dim Count   As Long
    Dim width   As Long:    Dim height  As Long
    Dim cx      As Long:    Dim cy      As Long
    Dim thr     As Long:    Dim min     As Long
    Dim fndIdx  As Long:    Dim size    As Long
    Dim offset  As Long
    
    ' // Extract prefer icon size
    cx = GetSystemMetrics(SM_CXSMICON)
    cy = GetSystemMetrics(SM_CYSMICON)
    
    dat() = LoadResData(ResId, "CUSTOM")
    
    ' // Get number of images in the icon
    GetMem4 dat(4), Count: Count = Count And &HFFFF&
    fndIdx = -1
    min = &H7FFFFFFF
    
    For index = 0 To Count - 1
        
        ' // Get offset to data
        GetMem4 dat(16& * index + 18&), offset
        GetMem4 dat(offset + 4), width
        GetMem4 dat(offset + 8), height
        
        height = height \ 2
        
        If width = cx And height = cy Then
            fndIdx = index
        Else
            thr = Abs(cx - width) + Abs(cy - height)
            If thr < min Then min = thr: fndIdx = index
        End If
        
    Next
    
    GetMem4 dat(16& * fndIdx + 18&), offset
    GetMem4 dat(16& * fndIdx + 14&), size
    
    hIcon = CreateIconFromResourceEx(dat(offset), size, 1, ICRESVER, 0&, 0&, LR_DEFAULTSIZE)
    
    If hIcon Then
        SendMessage hWnd, WM_SETICON, ICON_SMALL, ByVal hIcon
    End If
    
End Function

