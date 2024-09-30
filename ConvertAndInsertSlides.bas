Attribute VB_Name = "ConvertAndInsertSlides"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function StrCmpLogicalW Lib "shlwapi.dll" (ByVal psz1 As LongPtr, ByVal psz2 As LongPtr) As Long
#Else
    Private Declare Function StrCmpLogicalW Lib "shlwapi.dll" (ByVal psz1 As Long, ByVal psz2 As Long) As Long
#End If

Sub ConvertAndInsertSlides()
    Dim fileDialog As fileDialog
    Dim filePath As String
    Dim fileExtension As String
    Dim tempFolder As String
    Dim imgFolder As String
    Dim imgFiles() As String
    Dim imgFile As String
    Dim imgCount As Long
    Dim sizeReduction As Single
    Dim deleteFolderResponse As VbMsgBoxResult

    ' Initialize FileDialog
    Set fileDialog = Application.fileDialog(msoFileDialogOpen)
    With fileDialog
        .AllowMultiSelect = False
        .Title = "Select a PowerPoint or PDF file"
        .Filters.Clear
        .Filters.Add "PowerPoint and PDF Files", "*.ppt; *.pptx; *.pdf"
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With

    ' Determine file extension
    fileExtension = LCase(Mid(filePath, InStrRev(filePath, ".") + 1))

    ' Create temporary folders
    tempFolder = Environ("USERPROFILE") & "\Downloads\ConvertedImages_" & Format(Now, "yyyymmdd_hhmmss")
    MkDir tempFolder
    imgFolder = tempFolder & "\Images"
    MkDir imgFolder

    ' Convert file to images
    If fileExtension = "pptx" Or fileExtension = "ppt" Then
        Call ConvertPPTtoImages(filePath, imgFolder)
    ElseIf fileExtension = "pdf" Then
        Call ConvertPDFtoImages(filePath, imgFolder)
    Else
        MsgBox "Unsupported file type."
        Exit Sub
    End If

    ' Show the UserForm to select image size reduction
    frmImageSizeReduction.Show

    ' Determine the size reduction percentage
    If frmImageSizeReduction.opt1.Value = True Then
        sizeReduction = 0.01
    ElseIf frmImageSizeReduction.opt5.Value = True Then
        sizeReduction = 0.05
    ElseIf frmImageSizeReduction.opt10.Value = True Then
        sizeReduction = 0.1
    ElseIf frmImageSizeReduction.opt20.Value = True Then
        sizeReduction = 0.2
    ElseIf frmImageSizeReduction.opt25.Value = True Then
        sizeReduction = 0.25
    ElseIf frmImageSizeReduction.opt50.Value = True Then
        sizeReduction = 0.5
    Else
        sizeReduction = 0 ' Default to no reduction
    End If

    ' Unload the form
    Unload frmImageSizeReduction

    ' Read image files into an array
    imgCount = 0
    imgFile = Dir(imgFolder & "\*.png")
    Do While imgFile <> ""
        imgCount = imgCount + 1
        ReDim Preserve imgFiles(1 To imgCount)
        imgFiles(imgCount) = imgFolder & "\" & imgFile
        imgFile = Dir
    Loop

    ' Sort the image files using natural sort order
    If imgCount > 1 Then
        Call QuickSort(imgFiles, 1, imgCount)
    End If

    ' Insert sorted images
    Dim i As Long
    For i = 1 To imgCount
        Dim inlineShape As inlineShape
        Set inlineShape = Selection.InlineShapes.AddPicture(fileName:=imgFiles(i), LinkToFile:=False, SaveWithDocument:=True)

        ' Adjust the image size by the selected percentage
        If sizeReduction > 0 Then
            With inlineShape
                .Width = .Width * (1 - sizeReduction)
                .Height = .Height * (1 - sizeReduction)
            End With
        End If

        ' Move to the next line
        Selection.TypeParagraph
        Selection.TypeParagraph
    Next i

    MsgBox "Images have been inserted successfully."

    ' Ask the user whether to delete the temporary image folder
    deleteFolderResponse = MsgBox("Do you want to delete the temporary image folder?", vbYesNo + vbQuestion, "Delete Temporary Folder")

    If deleteFolderResponse = vbYes Then
        On Error Resume Next
        ' Delete the folder and its contents
        Call DeleteFolder(tempFolder)
        On Error GoTo 0
    End If
End Sub

' QuickSort algorithm using natural sort order
Sub QuickSort(arr() As String, low As Long, high As Long)
    Dim pivot As String
    Dim i As Long
    Dim j As Long
    Dim temp As String

    If low < high Then
        pivot = arr((low + high) \ 2)
        i = low
        j = high

        Do While i <= j
            Do While NaturalCompare(arr(i), pivot) < 0
                i = i + 1
            Loop
            Do While NaturalCompare(arr(j), pivot) > 0
                j = j - 1
            Loop
            If i <= j Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
                i = i + 1
                j = j - 1
            End If
        Loop

        If low < j Then QuickSort arr, low, j
        If i < high Then QuickSort arr, i, high
    End If
End Sub

' Function to compare strings using natural sort order
Function NaturalCompare(s1 As String, s2 As String) As Long
    NaturalCompare = StrCmpLogicalW(StrPtr(s1), StrPtr(s2))
End Function

' UserForm code for image size reduction
' (Assuming you have a UserForm named 'frmImageSizeReduction' with option buttons named 'opt1', 'opt5', etc.)

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    ' Unload the form and exit the macro
    Unload Me
    End
End Sub

