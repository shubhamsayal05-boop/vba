Attribute VB_Name = "Create_ODRIV_PQ"
Option Explicit

' Create Power Query queries from .pq files in a folder.
' Usage:
'   1) Extract this starter kit to a folder.
'   2) In Excel press ALT+F11, Import this .bas.
'   3) Run CreateODRIVQueries and pick the "M" folder.
Public Sub CreateODRIVQueries()
    Dim fso As Object, fld As Object, fil As Object
    Dim dlg As FileDialog
    Dim folderPath As String
    Dim mText As String, qName As String

    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    dlg.Title = "Select folder that contains .pq files"
    If dlg.Show <> -1 Then Exit Sub
    folderPath = dlg.SelectedItems(1)

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fld = fso.GetFolder(folderPath)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    For Each fil In fld.Files
        If LCase(fso.GetExtensionName(fil.Path)) = "pq" Then
            mText = ReadAllText(fil.Path)
            qName = fso.GetBaseName(fil.Path)
            On Error Resume Next
            ' Delete existing query with same name
            wb.Queries(qName).Delete
            On Error GoTo 0
            wb.Queries.Add Name:=qName, Formula:=mText
        End If
    Next fil

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Power Query definitions created.", vbInformation
End Sub

Private Function ReadAllText(ByVal filePath As String) As String
    Dim f As Integer, s As String
    f = FreeFile
    Open filePath For Input As #f
    s = Input$(LOF(f), f)
    Close #f
    ReadAllText = s
End Function
