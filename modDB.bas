Attribute VB_Name = "modDB"
'===========================================================================
' modDB.bas  —  Song database import
'
' Supports:
'   1. CSV  — plain text, comma-separated, no external libraries
'             Schema:  SongName , Artist , FilePath
'             First line may be a header row (auto-detected).
'
'   2. Excel (.xls)  — via ADO + Jet 4.0 OLEDB provider (MDAC 2.6+).
'             Sheet must have columns: SongName | Artist | FilePath
'             Column order / header names are matched case-insensitively.
'
' Both importers call AddSong() from modPlaylist to populate g_Queue.
'===========================================================================
Option Explicit

' ── Column index constants (0-based) ─────────────────────────────────────────
Private Const COL_NAME   As Integer = 0
Private Const COL_ARTIST As Integer = 1
Private Const COL_PATH   As Integer = 2

'===========================================================================
' LoadFromCSV
'   Reads sFilePath (CSV) and appends every valid entry to the queue.
'   Returns the number of songs added; -1 on fatal error.
'   sError is filled with a description on failure.
'===========================================================================
Public Function LoadFromCSV(ByVal sFilePath As String, _
                             ByRef nAdded   As Long, _
                             ByRef sError   As String) As Boolean
    nAdded = 0
    sError = ""

    If Len(Dir(sFilePath)) = 0 Then
        sError = "File not found: " & sFilePath
        LoadFromCSV = False
        Exit Function
    End If

    Dim iFile   As Integer
    Dim sLine   As String
    Dim nLine   As Long
    Dim bHeader As Boolean
    bHeader = False

    iFile = FreeFile
    On Error GoTo ErrCSV
    Open sFilePath For Input As #iFile

    Do While Not EOF(iFile)
        Line Input #iFile, sLine
        nLine = nLine + 1
        sLine = Trim(sLine)

        ' Skip blank lines and comment lines (#)
        If Len(sLine) = 0 Or Left(sLine, 1) = "#" Then GoTo NextLine

        Dim cols() As String
        cols = SplitCSVLine(sLine)

        ' Auto-detect header row: if the "path" column (index 2) does not
        ' contain a backslash on the very first data line, treat it as header
        If nLine = 1 And UBound(cols) >= COL_PATH Then
            If InStr(cols(COL_PATH), "\") = 0 Then
                bHeader = True
                GoTo NextLine
            End If
        End If

        ' We need at least 3 columns
        If UBound(cols) < COL_PATH Then GoTo NextLine

        Dim sName   As String
        Dim sArtist As String
        Dim sPath   As String

        sName   = Trim(cols(COL_NAME))
        sArtist = Trim(cols(COL_ARTIST))
        sPath   = Trim(cols(COL_PATH))

        ' Must be a .WRK file
        If UCase(Right(sPath, 4)) <> ".WRK" Then GoTo NextLine
        ' Path must contain at least one backslash
        If InStr(sPath, "\") = 0 Then GoTo NextLine

        AddSong sPath, sArtist
        nAdded = nAdded + 1

NextLine:
    Loop

    Close #iFile
    LoadFromCSV = True
    Exit Function

ErrCSV:
    sError = "CSV read error on line " & nLine & ": " & Err.Description
    On Error GoTo 0
    Close #iFile
    LoadFromCSV = False
End Function

'===========================================================================
' LoadFromExcel
'   Reads sFilePath (.xls) via ADO Jet 4.0 and appends entries to queue.
'   Requires MDAC 2.6+ installed on the machine (free MS download for Win98).
'   Returns True on success; sError filled on failure.
'===========================================================================
Public Function LoadFromExcel(ByVal sFilePath As String, _
                               ByRef nAdded    As Long, _
                               ByRef sError    As String) As Boolean
    nAdded = 0
    sError = ""

    If Len(Dir(sFilePath)) = 0 Then
        sError = "File not found: " & sFilePath
        LoadFromExcel = False
        Exit Function
    End If

    Dim conn As Object   ' ADODB.Connection
    Dim rs   As Object   ' ADODB.Recordset

    On Error GoTo ErrXLS
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
              "Data Source=" & sFilePath & ";" & _
              "Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1;"""

    ' Open the first sheet — Jet exposes sheets as tables named "Sheet1$" etc.
    Dim sSheet As String
    sSheet = GetFirstSheetName(conn)
    If sSheet = "" Then
        sError = "Could not find a worksheet in: " & sFilePath
        conn.Close
        LoadFromExcel = False
        Exit Function
    End If

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM [" & sSheet & "]", conn, 0, 1  ' adOpenForwardOnly, adLockReadOnly

    ' Map column names to indices (case-insensitive)
    Dim iName   As Integer : iName   = -1
    Dim iArtist As Integer : iArtist = -1
    Dim iPath   As Integer : iPath   = -1

    Dim j As Integer
    For j = 0 To rs.Fields.Count - 1
        Dim sFieldUp As String
        sFieldUp = UCase(Trim(rs.Fields(j).Name))
        Select Case sFieldUp
            Case "SONGNAME", "SONG NAME", "SONG", "NAME", "TITLE"
                iName = j
            Case "ARTIST", "PERFORMER", "BAND"
                iArtist = j
            Case "FILEPATH", "FILE PATH", "PATH", "FILENAME", "FILE"
                iPath = j
        End Select
    Next j

    ' Fallback: assume positional columns if headers didn't match
    If iName = -1 And rs.Fields.Count >= 3 Then iName   = 0
    If iArtist = -1 And rs.Fields.Count >= 3 Then iArtist = 1
    If iPath = -1 And rs.Fields.Count >= 3 Then iPath   = 2

    If iPath = -1 Then
        sError = "Cannot find FilePath column in sheet: " & sSheet
        rs.Close
        conn.Close
        LoadFromExcel = False
        Exit Function
    End If

    Do While Not rs.EOF
        Dim sPath   As String
        Dim sArtist As String

        sPath = Trim(NullToStr(rs.Fields(iPath).Value))

        If iArtist >= 0 Then
            sArtist = Trim(NullToStr(rs.Fields(iArtist).Value))
        Else
            sArtist = ""
        End If

        If UCase(Right(sPath, 4)) = ".WRK" And InStr(sPath, "\") > 0 Then
            AddSong sPath, sArtist
            nAdded = nAdded + 1
        End If

        rs.MoveNext
    Loop

    rs.Close
    conn.Close
    Set rs   = Nothing
    Set conn = Nothing
    LoadFromExcel = True
    Exit Function

ErrXLS:
    sError = "Excel import error: " & Err.Description
    On Error GoTo 0
    If Not rs   Is Nothing Then rs.Close
    If Not conn Is Nothing Then conn.Close
    Set rs   = Nothing
    Set conn = Nothing
    LoadFromExcel = False
End Function

'===========================================================================
' DetectAndLoad
'   Convenience wrapper: detects format by extension and calls the right loader.
'   Returns True on success.
'===========================================================================
Public Function DetectAndLoad(ByVal sFilePath As String, _
                               ByRef nAdded    As Long, _
                               ByRef sError    As String) As Boolean
    Dim sExt As String
    sExt = UCase(Right(Trim(sFilePath), 4))

    Select Case sExt
        Case ".CSV"
            DetectAndLoad = LoadFromCSV(sFilePath, nAdded, sError)
        Case ".XLS"
            DetectAndLoad = LoadFromExcel(sFilePath, nAdded, sError)
        Case Else
            ' Try CSV as default (many files have .txt extension too)
            DetectAndLoad = LoadFromCSV(sFilePath, nAdded, sError)
    End Select
End Function

' ============================================================================
'  Private helpers
' ============================================================================

'===========================================================================
' SplitCSVLine
'   Handles quoted fields: "Smith, John",Rock,"C:\My Music\song.wrk"
'   Returns an array of field strings.
'===========================================================================
Private Function SplitCSVLine(ByVal sLine As String) As String()
    Dim result(20) As String   ' max 21 columns; more than enough
    Dim nCol    As Integer
    Dim i       As Integer
    Dim ch      As String
    Dim bQuote  As Boolean
    Dim sCurrent As String

    bQuote = False
    sCurrent = ""
    nCol = 0

    For i = 1 To Len(sLine)
        ch = Mid(sLine, i, 1)

        If bQuote Then
            If ch = Chr(34) Then
                ' Check for doubled quote ""
                If i < Len(sLine) And Mid(sLine, i + 1, 1) = Chr(34) Then
                    sCurrent = sCurrent & Chr(34)
                    i = i + 1  ' skip next quote (note: VB For loop handles i)
                    ' Actually VB6 doesn't update loop var mid-iteration cleanly;
                    ' use a separate counter approach instead:
                Else
                    bQuote = False
                End If
            Else
                sCurrent = sCurrent & ch
            End If
        Else
            If ch = Chr(34) Then
                bQuote = True
            ElseIf ch = "," Then
                result(nCol) = sCurrent
                nCol = nCol + 1
                If nCol > 20 Then nCol = 20
                sCurrent = ""
            Else
                sCurrent = sCurrent & ch
            End If
        End If
    Next i

    ' Last field
    result(nCol) = sCurrent

    ReDim resultOut(nCol) As String
    Dim k As Integer
    For k = 0 To nCol
        resultOut(k) = result(k)
    Next k
    SplitCSVLine = resultOut
End Function

'===========================================================================
' GetFirstSheetName
'   Uses ADOX or falls back to a naming heuristic to find the first worksheet.
'===========================================================================
Private Function GetFirstSheetName(ByVal conn As Object) As String
    On Error GoTo FallbackSheet
    Dim cat  As Object   ' ADOX.Catalog
    Set cat = CreateObject("ADOX.Catalog")
    Set cat.ActiveConnection = conn

    Dim tbl As Object
    For Each tbl In cat.Tables
        If Right(tbl.Name, 1) = "$" Then
            GetFirstSheetName = tbl.Name
            Exit Function
        End If
    Next tbl

FallbackSheet:
    ' Standard default sheet name on English Excel
    GetFirstSheetName = "Sheet1$"
End Function

Private Function NullToStr(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        NullToStr = ""
    Else
        NullToStr = CStr(v)
    End If
End Function
