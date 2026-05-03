Attribute VB_Name = "ThisWorkbook"

' ================================================================
'  JOURNAL DES MODIFICATIONS  —  Backlog POD / NURDIC / LINE
'  À coller dans : Éditeur VBA > ThisWorkbook
' ================================================================

Private Const JOURNAL_SHEET As String = "Journal"
Private Const JOURNAL_PWD   As String = "journal"
Private Const MAX_ROWS      As Long   = 10000

Private mOldValue As String
Private mOldCell  As String
Private mOldSheet As String
Private mTracking As Boolean

' ── À l'ouverture du fichier ────────────────────────────────────
Private Sub Workbook_Open()
    mTracking = True
    Application.EnableEvents = True
End Sub

' ── Mémorise la valeur AVANT modification ───────────────────────
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If Not mTracking Then Exit Sub
    If Sh.Name = JOURNAL_SHEET Then Exit Sub
    On Error Resume Next
    mOldValue = Target.Cells(1, 1).Value
    mOldCell  = Target.Cells(1, 1).Address(False, False)
    mOldSheet = Sh.Name
    On Error GoTo 0
End Sub

' ── Enregistre la modification dans le Journal ──────────────────
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Not mTracking Then Exit Sub
    If Sh.Name = JOURNAL_SHEET Then Exit Sub

    Dim newVal  As String
    Dim detail  As String
    Dim jn      As Worksheet
    Dim nextRow As Long

    On Error GoTo ErrHandler
    mTracking = False

    newVal = Target.Cells(1, 1).Value
    detail = "Cellule " & mOldCell & " : [" & mOldValue & "] -> [" & newVal & "]"

    Set jn = Me.Sheets(JOURNAL_SHEET)
    jn.Unprotect Password:=JOURNAL_PWD

    nextRow = jn.Cells(jn.Rows.Count, "A").End(xlUp).Row + 1
    If nextRow < 3 Then nextRow = 3
    If nextRow > MAX_ROWS + 2 Then
        MsgBox "Journal plein (limite " & MAX_ROWS & " entrees).", vbInformation
        GoTo Cleanup
    End If

    ' Numéro séquentiel
    jn.Cells(nextRow, 1).Value = nextRow - 2
    ' Date
    jn.Cells(nextRow, 2).Value = Format(Now(), "yyyy-mm-dd")
    ' Heure
    jn.Cells(nextRow, 3).Value = Format(Now(), "hh:mm:ss")
    ' Utilisateur Windows (variable d'environnement)
    jn.Cells(nextRow, 4).Value = Environ("USERNAME")
    ' Onglet modifié
    jn.Cells(nextRow, 5).Value = mOldSheet
    ' Cellule(s)
    jn.Cells(nextRow, 6).Value = Target.Address(False, False)
    ' Détail
    jn.Cells(nextRow, 7).Value = detail

    ' Mise en forme
    Dim rng As Range
    Set rng = jn.Range(jn.Cells(nextRow, 1), jn.Cells(nextRow, 7))
    With rng
        .Font.Name = "Arial"
        .Font.Size = 10
        If nextRow Mod 2 = 0 Then
            .Interior.Color = RGB(214, 228, 247)
        Else
            .Interior.Color = RGB(255, 255, 255)
        End If
        .Borders(xlEdgeLeft).LineStyle        = xlContinuous
        .Borders(xlEdgeRight).LineStyle       = xlContinuous
        .Borders(xlEdgeTop).LineStyle         = xlContinuous
        .Borders(xlEdgeBottom).LineStyle      = xlContinuous
        .Borders(xlInsideVertical).LineStyle  = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    jn.Cells(nextRow, 7).HorizontalAlignment = xlLeft

Cleanup:
    jn.Protect Password:=JOURNAL_PWD, DrawingObjects:=True, _
               Contents:=True, Scenarios:=True
    mTracking = True
    Exit Sub

ErrHandler:
    On Error Resume Next
    jn.Protect Password:=JOURNAL_PWD, DrawingObjects:=True, _
               Contents:=True, Scenarios:=True
    mTracking = True
End Sub
