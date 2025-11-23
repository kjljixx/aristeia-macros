Attribute VB_Name = "debater"
' ===================================================================================
' - If text is selected, it runs ONLY on the selection.
' - If no text is selected, it runs on the ENTIRE document.
' It also resets styles at the end of every paragraph.
' ===================================================================================
Sub FormatStyleTags()
    Application.ScreenUpdating = False

    Dim searchRange As Range
    Dim processedArea As String

    If Selection.Type = wdSelectionNormal And Selection.Characters.Count > 1 Then
        ' User has selected a block of text.
        Set searchRange = Selection.Range
        processedArea = "selection"
    Else
        ' No text selected (or just a blinking cursor). Use the whole document.
        Set searchRange = ActiveDocument.Content
        processedArea = "entire document"
    End If
    
    ' State tracking variables
    Dim isUnderlined As Boolean: isUnderlined = False
    Dim isEmphasized As Boolean: isEmphasized = False
    Dim isHighlighted As Boolean: isHighlighted = False
    Dim isTag As Boolean: isTag = False
    Dim isCite As Boolean: isCite = False

    Do
        Dim nextEventRange As Range
        Set nextEventRange = FindNextEvent(searchRange)
        
        If nextEventRange Is Nothing Then Exit Do
        
        Dim contentRange As Range
        Set contentRange = ActiveDocument.Range(searchRange.Start, nextEventRange.Start)
        
        If contentRange.Characters.Count > 0 Then
            ' Apply formatting based on the current state of all our flags
            ApplyCurrentFormatting contentRange, isUnderlined, isEmphasized, isHighlighted, isTag, isCite
        End If
        
        If nextEventRange.Text = vbCr Then
            ' Reset all states at a paragraph break
            isUnderlined = False
            isEmphasized = False
            isHighlighted = False
            isTag = False
            isCite = False
            searchRange.Start = nextEventRange.End
        Else
            ' Update state based on the tag found
            Select Case LCase(nextEventRange.Text)
                Case "{underline}": isUnderlined = True
                Case "{/underline}": isUnderlined = False
                Case "{emphasize}": isEmphasized = True
                Case "{/emphasize}": isEmphasized = False
                Case "{highlight}": isHighlighted = True
                Case "{/highlight}": isHighlighted = False
                Case "{tag}": isTag = True
                Case "{/tag}": isTag = False
                Case "{cite}": isCite = True
                Case "{/cite}": isCite = False
            End Select
            ' Delete the tag after processing
            searchRange.Start = nextEventRange.Start
            nextEventRange.Delete
        End If
        
        If searchRange.Start >= searchRange.End Then Exit Do

    Loop While True
    
    ' Apply formatting to any remaining text at the end of the range
    If searchRange.Characters.Count > 0 Then
        ApplyCurrentFormatting searchRange, isUnderlined, isEmphasized, isHighlighted, isTag, isCite
    End If

    Application.ScreenUpdating = True
End Sub


' ===================================================================================
' Finds the closest tag OR paragraph mark from a starting point.
' ===================================================================================
Private Function FindNextEvent(searchRange As Range) As Range
    Dim tags As Variant
    ' Add the new tags to the array of items to search for
    tags = Array("{underline}", "{/underline}", "{emphasize}", "{/emphasize}", "{highlight}", "{/highlight}", "{tag}", "{/tag}", "{cite}", "{/cite}")
    
    Dim foundRange As Range, tempRange As Range
    Dim closestPos As Long
    closestPos = -1
    
    Dim tag As Variant
    For Each tag In tags
        Set tempRange = searchRange.Duplicate
        With tempRange.Find
            .Text = tag
            .Forward = True
            .Wrap = wdFindStop
            .MatchCase = False ' Make tag finding case-insensitive
            If .Execute Then
                If closestPos = -1 Or tempRange.Start < closestPos Then
                    closestPos = tempRange.Start
                    Set foundRange = tempRange.Duplicate
                End If
            End If
        End With
    Next tag
    
    Set tempRange = searchRange.Duplicate
    With tempRange.Find
        .Text = "^p" ' Special code for a paragraph mark
        .Forward = True
        .Wrap = wdFindStop
        If .Execute Then
            If closestPos = -1 Or tempRange.Start < closestPos Then
                closestPos = tempRange.Start
                Set foundRange = tempRange.Duplicate
            End If
        End If
    End With
    
    Set FindNextEvent = foundRange
End Function


' ===================================================================================
' Applies formatting to a range based on the current state.
' ===================================================================================
Private Sub ApplyCurrentFormatting(rng As Range, ByVal u As Boolean, ByVal e As Boolean, ByVal h As Boolean, ByVal t As Boolean, ByVal c As Boolean)
    ' The order of this If/ElseIf block defines priority.
    ' We check for the most specific styles (Tag, Cite) first.
    If t Then
        rng.Style = "Heading 4,Tag"
    ElseIf c Then
        rng.Style = "Style 13 pt Bold,Cite"
    ElseIf u And e Then
        rng.Style = "Emphasis"
    ElseIf u Then
        rng.Style = "Style Underline,Underline"
    ElseIf e Then
        rng.Style = "Emphasis"
    Else
        ' If no other style-based tags are active, apply Normal style.
        rng.Style = "Normal"
    End If
    
    ' Highlight is applied independently of the character style
    If (h And u) Or (h And e) Then
        rng.HighlightColorIndex = wdTurquoise
    Else
        rng.HighlightColorIndex = wdNoHighlight
    End If
End Sub


' ===================================================================================
' Converts selected formatted text back into tagged text.
' ===================================================================================
Sub DeformatStyleTags()

    If Selection.Type = wdSelectionIP Then
        MsgBox "Please select some text first.", vbInformation, "No Selection"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False

    Dim newString As String
    Dim char As Range
    
    ' State tracking variables to know when a format starts or stops.
    Dim isHighlighted As Boolean, isUnderlined As Boolean, isEmphasized As Boolean
    Dim isTag As Boolean, isCite As Boolean
    isHighlighted = False
    isUnderlined = False
    isEmphasized = False
    isTag = False
    isCite = False
    
    ' Loop through each character in the selection.
    For Each char In Selection.Characters
        ' Determine the formatting state of the CURRENT character.
        Dim currentHighlighted As Boolean, currentUnderlined As Boolean, currentEmphasized As Boolean
        Dim currentTag As Boolean, currentCite As Boolean
        
        currentHighlighted = (char.HighlightColorIndex = wdTurquoise)
        currentUnderlined = (char.Style = "Style Underline,Underline")
        currentEmphasized = (char.Style = "Emphasis")
        currentTag = (char.Style = "Heading 4,Tag")
        currentCite = (char.Style = "Style 13 pt Bold,Cite")

        ' Check if a format that was ON for the previous character is now OFF.
        If isCite And Not currentCite Then newString = newString & "{/cite}"
        If isTag And Not currentTag Then newString = newString & "{/tag}"
        If isEmphasized And Not currentEmphasized Then newString = newString & "{/emphasize}"
        If isUnderlined And Not currentUnderlined Then newString = newString & "{/underline}"
        If isHighlighted And Not currentHighlighted Then newString = newString & "{/highlight}"

        ' Check if a format that was OFF for the previous character is now ON.
        If Not isHighlighted And currentHighlighted Then newString = newString & "{highlight}"
        If Not isUnderlined And currentUnderlined Then newString = newString & "{underline}"
        If Not isEmphasized And currentEmphasized Then newString = newString & "{emphasize}"
        If Not isTag And currentTag Then newString = newString & "{tag}"
        If Not isCite And currentCite Then newString = newString & "{cite}"

        ' Add the actual text of the character.
        newString = newString & char.Text
        
        ' Update the state for the next loop iteration.
        isHighlighted = currentHighlighted
        isUnderlined = currentUnderlined
        isEmphasized = currentEmphasized
        isTag = currentTag
        isCite = currentCite
    Next char

    ' After the loop, close any tags that were still open at the very end.
    If isCite Then newString = newString & "{/cite}"
    If isTag Then newString = newString & "{/tag}"
    If isEmphasized Then newString = newString & "{/emphasize}"
    If isUnderlined Then newString = newString & "{/underline}"
    If isHighlighted Then newString = newString & "{/highlight}"
    
    ' Replace the user's formatted selection with the new tagged string.
    Selection.Text = newString
    
    Application.ScreenUpdating = True
End Sub
