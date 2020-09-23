Attribute VB_Name = "modSortListView"
Option Explicit

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Enum ListDataType
    ldtString = 0
    ldtNumber = 1
    ldtDateTime = 2
End Enum

Public Sub SortListView(ListView As ListView, ByVal Index As Integer, _
                ByVal DataType As ListDataType, ByVal Ascending As Boolean)

    On Error Resume Next
    Dim I As Integer
    Dim l As Long
    Dim strFormat As String
  
    Dim lngCursor As Long
    lngCursor = ListView.MousePointer
    ListView.MousePointer = vbHourglass
    
    LockWindowUpdate ListView.hwnd
    
    Dim blnRestoreFromTag As Boolean
    
    Select Case DataType
    Case ldtString
   
        blnRestoreFromTag = False
        
    Case ldtNumber
    
        strFormat = String$(20, "0") & "." & String$(10, "0")
        
        With ListView.ListItems
            If (Index = 1) Then
                For l = 1 To .Count
                    With .Item(l)
                        .Tag = .Text & Chr$(0) & .Tag
                        If IsNumeric(.Text) Then
                            If CDbl(.Text) >= 0 Then
                                .Text = Format(CDbl(.Text), strFormat)
                            Else
                                .Text = "&" & InvNumber(Format(0 - CDbl(.Text), strFormat))
                            End If
                        Else
                            .Text = ""
                        End If
                    End With
                Next l
            Else
                For l = 1 To .Count

                Next l
            End If
        End With
        
        blnRestoreFromTag = True
    
    Case ldtDateTime
    
        strFormat = "YYYYMMDDHhNnSs"
        
        Dim dte As Date
    
        With ListView.ListItems
            If (Index = 1) Then
                For l = 1 To .Count
                    With .Item(l)
                        .Tag = .Text & Chr$(0) & .Tag
                        dte = CDate(.Text)
                        .Text = Format$(dte, strFormat)
                    End With
                Next l
            Else
                For l = 1 To .Count
'                    With .Item(l).ListSubItems(Index - 1)
'                        .Tag = .Text & Chr$(0) & .Tag
'                        dte = CDate(.Text)
'                        .Text = Format$(dte, strFormat)
'                    End With
                Next l
            End If
        End With
        
        blnRestoreFromTag = True
        
    End Select
    
    ListView.SortOrder = IIf(Ascending, lvwAscending, lvwDescending)
    ListView.SortKey = Index - 1
    ListView.Sorted = True
    
    If blnRestoreFromTag Then
        
        With ListView.ListItems
            If (Index = 1) Then
                For l = 1 To .Count
                    With .Item(l)
                        I = InStr(.Tag, Chr$(0))
                        .Text = Left$(.Tag, I - 1)
                        .Tag = Mid$(.Tag, I + 1)
                    End With
                Next l
            Else
                For l = 1 To .Count
'                    With .Item(l).ListSubItems(Index - 1)
'                        i = InStr(.Tag, Chr$(0))
'                        .Text = Left$(.Tag, i - 1)
'                        .Tag = Mid$(.Tag, i + 1)
'                    End With
                Next l
            End If
        End With
    End If
    
    LockWindowUpdate 0&
    
    ListView.MousePointer = lngCursor
    
End Sub

Private Function InvNumber(ByVal Number As String) As String
    Static I As Integer
    For I = 1 To Len(Number)
        Select Case Mid$(Number, I, 1)
        Case "-": Mid$(Number, I, 1) = " "
        Case "0": Mid$(Number, I, 1) = "9"
        Case "1": Mid$(Number, I, 1) = "8"
        Case "2": Mid$(Number, I, 1) = "7"
        Case "3": Mid$(Number, I, 1) = "6"
        Case "4": Mid$(Number, I, 1) = "5"
        Case "5": Mid$(Number, I, 1) = "4"
        Case "6": Mid$(Number, I, 1) = "3"
        Case "7": Mid$(Number, I, 1) = "2"
        Case "8": Mid$(Number, I, 1) = "1"
        Case "9": Mid$(Number, I, 1) = "0"
        End Select
    Next
    InvNumber = Number
End Function
