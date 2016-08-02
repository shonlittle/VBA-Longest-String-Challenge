Attribute VB_Name = "mdlLongestString"
' --
' The Boston Conulting Group
' Data and Analytics Services (DaAS)
' Shon Little
' August 1, 2016
' --

' Options
Option Explicit

' Settings
Private Const MODULE    As String = "mdlLongestString"
Private Const DEBUGGING As Boolean = False

' Entry Point
Public Sub Main()
    Const strFirst      As String = "BCG's Advanced Analytics group delivers powerful analytics-based \n insights designed to help our clients tackle their most pressing \n business problems. The Analytics Solutions team is a global resource, \n working with clients and case teams in every industry area."
    Const strSecond     As String = "BCG has established the Analytics Solutions team to support case teams\n in managing and realizing insight and opportunity from the increasing \n analytical intensity of our clients' problems in every  industry area."
    Dim dblStartTime    As Double
    Dim strFinish       As String
    Dim strDebugMsg     As String
    Dim strTime         As String
    Dim varResult       As Variant
        
    ' Set mode
    If Not DEBUGGING Then
        On Error GoTo catch
        Application.ScreenUpdating = False
    End If
    
    ' Initialize
    dblStartTime = Timer
    strFinish = "Finished!"
    
    ' Routine
    varResult = LongestString(strFirst, strSecond)
    If IsArray(varResult) Then
        strFinish = strFinish & vbCrLf & Join(varResult, vbCrLf)
    Else
        strFinish = strFinish & vbCrLf & "No match found."
    End If
    
    ' Finish
    If DEBUGGING Then strDebugMsg = " with ""debugging"" turned on. " & vbCrLf & "Turning it off will improve performance."
    strTime = vbCrLf & "Time: " & Format((Timer - dblStartTime) / 86400, "hh:mm:ss")
       
finally:
    ' Reset application
    Application.ScreenUpdating = True
    ' Finish message
    MsgBox strFinish & strTime & strDebugMsg, vbOKOnly + vbInformation, "Finished"
    Exit Sub
    
catch:
    ' Error handler
    MsgBox "Error: " & Err.Description & " in " & MODULE & ".Main", vbCritical, "Error " & Err.Number
    Resume finally
End Sub

' Function to get an array of longest common substring
Private Function LongestString(ByVal strFirst As String, ByVal strSecond As String, Optional ByRef ignoreCase As Boolean, _
        Optional ByRef ignoreWhiteSpace As Boolean, Optional ByRef ignoreLineWrap As Boolean) As Variant
    Dim i               As Long
    Dim lngLen1         As Long
    Dim lngLen2         As Long
    Dim lngFull         As Long
    Dim lngSub          As Long
    Dim strShort        As String
    Dim strLong         As String
    Dim strSub          As String
    Dim strNeedle       As String
    Dim strHaystack     As String
    Dim objMatches      As Object

    ' Set mode
    If Not DEBUGGING Then On Error GoTo catch
    
    ' Check for blanks strings
    If strFirst = "" Or strSecond = "" Then GoTo finally

    ' Initialize
    lngLen1 = Len(strFirst)
    lngLen2 = Len(strSecond)
    Set objMatches = CreateObject("System.Collections.ArrayList")
    
    ' Handle line wrap
    If ignoreLineWrap Then
        strFirst = Replace(strFirst, "\n", "")
        strSecond = Replace(strSecond, "\n", "")
    End If
    
    ' Handle white space
    If ignoreWhiteSpace Then
        strFirst = Replace(strFirst, "  ", " ")
        strSecond = Replace(strSecond, "  ", " ")
    End If
    
    ' Get shorter and longer string
    If lngLen1 < lngLen2 Then
        strShort = strFirst
        strLong = strSecond
        lngFull = lngLen1
        lngSub = lngLen1
    Else
        strShort = strSecond
        strLong = strFirst
        lngFull = lngLen2
        lngSub = lngLen2
    End If
    
    ' Loop length of string to zero
    Do
        ' Loop substring
        For i = 0 To (lngFull - lngSub)
            ' Make substring
            strSub = Mid(strShort, i + 1, lngSub)
            strNeedle = strSub
            strHaystack = strLong
            ' Handle case
            If ignoreCase Then
                strNeedle = LCase(strNeedle)
                strHaystack = LCase(strHaystack)
            End If
            ' Check for match
            If InStr(strHaystack, strNeedle) > 0 Then
                ' Add matched substring to list
                If Not objMatches.Contains(strSub) Then objMatches.Add strSub
            End If
        Next i
        ' Deduct length
        lngSub = lngSub - 1
    ' Try again
    Loop Until objMatches.Count > 0 Or lngSub = 0
    
    ' Return
    LongestString = objMatches.ToArray()
       
finally:
    Exit Function
    
catch:
    ' Error handler
    MsgBox "Error: " & Err.Description & " in " & MODULE & ".", vbCritical, "Error " & Err.Number
    Resume finally
End Function
