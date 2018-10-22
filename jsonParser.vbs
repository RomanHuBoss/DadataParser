'------------------------------------------------------------------------------------------------------------
'-- ==================  CLASS JSON ==============================
'------------------------------------------------------------------------------------------------------------
' Usage: 
'  1)  Call  LoadJSON s      where "s" is a full JSON text string. (Do Not use snippets.) LoadJSON will parse the text for future calls.

'  2) Call s = GetProp(Propname, Multi, ParentLevel)
'
'  Propname - a "qualified property". (See accompanying explanation.) It can be like "obj1.obj2.value1" or just "value1"
'         The format will depend on the particular JSON text. 

' Multi - Boolean. Is a single value expected, or multiple? (Examples: A top-level property will be a single value. A property of a parent "object" may be part of an array.)

'  ParentLevel - This will usually be 0. ParentLevel is the context where the value could be. In other words, thinking of
'  PropName like an object hierarchy, what level should the search happen in? If PropName is a.b.c.d, is the search within
' "a" level (0), "b" level, or "c" level?
'  The only case where this might not be 0 is where an object must be within its parent. 
'  Example: GetProp("obj1.obj2.value1", False, 1) implies there is only one obj2, and value1 will be found there.
'  So the context is level 1, obj2 scope. In that case the code will find obj2 within obj1 and then assume value1
'   is there.
'      But if obj1 holds an array, of which each item holds an obj2, then there are multiple value1 properties, in multiple obj2's, so the search context must be obj1.
'   This is a dificult concept to explain but is obvious once you have real JSON with real parsing goals.

' Return is the property value as a string. In the case of multiple values, the values will be delimited by Chr(25): val1*val2*val3*val4

'-- It's important to remember that none of this is arbitrary discovery. In other words,
'-- if you're parsing JSON then you know what you're looking for, so you know how
'-- to phrase the PropName parameter. It's not self-documenting
'-- and not partially discoverable the way a COM object hierarchy might be.
'-- For instance, if you deal with Internet Explorer Document Object Model it's vast, but
'-- properties and methods can be looked up. And the basic function of the DOM provides
'-- a sense of how the object hierarchy will be set up.
'--    With JSON, any level of nesting is possible and the properties one is looking for
'--  might be anything. It can be storing a list of images in a webpage or a genealogy,
'--  with a particular family tree being stored in its hierarchical order. If it's stored using
'-- JSON then it *probably* required a nested structure. Otherwise it could be merely a list.

Class JSON
Private Q2, C20, C21, C22, C23, C24, C25
Private sJ1

Private Sub Class_Initialize()
  Q2 = Chr(34)
  C20 = Chr(20)
  C21 = Chr(21)  '[
  C22 = Chr(22)  ']
  C23 = Chr(23)  '{
  C24 = Chr(24)  '}
  C25 = Chr(25)  ' ,
End Sub


's2 = CJ.GetProp("legs.start_address", False, 0) & " to " & CJ.GetProp("legs.end_address", False, 0) & vbCrLf & vbCrLf
 ' sSteps = CJ.GetProp("steps.html_instructions", True, 0)
 ' A1 = Split(sSteps, Chr(25))

'  Explanation of this sample: First call to GetProp is looking for a start_address property and end_address property
' in the legs object. Only a single value is expected, so there's only one legs object, or at any rate, only the first is relevant.
' The next call finds the steps object and retrieves all html_instructions properties. In this case it's expected
' to be multiple properties.

Public Function GetProp(sProp, Multi, ParentLevel)
  Dim Pt1, PtSt, PtEnd, PtEnd2, PtMin, PtMax, PtMax2
  Dim SearchLevel, iParent, i2, iAdd, UB2
  Dim AProp 
  Dim sMark, sVal
        On Error Resume Next
   AProp = Split(sProp, ".")
   UB2 = UBound(AProp)
   Pt1 = 1: PtMax = Len(sJ1): PtMax2 = PtMax: PtMin = 1
   If PtMax = 0 Then Exit Function ' need to load json first.
   SearchLevel = 0
    If UB2 > 0 Then
        For iParent = 0 To ParentLevel
             PtSt = InStr(Pt1, sJ1, AProp(iParent) & ":")
             iAdd = (Len(AProp(iParent)) + 1)
             PtSt = PtSt + iAdd
             sMark = Mid(sJ1, PtSt, 1) ' parsing changed "name:{" to "name:*{"  where * is char. for iDepth.
             PtEnd = InStr(PtSt + 2, sJ1, sMark) 'find ending uni-character that marks end of this item.
                If PtEnd = 0 Then PtEnd = Len(sJ1)
             PtMin = PtSt: PtMax = PtEnd: PtMax2 = PtMax
             SearchLevel = SearchLevel + 1
        Next
    End If
     
  Do '--loop needed for multiple values.
    For i2 = SearchLevel To UB2
      If i2 = UB2 Then 'get property.
             PtSt = InStr(PtMin, sJ1, AProp(i2) & ":") ' find "propertyname:". must be after start of object and before end.
               If (PtSt = 0) Or (PtSt > PtMax2) Then
                 If Len(GetProp) > 0 Then GetProp = CleanUpMarkers(GetProp)
                 Exit Function
               End If
             iAdd = Len(AProp(i2)) + 1
             PtSt = PtSt + iAdd
             ' in case the value is an array.
              sVal = ""
             If Mid(sJ1, PtSt + 1, 1) = C21 Then  ' if [
                 PtEnd = InStr(PtSt, sJ1, C22)   ' find ]
                 If PtEnd = 0 Then Exit Function
                 sVal = Mid(sJ1, PtSt + 2, (PtEnd - PtSt) - 2)
                 sVal = Replace(sVal, C20, ",")
             Else
                 PtEnd = InStr(PtSt, sJ1, C20)   '  ,
                 PtEnd2 = InStr(PtSt, sJ1, C24)  ' }
                 If (PtEnd2 > PtSt) And ((PtEnd2 < PtEnd) Or (PtEnd = 0)) Then PtEnd = PtEnd2
                    If PtEnd = 0 Then Exit Function
                 sVal = Mid(sJ1, PtSt, PtEnd - PtSt)   
             End If      
             If Len(GetProp) > 0 Then
                GetProp = GetProp & C25 & sVal
             Else
                GetProp = sVal
             End If
              If Multi = False Then
                If Len(GetProp) > 0 Then GetProp = CleanUpMarkers(GetProp)
                Exit Function
              End If
            PtMin = PtEnd
       Else
           PtSt = InStr(PtMin, sJ1, AProp(i2) & ":")
           iAdd = (Len(AProp(i2)) + 1)
           PtSt = PtSt + iAdd
           sMark = Mid(sJ1, PtSt, 1) ' parsing changed "name:{" to "name:*{"  where * is char. for iDepth.
           PtEnd = InStr(PtSt + 1, sJ1, sMark) 'find ending uni-character that marks end of this item.
             If PtEnd = 0 Then PtEnd = Len(sJ1)
           PtMin = PtSt
           PtMax2 = PtEnd
      End If
    Next
  Loop
End Function

Private Function CleanUpMarkers(sIn) 
  Dim Len1, i2, iSaved 
  Dim s1
  Dim A1()
      On Error Resume Next
  CleanUpMarkers = ""    
  Len1 = Len(sIn)
   If Len1 = 0 Then Exit Function
  ReDim A1(Len1 - 1)
     iSaved = -1
     
    For i2 = 1 To Len1
      s1 = Mid(sIn, i2, 1)
      If Asc(s1) > 24 Then
         iSaved = iSaved + 1
         A1(iSaved) = s1
      End If
    Next
  If iSaved > -1 Then
      ReDim Preserve A1(iSaved)
      CleanUpMarkers = Join(A1, "")
  End If
End Function

Public Sub LoadJSON(s1 )
  Dim LT1, iPos,  iPosNew, iDepth
  Dim InQuote
  Dim sChar, sChar2
  Dim sUni
  Dim A1()
    On Error Resume Next
  LT1 = Len(s1)
  ReDim A1(LT1 - 1)
  iPos = 1
  iPosNew = 0
  iDepth = 1
  InQuote = False
        
 Do While iPos < LT1
     sChar = Mid(s1, iPos, 1)
          Select Case Asc(sChar)
     Case 92 '"\"
       sChar2 = Mid(s1, iPos + 1, 1)
             Select Case Asc(sChar2)
                Case 34, 47, 92  ' Q2, "/", "\" 
                   A1(iPosNew) = sChar2
                   iPos = iPos + 1: iPosNew = iPosNew + 1
                Case 85, 117  ' "U", "u" 
                   sUni = Mid(s1, iPos + 4, 2)
                   A1(iPosNew) = Chr(CByte("&H" & sUni))
                   iPos = iPos + 5: iPosNew = iPosNew + 1
                Case Else
                 iPos = iPos + 1  '-- dump any returns, tabs, etc.
           End Select
           
     Case 34    'Q2 '   "
       InQuote = Not InQuote 'dump " but keep track for spaces.
       
     Case 32   ' " "
       If InQuote = True Then
         A1(iPosNew) = " "
         iPosNew = iPosNew + 1
       End If
             
     Case 91, 123  ' "[", "{"
       If InQuote = False Then
         iDepth = iDepth + 1
         A1(iPosNew) = Chr(iDepth) 
         If sChar = "[" Then    
            A1(iPosNew + 1) = C21  '"["
         Else
            A1(iPosNew + 1) = C23  '"{"
         End If
           iPosNew = iPosNew + 2
       Else
           If sChar = "[" Then    
              A1(iPosNew) = "["
           Else
              A1(iPosNew) = "{"
           End If
         iPosNew = iPosNew + 1
       End If
       
     Case 93, 125   '"]", "}"  
       If InQuote = False Then
          If sChar = "]" Then
               A1(iPosNew) = C22  '"]"
          Else
               A1(iPosNew) = C24  ' "}"
         End If
           A1(iPosNew + 1) = Chr(iDepth)
           iDepth = iDepth - 1
          iPosNew = iPosNew + 2
       Else
           If sChar = "]" Then
               A1(iPosNew) = "]"
           Else
               A1(iPosNew) = "}"
         End If
            iPosNew = iPosNew + 1
       End If
       
     Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31
       '-- drop all returns, etc.
     Case 44  '  ,
       If InQuote = False Then
        A1(iPosNew) = C20 'use nonsense char. to replace comma outside of quotes.
         iPosNew = iPosNew + 1
       Else
         A1(iPosNew) = ","
         iPosNew = iPosNew + 1
       End If
       
     Case Else
        A1(iPosNew) = sChar
        iPosNew = iPosNew + 1
   End Select
     iPos = iPos + 1
 Loop

    ReDim Preserve A1(iPosNew - 1)
     sJ1 = Join(A1, "")
     
End Sub  

End Class