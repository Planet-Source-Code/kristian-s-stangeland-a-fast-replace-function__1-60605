Attribute VB_Name = "modStringReplace"
Option Explicit

'Copyright (C) 2004 Kristian. S. Stangeland

'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or (at your option) any later version.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (ptr() As Any) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SysAllocStringByteLen Lib "oleaut32" (ByVal lpString As Long, ByVal lLen As Long) As String
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private Type SAFEARRAYBOUND
    cElements As Long
    lLBound As Long
End Type

Private Type SafeArray1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0) As SAFEARRAYBOUND
End Type

Private Sub ReplaceUnequal(sText As String, sFind As String, sReplaceWith As String, sReturn As String, ByVal lStart As Long, ByVal lCount As Long)

    Dim aSource() As Integer, aChar() As Integer, Tell As Long, Temp As Long, ReplacePos As Long
    Dim lngFind() As Integer, lngResult() As Integer, lenghtFind As Long, lenghtResult As Long
    Dim lenghtText As Long, lngDifference As Long, diffValue As Long, bAdd As Boolean, bExit As Boolean
    Dim ByteArray As SafeArray1D, ReturnArray As SafeArray1D, lngFindZBound As Long, newLenght As Long
    Dim lngTempCount As Long, bUseCount As Boolean, lEnd As Long, lngResultZBound As Long, lngOverSize As Long

    ' See if there's anything to replace
    If lCount = 0 Then
        
        ' Just return the string unchanged
        sReturn = sText
        
        ' Exit procedure
        Exit Sub

    End If

    ' Should we replace anything
    bUseCount = CBool(lCount > 0)
    
    ' A extra temp variable holding the count
    lngTempCount = lCount

    ' Variables are much faster than functions, so save the lenght of the strings in variables
    lenghtFind = Len(sFind)
    lenghtResult = Len(sReplaceWith)
    lenghtText = Len(sText)
    lngFindZBound = lenghtFind - 1
    lngResultZBound = lenghtResult - 1

    ' Get the difference
    diffValue = (lenghtResult - lenghtFind)

    ' Use integer variables, since this is quite faster
    ReDim lngFind(lngFindZBound)
    
    ' Copy array
    CopyMemory lngFind(0), ByVal StrPtr(sFind), lenghtFind * 2
    
    ' Don't declare the array if it's supposed to be empty
    If lngResultZBound >= 0 Then
    
        ' Allocate the array
        ReDim lngResult(lngResultZBound)
        
        ' Copy the string
        CopyMemory lngResult(0), ByVal StrPtr(sReplaceWith), lenghtResult * 2
        
    End If

    ' Initialize array pointing to the string
    With ByteArray
        .cbElements = 0
        .cDims = 1
        .Bounds(0).lLBound = 1
        .Bounds(0).cElements = Len(sText)
        .pvData = StrPtr(sText)
    End With
    
    ' Set the safe array of this array
    CopyMemory ByVal VarPtrArray(aSource), VarPtr(ByteArray), 4
    
    ' Calculate where we need to read the data
    If lStart < 0 Then
        ' Simply use the beginning
        lStart = LBound(aSource)
    End If
    
    ' The position to search to (always the last character)
    lEnd = UBound(aSource)
     
    ' The new lenght is originaly the same size as the source string
    newLenght = (lEnd - lStart) + 1
    
    ' How many steps too far the search procedure might go
    lngOverSize = (lEnd Mod lenghtFind)
    
    ' This fixes the problem where the search prosedure sometimes goes to far
    lEnd = lEnd - lngOverSize
    
    ' Find all areas which needs to be replaced
    For Tell = lStart To lEnd
        
        ' If the character is equal to what we are looking for ...
        For Temp = 0 To lngFindZBound
            
            If aSource(Tell + Temp) = lngFind(Temp) Then
                
                If Temp = lngFindZBound Then
                
                    ' Increse or decrese the lenght of the result array
                    newLenght = newLenght + diffValue
                
                    ' Move forward
                    Tell = Tell + lngFindZBound
                    
                    ' See if we should stop
                    If bUseCount Then
                    
                        ' Decrese count
                        lngTempCount = lngTempCount - 1
                        
                        ' Exit if it gets to zero or below
                        If lngTempCount <= 0 Then
                    
                            ' Exit the next for-loop also
                            bExit = True
                    
                            ' Exit this loop
                            Exit For
                    
                        End If
                    
                    End If
                
                End If
            
            Else
                ' There's no match here
                Exit For
            End If
            
        Next
        
        ' If we are set to exit this for-loop, do so
        If bExit Then
            Exit For
        End If
    
    Next
    
    ' Reset variables
    bExit = False
    
    ' There's no point of doing all this if the new size is empty or equal to the original
    If newLenght > 0 And newLenght <> lenghtText Then
    
        ' Allocate the return buffer
        sReturn = SysAllocStringByteLen(0, newLenght)
    
        ' Initialize array pointing to the string
        With ReturnArray
            .cbElements = 0
            .cDims = 1
            .Bounds(0).lLBound = 1
            .Bounds(0).cElements = newLenght
            .pvData = StrPtr(sReturn)
        End With
        
        ' Set the new safe array
        CopyMemory ByVal VarPtrArray(aChar), VarPtr(ReturnArray), 4
        
        ' Add characters normaly
        bAdd = True
        
        ' Insert the un-searchable partition of string
        MoveString sReturn, sText, 1, 1, lStart - 1
        
        ' See if there actually is a bit where we not going to search
        If lngOverSize <> 0 Then
            ' Insert the last bit that isn't going to be search in
            MoveString sReturn, sText, newLenght, lEnd, lngOverSize
        End If
        
        ' Replace all characters corresponding to sFind
        For Tell = lStart To lEnd
         
            For Temp = 0 To lngFindZBound
            
                ' If the character is equal to what we are looking for ...
                If aSource(Tell + Temp) = lngFind(Temp) Then
                
                    ' ... replace it, if this is the last character
                    If Temp = lngFindZBound Then
                    
                        For ReplacePos = 0 To lngResultZBound
                            ' Replace the character
                            aChar(Tell + ReplacePos + lngDifference) = lngResult(ReplacePos)
                        Next
    
                        ' Move forward
                        Tell = Tell + lngFindZBound
                        
                        ' Increse or decrese the difference
                        lngDifference = lngDifference + diffValue
                        
                        ' Don't do anything more
                        bAdd = False
                        
                        ' See if we should stop
                        If bUseCount Then
                        
                            ' Decrese count
                            lCount = lCount - 1
                            
                            ' Exit if it gets to zero or below
                            If lCount <= 0 Then
                        
                                ' Copy the rest of the source into the string
                                MoveString sReturn, sText, Tell + lngDifference, Tell, UBound(aSource) - Tell + 1
                        
                                ' Exit the next for-loop also
                                bExit = True
                        
                                ' Exit this loop
                                Exit For
                        
                            End If
                        
                        End If
                        
                    End If
                
                Else
                
                    ' There's no match here
                    Exit For
                    
                End If
                
            Next

            If bAdd Then
            
                ' Set the character
                aChar(Tell + lngDifference) = aSource(Tell)
            
            Else
            
                ' Don't ignore anymore
                bAdd = True
            
            End If
        
        Next

    Else
    
        ' Just return the text
        sReturn = sText
    
    End If

    ' Clear the safe array and the returned string-array
    CopyMemory ByVal VarPtrArray(aSource), 0&, 4
    CopyMemory ByVal VarPtrArray(aChar), 0&, 4

End Sub

Private Sub ReplaceEqualSize(sText As String, sFind As String, sReplaceWith As String, ByVal lStart As Long, ByVal lCount As Long)

    Dim ByteArray As SafeArray1D, aChar() As Integer, Tell As Long, Temp As Long, ReplacePos As Long
    Dim lngFind() As Integer, lngResult() As Integer, lenghtFind As Long, lenghtResult As Long
    Dim lngFindZBound As Long, lngResultZBound As Long, bUseCount As Boolean

    ' Means that where not replacing anything
    If lCount = 0 Then
        Exit Sub
    End If

    ' Should we limit the amount of replacements?
    bUseCount = CBool(lCount > 0)

    ' Variables are much faster than functions, so save the lenght of the strings in variables
    lenghtFind = Len(sFind)
    lenghtResult = Len(sReplaceWith)
    lngFindZBound = lenghtFind - 1
    lngResultZBound = lenghtResult - 1

    ' Use integer variables, since this is quite faster
    ReDim lngFind(lngFindZBound)
    
    ' Copy array
    CopyMemory lngFind(0), ByVal StrPtr(sFind), lenghtFind * 2
    
    ' Don't declare the array if it's supposed to be empty
    If lngResultZBound >= 0 Then
    
        ' Allocate the array
        ReDim lngResult(lngResultZBound)
        
        ' Copy the string
        CopyMemory lngResult(0), ByVal StrPtr(sReplaceWith), lenghtResult * 2
        
    End If

    ' Initialize array pointing to the string
    With ByteArray
        .cbElements = 0
        .cDims = 1
        .Bounds(0).lLBound = 1
        .Bounds(0).cElements = Len(sText)
        .pvData = StrPtr(sText)
    End With
    
    ' Set the safe array of this array
    CopyMemory ByVal VarPtrArray(aChar), VarPtr(ByteArray), 4

    ' Calculate where we need to read the data
    If lStart < 0 Then
        ' Simply use the beginning
        lStart = LBound(aChar)
    End If

    ' Replace all characters corresponding to sFind
    For Tell = lStart To UBound(aChar)
    
        ' If the character is equal to what we are looking for ...
        For Temp = 0 To lngFindZBound
            If aChar(Tell + Temp) = lngFind(Temp) Then
            
                ' ... replace it, if this is the last character
                If Temp = lngFindZBound Then
                
                    For ReplacePos = 0 To lngResultZBound
                        ' Replace the character
                        aChar(Tell + ReplacePos) = lngResult(ReplacePos)
                    Next
                    
                    ' Move forward
                    Tell = Tell + lngFindZBound
                
                    ' Tell the number if replacements if specified
                    If bUseCount Then
                    
                        ' Decrese the count
                        lCount = lCount - 1
                        
                        ' If it gets till zero or lower
                        If lCount <= 0 Then
                            
                            ' Clear up
                            CopyMemory ByVal VarPtrArray(aChar), 0&, 4
                            
                            ' Stop the replacement-process
                            Exit Sub
                            
                        End If
                    
                    End If
                
                End If
            
            Else
                ' There's no match here
                Exit For
            End If
        Next
    
    Next

    ' Clear the safe array
    CopyMemory ByVal VarPtrArray(aChar), 0&, 4

End Sub

Private Function MoveString(sDest As String, sSource As String, lngDestPos As Long, lngSourcePos As Long, lngLenght As Long)

    ' The lenght must of course be valid
    If lngLenght > 0 Then
        
        ' Copy the source over to the destination using the two coordinates
        CopyMemory ByVal (StrPtr(sDest) + ((lngDestPos - 1) * 2)), ByVal (StrPtr(sSource) + ((lngSourcePos - 1) * 2)), lngLenght * 2
    
    End If

End Function

Public Function FastReplace(sText As String, sFind As String, sReplaceWith As String, Optional Start As Long = -1, Optional Count As Long = -1) As String

    If LenB(sText) = 0 Or LenB(sFind) = 0 Then
        ' There's no point of continuing; return text.
        FastReplace = sText
        Exit Function
    End If

    ' Use the fastest function when the lenght of sFind is equal to the lenght of sReplaceWith
    If LenB(sFind) = LenB(sReplaceWith) Then
            
        ' Use a buffer, since the function below is a sub
        FastReplace = sText
    
        ' Replace everything within the buffer using the optimized function
        ReplaceEqualSize FastReplace, sFind, sReplaceWith, Start, Count
    
    Else
    
        ' If the replacement string is lagrer or smaller, a different function needs to be used
        ReplaceUnequal sText, sFind, sReplaceWith, FastReplace, Start, Count
        
    End If

End Function

