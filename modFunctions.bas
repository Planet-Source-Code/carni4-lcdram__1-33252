Attribute VB_Name = "modFunctions"
Option Explicit

Public Function GetRAMInfo()
'Calls the API function
GlobalMemoryStatus MemoryInfo

'Calculates the RAM info
intTotal = Round((MemoryInfo.dwTotalPhys / 1024) / 1024)
intFree = Round((MemoryInfo.dwAvailPhys / 1024) / 1024)

'Puts the Integers into Strings, this is easier for ShowRAMInfo
strTotal = intTotal
strFree = intFree
End Function

Public Function ShowRAMInfo()
'This function checks the length of strTotal and strFree and puts their values
'in the ImageBoxes

With frmMain
Select Case Len(strTotal)
Case 1
    .imgT1.Picture = .picLCD(Val(strTotal)).Picture
    .imgT2.Picture = .picEmpty.Picture
    .imgT3.Picture = .picEmpty.Picture
    .imgT4.Picture = .picEmpty.Picture
Case 2
    .imgT1.Picture = .picLCD(Val(Right(strTotal, 1))).Picture
    .imgT2.Picture = .picLCD(Val(Left(strTotal, 1))).Picture
    .imgT3.Picture = .picEmpty.Picture
    .imgT4.Picture = .picEmpty.Picture
Case 3
    .imgT1.Picture = .picLCD(Val(Right(strTotal, 1))).Picture
    .imgT2.Picture = .picLCD(Val(Mid(strTotal, 2, 1))).Picture
    .imgT3.Picture = .picLCD(Val(Left(strTotal, 1))).Picture
    .imgT4.Picture = .picEmpty.Picture
Case 4
    .imgT1.Picture = .picLCD(Val(Right(strTotal, 1))).Picture
    .imgT2.Picture = .picLCD(Val(Mid(strTotal, 3, 1))).Picture
    .imgT3.Picture = .picLCD(Val(Mid(strTotal, 2, 1))).Picture
    .imgT4.Picture = .picLCD(Val(Left(strTotal, 1))).Picture
End Select

Select Case Len(strFree)
Case 1
    .imgF1.Picture = .picLCD(Val(strFree)).Picture
    .imgF2.Picture = .picEmpty.Picture
    .imgF3.Picture = .picEmpty.Picture
    .imgF4.Picture = .picEmpty.Picture
Case 2
    .imgF1.Picture = .picLCD(Val(Right(strFree, 1))).Picture
    .imgF2.Picture = .picLCD(Val(Left(strFree, 1))).Picture
    .imgF3.Picture = .picEmpty.Picture
    .imgF4.Picture = .picEmpty.Picture
Case 3
    .imgF1.Picture = .picLCD(Val(Right(strFree, 1))).Picture
    .imgF2.Picture = .picLCD(Val(Mid(strFree, 2, 1))).Picture
    .imgF3.Picture = .picLCD(Val(Left(strFree, 1))).Picture
    .imgF4.Picture = .picEmpty.Picture
Case 4
    .imgF1.Picture = .picLCD(Val(Right(strFree, 1))).Picture
    .imgF2.Picture = .picLCD(Val(Mid(strFree, 3, 1))).Picture
    .imgF3.Picture = .picLCD(Val(Mid(strFree, 2, 1))).Picture
    .imgF4.Picture = .picLCD(Val(Left(strFree, 1))).Picture
End Select
End With
End Function
