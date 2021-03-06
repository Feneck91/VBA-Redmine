Attribute VB_Name = "RemineHelper"
' All the code below is used to encore URL without using ScriptControl
' Found here : https://stackoverflow.com/questions/218181/how-can-i-url-encode-a-string-in-excel-vba

Private Const CP_UTF8 = 65001

#If VBA7 Then
    Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, _
                                                                         ByVal dwFlags As Long, _
                                                                         ByVal lpWideCharStr As LongPtr, _
                                                                         ByVal cchWideChar As Long, _
                                                                         ByVal lpMultiByteStr As LongPtr, _
                                                                         ByVal cbMultiByte As Long, _
                                                                         ByVal lpDefaultChar As Long, _
                                                                         ByVal lpUsedDefaultChar As Long _
                                                                        ) As Long
#Else
    Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, _
                                                                 ByVal dwFlags As Long, _
                                                                 ByVal lpWideCharStr As Long, _
                                                                 ByVal cchWideChar As Long, _
                                                                 ByVal lpMultiByteStr As Long, _
                                                                 ByVal cbMultiByte As Long, _
                                                                 ByVal lpDefaultChar As Long, _
                                                                 ByVal lpUsedDefaultChar As Long _
                                                                ) As Long
#End If

Public Function UTF16To8(ByVal UTF16 As String) As String
    Dim sBuffer As String
    Dim lLength As Long
    If UTF16 <> "" Then
        #If VBA7 Then
            lLength = WideCharToMultiByte(CP_UTF8, 0, CLngPtr(StrPtr(UTF16)), -1, 0, 0, 0, 0)
        #Else
            lLength = WideCharToMultiByte(CP_UTF8, 0, StrPtr(UTF16), -1, 0, 0, 0, 0)
        #End If
        
        sBuffer = Space$(lLength)
        #If VBA7 Then
            lLength = WideCharToMultiByte(CP_UTF8, 0, CLngPtr(StrPtr(UTF16)), -1, CLngPtr(StrPtr(sBuffer)), LenB(sBuffer), 0, 0)
        #Else
            lLength = WideCharToMultiByte(CP_UTF8, 0, StrPtr(UTF16), -1, StrPtr(sBuffer), LenB(sBuffer), 0, 0)
        #End If
        
        sBuffer = StrConv(sBuffer, vbUnicode)
        UTF16To8 = Left$(sBuffer, lLength - 1)
    Else
        UTF16To8 = ""
    End If
End Function

Public Function URLEncode(StringVal As String, _
                          Optional SpaceAsPlus As Boolean = False, _
                          Optional UTF8Encode As Boolean = True _
                         ) As String

    Dim StringValCopy As String: StringValCopy = IIf(UTF8Encode, UTF16To8(StringVal), StringVal)
    Dim StringLen As Long: StringLen = Len(StringValCopy)

    If StringLen > 0 Then
        ReDim Result(StringLen) As String
        Dim I As Long, CharCode As Integer
        Dim Char As String, Space As String

        If SpaceAsPlus Then Space = "+" Else Space = "%20"

        For I = 1 To StringLen
            Char = Mid$(StringValCopy, I, 1)
            CharCode = Asc(Char)
            Select Case CharCode
                Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
                    Result(I) = Char
                Case 32
                    Result(I) = Space
                Case 0 To 15
                    Result(I) = "%0" & Hex(CharCode)
                Case Else
                    Result(I) = "%" & Hex(CharCode)
            End Select
        Next I
        URLEncode = Join(Result, "")
    End If
End Function

Function CloneDictionary(ByRef dict As Dictionary)
    Dim newDict
    Set newDict = CreateObject("Scripting.Dictionary")

    For Each key In dict.Keys
        Call newDict.Add(key, dict(key))
    Next
    newDict.CompareMode = dict.CompareMode

    Set CloneDictionary = newDict
End Function

