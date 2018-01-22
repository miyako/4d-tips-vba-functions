# 4d-tips-vba-functions
Helper functions for Unicode support in VBA

``WebHelpers``

```vba
' in Function json_ParseString
' because SJIS converts antislash to yen

    ' Application.LanguageSettings not available on Mac?
    ' https://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.applicationclass.languagesettings(v=office.14).aspx
    Dim IsCountryCodeJapan As Boolean
    IsCountryCodeJapan = (81 = Application.International(xlCountryCode))

    If IsCountryCodeJapan And json_Char = "¥" Then
        json_Char = "\"
    End If
```

```vba
'AscW returns signed integer, which can be negative for 0x8000 and above
Public Function AscU(char As String) As Long
  AscU = VBA.CLng("&H0000" + (VBA.Hex(VBA.AscW(char))))
End Function
```

```vba
'Based on:
' WebHelpers v4.1.3
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web

'Modification:
' UTF-8 support for 2 and 3 byte sequences (BMP)

''
' Encode string for URLs
'
' See https://github.com/VBA-tools/VBA-Web/wiki/Url-Encoding for details
'
' References:
' - RFC 3986, https://tools.ietf.org/html/rfc3986
' - form-urlencoded encoding algorithm,
'   https://www.w3.org/TR/html5/forms.html#application/x-www-form-urlencoded-encoding-algorithm
' - RFC 6265 (Cookies), https://tools.ietf.org/html/rfc6265
'   Note: "%" is allowed in spec, but is currently excluded due to parsing issues
'
' @method UrlEncode
' @param {Variant} Text Text to encode
' @param {Boolean} [SpaceAsPlus = False] `%20` if `False` / `+` if `True`
'   DEPRECATED Use EncodingMode:=FormUrlEncoding
' @param {Boolean} [EncodeUnsafe = True] Encode characters that could be misunderstood within URLs.
'   (``SPACE, ", <, >, #, %, {, }, |, \, ^, ~, `, [, ]``)
'   DEPRECATED This was based on an outdated URI spec and has since been removed.
'     EncodingMode:=CookieUrlEncoding is the closest approximation of this behavior
' @param {UrlEncodingMode} [EncodingMode = StrictUrlEncoding]
' @return {String} Encoded string
''
Public Function UrlEncode(text As Variant, _
    Optional SpaceAsPlus As Boolean = False, Optional EncodeUnsafe As Boolean = True, _
    Optional EncodingMode As UrlEncodingMode = UrlEncodingMode.StrictUrlEncoding) As String

    If SpaceAsPlus = True Then
        LogWarning "SpaceAsPlus is deprecated and will be removed in VBA-Web v5. " & _
            "Use EncodingMode:=FormUrlEncoding instead", "WebHelpers.UrlEncode"
    End If
    If EncodeUnsafe = False Then
        LogWarning "EncodeUnsafe has been removed as it was based on an outdated url encoding specification. " & _
            "Use EncodingMode:=CookieUrlEncoding to approximate this behavior", "WebHelpers.UrlEncode"
    End If

    Dim web_UrlVal As String
    Dim web_StringLen As Long

    web_UrlVal = VBA.CStr(text)
    web_StringLen = VBA.Len(web_UrlVal)

    If web_StringLen > 0 Then
        Dim web_Result() As String
        Dim web_i As Long
        Dim web_CharCode As Long
        Dim web_Char As String
        Dim web_Space As String
        ReDim web_Result(web_StringLen)
        Dim b1 As Long
        Dim b2 As Long
        Dim b3 As Long

        ' StrictUrlEncoding - ALPHA / DIGIT / "-" / "." / "_" / "~"
        ' FormUrlEncoding   - ALPHA / DIGIT / "-" / "." / "_" / "*" / (space) -> "+"
        ' QueryUrlEncoding  - ALPHA / DIGIT / "-" / "." / "_"
        ' CookieUrlEncoding - strict / "!" / "#" / "$" / "&" / "'" / "(" / ")" / "*" / "+" /
        '   "/" / ":" / "<" / "=" / ">" / "?" / "@" / "[" / "]" / "^" / "`" / "{" / "|" / "}"
        ' PathUrlEncoding   - strict / "!" / "$" / "&" / "'" / "(" / ")" / "*" / "+" / "," / ";" / "=" / ":" / "@"

        ' Set space value
        If SpaceAsPlus Or EncodingMode = UrlEncodingMode.FormUrlEncoding Then
            web_Space = "+"
        Else
            web_Space = "%20"
        End If

        ' Loop through string characters
        For web_i = 1 To web_StringLen
            ' Get character and ascii code
            web_Char = VBA.Mid$(web_UrlVal, web_i, 1)
            web_CharCode = AscU(web_Char)

            Select Case web_CharCode
                Case 65 To 90, 97 To 122
                    ' ALPHA
                    web_Result(web_i) = web_Char
                Case 48 To 57
                    ' DIGIT
                    web_Result(web_i) = web_Char
                Case 45, 46, 95
                    ' "-" / "." / "_"
                    web_Result(web_i) = web_Char

                Case 32
                    ' (space)
                    ' FormUrlEncoding -> "+"
                    ' Else -> "%20"
                    web_Result(web_i) = web_Space

                Case 33, 36, 38, 39, 40, 41, 43, 58, 61, 64
                    ' "!" / "$" / "&" / "'" / "(" / ")" / "+" / ":" / "=" / "@"
                    ' PathUrlEncoding, CookieUrlEncoding -> Unencoded
                    ' Else -> Percent-encoded
                    If EncodingMode = UrlEncodingMode.PathUrlEncoding Or EncodingMode = UrlEncodingMode.CookieUrlEncoding Then
                        web_Result(web_i) = web_Char
                    Else
                        web_Result(web_i) = "%" & VBA.Hex(web_CharCode)
                    End If

                Case 35, 45, 46, 47, 60, 62, 63, 91, 93, 94, 95, 96, 123, 124, 125
                    ' "#" / "-" / "." / "/" / "<" / ">" / "?" / "[" / "]" / "^" / "_" / "`" / "{" / "|" / "}"
                    ' CookieUrlEncoding -> Unencoded
                    ' Else -> Percent-encoded
                    If EncodingMode = UrlEncodingMode.CookieUrlEncoding Then
                        web_Result(web_i) = web_Char
                    Else
                        web_Result(web_i) = "%" & VBA.Hex(web_CharCode)
                    End If

                Case 42
                    ' "*"
                    ' FormUrlEncoding, PathUrlEncoding, CookieUrlEncoding -> "*"
                    ' Else -> "%2A"
                    If EncodingMode = UrlEncodingMode.FormUrlEncoding _
                        Or EncodingMode = UrlEncodingMode.PathUrlEncoding _
                        Or EncodingMode = UrlEncodingMode.CookieUrlEncoding Then

                        web_Result(web_i) = web_Char
                    Else
                        web_Result(web_i) = "%" & VBA.Hex(web_CharCode)
                    End If

                Case 44, 59
                    ' "," / ";"
                    ' PathUrlEncoding -> Unencoded
                    ' Else -> Percent-encoded
                    If EncodingMode = UrlEncodingMode.PathUrlEncoding Then
                        web_Result(web_i) = web_Char
                    Else
                        web_Result(web_i) = "%" & VBA.Hex(web_CharCode)
                    End If

                Case 126
                    ' "~"
                    ' FormUrlEncoding, QueryUrlEncoding -> "%7E"
                    ' Else -> "~"
                    If EncodingMode = UrlEncodingMode.FormUrlEncoding Or EncodingMode = UrlEncodingMode.QueryUrlEncoding Then
                        web_Result(web_i) = "%7E"
                    Else
                        web_Result(web_i) = web_Char
                    End If

                Case 0 To 15
                    web_Result(web_i) = "%0" & VBA.Hex(web_CharCode)
                Case Else
                    web_Result(web_i) = "%" & VBA.Hex(web_CharCode)

                ' TODO For non-ASCII characters,
                '
                ' FormUrlEncoded:
                '
                ' Replace the character by a string consisting of a U+0026 AMPERSAND character (&), a "#" (U+0023) character,
                ' one or more ASCII digits representing the Unicode code point of the character in base ten, and finally a ";" (U+003B) character.
                '
                ' Else:
                '
                  If web_CharCode > &H7F And web_CharCode < &H800 Then
                    b1 = (web_CharCode And &HFFC0)
                    b1 = (web_CharCode / &H40)
                    b1 = b1 And &H1F
                    b1 = b1 + &HC0
                    b2 = (web_CharCode And &H3F)
                    b2 = b2 + &H80
                    web_Result(web_i) = "%" & Hex(b1) & "%" & Hex(b2)
                  ElseIf web_CharCode > &H7FF Then
                    b1 = (web_CharCode And &HF000)
                    b1 = (b1 / &H1000)
                    b1 = b1 And &HF
                    b1 = b1 + &HE0
                    b2 = (web_CharCode And &HFFC0)
                    b2 = (b2 / &H40)
                    b2 = (b2 And &H3F)
                    b2 = b2 + &H80
                    b3 = (web_CharCode And &H3F)
                    b3 = b3 + &H80
                    web_Result(web_i) = "%" & Hex(b1) & "%" & Hex(b2) & "%" & Hex(b3)
                End If
            End Select
        Next web_i
        UrlEncode = VBA.Join$(web_Result, "")
    End If
End Function
```

```vba
'Based on:
' WebHelpers v4.1.3
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web

'Modification:
' UTF-8 support for 2 and 3 byte sequences (BMP)

''
' Decode Url-encoded string.
'
' @method UrlDecode
' @param {String} Encoded Text to decode
' @param {Boolean} [PlusAsSpace = True] Decode plus as space
'   DEPRECATED Use EncodingMode:=FormUrlEncoding Or QueryUrlEncoding
' @param {UrlEncodingMode} [EncodingMode = StrictUrlEncoding]
' @return {String} Decoded string
''
Public Function UrlDecode(Encoded As String, _
    Optional PlusAsSpace As Boolean = True, _
    Optional EncodingMode As UrlEncodingMode = UrlEncodingMode.StrictUrlEncoding) As String

    Dim web_StringLen As Long
    web_StringLen = VBA.Len(Encoded)

    If web_StringLen > 0 Then
        Dim web_i As Long
        Dim web_Result As String
        Dim web_Temp As String
        Dim web_Code As Long
        Dim b1 As Long
        Dim b2 As Long
        Dim b3 As Long

        For web_i = 1 To web_StringLen
            web_Temp = VBA.Mid$(Encoded, web_i, 1)

            If web_Temp = "+" And _
                (PlusAsSpace _
                 Or EncodingMode = UrlEncodingMode.FormUrlEncoding _
                 Or EncodingMode = UrlEncodingMode.QueryUrlEncoding) Then

                web_Temp = " "
            ElseIf web_Temp = "%" And web_StringLen >= web_i + 2 Then
                web_Temp = VBA.Mid$(Encoded, web_i + 1, 2)
                web_Code = VBA.CInt("&H" & web_Temp)
                If web_Code <= &H7F Then
                web_Temp = VBA.ChrW(web_Code)
                web_i = web_i + 2
                ElseIf web_Code >= &HE0 And web_Code <= &HF00 Then
                  '3 bytes
                  If web_StringLen >= web_i + 8 Then
                    b1 = (web_Code Mod &HE0) * &H1000
                    web_Code = VBA.CInt("&H" & VBA.Mid$(Encoded, web_i + 4, 2))
                    b2 = (web_Code Mod &H80) * &H40
                    b3 = VBA.CInt("&H" & VBA.Mid$(Encoded, web_i + 7, 2))
                    b3 = b3 And &H3F
                    web_Temp = VBA.ChrW(b1 + b2 + b3)
                    web_i = web_i + 8
                  Else
                    web_i = web_StringLen
                  End If
                ElseIf web_Code >= &H7F And web_Code <= &HE0 Then
                  '2 bytes
                  If web_StringLen >= web_i + 5 Then
                    b1 = (web_Code Mod &HC0) * &H40
                    b2 = VBA.CInt("&H" & VBA.Mid$(Encoded, web_i + 4, 2))
                    b2 = b2 And &H3F
                    web_Temp = VBA.ChrW(b1 + b2)
                    web_i = web_i + 5
                  Else
                    web_i = web_StringLen
                  End If
                End If
            End If
            web_Result = web_Result & web_Temp
        Next web_i
        UrlDecode = web_Result
    End If
End Function
```
