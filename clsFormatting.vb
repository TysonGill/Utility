Imports System
Imports System.IO
Imports System.Text
Imports System.Xml
Imports System.Xml.Xsl
Imports System.IO.Compression
Imports System.Security.Cryptography

''' <summary>General library for formatting and parsing routines.</summary>
Public Class clsFormatting

    Dim rnd As New Random()

    ''' <summary>Prepare a variable for inclusion in a SQL string.</summary>
    ''' <param name="Fld">Text to prepare.</param>
    ''' <remarks>Use for formatting variables in SQL statements. This also filters against SQL injection attacks.</remarks>
    ''' <example>sql = "SELECT * FROM MyTable WHERE ID = " + fmt.q(sID)</example>
    ''' <returns>A string formatted for inclusion in a SQL query.</returns>
    Public Function q(ByVal Fld As Object) As String
        Try
            If IsDBNull(Fld) Then Return "NULL"
            If Fld Is Nothing Then Return "NULL"
            If VarType(Fld) = vbNull Then Return "NULL"
            Select Case VarType(Fld)
                Case VariantType.String
                    Fld = Replace(Fld, "\'", "")
                    Fld = Replace(Fld, "'", "''")
                    Return "'" + CType(Fld, String) + "'"
                Case VariantType.Date
                    Return "'" + CType(Fld, String) + "'"
                Case VariantType.Boolean
                    Return Val(Fld)
                Case Else
                    Return CType(Fld, String)
            End Select
        Catch
            Return Fld
        End Try
    End Function

    ''' <summary>Parse the name from an email address.</summary>
    ''' <param name="EmailAddress">A valid email address.</param>
    ''' <returns>The name part of the address.</returns>
    Function EmailName(ByVal EmailAddress As String) As String
        Dim intPos As Integer = InStr(EmailAddress, "@")
        If intPos = -1 Then Return EmailAddress
        Return EmailAddress.Substring(0, intPos - 1)
    End Function

    ''' <summary>Get the age of a person.</summary>
    ''' <param name="BirthDate">The birthdate of the person.</param>
    ''' <returns>The age in years.</returns>
    Function GetAge(ByVal BirthDate As Date) As Integer
        Return DateDiff(DateInterval.Month, BirthDate, Now) \ 12
    End Function

    ''' <summary>Get the date of the first day in a month.</summary>
    ''' <param name="dtDate">Any date in the month to return.</param>
    ''' <returns>The date of the first day of the month.</returns>
    Function GetFirstDayInMonth(Optional ByVal dtDate As Date = Nothing) As Date
        If dtDate = Nothing Then dtDate = Now
        Return DateAdd(DateInterval.Day, (dtDate.Day - 1) * -1, dtDate)
    End Function

    ''' <summary>Get the date of the last day in a month.</summary>
    ''' <param name="dtDate">Any date in the month to return.</param>
    ''' <returns>The date of the last day of the month.</returns>
    Function GetLastDayInMonth(Optional ByVal dtDate As Date = Nothing) As Date
        If dtDate = Nothing Then dtDate = Now
        Return DateAdd(DateInterval.Day, (Day(DateAdd(DateInterval.Month, 1, dtDate))) * -1, DateAdd(DateInterval.Month, 1, dtDate))
    End Function

    ''' <summary>Set the time for a date to the start of day.</summary>
    ''' <param name="dtDate">Any date for which to set the time.</param>
    ''' <returns>The same date at 12:00 am.</returns>
    Function GetStartOfDay(Optional ByVal dtDate As Date = Nothing) As DateTime
        If dtDate = Nothing Then dtDate = Now
        Return dtDate.ToShortDateString + " 12:00 AM"
    End Function

    ''' <summary>Set the time for a date to the end of day.</summary>
    ''' <param name="dtDate">Any date for which to set the time.</param>
    ''' <returns>The same date at 11:59 pm.</returns>
    Function GetEndOfDay(Optional ByVal dtDate As Date = Nothing) As DateTime
        If dtDate = Nothing Then dtDate = Now
        Return dtDate.ToShortDateString + " 11:59 PM"
    End Function

    ''' <summary>Get the text in a string between two fragments.</summary>
    ''' <param name="Full">The string to parse.</param>
    ''' <param name="FromPart">The start fragment.</param>
    ''' <param name="ToPart">The end fragment.</param>
    ''' <returns>Inner string.</returns>
    Public Function GetTextBetween(ByVal Full As String, ByVal FromPart As String, ByVal ToPart As String) As String
        Dim startpos As Integer = InStr(Full, FromPart, CompareMethod.Text)
        If startpos = 0 Then Return ""
        startpos += FromPart.Length
        Dim endpos As Integer = InStr(startpos, Full, ToPart, CompareMethod.Text)
        If endpos = 0 Then Return ""
        Return Microsoft.VisualBasic.Mid(Full, startpos, endpos - startpos).Trim
    End Function

    ''' <summary>Replace delimiter pairs.</summary>
    ''' <param name="Full">The string to parse.</param>
    ''' <param name="SearchLeft">The left delimiter to be replaced.</param>
    ''' <param name="SearchRight">The right delimiter to be replaced.</param>
    ''' <param name="RepLeft">The left delimiter to replace with.</param>
    ''' <param name="RepRight">The right delimiter to replace with.</param>
    ''' <returns>Modified string.</returns>
    Public Function ReplacePairs(ByVal Full As String, Optional ByVal SearchLeft As String = Chr(34), Optional ByVal SearchRight As String = Chr(34), Optional ByVal RepLeft As String = "[", Optional ByVal RepRight As String = "]") As String
        Dim RightPos As Integer = 0
        Dim LeftPos As Integer = InStr(1, Full, SearchLeft, CompareMethod.Text)
        Do While LeftPos > 0
            RightPos = InStr(LeftPos + 1, Full, SearchRight, CompareMethod.Text)
            If RightPos > 0 Then
                Mid(Full, LeftPos, 1) = RepLeft
                Mid(Full, RightPos, 1) = RepRight
                LeftPos = InStr(RightPos + 1, Full, SearchLeft, CompareMethod.Text)
            Else
                LeftPos = 0
            End If
        Loop
        Return Full
    End Function

    ''' <summary>Safely returns a boolean from a bit field.</summary>
    ''' <param name="BitField">Bit field to convert.</param>
    ''' <returns>A valid boolean.</returns>
    Public Function ToBoolean(ByVal BitField As Object) As Boolean
        If BitField Is DBNull.Value OrElse Val(BitField) = 0 Then Return False
        Return True
    End Function

    ''' <summary>Safely returns a string even if str is Nothing.</summary>
    ''' <param name="str">str to prepare.</param>
    ''' <returns>A valid string.</returns>
    Public Function ToStr(ByVal str As Object) As String
        If str Is Nothing Then Return ""
        Return str.ToString
    End Function

    ''' <summary>Prepare a URL string.</summary>
    ''' <param name="URL">URL to prepare.</param>
    ''' <remarks>Use to encode special characters for URL.</remarks>
    ''' <returns>A valid URL.</returns>
    Public Function ToURL(ByVal URL As String) As String
        If URL.Trim = "" Then Return ""
        Dim i As Integer
        Dim c As Integer
        ToURL = ""
        URL = Replace(URL, "%", "%25")
        For i = 1 To URL.Length
            c = Asc(Microsoft.VisualBasic.Mid(URL, i, 1))
            If (c <= 47 And c <> 37) OrElse (c >= 58 And c <= 64) OrElse (c >= 91 And c <= 96) OrElse (c >= 123) Then
                ToURL += "%" + Hex(c)
            Else
                ToURL += Chr(c)
            End If
        Next
    End Function

    ''' <summary>Prepare an XML string.</summary>
    ''' <param name="XML">XML string to prepare.</param>
    ''' <remarks>Use to encode special characters for XML.</remarks>
    ''' <returns>A valid XML string.</returns>
    Public Function ToXML(ByVal XML As String) As String
        ToXML = XML.ToString.Trim
        If ToXML = "" Then Exit Function
        ToXML = Replace(ToXML, "&", "&amp;")
        ToXML = Replace(ToXML, "<", "&lt;")
        ToXML = Replace(ToXML, ">", "&gt;")
        ToXML = Replace(ToXML, """", "&quot;")
        ToXML = Replace(ToXML, "'", "&apos;")
    End Function

    ''' <summary>Decode an XML string.</summary>
    ''' <param name="XML">XML string to decode.</param>
    ''' <remarks>Use to decode special characters in XML.</remarks>
    ''' <returns>A valid text string.</returns>
    Public Function FromXML(ByVal XML As String) As String
        FromXML = XML.ToString.Trim
        If FromXML = "" Then Exit Function
        FromXML = Replace(FromXML, "&amp;", "&")
        FromXML = Replace(FromXML, "&lt;", "<")
        FromXML = Replace(FromXML, "&gt;", ">")
        FromXML = Replace(FromXML, "&quot;", """")
        FromXML = Replace(FromXML, "&apos;", "'")
    End Function

    ''' <summary>Convert a SSN to standard format.</summary>
    ''' <param name="SSN">SSN to format.</param>
    ''' <returns>Formatted SSN.</returns>
    Public Function ToSSN(ByVal SSN As String) As String
        Dim s As String = ToNumericCharsOnly(SSN)
        If s.Length <> 9 Then Return SSN
        Return Microsoft.VisualBasic.Left(s, 3) + "-" + Microsoft.VisualBasic.Mid(s, 4, 2) + "-" + Microsoft.VisualBasic.Right(s, 4)
    End Function

    ''' <summary>Verify that a string is formattable to an SSN.</summary>
    ''' <param name="SSN">Phone number to verify.</param>
    ''' <returns>True if the string is a 9 digit number.</returns>
    Public Function IsSSN(ByVal SSN As String) As Boolean
        Dim NumChars As Integer = fmt.ToNumericCharsOnly(SSN).Length
        Return (NumChars = 9)
    End Function

    ''' <summary>Create a phone number string from component parts.</summary>
    ''' <param name="ph1">Area code.</param>
    ''' <param name="ph2">Exchange.</param>
    ''' <param name="ph3">Number.</param>
    ''' <returns>Formatted phone number.</returns>
    Public Function AssemblePhone(ByVal ph1 As String, ByVal ph2 As String, ByVal ph3 As String) As String
        ph1 = ph1.Trim
        ph2 = ph2.Trim
        ph3 = ph3.Trim
        If ph1 = "" And ph2 = "" And ph3 = "" Then Return ""
        If ph1 = "000" AndAlso ph2 = "000" AndAlso ph3 = "0000" Then Return ""
        If ph1 <> "" Then
            Return "(" + ph1 + ") " + ph2 + "-" + ph3
        Else
            Return ph2 + "-" + ph3
        End If
    End Function

    ''' <summary>Check if two street addresses match.</summary>
    ''' <param name="addr1">The first street address.</param>
    ''' <param name="addr2">The second street address.</param>
    ''' <remarks>Should be street address only, without city, state, and zip.</remarks>
    ''' <returns>True if the addresses are the same.</returns>
    Public Function IsAddressSame(ByVal addr1 As String, ByVal addr2 As String) As Boolean

        ' Split up the address elements, putting the one with fewer elements into a1
        Dim a1() As String
        Dim a2() As String
        If Split(addr1, " ").Length <= Split(addr2, " ").Length Then
            a1 = Split(addr1.ToUpper, " ")
            a2 = Split(addr2.ToUpper, " ")
        Else
            a1 = Split(addr2.ToUpper, " ")
            a2 = Split(addr1.ToUpper, " ")
        End If

        ' Check that each number or string element longer than 2 characters in the short address is found in the longer one
        For Each e1 As String In a1
            If IsNumeric(e1) OrElse e1.Length > 2 Then
                For Each e2 As String In a2
                    If e2.Length <= 2 Then GoTo nextelement
                    If e1 = e2 Then GoTo NextElement
                Next
                Return False
            End If
NextElement:
        Next
        Return True

    End Function

    ''' <summary>Convert a Zip Code to standard format.</summary>
    ''' <param name="zip">Zip to format.</param>
    ''' <remarks>Handles Zip+4 format.</remarks>
    ''' <returns>Formatted zip code.</returns>
    Public Function ToZip(ByVal zip As String) As String
        If Not IsZip(zip) Then Return zip
        Dim s As String = ToNumericCharsOnly(zip)
        If s.Length <> 9 Then Return s
        Return Microsoft.VisualBasic.Left(s, 5) + "-" + Microsoft.VisualBasic.Right(s, 4)
    End Function

    ''' <summary>Validates a Zip Code to standard format.</summary>
    ''' <param name="zip">Zip to validate.</param>
    ''' <remarks>Validates that the code can be formatted as a zip code. Use ToZip to ensure proper formatting.</remarks>
    ''' <returns>True if valid (zip has 5 or 9 numeric characters with an optional hyphen).</returns>
    Public Function IsZip(ByVal zip As String) As Boolean
        Dim NumChars As Integer = fmt.ToNumericCharsOnly(zip).Length
        Return IsNumeric(NumChars = 5 OrElse NumChars = 9)
    End Function

    ''' <summary>Convert a phone number to standard format.</summary>
    ''' <param name="Phone">Phone number to format.</param>
    ''' <returns>Formatted phone Number.</returns>
    Public Function ToPhone(ByVal Phone As String) As String
        Dim s As String = ToNumericCharsOnly(Phone)
        If s.Length <> 7 AndAlso s.Length <> 10 Then Return Phone.ToString.Trim
        If s.Length = 7 Then Return Microsoft.VisualBasic.Left(s, 3) + "-" + Microsoft.VisualBasic.Right(s, 4)
        Return "(" + Microsoft.VisualBasic.Left(s, 3) + ") " + Mid(s, 4, 3) + "-" + Microsoft.VisualBasic.Right(s, 4)
    End Function

    ''' <summary>Verify that a string is formattable to a 10 digit phone number.</summary>
    ''' <param name="Phone">Phone number to verify.</param>
    ''' <returns>True if the string is a 10 digit number.</returns>
    Public Function IsPhone(ByVal Phone As String, Optional ByVal Allow7 As Boolean = False) As Boolean
        Dim NumChars As Integer = fmt.ToNumericCharsOnly(Phone).Length
        Return (NumChars = 10 OrElse (Allow7 AndAlso NumChars = 7))
    End Function

    ''' <summary>Strips a string of any non-numeric characters.</summary>
    ''' <param name="Text">Text string to format.</param>
    ''' <returns>Stripped text string.</returns>
    Public Function ToNumericCharsOnly(ByVal Text As String, Optional ByVal DecimalAllowed As Boolean = False, Optional ByVal NegativeAllowed As Boolean = False) As String
        Dim i As Integer
        Dim c As String
        ToNumericCharsOnly = ""
        For i = 1 To Text.Length
            c = Mid(Text, i, 1)
            If IsNumeric(c) Or (DecimalAllowed And c = ".") Or (NegativeAllowed And c = "-") Then ToNumericCharsOnly += c
        Next
    End Function

    ''' <summary>Strips a string of any white space.</summary>
    ''' <param name="Text">Text string to format.</param>
    ''' <returns>Stripped text string.</returns>
    Public Function StripWhiteSpace(ByVal Text) As String
        Dim rex As New System.Text.RegularExpressions.Regex("\s+")
        Return rex.Replace(Text, " ")
    End Function

    ''' <summary>Truncates the decimal places.</summary>
    ''' <param name="Number">Number object to format.</param>
    ''' <returns>A number string.</returns>
    Public Function ToDecimalPlaces(ByVal Number As Object, Optional ByVal DecimalPlaces As Integer = 0) As String
        Dim NumText As String = ToNumericCharsOnly(Number.ToString, True, True)
        Dim FormatString As String = "{0:0." + New String("0", DecimalPlaces) + "}"
        ToDecimalPlaces = String.Format(FormatString, Decimal.Parse(NumText))
    End Function

    ''' <summary>Create an address string from component parts.</summary>
    ''' <param name="ad1">Address Line 1.</param>
    ''' <param name="ad2">Address Line 2.</param>
    ''' <param name="city">City.</param>
    ''' <param name="state">State abbreviation.</param>
    ''' <param name="zip">Zip Code.</param>
    ''' <returns>A formatted address.</returns>
    Public Function ToAddress(ByVal ad1 As String, ByVal ad2 As String, ByVal city As String, ByVal state As String, ByVal zip As String) As String
        ad1 = ad1.Trim
        ad2 = ad2.Trim
        city = city.Trim
        state = state.ToUpper.Trim
        zip = zip.Trim
        If ad1 = "" And ad2 = "" Then
            ToAddress = city
        ElseIf ad1 <> "" And ad2 = "" Then
            ToAddress = ad1 + ", " + city
        ElseIf ad1 = "" And ad2 <> "" Then
            ToAddress = ad2 + ", " + city
        Else
            ToAddress = ad1 + " " + ad2 + ", " + city
        End If
        ToAddress = ToTitle(ToAddress)
        ToAddress += " " + state + " " + zip
        Return StripWhiteSpace(ToAddress.Trim)
    End Function

    ''' <summary>Create a full name string from component parts.</summary>
    ''' <param name="FirstName">First Name.</param>
    ''' <param name="MiddleName">Middle Name.</param>
    ''' <param name="LastName">Last Name.</param>
    ''' <returns>A full name string.</returns>
    Public Function ToFullName(ByVal FirstName As String, ByVal MiddleName As String, ByVal LastName As String, Optional ByVal Title As String = "") As String
        FirstName = FirstName.Trim
        MiddleName = MiddleName.Trim
        LastName = LastName.Trim
        Title = Title.Trim
        If FirstName = "" And MiddleName = "" And LastName = "" Then Return ""
        If MiddleName <> "" Then
            If MiddleName.Length = 1 Then MiddleName += "."
            ToFullName = ToTitle(FirstName + " " + MiddleName + " " + LastName)
        Else
            ToFullName = ToTitle(FirstName + " " + LastName)
        End If
        If Title <> "" Then
            ToFullName += ", " + Title
        End If
    End Function

    ''' <summary>Convert a date to MHS format.</summary>
    ''' <param name="d">The DateTime to convert.</param>
    ''' <remarks>If not provided, d defaults to Now.</remarks>
    ''' <returns>An MHS format date.</returns>
    Public Function ToMhsDate(Optional ByVal d As DateTime = Nothing) As String
        ' MHS Dates are in format CYYMMDD where C of 0 = 1800, C of 1 = 1900 etc
        If d = Nothing Then d = DateTime.Now
        ToMhsDate = (Microsoft.VisualBasic.Left(d.Year.ToString, 2) - 18).ToString + Microsoft.VisualBasic.Mid(d.Year.ToString, 3, 2) + d.Month.ToString("0#") + d.Day.ToString("0#")
    End Function

    ''' <summary>Make sure a time string is formatted properly.</summary>
    ''' <param name="TimeString">The time string to convert.</param>
    ''' <returns>A properly formatted time string.</returns>
    Public Function ToTime(ByVal TimeString As String) As String
        If TimeString.Trim = "" Then Return ""
        ToTime = TimeString.Trim.ToUpper
        ToTime = Replace(ToTime, ".", ":") 'some people put in a . instead of a :
        ToTime = Replace(ToTime, " ", "") ' remove any extra spaces
        If ToTime.EndsWith("P") Then ToTime += "M" 'some people leave off the M
        If ToTime.EndsWith("A") Then ToTime += "M" 'some people leave off the M
        ToTime = Replace(ToTime, "AM", " AM") 'make sure there is a space before the AM
        ToTime = Replace(ToTime, "PM", " PM") 'make sure there is a space before the PM
    End Function

    ''' <summary>Convert a military time to standard am/pm format.</summary>
    ''' <param name="MilitaryTime">The miltiary time string to convert.</param>
    ''' <returns>A properly formatted standard time string.</returns>
    Public Function ToStdTime(ByVal MilitaryTime As String) As String
        If MilitaryTime.Trim = "" Then Return ""
        If Val(MilitaryTime) < 0 Or Val(MilitaryTime) > 2400 Then Return MilitaryTime
        Dim h As Integer = Int(Val(MilitaryTime) / 100)
        Dim m As Integer = Val(MilitaryTime) - (h * 100)
        Dim p As String
        If h = 0 Then
            h += 12
            p = " am"
        ElseIf h < 12 Then
            p = " am"
        ElseIf h = 12 Then
            p = " pm"
        Else
            h -= 12
            p = " pm"
        End If
        ToStdTime = h.ToString + ":" + String.Format("{0:00}", m) + p
    End Function

    ''' <summary>Return the friendly form for a gender abbreviation.</summary>
    ''' <param name="sText">The gender abbreviation to convert.</param>
    ''' <returns>The full gender description string.</returns>
    Public Function ToGender(ByVal sText As String) As String
        Select Case sText.Trim.ToUpper
            Case "M"
                Return "Male"
            Case "F"
                Return "Female"
            Case "X"
                Return ""
            Case Else
                Return sText
        End Select
    End Function

    ''' <summary>Extract a zip code from an address.</summary>
    ''' <param name="AddrString">The address string containing a zip code.</param>
    ''' <returns>The parsed zip code, if any.</returns>
    Public Function ZipFromAddr(ByVal AddrString As String) As String
        AddrString = AddrString.Trim
        ZipFromAddr = ""

        ' See if we have a zip only already
        If AddrString.Length = 5 Then
            If IsNumeric(AddrString) Then
                ZipFromAddr = AddrString
                Exit Function
            End If
        ElseIf AddrString.Length = 9 Then
            If IsNumeric(AddrString) Then
                ZipFromAddr = Microsoft.VisualBasic.Left(AddrString, 5)
                Exit Function
            End If
        ElseIf AddrString.Length = 10 Then
            If Mid(AddrString, 6, 1) = "-" Then
                If IsNumeric(Microsoft.VisualBasic.Left(AddrString, 5)) Then
                    ZipFromAddr = Microsoft.VisualBasic.Left(AddrString, 5)
                    Exit Function
                End If
            End If
        End If

        ' Search the end of the string to find the zip code
        Dim c As String
        For i As Integer = AddrString.Length To 1 Step -1
            c = Microsoft.VisualBasic.Mid(AddrString, i, 1)
            If IsNumeric(c) OrElse c = "-" Then
                ZipFromAddr = c + ZipFromAddr
            Else
                Exit For
            End If
        Next
        If ZipFromAddr = "" Then Exit Function
        ZipFromAddr = Replace(ZipFromAddr, "-", "")
        If ZipFromAddr.Length = 5 Then Exit Function
        If ZipFromAddr.Length = 9 Then Return Left(ZipFromAddr, 5)
        Return ""
    End Function

    ''' <summary>Put double quotes around a string.</summary>
    ''' <param name="strText">String to quote.</param>
    Public Function Quote(ByVal strText As String) As String
        Return Chr(34) & strText & Chr(34)
    End Function

    ''' <summary>Scramble characters to protect the text.</summary>
    ''' <param name="RawText">String to scramble.</param>
    Public Function Scramble(ByVal RawText As String) As String
        Scramble = ""
        Dim RawChar As Char
        For i As Integer = 1 To RawText.Length
            RawChar = Mid(RawText, i, 1)
            If Char.IsNumber(RawChar) Then
                Scramble += Chr(rnd.Next(Asc("0"), Asc("9") + 1))
            ElseIf Char.IsLower(RawChar) Then
                Scramble += Chr(rnd.Next(Asc("a"), Asc("z") + 1))
            ElseIf Char.IsUpper(RawChar) Then
                Scramble += Chr(rnd.Next(Asc("A"), Asc("Z") + 1))
            Else
                Scramble += RawChar
            End If
        Next
    End Function


    ''' <summary>Append a string with another string if it is not already there.</summary>
    ''' <param name="sText">The text string to be appended.</param>
    ''' <param name="sAppend">The text string to append.</param>
    ''' <remarks>Comparison is case insensitive.</remarks>
    ''' <example>sConn = AppendIfNeeded(sConn, ";")</example>
    Public Function AppendIfNeeded(ByVal sText As String, ByVal sAppend As String) As String
        If Microsoft.VisualBasic.Right(sText, sAppend.Length).ToUpper <> sAppend.ToUpper Then
            Return sText + sAppend
        Else
            Return sText
        End If
    End Function

    ''' <summary>Replace the first occurrence only of a substring.</summary>
    ''' <param name="sText">The text string to be appended.</param>
    ''' <param name="sSearch">The text string to search for.</param>
    ''' <param name="sReplace">The text string to replace it with.</param>
    ''' <remarks>If sSearch is not found, returns original string.</remarks>
    Public Function ReplaceFirst(ByVal sText As String, ByVal sSearch As String, ByVal sReplace As String)
        Dim pos As Integer
        pos = sText.IndexOf(sSearch)
        If pos = -1 Then Return sText
        sText = sText.Remove(pos, sSearch.Length)
        sText = sText.Insert(pos, sReplace)
        Return sText
    End Function

    ''' <summary>Convert a string to Title Case.</summary>
    ''' <param name="sText">The text to convert.</param>
    Public Function ToTitle(ByVal sText As String) As String
        ToTitle = ""
        sText = sText.Trim
        If sText = "" Then Exit Function
        ToTitle = Microsoft.VisualBasic.Left(sText, 1).ToUpper
        Dim i As Integer
        Dim c As Char
        For i = 2 To sText.Length
            c = sText.Chars(i - 2)
            If Not Char.IsLetter(c) AndAlso c <> "'" Then
                ToTitle += Microsoft.VisualBasic.Mid(sText, i, 1).ToUpper
            Else
                ToTitle += Microsoft.VisualBasic.Mid(sText, i, 1).ToLower
            End If
        Next
    End Function

    ''' <summary>Convert a string to currency.</summary>
    ''' <param name="sText">The text to convert.</param>
    Public Function ToCurrency(ByVal sText As String) As String
        If sText.Trim = "" Then Return ""
        ToCurrency = ""
        sText = ToNumericCharsOnly(sText, True, True)
        Dim amt As Double = Double.Parse(Replace(sText, "$", "").ToString)
        ToCurrency = "$" + String.Format("{0:0.00}", amt)
    End Function

    ''' <summary>Convert a string to percent.</summary>
    ''' <param name="sText">The text to convert.</param>
    Public Function ToPercent(ByVal sText As String) As String
        If sText.Trim = "" Then Return ""
        ToPercent = ""
        sText = ToNumericCharsOnly(sText, True, True)
        Dim amt As Double = Double.Parse(Replace(sText, "%", "").ToString)
        ToPercent = String.Format("{0:0.00}", amt) + "%"
    End Function

    ''' <summary>Compress a string.</summary>
    ''' <param name="OriginalText">The text to compress.</param>
    ''' <returns>The compressed string.</returns>
    ''' <remarks>Might increase the length of very short strings.</remarks>
    ''' <seealso cref="Utility.clsFormatting.Decompress"/>
    Public Function Compress(ByVal OriginalText As String) As String
        Dim buffer() As Byte = Encoding.UTF8.GetBytes(OriginalText)
        Dim ms As New MemoryStream()
        Dim zipStream As New GZipStream(ms, CompressionMode.Compress, True)
        zipStream.Write(buffer, 0, buffer.Length)
        zipStream.Flush()
        zipStream.Close()
        Return Convert.ToBase64String(ms.ToArray)
    End Function

    ''' <summary>Decompress a string.</summary>
    ''' <param name="CompressedText">The compressed text string to decompress.</param>
    ''' <returns>The uncompressed string.</returns>
    ''' <seealso cref="Utility.clsFormatting.Compress"/>
    Public Function Decompress(ByVal CompressedText As String) As String
        Dim compressed As Byte() = Convert.FromBase64String(CompressedText)
        Dim lastFour(3) As Byte
        Array.Copy(compressed, compressed.Length - 4, lastFour, 0, 4)
        Dim bufferLength As Integer = BitConverter.ToInt32(lastFour, 0)
        Dim buffer(bufferLength - 1) As Byte
        Dim ms As New MemoryStream(compressed)
        Dim decompressedStream As New GZipStream(ms, CompressionMode.Decompress, True)
        decompressedStream.Read(buffer, 0, bufferLength)
        Return Encoding.UTF8.GetString(buffer)
    End Function

#Region "XML"

    ''' <summary>Get a node attribute.</summary>
    ''' <param name="n">The XML Node containing the attribute.</param>
    ''' <param name="Attribute">The name of the attribute to get.</param>
    ''' <param name="DefaultValue">An optional default value if no attribute is found.</param>
    ''' <returns>The value of the specified attribute.</returns>
    Public Function GetAttribute(ByVal n As XmlNode, ByVal Attribute As String, Optional ByVal DefaultValue As Object = "") As Object
        If n Is Nothing OrElse n.Attributes(Attribute) Is Nothing Then Return DefaultValue
        Return fmt.FromXML(n.Attributes(Attribute).Value).Trim
    End Function

    ''' <summary>Set a node attribute.</summary>
    ''' <param name="n">The XML Node containing the attribute.</param>
    ''' <param name="Attribute">The name of the attribute to set.</param>
    ''' <param name="Value">The value of the attribute to set.</param>
    Public Sub SetAttribute(ByVal n As XmlNode, ByVal Attribute As String, ByVal Value As String)
        If n Is Nothing Then Exit Sub
        If n.Attributes(Attribute) Is Nothing Then
            Dim attr As XmlAttribute
            attr = n.OwnerDocument.CreateAttribute(Attribute)
            n.Attributes.Append(attr)
        End If
        n.Attributes(Attribute).Value = fmt.ToXML(Value.Trim)
    End Sub

    ''' <summary>Apply an XSLT transform string to an HTML/XML string.</summary>
    ''' <param name="document">The XML or HTML document to transform.</param>
    ''' <param name="stylesheet">The XSLT stylesheet to apply.</param>
    ''' <returns>The transformed document.</returns>
    Public Function ApplyXSLT(ByVal document As String, ByVal stylesheet As String) As String

        ' Create a TextReader object for the document
        Dim xml As New XmlTextReader(document, XmlNodeType.Document, Nothing)

        ' Create a TextReader object for the stylesheet
        Dim ss As New XmlTextReader(stylesheet, XmlNodeType.Document, Nothing)

        ' Create a TextWriter for the transformed document
        Dim writer As New StringWriter()
        Dim xmlwriter As New XmlTextWriter(writer)

        ' Apply the stylesheet
        Dim xslt As New XslCompiledTransform
        xslt.Load(ss)
        xslt.Transform(xml, xmlwriter)
        xmlwriter.Close()

        ' Return the new html
        Return writer.ToString

    End Function

    ''' <summary>Return the factorial of a given number.</summary>
    ''' <param name="n">The number to compute.</param>
    ''' <returns>The factorial of n.</returns>
    Public Function Factorial(ByVal n As Integer) As Long
        If n = 0 Then Return 1
        Factorial = 1
        For i As Integer = 2 To n
            Factorial *= i
        Next
    End Function

    ''' <summary>Determines if a number is a square of an integer value.</summary>
    ''' <param name="PossibleSquare">The number to check.</param>
    ''' <returns>True if the value provided has an integer square root.</returns>
    Public Function IsSquare(ByVal PossibleSquare As Double) As Boolean
        If PossibleSquare < 0 Then Return False
        Dim sqrt As Double = Math.Sqrt(PossibleSquare)
        Return ((sqrt - Math.Floor(sqrt)) = 0)
    End Function

#End Region

#Region "Encryption"

    ''' <summary>Encrypt or Unencrypt a text string.</summary>
    ''' <example>
    ''' Dim enc As New clsFormatting.Encrypter("this is the optional encryption key")
    ''' Dim b() As Byte = enc.Encrypt("Please encrypt this important secret message.")
    ''' Dim s As Object = enc.Decrypt(b) 'Returns Nothing if there is an error
    ''' Console.WriteLine(s.ToString)
    ''' </example>
    Public Class Encrypter

        Public EncryptKey As String = ""

        Private DefaultKey As String = "ki2d7nvqI]G4*vmaW@k?u73h^he2l:%x"
        Private key() As Byte = {6, 243, 231, 74, 9, 222, 146, 120, 43, 86, 23, 95, 184, 201, 137, 111, 55, 141, 74, 69, 194, 32, 74, 150}
        Private iv() As Byte = {43, 67, 111, 63, 212, 155, 76, 95}

        Public Sub New(Optional ByVal EncryptionKey As String = "")
            EncryptKey = EncryptionKey
        End Sub

        ''' <summary>Encrypt a string.</summary>
        ''' <returns>An encrypted byte array.</returns>
        ''' <seealso cref="Utility.clsFormatting.Encrypter.Decrypt"/>
        Public Function Encrypt(ByVal plainText As String) As Byte()

            ' Convert the EncryptKey into the key() and iv() arrays needed for encryption
            CreateKeys()

            ' Declare a UTF8Encoding object so we may use the GetByte 
            ' method to transform the plainText into a Byte array. 
            Dim utf8encoder As UTF8Encoding = New UTF8Encoding()
            Dim inputInBytes() As Byte = utf8encoder.GetBytes(plainText)

            ' Create a new TripleDES service provider 
            Dim tdesProvider As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider()

            ' The ICryptTransform interface uses the TripleDES 
            ' crypt provider along with encryption key and init vector 
            ' information 
            Dim cryptoTransform As ICryptoTransform = tdesProvider.CreateEncryptor(Me.key, Me.iv)

            ' All cryptographic functions need a stream to output the 
            ' encrypted information. Here we declare a memory stream 
            ' for this purpose. 
            Dim encryptedStream As MemoryStream = New MemoryStream()
            Dim cryptStream As CryptoStream = New CryptoStream(encryptedStream, cryptoTransform, CryptoStreamMode.Write)

            ' Write the encrypted information to the stream. Flush the information 
            ' when done to ensure everything is out of the buffer. 
            cryptStream.Write(inputInBytes, 0, inputInBytes.Length)
            cryptStream.FlushFinalBlock()
            encryptedStream.Position = 0

            ' Read the stream back into a Byte array and return it to the calling 
            ' method. 
            Dim result(encryptedStream.Length - 1) As Byte
            encryptedStream.Read(result, 0, encryptedStream.Length)
            cryptStream.Close()
            Return result
        End Function

        ''' <summary>Decrypt a byte array.</summary>
        ''' <returns>An unencrypted string.</returns>
        ''' <seealso cref="Utility.clsFormatting.Encrypter.Encrypt"/>
        Public Function Decrypt(ByVal inputInBytes() As Byte) As String

            Try

                ' Convert the EncryptKey into the key() and iv() arrays needed for encryption
                CreateKeys()

                ' UTFEncoding is used to transform the decrypted Byte Array 
                ' information back into a string. 
                Dim utf8encoder As UTF8Encoding = New UTF8Encoding()
                Dim tdesProvider As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider()

                ' As before we must provide the encryption/decryption key along with 
                ' the init vector. 
                Dim cryptoTransform As ICryptoTransform = tdesProvider.CreateDecryptor(Me.key, Me.iv)

                ' Provide a memory stream to decrypt information into 
                Dim decryptedStream As MemoryStream = New MemoryStream()
                Dim cryptStream As CryptoStream = New CryptoStream(decryptedStream, cryptoTransform, CryptoStreamMode.Write)
                cryptStream.Write(inputInBytes, 0, inputInBytes.Length)
                cryptStream.FlushFinalBlock()
                decryptedStream.Position = 0

                ' Read the memory stream and convert it back into a string 
                Dim result(decryptedStream.Length - 1) As Byte
                decryptedStream.Read(result, 0, decryptedStream.Length)
                cryptStream.Close()
                Dim myutf As UTF8Encoding = New UTF8Encoding()
                Return myutf.GetString(result)

            Catch ex As Exception
                Return Nothing
            End Try

        End Function

        ' Convert the EncryptKey into the key() and iv() arrays needed for encryption
        Private Sub CreateKeys()

            Dim s As String = EncryptKey + DefaultKey
            For i As Integer = 0 To 23
                key(i) = Asc(s.Chars(i + 1))
            Next
            For i As Integer = 0 To 7
                iv(i) = Asc(s.Chars(i + 24))
            Next

        End Sub
    End Class

#End Region

#Region "Sorted Lists"

    ''' <summary>Creates a sorted list of unique values with counts.</summary>
    ''' <param name="list">A SortedList to which to add a new value.</param>
    ''' <param name="key">The value of the list item.</param>
    ''' <remarks>The value passed will be the key and the value will hold the count.</remarks>
    Public Sub AddToSortedCounterList(ByVal list As SortedList, ByVal key As Object)

        ' Add or update list
        If list.ContainsKey(key) Then
            ' Increment the counter held in the value
            Dim index As Integer = list.IndexOfKey(key)
            list.SetByIndex(index, list.GetByIndex(index) + 1)
        Else
            ' Add a new item with a count of one
            list.Add(key, 1)
            Exit Sub
        End If

    End Sub

#End Region

#Region "Permutations"

    ''' <summary>Return a collection of string permutations keys with duplicate count values.</summary>
    ''' <param name="input">The string to be processed</param>
    ''' <returns>A SortedList of all unique key values along with their counts</returns>
    Friend Function GetPermutations(ByVal input As String) As SortedList

        ' Initialize the returned list and check for standard conditions
        Dim PermList As New SortedList
        If input Is Nothing OrElse input.Length = 0 Then Return PermList
        If input.Length = 1 Then
            AddToSortedCounterList(PermList, input)
            Return PermList
        End If

        ' Recursively add all unique permutations to PermList
        GetPermutations(PermList, input, "")
        Return PermList

    End Function

    Private Sub GetPermutations(ByVal PermList As SortedList, ByVal input As String, ByVal working As String)

        If input.Length = 0 Then
            AddToSortedCounterList(PermList, working)
            Exit Sub
        End If

        For i As Integer = 0 To input.Length - 1
            Dim remaining As String = input.Substring(0, i) + input.Substring(i + 1)
            GetPermutations(PermList, remaining, working + input(i))
        Next

    End Sub

#End Region

    ''' <summary>Determines if a string is included in a delimted string.</summary>
    ''' <param name="DelimitedString">The delimited string to search.</param>
    ''' <param name="Delimiter">The string delimiter.</param>
    ''' <param name="SearchString">The string to search for.</param>
    ''' <returns>True if the string is found.</returns>
    Public Function IsInString(ByVal DelimitedString As String, ByVal Delimiter As String, ByVal SearchString As String) As Boolean
        Dim StringArray() As String = Split(DelimitedString, Delimiter)
        Return IsInArray(StringArray, SearchString)
    End Function

    ''' <summary>Determines if a string is included in a string array.</summary>
    ''' <param name="StringArray">The string array to search.</param>
    ''' <param name="SearchString">The string to search for.</param>
    ''' <returns>True if the string is found.</returns>
    Public Function IsInArray(ByVal StringArray() As String, ByVal SearchString As String) As Boolean
        SearchString = SearchString.Trim.ToUpper
        For i As Integer = 0 To UBound(StringArray)
            If SearchString = StringArray(i).Trim.ToUpper Then Return True
        Next
        Return False
    End Function

    ''' <summary>Determines if a number is included in a delimted string.</summary>
    ''' <param name="DelimitedString">The delimited string to search.</param>
    ''' <param name="Delimiter">The string delimiter.</param>
    ''' <param name="SearchString">The number to search for.</param>
    ''' <returns>True if the number is found.</returns>
    Public Function IsInStringNumeric(ByVal DelimitedString As String, ByVal Delimiter As String, ByVal SearchString As String) As Boolean
        Dim StringArray() As String = Split(DelimitedString, Delimiter)
        Return IsInArrayNumeric(StringArray, SearchString)
    End Function

    ''' <summary>Determines if a number is included in a string array.</summary>
    ''' <param name="StringArray">The string array to search.</param>
    ''' <param name="SearchString">The number to search for.</param>
    ''' <returns>True if the number is found.</returns>
    Public Function IsInArrayNumeric(ByVal StringArray() As String, ByVal SearchString As String) As Boolean
        SearchString = SearchString.Trim.ToUpper
        For i As Integer = 0 To UBound(StringArray)
            If Val(SearchString) = Val(StringArray(i)) Then Return True
        Next
        Return False
    End Function

    ''' <summary>Parses a VB file and returns a string of all public functions with their XML descriptions.</summary>
    ''' <param name="filename">The VB source code file to parse.</param>
    ''' <returns>A formatted string.</returns>
    Public Function DocumentVB(ByVal filename As String) As String

        Dim rtb As New Windows.Forms.RichTextBox
        Dim NormalFont As Drawing.Font = rtb.Font
        Dim BoldFont As Drawing.Font = New Drawing.Font(rtb.Font, Drawing.FontStyle.Bold)
        rtb.Text = ""

        ' Load the VB module
        Dim code() As String = IO.File.ReadAllLines(filename)

        ' Create an array to map the line types
        Dim linetype(code.Length) As String

        ' Find all public functions
        For i As Integer = 0 To code.Length - 1
            If code(i).Trim.StartsWith("'''") Then
                linetype(i) = "XML"
            ElseIf code(i).Trim.StartsWith("Public") Then
                linetype(i) = "PUB"
            Else
                linetype(i) = ""
            End If
        Next

        ' Create the doc
        For i As Integer = 0 To code.Length - 1
            If linetype(i) = "PUB" Then

                ' Find all code comments preceeding the function
                Dim comment As String = ""
                For j As Integer = i - 1 To 0 Step -1
                    If linetype(j) <> "XML" Then Exit For
                    comment = code(j) + vbCrLf + comment
                Next

                ' Add the function call
                rtb.SelectionStart = rtb.Text.Length
                rtb.SelectionFont = NormalFont
                rtb.AppendText(comment)
                rtb.SelectionFont = BoldFont
                rtb.AppendText(code(i) + vbCrLf + vbCrLf)

            End If
        Next

        Return rtb.Rtf

    End Function

End Class
