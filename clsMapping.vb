Imports System.Net
Imports System.IO
Imports System.XML

''' <summary>General library for mapping routines.</summary>
Public Class clsMapping

    ''' <summary>Get the distance between two locations.</summary>
    ''' <param name="FromLoc">Starting location, address and/or zip code.</param>
    ''' <param name="ToLoc">Ending location, address and/or zip code.</param>
    ''' <returns>Distance in miles.</returns>
    Public Function GetDistance(ByVal FromLoc As String, ByVal ToLoc As String) As Double
        Try
            If FromLoc.Trim = "" OrElse ToLoc.Trim = "" Then Return 0
            Dim lat1 As String = ""
            Dim lon1 As String = ""
            Dim lat2 As String = ""
            Dim lon2 As String = ""
            Dim numlocs As Integer
            numlocs = GetLadLon(FromLoc, lat1, lon1)
            If numlocs <> 1 Then Return 0
            numlocs = GetLadLon(ToLoc, lat2, lon2)
            If numlocs <> 1 Then Return 0
            Return CalcDistance(lat1, lon1, lat2, lon2)
        Catch ex As Exception
            ReportError(ex)
        End Try
    End Function

    ''' <summary>Display Google map location in the browser.</summary>
    ''' <param name="Location">Location to display.</param>
    Public Sub ShowMap(ByVal Location As String, Optional ByVal zoom As Integer = 0)
        Dim url As String
        'url = http://maps.google.com/?q=2233+s+columbus+blvd+philadelphia&z=1
        'url = http://maps.google.com/maps?q=from+2233+s+columbus+blvd+philadelphia+to+1600+pennsylvania
        'url = "http://maps.google.com/?sll=39.918,-75.136"
        url = "http://maps.google.com/?q=" & fmt.ToURL(Location)
        If zoom > 0 Then url += "&z=" + zoom.ToString
        ui.OpenFile(url)
    End Sub

    ''' <summary>Display Google map directions in the browser.</summary>
    ''' <param name="FromLoc">Starting location.</param>
    ''' <param name="ToLoc">Ending location..</param>
    Public Sub ShowDirections(ByVal FromLoc As String, ByVal ToLoc As String)
        Dim url As String
        'url = http://maps.google.com/?q=2233+s+columbus+blvd+philadelphia&z=1
        'url = http://maps.google.com/maps?q=from+2233+s+columbus+blvd+philadelphia+to+1600+pennsylvania
        'url = "http://maps.google.com/?sll=39.918,-75.136"
        url = "http://maps.google.com/?q=from+" + fmt.ToURL(FromLoc) & "+to+" + fmt.ToURL(ToLoc)
        ui.OpenFile(url)
    End Sub

    ''' <summary>Return the latitude and longitude for a given location.</summary>
    ''' <param name="Location">Starting location, an address and/or zip code.</param>
    ''' <param name="Latitude">Returned Latitude.</param>
    ''' <param name="Longitude">Returned Longitude.</param>
    ''' <returns>Number of results found.</returns>
    ''' <remarks>Results should only be considered reliable only if exactly 1 is returned.</remarks>
    Public Function GetLadLon(ByVal Location As String, ByRef Latitude As String, ByRef Longitude As String) As Integer
        Dim url As String = "http://local.yahooapis.com/MapsService/V1/geocode?appid=Healthfirst&location=" & fmt.ToURL(Location)
        Try
            Dim numlocs As Integer = 0
            Dim tr As XmlTextReader = New XmlTextReader(url)
            While tr.Read
                If tr.Name = "Latitude" Then
                    Latitude = tr.ReadElementString
                    numlocs += 1
                End If
                If tr.Name = "Longitude" Then
                    Longitude = tr.ReadElementString
                End If
            End While
            Return numlocs
        Catch ex As Exception
            If ex.Message.Contains("(400) Bad Request.") Then
                Throw New Exception("Lookup failed for: " + Location)
            Else
                Throw
            End If
        End Try
    End Function

    ''' <summary>Returns an address in standardized format.</summary>
    ''' <param name="Address">The address to standardize.</param>
    ''' <returns>The standardized address.</returns>
    Public Function GetStandardAddress(ByVal Address As String) As String
        Dim URL As String = "http://www.trynt.com/address-standardization-api/v1/?s="
        URL += fmt.ToURL(Address)
        Dim request As HttpWebRequest = WebRequest.Create(URL)
        Dim response As HttpWebResponse = request.GetResponse()
        Dim reader As StreamReader = New StreamReader(response.GetResponseStream())
        Dim resp As String = reader.ReadToEnd
        Return resp
    End Function

    ''' <summary>Calculate the distance between two sets of geo-coordinates.</summary>
    ''' <param name="lat1">Latitude for the first location.</param>
    ''' <param name="lon1">Longitude for the first location.</param>
    ''' <param name="lat2">Latitude for the second location.</param>
    ''' <param name="lon2">Longitude for the second location.</param>
    ''' <returns>The direct distance between the two coordinates.</returns>
    Public Function CalcDistance(ByVal lat1 As Double, ByVal lon1 As Double, ByVal lat2 As Double, ByVal lon2 As Double) As Double
        Try
            Dim dist As Double = Math.Sin(deg2rad(lat1)) * Math.Sin(deg2rad(lat2)) + Math.Cos(deg2rad(lat1)) * Math.Cos(deg2rad(lat2)) * Math.Cos(deg2rad(lon1 - lon2))
            dist = rad2deg(Math.Acos(dist)) * 69.09
            Return dist
        Catch ex As Exception
            ReportError(ex)
        End Try
    End Function

    ' Convert degrees to radians
    Private Function deg2rad(ByVal deg As Double) As Double
        Return (deg * Math.PI / 180.0)
    End Function

    ' Convert radians to degrees
    Private Function rad2deg(ByVal rad As Double) As Double
        Return rad / Math.PI * 180.0
    End Function

End Class
