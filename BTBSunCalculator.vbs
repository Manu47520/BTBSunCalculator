Option Explicit

' Définissez les coordonnées GPS et la date/heure en UTC
Dim latitude, longitude, year, month, day, hour, minute, second, time, date
latitude = InputBox("Please enter the latitude of the GPS position:", "BTB Sun Calculator")
latitude = Replace(latitude, ".", ",")
longitude = InputBox("Please enter the longitude of the GPS position:", "BTB Sun Calculator")
longitude = Replace(longitude, ".", ",")

' Boucle pour la saisie de la date avec vérification
Do
    date = InputBox("Please enter a date in DD/MM/YYYY format:", "BTB Sun Calculator")
    
    ' Vérifie si le format est JJ/MM/AAAA avec une expression régulière
    Dim regexDate
    Set regexDate = New RegExp
    regexDate.Pattern = "^\d{2}/\d{2}/\d{4}$"
    regexDate.IgnoreCase = True

    If regexDate.Test(date) Then
        ' Extrait le jour, le mois et l'année
        day = Left(date, 2)
        month = Mid(date, 4, 2)
        year = Right(date, 4)
        
        ' Vérifie si c'est une date valide
        If IsDate(day & "/" & month & "/" & year) Then
            Exit Do ' Sortie de la boucle si la date est correcte
        Else
            MsgBox "The date entered is invalid. Please try again.", 0+48+0, "BTB Sun Calculator"
        End If
    Else
        MsgBox "Incorrect date format. Please use DD/MM/YYYY.", 0+48+0, "BTB Sun Calculator"
    End If
Loop

' Boucle pour la saisie de l'heure avec vérification
Do
    time = InputBox("Please enter the time in HH:MM format:", "BTB Sun Calculator")
    
    ' Vérifie si le format est HH:MM avec une expression régulière
    Dim regexTime
    Set regexTime = New RegExp
    regexTime.Pattern = "^\d{2}:\d{2}$"
    regexTime.IgnoreCase = True

    If regexTime.Test(time) Then
        ' Extrait l'heure et les minutes
        hour = Left(time, 2)
        minute = Right(time, 2)
        
        ' Vérifie que l'heure et les minutes sont valides
        If hour >= 0 And hour < 24 And minute >= 0 And minute < 60 Then
            Exit Do ' Sortie de la boucle si l'heure est correcte
        Else
            MsgBox "The time entered is invalid. Please try again.", 0+48+0, "BTB Sun Calculator"
        End If
    Else
        MsgBox "Incorrect time format. Please use HH:MM.", 0+48+0, "BTB Sun Calculator"
    End If
Loop

second = 0

' Convertir la date en Jour Julien
Dim JD, UT
UT = hour + (minute / 60) + (second / 3600)
JD = CalcJulianDate(year, month, day, UT)

' Calcul du Temps Sidéral Local
Dim GST, LST
GST = CalcGST(JD)
LST = GST + longitude

' Calcul de la position du Soleil (anomalie moyenne, longitude écliptique, etc.)
Dim M, lambda, RA, DEC
M = 357.5291 + 0.98560028 * (JD - 2451545.0)
lambda = M + 1.9148 * Sin(DegToRad(M)) + 0.0200 * Sin(DegToRad(2 * M)) + 282.9372
RA = CalcRA(lambda)
DEC = CalcDEC(lambda)

' Calcul de l'angle horaire et de l'altitude/azimut
Dim H, altitude, azimuth
H = LST - RA
altitude = CalcAltitude(latitude, DEC, H)
azimuth = CalcAzimuth(latitude, DEC, H)

' Calcul des composants x, y, z pour sunDirection
Dim x, y, z
x = Sin(DegToRad(azimuth)) * Sin(DegToRad(altitude))
z = Cos(DegToRad(azimuth)) * Sin(DegToRad(altitude))
y = Cos(DegToRad(altitude))

x = Round(x, 7)
y = Round(y, 7)
z = Round(z, 7)

x = Replace(x, ",", ".")
y = Replace(y, ",", ".")
z = Replace(z, ",", ".")

' Écrire le résultat dans un fichier settings.ini
Dim fso, file
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.CreateTextFile("settings.ini", True)
file.WriteLine "[SETTINGS]"
file.WriteLine "sunDirection = " & x & ", " & z & ", " & y & "    ; " & date & " " & time
file.Close

' Afficher les résultats
' WScript.Echo "Azimuth of the Sun: " & Round(azimuth, 2) & vbCrLf & "Altitude of the Sun: " & Round(altitude, 2)
' WScript.Echo " The settings.ini file was created with the values: " & vbCrLf & "sunDirection = " & x & ", " & z & ", " & y
MsgBox "Azimuth of the Sun: " & Round(azimuth, 2) & vbCrLf & "Altitude of the Sun: " & Round(altitude, 2) & "", 0+64+0, "BTB Sun Calculator"
MsgBox "The settings.ini file was created with the values: " & vbCrLf & "sunDirection = " & x & ", " & z & ", " & y & "", 0+64+0, "BTB Sun Calculator"


' Fonction pour calculer le Jour Julien
Function CalcJulianDate(y, m, d, UT)
    CalcJulianDate = 367 * y - Int((7 * (y + Int((m + 9) / 12))) / 4) + Int(275 * m / 9) + d + 1721013.5 + (UT / 24)
End Function

' Fonction pour calculer le temps sidéral de Greenwich
Function CalcGST(JD)
    Dim GST
    GST = 280.46061837 + 360.98564736629 * (JD - 2451545.0)
    CalcGST = GST Mod 360
End Function

' Fonction pour calculer l'ascension droite (RA)
Function CalcRA(lambda)
    CalcRA = Atan2(Cos(DegToRad(23.44)) * Sin(DegToRad(lambda)), Cos(DegToRad(lambda)))
End Function

' Fonction pour calculer la déclinaison (DEC)
Function CalcDEC(lambda)
    CalcDEC = Asin(Sin(DegToRad(23.44)) * Sin(DegToRad(lambda)))
End Function

' Fonction pour calculer l'altitude du Soleil
Function CalcAltitude(lat, DEC, H)
    CalcAltitude = Asin(Sin(DegToRad(lat)) * Sin(DegToRad(DEC)) + Cos(DegToRad(lat)) * Cos(DegToRad(DEC)) * Cos(DegToRad(H)))
End Function

' Fonction pour calculer l'azimut du Soleil
Function CalcAzimuth(lat, DEC, H)
    CalcAzimuth = Atan2(-Sin(DegToRad(H)), Tan(DegToRad(DEC)) * Cos(DegToRad(lat)) - Sin(DegToRad(lat)) * Cos(DegToRad(H)))
End Function

' Fonctions utilitaires pour conversion en radians/degrés et arrondi
Function DegToRad(deg)
    DegToRad = deg * (3.14159265358979 / 180)
End Function

Function RadToDeg(rad)
    RadToDeg = rad * (180 / 3.14159265358979)
End Function

Function Asin(x)
    Asin = RadToDeg(Atn(x / Sqr(-x * x + 1)))
End Function

Function Atan2(y, x)
    If x > 0 Then
        Atan2 = RadToDeg(Atn(y / x))
    ElseIf x < 0 And y >= 0 Then
        Atan2 = RadToDeg(Atn(y / x)) + 180
    ElseIf x < 0 And y < 0 Then
        Atan2 = RadToDeg(Atn(y / x)) - 180
    ElseIf x = 0 And y > 0 Then
        Atan2 = 90
    ElseIf x = 0 And y < 0 Then
        Atan2 = -90
    Else
        Atan2 = 0
    End If
End Function
