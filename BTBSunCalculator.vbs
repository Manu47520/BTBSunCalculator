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
    Dim regexDate
    Set regexDate = New RegExp
    regexDate.Pattern = "^\d{2}/\d{2}/\d{4}$"
    regexDate.IgnoreCase = True

    If regexDate.Test(date) Then
        day = Left(date, 2)
        month = Mid(date, 4, 2)
        year = Right(date, 4)

        If IsDate(day & "/" & month & "/" & year) Then
            Exit Do
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
    Dim regexTime
    Set regexTime = New RegExp
    regexTime.Pattern = "^\d{2}:\d{2}$"
    regexTime.IgnoreCase = True

    If regexTime.Test(time) Then
        hour = Left(time, 2)
        minute = Right(time, 2)

        If hour >= 0 And hour < 24 And minute >= 0 And minute < 60 Then
            Exit Do
        Else
            MsgBox "The time entered is invalid. Please try again.", 0+48+0, "BTB Sun Calculator"
        End If
    Else
        MsgBox "Incorrect time format. Please use HH:MM.", 0+48+0, "BTB Sun Calculator"
    End If
Loop

second = 0

' Créer le fichier settings.ini
Dim fso, file
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.CreateTextFile("settings.ini", True)
file.WriteLine "[SETTINGS]"

' Calcul de la position du soleil pour l'heure saisie
Call CalculateAndWriteSunPosition(CInt(hour), CInt(minute), "User-Specified", False)

' Calcul pour les autres moments de la journée avec un point-virgule en début de ligne
Call CalculateAndWriteSunPosition(6, 0, "Sunrise", True)
Call CalculateAndWriteSunPosition(13, 0, "Noon", True)
Call CalculateAndWriteSunPosition(18, 0, "Sunset", True)
Call CalculateAndWriteSunPosition(3, 0, "Night", True)

' Fermer le fichier settings.ini
file.Close

MsgBox "The settings.ini file was created with sun positions for specified time, Sunrise (06:00), Noon (13:00), Sunset (18:00), and Night (03:00).", vbInformation, "BTB Sun Calculator"

' Demander à l'utilisateur s'il souhaite ouvrir le fichier
Dim openFile
openFile = MsgBox("Do you want to open the settings.ini file?", vbYesNo + vbQuestion, "BTB Sun Calculator")

If openFile = vbYes Then
    Set file = Nothing
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.OpenTextFile("settings.ini").Close  ' Force la fermeture au cas où

    ' Ouvrir le fichier avec l'éditeur de texte par défaut
    Dim shell
    Set shell = CreateObject("WScript.Shell")
    shell.Run "notepad.exe settings.ini"
End If

' Subroutine pour calculer et écrire la position du soleil dans le fichier
Sub CalculateAndWriteSunPosition(calcHour, calcMinute, label, addCommentSymbol)
    Dim UT, JD, GST, LST, M, lambda, RA, DEC, H, altitude, azimuth, x, y, z
    UT = calcHour + (calcMinute / 60)

    ' Calcul du Jour Julien
    JD = CalcJulianDate(year, month, day, UT)

    ' Calcul du Temps Sidéral Local
    GST = CalcGST(JD)
    LST = GST + longitude

    ' Calcul de la position du Soleil
    M = 357.5291 + 0.98560028 * (JD - 2451545.0)
    lambda = M + 1.9148 * Sin(DegToRad(M)) + 0.0200 * Sin(DegToRad(2 * M)) + 282.9372
    RA = CalcRA(lambda)
    DEC = CalcDEC(lambda)

    ' Calcul de l'angle horaire et de l'altitude/azimut
    H = LST - RA
    altitude = CalcAltitude(latitude, DEC, H)
    azimuth = CalcAzimuth(latitude, DEC, H)

    ' Calcul des composants x, y, z pour sunDirection
    x = Sin(DegToRad(azimuth)) * Sin(DegToRad(altitude))
    z = Cos(DegToRad(azimuth)) * Sin(DegToRad(altitude))
    y = Cos(DegToRad(altitude))

    x = Round(x, 7)
    y = Round(y, 7)
    z = Round(z, 7)

    x = Replace(x, ",", ".")
    y = Replace(y, ",", ".")
    z = Replace(z, ",", ".")

    ' Formater l'heure et les minutes en deux chiffres
    Dim formattedHour, formattedMinute
    formattedHour = Right("0" & calcHour, 2)
    formattedMinute = Right("0" & calcMinute, 2)

    ' Ajouter ";" devant la ligne si c'est pour un moment spécifique (sunrise, noon, etc.)
    Dim lineContent
    lineContent = "sunDirection = " & x & ", " & z & ", " & y & "    ; " & date & " " & formattedHour & ":" & formattedMinute & " " & label
    If addCommentSymbol Then
        lineContent = ";" & lineContent
    End If

    ' Écrire la ligne dans le fichier settings.ini
    file.WriteLine lineContent
End Sub

' Fonctions pour les calculs (inchangées)
Function CalcJulianDate(y, m, d, UT)
    CalcJulianDate = 367 * y - Int((7 * (y + Int((m + 9) / 12))) / 4) + Int(275 * m / 9) + d + 1721013.5 + (UT / 24)
End Function

Function CalcGST(JD)
    Dim GST
    GST = 280.46061837 + 360.98564736629 * (JD - 2451545.0)
    CalcGST = GST Mod 360
End Function

Function CalcRA(lambda)
    CalcRA = Atan2(Cos(DegToRad(23.44)) * Sin(DegToRad(lambda)), Cos(DegToRad(lambda)))
End Function

Function CalcDEC(lambda)
    CalcDEC = Asin(Sin(DegToRad(23.44)) * Sin(DegToRad(lambda)))
End Function

Function CalcAltitude(lat, DEC, H)
    CalcAltitude = Asin(Sin(DegToRad(lat)) * Sin(DegToRad(DEC)) + Cos(DegToRad(lat)) * Cos(DegToRad(DEC)) * Cos(DegToRad(H)))
End Function

Function CalcAzimuth(lat, DEC, H)
    CalcAzimuth = Atan2(-Sin(DegToRad(H)), Tan(DegToRad(DEC)) * Cos(DegToRad(lat)) - Sin(DegToRad(lat)) * Cos(DegToRad(H)))
End Function

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
