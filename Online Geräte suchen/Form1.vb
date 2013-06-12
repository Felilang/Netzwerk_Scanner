Imports System.Threading
Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Management
Imports System.IO

Public Class Form1
    Public Gesamtausgabe As String
    Public abbruch As Integer
    Public stop1 As Integer
    Dim IPThread As New Thread(AddressOf IPSub)
    Dim prüfenThread As New Thread(AddressOf Prüfen)
    Public treen As TreeNode

    Public Sub IPSub()
        'Progress Bar wieder auf 0 setzen
        Me.Invoke(Sub() ProgressBar1.Value = 0)
        'deklarieren aller Benötigten Variablen
        Dim pingResult As Boolean
        Dim IPQ1 As Integer
        Dim IPQ2 As Integer
        Dim IPQ3 As Integer
        Dim IPQ4 As Integer
        Dim IPQ5 As Integer
        Dim IPQ6 As Integer
        Dim IPQ7 As Integer
        Dim IPQ8 As Integer
        Dim IPQ10 As Integer
        Dim IPQ20 As Integer
        Dim IPQ30 As Integer
        Dim IPQ40 As Integer
        Dim IPQ50 As Integer
        Dim IPQ60 As Integer
        Dim IPQ70 As Integer
        Dim IPQ80 As Integer
        Dim Loop1 As String
        Dim Loop2 As String
        Dim Loop3 As String
        Dim Zahler As Integer
        Dim zahler1 As Integer
        Dim Zahler2 As Integer
        Dim kIP As Integer
        Dim gIP As Integer
        Dim ausgabe As String
        Dim cmd As String
        Dim cmd2 As String
        Dim macProzess As New Process
        Dim cmdmac As String
        Dim pcProzess As New Process
        Dim cmdpc As String
        Dim mac As String
        Dim ipzahler As Integer
        Dim offlinezahler As Integer
        Dim pcname As String
        Dim status As String
        Dim leer As String
        Dim ip As String
        'gip = größte mögliche IP Adresse
        gIP = 255
        'kip = kleinste mögliche IP Adresse
        kIP = 0
        Loop1 = IPQ8
        Loop2 = IPQ7
        Loop3 = IPQ6
        Zahler = 0
        Zahler2 = 0
        mac = ""
        'Prüfen ob eine Ganze IP in der Oberen Zeile steht
        If IPQ1Box.Text = "" Or IPQ2Box.Text = "" Or IPQ3Box.Text = "" Or IPQ4Box.Text = "" Then
            MsgBox("Geben sie bitte eine vollständige IP Adresse ein", Buttons:=vbCritical, Title:="IP Fehlt")
        Else
            'Prüfen ob eine Ganze IP in der Unteren Zeile steht fals nicht wird dort die IP von oben hineingeschrieben
            If IPQ5Box.Text = "" And IPQ6Box.Text = "" And IPQ7Box.Text = "" And IPQ8Box.Text = "" Then
                Me.Invoke(Sub() IPQ5Box.Text = IPQ1Box.Text)
                Me.Invoke(Sub() IPQ6Box.Text = IPQ2Box.Text)
                Me.Invoke(Sub() IPQ7Box.Text = IPQ3Box.Text)
                Me.Invoke(Sub() IPQ8Box.Text = IPQ4Box.Text)
            End If
            'Den Quad Variabeln die Werte der Texboxen zuweisen
            IPQ1 = IPQ1Box.Text
            IPQ2 = IPQ2Box.Text
            IPQ3 = IPQ3Box.Text
            IPQ4 = IPQ4Box.Text
            IPQ5 = IPQ5Box.Text
            IPQ6 = IPQ6Box.Text
            IPQ7 = IPQ7Box.Text
            IPQ8 = IPQ8Box.Text
            'Überprüfen ob oben die Kleinere IP steht, fals nicht die IPs Tauschen
            If IPQ5 < IPQ1 Or IPQ6 < IPQ2 And IPQ5 = IPQ1 Or IPQ7 < IPQ3 And IPQ6 <= IPQ2 And IPQ5 <= IPQ1 Or IPQ8 < IPQ4 And IPQ7 <= IPQ3 And IPQ6 <= IPQ2 And IPQ5 <= IPQ1 Then
                Dim tausch As Integer
                tausch = IPQ1Box.Text
                Me.Invoke(Sub() IPQ1Box.Text = IPQ5Box.Text)
                Me.Invoke(Sub() IPQ5Box.Text = tausch)
                tausch = IPQ2Box.Text
                Me.Invoke(Sub() IPQ2Box.Text = IPQ6Box.Text)
                Me.Invoke(Sub() IPQ6Box.Text = tausch)
                tausch = IPQ3Box.Text
                Me.Invoke(Sub() IPQ3Box.Text = IPQ7Box.Text)
                Me.Invoke(Sub() IPQ7Box.Text = tausch)
                tausch = IPQ4Box.Text
                Me.Invoke(Sub() IPQ4Box.Text = IPQ8Box.Text)
                Me.Invoke(Sub() IPQ8Box.Text = tausch)
                'Den Quad Variabeln die Werte der Texboxen erneut zuweisen
                IPQ1 = IPQ1Box.Text
                IPQ2 = IPQ2Box.Text
                IPQ3 = IPQ3Box.Text
                IPQ4 = IPQ4Box.Text
                IPQ5 = IPQ5Box.Text
                IPQ6 = IPQ6Box.Text
                IPQ7 = IPQ7Box.Text
                IPQ8 = IPQ8Box.Text
                MsgBox("End und Start IP wurden vertauscht", Buttons:=vbInformation, Title:="IPs wurden vertauscht")
            End If
            'den QuadVariablen zum ermitteln wie viele IPs geprüft werden müssen die Werte der Textboxen zuweisen
            IPQ10 = IPQ1Box.Text
            IPQ20 = IPQ2Box.Text
            IPQ30 = IPQ3Box.Text
            IPQ40 = IPQ4Box.Text
            IPQ50 = IPQ5Box.Text
            IPQ60 = IPQ6Box.Text
            IPQ70 = IPQ7Box.Text
            IPQ80 = IPQ8Box.Text
            'die Schleifen mit denen die Anzahl der IPs ermittelt werden
            'in diesen Schleifen wird überprüft, ob die Obere und Untere IP gleich sind fals nicht wird die untere Um eine IP erhöht
            Do
                Do
                    Do
                        Do
                            Loop1 = IPQ8
                            'der zahler der die Anzahl der IPs zahlt
                            ipzahler = ipzahler + 1
                            IPQ40 = IPQ40 + 1
                            If IPQ30 < IPQ70 Then
                                Loop1 = gIP
                            End If
                            If IPQ20 < IPQ60 Then
                                Loop1 = gIP
                            End If
                            If IPQ10 < IPQ50 Then
                                Loop1 = gIP
                            End If
                        Loop While IPQ40 <= Loop1
                        If abbruch = 1 Then
                            Exit Do
                        End If
                        IPQ40 = kIP
                        If IPQ20 < IPQ60 Then
                            IPQ30 = IPQ30 + 1
                            Loop2 = gIP
                        Else
                            If IPQ10 < IPQ50 Then
                                IPQ30 = IPQ30 + 1
                                Loop2 = gIP
                            Else
                                If IPQ30 < IPQ70 Then
                                    IPQ30 = IPQ30 + 1
                                    Loop2 = IPQ70
                                Else
                                    If IPQ30 = IPQ70 Then
                                        IPQ30 = IPQ30 + 1
                                    End If
                                End If
                            End If
                        End If
                    Loop While IPQ30 <= Loop2
                    If abbruch = 1 Then
                        Exit Do
                    End If
                    IPQ30 = kIP
                    If IPQ10 < IPQ50 Then
                        IPQ20 = IPQ20 + 1
                        Loop3 = gIP
                    Else
                        If IPQ20 < IPQ60 Then
                            IPQ20 = IPQ20 + 1
                            Loop3 = IPQ60
                        Else
                            If IPQ20 = IPQ60 Then
                                IPQ20 = IPQ20 + 1
                            End If
                        End If
                    End If
                Loop While IPQ20 <= Loop3
                If abbruch = 1 Then
                    Exit Do
                End If
                IPQ20 = kIP
                If IPQ10 <= IPQ50 Then
                    IPQ10 = IPQ10 + 1
                End If
            Loop While IPQ10 <= IPQ50


            'der Progressbar den maximalen Wert zuweisen, in diesem Fall die Anzahl der IP Adressen, welche von uns inden Schleifen ermittelet wurden
            Me.Invoke(Sub() ProgressBar1.Maximum = ipzahler)


            'die Oberern vier verschachtelten Schleifen erneut, nur diesemal werden die IPs überprüft
            'die mac adresse gesucht und der Pc Name ermittelt
            Do
                Do
                    Do
                        Do
                            Loop1 = IPQ8
                            'fals ein Fehler in diesen Funktionen auftritt, wird er durch Try abgefangen
                            Try
                                'der Variablen Ip die Einzelnen Quads zuweisen
                                ip = IPQ1.ToString & "." & IPQ2.ToString & "." & IPQ3.ToString & "." & IPQ4.ToString
                                'die Ip überprüfen die gerade in der Variabel Ip steht
                                pingResult = My.Computer.Network.Ping(ip)
                                'Abbrechen fals von einem Anderem Programm die Globale Variable auf 1 gesetzt wird
                                If abbruch = 1 Then
                                    Exit Do
                                End If
                                'Der Variablen ausgabe den Wert online oder Offline zuweisen
                                If pingResult = False Then
                                    ausgabe = "Offline"
                                Else
                                    ausgabe = "Online"
                                End If
                                'der Variablen status den Wert von ausgabe zuweisen
                                status = ausgabe
                                'Wenn die Ip Online war wird jetzt die Mac und der PC Name ermittelt, 
                                'Wenn sie Offline war wird dieser Teil übersprungen
                                If ausgabe = "Online" Then
                                    'der Befehlt mit dem Die Mac ermittelt wird
                                    'das ganze wird in Cmd ausgeführt und alles wird als String zurückgegeben
                                    cmd = "/c arp -a " & ip
                                    macProzess.StartInfo.RedirectStandardOutput = True
                                    macProzess.StartInfo.RedirectStandardInput = True
                                    macProzess.StartInfo.CreateNoWindow = True
                                    macProzess.StartInfo.UseShellExecute = False
                                    macProzess.StartInfo.FileName = "cmd.exe"
                                    macProzess.StartInfo.Arguments = cmd
                                    'hier wird Cmd gestartet die Befehle davor haben cmd gesagt, wie es ausgeführt werden soll
                                    macProzess.Start()
                                    'der Variablen Cmdmac wird die rückgabe von cmd zugewiesen
                                    cmdmac = macProzess.StandardOutput.ReadToEnd()
                                    'die Variable cmdmac wird unterteilt in ein Array, und zwar wird die Variable immer wenn eine Leertaste kommt unterteilt
                                    Dim strTeil() As String = cmdmac.Split(" ")
                                    'überprüfen ob nach  der Teilung der Variable an zweiter Stelle drei sriche stehen, fals dies der Fall ist überspringt er das suchen der mac
                                    'da dann keine Mac adresse gefunden wurde
                                    If strTeil(2) = "---" Then
                                        'die mac kann an verschiedenen Punkten stehen, sie steht zwischen 27 und 32 stell die anderen sind immer leer
                                        If strTeil(27) = "" Then
                                        Else
                                            mac = strTeil(27)
                                        End If
                                        If strTeil(28) = "" Then
                                        Else
                                            mac = strTeil(28)
                                        End If
                                        If strTeil(29) = "" Then
                                        Else
                                            mac = strTeil(29)
                                        End If
                                        If strTeil(30) = "" Then
                                        Else
                                            mac = strTeil(30)
                                        End If
                                        If strTeil(31) = "" Then
                                        Else
                                            mac = strTeil(31)
                                        End If
                                        If strTeil(32) = "" Then
                                        Else
                                            mac = strTeil(32)
                                        End If
                                    End If
                                    'hier wird der PC Name ermittelt anhand des Ping -a Befehls
                                    'Im grunde funktioniert das wieder wie bei der Mac cmd Wird gesagt wie es starten soll, dann startet es und 
                                    'schreibt das ermittelte in die Variable Cmdpc
                                    cmd2 = "/c ping -a " & ip & " -n 1"
                                    pcProzess.StartInfo.RedirectStandardOutput = True
                                    pcProzess.StartInfo.RedirectStandardInput = True
                                    pcProzess.StartInfo.CreateNoWindow = True
                                    pcProzess.StartInfo.UseShellExecute = False
                                    pcProzess.StartInfo.FileName = "cmd.exe"
                                    pcProzess.StartInfo.Arguments = cmd2
                                    pcProzess.Start()
                                    cmdpc = pcProzess.StandardOutput.ReadToEnd()
                                    'hier wird dann die Variable cmdpc wieder unterteil
                                    Dim strTeil2() As String = cmdpc.Split(" ")
                                    'der Pcname steht immer an vierter stelle also wird der Variablen PCname der vierte Teil der unterteilten cmdpc variablen zugewiesen
                                    pcname = strTeil2(4)
                                    'zusätzlich wird noch der Zahler1 um 1 erhöht, so kann man ermittelen wie viele IPs Online waren
                                    zahler1 = zahler1 + 1
                                Else
                                    'fals der Variablem Ausgabe der Wert Online nicht zugewiesen war kommt man hier an
                                    'hier wird hochgezählt, wie viele IPs Offline waren 
                                    offlinezahler = offlinezahler + 1
                                    'der Variablen Mac wird nichts zugewiesen
                                    mac = ""
                                    'der Variablen PCName wird ebenfalls nichts zugewiesen, damit in den Variablen keine alten Werte mehr stehen
                                    pcname = ""
                                End If
                                'wie viele Durchläufe es insgesamt sind wird hier ermittelt
                                Zahler = Zahler + 1
                                'den Labels10 und 11 werden die Akktuellen anzahlen der Offlinen und Onlinen Geräte zugewiesen
                                Me.Invoke(Sub() Label10.Text = "Online: " & zahler1)
                                Me.Invoke(Sub() Label11.Text = "Offline: " & offlinezahler)
                                'hier wird ermittelt wie groß der Abstand zwischen IP und status sein muss, damit der Status immer schön untereinander angezeigt wird
                                Dim laenge As Integer
                                laenge = Len(ip)
                                Dim anleer As Integer
                                anleer = 16 - laenge
                                leer = New String(" ", 2 * anleer)
                                'Alle Werte werden der Variablen Gesamtausgabe zugewiesen
                                Gesamtausgabe = ip & leer & status & "  " & mac & "  " & pcname
                                'der Treeview Box wird gesagt das die Farbe der Ausgabe über die Gasamte Spalte gehen soll und im Hintergrund sein soll
                                Me.Invoke(Sub() TreeView1.ShowLines = False)
                                Me.Invoke(Sub() TreeView1.FullRowSelect = True)
                                Me.Invoke(Sub() TreeView1.ShowRootLines = False)
                                'hier wird entschieden ob der Hintergrund rot für Offline oder grün für Online sein muss
                                'ebensowird dirkt bei der Entscheidung Der Wert Gesamtausgabe ausgegeben
                                If ausgabe = "Online" Then
                                    Me.Invoke(Sub() TreeView1.Nodes.Add(Gesamtausgabe).BackColor = Color.Green)
                                Else
                                    Me.Invoke(Sub() TreeView1.Nodes.Add(Gesamtausgabe).BackColor = Color.Red)
                                End If
                                'Hier wird dann die ProgressBar um eins erhöht
                                Me.Invoke(Sub() ProgressBar1.Value = ProgressBar1.Value + 1)
                            Catch ex As Exception
                                'fals ein Fehler aufgetaucht war ist man bis hierher gesprungen
                            End Try
                            'jetzt kommen noch die Schleifen die fürs hochzählen der IPs zuständig sind
                            IPQ4 = IPQ4 + 1
                            If IPQ3 < IPQ7 Then
                                Loop1 = gIP
                            End If
                            If IPQ2 < IPQ6 Then
                                Loop1 = gIP
                            End If
                            If IPQ1 < IPQ5 Then
                                Loop1 = gIP
                            End If
                        Loop While IPQ4 <= Loop1
                        If abbruch = 1 Then
                            Exit Do
                        End If
                        IPQ4 = kIP
                        If IPQ2 < IPQ6 Then
                            IPQ3 = IPQ3 + 1
                            Loop2 = gIP
                        Else
                            If IPQ1 < IPQ5 Then
                                IPQ3 = IPQ3 + 1
                                Loop2 = gIP
                            Else
                                If IPQ3 < IPQ7 Then
                                    IPQ3 = IPQ3 + 1
                                    Loop2 = IPQ7
                                Else
                                    If IPQ3 = IPQ7 Then
                                        IPQ3 = IPQ3 + 1
                                    End If
                                End If
                            End If
                        End If
                    Loop While IPQ3 <= Loop2
                    If abbruch = 1 Then
                        Exit Do
                    End If
                    IPQ3 = kIP
                    If IPQ1 < IPQ5 Then
                        IPQ2 = IPQ2 + 1
                        Loop3 = gIP
                    Else
                        If IPQ2 < IPQ6 Then
                            IPQ2 = IPQ2 + 1
                            Loop3 = IPQ6
                        Else
                            If IPQ2 = IPQ6 Then
                                IPQ2 = IPQ2 + 1
                            End If
                        End If
                    End If
                Loop While IPQ2 <= Loop3
                If abbruch = 1 Then
                    Exit Do
                End If
                IPQ2 = kIP
                If IPQ1 <= IPQ5 Then
                    IPQ1 = IPQ1 + 1
                End If
            Loop While IPQ1 <= IPQ5
        End If
        'nach durchlauf aller IPs werden noch die Buttons wieder auf Starten gesetzt
        'ebenso werden die Importieren und Exportieren Button wieder Aktiviert
        Me.Invoke(Sub() Button1.Text = "Starten")
        Me.Invoke(Sub() StartenToolStripMenuItem.Text = "Starten")
        Me.Invoke(Sub() ExportierenToolStripMenuItem.Enabled = True)
        Me.Invoke(Sub() ImportierenToolStripMenuItem.Enabled = True)
        'bei einem Abbruch des Durchlaufs wird die ProgressBar wieder auf 0 gesetzt
        If abbruch = 1 Then
            Me.Invoke(Sub() ProgressBar1.Value = 0)
        End If
        'zum schluss werden auch noch die Abbruch die bei einem Manuellen Abbruch auf eins steht wieder auf 0 gestezt
        'und die Stop Variable anhand man prüfen kann ob das Prügramm läuft wieder auf 0 gesetzt
        abbruch = 0
        stop1 = "0"
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'ruft das Programm starten auf
        Call starten()
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'ruft das Programm Ende auf
        Call ende()
    End Sub
    Private Sub IPQ1Box_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles IPQ1Box.KeyPress
        'Fehlermeldung fals die IP nicht im IP Bereich (0-255) liegt
        Dim IPQ1 As String = 1
        If IPQ1Box.Text = "" Then
            IPQ1 = 1
        Else
            IPQ1 = IPQ1Box.Text
        End If
        Dim pruefung As Integer
        pruefung = IPQ1Box.SelectedText & Asc(e.KeyChar)
        If pruefung > 255 Then
            IPQ1Box.Text = ""
        Else
            If IPQ1 > 25 And Asc(e.KeyChar) >= 48 Or IPQ1 = 25 And Asc(e.KeyChar) >= 54 Or IPQ1 = "00" And Asc(e.KeyChar) = 48 Then
                MsgBox("Falsche Eingabe, geben sie bitte eine Zahl zwischen 0 und 255 ein", Buttons:=vbCritical, Title:="Falscher Zahlenbereich")
                e.Handled = True
            Else
            End If
            'verhindern von gehen ins nächste Kästchen wenn das jetztige Kästchen noch leer ist 
            If Asc(e.KeyChar) = 46 Then
                If IPQ1Box.Text = "" Then
                    e.Handled = True
                Else
                    IPQ2Box.Focus()
                    IPQ2Box.SelectAll()
                End If
            End If
            'Verhindern von eingabe mehrerer nullen
            If Asc(e.KeyChar) = 48 Then
                If IPQ1 = "0" Then
                    e.Handled = True
                End If
            End If



            If Asc(e.KeyChar) = 8 Then
                If IPQ1Box.Text = "" Then
                    IPQ1Box.Focus()
                End If
            End If
            Select Case Asc(e.KeyChar)
                Case 48 To 57, 8, 28 To 29
                    ' Zahlen, Backspace und Space zulassen
                Case Else
                    ' alle anderen Eingaben unterdrücken
                    e.Handled = True
            End Select
        End If

    End Sub
    Private Sub IPQ2Box_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles IPQ2Box.KeyPress
        'Fehlermeldung fals die IP nicht im IP Bereich (0-255) liegt
        Dim IPQ2 As String = 1
        If IPQ2Box.Text = "" Then
            IPQ2 = 1
        Else
            IPQ2 = IPQ2Box.Text
        End If
        Dim pruefung As Integer
        pruefung = IPQ2Box.SelectedText & Asc(e.KeyChar)
        If pruefung > 255 Then
            IPQ2Box.Text = ""
        Else
            If IPQ2 > 25 And Asc(e.KeyChar) >= 48 Or IPQ2 = 25 And Asc(e.KeyChar) >= 54 Or IPQ2 = "00" And Asc(e.KeyChar) = 48 Then
                MsgBox("Falsche Eingabe, geben sie bitte eine Zahl zwischen 0 und 255 ein", Buttons:=vbCritical, Title:="Falscher Zahlenbereich")
                e.Handled = True
            Else
            End If
            'verhindern von gehen ins nächste Kästchen wenn das jetztige Kästchen noch leer ist 
            If Asc(e.KeyChar) = 46 Then
                If IPQ2Box.Text = "" Then
                    e.Handled = True
                Else
                    IPQ3Box.Focus()
                    IPQ3Box.SelectAll()
                End If
            End If
            'Verhindern von eingabe mehrerer nullen
            If Asc(e.KeyChar) = 48 Then
                If IPQ2 = "0" Then
                    e.Handled = True
                End If
            End If




            If Asc(e.KeyChar) = 8 Then
                If IPQ2Box.Text = "" Then
                    IPQ1Box.Focus()
                End If
            End If
            Select Case Asc(e.KeyChar)
                Case 48 To 57, 8, 28 To 29
                    ' Zahlen, Backspace und Space zulassen
                Case Else
                    ' alle anderen Eingaben unterdrücken
                    e.Handled = True
            End Select
        End If
    End Sub
    Private Sub IPQ3Box_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles IPQ3Box.KeyPress
        'Fehlermeldung fals die IP nicht im IP Bereich (0-255) liegt
        Dim IPQ3 As String = 1
        If IPQ3Box.Text = "" Then
            IPQ3 = 1
        Else
            IPQ3 = IPQ3Box.Text
        End If
        Dim pruefung As Integer
        pruefung = IPQ3Box.SelectedText & Asc(e.KeyChar)
        If pruefung > 255 Then
            IPQ3Box.Text = ""
        Else
            If IPQ3 > 25 And Asc(e.KeyChar) >= 48 Or IPQ3 = 25 And Asc(e.KeyChar) >= 54 Or IPQ3 = "00" And Asc(e.KeyChar) = 48 Then
                MsgBox("Falsche Eingabe, geben sie bitte eine Zahl zwischen 0 und 255 ein", Buttons:=vbCritical, Title:="Falscher Zahlenbereich")
                e.Handled = True
            Else
            End If
            'verhindern von gehen ins nächste Kästchen wenn das jetztige Kästchen noch leer ist 
            If Asc(e.KeyChar) = 46 Then
                If IPQ3Box.Text = "" Then
                    e.Handled = True
                Else
                    IPQ4Box.Focus()
                    IPQ4Box.SelectAll()
                End If
            End If
            'Verhindern von eingabe mehrerer nullen
            If Asc(e.KeyChar) = 48 Then
                If IPQ3 = "0" Then
                    e.Handled = True
                End If
            End If




            If Asc(e.KeyChar) = 8 Then
                If IPQ3Box.Text = "" Then
                    IPQ2Box.Focus()
                End If
            End If

            Select Case Asc(e.KeyChar)
                Case 48 To 57, 8, 28 To 29
                    ' Zahlen, Backspace und Space zulassen
                Case Else
                    ' alle anderen Eingaben unterdrücken
                    e.Handled = True
            End Select
        End If
    End Sub
    Private Sub IPQ4Box_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles IPQ4Box.KeyPress
        'Fehlermeldung fals die IP nicht im IP Bereich (0-255) liegt
        Dim IPQ4 As String = 1
        If IPQ4Box.Text = "" Then
            IPQ4 = 1
        Else
            IPQ4 = IPQ4Box.Text
        End If
        Dim pruefung As Integer
        pruefung = IPQ4Box.SelectedText & Asc(e.KeyChar)
        If pruefung > 255 Then
            IPQ4Box.Text = ""
        Else
            If IPQ4 > 25 And Asc(e.KeyChar) >= 48 Or IPQ4 = 25 And Asc(e.KeyChar) >= 54 Or IPQ4 = "00" And Asc(e.KeyChar) = 48 Then
                MsgBox("Falsche Eingabe, geben sie bitte eine Zahl zwischen 0 und 255 ein", Buttons:=vbCritical, Title:="Falscher Zahlenbereich")
                e.Handled = True
            Else
            End If
            'verhindern von gehen ins nächste Kästchen wenn das jetztige Kästchen noch leer ist 
            If Asc(e.KeyChar) = 46 Then
                If IPQ4Box.Text = "" Then
                    e.Handled = True
                Else
                    IPQ5Box.Focus()
                    IPQ5Box.SelectAll()
                End If
            End If
            'Verhindern von eingabe mehrerer nullen
            If Asc(e.KeyChar) = 48 Then
                If IPQ4 = "0" Then
                    e.Handled = True
                End If
            End If
            If Asc(e.KeyChar) = 13 Then
                If IPQ4Box.Text = "" Then
                Else
                    Call starten()
                End If
            End If



            If Asc(e.KeyChar) = 8 Then
                If IPQ4Box.Text = "" Then
                    IPQ3Box.Focus()
                End If
            End If

            Select Case Asc(e.KeyChar)
                Case 48 To 57, 8, 28 To 29
                    ' Zahlen, Backspace und Space zulassen
                Case Else
                    ' alle anderen Eingaben unterdrücken
                    e.Handled = True
            End Select
        End If
    End Sub
    Private Sub IPQ5Box_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles IPQ5Box.KeyPress
        'Fehlermeldung fals die IP nicht im IP Bereich (0-255) liegt
        Dim IPQ5 As String = 1
        If IPQ5Box.Text = "" Then
            IPQ5 = 1
        Else
            IPQ5 = IPQ5Box.Text
        End If
        Dim pruefung As Integer
        pruefung = IPQ5Box.SelectedText & Asc(e.KeyChar)
        If pruefung > 255 Then
            IPQ5Box.Text = ""
        Else
            If IPQ5 > 25 And Asc(e.KeyChar) >= 48 Or IPQ5 = 25 And Asc(e.KeyChar) >= 54 Or IPQ5 = "00" And Asc(e.KeyChar) = 48 Then
                MsgBox("Falsche Eingabe, geben sie bitte eine Zahl zwischen 0 und 255 ein", Buttons:=vbCritical, Title:="Falscher Zahlenbereich")
                e.Handled = True
            Else
            End If
            'verhindern von gehen ins nächste Kästchen wenn das jetztige Kästchen noch leer ist 
            If Asc(e.KeyChar) = 46 Then
                If IPQ5Box.Text = "" Then
                    e.Handled = True
                Else
                    IPQ6Box.Focus()
                    IPQ6Box.SelectAll()
                End If
            End If
            'Verhindern von eingabe mehrerer nullen
            If Asc(e.KeyChar) = 48 Then
                If IPQ5 = "0" Then
                    e.Handled = True
                End If
            End If





            If Asc(e.KeyChar) = 8 Then
                If IPQ5Box.Text = "" Then
                    IPQ4Box.Focus()
                End If
            End If

            Select Case Asc(e.KeyChar)
                Case 48 To 57, 8, 28 To 29
                    ' Zahlen, Backspace und Space zulassen
                Case Else
                    ' alle anderen Eingaben unterdrücken
                    e.Handled = True
            End Select
        End If
    End Sub
    Private Sub IPQ6Box_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles IPQ6Box.KeyPress
        'Fehlermeldung fals die IP nicht im IP Bereich (0-255) liegt
        Dim IPQ6 As String = 1
        If IPQ6Box.Text = "" Then
            IPQ6 = 1
        Else
            IPQ6 = IPQ6Box.Text
        End If
        Dim pruefung As Integer
        pruefung = IPQ6Box.SelectedText & Asc(e.KeyChar)
        If pruefung > 255 Then
            IPQ6Box.Text = ""
        Else
            If IPQ6 > 25 And Asc(e.KeyChar) >= 48 Or IPQ6 = 25 And Asc(e.KeyChar) >= 54 Or IPQ6 = "00" And Asc(e.KeyChar) = 48 Then
                MsgBox("Falsche Eingabe, geben sie bitte eine Zahl zwischen 0 und 255 ein", Buttons:=vbCritical, Title:="Falscher Zahlenbereich")
                e.Handled = True
            Else
            End If
            'verhindern von gehen ins nächste Kästchen wenn das jetztige Kästchen noch leer ist 
            If Asc(e.KeyChar) = 46 Then
                If IPQ6Box.Text = "" Then
                    e.Handled = True
                Else
                    IPQ7Box.Focus()
                    IPQ7Box.SelectAll()
                End If
            End If
            'Verhindern von eingabe mehrerer nullen
            If Asc(e.KeyChar) = 48 Then
                If IPQ6 = "0" Then
                    e.Handled = True
                End If
            End If




            If Asc(e.KeyChar) = 8 Then
                If IPQ6Box.Text = "" Then
                    IPQ5Box.Focus()
                End If
            End If

            Select Case Asc(e.KeyChar)
                Case 48 To 57, 8, 28 To 29
                    ' Zahlen, Backspace und Space zulassen
                Case Else
                    ' alle anderen Eingaben unterdrücken
                    e.Handled = True
            End Select
        End If
    End Sub
    Private Sub IPQ7Box_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles IPQ7Box.KeyPress
        'Fehlermeldung fals die IP nicht im IP Bereich (0-255) liegt
        Dim IPQ7 As String = 1
        If IPQ7Box.Text = "" Then
            IPQ7 = 1
        Else
            IPQ7 = IPQ7Box.Text
        End If
        Dim pruefung As Integer
        pruefung = IPQ7Box.SelectedText & Asc(e.KeyChar)
        If pruefung > 255 Then
            IPQ7Box.Text = ""
        Else
            If IPQ7 > 25 And Asc(e.KeyChar) >= 48 Or IPQ7 = 25 And Asc(e.KeyChar) >= 54 Or IPQ7 = "00" And Asc(e.KeyChar) = 48 Then
                MsgBox("Falsche Eingabe, geben sie bitte eine Zahl zwischen 0 und 255 ein", Buttons:=vbCritical, Title:="Falscher Zahlenbereich")
                e.Handled = True
            Else
            End If
            'verhindern von gehen ins nächste Kästchen wenn das jetztige Kästchen noch leer ist 
            If Asc(e.KeyChar) = 46 Then
                If IPQ7Box.Text = "" Then
                    e.Handled = True
                Else
                    IPQ8Box.Focus()
                    IPQ8Box.SelectAll()
                End If
            End If
            'Verhindern von eingabe mehrerer nullen
            If Asc(e.KeyChar) = 48 Then
                If IPQ7 = "0" Then
                    e.Handled = True
                End If
            End If




            If Asc(e.KeyChar) = 8 Then
                If IPQ7Box.Text = "" Then
                    IPQ6Box.Focus()
                End If
            End If

            Select Case Asc(e.KeyChar)
                Case 48 To 57, 8, 28 To 29
                    ' Zahlen, Backspace und Space zulassen
                Case Else
                    ' alle anderen Eingaben unterdrücken
                    e.Handled = True
            End Select
        End If
    End Sub
    Private Sub IPQ8Box_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles IPQ8Box.KeyPress
        'Fehlermeldung fals die IP nicht im IP Bereich (0-255) liegt
        Dim IPQ8 As String = 1
        If IPQ8Box.Text = "" Then
            IPQ8 = 1
        Else
            IPQ8 = IPQ8Box.Text
        End If
        Dim pruefung As Integer
        pruefung = IPQ8Box.SelectedText & Asc(e.KeyChar)
        If pruefung > 255 Then
            IPQ8Box.Text = ""
        Else
            If IPQ8 > 25 And Asc(e.KeyChar) >= 48 Or IPQ8 = 25 And Asc(e.KeyChar) >= 54 Or IPQ8 = "00" And Asc(e.KeyChar) = 48 Then
                MsgBox("Falsche Eingabe, geben sie bitte eine Zahl zwischen 0 und 255 ein", Buttons:=vbCritical, Title:="Falscher Zahlenbereich")
                e.Handled = True
            Else
            End If
            'verhindern von gehen ins nächste Kästchen wenn das jetztige Kästchen noch leer ist 
            If Asc(e.KeyChar) = 46 Then
                If IPQ8Box.Text = "" Then
                    e.Handled = True
                Else
                    Button1.Focus()
                End If
            End If
            'Verhindern von eingabe mehrerer nullen
            If Asc(e.KeyChar) = 48 Then
                If IPQ8 = "0" Then
                    e.Handled = True
                End If
            End If
            If Asc(e.KeyChar) = 13 Then
                If IPQ8Box.Text = "" Then
                Else
                    Call starten()
                End If
            End If



            If Asc(e.KeyChar) = 8 Then
                If IPQ8Box.Text = "" Then
                    IPQ7Box.Focus()
                End If
            End If

            Select Case Asc(e.KeyChar)
                Case 48 To 57, 8, 28 To 29
                    ' Zahlen, Backspace und Space zulassen
                Case Else
                    ' alle anderen Eingaben unterdrücken
                    e.Handled = True
            End Select
        End If
    End Sub
    Private Sub IPQ1Box_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IPQ1Box.TextChanged
        'wenn 3 Zahlen eingeben sind automatisch das nächste Feld als Focus nehmen
        If IPQ1Box.TextLength = 3 Then IPQ2Box.Focus()
        IPQ2Box.SelectAll()
    End Sub
    Private Sub IPQ2Box_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IPQ2Box.TextChanged
        'wenn 3 Zahlen eingeben sind automatisch das nächste Feld als Focus nehmen
        If IPQ2Box.TextLength = 3 Then IPQ3Box.Focus()
        IPQ3Box.SelectAll()
    End Sub
    Private Sub IPQ3Box_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IPQ3Box.TextChanged
        'wenn 3 Zahlen eingeben sind automatisch das nächste Feld als Focus nehmen
        If IPQ3Box.TextLength = 3 Then IPQ4Box.Focus()
        IPQ4Box.SelectAll()
    End Sub
    Private Sub IPQ4Box_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IPQ4Box.TextChanged
        'wenn 3 Zahlen eingeben sind automatisch das nächste Feld als Focus nehmen
        If IPQ4Box.TextLength = 3 Then IPQ5Box.Focus()
        IPQ5Box.SelectAll()
    End Sub
    Private Sub IPQ5Box_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IPQ5Box.TextChanged
        'wenn 3 Zahlen eingeben sind automatisch das nächste Feld als Focus nehmen
        If IPQ5Box.TextLength = 3 Then IPQ6Box.Focus()
        IPQ6Box.SelectAll()
    End Sub
    Private Sub IPQ6Box_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IPQ6Box.TextChanged
        'wenn 3 Zahlen eingeben sind automatisch das nächste Feld als Focus nehmen
        If IPQ6Box.TextLength = 3 Then IPQ7Box.Focus()
        IPQ7Box.SelectAll()
    End Sub
    Private Sub IPQ7Box_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IPQ7Box.TextChanged
        'wenn 3 Zahlen eingeben sind automatisch das nächste Feld als Focus nehmen
        If IPQ7Box.TextLength = 3 Then IPQ8Box.Focus()
        IPQ8Box.SelectAll()
    End Sub
    Private Sub IPQ8Box_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IPQ8Box.TextChanged
        'wenn 3 Zahlen eingeben sind automatisch das nächste Feld als Focus nehmen
        If IPQ8Box.TextLength = 3 Then Button1.Focus()
    End Sub
    Private Sub ipq1box_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IPQ1Box.Click
        'mit einem Click in das Feld alles Makieren
        IPQ1Box.SelectAll()
    End Sub
    Private Sub ipq2box_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IPQ2Box.Click
        'mit einem Click in das Feld alles Makieren
        IPQ2Box.SelectAll()
    End Sub
    Private Sub ipq3box_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IPQ3Box.Click
        'mit einem Click in das Feld alles Makieren
        IPQ3Box.SelectAll()
    End Sub
    Private Sub ipq4box_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IPQ4Box.Click
        'mit einem Click in das Feld alles Makieren
        IPQ4Box.SelectAll()
    End Sub
    Private Sub ipq5box_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IPQ5Box.Click
        'mit einem Click in das Feld alles Makieren
        IPQ5Box.SelectAll()
    End Sub
    Private Sub ipq6box_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IPQ6Box.Click
        'mit einem Click in das Feld alles Makieren
        IPQ6Box.SelectAll()
    End Sub
    Private Sub ipq7box_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IPQ7Box.Click
        'mit einem Click in das Feld alles Makieren
        IPQ7Box.SelectAll()
    End Sub
    Private Sub ipq8box_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IPQ8Box.Click
        'mit einem Click in das Feld alles Makieren
        IPQ8Box.SelectAll()
    End Sub
    Private Sub StartenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StartenToolStripMenuItem.Click
        'ruft das Programm starten auf
        Call starten()
    End Sub
    Private Sub BeendenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BeendenToolStripMenuItem.Click
        'ruft das Programm Ende auf
        Call ende()
    End Sub
    Private Sub ExportierenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExportierenToolStripMenuItem.Click
        Try
            'die Variablen die Ich benötige deklarieren
            Dim stop1 As Integer
            Dim zahler As Integer
            Dim fbd As New FolderBrowserDialog
            Dim tn As TreeNode
            Dim laenge As Integer
            Dim anleer As Integer
            Dim leer As String
            Dim ip As String
            Dim mac As Integer
            Dim pc As Integer
            Dim saveFileDialog1 As New SaveFileDialog()

            'fals die Treeview Box leer sagen das nichts exportiert werden kann
            If TreeView1.Nodes.Count = 0 Then
                MsgBox(" Zum Exportieren muss vorher ein Scan gemacht worden sein", Buttons:=vbCritical, Title:="Kein Scan Durchgeführt")
            Else
                'fals die Treeview Box nicht leer ist festlegen was man im speichen Fenster von Windows auswählen kann
                'in meinem Fall CSV
                saveFileDialog1.Filter = "Comma-separated values|*.csv"
                saveFileDialog1.Title = "Exportieren der Liste"
                'das Speichern Fenster von Windows Öffnen
                saveFileDialog1.ShowDialog()
                'zahler auf 0 setzen
                zahler = 0
                'die Variable Mywriter deklarien als variable die in die Datei schreibt 
                Dim myWriter As New StreamWriter(saveFileDialog1.FileName, True)
                'der Stop Variablen sagen wie viele Zeilen es zu Exportieren gibt
                stop1 = TreeView1.Nodes.Count
                Do
                    'der Variable tn eine Zeile von Treeview zuweisen
                    tn = TreeView1.Nodes(zahler)
                    'Die Zeile aufteilen in kleine Variablen und zwar immer bei einer Leerzweile
                    Dim strTeil() As String = tn.Text.Split(" ")
                    'die länge der IP ermitteln
                    laenge = Len(strTeil.First)
                    'die Leerzeilen zwischen IP und status ermitteln
                    anleer = 16 - laenge
                    leer = 2 * anleer
                    'der Variabelem IP den ersten Teil von Tn zuweisen
                    ip = strTeil.First
                    'der Variblen Mac sagen sie wo die mac steht in tn
                    mac = leer + 2
                    'der Variblen pc sagen sie wo die mac steht in tn
                    pc = mac + 2
                    'eine Zeile in Die CSV Datei Schreiben
                    myWriter.WriteLine(ip & ";" & strTeil(leer) & ";" & strTeil(mac) & ";" & strTeil(pc))
                    'den zahler hochzählen, als stopvariable
                    zahler = zahler + 1
                Loop While zahler < stop1
                'die CSV datei schließen damit sie nicht mehr in benutzung ist 
                myWriter.Close()
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub ImportierenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportierenToolStripMenuItem.Click
        Try
            'alle Variablem Deklarierem
            Dim fbd As New FolderBrowserDialog
            Dim strTeil(2) As String
            Dim laenge As Integer
            Dim anleer As Integer
            Dim leer As String
            Dim status As Integer
            Dim mac As Integer
            Dim openFileDialog1 As New OpenFileDialog()
            Dim pc As Integer
            'Treeview sagen er soll die Farben im Ganzen Hintergrund einbleden
            TreeView1.ShowLines = False
            TreeView1.FullRowSelect = True
            TreeView1.ShowRootLines = False

            'den Programm sagen es darf nur eine CSV Datei geöffnet werden
            openFileDialog1.Filter = "Comma-separated values|*.csv"
            openFileDialog1.Title = "Exportieren der Liste"
            'den Benutzer fragen wo die Datei liegt
            openFileDialog1.ShowDialog()
            'fs als variable die die Datei ausließt deklarieren
            Dim fs As New FileStream(openFileDialog1.FileName, FileMode.Open)
            'mit sr wird nun fs ausglesen
            Dim sr As New StreamReader(fs)

            'zeile wird als String deklariert
            Dim zeile As String
            'treeview leeren
            TreeView1.Nodes.Clear()

            Try
                'wiederhole solange in der Excel liste noch daten sind
                Do Until sr.Peek() = -1
                    'lese eine Zeile
                    zeile = sr.ReadLine()
                    'Teile diese Zeile immer wenn ein ; steht 
                    strTeil = zeile.Split(";")
                    'ermittle die Benötigten Leerzeilen zwischen IP und Status
                    laenge = Len(strTeil.First)
                    anleer = 16 - laenge
                    'setze Status auf 1 
                    status = 1
                    leer = New String(" ", 2 * anleer)
                    'die mac steht 1 neben dem Status in der Excel Liste
                    mac = status + 1
                    'der PCName steht eins neben der MAC
                    pc = mac + 1
                    'wenn der Status Teil der Zeile Online ist dann wird der Text in grün geschrieben sonst in Rot
                    If strTeil(status) = "Online" Then
                        TreeView1.Nodes.Add(strTeil.First & leer & strTeil(status) & "  " & strTeil(mac) & "  " & strTeil(pc)).BackColor = Color.Green
                    Else
                        TreeView1.Nodes.Add(strTeil.First & leer & strTeil(status) & "  " & strTeil(mac) & "  " & strTeil(pc)).BackColor = Color.Red
                    End If
                Loop
                'blende den Prüfen Button ein
                Button3.Visible = True
            Catch ex As Exception
                'bei nicht vorhanden sein der Datei wird  eine Fehlermeldung erzeugt die besagt, datei nicht vorhanden.
                MsgBox("Datei nicht vorhanden", Buttons:=vbCritical, Title:="Fehlende Datei")
            End Try
        Catch ex As Exception
            'wenn die Datei nicht eingelesen werden kann, zum beispiel weil sie noch geöffnet oder defekt ist kommt eine Fehlermeldung Datei Fehlerhaft oder Geöffnet
            MsgBox("Datei Fehlerhaft, oder noch geöffnet", Buttons:=vbCritical, Title:="Datei Fehler")
        End Try
    End Sub
    Private Sub VersionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VersionToolStripMenuItem.Click
        'Es öffnet sich ein Versionsbox mit der Version des Programms
        MsgBox(Prompt:="Version 1.0", Buttons:=vbQuestion, Title:="Version")
    End Sub
    Private Sub HerstellerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HerstellerToolStripMenuItem.Click
        'es Öffnet sich ein Hersteller Box mit Infos über den Hersteller
        MsgBox("Hersteller: Felix Langner" & vbCrLf & "Firma: Langner", Buttons:=vbOKOnly + vbInformation, Title:="Info")
    End Sub
    Private Sub ende()
        'Hier wird gefragt ob sie das Programm wirklich beenden wollen
        If MsgBox("Wirklich Beenden?", Buttons:=vbYesNo + vbQuestion, Title:="Beenden") = vbYes Then
            End
        End If
    End Sub
    Private Sub starten()
        'Die Start Button werden auf Abbrechen gesetzt
        Button1.Text = "Abbrechen"
        StartenToolStripMenuItem.Text = "Abbrechen"
        'Button3 (prüfen) wird wieder ausgeblendet
        Button3.Visible = False
        Button3.Text = "Prüfen"
        'die Exportier, Importier Funktionnen werden während des Durchlaufs dektivirt
        ExportierenToolStripMenuItem.Enabled = False
        ImportierenToolStripMenuItem.Enabled = False
        'es wird geprüft, ob das Programm vielleicht schon läuft
        If stop1 = 0 Then
            'läuft es nicht wird es gestartet
            If IPThread.ThreadState = ThreadState.Unstarted Then
                'die Variable stop wird auf 1 gesetzt, was symbolysiert, das Programm läuft
                stop1 = 1
                'Treeview wird geleert
                TreeView1.Nodes.Clear()
                'der Thread wird nur beim ersten mal gestartet
                Call IPThread.Start()
            Else
                'die Variable stop wird auf 1 gesetzt, was symbolysiert, das Programm läuft
                stop1 = 1
                'Treeview wird geleert
                TreeView1.Nodes.Clear()
                'da der Thread jetzt gestartet ist springt man nur in ihn hinein
                Dim IPThread As New Thread(AddressOf IPSub)
                Call IPThread.Start()
            End If
        Else
            'läuft das Programm schon, so wird gefragt, ob es beendet werden soll
            If MsgBox(Prompt:="Task wirklich Abbrechen?", Buttons:=vbYesNo + vbQuestion, Title:="Abbrechen") = vbNo Then
            Else
                'wenn ja so wird die Abbruch Variable auf 1 gesetzt
                abbruch = 1
                'und die Buttons wieder auf start
                Button1.Text = "Starten"
                StartenToolStripMenuItem.Text = "Starten"
            End If

        End If
    End Sub
    Private Sub Prüfen()
        Me.Invoke(Sub() ProgressBar1.Value = 0)
        'alle Benötigten Variabelen
        Dim zahleraufnehmen As Integer
        Dim laenge As Integer
        Dim anleer As Integer
        Dim leer As String
        Dim ip As String
        Dim mac As Integer
        Dim pingresult As String
        Dim ausgabe As String
        Dim counter As Integer
        Dim pc As Integer
        Dim cmd As String
        Dim cmd2 As String
        Dim macProzess As New Process
        Dim cmdmac As String
        Dim pcProzess As New Process
        Dim cmdpc As String
        Dim mac2 As String
        Dim pcname As String
        Dim offlinezahler As Integer
        Dim zahler1 As Integer
        mac2 = ""
        'der Counter ist dafür da um zu prüfen ob schon die gesamte Box geprüft wurde 
        counter = TreeView1.Nodes.Count
        'hier wird die Progressbar auf maximum gesetz, was in diesem Fall die anzahl der Zeilen von treeview oder auch dem Counter entspricht
        Me.Invoke(Sub() ProgressBar1.Maximum = counter)
        'durch die Schleife wird jede Zeile von Treeview Geprüft
        Do
            'treen werden die Zeilen der Reihenfolge nach eingelesen
            treen = TreeView1.Nodes(zahleraufnehmen)
            'hier wird die Zeile gesplittet
            Dim strTeil() As String = treen.Text.Split(" ")
            'die Länge der Leerzeilen zwischen Ip und Status wird ermittelt
            laenge = Len(strTeil.First)
            anleer = 16 - laenge
            leer = New String(" ", 2 * anleer)
            'der variablen ip wird der erste Teil zugwiesen
            ip = strTeil.First
            'die Mac steht 2 * die leerzeilen + 2 vom anfang weg
            mac = (2 * anleer) + 2
            'der Pc steht die Mac plus 2 weg
            pc = mac + 2

            Try
                'hier wird die Ip erneut geprüft
                pingresult = My.Computer.Network.Ping(ip)
                'der Variable Ausgabe wird entweder Online oder Offline zugewiesen
                If pingresult = False Then
                    ausgabe = "Offline"
                Else
                    ausgabe = "Online"
                End If

                'fals die Variabel Online ist wird die MAc und Der PC Name ermittelt 
                'infos auch IPsub. da dieser Bereich doppelt vorhanden ist
                If ausgabe = "Online" Then
                    cmd = "/c arp -a " & ip
                    macProzess.StartInfo.RedirectStandardOutput = True
                    macProzess.StartInfo.RedirectStandardInput = True
                    macProzess.StartInfo.CreateNoWindow = True
                    macProzess.StartInfo.UseShellExecute = False
                    macProzess.StartInfo.FileName = "cmd.exe"
                    macProzess.StartInfo.Arguments = cmd
                    macProzess.Start()
                    cmdmac = macProzess.StandardOutput.ReadToEnd()
                    Dim strTeil3() As String = cmdmac.Split(" ")
                    If strTeil3(2) = "---" Then
                        If strTeil3(27) = "" Then
                        Else
                            mac2 = strTeil3(27)
                        End If
                        If strTeil3(28) = "" Then
                        Else
                            mac2 = strTeil3(28)
                        End If
                        If strTeil3(29) = "" Then
                        Else
                            mac2 = strTeil3(29)
                        End If
                        If strTeil3(30) = "" Then
                        Else
                            mac2 = strTeil3(30)
                        End If
                        If strTeil3(31) = "" Then
                        Else
                            mac2 = strTeil3(31)
                        End If
                        If strTeil3(32) = "" Then
                        Else
                            mac2 = strTeil3(32)
                        End If
                    End If
                    cmd2 = "/c ping -a " & ip & " -n 1"
                    pcProzess.StartInfo.RedirectStandardOutput = True
                    pcProzess.StartInfo.RedirectStandardInput = True
                    pcProzess.StartInfo.CreateNoWindow = True
                    pcProzess.StartInfo.UseShellExecute = False
                    pcProzess.StartInfo.FileName = "cmd.exe"
                    pcProzess.StartInfo.Arguments = cmd2
                    pcProzess.Start()
                    cmdpc = pcProzess.StandardOutput.ReadToEnd()
                    Dim strTeil2() As String = cmdpc.Split(" ")
                    pcname = strTeil2(4)
                    zahler1 = zahler1 + 1
                Else
                    offlinezahler = offlinezahler + 1
                    mac2 = ""
                    pcname = ""
                End If



                Gesamtausgabe = ip & leer & ausgabe & "  " & mac2 & "  " & pcname & "   geprüft"
                If abbruch = 1 Then
                    Exit Do
                End If
                If ausgabe = "Online" Then
                    Me.Invoke(Sub() TreeView1.Nodes.Insert(zahleraufnehmen, Gesamtausgabe).BackColor = Color.Green)
                Else
                    Me.Invoke(Sub() TreeView1.Nodes.Insert(zahleraufnehmen, Gesamtausgabe).BackColor = Color.Red)
                End If
                Me.Invoke(Sub() TreeView1.Nodes.RemoveAt(zahleraufnehmen + 1))
                zahleraufnehmen = zahleraufnehmen + 1
                Me.Invoke(Sub() ProgressBar1.Value = ProgressBar1.Value + 1)
            Catch ex As Exception
            End Try
        Loop While zahleraufnehmen < counter
        Me.Invoke(Sub() Button3.Text = "Prüfen")
        Me.Invoke(Sub() Button1.Enabled = True)
        Me.Invoke(Sub() StartenToolStripMenuItem.Enabled = True)
        Me.Invoke(Sub() ExportierenToolStripMenuItem.Enabled = True)
        Me.Invoke(Sub() ImportierenToolStripMenuItem.Enabled = True)
        If abbruch = 1 Then
            Me.Invoke(Sub() ProgressBar1.Value = 0)
        End If
        abbruch = 0
        stop1 = 0

    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'bestimmte Button werden deaktiviert 
        Button1.Enabled = False
        StartenToolStripMenuItem.Enabled = False
        ExportierenToolStripMenuItem.Enabled = False
        ImportierenToolStripMenuItem.Enabled = False
        'alle Variablen werden deklariert
        Dim laenge As Integer
        Dim anleer As Integer
        Dim leer As String
        Dim ip As String
        Dim mac As Integer
        Dim counter As Integer
        Dim pc As Integer
        Dim zahler As Integer
        Dim status1 As String
        'der Variabel Counter wird Die anzahl der Zeilen von Counter zugewiesen
        counter = TreeView1.Nodes.Count
        'es wird geprüft ob das Programm noch läuft
        If stop1 = 0 Then
            'fals nicht, wird Button3 auf abbrechen gesetzt
            Button3.Text = "Abbruch"
            'und beim ersten durchlauf der Prüfenthread gestartet danach wird immer nur noch wieder in den Prüfenthread hinein gesprungen
            If prüfenThread.ThreadState = ThreadState.Unstarted Then
                'stop wird 1 gestzt, als symbol für das laufen das Programms
                stop1 = 1
                Call prüfenThread.Start()
            Else
                'stop wird 1 gestzt, als symbol für das laufen das Programms
                stop1 = 1
                Dim prüfenThread As New Thread(AddressOf Prüfen)
                Call prüfenThread.Start()
            End If
        Else
            'fals das Programm noch läuft, wird gefragt ob das Programm beendet werden soll
            If MsgBox(Prompt:="Task wirklich Abbrechen?", Buttons:=vbYesNo + vbQuestion, Title:="Abbrechen") = vbNo Then
            Else
                ' fals das Programm beendet werden soll, so wird die abbruch Variable auf 1 gesetzt
                abbruch = 1
                'und mit dieser Schleife werden die geprüften Zeilen zurück gesetzt
                Do
                    treen = TreeView1.Nodes(zahler)
                    Dim strTeil() As String = treen.Text.Split(" ")
                    laenge = Len(strTeil.First)
                    anleer = 16 - laenge
                    leer = New String(" ", 2 * anleer)
                    status1 = anleer * 2
                    ip = strTeil.First
                    mac = (2 * anleer) + 2
                    pc = mac + 2
                    Try
                        Gesamtausgabe = strTeil.First & leer & strTeil(status1) & "  " & strTeil(mac) & "  " & strTeil(pc)
                        If strTeil(status1) = "Online" Then
                            Me.Invoke(Sub() TreeView1.Nodes.Insert(zahler, Gesamtausgabe).BackColor = Color.Green)
                        Else
                            Me.Invoke(Sub() TreeView1.Nodes.Insert(zahler, Gesamtausgabe).BackColor = Color.Red)
                        End If
                        Me.Invoke(Sub() TreeView1.Nodes.RemoveAt(zahler + 1))
                        zahler = zahler + 1
                    Catch ex As Exception
                    End Try
                Loop While zahler < counter
            End If
        End If

    End Sub
    Private Sub Form1_formclosing(ByVal sender As System.Object, ByVal e As FormClosingEventArgs) Handles MyBase.FormClosing
        'soll das Programm wirklich beendet werden
        If MsgBox("Wirklich Beenden?", Buttons:=vbYesNo + vbQuestion, Title:="Beenden") = vbYes Then
            'fals ja beende es
            End
        Else
            'fals nein brech den befehlt ab
            e.Cancel = True
        End If
    End Sub
End Class