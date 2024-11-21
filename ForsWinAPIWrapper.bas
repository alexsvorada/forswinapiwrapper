Attribute VB_Name = "ForsWinAPIWrapper"
Option Explicit

' Declare Windows API functions / Deklar�cia funkci� Windows API
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As LongPtr, ByVal wParam As Long, lParam As Any) As LongPtr
Private Declare PtrSafe Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As LongPtr, ByVal wMapType As LongPtr) As LongPtr

' TODO: Add commands sheet, replace wrapper runCount with FindFirstEmptyRow and set it into an internal variable, Add Error handling / TODO: Pridat list pr�kazov, nahradit wrapper runCount funkciou FindFirstEmptyRow a nastavit ju do internej premennej, pridat spracovanie ch�b

Public trnsactions As Transactions

' Define a type to hold information for the ForsWinAPIWrapper / Definovanie typu na uchov�vanie inform�ci� pre ForsWinAPIWrapper
Public Type t_ForsWinAPIWrapper
    userName As String
    serverName As String
    runCount As Range
    hwnd As LongPtr
    hexadecimalCodes As Object
    commands As Variant
End Type

' Define a type to hold command information / Definovanie typu na uchov�vanie inform�ci� o pr�kazoch
Private Type t_Command
    type As String
    value As Variant
    repeatCount As Integer
End Type

' Initialize the ForsWinAPIWrapper / Inicializ�cia ForsWinAPIWrapper
Public Function initForsWinAPIWrapper(ByRef wrapper As t_ForsWinAPIWrapper)
    Dim forsWindowTitle As String
    
    Set trnsactions = New Transactions
    
    ' Get the username from the environment variables / Z�skanie pou��vatelsk�ho mena z environment�lnych premenn�ch
    wrapper.userName = Environ("username")
    
    ' Retrieve server name and run count from the "Main" worksheet / Z�skanie n�zvu servera a poctu spusten� z listu "Main"
    With ThisWorkbook.Worksheets("Main")
        wrapper.serverName = .Range("B3").value
        Set wrapper.runCount = .Range("A13")
    End With
    
    ' Construct the window title for the Fors window / Kon�trukcia n�zvu okna pre okno Fors
    forsWindowTitle = wrapper.userName & "@lxhv1fors08[" & wrapper.serverName & "] || {}"
    
    ' Find the window handle using the constructed title / N�jst handle okna pomocou kon�truovan�ho n�zvu
    wrapper.hwnd = FindWindow(vbNullString, forsWindowTitle)
    
    ' Initialize hexadecimal codes for shortcuts / Inicializ�cia hexadecim�lnych k�dov pre skratky
    initHexadecimalCodes wrapper
End Function

' Process the commands / Spracovanie pr�kazov
Public Function processCommands(ByRef wrapper As t_ForsWinAPIWrapper, ByVal commandValues As Variant)
    Dim value As Variant
    Dim values As Variant
    
    values = Split(Trim(commandValues), ",")
    
    For Each value In values
        If (Trim(value) Like "[$]*") Then
            setPosition Trim(value), wrapper
        Else
            Dim command As t_Command
            initCommand command, Trim(value), wrapper.hexadecimalCodes
            executeCommand command, wrapper
        End If
    Next value
End Function

' This function sets the position based on the provided value / T�to funkcia nastav� poz�ciu na z�klade poskytnutej hodnoty
Private Function setPosition(ByVal value As String, ByRef wrapper As t_ForsWinAPIWrapper)
    Dim positionParts() As String
    Dim moduleName As String
    Dim position As String
    
    ' Split the input value into parts based on the dot separator / Rozdelte vstupn� hodnotu na casti na z�klade bodky
    positionParts = Split(Split(value, "$")(1), ".")
    ' Convert the first part to uppercase to get the module name / Prevedte prv� cast na velk� p�smen�, aby ste z�skali n�zov modulu
    moduleName = UCase(positionParts(0))
    ' Get the position name from the second part / Z�skajte n�zov poz�cie z druhej casti
    If UBound(positionParts) = 0 Then
        position = trnsactions.position(moduleName)("transaction")
    ElseIf UBound(positionParts) = 2 Then
        position = trnsactions.position(moduleName)(positionParts(1))(CInt(positionParts(2)))
    Else
        position = trnsactions.position(moduleName)(positionParts(1))
    End If
    
    
    
    ' Select the appropriate module and process the commands / Vyberte pr�slu�n� modul a spracujte pr�kazy
    processCommands wrapper, position
End Function

' Initialize a command / Inicializ�cia pr�kazu
Private Function initCommand(ByRef command As t_Command, ByVal value As String, ByVal hexadecimalCodes As Object)
    Select Case True
        ' Case for address commands, e.g., "&A1" / Pr�pad pre pr�kazy adresy, napr. "&A1"
        Case value Like "[&]*"
            command.value = value
            command.type = "Address"
            command.repeatCount = 0
        
        ' Case for repeated commands, e.g., "TEXT*3" or "CTRL*2" / Pr�pad pre opakovan� pr�kazy, napr. "TEXT*3" alebo "CTRL*2"
        Case value Like "*[*]*"
            Dim commandParts() As String

            ' Split the value into command and repeat count / Rozdelenie hodnoty na pr�kaz a pocet opakovan�
            commandParts = Split(value, "*")
            command.value = commandParts(0)
            command.repeatCount = CInt(commandParts(1))
            
            ' Check if the command is a shortcut / Kontrola, ci je pr�kaz skratkou
            If isShortcut(commandParts(0), hexadecimalCodes) Then
                command.type = "RepeatedShortcut"
            Else
                command.type = "RepeatedText"
            End If
        
        ' Case for single shortcut commands, e.g., "CTRL" / Pr�pad pre jednotliv� pr�kazy skratky, napr. "CTRL"
        Case isShortcut(value, hexadecimalCodes)
            command.value = value
            command.type = "Shortcut"
            command.repeatCount = 0
        
        ' Default case for text commands, e.g., "Hello" / Predvolen� pr�pad pre textov� pr�kazy, napr. "Hello"
        Case Else
            command.value = value
            command.type = "Text"
            command.repeatCount = 0
            
    End Select
End Function

' Execute a command / Vykonanie pr�kazu
Private Function executeCommand(ByRef command As t_Command, ByRef wrapper As t_ForsWinAPIWrapper)
    Dim i As Integer
    
    Select Case command.type
        Case "Address"
            address command.value, wrapper.hwnd
        Case "RepeatedShortcut"
            For i = 1 To command.repeatCount
                sendShortcut command.value, wrapper.hwnd, wrapper.hexadecimalCodes
            Next i
        Case "RepeatedText"
            For i = 1 To command.repeatCount
                sendText command.value, wrapper.hwnd
            Next i
        Case "Shortcut"
            sendShortcut command.value, wrapper.hwnd, wrapper.hexadecimalCodes
        Case "Text"
            sendText command.value, wrapper.hwnd
    End Select
    
    Select Case True
        Case command.repeatCount > 0
            log command.value & ": " & command.type & "*" & command.repeatCount, wrapper
        Case Else
            log command.value & ": " & command.type, wrapper
    End Select
End Function

' Handle address commands / Spracovanie pr�kazov adresy
Private Function address(ByVal value, ByVal hwnd)
    Dim column As String
    Dim row As Integer
    
    ' Extract column and row from the command value / Extrahovanie stlpca a riadku z hodnoty pr�kazu
    column = Mid(command.value, 2, 1)
    row = CInt(Mid(command.value, 3))
    
    ' Send the text / Odoslanie textu
    sendText ThisWorkbook.Worksheets("Data").Range(column & row).value, hwnd
End Function

' Send a shortcut to the window / Odoslanie skratky do okna
Private Function sendShortcut(ByVal value As String, ByVal hwnd As LongPtr, ByVal hexadecimalCodes As Object)
    Dim shortcutHexCode As LongPtr
    Dim shortcutSendCode As LongPtr
    
    ' Get the hexadecimal code and virtual key code for the shortcut / Z�skanie hexadecim�lneho k�du a virtu�lneho k�du kl�vesy pre skratku
    shortcutHexCode = hexadecimalCodes.Item(value)
    shortcutSendCode = MapVirtualKey(shortcutHexCode, 0)
    
    ' Send the shortcut to the window / Odoslanie skratky do okna
    SendMessage hwnd, &H101, shortcutHexCode, shortcutSendCode
End Function

' Send text to the window / Odoslanie textu do okna
Private Function sendText(ByVal value As String, ByVal hwnd As LongPtr)
    Dim char As Variant
    Dim i As Integer
    
    For Each char In toChars(value)
        sendCharacter char, hwnd
    Next char
End Function

' Send a single character to the window / Odoslanie jedn�ho znaku do okna
Private Function sendCharacter(ByVal char As String, ByVal hwnd As LongPtr)
    Dim charAscii As Integer
    Dim charSendCode As LongPtr
    
    ' Get the ASCII code and virtual key code for the character / Z�skanie ASCII k�du a virtu�lneho k�du kl�vesy pre znak
    charAscii = Asc(char)
    charSendCode = MapVirtualKey(charAscii, 0)
    
    ' Send the character to the window / Odoslanie znaku do okna
    SendMessage hwnd, &H102, charAscii, charSendCode
End Function

' Log information / Zaznamenanie inform�ci�
Private Function log(ByVal info As String, ByRef wrapper As t_ForsWinAPIWrapper)
    ' Initialize run count if necessary / Inicializ�cia poctu spusten�, ak je to potrebn�
    If IsEmpty(ThisWorkbook.Worksheets("Logger").Range("A1")) Then
        wrapper.runCount.value = 0
    End If
    
    ' Increment run count and log the information / Zv��enie poctu spusten� a zaznamenanie inform�ci�
    wrapper.runCount.value = wrapper.runCount.value + 1
    ThisWorkbook.Worksheets("Logger").Range("A" & wrapper.runCount.value) = constructLogEntry(wrapper, generateTimeStamp(), info)
End Function

' Generate a timestamp / Generovanie casovej peciatky
Private Function generateTimeStamp() As String
    generateTimeStamp = Format(Now(), "dd/mm/yyyy hh:mm:ss")
End Function

' Construct a log entry / Kon�trukcia z�znamu do logu
Private Function constructLogEntry(ByRef wrapper As t_ForsWinAPIWrapper, ByVal timeStamp As String, ByVal info As String) As String
    constructLogEntry = "[" & timeStamp & "] " & wrapper.userName & "@" & wrapper.serverName & " | " & info
End Function

' Check if a command is a shortcut / Kontrola, ci je pr�kaz skratkou
Private Function isShortcut(ByVal value As String, ByVal hexadecimalCodes As Object) As Boolean
    isShortcut = hexadecimalCodes.Exists(value)
End Function

' Initialize map with hexadecimal codes for shortcuts (shortcut to hexadecimal code pairs) / Inicializ�cia mapy s hexadecim�lnymi k�dmi pre skratky (skratka na hexadecim�lny k�d)
Private Function initHexadecimalCodes(ByRef wrapper As t_ForsWinAPIWrapper)
    Set wrapper.hexadecimalCodes = CreateObject("Scripting.Dictionary")
    
    With wrapper.hexadecimalCodes
        .CompareMode = vbTextCompare
        .Add "BACK", &H8 ' Backspace key / Kl�ves Backspace
        .Add "TAB", &H9 ' Tab key / Kl�ves Tab
        .Add "ENTER", &H9 ' Enter key / Kl�ves Enter
        .Add "SHIFT", &H10 ' Shift key / Kl�ves Shift
        .Add "CTRL", &H11 ' Control key / Kl�ves Control
        .Add "ALT", &H12 ' Alt key / Kl�ves Alt
        .Add "SPACE", &H20 ' Spacebar / Medzern�k
        .Add "PAGEUP", &H21 ' Page Up key / Kl�ves Page Up
        .Add "PAGEDOWN", &H22 ' Page Down key / Kl�ves Page Down
        .Add "END", &H23 ' End key / Kl�ves End
        .Add "HOME", &H24 ' Home key / Kl�ves Home
        .Add "LEFT", &H25 ' Left arrow key / Kl�ves so ��pkou dolava
        .Add "UP", &H26 ' Up arrow key / Kl�ves so ��pkou nahor
        .Add "RIGHT", &H27 ' Right arrow key / Kl�ves so ��pkou doprava
        .Add "DOWN", &H28 ' Down arrow key / Kl�ves so ��pkou nadol
        .Add "INSERT", &H2D ' Insert key / Kl�ves Insert
        .Add "DELETE", &H2E ' Delete key / Kl�ves Delete
        .Add "F1", &H70 ' F1 key / Kl�ves F1
        .Add "F2", &H71 ' F2 key / Kl�ves F2
        .Add "F3", &H72 ' F3 key / Kl�ves F3
        .Add "F4", &H73 ' F4 key / Kl�ves F4
        .Add "F5", &H74 ' F5 key / Kl�ves F5
        .Add "F6", &H75 ' F6 key / Kl�ves F6
        .Add "F7", &H76 ' F7 key / Kl�ves F7
        .Add "F8", &H77 ' F8 key / Kl�ves F8
        .Add "F9", &H78 ' F9 key / Kl�ves F9
        .Add "F10", &H79 ' F10 key / Kl�ves F10
        .Add "F11", &H7A ' F11 key / Kl�ves F11
        .Add "F12", &H7B ' F12 key / Kl�ves F12
        .Add "LSHIFT", &HA0 ' Left Shift key / Lav� kl�ves Shift
        .Add "RSHIFT", &HA1 ' Right Shift key / Prav� kl�ves Shift
        .Add "LCONTROL", &HA2 ' Left Control key / Lav� kl�ves Control
        .Add "RCONTROL", &HA3 ' Right Control key / Prav� kl�ves Control
        .Add "NENTER", &H7B ' Numeric Enter key / Numerick� kl�ves Enter
    End With
End Function

Public Function toChars(ByVal str As String) As Variant
    Dim chars As Variant
    chars = Split(StrConv(str, 64), Chr(0))
    ReDim Preserve chars(UBound(chars) - 1)
    
    toChars = chars
End Function
