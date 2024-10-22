Attribute VB_Name = "ForsWinAPIWrapper"
Option Explicit

' Declare Windows API functions / Deklarácia funkcií Windows API
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As LongPtr, ByVal wParam As Long, lParam As Any) As LongPtr
Private Declare PtrSafe Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As LongPtr, ByVal wMapType As LongPtr) As LongPtr

' TODO: Add commands sheet, replace wrapper runCount with FindFirstEmptyRow and set it into an internal variable, Add Error handling / TODO: Pridat list príkazov, nahradit wrapper runCount funkciou FindFirstEmptyRow a nastavit ju do internej premennej, pridat spracovanie chýb

Public trnsactions As Transactions

' Define a type to hold information for the ForsWinAPIWrapper / Definovanie typu na uchovávanie informácií pre ForsWinAPIWrapper
Public Type t_ForsWinAPIWrapper
    userName As String
    serverName As String
    runCount As Range
    hwnd As LongPtr
    hexadecimalCodes As Object
    commands As Variant
End Type

' Define a type to hold command information / Definovanie typu na uchovávanie informácií o príkazoch
Private Type t_Command
    type As String
    value As Variant
    repeatCount As Integer
End Type

' Initialize the ForsWinAPIWrapper / Inicializácia ForsWinAPIWrapper
Public Function initForsWinAPIWrapper(ByRef wrapper As t_ForsWinAPIWrapper)
    Dim forsWindowTitle As String
    
    Set trnsactions = New Transactions
    
    ' Get the username from the environment variables / Získanie používatelského mena z environmentálnych premenných
    wrapper.userName = Environ("username")
    
    ' Retrieve server name and run count from the "Main" worksheet / Získanie názvu servera a poctu spustení z listu "Main"
    With ThisWorkbook.Worksheets("Main")
        wrapper.serverName = .Range("B3").value
        Set wrapper.runCount = .Range("A13")
    End With
    
    ' Construct the window title for the Fors window / Konštrukcia názvu okna pre okno Fors
    forsWindowTitle = wrapper.userName & "@lxhv1fors08[" & wrapper.serverName & "] || {}"
    
    ' Find the window handle using the constructed title / Nájst handle okna pomocou konštruovaného názvu
    wrapper.hwnd = FindWindow(vbNullString, forsWindowTitle)
    
    ' Initialize hexadecimal codes for shortcuts / Inicializácia hexadecimálnych kódov pre skratky
    initHexadecimalCodes wrapper
End Function

' Process the commands / Spracovanie príkazov
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

' This function sets the position based on the provided value / Táto funkcia nastaví pozíciu na základe poskytnutej hodnoty
Private Function setPosition(ByVal value As String, ByRef wrapper As t_ForsWinAPIWrapper)
    Dim positionParts() As String
    Dim moduleName As String
    Dim position As String
    
    ' Split the input value into parts based on the dot separator / Rozdelte vstupnú hodnotu na casti na základe bodky
    positionParts = Split(Split(value, "$")(1), ".")
    ' Convert the first part to uppercase to get the module name / Prevedte prvú cast na velké písmená, aby ste získali názov modulu
    moduleName = UCase(positionParts(0))
    ' Get the position name from the second part / Získajte názov pozície z druhej casti
    If UBound(positionParts) = 0 Then
        position = trnsactions.position(moduleName)("transaction")
    ElseIf UBound(positionParts) = 2 Then
        position = trnsactions.position(moduleName)(positionParts(1))(CInt(positionParts(2)))
    Else
        position = trnsactions.position(moduleName)(positionParts(1))
    End If
    
    
    
    ' Select the appropriate module and process the commands / Vyberte príslušný modul a spracujte príkazy
    processCommands wrapper, position
End Function

' Initialize a command / Inicializácia príkazu
Private Function initCommand(ByRef command As t_Command, ByVal value As String, ByVal hexadecimalCodes As Object)
    Select Case True
        ' Case for address commands, e.g., "&A1" / Prípad pre príkazy adresy, napr. "&A1"
        Case value Like "[&]*"
            command.value = value
            command.type = "Address"
            command.repeatCount = 0
        
        ' Case for repeated commands, e.g., "TEXT*3" or "CTRL*2" / Prípad pre opakované príkazy, napr. "TEXT*3" alebo "CTRL*2"
        Case value Like "*[*]*"
            Dim commandParts() As String

            ' Split the value into command and repeat count / Rozdelenie hodnoty na príkaz a pocet opakovaní
            commandParts = Split(value, "*")
            command.value = commandParts(0)
            command.repeatCount = CInt(commandParts(1))
            
            ' Check if the command is a shortcut / Kontrola, ci je príkaz skratkou
            If isShortcut(commandParts(0), hexadecimalCodes) Then
                command.type = "RepeatedShortcut"
            Else
                command.type = "RepeatedText"
            End If
        
        ' Case for single shortcut commands, e.g., "CTRL" / Prípad pre jednotlivé príkazy skratky, napr. "CTRL"
        Case isShortcut(value, hexadecimalCodes)
            command.value = value
            command.type = "Shortcut"
            command.repeatCount = 0
        
        ' Default case for text commands, e.g., "Hello" / Predvolený prípad pre textové príkazy, napr. "Hello"
        Case Else
            command.value = value
            command.type = "Text"
            command.repeatCount = 0
            
    End Select
End Function

' Execute a command / Vykonanie príkazu
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

' Handle address commands / Spracovanie príkazov adresy
Private Function address(ByVal value, ByVal hwnd)
    Dim column As String
    Dim row As Integer
    
    ' Extract column and row from the command value / Extrahovanie stlpca a riadku z hodnoty príkazu
    column = Mid(command.value, 2, 1)
    row = CInt(Mid(command.value, 3))
    
    ' Send the text / Odoslanie textu
    sendText ThisWorkbook.Worksheets("Data").Range(column & row).value, hwnd
End Function

' Send a shortcut to the window / Odoslanie skratky do okna
Private Function sendShortcut(ByVal value As String, ByVal hwnd As LongPtr, ByVal hexadecimalCodes As Object)
    Dim shortcutHexCode As LongPtr
    Dim shortcutSendCode As LongPtr
    
    ' Get the hexadecimal code and virtual key code for the shortcut / Získanie hexadecimálneho kódu a virtuálneho kódu klávesy pre skratku
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

' Send a single character to the window / Odoslanie jedného znaku do okna
Private Function sendCharacter(ByVal char As String, ByVal hwnd As LongPtr)
    Dim charAscii As Integer
    Dim charSendCode As LongPtr
    
    ' Get the ASCII code and virtual key code for the character / Získanie ASCII kódu a virtuálneho kódu klávesy pre znak
    charAscii = Asc(char)
    charSendCode = MapVirtualKey(charAscii, 0)
    
    ' Send the character to the window / Odoslanie znaku do okna
    SendMessage hwnd, &H102, charAscii, charSendCode
End Function

' Log information / Zaznamenanie informácií
Private Function log(ByVal info As String, ByRef wrapper As t_ForsWinAPIWrapper)
    ' Initialize run count if necessary / Inicializácia poctu spustení, ak je to potrebné
    If IsEmpty(ThisWorkbook.Worksheets("Logger").Range("A1")) Then
        wrapper.runCount.value = 0
    End If
    
    ' Increment run count and log the information / Zvýšenie poctu spustení a zaznamenanie informácií
    wrapper.runCount.value = wrapper.runCount.value + 1
    ThisWorkbook.Worksheets("Logger").Range("A" & wrapper.runCount.value) = constructLogEntry(wrapper, generateTimeStamp(), info)
End Function

' Generate a timestamp / Generovanie casovej peciatky
Private Function generateTimeStamp() As String
    generateTimeStamp = Format(Now(), "dd/mm/yyyy hh:mm:ss")
End Function

' Construct a log entry / Konštrukcia záznamu do logu
Private Function constructLogEntry(ByRef wrapper As t_ForsWinAPIWrapper, ByVal timeStamp As String, ByVal info As String) As String
    constructLogEntry = "[" & timeStamp & "] " & wrapper.userName & "@" & wrapper.serverName & " | " & info
End Function

' Check if a command is a shortcut / Kontrola, ci je príkaz skratkou
Private Function isShortcut(ByVal value As String, ByVal hexadecimalCodes As Object) As Boolean
    isShortcut = hexadecimalCodes.Exists(value)
End Function

' Initialize map with hexadecimal codes for shortcuts (shortcut to hexadecimal code pairs) / Inicializácia mapy s hexadecimálnymi kódmi pre skratky (skratka na hexadecimálny kód)
Private Function initHexadecimalCodes(ByRef wrapper As t_ForsWinAPIWrapper)
    Set wrapper.hexadecimalCodes = CreateObject("Scripting.Dictionary")
    
    With wrapper.hexadecimalCodes
        .CompareMode = vbTextCompare
        .Add "BACK", &H8 ' Backspace key / Kláves Backspace
        .Add "TAB", &H9 ' Tab key / Kláves Tab
        .Add "ENTER", &H9 ' Enter key / Kláves Enter
        .Add "SHIFT", &H10 ' Shift key / Kláves Shift
        .Add "CTRL", &H11 ' Control key / Kláves Control
        .Add "ALT", &H12 ' Alt key / Kláves Alt
        .Add "SPACE", &H20 ' Spacebar / Medzerník
        .Add "PAGEUP", &H21 ' Page Up key / Kláves Page Up
        .Add "PAGEDOWN", &H22 ' Page Down key / Kláves Page Down
        .Add "END", &H23 ' End key / Kláves End
        .Add "HOME", &H24 ' Home key / Kláves Home
        .Add "LEFT", &H25 ' Left arrow key / Kláves so šípkou dolava
        .Add "UP", &H26 ' Up arrow key / Kláves so šípkou nahor
        .Add "RIGHT", &H27 ' Right arrow key / Kláves so šípkou doprava
        .Add "DOWN", &H28 ' Down arrow key / Kláves so šípkou nadol
        .Add "INSERT", &H2D ' Insert key / Kláves Insert
        .Add "DELETE", &H2E ' Delete key / Kláves Delete
        .Add "F1", &H70 ' F1 key / Kláves F1
        .Add "F2", &H71 ' F2 key / Kláves F2
        .Add "F3", &H72 ' F3 key / Kláves F3
        .Add "F4", &H73 ' F4 key / Kláves F4
        .Add "F5", &H74 ' F5 key / Kláves F5
        .Add "F6", &H75 ' F6 key / Kláves F6
        .Add "F7", &H76 ' F7 key / Kláves F7
        .Add "F8", &H77 ' F8 key / Kláves F8
        .Add "F9", &H78 ' F9 key / Kláves F9
        .Add "F10", &H79 ' F10 key / Kláves F10
        .Add "F11", &H7A ' F11 key / Kláves F11
        .Add "F12", &H7B ' F12 key / Kláves F12
        .Add "LSHIFT", &HA0 ' Left Shift key / Lavý kláves Shift
        .Add "RSHIFT", &HA1 ' Right Shift key / Pravý kláves Shift
        .Add "LCONTROL", &HA2 ' Left Control key / Lavý kláves Control
        .Add "RCONTROL", &HA3 ' Right Control key / Pravý kláves Control
        .Add "NENTER", &H7B ' Numeric Enter key / Numerický kláves Enter
    End With
End Function

Public Function toChars(ByVal str As String) As Variant
    Dim chars As Variant
    chars = Split(StrConv(str, 64), Chr(0))
    ReDim Preserve chars(UBound(chars) - 1)
    
    toChars = chars
End Function
