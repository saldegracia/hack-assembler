Dim Shared symbol$(0 To 5000)
Dim Shared address%(0 To 5000)
Dim Shared count%
Dim Shared nextVarAddress%

count% = 0
InitSymbolTable
For x = 0 To count%
    Print symbol$(x);
    Print address%(x)
Next x

' --- Pass 1 ---
Print "Enter .asm filename:" 'Enter file name
Input filename$
base$ = Left$(filename$, _InStrRev(filename$, ".") - 1) 'create output file name
out$ = base$ + ".hack"
On Error GoTo errorhandler 'Error message if file does not exist
Open filename$ For Input As #1
romAddress% = 0
Do While Not EOF(1)
    Line Input #1, line$
    cleanline$ = StripLine$(line$)
    If cleanline$ <> "" Then
        If Left$(cleanline$, 1) = "(" And Right$(cleanline$, 1) = ")" Then
            ' Add label to symbol table
            label$ = Mid$(cleanline$, 2, Len(cleanline$) - 2)
            AddSymbol label$, romAddress%
        Else
            romAddress% = romAddress% + 1
        End If
    End If
Loop
Close #1

' --- Pass 2 ---
Open filename$ For Input As #1
Open out$ For Output As #2

Do While Not EOF(1)
    'open file to read line by line
    Line Input #1, line$
    'remove comments and whitespace from line
    cleanline$ = StripLine$(line$)
    'Skip empty lines and write new lines into array
    If cleanline$ <> "" Then
        ' Skip label lines
        If Left$(cleanline$, 1) <> "(" Then
            'Detect instruction type and convert to binary
            first$ = Left$(cleanline$, 1)
            If first$ = "@" Then
                binary$ = AInst$(cleanline$)
            Else
                binary$ = CInst$(cleanline$)
            End If
            Print cleanline$, binary$
            Print #2, binary$ 'print binary values into .hack file
        End If
    End If
Loop

Print ""
Print "File processing complete!"
Close #1
Close #2
End

errorhandler:
Print "File not found!"

Function StripLine$ (raw$)
    'Detect and remove comment
    comment = InStr(raw$, "//")
    If comment > 0 Then
        raw$ = Left$(raw$, comment - 1)
    End If
    'Remove whitespace
    raw$ = _Trim$(raw$)
    ' Remove all spaces
    While InStr(raw$, " ") > 0 'returns 0 when no blank is found
        blank = InStr(raw$, " ")
        raw$ = Left$(raw$, blank - 1) + Mid$(raw$, blank + 1)
    Wend
    StripLine$ = raw$
End Function

Function AInst$ (line$) 'converts A-instruction into binary value
    decimal$ = Mid$(line$, 2) 'remove @ from string

    If IsNumber%(decimal$) Then
        decimal = Val(decimal$) 'convert string to integer
    Else
        decimal = ResolveSymbol%(decimal$)
    End If
    bin$ = "" 'convert integer to binary (string)
    Do While decimal > 0
        bin$ = LTrim$(Str$(decimal Mod 2)) + bin$
        decimal = decimal \ 2
    Loop
    zero$ = "0"
    Do While Len(bin$) < 16 'insert 0s to create 16bit binary number
        bin$ = zero$ + bin$
    Loop
    AInst$ = bin$
End Function

Function CInst$ (line$)
    splitPos = InStr(line$, ";")
    If splitPos > 0 Then 'Jump instruction
        Comp$ = Left$(line$, splitPos - 1) 'Comp bits
        Jump$ = Mid$(line$, splitPos + 1) 'Dest bits
        DestBin$ = "000"
        CompBin$ = CompBits$(Comp$)
        JumpBin$ = JumpBits$(Jump$)
    ElseIf InStr(line$, "=") > 0 Then 'Non-jump instruction"
        splitPos = InStr(line$, "=")
        Dest$ = Left$(line$, splitPos - 1) 'Dest bits
        Comp$ = Mid$(line$, splitPos + 1) 'Comp bits
        DestBin$ = DestBits$(Dest$)
        CompBin$ = CompBits$(Comp$)
        JumpBin$ = "000"
    Else
        Print "No valid assembler instruction found!"
    End If
    CInst$ = "111" + CompBin$ + DestBin$ + JumpBin$
End Function

Function CompBits$ (Comp$) 'Create Comp part
    Select Case Comp$
        Case Is = "0"
            CompBits$ = "0101010"
        Case Is = "1"
            CompBits$ = "0111111"
        Case Is = "-1"
            CompBits$ = "0111010"
        Case Is = "D"
            CompBits$ = "0001100"
        Case Is = "A"
            CompBits$ = "0110000"
        Case Is = "!D"
            CompBits$ = "0001101"
        Case Is = "!A"
            CompBits$ = "0110001"
        Case Is = "-D"
            CompBits$ = "0001111"
        Case Is = "-A"
            CompBits$ = "0110011"
        Case Is = "D+1"
            CompBits$ = "0011111"
        Case Is = "A+1"
            CompBits$ = "0110111"
        Case Is = "D-1"
            CompBits$ = "0001110"
        Case Is = "A-1"
            CompBits$ = "0110010"
        Case Is = "D+A"
            CompBits$ = "0000010"
        Case Is = "D-A"
            CompBits$ = "0010011"
        Case Is = "A-D"
            CompBits$ = "0000111"
        Case Is = "D&A"
            CompBits$ = "0000000"
        Case Is = "D|A"
            CompBits$ = "0010101"
        Case Is = "M"
            CompBits$ = "1110000"
        Case Is = "!M"
            CompBits$ = "1110001"
        Case Is = "-M"
            CompBits$ = "1110011"
        Case Is = "M+1"
            CompBits$ = "1110111"
        Case Is = "M-1"
            CompBits$ = "1110010"
        Case Is = "D+M"
            CompBits$ = "1000010"
        Case Is = "D-M"
            CompBits$ = "1010011"
        Case Is = "M-D"
            CompBits$ = "1000111"
        Case Is = "D&M"
            CompBits$ = "1000000"
        Case Is = "D|M"
            CompBits$ = "1010101"
        Case Else
            Print "Comp error!"
    End Select
End Function

Function DestBits$ (Dest$) 'Create Dest part
    Select Case Dest$
        Case Is = " "
            DestBits$ = "000"
        Case Is = "M"
            DestBits$ = "001"
        Case Is = "D"
            DestBits$ = "010"
        Case Is = "MD"
            DestBits$ = "011"
        Case Is = "A"
            DestBits$ = "100"
        Case Is = "AM"
            DestBits$ = "101"
        Case Is = "AD"
            DestBits$ = "110"
        Case Is = "AMD"
            DestBits$ = "111"
        Case Else
            Print "Dest error!"
    End Select
End Function

Function JumpBits$ (Jump$) 'Create Jump part
    Select Case Jump$
        Case Is = " "
            JumpBits$ = "000"
        Case Is = "JGT"
            JumpBits$ = "001"
        Case Is = "JEQ"
            JumpBits$ = "010"
        Case Is = "JGE"
            JumpBits$ = "011"
        Case Is = "JLT"
            JumpBits$ = "100"
        Case Is = "JNE"
            JumpBits$ = "101"
        Case Is = "JLE"
            JumpBits$ = "110"
        Case Is = "JMP"
            JumpBits$ = "111"
        Case Else
            Print "Jump error!"
    End Select
End Function

Sub InitSymbolTable 'Creates table with predefined symbols
    count% = 0
    nextVarAddress% = 16
    AddSymbol "SP", 0
    AddSymbol "LCL", 1
    AddSymbol "ARG", 2
    AddSymbol "THIS", 3
    AddSymbol "THAT", 4
    For a = 0 To 15
        name$ = "R" + LTrim$(Str$(a))
        AddSymbol name$, a
    Next a
    AddSymbol "SCREEN", 16384
    AddSymbol "KBD", 24576
End Sub

Sub AddSymbol (name$, addr%) 'adds new symbol to table
    symbol$(count%) = name$
    address%(count%) = addr%
    count% = count% + 1

End Sub

Function FindSymbolIndex% (name$) 'Checks if symbol is in table, returns address or -1 if not in table
    'Symbol not found
    FindSymbolIndex% = -1
    For n = 0 To count% - 1
        If name$ = symbol$(n) Then
            FindSymbolIndex% = n
            Exit Function
        End If
    Next n
End Function

Function ResolveSymbol% (name$)
    Dim idx%

    ' Look for symbol in the table
    idx% = FindSymbolIndex%(name$)
    'Print "Resolve:", name$, " idx=", idx%
    If idx% <> -1 Then
        ' Symbol already exists ? return its address
        ResolveSymbol% = address%(idx%)
    Else
        ' New variable ? assign next free RAM address
        AddSymbol name$, nextVarAddress%
        ResolveSymbol% = nextVarAddress%
        nextVarAddress% = nextVarAddress% + 1
    End If
End Function

Function IsNumber% (s$)
    Dim i%, ch$

    If s$ = "" Then
        IsNumber% = 0
        Exit Function
    End If

    For i% = 1 To Len(s$)
        ch$ = Mid$(s$, i%, 1)
        If ch$ < "0" Or ch$ > "9" Then
            IsNumber% = 0
            Exit Function
        End If
    Next i%

    IsNumber% = -1 ' QB TRUE
End Function











