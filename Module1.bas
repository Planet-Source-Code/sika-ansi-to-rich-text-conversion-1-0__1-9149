Attribute VB_Name = "Module1"
'                           STANDARD ANSI TABLE
'                       --------------------------
'                       Some of the information
'                       Supplied by the table may
'                       be incorrect or inaccurate
'                       You may modify it to best
'                       Fit your needs.

'   ESC[#;#;#mText
'   |________||___Text in which is effect by the ANSI Protocol
'   |--
    '   \>format of ANSI Protocol
    
'   Example: Esc[0;32mString will display a green normal text of "String"
'            Esc[1;4;32mString will display bold underlined green text of "String"
'    ##   Effect                     Implementable [using standard RTF textbox]
'   ===============================================
'    0   For normal display             *
'    1   For bold on                    *
'    4   underline (mono only)          *
'    5   blink On
'    7   reverse video On
'    8   nondisplayed (invisible)       *

'   30   black foreground               *
'   31   red foreground                 *
'   32   green foreground               *
'   33   yellow foreground              *
'   34   blue foreground                *
'   35   magenta foreground             *
'   36   cyan foreground                *
'   37   white foreground               *

'   40   black background
'   41   red background
'   42   green background
'   43   yellow background
'   44   blue background
'   45   magenta background
'   46   cyan background
'   47   white background

Dim ac_bold As String
Dim ac_underline As String
Dim ac_color As String
Dim Esc As String
Dim ExtractIndicator(), ExtractMove(), ExtractExpirence(), ExtractStyle As String
Sub OutputText(Rich As RichTextBox, ansicode As String)
Dim isEscape As Boolean
Dim sBuffer As String, i As Integer
Dim curChr As String
Dim j As Integer, modes() As String
Dim INFO() As String, RawDATA As String

'                 \ /
' HUGE thanks to >vcv<
'                 / \
' visit www.dosfx.com for a great site for programming related stuff
' visit www.dosfx.com/~thezone if you want to hang out in a knowledgable chat room.heh heh heh

ansicode = Replace(ansicode, "ÿ", "")   '}\
ansicode = Replace(ansicode, "", "")   '}} \_I know this ***t means something and I figured out a couple of them but I dont find use for them in my programs.
ansicode = Replace(ansicode, "û", "")   '}} /_They are combined to tell the client to do a certain task. Example 'ÿRJ' tells the client that a password is going to be entered
ansicode = Replace(ansicode, "ü", "")   '}/
ansicode = Replace(ansicode, Chr(13), "")
Esc = Chr(27)
isEscape = False
For i = 1 To Len(ansicode)
    curChr = Mid(ansicode, i, 1)
    If curChr = "[" And isEscape Then GoTo nex
    If curChr = Esc Then
        Rich.SelText = sBuffer
        Rich.SelStart = Len(Rich.Text)
        sBuffer = ""
        isEscape = True
    ElseIf curChr = "m" And isEscape Then
        If sBuffer = "" Then isEscape = False: GoTo nex
        modes = Split(sBuffer, ";")
        For j = LBound(modes) To UBound(modes)
            If IsNumeric(modes(j)) = False Then GoTo nex
            Select Case modes(j)
                Case 0:     If Rich.SelBold Then Rich.SelBold = False
                Case 4:     If Rich.SelUnderline = False Then Rich.SelUnderline = True
                Case 30:    Rich.SelColor = vbBlack
                Case 31:    Rich.SelColor = vbRed
                Case 32:    Rich.SelColor = vbGreen
                Case 33:    Rich.SelColor = vbYellow
                Case 34:    Rich.SelColor = vbBlue
                Case 35:    Rich.SelColor = vbMagenta
                Case 36:    Rich.SelColor = vbCyan
                Case 37:    Rich.SelColor = vbWhite
               '--------------------------------------------------------------------------------------
               'Case #*:    effect it has on text
               '*see the chart above or use a different value if your server supports it.
               ' If you customize those values for your server you can integrate font and sound
               ' support in your server dont ask me how to make your own ANSI in C++ cause i dont know how
               ' For example if you wish to use bold text integrated with this you would...
               ' Case 1(symbolizes bold text according to table above):If Rich.SelBold = False Then Rich.SelBold = True
               '---------------------------------------------------------------------------------------
            End Select
        Next j
        sBuffer = ""
        isEscape = False
    Else
        sBuffer = sBuffer & curChr
        If i = Len(ansicode) Then Rich.SelText = sBuffer
    End If
nex:
Next i
End Sub

