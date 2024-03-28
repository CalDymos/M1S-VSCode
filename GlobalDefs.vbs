''' <summary>Returns the absolute value of a number.</summary>
''' <param name="Number">Any valid numeric expression.</param>
Function Abs(ByVal Number As Variant) As Double
End Function

''' <summary>Returns a Variant containing an array.</summary>
''' <param name="arglist">arglist argument is a comma-delimited list of values that are assigned to the elements of an array</param>
Function Array(ByVal arglist As Variant) As Variant
End Function

''' <summary>Returns the unicode code of a character.</summary>
''' <param name="String">The character to get the code for. If a string is used, the code for the first character is given.</param>
Function Asc(ByVal String As String) As Integer
End Function

''' <summary>Returns the arctangent of a number.</summary>
Function Atn(ByVal Number As Variant) As Double
End Function
''' <summary>Returns an expression that has been converted to a Variant of subtype Boolean.</summary>
Function CBool(ByVal expression As Variant) As Boolean
End Function

Function CDate(ByVal expression As Variant) As Date
End Function

Function CDbl(expr) ' As Double
End Function

Function Chr(charcode)
End Function

Function ChrB(charcode)
End Function

Function ChrW(charcode)
End Function

Function CInt(expr) ' As Integer
End Function

Function CLng(expr) ' As Long
End Function

Function Cos(number)
End Function

Function CreateObject(classname)
End Function

Function CreateObject(classname, location)
End Function

Function CSng(expr) ' As Single
End Function

Function CStr(expr) ' As String
End Function

Function Date()
End Function

''' <summary>Returns a date to which a specified time interval has been added.</summary>
''' <param name="interval">String expression that is the interval you want to add</param>
''' <param name="number">Numeric expression that is the number of interval you want to add. The numeric expression can either be positive, for dates in the future, or negative, for dates in the past.</param>
''' <param name="date">Variant or literal representing the date to which interval is added.</param>
Function DateAdd(interval, number, date)
End Function

''' <summary>Returns the number of intervals between two dates.</summary>
''' <param name="interval">String expression that is the interval you want to use to calculate the differences between date1 and date2</param>
''' <param name="date1">Date expressions. Two dates you want to use in the calculation.</param>
''' <param name="date2">Date expressions. Two dates you want to use in the calculation.</param>
Function DateDiff(interval, date1, date2)
End Function

''' <summary>Returns the number of intervals between two dates.</summary>
''' <param name="interval">String expression that is the interval you want to use to calculate the differences between date1 and date2</param>
''' <param name="date1">Date expressions. Two dates you want to use in the calculation.</param>
''' <param name="date2">Date expressions. Two dates you want to use in the calculation.</param>
''' <param name="firstdayofweek">Constant that specifies the day of the week. If not specified, Sunday is assumed</param>
Function DateDiff(interval, date1, date2, firstdayofweek)
End Function

''' <summary>Returns the number of intervals between two dates.</summary>
''' <param name="interval">String expression that is the interval you want to use to calculate the differences between date1 and date2</param>
''' <param name="date1">Date expressions. Two dates you want to use in the calculation.</param>
''' <param name="date2">Date expressions. Two dates you want to use in the calculation.</param>
''' <param name="firstdayofweek">Constant that specifies the day of the week. If not specified, Sunday is assumed</param>
''' <param name="firstweekofyear">Constant that specifies the first week of the year. If not specified, the first week is assumed to be the week in which January 1 occurs</param>
Function DateDiff(interval, date1, date2, firstdayofweek, firstweekofyear)
End Function

''' <summary>Returns the specified part of a given date.</summary>
''' <param name="interval">String expression that is the interval of time you want to return</param>
''' <param name="date">Date expression you want to evaluate</param>
Function DatePart(interval, date)
End Function

''' <summary>Returns the specified part of a given date.</summary>
''' <param name="interval">String expression that is the interval of time you want to return</param>
''' <param name="date">Date expression you want to evaluate</param>
''' <param name="firstdayofweek">Constant that specifies the day of the week. If not specified, Sunday is assumed</param>
Function DatePart(interval, date, firstdayofweek)
End Function

''' <summary>Returns the specified part of a given date.</summary>
''' <param name="interval">String expression that is the interval of time you want to return</param>
''' <param name="date">Date expression you want to evaluate</param>
''' <param name="firstdayofweek">Constant that specifies the day of the week. If not specified, Sunday is assumed</param>
''' <param name="firstweekofyear">Constant that specifies the first week of the year. If not specified, the first week is assumed to be the week in which January 1 occurs</param>
Function DatePart(interval, date, firstdayofweek, firstweekofyear)
End Function

''' <summary>Returns a Variant of subtype Date for a specified year, month, and day.</summary>
''' <param name="year">Number between 100 and 9999, inclusive, or a numeric expression.</param>
Function DateSerial(year, month, day)
End Function

''' <summary>Returns a Variant of subtype Date.</summary>
Function DateValue(date)
End Function

''' <summary>Returns a whole number between 1 and 31, inclusive, representing the day of the month.</summary>
Function Day(date)
End Function

''' <summary>Returns</summary>
Function Escape(str) ' As String
End Function

''' <summary>Evaluates an expression and returns the result.</summary>
Function Eval(expr)
End Function

''' <summary>Returns e (the base of natural logarithms) raised to a power.</summary>
Function Exp(number)
End Function

''' <summary>Returns a zero-based array containing a subset of a string array based on a specified filter criteria.</summary>
''' <param name="InputStrings">One-dimensional array of strings to be searched.</param>
''' <param name="Value">String to search for.</param>
Function Filter(InputStrings, Value)
End Function

''' <summary>Returns a zero-based array containing a subset of a string array based on a specified filter criteria.</summary>
''' <param name="InputStrings">One-dimensional array of strings to be searched.</param>
''' <param name="Value">String to search for.</param>
''' <param name="Include">Boolean value indicating whether to return substrings that include or exclude Value. If Include is True, Filter returns the subset of the array that contains Value as a substring. If Include is False, Filter returns the subset of the array that does not contain Value as a substring.</param>
Function Filter(InputStrings, Value, Include)
End Function

''' <summary>Returns a zero-based array containing a subset of a string array based on a specified filter criteria.</summary>
''' <param name="InputStrings">One-dimensional array of strings to be searched.</param>
''' <param name="Value">String to search for.</param>
''' <param name="Include">Boolean value indicating whether to return substrings that include or exclude Value. If Include is True, Filter returns the subset of the array that contains Value as a substring. If Include is False, Filter returns the subset of the array that does not contain Value as a substring.</param>
''' <param name="Compare">Numeric value indicating the kind of string comparison to use</param>
Function Filter(InputStrings, Value, Include, Compare)
End Function

Function Fix(number)
End Function

''' <summary>Returns an expression formatted as a currency value using the currency symbol defined in the system control panel.</summary>
''' <param name="Expression">Expression to be formatted.</param>
Function FormatCurrency(Expression) ' As String
End Function

''' <summary>Returns an expression formatted as a currency value using the currency symbol defined in the system control panel.</summary>
''' <param name="Expression">Expression to be formatted.</param>
''' <param name="NumDigitsAfterDecimal">Numeric value indicating how many places to the right of the decimal are displayed. Default value is -1, which indicates that the computer's regional settings are used</param>
Function FormatCurrency(Expression, NumDigitsAfterDecimal) ' As String
End Function

''' <summary>Returns an expression formatted as a currency value using the currency symbol defined in the system control panel.</summary>
''' <param name="Expression">Expression to be formatted.</param>
''' <param name="NumDigitsAfterDecimal">Numeric value indicating how many places to the right of the decimal are displayed. Default value is -1, which indicates that the computer's regional settings are used</param>
''' <param name="IncludeLeadingDigit">Tristate constant that indicates whether or not a leading zero is displayed for fractional values</param>
Function FormatCurrency(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit) ' As String
End Function

''' <summary>Returns an expression formatted as a currency value using the currency symbol defined in the system control panel.</summary>
''' <param name="Expression">Expression to be formatted.</param>
''' <param name="NumDigitsAfterDecimal">Numeric value indicating how many places to the right of the decimal are displayed. Default value is -1, which indicates that the computer's regional settings are used</param>
''' <param name="IncludeLeadingDigit">Tristate constant that indicates whether or not a leading zero is displayed for fractional values</param>
''' <param name="UseParensForNegativeNumbers">Tristate constant that indicates whether or not to place negative values within parentheses</param>
Function FormatCurrency(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers) ' As String
End Function

''' <summary>Returns an expression formatted as a currency value using the currency symbol defined in the system control panel.</summary>
''' <param name="Expression">Expression to be formatted.</param>
''' <param name="NumDigitsAfterDecimal">Numeric value indicating how many places to the right of the decimal are displayed. Default value is -1, which indicates that the computer's regional settings are used</param>
''' <param name="IncludeLeadingDigit">Tristate constant that indicates whether or not a leading zero is displayed for fractional values</param>
''' <param name="UseParensForNegativeNumbers">Tristate constant that indicates whether or not to place negative values within parentheses</param>
''' <param name="GroupDigits">Tristate constant that indicates whether or not numbers are grouped using the group delimiter specified in the computer's regional settings</param>
Function FormatCurrency(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits) ' As String
End Function

Function FormatDateTime(Date) ' As String
End Function

Function FormatDateTime(Date, NamedFormat) ' As String
End Function

Function FormatNumber(Expression) ' As String
End Function

Function FormatNumber(Expression, NumDigitsAfterDecimal) ' As String
End Function

Function FormatNumber(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit) ' As String
End Function

Function FormatNumber(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers) ' As String
End Function

Function FormatNumber(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits) ' As String
End Function

Function FormatPercent(Expression) ' As String
End Function

Function FormatPercent(Expression, NumDigitsAfterDecimal) ' As String
End Function

Function FormatPercent(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit) ' As String
End Function

Function FormatPercent(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers) ' As String
End Function

Function FormatPercent(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits) ' As String
End Function

Function GetLocale() ' As Long
End Function

''' <summary>?</summary>
Function GetObject(pathname) ' As Object
End Function

''' <summary>?</summary>
Function GetObject(pathname, classname) ' As Object
End Function

Function GetRef(procname)
End Function

Function GetUILanguage() ' As Integer
End Function

Function Hex(number) ' As String
End Function

''' <summary>Returns a whole number between 0 and 23, inclusive, representing the hour of the day.</summary>
Function Hour(time) ' As Integer
End Function

''' <summary>Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns the contents of the text box.</summary>
Function InputBox(prompt)
End Function

''' <summary>Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns the contents of the text box.</summary>
Function InputBox(prompt, title)
End Function

''' <summary>Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns the contents of the text box.</summary>
Function InputBox(prompt, title, default)
End Function

''' <summary>Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns the contents of the text box.</summary>
Function InputBox(prompt, title, default, xpos)
End Function

''' <summary>Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns the contents of the text box.</summary>
Function InputBox(prompt, title, default, xpos, ypos)
End Function

''' <summary>Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns the contents of the text box.</summary>
Function InputBox(prompt, title, default, xpos, ypos, helpfile, context)
End Function

''' <summary>Returns the position of the first occurrence of one string within another.</summary>
''' <param name="string1">The string to be searched</param>
''' <param name="string2">The string expression to search for</param>
Function InStr(string1, string2) ' As Long
End Function

''' <summary>Returns the position of the first occurrence of one string within another.</summary>
''' <param name="start">Specifies the starting position for search</param>
''' <param name="string1">The string to be searched</param>
''' <param name="string2">The string expression to search for</param>
Function InStr(start, string1, string2) ' As Long
End Function

''' <summary>Returns the position of the first occurrence of one string within another.</summary>
''' <param name="string1">The string to be searched</param>
''' <param name="string2">The string expression to search for</param>
''' <param name="compare">Specifies the string comparison to use</param>
Function InStr(string1, string2, compare) ' As Long
End Function

''' <summary>Returns the position of the first occurrence of one string within another.</summary>
''' <param name="start">Specifies the starting position for search</param>
''' <param name="string1">The string to be searched</param>
''' <param name="string2">The string expression to search for</param>
''' <param name="compare">Specifies the string comparison to use</param>
Function InStr(start, string1, string2, compare) ' As Long
End Function

''' <summary>Returns the position of the first occurrence of one string within another.</summary>
Function InStrB(string1, string2) ' As Long
End Function
''' <summary>Returns the position of the first occurrence of one string within another.</summary>
Function InStrB(start, string1, string2) ' As Long
End Function
''' <summary>Returns the position of the first occurrence of one string within another.</summary>
Function InStrB(string1, string2, compare) ' As Long
End Function
''' <summary>Returns the position of the first occurrence of one string within another.</summary>
Function InStrB(start, string1, string2, compare) ' As Long
End Function

''' <summary>Returns the position of an occurrence of one string within another, from the end of string.</summary>
Function InStrRev(string1, string2) ' As Long
End Function

''' <summary>Returns the position of an occurrence of one string within another, from the end of string.</summary>
Function InStrRev(string1, string2, start) ' As Long
End Function

''' <summary>Returns the position of an occurrence of one string within another, from the end of string.</summary>
Function InStrRev(string1, string2, start, compare) ' As Long
End Function

Function Int(number)
End Function

Function IsArray(var) ' As Boolean
End Function

Function IsDate(expr) ' As Boolean
End Function

Function IsEmpty(expr) ' As Boolean
End Function

Function IsNull(expr) ' As Boolean
End Function

''' <summary>Returns a Boolean value indicating whether an expression can be evaluated as a number.</summary>
Function IsNumeric(expr) ' As Boolean
End Function

Function IsObject(expr) ' As Boolean
End Function

''' <summary>Returns a string created by joining a number of substrings contained in an array.</summary>
Function Join(list) ' As String
End Function

''' <summary>Returns a string created by joining a number of substrings contained in an array.</summary>
Function Join(list, delimiter) ' As String
End Function

Function LBound(arrayname)
End Function

Function LBound(arrayname, dimension)
End Function

Function LCase(str) ' As String
End Function

Function Left(str, length) ' As String
End Function

Function LeftB(str, length) ' As String
End Function

''' <summary>Returns the number of characters in a string or the number of bytes required to store a variable.</summary>
Function Len(str) ' As Long
End Function

Function LenB(str) ' As Long
End Function

Function LoadPicture(picturename)
End Function

Function Log(number)
End Function

Function LTrim(str) ' As String
End Function

''' <summary>Returns a specified number of characters from a string.</summary>
Function Mid(str, start) ' As String
End Function

''' <summary>Returns a specified number of characters from a string.</summary>
Function Mid(str, start, length) ' As String
End Function

''' <summary>Returns a specified number of characters from a string.</summary>
Function MidB(str, start) ' As String
End Function

''' <summary>Returns a specified number of characters from a string.</summary>
Function MidB(str, start, length) ' As String
End Function

Function Minute(time) ' As Integer
End Function

Function Month(date) ' As Integer
End Function

Function MonthName(date) ' As String
End Function

Function MonthName(date, abbrevation) ' As String
End Function

''' <summary>Displays a message in a dialog box, waits for the user to click a button, and returns a value indicating which button the user clicked.</summary>
''' <param name="prompt">String expression displayed as the message in the dialog box</param>
Function MsgBox(prompt)
End Function

''' <summary>Displays a message in a dialog box, waits for the user to click a button, and returns a value indicating which button the user clicked.</summary>
''' <param name="prompt">String expression displayed as the message in the dialog box</param>
''' <param name="buttons">Numeric expression that is the sum of values specifying the number and type of buttons to display, the icon style to use, the identity of the default button, and the modality of the message box. If omitted, the default value for buttons is 0.</param>
Function MsgBox(prompt, buttons)
End Function

''' <summary>Displays a message in a dialog box, waits for the user to click a button, and returns a value indicating which button the user clicked.</summary>
''' <param name="prompt">String expression displayed as the message in the dialog box</param>
''' <param name="buttons">Numeric expression that is the sum of values specifying the number and type of buttons to display, the icon style to use, the identity of the default button, and the modality of the message box. If omitted, the default value for buttons is 0.</param>
''' <param name="title">String expression displayed in the title bar of the dialog box. If you omit title, the application name is placed in the title bar.</param>
Function MsgBox(prompt, buttons, title)
End Function

''' <summary>Displays a message in a dialog box, waits for the user to click a button, and returns a value indicating which button the user clicked.</summary>
''' <param name="prompt">String expression displayed as the message in the dialog box</param>
''' <param name="buttons">Numeric expression that is the sum of values specifying the number and type of buttons to display, the icon style to use, the identity of the default button, and the modality of the message box. If omitted, the default value for buttons is 0.</param>
''' <param name="title">String expression displayed in the title bar of the dialog box. If you omit title, the application name is placed in the title bar.</param>
''' <param name="helpfile">String expression that identifies the Help file to use to provide context-sensitive Help for the dialog box. If helpfile is provided, context must also be provided. Not available on 16-bit platforms.</param>
''' <param name="context">Numeric expression that identifies the Help context number assigned by the Help author to the appropriate Help topic. If context is provided, helpfile must also be provided. Not available on 16-bit platforms.</param>
Function MsgBox(prompt, buttons, title, helpfile, context)
End Function

''' <summary>Returns the current date and time according to the setting of your computer's system date and time.</summary>
Function Now ' As Date
End Function

Function Oct(number) ' As String
End Function

''' <summary>Returns a string in which a specified substring has been replaced with another substring.</summary>
Function Replace(str, find, replacewith) ' As String
End Function

''' <summary>Returns a string in which a specified substring has been replaced with another substring.</summary>
Function Replace(str, find, replacewith, start) ' As String
End Function

''' <summary>Returns a string in which a specified substring has been replaced with another substring a specified number of times.</summary>
Function Replace(str, find, replacewith, start, count) ' As String
End Function

''' <summary>Returns a string in which a specified substring has been replaced with another substring a specified number of times.</summary>
Function Replace(str, find, replacewith, start, count, compare) ' As String
End Function

''' <summary>Returns a whole number representing an RGB color value.</summary>
Function RGB(red, green, blue) ' As Long
End Function

Function Right(str, length) ' As String
End Function

Function RightB(str, length) ' As String
End Function

Function Rnd()
End Function

Function Rnd(number)
End Function

Function Round(number, digits)
End Function

Function RTrim(str) ' As String
End Function

Function Second(time)
End Function

''' <summary>undocumented</summary>
Function SetLocale(int)
End Function

Function Sgn(number)
End Function

Function Sin(number)
End Function

Function Space(number) ' As String
End Function

Function Sqr(number)
End Function

Function StrComp(string1, string2)
End Function

Function StrComp(string1, string2, compare)
End Function

Function Tan(number)
End Function

Function Time
End Function

Function Timer
End Function

Function TimeSerial(hour, minute, second)
End Function

Function TimeValue(time)
End Function

Function Trim(str) ' As String
End Function

Function TypeName(var) ' As String
End Function

''' <summary>Returns the largest available subscript for the indicated dimension of an array.</summary>
Function UBound(arrayname) ' As Long
End Function

''' <summary>Returns the largest available subscript for the indicated dimension of an array.</summary>
Function UBound(arrayname, dimension) ' As Long
End Function

''' <summary>Returns a string that has been converted to uppercase.</summary>
Function UCase(str) ' As String
End Function

Function Unescape(str) ' As String
End Function

Function VarType(var) ' as Integer
End Function

Function Weekday(date) ' as Integer
End Function

Function Weekday(date, firstdayofweek) ' as Integer
End Function

Function WeekdayName(weekday) ' As String
End Function

Function WeekdayName(weekday, abbreviate) ' As String
End Function

Function WeekdayName(weekday, abbreviate, firstdayofweek) ' As String
End Function

''' <summary>Returns a whole number representing the year.</summary>
''' <param name="date">Any expression that can represent a date</param>
Function Year(date)
End Function

'Mach 3 Functions and Subs

Sub StraightFeed(ByVal x As Double, ByVal Y As Double, ByVal Z As Double, ByVal a As Double, ByVal B As Double, ByVal C As Double)
End Sub

Sub HelpAbout()
End Sub

Function GetSafeZ() As Double
End Function

Sub SetSafeZ(ByVal SafeZ As Double)
End Sub

Sub SetCurrentTool(ByVal Tool As Integer)
End Sub

Function GetSelectedTool() As Integer
End Function

Function GetToolChangeStart(ByVal Axis As Integer) As Double
End Function

Sub StraightTraverse(ByVal x As Double, ByVal Y As Double, ByVal Z As Double, ByVal a As Double, ByVal B As Double, ByVal C As Double)
End Sub

Function ToolLengthOffset() As Double
End Function

Function CommandedFeed() As Double
End Function

Sub SetFeedRate(ByVal Rate As Double)
End Sub

Sub ActivateSignal(ByVal Signal As Integer)
End Sub

Function IsActive(ByVal Signal As Integer) As Byte
End Function

Sub DeActivateSignal(ByVal Signal As Integer)
End Sub

Sub SystemWaitFor(ByVal Signal As Integer)
End Sub

Function Param1() As Double
End Function

Function Param2() As Double
End Function

Function Param3() As Double
End Function

Sub VerifyAxis(ByVal Silent As Long)
End Sub

Function GetVar(ByVal var As Integer) As Double
End Function

Sub SetVar(ByVal var As Integer, ByVal Value As Double)
End Sub

Sub OpenDigFile()
End Sub

Sub CloseDigFile()
End Sub

Sub THCOn()
End Sub

Sub THCOff()
End Sub

Sub Code(ByVal Command As String)
End Sub

Function GetScale(ByVal Axis As Integer) As Double
End Function

Sub SetScale(ByVal Axis As Integer, ByVal dScale As Double)
End Sub

Sub SendSerial(ByVal sData As String)
End Sub

Function GetPortByte(ByVal PortAddress As Integer) As Byte
End Function

Function PutPortByte(ByVal PortAddress As Integer, ByVal data As Byte) As Integer
End Function

Function IsMoving() As Integer
End Function

Sub ResetTHC()
End Sub

Function GetParam(ByVal Param As String) As Double
End Function

Sub SetParam(ByVal Param As String, ByVal Value As Double)
End Sub

Sub setobj(ByVal thevar As String, ByVal thevalue As Double)
End Sub

Function GetCurrentTool() As Integer
End Function

Sub DoOEMButton(ByVal Button As Integer)
End Sub

Sub LoadRun(ByVal Filename As String)
End Sub

Sub DoButton(ByVal ButtonNum As Integer)
End Sub

Function GetLED(ByVal Led As Integer) As Boolean
End Function

Sub SetOEMLED(ByVal Led As Integer, ByVal state As Integer)
End Sub

Function GetOEMLed(ByVal Led As Integer) As Boolean
End Function

Function GetOEMDRO(ByVal dro As Integer) As Double
End Function

Function GetDRO(ByVal dro As Integer) As Double
End Function

Sub SetDRO(ByVal dro As Integer, ByVal Value As Double)
End Sub

Sub SetOEMDRO(ByVal dro As Integer, ByVal Value As Double)
End Sub

Sub GetCoord(ByVal Question As String)
End Sub

Function GetXCoor() As Double
End Function

Function GetYCoor() As Double
End Function

Function GetZCoor() As Double
End Function

Function GetACoor() As Double
End Function

Function IsFirst() As Boolean
End Function

Sub SetMachZero(ByVal Axis As Integer)
End Sub

Sub RunFile()
End Sub

Sub SetUserDRO(ByVal dro As Integer, ByVal Value As Double)
End Sub

Sub SetUserLED(ByVal Led As Integer, ByVal state As Integer)
End Sub

Function GetUserDRO(ByVal dro As Integer) As Double
End Function

Function GetUserLED(ByVal Led As Integer) As Integer
End Function

Function tXStart() As Double
End Function

Function tZStart() As Double
End Function

Function tEndX() As Double
End Function

Function tEndZ() As Double
End Function

Function tClearX() As Double
End Function

Function tLead() As Double
End Function

Function tSpring() As Integer
End Function

Function tPasses() As Integer
End Function

Function tChamfer() As Double
End Function

Function tTaper() As Double
End Function

Function tInFeed() As Double
End Function

Function tDepthLastPass() As Double
End Function

Function IsLoading() As Integer
End Function

Function GetABSPostion(Axis) As Double
End Function

Function GetRPM() As Double
End Function

Sub DoSpinCW()
End Sub

Sub DoSpinCCW()
End Sub

Sub DoSpinStop()
End Sub

Sub SetSpinSpeed(ByVal rpm As Double)
End Sub

Sub InFeeds(ByVal iteration As Integer)
End Sub

Sub ThreadDepth(ByVal iteration As Integer)
End Sub

Sub tTapers(ByVal iteration As Integer)
End Sub

Sub SetTicker(ByVal TickerNum As Integer, ByVal Message As String)
End Sub

Sub SetUserLabel(ByVal LabelNum As Integer, ByVal LabelText As String)
End Sub

Sub RefCombination(ByVal combo As Integer)
End Sub

Sub IsSuchSignal(ByVal Signal As Integer)
End Sub

Sub OpenTeachFile(ByVal name As String)
End Sub

Sub CloseTeachFile()
End Sub

Sub SetButtonText(ByVal Text As String)
End Sub

Sub LoadStandardLayout()
End Sub

Sub LoadTeachFile()
End Sub

Function IsOutputActive(ByVal Signal As Integer) As Boolean
End Function

Sub SingleVerify()
End Sub

Function GetPage() As Integer
End Function

Sub SetPage(ByVal page As Integer)
End Sub

Sub ToggleScreens()
End Sub

Sub PlayWave(ByVal wavename As String)
End Sub

Sub Speak(ByVal Text As String)
End Sub

Sub Message(ByVal Text As String)
End Sub

Function IsStopped() As Integer
End Function

Sub SaveWizard()
End Sub

Sub SingleVerifyReport(ByVal Axis As Integer)
End Sub

Sub SetIJMode(ByVal mode As Integer)
End Sub

Function GetIJMode() As Integer
End Function

Function GetMinPass() As Double
End Function

Function AppendTeachFile(ByVal name As String) As Integer
End Function

Function Random() As Double
End Function

Function IsDiameter() As Integer
End Function

Function ModGetInputWord(ByVal bit As Integer) As Integer
End Function

Function FillFromCoil(ByVal slave As Integer, ByVal startAddress As Integer, ByVal nBytes As Integer) As Integer
End Function

Function FillFromStatus(ByVal slave As Integer, ByVal startAddress As Integer, ByVal nBytes As Integer) As Integer
End Function

Function FillFromHolding(ByVal slave As Integer, ByVal startAddress As Integer, ByVal nBytes As Integer) As Integer
End Function

Function FillFromInput(ByVal slave As Integer, ByVal startAddress As Integer, ByVal nBytes As Integer) As Integer
End Function

Function GetModWord(ByVal index As Integer) As Integer
End Function

Function tMinDepth() As Double
End Function

Function tGetCutType() As Integer
End Function

Function tGetInfeedType() As Integer
End Function

Sub tSetCutType(ByVal iType As Integer)
End Sub

Sub tSetInFeedType(ByVal iType As Integer)
End Sub

Function OpenSubroutineFile(ByVal name As String) As Integer
End Function

Function tZClear() As Double
End Function

Function FeedRate() As Double
End Function

Function tCutDepth() As Double
End Function

Function Filename() As String
End Function

Function MinX() As Double
End Function

Function MaxX() As Double
End Function

Function MinY() As Double
End Function

Function MaxY() As Double
End Function

Function RetractMode() As Integer
End Function

Sub GotoSafeZ()
End Sub

Sub SetPulley(ByVal pulley As Integer)
End Sub

Sub LoadWizard(ByVal name As String)
End Sub

Sub SetOutput(ByVal word As Integer)
End Sub

Function GetInput(ByVal word As Integer)
End Function

Sub SetModOutput(ByVal reg As Integer, ByVal Value As Integer)
End Sub

Sub LoadFile(ByVal Filename As String)
End Sub

Sub SetTriggerMacro(ByVal macro As Integer)
End Sub

Sub SetHomannString(ByVal x As Integer, ByVal Y As Integer, ByVal Text As String)
End Sub

Sub JogOn(ByVal Axis As Integer, ByVal dir As Integer)
End Sub

Sub JogOff(ByVal Axis As Integer)
End Sub

Sub OpenSubFile(ByVal Filename As String)
End Sub

Sub SetToolX(ByVal Value As Double)
End Sub

Sub SetToolZ(ByVal Value As Double)
End Sub

Function GetTurretAng() As Double
End Function

Function GetFiFoEntry() As String
End Function

Sub RunProgram(ByVal Program As String)
End Sub

Sub SendFIFO(ByVal com As String)
End Sub

Function Round(val) As Double
End Function

Function roun(val) As Double
End Function

Function nFmt(ByVal val As Double, ByVal DEC As Integer) As Double
End Function

Sub ZeroTHC()
End Sub

Function IsWiz() As Integer
End Function

Sub WaitForPoll()
End Sub

Sub SetModIOString(ByVal slave As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal OutText As String)
End Sub

Sub Sleep(ByVal iTime As Integer)
End Sub

Sub SendFIFOcr(ByVal data As String)
End Sub

Sub SendFIFObyte(ByVal data As Integer)
End Sub

Sub SendFIFOword(ByVal data As Integer)
End Sub

Sub SetTimer(ByVal iTimer As Integer)
End Sub

Function GetTimer(ByVal iTimer As Integer) As Double
End Function

Function AskTextQuestion(ByVal Question As String) As String
End Function

Function GetDROString(ByVal nDRO As Integer)
End Function

Sub LoadLinTable()
End Sub

Sub SwapAxis(ByVal Primary As Integer, ByVal Secondary As Integer)
End Sub

Sub ResetAxisSwap()
End Sub

Sub SetFormula(ByVal Formula As String, ByVal Axis As Integer)
End Sub

Sub IsProbing()
End Sub

Function GetInBit(ByVal index As Integer, ByVal bit As Integer) As Integer
End Function

Sub SetOutBit(ByVal index As Integer, ByVal bit As Integer)
End Sub

Sub ResetOutBit(ByVal index As Integer, ByVal bit As Integer)
End Sub

Sub SetProbeActive(ByVal level As Integer)
End Sub

Sub StartVideo()
End Sub

Sub VBTest()
End Sub

Sub StartTHC()
End Sub

Sub EndTHC()
End Sub

Sub CoupleSlave(ByVal state As Integer)
End Sub

Function GetToolParam(ByVal toolnum As Integer, ByVal Param As Integer) As Double
End Function

Sub SetToolParam(ByVal toolnum As Integer, ByVal Param As Integer, ByVal data As Double)
End Sub

Sub SetProbeState(ByVal state As Integer)
End Sub

Function GetMainFolder() As String
End Function

Sub DoMenu(ByVal MenuIndex As Integer, ByVal MenuItem As Integer)
End Sub

Function GetMasterInput(ByVal Param As Integer) As Integer
End Function

Function GetMasterOutput(ByVal Param As Integer) As Integer
End Function

Sub SetMasterOutput(ByVal Param As Integer, ByVal Value As Integer)
End Sub

Sub SetMasterInput(ByVal Param As Integer, ByVal Value As Integer)
End Sub

Function GetUserLabel(ByVal nLabel As Integer) As String
End Function

Function GetToolDesc(ByVal Tool As Integer) As String
End Function

Sub NotifyPlugins(ByVal Message As Integer)
End Sub

Function IsEstop() As Integer
End Function

Function IsSafeZ() As Integer
End Function

Sub MachMSG(ByVal Msg As String, ByVal Title As String, ByVal iType As Integer)
End Sub

Function GetMessage() As String
End Function

Sub SetModPlugString(ByVal Message As String, ByVal cfg As Integer, ByVal address As Integer)
End Sub

Sub ClearFiFo()
End Sub

Function GetMachVersion(ByRef major As Integer, ByRef minor As Integer, ByRef Build As Integer) As Boolean
End Function

Function SetToolDesc(ByVal ToolNumber As Integer, ByVal Description As String) As Boolean
End Function

Function RunScript(ByVal macro As String) As Boolean
End Function

Function GetActiveProfileName() As String
End Function

Function NumberPad(ByVal MessageString As String) As Boolean
End Function

Function GetMyWindowsHandle() As Long
End Function

Function GetLoadedGCodeDir() As String
End Function

Function GetLoadedGCodeFileName() As String
End Function

Function GetActiveProfileDir() As String
End Function

Function GetActiveScreenSetName() As String
End Function

Function IncludeTLOinZFromG31() As Boolean
End Function

Function ProgramSafetyLockout() As Boolean
End Function

Function GetSetupUnits() As Integer
End Function

Function StartPeriodicScript(ByVal ScriptPath As String, ByVal UpdateTime As Double) As Boolean
End Function

Function StopPeriodicScript(ByVal ScriptPath As String) As Boolean
End Function

Function IsPeriodicScriptRunning(ByVal PathString As String) As Boolean
End Function

Function GetIODevName(ByVal DeviceNumber As Integer) As String
End Function

Function SetDevOutput(ByVal DeviceID As Integer, ByVal OutputNum As Integer, ByVal Value As Integer) As Integer
End Function

Function GetIODevInput(ByVal DeviceID As Integer, ByVal InputNum As Integer) As Double
End Function

Function GetIODevIOName(ByVal DeviceID As Integer, ByVal InputNumber As Integer) As String
End Function

Function SetMachPos(ByVal Axis As Integer, ByVal Value As Double) As Integer
End Function

Function SetInputData(ByVal InputNumber As Integer, ByVal PinNumber As Integer, ByVal ActiveLow) As Integer
End Function

Function SetIODevDataPointers(ByVal DeviceID As Integer, ByVal PinNumber As Integer, ByVal SetValueDRO As Integer, ByVal SetOffsetDRO As Integer) As Integer
End Function

Sub NoDelaySpindle(ByVal StopForSpeedChange As Integer)
End Sub

Sub SetEnaContext(ByVal Ctx As Long)
End Sub

Sub WriteProtocol(ByVal Value As String, Optional bRuntimeInfo As Boolean = True)
End Sub


''' Enum VbVarType
Const vbEmpty = 0
Const vbNull = 1
Const vbInteger = 2
Const vbLong = 3
Const vbSingle = 4
Const vbDouble = 5
Const vbCurrency = 6
Const vbDate = 7
Const vbString = 8
Const vbObject = 9
Const vbError = 10
Const vbBoolean = 11
Const vbVariant = 12
Const vbDataObject = 13
Const vbDecimal = 14
Const vbByte = 17
Const vbArray = 8192
''' End Enum ' VbVarType

Const Nothing = Nothing
Const Empty = Empty ' The Empty keyword is used to indicate an uninitialized variable value. This is not the same thing as Null. You can use the IsEmpty Function to determine whether a variable is initialized.
Const Null = Null

Const False = False ' Boolean
Const True = True

''' Enum VbTriState
Const vbUseDefault = -2
Const vbTrue = -1
Const vbFalse = 0
''' End Enum

''' Enum VbCompareMethod
Const vbBinaryCompare = 0 ' Perform a binary comparison
Const vbTextCompare = 1 ' Perform a textual comparison
Const vbDatabaseCompare = 2 ' Only in Access
''' End Enum

Const vbCr = Chr(13)
Const vbCrLf = Chr(13) & Chr(10)
Const vbFormFeed = Chr(12)
Const vbLf = Chr(10)
Const vbNewLine = Chr(13) & Chr(10)
Const vbNullChar = Chr(0)
Const vbNullString = Empty
Const vbTab = Chr(9)
Const vbVerticalTab = Chr(11)

''' Enum VbDateTimeFormat
Const vbGeneralDate = 0
Const vbLongDate = 1 ' Display a date using the long date format specified in your computer's regional settings.
Const vbShortDate = 2 ' Display a date using the short date format specified in your computer's regional settings.
Const vbLongTime = 3 ' Display a time using the time format specified in your computer's regional settings.
Const vbShortTime = 4
''' End Enum

''' Enum VbFirstWeekOfYear
Const vbUseSystemDayOfWeek = 0
Const vbFirstJan1 = 1
Const vbFirstFourDays = 2
Const vbFirstFullWeek = 3
''' End Enum

Const vbObjectError = &h80040000 ' User-defined error numbers should be greater than this value.

''' Enum VbDayOfWeek
Const vbMonday = 2
Const vbTuesday = 3
Const vbWednesday = 4
Const vbThursday = 5
Const vbFriday = 6
Const vbSaturday = 7
Const vbSunday = 1
Const vbUseSystem = 0
''' End Enum

''' Enum VbMsgBoxStyle
Const vbOKOnly = 0 ' Display OK button only.
Const vbOKCancel = 1 ' Display OK and Cancel buttons
Const vbAbortRetryIgnore = 2
Const vbYesNoCancel = 3
Const vbYesNo = 4
Const vbRetryCancel = 5 ' Display Retry and Cancel buttons.
Const vbCritical = 16
Const vbQuestion = 32 ' Display Warning Query icon.
Const vbExclamation = 48
Const vbInformation = 64
Const vbDefaultButton1 = 0
Const vbDefaultButton2 = 256
Const vbDefaultButton3 = 512
Const vbDefaultButton4 = 768
Const vbApplicationModal = 0
Const vbSystemModal         = &h01000
Const vbMsgBoxHelpButton    = &h04000
Const VbMsgBoxSetForeground = &h010000
Const vbMsgBoxRight         = &h080000
Const vbMsgBoxRtlReading    = &h100000
''' End Enum

''' Enum VbMsgBoxResult
Const vbOK = 1
Const vbCancel = 2
Const vbAbort = 3
Const vbRetry = 4 ' Retry button was clicked
Const vbIgnore = 5
Const vbYes = 6
Const vbNo = 7
''' End Enum

''' Enum ColorConstants
Const vbBlack   = &h000000
Const vbBlue    = &hFF0000
Const vbCyan    = &hFFFF00
Const vbGreen   = &h00FF00
Const vbMagenta = &hFF00FF
Const vbRed     = &h0000FF
Const vbWhite   = &hFFFFFF
Const vbYellow  = &h00FFFF
''' End Enum ' ColorConstants


' Const SystemFolder = 1
' Const TemporaryFolder = 2
' Const WindowsFolder = 0