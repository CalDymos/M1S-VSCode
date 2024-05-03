var m1sparser = function m1sparser_(options) {
  var
    tokens = [],
    tokenTypes = [],
    tokenTable = {
      "step": { "label": "Step", "type": "STEP"},
      "const": { "label": "Const", "type": "CONST"},
      "function": { "label": "Function", "type": "FUNCTION"},
      "type": { "label": "Type", "type": "TYPE"},
      "sub": { "label": "Sub", "type": "SUB"},
      "goto": { "label": "Goto", "type": "GOTO"},
      "xor": { "label": "Xor", "type": "BINARY_OPERATOR"},
      "or": { "label": "Or", "type": "BINARY_OPERATOR"},
      "and": { "label": "And", "type": "BINARY_OPERATOR"},
      "not": { "label": "Not", "type": "BINARY_OPERATOR"},
      "eqv": { "label": "Eqv", "type": "BINARY_OPERATOR"},
      "imp": { "label": "Imp", "type": "BINARY_OPERATOR"},
      "=": { "label": "=", "type": "COMPARISON_OPERATOR"},
      "<=": { "label": "<=", "type": "COMPARISON_OPERATOR"},
      ">=": { "label": ">=", "type": "COMPARISON_OPERATOR"},
      "<>": { "label": "<>", "type": "COMPARISON_OPERATOR"},
      "is": { "label": "Is", "type": "COMPARISON_OPERATOR"},
      "<": { "label": "<", "type": "COMPARISON_OPERATOR"},
      ">": { "label": ">", "type": "COMPARISON_OPERATOR"},
      "mod": { "label": "Mod", "type": "ARTHMETIC_OPERATOR"},
      "dim": { "label": "Dim", "type": "DIM"},
      "redim": { "label": "ReDim", "type": "REDIM"},
      "private": { "label": "Private", "type": "PRIVATE"},
      "global": { "label": "Global", "type": "GLOBAL"},
      "default": { "label": "Default", "type": "DEFAULT"},
      "next": { "label": "Next", "type": "FOR_LOOP_NEXT"},
      "nothing": { "label": "Nothing", "type": "OBJECT_NOTHING"},
      "null": { "label": "Null", "type": "VALUE_NULL"},
      "true": { "label": "True", "type": "VALUE_TRUE"},
      "false": { "label": "False", "type": "VALUE_FALSE"},
      "empty": { "label": "Empty", "type": "VALUE_EMPTY"},
      "byval": { "label": "ByVal", "type": "BYVAL"},
      "byref": { "label": "ByRef", "type": "BYREF"},
      "select": { "label": "Select", "type": "SELECT"},
      "case": { "label": "Case", "type": "CASE"},
      "if": { "label": "If", "type": "IF"},
      "else": { "label": "Else", "type": "ELSE"},
      "elseif": { "label": "ElseIf", "type": "ELSE_IF"},
      "exit": { "label": "Exit", "type": "EXIT"},
      "end": { "label": "End", "type": "END"},
      "then": { "label": "Then", "type": "THEN"},
      "err": { "label": "Err", "type": "ERR"},
      "call": { "label": "Call", "type": "CALL"},
      "erase": { "label": "Erase", "type": "ERASE"},
      "with": { "label": "With", "type": "WITH"},
      "stop": { "label": "Stop", "type": "STOP"},
      "on": { "label": "On", "type": "ON"},
      "error": { "label": "Error", "type": "ERROR"},
      "resume": { "label": "Resume", "type": "RESUME"},
      "option": { "label": "Option", "type": "OPTION"},
      "explicit": { "label": "Explicit", "type": "EXPLICIT"},
      "do": { "label": "Do", "type": "DO_LOOP"},
      "while": { "label": "While", "type": "WHILE_LOOP"},
      "wend": { "label": "Wend", "type": "WHILE_LOOP_WEND"},
      "until": { "label": "Until", "type": "DO_LOOP_UNTIL"},
      "loop": { "label": "Loop", "type": "DO_LOOP_END"},
      "for": { "label": "For", "type": "FOR_LOOP"},
      "to": { "label": "To", "type": "FOR_LOOP_TO"},
      "in": { "label": "In", "type": "FOR_LOOP_IN"},
      "set": { "label": "Set", "type": "SET_OPERATOR"},
      "new": { "label": "New", "type": "NEW_OPERATOR"},
      "abs": { "label": "Abs", "type": "CBSCRIPT_FUNCTION"},
      "array": { "label": "Array", "type": "CBSCRIPT_FUNCTION"},
      "asc": { "label": "Asc", "type": "CBSCRIPT_FUNCTION"},
      "atn": { "label": "Atn", "type": "CBSCRIPT_FUNCTION"},
      "cbool": { "label": "CBool", "type": "CBSCRIPT_FUNCTION"},
      "cbyte": { "label": "CByte", "type": "CBSCRIPT_FUNCTION"},
      "cdate": { "label": "CDate", "type": "CBSCRIPT_FUNCTION"},
      "cdbl": { "label": "CDbl", "type": "CBSCRIPT_FUNCTION"},
      "chr": { "label": "Chr", "type": "CBSCRIPT_FUNCTION"},
      "cint": { "label": "CInt", "type": "CBSCRIPT_FUNCTION"},
      "clng": { "label": "CLng", "type": "CBSCRIPT_FUNCTION"},
      "cos": { "label": "Cos", "type": "CBSCRIPT_FUNCTION"},
      "createobject": { "label": "CreateObject", "type": "CBSCRIPT_FUNCTION"},
      "csng": { "label": "CSng", "type": "CBSCRIPT_FUNCTION"},
      "cstr": { "label": "CStr", "type": "CBSCRIPT_FUNCTION"},
      "date": { "label": "Date", "type": "CBSCRIPT_FUNCTION"},
      "dateserial": { "label": "DateSerial", "type": "CBSCRIPT_FUNCTION"},
      "datevalue": { "label": "DateValue", "type": "CBSCRIPT_FUNCTION"},
      "day": { "label": "Day", "type": "CBSCRIPT_FUNCTION"},
      "derived math": { "label": "Derived Math", "type": "CBSCRIPT_FUNCTION"},
      "escape": { "label": "Escape", "type": "CBSCRIPT_FUNCTION"},
      "eval": { "label": "Eval", "type": "CBSCRIPT_FUNCTION"},
      "exp": { "label": "Exp", "type": "CBSCRIPT_FUNCTION"},
      "format": { "label": "Format", "type": "CBSCRIPT_FUNCTION"},
      "getobject": { "label": "GetObject", "type": "CBSCRIPT_FUNCTION"},
      "hex": { "label": "Hex", "type": "CBSCRIPT_FUNCTION"},
      "hour": { "label": "Hour", "type": "CBSCRIPT_FUNCTION"},
      "inputbox": { "label": "InputBox", "type": "CBSCRIPT_FUNCTION"},
      "instr": { "label": "InStr", "type": "CBSCRIPT_FUNCTION"},
      "int, fix": { "label": "Int, Fix", "type": "CBSCRIPT_FUNCTION"},
      "isarray": { "label": "IsArray", "type": "CBSCRIPT_FUNCTION"},
      "isdate": { "label": "IsDate", "type": "CBSCRIPT_FUNCTION"},
      "isempty": { "label": "IsEmpty", "type": "CBSCRIPT_FUNCTION"},
      "isnull": { "label": "IsNull", "type": "CBSCRIPT_FUNCTION"},
      "isnumeric": { "label": "IsNumeric", "type": "CBSCRIPT_FUNCTION"},
      "isobject": { "label": "IsObject", "type": "CBSCRIPT_FUNCTION"},
      "lbound": { "label": "LBound", "type": "CBSCRIPT_FUNCTION"},
      "lcase": { "label": "LCase", "type": "CBSCRIPT_FUNCTION"},
      "left": { "label": "Left", "type": "CBSCRIPT_FUNCTION"},
      "len": { "label": "Len", "type": "CBSCRIPT_FUNCTION"},
      "log": { "label": "Log", "type": "CBSCRIPT_FUNCTION"},
      "ltrim": { "label": "LTrim", "type": "CBSCRIPT_FUNCTION"},
      "maths": { "label": "Maths", "type": "CBSCRIPT_FUNCTION"},
      "mid": { "label": "Mid", "type": "CBSCRIPT_FUNCTION"},
      "minute": { "label": "Minute", "type": "CBSCRIPT_FUNCTION"},
      "month": { "label": "Month", "type": "CBSCRIPT_FUNCTION"},
      "msgbox": { "label": "MsgBox", "type": "CBSCRIPT_FUNCTION"},
      "now": { "label": "Now", "type": "CBSCRIPT_FUNCTION"},
      "oct": { "label": "Oct", "type": "CBSCRIPT_FUNCTION"},
      "replace": { "label": "Replace", "type": "CBSCRIPT_FUNCTION"},
      "right": { "label": "Right", "type": "CBSCRIPT_FUNCTION"},
      "rnd": { "label": "Rnd", "type": "CBSCRIPT_FUNCTION"},
      "round": { "label": "Round", "type": "CBSCRIPT_FUNCTION"},
      "rtrim": { "label": "RTrim", "type": "CBSCRIPT_FUNCTION"},
      "second": { "label": "Second", "type": "CBSCRIPT_FUNCTION"},
      "sgn": { "label": "Sgn", "type": "CBSCRIPT_FUNCTION"},
      "sin": { "label": "Sin", "type": "CBSCRIPT_FUNCTION"},
      "space": { "label": "Space", "type": "CBSCRIPT_FUNCTION"},
      "sqr": { "label": "Sqr", "type": "CBSCRIPT_FUNCTION"},
      "strcomp": { "label": "StrComp", "type": "CBSCRIPT_FUNCTION"},
      "string": { "label": "String", "type": "CBSCRIPT_FUNCTION"},
      "tan": { "label": "Tan", "type": "CBSCRIPT_FUNCTION"},
      "time": { "label": "Time", "type": "CBSCRIPT_FUNCTION"},
      "timer": { "label": "Timer", "type": "CBSCRIPT_FUNCTION"},
      "timeserial": { "label": "TimeSerial", "type": "CBSCRIPT_FUNCTION"},
      "timevalue": { "label": "TimeValue", "type": "CBSCRIPT_FUNCTION"},
      "trim": { "label": "Trim", "type": "CBSCRIPT_FUNCTION"},
      "typename": { "label": "TypeName", "type": "CBSCRIPT_FUNCTION"},
      "ubound": { "label": "UBound", "type": "CBSCRIPT_FUNCTION"},
      "ucase": { "label": "UCase", "type": "CBSCRIPT_FUNCTION"},
      "vartype": { "label": "VarType", "type": "CBSCRIPT_FUNCTION"},
      "weekday": { "label": "Weekday", "type": "CBSCRIPT_FUNCTION"},
      "year": { "label": "Year", "type": "CBSCRIPT_FUNCTION"},
    },
    source = options.source || '',
    lastParsedToken = '',
    lastNonWSParsedToken = '',
    pushToken = function m1sparser_pushToken(token, tokenType) {
      tokens.push(token);
      tokenTypes.push(tokenType);
      lastParsedToken = tokenType;
      if (tokenType === 'WHITESPACE' || tokenType === 'NEWLINE')
        lastNonWSParsedToken = tokenType;
    };


  //Lexical Analysis
  (function m1sparser_tokenizer() {
    var
      index = 0,
      bLength = source.length,
      buffer = options.source.split(''),
      n = 0,
      p = 0,
      isSpace = function m1sparser_tokenizer_isSpace(char) {
        return (char === ' ' || char === '\t' || char === '\f' || char ===
          '\v');
      },
      isEqualSign = function m1sparser_tokenizer_isEqualSign(char) {
        return (char === '=');
      },
      isEOLorEOF = function m1sparser_tokenizer_isEOLorEOF(char) {
        char = char || buffer[index];
        if (char === -1) return true;
        if (char === '\r' && nextChar() === '\n') {
          return true;
        }
        if (char === '\n') return true;
        return false;
      },
      isAlphaNumeric = function m1sparser_tokenizer_isAlphaNumeric(char) {
        return char !== -1 && /[a-zA-Z0-9_]/.test(char);
      },
      isDigit = function m1sparser_tokenizer_isDigit(char) {
        return /[0-9]/.test(char);
      },
      read = function m1sparser_tokenizer_read(length) {
        var str = '';
        length = length || 1;
        if (index + length > bLength) {
          return -1;
        }
        str = buffer.slice(index, index + length).join('');
        index += length;
        return str;
      },
      currentChar = function m1sparser_tokenizer_currentChar() {
        return (index >= 0 && index < bLength) ? buffer[index] : -1;
      },
      charAt = function m1sparser_tokenizer_charAt(chrIndex) {
        if (chrIndex >= 0 && chrIndex < bLength) {
          return buffer[chrIndex]
        }
        return -1;
      },
      prevChar = function m1sparser_tokenizer_prevChar() {
        if (index - 1 >= 0 && bLength > 0) {
          return buffer[index - 1]
        }
        return -1;
      },
      nextChar = function m1sparser_tokenizer_nextChar() {
        if (index + 1 >= bLength) {
          return -1;
        }
        return buffer[index + 1]
      },
      readTill = function m1sparser_tokenizer_readTill(fn) {
        var n = 0,
          str = '',
          peeked;

        while ((peeked = charAt(index + n)) !== -1 && !(isEOLorEOF(peeked)) &&
          fn(peeked, buffer, index)) {
          n++;
        }
        if (n === 0) return '';
        str = read(n);
        return str;
      },
      readSpace = function m1sparser_tokenizer_readSpace() {
        return readTill(function(char) {
          return isSpace(char);
        });
      },
      readNextWord = function m1sparser_tokenizer_readNextWord() {
        var ch,
          str = '';
        //n = 0;
        do {
          ch = charAt(index + n++);
        } while (isSpace(ch) || ch === '_' || ch === '\r' || ch === '\n');
        n--;
        if (n === 0) {
          return '';
        }
        while (isAlphaNumeric(ch = charAt(index + n++))) {
          str += ch;
        }
        n--;
        return str;
      },
      readPrevWord = function m1sparser_tokenizer_readPrevWord() {
        var ch,
          str = '';
        do {
          ch = charAt(index + p--);
        } while (index + p != 0 && (isSpace(ch) || ch === '_' || ch === '\r' || ch === '\n'));
        if (index + p === 0) {
          return '';
        }
        do {
          ch = charAt(index + p--);
        } while (index + p != 0 && (!isSpace(ch) && ch != '_' && ch != '\r' && ch != '\n'));
        if (index + p === 0) {
          return '';
        }
        do {
          ch = charAt(index + p--);
        } while (index + p != 0 && (isSpace(ch) || ch === '_' || ch === '\r' || ch === '\n'));
        if (index + p === 0) {
          return '';
        }
        do {
          ch = charAt(index + p--);
        } while (index + p != 0 && (!isSpace(ch) && ch != '_' && ch != '\r' && ch != '\n'));
        p++;
        if (index + p === 0) {
          return '';
        }
        while (isAlphaNumeric(ch = charAt(index + p++))) {
          str += ch;
        }
        p--;
        return str;
      },
      readString = function m1sparser_tokenizer_readString() {
        var n = 1,
          str = '',
          peeked;

        while ((peeked = charAt(index + n)) !== -1 && !(isEOLorEOF(peeked))) {
          if(peeked !== '\"'){
            str += peeked;
          } else if (charAt(index + n + 1) === '\"') {
            n++;
          } else {
            n++;
            break;
          }
          n++;
        }
        if (n === 0) return '';
        str = read(n);
        return str;
      },
      readLine = function m1sparser_tokenizer_readLine() {
        return readTill(function(char) {
          return !isEOLorEOF(char);
        });
      },
      readAlphaNumeric = function m1sparser_tokenizer_readAlphaNumeric() {
        return readTill(function(char) {
          return isAlphaNumeric(char);
        });
      },
      readNumber = function m1sparser_tokenizer_readNumber() {
        var str = '';

        str += readTill(function(chr) {
          return isDigit(chr);
        });

        if (currentChar() === '.') str += read();

        if (str.length === 0) return '';
        str += readTill(function(chr) {
          return isDigit(chr);
        });

        if (currentChar() === 'e' || currentChar() === 'E') {
          str += read();
          if (currentChar() === '+' || currentChar() === '-') str += read();
          str += readTill(function(chr) {
            return isDigit(chr);
          });
        }
        return str;
      };

    var
      curChar,
      nextChr,
      prevChr,
      prevChr2,
      ch,
      word,
      nextWord,
      prevWord;

    while ((ch = currentChar()) !== -1) {
      word = '';
      switch (ch) {
        /*case '~':
        case ';':
        case '?':
        case '|':
        case '`':
        case '!':
        case '{':
        case '}':
            pushToken(ch, UNKNOWN);
            break;*/
        case '\t':
        case '\v':
        case ' ':
        case '\f':
          pushToken(readSpace(), 'WHITESPACE');
          break;
        case '(':
          pushToken(read(), 'OPEN_BRACKET');
          break;
        case ')':
          pushToken(read(), 'CLOSE_BRACKET');
          break;
        case '\"':
          pushToken(readString(), 'STRING');
          break;
        case '\'':

          pushToken(readLine(), 'COMMENT');
          break;
        case '#':
          word = readLine();
          if (word.indexOf("#expand ") != -1){
            pushToken(word, 'EXPAND');
            continue;
          }else { 
            pushToken(word, 'UNKNOWN');
          }
          break; 
        case '[':
          word = readTill(function(char) {
            return char !== ']';
          }) + "]";
          read();
          pushToken(word, 'IDENTIFIER');
          break;
        case '_':
          pushToken(read(), 'STATEMENT_CONTINUATION');
          break;
        case ':':
          pushToken(read(), 'STATEMENT_SEPARATOR');
          break;
        case '@':
          if (nextChar() === '@') pushToken(readLine(), 'COMMENT');
          else pushToken(read(), 'UNKNOWN');
          break;
        case '+':
        case '^':
        case '*':
        case '/':
        case '%':
        case '\\':
        case '=':
          pushToken(read(), 'ARTHMETIC_OPERATOR');
          break;
        case '-':
          prevChr = prevChar();
          if ((charAt(index - 2) === ',') && (prevChr === ' ')) {
            pushToken(read(), 'INVALID');
            continue;
          }
          pushToken(read(), 'ARTHMETIC_OPERATOR');
          break;          
        case '<':
          nextChr = nextChar();

          if (nextChr === '>') {
            pushToken(read(2), 'COMPARISON_OPERATOR');
          } else if (nextChr === '=') {
            pushToken(read(2), 'COMPARISON_OPERATOR');
          } else {
            pushToken(read(), 'COMPARISON_OPERATOR');
          }

          break;
        case '>':
          nextChr = nextChar();

          if (nextChr === '=') {
            pushToken(read(2), 'COMPARISON_OPERATOR');
          } else {
            pushToken(read(), 'COMPARISON_OPERATOR');
          }

          break;
        case '&':
          nextChr = nextChar();

          if (nextChr === 'H' || nextChr === 'h') {
            //this is a hexa decimal value
            pushToken(read() + readAlphaNumeric(), 'HEXNUMBER');
          } else {
            pushToken(read(), 'ARTHMETIC_OPERATOR');
          }

          break;
        case '.':
          nextChr = nextChar();

          if (isDigit(nextChr)) {
            pushToken(readNumber(), 'NUMBER');
          } else {
            //record dot operator
            pushToken(read(), 'DOT_OPERATOR');
          }
          break;
        case '\r':
          if (nextChar() === '\n') {
            pushToken(read(2), 'NEWLINE');
          } else {
            pushToken(read(), 'NEWLINE');
          }

          break;
        case '\n':
          pushToken(read(), 'NEWLINE');
          break;
        case '0':
        case '1':
        case '2':
        case '3':
        case '4':
        case '5':
        case '6':
        case '7':
        case '8':
        case '9':
          pushToken(readNumber(), 'NUMBER');
          break;
        case ',':
          pushToken(read(), 'COMMA');
          break;
        default:

          if (!isAlphaNumeric(ch)) {
            pushToken(read(), 'INVALID');
            continue;
          }

          word = readAlphaNumeric();
          n = 0;

          if (lastNonWSParsedToken !== 'DOT_OPERATOR' && tokenTable[word.toLowerCase()] !== undefined) {

            switch (word.toLowerCase()) {
              case 'do':
                nextWord = readNextWord().toLowerCase();

                if (nextWord === 'while') {
                  read(n);
                  pushToken('Do While', 'DO_LOOP_START_WHILE');
                } else if (nextWord === 'until') {
                  read(n);
                  pushToken('Do Until', 'DO_LOOP_START_UNTIL');
                } else {
                  pushToken(tokenTable[word.toLowerCase()].label, tokenTable[word.toLowerCase()].type);
                }

                break;
              case 'loop':
                nextWord = readNextWord().toLowerCase();

                if (nextWord === 'while') {
                  read(n);
                  pushToken('Loop While', 'DO_LOOP_END_WHILE');
                } else if (nextWord === 'until') {
                  read(n);
                  pushToken('Loop Until', 'DO_LOOP_END_UNTIL');
                }else {
                  pushToken('Loop', 'DO_LOOP_END');
                }

                break;
              case 'for':
                nextWord = readNextWord();

                if (nextWord.toLowerCase() === 'each') {
                  read(n);
                  pushToken('For Each', 'FOR_EACHLOOP');
                } else if (nextWord.toLowerCase() === 'input' || nextWord.toLowerCase() === 'output' || nextWord.toLowerCase() === 'binary' || nextWord.toLowerCase() === 'append'){
                  pushToken(tokenTable[word.toLowerCase()].label, 'FOR_MODE'); 
                } else {
                  
                  pushToken(tokenTable[word.toLowerCase()].label, tokenTable[word.toLowerCase()].type);
                }

                break;
              case 'on':
                nextWord = readNextWord();

                if (nextWord.toLowerCase() === 'error') {
                  nextWord1 = readNextWord().toLowerCase();
                  nextWord2 = readNextWord().toLowerCase();

                  if (nextWord1 === 'resume' && nextWord2 === 'next') {
                    read(n);
                    pushToken('On Error Resume Next',
                      'ON_ERROR_RESUME_NEXT');
                  } else if (nextWord1 === 'goto' && nextWord2 === '0') {
                    read(n);
                    pushToken('On Error GoTo 0', 'ON_ERROR_GOTO_0');
                  } else {
                    pushToken(tokenTable[word.toLowerCase()].label, tokenTable[word.toLowerCase()].type);
                  }
                }

                break;
              case 'case':
                nextWord = readNextWord();

                if (nextWord.toLowerCase() === 'else') {
                  read(n);
                  pushToken('Case Else', 'CASE_ELSE');
                } else {
                  pushToken(tokenTable[word.toLowerCase()].label, tokenTable[word.toLowerCase()].type);
                }

                break;
              case 'select':
                nextWord = readNextWord();

                if (nextWord.toLowerCase() === 'case') {
                  read(n);
                  pushToken('Select Case', 'SELECT_CASE');
                } else {
                  pushToken(tokenTable[word.toLowerCase()].label, tokenTable[word.toLowerCase()].type);
                }

                break;
              case 'end':
                nextWord = readNextWord();

                switch (nextWord.toLowerCase()) {
                  case 'function':
                    read(n);
                    pushToken('End Function', 'END_FUNCTION');

                    break;
                  case 'sub':
                    read(n);
                    pushToken('End Sub', 'END_SUB');

                    break;
                  case 'type':
                    read(n);
                    pushToken('End Type', 'END_TYPE');

                    break;
                  case 'if':
                    read(n);
                    pushToken('End If', 'END_IF');

                    break;
                  case 'with':
                    read(n);
                    pushToken('End With', 'END_WITH');

                    break;
                  case 'select':
                    read(n);
                    pushToken('End Select', 'END_SELECT');

                    break;
                  default:
                    pushToken(tokenTable[word.toLowerCase()].label, tokenTable[word.toLowerCase()].type);

                    break;
                }
                break;
              case 'exit':
                nextWord = readNextWord();

                switch (nextWord.toLowerCase()) {
                  case 'function':
                    read(n);
                    pushToken('Exit Function', 'EXIT_FUNCTION');

                    break;
                  case 'for':
                    read(n);
                    pushToken('Exit For', 'EXIT_FOR');

                    break;
                  case 'do':
                    read(n);
                    pushToken('Exit Do', 'EXIT_DO');

                    break;
                  case 'sub':
                    read(n);
                    pushToken('Exit Sub', 'EXIT_SUB');

                    break;
                  default:
                    pushToken(tokenTable[word.toLowerCase()].label, tokenTable[word.toLowerCase()].type);
                    break;
                }

                break;
              default:
                pushToken(tokenTable[word.toLowerCase()].label, tokenTable[word.toLowerCase()].type);
                break;
            }
          } else {

              switch (word.toLowerCase()){

                default:
                  curChar = currentChar()
                  nextChr = nextChar();
    
                  switch (curChar) {
                  
                    case ':':
                      if (nextChr === '='){
                        read(2);
                        nextWord = readTill(function(char) {
                          return isAlphaNumeric(char);
                        });
                        pushToken(word + ':=' + nextWord, "ARGUMENT_ASSIGNMENT")
                        break;
                      }
                      if ((nextChr === '\n') || (nextChr === '\r')){
                        read();
                        pushToken(word + ':', "LABEL");
                        break;
                      }
                    default:
                      pushToken(word, 'UNKNOWN');
                      break;
                
                    }
                break;
                
              }
            /*switch (word.toUpperCase()) {
                            case 'REM':
                                pushToken(word + readLine(), 'COMMENT');
                                break;
                            default:*/
            //pushToken(word, 'UNKNOWN');
            /*break;
                        }*/
          }

          break;
      }
    }

    pushToken(null, 'EOF');
  })();
  //Syntatic Analysis
  (function m1sparser_analizer() {
    var
      i = 0,
      lTokens = tokens.length,
      n = 0,
      curToken = null,
      curTokenType = null,
      nextToken = null,
      skipToToken = function m1sparser_analizer_skipToToken(tokenType) {

        while (n < lTokens && tokenTypes[++n] !== tokenType);
      },
      skipToEndOfStatement = function m1sparser_analizer_skipToEndOfStatement(fn) {
        var
          bIgnoreNewLine = false,
          tokenType = null;

        while (n < lTokens) {
          tokenType = tokenTypes[++n];
          switch (tokenType) {
            case 'STATEMENT_SEPARATOR':
              if (options.breakOnSeperator) return;
              break;
            case 'NEWLINE':
              if (bIgnoreNewLine) {
                bIgnoreNewLine = false;
                break;
              }
              return;
            case 'STATEMENT_CONTINUATION':
              bIgnoreNewLine = true;
              break;
            default:
              break;
          }
          if (fn) {
            fn(tokenType);
          }
        }
      };

    for (i = 0; i < lTokens - 1; i++) {

      curToken = tokens[i];
      curTokenType = tokenTypes[i];

      n = i + 1;

      while (n < lTokens && tokenTypes[n] === 'WHITESPACE') {
        n++;
      }

      if (n < lTokens) {
        nextToken = tokens[n];
      }

      n = i;

      switch (curTokenType) {
        case 'IF':
          skipToToken('THEN')
          while (tokenTypes[++n] !== 'NEWLINE') {
            if (!(tokenTypes[n] === 'WHITESPACE' || tokenTypes[n] ===
                'COMMENT')) {
              tokenTypes[i] = 'IF_ELSE_ONE_LINE';
              break;
            }
          }
          break;
        case 'DIM':
          skipToEndOfStatement(function(t) {
            if (t === 'UNKNOWN') {
              tokenTypes[n] = 'VARIABLE_NAME';
            }
          });

          break;
        case 'REDIM':
          skipToEndOfStatement(function(t) {
            if (t === 'UNKNOWN') {
              tokenTypes[n] = 'VARIABLE_NAME';
            }
          });

          break;
        case 'CONST':
          skipToEndOfStatement(function(t) {
            if (t === 'UNKNOWN') {
              tokenTypes[n] = 'CONST_VARIABLE_NAME';
            }
          });

          break;
        default:
          break;
      }

      lastTokenType = tokenTypes[i];

      if (lastTokenType !== 'WHITESPACE' && lastTokenType !== 'NEWLINE') {
        lastNonWSParsedToken = lastTokenType;
      }

      if (lastNonWSParsedToken != 'WHITESPACE') {
        lastNonWSLNParsedToken = lastTokenType;
      }
    }
  })();

  return {
    tokens: tokens,
    tokenTypes: tokenTypes
  };
};
module.exports = m1sparser;
