'
'  Adapted from
'     https://support.office.com/de-de/article/konvertieren-von-zahlen-in-w%C3%B6rter-a0d166fb-e1ea-4090-95c8-69442cd55d98
'
option explicit

' { spellNumber
public function spellNumber( _
           byVal    num       as string            , _
           optional lang      as string = "de"     , _
           optional currency_ as string = "Franken", _
           optional fraction  as string = "Rappen"   _
       )

   dim currencyStr   as string
   dim fractionalStr as string
'  dim Dollars, fractionalStr, Temp

'  redim unitStr(9) as string
   dim unitStr(5) as string

   if lang = "en" then
      unitStr(2) = " Thousand"
      unitStr(3) = " Million"
      unitStr(4) = " Billion"
      unitStr(5) = " Trillion"
   elseif lang = "de" then
      unitStr(2) = " Tausend"
      unitStr(3) = " Million"
      unitStr(4) = " Milliarde"
      unitStr(5) = " Billion"
   end if

    ' string representation of amount.
    
    num = trim(str(num))
    
  ' Position of decimal place 0 if none.
    
    dim decimalPlace
    decimalPlace = instr(num, ".")
    
  ' Convert cents and set num to dollar amount.
    
    if decimalPlace > 0 then
    
       fractionalStr = num_10_to_99_to_string(Left(mid(num, decimalPlace + 1) & "00", 2), lang)
    
       num = trim(Left(num, decimalPlace - 1))
    
    end if
    
    dim unitNo
    unitNo = 1
    
    do while num <> "" ' {
    
       dim hundredsStr
       hundredsStr = GetHundreds(Right(num, 3), lang)
       
       if hundredsStr <> "" then currencyStr = hundredsStr & unitStr(unitNo) & currencyStr
       
       if len(num) > 3 then
          num = Left(num, len(num) - 3)
       else
          num = ""
       end if
       
       unitNo = unitNo + 1
    
    loop ' }
        
    select case currencyStr ' {
       case ""   : currencyStr = "No Dollars"
       case "One": currencyStr = "One Dollar"
       case else : currencyStr = currencyStr & " " & currency_
    end select ' }
    
    select case fractionalStr ' {
    case "" ' {
    
       if lang = "en" then
          fractionalStr = " and No " & fraction
       elseif lang = "de" then
          fractionalStr = " und kein " & fraction
       end if
    ' } 
    case "One" ' {
       if lang = "en" then
          fractionalStr = " and One Cent"
       elseIf lang = "de" then
          fractionalStr = " und ein " & fraction
       end if
    ' }
    case else ' {
    
       if lang = "en" then
          fractionalStr = " and " & fractionalStr & " " & fraction
       elseIf lang = "de" then
          fractionalStr = " und " & fractionalStr & " " & fraction
       end if
    ' } 
    end select ' }
    
    spellNumber = currencyStr & fractionalStr

end function ' }

public sub test_SpellNumber() ' {

    call checkSpelledNumber("1", "de", "Franken", "Rappen", "Ein Franken und kein Rappen")
    call checkSpelledNumber("1", "en", "Dollar" , "Cent"  , "One Dollar and No Cent")

end sub ' }

private sub checkSpelledNumber(number as string, lang as string, currency_ as string, fraction as string, expected as string) ' {

   dim gotten as string
   gotten = spellNumber(number, lang, currency_, fraction)

   if gotten <> expected then ' {
      debug.print(number & ", " & lang & ", " & currency_ & ", " & fraction & ": " & gotten & " != " & expected)
   end if ' }

end sub ' }

' Converts a number from 100-999 into text

function GetHundreds(byVal MyNumber as string, lang as string) as string ' {

    dim Result as string
    
    if val(MyNumber) = 0 then exit function
    
    MyNumber = Right("000" & MyNumber, 3)
    
    ' Convert the hundreds place.
    
    if mid(MyNumber, 1, 1) <> "0" then
    
       if lang = "en" then
          Result = num_1_to_9_to_string(mid(MyNumber, 1, 1), lang) & " Hundred "
       elseif lang = "de" then
          Result = num_1_to_9_to_string(mid(MyNumber, 1, 1), lang) & " Hundert "
       end if
    
    end if
    
    ' Convert the tens and ones place.
    
    if mid(MyNumber, 2, 1) <> "0" then
    
       Result = Result & num_10_to_99_to_string(mid(MyNumber, 2), lang)
    
    else
    
       Result = Result & num_1_to_9_to_string(mid(MyNumber, 3), lang)
    
    end if
    
    GetHundreds = Result

end function ' }


' Converts a number from 10 to 99 into text.


private function num_10_to_99_to_string(TensText as string, lang as string) as string ' {

    dim Result as string
    
    Result = "" ' Null out the temporary function value.
    
    if val(Left(TensText, 1)) = 1 then ' if value between 10-19...
     
       if lang = "en" then
       
        select case val(TensText)
        
            case 10: Result = "Ten"
            case 11: Result = "Eleven"
            case 12: Result = "Twelve"
            case 13: Result = "Thirteen"
            case 14: Result = "Fourteen"
            case 15: Result = "Fifteen"
            case 16: Result = "Sixteen"
            case 17: Result = "Seventeen"
            case 18: Result = "Eighteen"
            case 19: Result = "Nineteen"
            case else
        
        end select
        
       elseif lang = "de" then
        
        
          select case val(TensText)
        
            case 10: Result = "Zehn"
            case 11: Result = "Elf"
            case 12: Result = "Zwölf"
            case 13: Result = "Dreizehn"
            case 14: Result = "Vierzehn"
            case 15: Result = "Fünfzehn"
            case 16: Result = "Sechszehn"
            case 17: Result = "Siebzehn"
            case 18: Result = "Achtzehn"
            case 19: Result = "Neunzehn"
            case else
        
        end select
         
        end if
    
    else ' if value between 20-99...
    
       if lang = "en" then
        select case val(Left(TensText, 1))
        

            case 2: Result = "Twenty"
            case 3: Result = "Thirty"
            case 4: Result = "Forty"
            case 5: Result = "Fifty"
            case 6: Result = "Sixty"
            case 7: Result = "Seventy"
            case 8: Result = "Eighty"
            case 9: Result = "Ninety"
            case else
         end select
        
       elseif lang = "de" then
          select case val(Left(TensText, 1))
            case 2: Result = "Zwanzig"
            case 3: Result = "Dreissig"
            case 4: Result = "Vierzig"
            case 5: Result = "Fünfzig"
            case 6: Result = "Sechzig"
            case 7: Result = "Siebzig"
            case 8: Result = "Achtzig"
            case 9: Result = "Neunzig"
            case else
          end select
          end if
        
    Result = Result & num_1_to_9_to_string(Right(TensText, 1), lang)  ' Retrieve ones place.
    
    end if
    
    num_10_to_99_to_string = Result

end function ' }

private function num_1_to_9_to_string(Digit as string, lang as string) as string ' {

   if lang = "en" then

        select case val(Digit)
           case 1: num_1_to_9_to_string = "One"
           case 2: num_1_to_9_to_string = "Two"
           case 3: num_1_to_9_to_string = "Three"
           case 4: num_1_to_9_to_string = "Four"
           case 5: num_1_to_9_to_string = "Five"
           case 6: num_1_to_9_to_string = "Six"
           case 7: num_1_to_9_to_string = "Seven"
           case 8: num_1_to_9_to_string = "Eight"
           case 9: num_1_to_9_to_string = "Nine"
           case else: num_1_to_9_to_string = ""
        end select

   elseIf lang = "de" then

        select case val(Digit)
           case 1: num_1_to_9_to_string = "Ein"
           case 2: num_1_to_9_to_string = "Zwei"
           case 3: num_1_to_9_to_string = "Drei"
           case 4: num_1_to_9_to_string = "Vier"
           case 5: num_1_to_9_to_string = "Fünf"
           case 6: num_1_to_9_to_string = "Sechs"
           case 7: num_1_to_9_to_string = "Sieben"
           case 8: num_1_to_9_to_string = "Acht"
           case 9: num_1_to_9_to_string = "Neun"
           case else: num_1_to_9_to_string = ""
        end select

   end if

end function ' }
