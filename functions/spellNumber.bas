'
'  https://github.com/ReneNyffenegger/temp-VBA/blob/master/functions/spellNumber.bas
'
'  Adapted from
'     https://support.office.com/de-de/article/konvertieren-von-zahlen-in-w%C3%B6rter-a0d166fb-e1ea-4090-95c8-69442cd55d98
'
option explicit


' { spellNumber
public function spellNumber (                     _
           byVal    num    as longLong          , _
           optional lang   as string = "de"     , _
           optional de_1   as string = "eins"   , _
           optional de_ss  as string = "ss"       _
           ) as string

    dim numText as string
    numText = num

   '  redim unitStr(9) as string
    dim unitStr(5, 2) as string

    if lang = "en" then ' {
       unitStr(2, 1) = " Thousand"
       unitStr(3, 1) = " Million"
       unitStr(4, 1) = " Billion"
       unitStr(5, 1) = " Trillion"

       unitStr(2, 2) = " Thousand"
       unitStr(3, 2) = " Million"
       unitStr(4, 2) = " Billion"
       unitStr(5, 2) = " Trillion"
    ' }
    elseif lang = "de" then ' {
       unitStr(2, 1 ) = "tausend"
       unitStr(3, 1) = " Million"
       unitStr(4, 1) = " Milliarde"
       unitStr(5, 1) = " Billion"

       unitStr(2, 2) = "tausend"
       unitStr(3, 2) = " Millionen"
       unitStr(4, 2) = " Milliarden"
       unitStr(5, 2) = " Billionen"
    end if ' }

    dim unitNo
    unitNo = 1

    do while numText <> "" ' {

       dim hundredsStr

       dim de_ein as string
       de_ein = "ein"
       if unitNo >= 3 then
          de_ein = "eine"
       end if

       hundredsStr = num_100_to_999( right(numText, 3), lang, de_ein := de_ein)

       dim pl_sg as integer
       pl_sg = 1

       if cLng(right(numText, 3)) > 1 then pl_sg = 2

'      if lang = "de" and unitNo >= 3 and cLng(right(numText, 3)) > 1 then


       if lang = "de" and unitNo >= 3 and spellNumber <> "" then
          spellNumber = " " & spellNumber
       end if

'      hundredsStr = num_100_to_999( right(numText, 3), lang)

       if hundredsStr <> "" then spellNumber = hundredsStr & unitStr(unitNo, pl_sg) & spellNumber

       if len(numText) > 3 then
          numText = left(numText, len(numText) - 3)
       else
          numText = ""
       end if

       unitNo = unitNo + 1

    loop ' }

    if (num = 1 or right(num, 2) = "01") and de_1 = "eins" then ' {
        spellNumber = spellNumber & "s"
    end if ' }

end function ' }

' { spellNumberWithCurrency
public function spellNumberWithCurrency( _
           byVal    num       as string            , _
           optional lang      as string = "de"     , _
           optional currency_ as string = "Franken", _
           optional fraction  as string = "Rappen" , _
           optional de_ss     as string = "ss"       _
       ) as string

   dim currencyStr   as string
   dim fractionalStr as string
'  dim Dollars, fractionalStr, Temp

' '  redim unitStr(9) as string
'  dim unitStr(5) as string

'  if lang = "en" then ' {
'     unitStr(2) = " Thousand"
'     unitStr(3) = " Million"
'     unitStr(4) = " Billion"
'     unitStr(5) = " Trillion"
'  ' }
'  elseif lang = "de" then ' {
'     unitStr(2) = " Tausend"
'     unitStr(3) = " Million"
'     unitStr(4) = " Milliarde"
'     unitStr(5) = " Billion"
'  end if ' }

  ' string representation of amount.

    num = trim(str(num))

  ' Position of decimal place 0 if none.

    dim decimalPlace
    decimalPlace = instr(num, ".")

  ' Convert cents and set num to dollar amount.

    if decimalPlace > 0 then ' {

       fractionalStr = num_10_to_99_to_string(Left(mid(num, decimalPlace + 1) & "00", 2), lang, de_ss := de_ss)

       num = trim(Left(num, decimalPlace - 1))

    end if ' }


    currencyStr = spellNumber(num, lang)

'   dim unitNo
'   unitNo = 1
'   do while num <> "" ' {
'
'      dim hundredsStr
'      hundredsStr = num_100_to_999(Right(num, 3), lang)
'
'      if hundredsStr <> "" then currencyStr = hundredsStr & unitStr(unitNo) & currencyStr
'
'      if len(num) > 3 then
'         num = Left(num, len(num) - 3)
'      else
'         num = ""
'      end if
'
'      unitNo = unitNo + 1
'
'   loop ' }

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

    spellNumberWithCurrency = currencyStr & fractionalStr

end function ' }

public sub test_SpellNumber() ' {

    call checkSpelledNumber(            1, "de", "eins"                                                                                                )
    call checkSpelledNumber(            2, "de", "zwei"                                                                                                )
    call checkSpelledNumber(           10, "de", "zehn"                                                                                                )
    call checkSpelledNumber(           11, "de", "elf"                                                                                                 )
    call checkSpelledNumber(           12, "de", "zwölf"                                                                                               )
    call checkSpelledNumber(           13, "de", "dreizehn"                                                                                            )
    call checkSpelledNumber(           20, "de", "zwanzig"                                                                                             )
    call checkSpelledNumber(           22, "de", "zweiundzwanzig"                                                                                      )
    call checkSpelledNumber(           30, "de", "dreissig"                                                                                            )
    call checkSpelledNumber(           35, "de", "fünfunddreissig"                                                                                     )
    call checkSpelledNumber(          100, "de", "einhundert"                                                                                          ) ' Oder hundert
    call checkSpelledNumber(          101, "de", "einhunderteins"                                                                                      ) ' Oder hunderteins
    call checkSpelledNumber(          150, "de", "einhundertfünfzig"                                                                                   ) ' Oder hundertfünfzig
    call checkSpelledNumber(          180, "de", "einhundertachtzig"                                                                                   ) ' Oder hundertachtzig
    call checkSpelledNumber(          186, "de", "einhundertsechsundachtzig"                                                                           )
    call checkSpelledNumber(          999, "de", "neunhundertneunundneunzig"                                                                           )
    call checkSpelledNumber(         1000, "de", "eintausend"                                                                                          ) ' Oder tausend
    call checkSpelledNumber(         1001, "de", "eintausendeins"                                                                                      ) ' Oder tausend
    call checkSpelledNumber(       100001, "de", "einhunderttausendeins"                                                                               ) ' Oder hunderttaussendeins, einhunderttausendundeins, hunderttausendundeins
    call checkSpelledNumber(      1000000, "de", "eine Million"                                                                                        )
    call checkSpelledNumber(      1500000, "de", "eine Million fünfhunderttausend"                                                                     )
    call checkSpelledNumber(      4950300, "de", "vier Millionen neunhundertfünfzigtausenddreihundert"                                                 )
    call checkSpelledNumber(   1234678901, "de", "eine Milliarde zweihundertvierunddreissig Millionen sechshundertachtundsiebzigtausendneunhunderteins")
    call checkSpelledNumber(1000000000000, "de", "eine Billion"                                                                                        )
    call checkSpelledNumber(2000000000000, "de", "zwei Billionen"                                                                                      )

 '
 '  TODO: Include me again:
 '
 '  call checkSpelledNumberWithCurrency("1", "de", "Franken", "Rappen", "Ein Franken und kein Rappen")
 '  call checkSpelledNumberWithCurrency("1", "en", "Dollar" , "Cent"  , "One Dollar and No Cent")

    debug.print "test finished."
end sub ' }

private sub checkSpelledNumberWithCurrency(number as string, lang as string, currency_ as string, fraction as string, expected as string) ' {

   dim gotten as string
   gotten = spellNumberWithCurrency(number, lang, currency_, fraction)

   if gotten <> expected then ' {
      debug.print(number & ", " & lang & ", " & currency_ & ", " & fraction & ": " & gotten & " != " & expected)
   end if ' }

end sub ' }

private sub checkSpelledNumber(num as longLong, lang as string, expected as string) ' {

   dim gotten as string
   gotten = spellNumber(num, lang)

   if gotten <> expected then ' {
      debug.print(num & ", " & lang & ": " & gotten & " != " & expected)
   end if ' }

end sub ' }

function num_100_to_999(byVal MyNumber as string, lang as string, optional de_ein as string = "ein", optional de_ss as string = "ss") as string ' {
  '
  ' Spell a number between 100 and 999
  '

'   dim Result as string

    if val(MyNumber) = 0 then exit function

    MyNumber = Right("000" & MyNumber, 3)

    ' Convert the hundreds place.

    if mid(MyNumber, 1, 1) <> "0" then

       if lang = "en" then
          num_100_to_999 = num_1_to_9_to_string(mid(MyNumber, 1, 1), lang) & " Hundred "
       elseif lang = "de" then
          num_100_to_999 = num_1_to_9_to_string(mid(MyNumber, 1, 1), lang) & "hundert"
       end if

    end if

    ' Convert the tens and ones place.

    if mid(MyNumber, 2, 1) <> "0" then

       num_100_to_999 = num_100_to_999 & num_10_to_99_to_string(mid(MyNumber, 2), lang, de_ss := de_ss)

    else

       num_100_to_999 = num_100_to_999 & num_1_to_9_to_string(mid(MyNumber, 3), lang, de_ein := de_ein)

    end if

'   num_100_to_999 = Result

end function ' }

private function num_10_to_99_to_string(TensText as string, lang as string, optional de_ss as string = "ss") as string ' {

  '
  ' Spell a number between 10 and 99
  '

'   dim Result as string

    num_10_to_99_to_string = "" ' Null out the temporary function value.

    if val(Left(TensText, 1)) = 1 then ' { if value between 10-19...

       if lang = "en" then  ' {

          select case val(TensText) ' {

              case 10: num_10_to_99_to_string = "Ten"
              case 11: num_10_to_99_to_string = "Eleven"
              case 12: num_10_to_99_to_string = "Twelve"
              case 13: num_10_to_99_to_string = "Thirteen"
              case 14: num_10_to_99_to_string = "Fourteen"
              case 15: num_10_to_99_to_string = "Fifteen"
              case 16: num_10_to_99_to_string = "Sixteen"
              case 17: num_10_to_99_to_string = "Seventeen"
              case 18: num_10_to_99_to_string = "Eighteen"
              case 19: num_10_to_99_to_string = "Nineteen"
              case else

          end select ' }
       ' }
       elseif lang = "de" then ' {

          select case val(TensText) ' {

            case 10: num_10_to_99_to_string = "zehn"
            case 11: num_10_to_99_to_string = "elf"
            case 12: num_10_to_99_to_string = "zwölf"
            case 13: num_10_to_99_to_string = "dreizehn"
            case 14: num_10_to_99_to_string = "vierzehn"
            case 15: num_10_to_99_to_string = "fünfzehn"
            case 16: num_10_to_99_to_string = "sechszehn"
            case 17: num_10_to_99_to_string = "siebzehn"
            case 18: num_10_to_99_to_string = "achtzehn"
            case 19: num_10_to_99_to_string = "neunzehn"
            case else

        end select ' }

       end if ' }
    ' }
    else ' { if value between 20-99...

       if lang = "en" then ' {

          select case val(Left(TensText, 1)) ' {
             case 2: num_10_to_99_to_string = "Twenty"
             case 3: num_10_to_99_to_string = "Thirty"
             case 4: num_10_to_99_to_string = "Forty"
             case 5: num_10_to_99_to_string = "Fifty"
             case 6: num_10_to_99_to_string = "Sixty"
             case 7: num_10_to_99_to_string = "Seventy"
             case 8: num_10_to_99_to_string = "Eighty"
             case 9: num_10_to_99_to_string = "Ninety"
             case else
          end select ' }

       ' }
       elseif lang = "de" then ' {
          select case val(Left(TensText, 1)) ' {
            case 2: num_10_to_99_to_string = "zwanzig"
            case 3: num_10_to_99_to_string = "drei" & de_ss & "ig"
            case 4: num_10_to_99_to_string = "vierzig"
            case 5: num_10_to_99_to_string = "fünfzig"
            case 6: num_10_to_99_to_string = "sechzig"
            case 7: num_10_to_99_to_string = "siebzig"
            case 8: num_10_to_99_to_string = "achtzig"
            case 9: num_10_to_99_to_string = "neunzig"
            case else
          end select ' }
       end if ' }

       if lang = "en" then ' {
          num_10_to_99_to_string = num_10_to_99_to_string & num_1_to_9_to_string(Right(TensText, 1), lang)  ' Retrieve ones place.
       else
          if right(tensText, 1) >= "2" and right(tensText, 1) <= "9" then
             num_10_to_99_to_string = num_1_to_9_to_string(right(tensText, 1), lang) & "und" & num_10_to_99_to_string
          end if
       end if ' }

    end if ' }

    num_10_to_99_to_string = num_10_to_99_to_string

end function ' }

  private function num_1_to_9_to_string(digit as string, lang as string, optional de_ein as string = "ein") as string ' {

   if lang = "en" then ' {

        select case val(digit) ' {
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
        end select ' }

   ' }
   elseIf lang = "de" then ' {

        select case val(digit) ' {
           case 1: num_1_to_9_to_string = de_ein
'                if de_eins then
'                  num_1_to_9_to_string = "eins"
'                else
'                  num_1_to_9_to_string = "ein"
'               end if
           case 2: num_1_to_9_to_string = "zwei"
           case 3: num_1_to_9_to_string = "drei"
           case 4: num_1_to_9_to_string = "vier"
           case 5: num_1_to_9_to_string = "fünf"
           case 6: num_1_to_9_to_string = "sechs"
           case 7: num_1_to_9_to_string = "sieben"
           case 8: num_1_to_9_to_string = "acht"
           case 9: num_1_to_9_to_string = "neun"
           case else: num_1_to_9_to_string = ""
        end select ' }

   end if ' }

end function ' }
