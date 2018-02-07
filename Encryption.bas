Attribute VB_Name = "Encryption"
'################################################################################################################################################
'NAME: Public Function hEncrypt(whatToEncrypt As String, key As String) As String
'DESCRIPTION: Encrypts a string.
'EXPECTS: 'whatToEncrypt' as a string containing the text that we wish to encrypt.
'RETURNS: 'newString' as a string containing the encrypted text.
'PRECONDITIONS: NONE
'POSTCONDITIONS: NONE
'AVAILABLE ENCRYPTION METHODS:
'        1) hCase (string)
'        2) hInverse (string)
'        3) hFold (string)
'        4) hReverse (string)
'        5) hXor (string,key)
'################################################################################################################################################
Public Function hEncrypt(whatToEncrypt As String, key As String) As String
    Dim newString As String
    Dim howMany As Long
    newString = whatToEncrypt
    newString = hXor(newString, key)
    For howMany = 1 To 4
        newString = hCase(newString)
        newString = hInverse(newString)
    Next
    For howMany = 1 To 3
        newString = hFold(newString)
    Next
    newString = hReverse(newString)
    newString = hXor(newString, key)
    hEncrypt = newString
End Function

'################################################################################################################################################
'NAME: Public Function hDecrypt(whatToDecrypt As String, key As String) As String
'DESCRIPTION: Decrypts a string.
'EXPECTS: 'whatToDecrypt' as a string containing the text that we wish to encrypt.
'RETURNS: 'newString' as a string containing the encrypted text.
'PRECONDITIONS: NONE
'POSTCONDITIONS: NONE
'AVAILABLE DECRYPTION METHODS:
'        1) hXor (string,key)
'        2) hReverse (string)
'        3) hUnfold (string)
'        4) hInverse (string)
'        5) hUncase (string)
'################################################################################################################################################
Public Function hDecrypt(whatToDecrypt As String, key As String) As String
    Dim newString As String
    Dim howMany As Long
    newString = whatToDecrypt
    newString = hXor(newString, key)
    newString = hReverse(newString)
    For howMany = 1 To 3
        newString = hUnfold(newString)
    Next
    For howMany = 1 To 4
        newString = hInverse(newString)
        newString = hUncase(newString)
    Next
    newString = hXor(newString, key)
    hDecrypt = newString
End Function

'################################################################################################################################################
'NAME: Private Function hXor(origString As String, key As String) As String
'DESCRIPTION: Performs an Xor operation on a string ('origString') using a string ('key') as the key.
'EXPECTS:
'        1) 'origString' as a string containing the data with which we are going to 'Xor' the key.
'        2) 'key' as a string containing the key of our 'Xor' operation.
'RETURNS: 'temp' as a string containing the 'Xor'ed data.
'PRECONDITIONS: NONE
'POSTCONDITIONS: NONE
'################################################################################################################################################
Private Function hXor(origString As String, key As String) As String
    Dim temp As String
    Dim counter As Long
    For counter = 1 To Len(origString)
        temp = temp + Chr(Asc(Mid(key, (counter Mod Len(key) + 1))) Xor Asc(Mid(origString, counter, 1)))
    Next
    hXor = temp
End Function

'################################################################################################################################################
'NAME: Public Function hGetMessage() As String
'DESCRIPTION: Retrieves and returns an embedded message from an image ('Picture2') containing an embedded message.
'EXPECTS: NOTHING
'RETURNS: 'phynalString' as a string containing the retrieved text.
'PRECONDITIONS: 'AppMainForm' as a form containing 'Picture2' as a PictureBox control.
'POSTCONDITIONS: NONE
'################################################################################################################################################
Public Function hGetMessage() As String 'Return a string containing the decrypted text message.
    Dim i As Long, j As Long, k As Long, n As Long, pix(0 To 2) As Long
    Dim tx As String, nmd As Long, start As Integer
    Dim endmess As String, comp(1 To 8) As Long, Ch As Long
    Dim phynalString As String
    For i = 0 To AppMainForm.Picture2.ScaleWidth - 1 'Do all this for each column of the picture.
        For j = 0 To AppMainForm.Picture2.ScaleHeight - 1 'Do all this for each row of the picture.
            nmd = n Mod 3 'n starts as zero. It is divided by 3, and the remainder is placed in 'nmd'.
            If nmd = 0 Then 'If we are on an 'nmd' that is divisible by 3 then
                If start < 14 Then 'If 'start' is less than 14 then
                    start = start + 1 'Increment 'start' by 1
                    If start = 14 And tx <> "start message" Then 'if start is 14 and tx isn't "start message" then
                        phynalString = "THIS PICTURE HAS NO SECRET MESSAGE" 'Tell the phools that there is no phreakin' message in the piccy.
                        hGetMessage = phynalString 'Exit the subroutine.
                        Exit Function
                    ElseIf start = 14 Then 'otherwise if start is 14 then
                        tx = "" 'set tx to nothing
                    End If
                End If
                Ch = 0 'ch' gets set to zero
                pix(nmd) = AppMainForm.Picture2.Point(i, j) 'We get a pixel out of the picture and place it into the nmd'th element of the 'pix' array.
                comp(8) = ((pix(nmd) And RGB(255, 0, 0)) Mod 2) 'We get the 8th bit of the 'comp' array from our data and junk.
                comp(7) = (((pix(nmd) And RGB(0, 255, 0)) \ 256) Mod 2) 'We get the 7th bit value of the 'comp' array.
                comp(6) = (((pix(nmd) And RGB(0, 0, 255)) \ 65536) Mod 2) 'We get the 6th bit value of the 'comp' array.
                For k = 8 To 6 Step -1 'Do something phunky with our 'ch' thingy.
                    Ch = Ch + (2 ^ (k - 1)) * comp(k) 'Translate the bit into a decimal value
                Next k
            End If
            If nmd = 1 Then 'If we have a remainder of 1 when we calculate our 'nmd' position, then do the following.
                pix(nmd) = AppMainForm.Picture2.Point(i, j) 'We get a pixel out of the picture and place it into the nmd'th element of the 'pix' array.
                comp(5) = ((pix(nmd) And RGB(255, 0, 0)) Mod 2) 'We get the 5th bit value of the 'comp' array.
                comp(4) = (((pix(nmd) And RGB(0, 255, 0)) \ 256) Mod 2) 'We get the 4th bit value of the 'comp' array.
                comp(3) = (((pix(nmd) And RGB(0, 0, 255)) \ 65536) Mod 2) 'We get the 3rd bit value of the 'comp' array.
                For k = 5 To 3 Step -1 'Do some wacky vomit with the 'ch' thingy.
                    Ch = Ch + (2 ^ (k - 1)) * comp(k) 'Translate the bit into a decimal value
                Next k
            End If
            If nmd = 2 Then 'If we have a reaminder of 2 when we calculate our 'nmd' position, then do the following.
                pix(nmd) = AppMainForm.Picture2.Point(i, j) 'We get a pixel out of the picture and place it into the nmd'th element of the 'pix' array.
                comp(2) = ((pix(nmd) And RGB(255, 0, 0)) Mod 2) 'We get the 4th bit value of the 'comp' array.
                comp(1) = (((pix(nmd) And RGB(0, 255, 0)) \ 256) Mod 2) 'We get the 3rd bit value of the 'comp' array.
                For k = 2 To 1 Step -1 'Do some wacky vomit with the 'ch' thingy.
                    Ch = Ch + (2 ^ (k - 1)) * comp(k) 'Translate the bit into a decimal value.  At this point, it's the ascii value
                Next k                                'of the character that we are looking for.
            End If
            n = n + 1 'Go to the next 'n'
            If n = 3 Then 'If we are at 3, go back to zero
                n = 0
                tx = tx & Chr(Ch) 'stick the character onto the end of our string (or something)
            End If
            endmess = Right(tx, 11) ' Check to see if the message is finished
            If endmess = "end message" Then 'If it is, then clip off the 'end message' string from the end of the message and return the fetcher.
                phynalString = Left(tx, Len(tx) - 11)
                hGetMessage = phynalString
                Exit Function 'Exit the subroutine.
            End If
        Next j
    Next i
    
End Function

'################################################################################################################################################
'NAME: Private Function ByteToBin(n As Integer) As String
'DESCRIPTION: Translates the incoming decimal value of a character into an 8-character string representing the binary values of that incoming value
'EXPECTS: 'n' as an integer which contains the ascii value of some character.
'RETURNS: 'j' as a string containing the 8-character string representing the binary values of the decimal value that was passed in
'PRECONDITIONS: NONE
'POSTCONDITIONS: NONE
'################################################################################################################################################
Private Function ByteToBin(n As Integer) As String   'This function transforms an integer (which is the
    Dim j As String                                 'the ascii code of a character) into a string (which
    Do While n >= 1                                 'is the binary representation of the ascii code)
    j = n Mod 2 & j
    n = n \ 2
    Loop
    If Len(j) < 8 Then j = String(8 - Len(j), "0") & j
    ByteToBin = j
End Function

'################################################################################################################################################
'NAME: Public Sub hPutMessage(whatToEmbed As String)
'DESCRIPTION: Embeds the incoming text ('whatToEmbed') into a graphic ('Picture2' on 'AppMainForm').
'EXPECTS: 'whatToEmbed' as a string containing the information that is going to be embedded into the graphic.
'RETURNS: NOTHING
'PRECONDITIONS:
'        1) 'AppMainForm' as a form containing 'Picture2' as a PictureBox control.
'        2) 'ByteToBin' function (in this module, I hope).
'POSTCONDITIONS: NONE
'################################################################################################################################################
Public Function hPutMessage(whatToEmbed As String) As Boolean 'Pass in the text.
    Dim i As Long, j As Long, tx As String, binaryByteString As String, NrPix As Long
    Dim pix(0 To 2) As Long, bitmapWidth As Long, bitmapHeight As Long
    Dim r As Long, g As Long, b As Long, comp(1 To 8) As Long
    Dim aa(0 To 2) As Long, bb(0 To 2) As Long
    tx = "start message" & whatToEmbed & "end message"  'This contains the text of the message to be encrypted.
    bitmapWidth = AppMainForm.Picture2.ScaleWidth 'This is the number of pixels wide of the bitmap graphic
    bitmapHeight = AppMainForm.Picture2.ScaleHeight 'This is the number of pixels high of the bitmap graphic
    If Len(tx) * 3 > bitmapHeight * bitmapWidth Then 'If the length of the text is larger than one third of the number of pixels in the bitmap
        tx = MsgBox("Text is " & Len(tx) * 3 - bitmapWidth * bitmapHeight & " characters longer than this picture can store", vbCritical) 'Display ERROR!
        hPutMessage = False
        Exit Function 'Exit the Function
    End If
    For i = 1 To Len(tx) 'For each character in the string that we are embedding into the graphic, do the following:
        binaryByteString = ByteToBin(Asc(Mid(tx, i, 1))) 'Get the ith byte and make it's ascii value into it's binary equivalent.
        NrPix = (CLng(i) - 1) * 3 'Take i, change it to a long, subtract 1, and multply by three, and stick it into NrPix. Keeps track of the current pixel.
        aa(0) = (NrPix Mod bitmapHeight) 'Take NrPix and divide it by 'bitmapHeight', and put the remainder into the first element of the 'aa' array.
        bb(0) = (NrPix \ bitmapHeight) 'Take NrPix and divide it by 'bitmapHeight', and put the whole integer result into the first element of the 'bb' array.
        pix(0) = AppMainForm.Picture2.Point(bb(0), aa(0)) 'Retrieve the first pixel in the group of three and put it into the first element of the 'pix' array.
        r = (pix(0) And RGB(255, 0, 0)) - (pix(0) And RGB(255, 0, 0)) Mod 2: comp(1) = r
        'Make the 'RED' value of the pixel even; place it into the 1st element of the 'comp' array.
        g = ((pix(0) And RGB(0, 255, 0)) \ 256) - ((pix(0) And RGB(0, 255, 0)) \ 256) Mod 2: comp(2) = g
        'Make the 'GREEN' value of the pixel even; place it into the 2nd element of the 'comp' array.
        b = ((pix(0) And RGB(0, 0, 255)) \ 65536) - ((pix(0) And RGB(0, 0, 255)) \ 65536) Mod 2: comp(3) = b
        'Make the 'BLUE' value of the pixel even; place it into the 3rd element of the 'comp' array.
        NrPix = NrPix + 1 'Go to the next pixel, essentially.
        aa(1) = (NrPix Mod bitmapHeight) 'Find out which pixel column it corresponds to. 'X' axis value.  Horizontal value.
        bb(1) = (NrPix \ bitmapHeight) 'Find out which pixel row it corresponds to. 'Y' axis value.  Vertical value.
        pix(1) = AppMainForm.Picture2.Point(bb(1), aa(1)) 'Retrieve the second pixel in the group of three and put it into the second element of the 'pix' array.
        r = (pix(1) And RGB(255, 0, 0)) - (pix(1) And RGB(255, 0, 0)) Mod 2: comp(4) = r
        'Make the 'RED' value of the pixel even; place it into the 4th element of the 'comp' array.
        g = ((pix(1) And RGB(0, 255, 0)) \ 256) - ((pix(1) And RGB(0, 255, 0)) \ 256) Mod 2: comp(5) = g
        'Make the 'GREEN' value of the pixel even; place it into the 5th element of the 'comp' array.
        b = ((pix(1) And RGB(0, 0, 255)) \ 65536) - ((pix(1) And RGB(0, 0, 255)) \ 65536) Mod 2: comp(6) = b
        'Make the 'BLUE' value of the pixel even; place it into the 6th element of the 'comp' array.
        NrPix = NrPix + 1 'Go to the next pixel.
        aa(2) = (NrPix Mod bitmapHeight) 'Find out which pixel column it corresponds to. 'X' axis value.  Horizontal value.
        bb(2) = (NrPix \ bitmapHeight) 'Find out which pixel row it corresponds to. 'Y' axis value.  Vertical value.
        pix(2) = AppMainForm.Picture2.Point(bb(2), aa(2)) 'Retrieve the third pixel in the group of three and put it into the third element of the 'pix' array.
        r = (pix(2) And RGB(255, 0, 0)) - (pix(2) And RGB(255, 0, 0)) Mod 2: comp(7) = r
        'Make the 'RED' value of the pixel even; place it into the 7th element of the 'comp' array.
        g = ((pix(2) And RGB(0, 255, 0)) \ 256) - ((pix(2) And RGB(0, 255, 0)) \ 256) Mod 2: comp(8) = g
        'Make the 'GREEN' value of the pixel even; place it into the 8th element of the 'comp' array.
        b = ((pix(2) And RGB(0, 0, 255)) \ 65536) 'This is the 9th element, and since we only have 8 bits, we don't touch this element.
        For j = 1 To 8                                            'This 'for' loop is where we change all of the elements of the 'comp' array from
        comp(j) = comp(j) + CInt(Mid(binaryByteString, j, 1)) * 1 'the even number that they represent into the even/odd bit representation of the
        Next j                                                    'binary string from the current character.
        AppMainForm.Picture2.PSet (bb(0), aa(0)), RGB(comp(1), comp(2), comp(3)) 'This is where, using our arrays, we re-assign all of the pixel information
        AppMainForm.Picture2.PSet (bb(1), aa(1)), RGB(comp(4), comp(5), comp(6)) 'back into the picture.  The pixels are modified according to the bit value
        AppMainForm.Picture2.PSet (bb(2), aa(2)), RGB(comp(7), comp(8), b)       'at their corresponding positions.
    Next i 'Go to the next character in our string to encrypt.
    hPutMessage = True
End Function
'################################################################################################################################################
'NAME: Private Function hCase(whatToCase As String) As String
'DESCRIPTION: Subtracts 23 from the Ascii code of the first character, adds 13 to the ascii code of the second character, subtracts 69 from the
'             ascii code of the third character, simply adds the fourth character, and adds 27 to the ascii code of the fifth character.  It then
'             treats the sixth character as it does the first character and thus starts the entire process all over again.
'EXPECTS: 'whatToCase' as a string containing the text that we wish to encrypt.
'RETURNS: 'tempString' as a string containing the encrypted text.
'PRECONDITIONS: NONE
'POSTCONDITIONS: NONE
'################################################################################################################################################
Private Function hCase(whatToCase As String) As String
Dim tempString As String
Dim counter As Long
Dim tempAscii As Integer
tempString = ""
For counter = 1 To Len(whatToCase)
    Select Case (counter Mod 5)
        Case 0:
            tempAscii = Asc(Mid(whatToCase, counter, 1))
            tempAscii = tempAscii + 27
            If tempAscii > 255 Then tempAscii = tempAscii - 256
            tempString = tempString + Chr(tempAscii)
        Case 1:
            tempAscii = Asc(Mid(whatToCase, counter, 1))
            tempAscii = tempAscii - 23
            If tempAscii < 0 Then tempAscii = tempAscii + 256
            tempString = tempString + Chr(tempAscii)
        Case 2:
            tempAscii = Asc(Mid(whatToCase, counter, 1))
            tempAscii = tempAscii + 13
            If tempAscii > 255 Then tempAscii = tempAscii - 256
            tempString = tempString + Chr(tempAscii)
        Case 3:
            tempAscii = Asc(Mid(whatToCase, counter, 1))
            tempAscii = tempAscii - 69
            If tempAscii < 0 Then tempAscii = tempAscii + 256
            tempString = tempString + Chr(tempAscii)
        Case 4:
            tempString = tempString + Mid(whatToCase, counter, 1)
    End Select
Next
hCase = tempString
End Function

'################################################################################################################################################
'NAME: Private Function hUncase(whatToUncase As String) As String
'DESCRIPTION: Adds 23 to the Ascii code of the first character, subtracts 13 from the ascii code of the second character, adds 69 to the ascii
'             code of the third character, simply adds the fourth character, and subtracts 27 from the ascii code of the fifth character.  It
'             then treats the sixth character as it does the first character and thus starts the entire process all over again.
'EXPECTS: 'whatToUncase' as a string containing the text that we wish to decrypt.
'RETURNS: 'tempString' as a string containing the decrypted text.
'PRECONDITIONS: NONE
'POSTCONDITIONS: NONE
'################################################################################################################################################
Private Function hUncase(whatToUncase As String) As String
Dim tempString As String
Dim counter As Long
Dim tempAscii As Integer
tempString = ""
For counter = 1 To Len(whatToUncase)
    Select Case (counter Mod 5)
        Case 0:
            tempAscii = Asc(Mid(whatToUncase, counter, 1))
            tempAscii = tempAscii - 27
            If tempAscii < 0 Then tempAscii = tempAscii + 256
            tempString = tempString + Chr(tempAscii)
        Case 1:
            tempAscii = Asc(Mid(whatToUncase, counter, 1))
            tempAscii = tempAscii + 23
            If tempAscii > 255 Then tempAscii = tempAscii - 256
            tempString = tempString + Chr(tempAscii)
        Case 2:
            tempAscii = Asc(Mid(whatToUncase, counter, 1))
            tempAscii = tempAscii - 13
            If tempAscii < 0 Then tempAscii = tempAscii + 256
            tempString = tempString + Chr(tempAscii)
        Case 3:
            tempAscii = Asc(Mid(whatToUncase, counter, 1))
            tempAscii = tempAscii + 69
            If tempAscii > 255 Then tempAscii = tempAscii - 256
            tempString = tempString + Chr(tempAscii)
        Case 4:
            tempString = tempString + Mid(whatToUncase, counter, 1)
    End Select
Next
hUncase = tempString
End Function

'################################################################################################################################################
'NAME: Private Function hInverse(whatToInverse As String) As String
'DESCRIPTION: Subtracts the ascii value of the current character from 255 and replaces the original character with the result.  This algorithm
'             both encrypts plain text and decrypts encrypted text.
'EXPECTS: 'whatToInverse' as a string containing the text that we wish to encrypt.
'RETURNS: 'tempString' as a string containing the encrypted text.
'PRECONDITIONS: NONE
'POSTCONDITIONS: NONE
'################################################################################################################################################
Private Function hInverse(whatToInverse As String) As String
    Dim tempString As String
    Dim counter As Long
    tempString = ""
    For counter = 1 To Len(whatToInverse)
        tempString = tempString + Chr(255 - Asc(Mid(whatToInverse, counter, 1)))
    Next
    hInverse = tempString
End Function

'################################################################################################################################################
'NAME: Private Function hFold(whatToFold As String) As String
'DESCRIPTION: "Folds" the bits of a string in half.  It takes the last bit of a string and puts it first.  It takes the first bit of the string
'             and places it second.  It takes the second-to-the-last bit and places it third.  It takes the second bit and places it fourth.  It
'             takes the third-to-the-last bit and places it fifth.  It takes the third bit and places it sixth and so on until the middle two
'             bits are switched.  It then makes characters back out of the entire thing and returns that whole string encrypted.
'EXPECTS: 'whatToFold' as a string containing the text that we wish to encrypt.
'RETURNS: 'tempString' as a string containing the encrypted text.
'PRECONDITIONS: NONE
'POSTCONDITIONS: NONE
'################################################################################################################################################
Private Function hFold(whatToFold As String) As String
    Dim tempString As String
    Dim byteA As String
    Dim byteB As String
    Dim strOrBin As String
    Dim currDown As Long
    Dim currUp As Long
    Dim lenModTwo As Integer
    Dim counter As Long
    Dim superTempString As String
    strOrBin = "binary"
    For counter = 1 To Len(whatToFold)
        If Mid(whatToFold, counter, 1) <> "0" And Mid(whatToFold, counter, 1) <> "1" Then strOrBin = "string"
    Next
    currUp = 0
    currDown = Len(whatToFold) + 1
    lenModTwo = ((currDown + 1) Mod 2) + 1
    tempString = ""
    If Len(whatToFold) > 1 Then
        Do
            currUp = currUp + 1
            currDown = currDown - 1
            byteA = Mid(whatToFold, currDown, 1)
            byteB = Mid(whatToFold, currUp, 1)
            If strOrBin = "string" Then
                superTempString = hFold(ByteToBin(Asc(byteA)) + ByteToBin(Asc(byteB)))
                byteA = Chr(BinToDec(Left(superTempString, 8)))
                byteB = Chr(BinToDec(Right(superTempString, 8)))
            End If
                tempString = tempString + byteA + byteB
        Loop Until ((currUp + lenModTwo = currDown) Or (currUp = currDown))
        If lenModTwo = 2 Then
            currDown = currDown - 1
            tempString = tempString + Chr(BinToDec(hFold(ByteToBin(Asc(Mid(whatToFold, currDown, 1))))))
        End If
    Else
        tempString = whatToFold
    End If
    hFold = tempString
End Function

'################################################################################################################################################
'NAME: Public Function BinToDec(stringToConvert As String) As Integer
'DESCRIPTION: Converts a string that represents an 8 bit binary number into a decimal integer.
'EXPECTS: 'stringToConvert' as a string containing the 8 bit binary number representation.
'RETURNS: 'newNumber' as a decimal integer of the 8 bit binary number string representation that was passed in.
'PRECONDITIONS: NONE
'POSTCONDITIONS: NONE
'################################################################################################################################################
Public Function BinToDec(stringToConvert As String) As Integer
    Dim newNumber As String
    Dim counter As Integer
    newNumber = 0
    For counter = 0 To (Len(stringToConvert) - 1)
        If Mid(stringToConvert, (Len(stringToConvert) - counter), 1) = "1" Then newNumber = newNumber + (2 ^ counter)
    Next
    BinToDec = newNumber
End Function

'################################################################################################################################################
'NAME: Private Function hUnfold(whatToUnfold As String) As String
'DESCRIPTION: "Unfolds" a "folded" string by taking the first character and placing it at the end.  Then takes the second character and places it
'             at the beginning.  Then takes the third character and places it next-to-last.  Then takes the fourth character and places it second.
'             Does this until the end of the encrypted string is reached.
'EXPECTS: 'whatToUnfold' as a string containing the text that we wish to decrypt.
'RETURNS: 'tempString' as a string containing the decrypted text.
'PRECONDITIONS: NONE
'POSTCONDITIONS: NONE
'################################################################################################################################################
Private Function hUnfold(whatToUnfold As String) As String
    Dim tempString As String
    Dim lastHalf As String
    Dim firstHalf As String
    Dim charA As String
    Dim charB As String
    Dim strOrBin As String
    Dim superTempString As String
    Dim counter As Long
    strOrBin = "binary"
    For counter = 1 To Len(whatToUnfold)
        If Mid(whatToUnfold, counter, 1) <> "0" And Mid(whatToUnfold, counter, 1) <> "1" Then strOrBin = "string"
    Next
    tempString = ""
    lastHalf = ""
    firstHalf = ""
    lastHalfCounter = 1
    firstHalfCounter = 2
    If Len(whatToUnfold) > 1 Then
    For counter = 1 To Len(whatToUnfold)
        charA = Mid(whatToUnfold, counter, 1)
        charB = Mid(whatToUnfold, counter + 1, 1)
        If strOrBin = "string" Then
            If Len(charB) > 0 Then
                superTempString = hUnfold(ByteToBin(Asc(charA)) + ByteToBin(Asc(charB)))
            Else
                superTempString = hUnfold(ByteToBin(Asc(charA)))
            End If
            charA = Chr(BinToDec(Left(superTempString, 8)))
            If Len(superTempString) > 8 Then charB = Chr(BinToDec(Right(superTempString, 8)))
        End If
        lastHalf = charA + lastHalf
        firstHalf = firstHalf + charB
        counter = counter + 1
    Next
    tempString = firstHalf + lastHalf
    Else
        tempString = whatToUnfold
    End If
    hUnfold = tempString
End Function

'################################################################################################################################################
'NAME: Private Function hReverse(whatToReverse as String) As String
'DESCRIPTION: Reverses a string.
'EXPECTS: 'whatToReverse' as a string containing the text that we wish to reverse.
'RETURNS: 'modifiedString' as a string containing the reversed text.
'PRECONDITIONS: NONE
'POSTCONDITIONS: NONE
'################################################################################################################################################
Private Function hReverse(whatToReverse As String) As String
    Dim modifiedString As String
    Dim counter As Long
    modifiedString = ""
    For counter = 1 To Len(whatToReverse)
        modifiedString = modifiedString + Mid(whatToReverse, (Len(whatToReverse) - counter + 1), 1)
    Next
    hReverse = modifiedString
End Function
