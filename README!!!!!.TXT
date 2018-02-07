#Description#

Written in VB6.0 in 2002

This program allows you to:

1) Encrypt text and save it to a text file.
2) Embed text into an image and save the image (with
   text embedded) out to a file.
3) Encrypt text and embed it into an image file and
   save the image (with text embedded) out to a file.

The buttons across the top of the interface are for
the text boxes.  The buttons across the middle are
for the image.

If you can't figure it out, I think an excellent 
career oportunity awaits you in the "Worm Farming"
industry.

I am including a couple of image files.  "None.gif" 
and "Message.bmp".  "None.gif" is a plain GIF with
no message.  "Message.bmp" is a bitmap with an
encrypted message embedded.  The key for decoding
that message is README!!!!!.TXT . Also, take note:
Normal bitmaps and those with embedded text in them
are very nearly the same size.

Please note that you can open gifs, bitmaps, and jpegs.
You can only save out bitmaps!  This may change in 
later versions, but for this one, it didn't.

Test it out and let me know if anything happens that
you think shouldn't happen or doesn't that should.

Updates:
1.1 - Fixed bug in encryption algorithms.
    - Added ability to do 'Xor' operations using key.
    - Added menuing system.
    - Added 'Help' and 'About' screens.
    - Loads characters with the ascii value of 8 
      properly (as this is a backspace, it was
      actually deleting characters when it shouldn't)
