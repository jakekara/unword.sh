# unword.sh

Extract images from .docx files.

### deets

This OS X* shell script extracts images from a .docx word document. Useful
for my work as a newspaper editor. I receive a lot of word documents, but
we don’t use Word for our copy flow, so I try not to open Word if I can
avoid it. All this does is automate the ".zip trick" — changing the file
extension from .docx to .zip and then unzipping, and pulling the image
files out. That trick doesn’t work with .doc files, so neither does this
script.

Optional command line arguments:

-h (help screen); d (delete .docx file after extracting images) o (open
-.docx file in TextEdit after extracting images) By default, the script
-does not delete the .docx file or open it in TextEdit.

This code can be copy and pasted right into an Automator app (using
"execute shell script"), so you can drag files onto the app and they’ll be
run through this script with the default settings.

*This only thing that makes this OS X-specific is that it has an option
 to open the word doc in TextEdit. That can easily be removed

### Issues

So, when I wrote this, a couple of years ago now, I didn't realize this zip
trick also works for .pptx files, since I had no use for that. This could
be modified to work with .pptx and other types of files, by taking out the
file extension check.

I probably won't update this because I haven't used it in years, but it was
very handy when I was processing dozens of word documents per day.