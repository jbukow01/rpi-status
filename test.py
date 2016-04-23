#!/usr/bin/env python3
#!/usr/bin/python

import cgi
import cgitb

print("Content-type: text/html\n")

cgitb.enable()

print("<html>\n<body>")
args = cgi.FieldStorage()
print("<h1>Lista wszystkich argumentów</h1>")
print("<pre>")
for x in args:
    print(x + "=" + args[x].value)
print("</pre>\n</body>\n</html>")