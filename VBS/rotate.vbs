set a= createobject("wscript.shell")
set b= createobject("sapi.spvoice")
b.speak "lets rotate left"
a.sendkeys "^%{left}"
b.speak "more left"
wscript.sleep 150
a.sendkeys "^%{down}"
b.speak "more left"
wscript.sleep 150
a.sendkeys "^%{right}"
b.speak "last left"
wscript.sleep 150
a.sendkeys "^%{up}"
wscript.sleep 150
b.speak "hope you like it"