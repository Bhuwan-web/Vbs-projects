option explicit
dim a, b, x
set x= createobject("wscript.shell")
a= inputbox("Type facebook id","facebook login")
b= inputbox("type facebook password","facebook login")
x.run "https://www.facebook.com"
wscript.sleep 6000
x.sendkeys a
x.sendkeys "{enter}"
x.sendkeys b
x.sendkeys "{entre}"
wscript.quit