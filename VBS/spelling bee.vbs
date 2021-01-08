do
x=inputbox("type text you want to make speak...                          note:- press 1 to exit this file","speaking s/w")
set speak=createobject("sapi.spvoice")
speak.speak x
loop while x<>a
