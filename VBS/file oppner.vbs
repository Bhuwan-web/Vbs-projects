option explicit
dim a, b, c,d
set a= createobject("wscript.shell")
b=inputbox("type file name with extention or full web address","File oppner")

a.run b
