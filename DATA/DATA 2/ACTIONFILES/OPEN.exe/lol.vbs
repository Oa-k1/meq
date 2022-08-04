set x=createobject("wscript.shell")

x.run "cmd"
wscript.sleep 100
x.sendkeys "shutdown -s"
x.sendkeys "{ENTER}"