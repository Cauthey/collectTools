Msg=Inputbox("input content")
num=Inputbox("input times")

a=1
b=num

set wshshell=CreateObject("wscript.shell")  '创建Windows的shell对象打开shell窗口
wscript.sleep 5000

for i=a to b   '循环发送
	str = Msg + " x " + cstr(i) 
	wshshell.Run "cmd.exe /c echo " & str & " | clip.exe",0,True '内容放入剪切板
	wshshell.sendKeys "^v"  '粘贴要发送的消息内容

	wscript.sleep 100
	wshshell.sendKeys "{ENTER}"    '按enter键进行发送
next
wscript.quit
