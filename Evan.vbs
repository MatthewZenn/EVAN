Dim Message, EVAN
Set shell=CreateObject("wscript.shell")

Start=MsgBox("welcome to EVAN 2016 v2.3! the EVAN program was designed by and is property of DroneTech under the direction of Thinkit.inc. You are not permitted to modify EVAN in any way. press ok to continue.",1+0,"EVAN �copyright Thinkit.inc 2016 all rights reserved")

If Start=vbOk then User=InputBox("Please enter your name","EVAN")
Set Speak=CreateObject("sapi.spvoice")
If User="[User_Name]" then Speak.Speak "welcome... master, today is " & Date & " and the time is " & Time else if Start=vbOk then Speak.Speak "hello..."+User+"...today is " & Date & "and the time is " & Time & "......allow me to introduce myself ...i am evan, the elite virtual assistance network"

If Start=vbOk then Speak.Speak "what can i do for you "+User
If Start=vbOk then Value=InputBox("input command: application name, search","EVAN")
If Start=vbOk then Speak.Speak "ok, I'm opening..."+Value

If Value="search" then Search=InputBox("what should i search for","EVAN") else Shell.Run (Value)
Set shell=CreateObject("wscript.shell")
If Value= "search" then Shell.Run ("https://www.google.com/?gws_rd=ssl#q="+Search)
If Value="search" then Speak.Speak "ok, i'm searching for..."+Search
