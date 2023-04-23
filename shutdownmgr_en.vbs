option explicit
dim a, b, c, d, e, f, h, m, obj'a-main interface, b-input box, c-command time, d-system time, e-Scheduled shutdown box, f-length of time, h-input hours, m-input minutes, obj-command
const NAME="Schedule Shutdown Manager"
sub command'shutdown command section
   dim fso, f'Create a command file
       set fso = CreateObject("Scripting.FileSystemObject")
       set f = fso.CreateTextFile(".\sd.bat", true)
           f.Write("shutdown /s /t ")
           f.WriteLine(c)'Write the command time and wrap
           f.Write("del sd.bat")'Write delete command
           f.Close()
       set f = nothing
       set fso = nothing
   wscript.createobject("wscript.shell").run ".\sd.bat"'Run the command file
 end sub
sub cancel'Cancel the shutdown command section
   Set obj = createobject("wscript.shell")
       obj.run "cmd /c shutdown /a"
 end sub
sub dsgj'Scheduled shutdown section
  do
    b=(inputbox("Enter the timed shutdown time (HM)","Scheduled Shutdown"))
    f=len(b)
    h=left(b,2)
    m=right(b,2)
    d=int(Hour(now)&Minute(now))
    if int(b)>0 and f=4 then'Scheduled shutdown command
      if int(b)>int(d) then
        c=(int(h)-int(Hour(Now)))*3600+(int(m)-int(Minute(Now)))*60-int(Second(Now))
      elseif int(b)<int(d) then
        c=(23+int(h)-int(Hour(Now)))*3600+(60+int(m)-int(Minute(Now)))*60-int(Second(Now))
      end if
      call cancel
      call command
    elseif b=vbNullString then
    else'Error prompt
      msgbox "Error",,"Scheduled Shutdown"
    end if
  loop until int(b)>0 and f=4 or b=vbNullString
 end sub
sub ysgj'Delayed shutdown section
  do
    b=(inputbox("Input Delayed Shutdown Time (min)","Delayed Shutdown"))
    if int(b)>0 then'Delayed shutdown command
      c=int(b)*60
      call cancel
      call command
    elseif b=vbNullString then
    else'Error prompt
      msgbox "Error",,"Delayed shutdown"
    end if
  loop until b=vbNullString or int(b)>0
 end sub
do
a=(msgbox("Select the action you want to take:" & vbcrlf & "  Yes = Scheduled shutdown" & vbcrlf & "  No = Cancel the plan",3,[NAME]))'Main interface
if a=vbyes then'Scheduled shutdown box
  e=(msgbox("Select the action you want to take:" & vbcrlf & "  Yes = Scheduled shutdown" & vbcrlf & "  No = Delayed shutdown",3,[NAME]))
  if e=vbyes then
    call dsgj
  elseif e=vbno then
    call ysgj
  elseif e=vbcancel then
  end if
elseif a=vbno then'Cancel the plan
  call cancel
elseif a=vbcancel then'Exit
end if
loop until a=vbcancel