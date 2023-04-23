option explicit
dim a, b, c, d, e, f, h, m, obj'a-主界面，b-输入框，c-命令时间，d-系统时间，e-计划关机选择框，f-时间长度，h-输入小时，m-输入分钟，obj-命令
const NAME="计划关机管理器"
sub command'关机命令部分
   dim fso, f'创建命令文件
       set fso = CreateObject("Scripting.FileSystemObject")
       set f = fso.CreateTextFile(".\sd.bat", true) '第二个参数表示目标文件存在时是否覆盖
           f.Write("shutdown /s /t ")
           f.WriteLine(c)'写入命令时间并换行
           f.Write("del sd.bat")'写入删除命令
           f.Close()
       set f = nothing
       set fso = nothing
   wscript.createobject("wscript.shell").run ".\sd.bat"'运行命令文件
 end sub
sub cancel'取消关机命令部分
   Set obj = createobject("wscript.shell")
       obj.run "cmd /c shutdown /a"
 end sub
sub dsgj'定时关机部分
  do
    b=(inputbox("输入定时关机时间（时分）","定时关机"))
    f=len(b)
    h=left(b,2)
    m=right(b,2)
    d=int(Hour(now)&Minute(now))
    if int(b)>0 and f=4 then'定时关机命令
      if int(b)>int(d) then
        c=(int(h)-int(Hour(Now)))*3600+(int(m)-int(Minute(Now)))*60-int(Second(Now))
      elseif int(b)<int(d) then
        c=(23+int(h)-int(Hour(Now)))*3600+(60+int(m)-int(Minute(Now)))*60-int(Second(Now))
      end if
      call cancel
      call command
    elseif b=vbNullString then
    else'错误提示
      msgbox "错误",,"定时关机"
    end if
  loop until int(b)>0 and f=4 or b=vbNullString
 end sub
sub ysgj'延时关机部分
  do
    b=(inputbox("输入延时关机时间（min）","延时关机"))
    if int(b)>0 then'延时关机命令
      c=int(b)*60
      call cancel
      call command
    elseif b=vbNullString then
    else'错误提示
      msgbox "错误",,"延时关机"
    end if
  loop until b=vbNullString or int(b)>0
 end sub
do
a=(msgbox("选择要进行的操作：" & vbcrlf & "   是=计划关机" & vbcrlf & "   否=取消关机",3,[NAME]))'操作界面
if a=vbyes then'计划关机选择框
  e=(msgbox("选择要进行的操作：" & vbcrlf & "   是=定时关机" & vbcrlf & "   否=延时关机",3,[NAME]))
  if e=vbyes then
    call dsgj
  elseif e=vbno then
    call ysgj
  elseif e=vbcancel then
  end if
elseif a=vbno then'取消关机
  call cancel
elseif a=vbcancel then'退出
end if
loop until a=vbcancel