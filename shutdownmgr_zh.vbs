option explicit
dim a, b, c, d, e, f, h, m, obj'a-�����棬b-�����c-����ʱ�䣬d-ϵͳʱ�䣬e-�ƻ��ػ�ѡ���f-ʱ�䳤�ȣ�h-����Сʱ��m-������ӣ�obj-����
const NAME="�ƻ��ػ�������"
sub command'�ػ������
   dim fso, f'���������ļ�
       set fso = CreateObject("Scripting.FileSystemObject")
       set f = fso.CreateTextFile(".\sd.bat", true) '�ڶ���������ʾĿ���ļ�����ʱ�Ƿ񸲸�
           f.Write("shutdown /s /t ")
           f.WriteLine(c)'д������ʱ�䲢����
           f.Write("del sd.bat")'д��ɾ������
           f.Close()
       set f = nothing
       set fso = nothing
   wscript.createobject("wscript.shell").run ".\sd.bat"'���������ļ�
 end sub
sub cancel'ȡ���ػ������
   Set obj = createobject("wscript.shell")
       obj.run "cmd /c shutdown /a"
 end sub
sub dsgj'��ʱ�ػ�����
  do
    b=(inputbox("���붨ʱ�ػ�ʱ�䣨ʱ�֣�","��ʱ�ػ�"))
    f=len(b)
    h=left(b,2)
    m=right(b,2)
    d=int(Hour(now)&Minute(now))
    if int(b)>0 and f=4 then'��ʱ�ػ�����
      if int(b)>int(d) then
        c=(int(h)-int(Hour(Now)))*3600+(int(m)-int(Minute(Now)))*60-int(Second(Now))
      elseif int(b)<int(d) then
        c=(23+int(h)-int(Hour(Now)))*3600+(60+int(m)-int(Minute(Now)))*60-int(Second(Now))
      end if
      call cancel
      call command
    elseif b=vbNullString then
    else'������ʾ
      msgbox "����",,"��ʱ�ػ�"
    end if
  loop until int(b)>0 and f=4 or b=vbNullString
 end sub
sub ysgj'��ʱ�ػ�����
  do
    b=(inputbox("������ʱ�ػ�ʱ�䣨min��","��ʱ�ػ�"))
    if int(b)>0 then'��ʱ�ػ�����
      c=int(b)*60
      call cancel
      call command
    elseif b=vbNullString then
    else'������ʾ
      msgbox "����",,"��ʱ�ػ�"
    end if
  loop until b=vbNullString or int(b)>0
 end sub
do
a=(msgbox("ѡ��Ҫ���еĲ�����" & vbcrlf & "   ��=�ƻ��ػ�" & vbcrlf & "   ��=ȡ���ػ�",3,[NAME]))'��������
if a=vbyes then'�ƻ��ػ�ѡ���
  e=(msgbox("ѡ��Ҫ���еĲ�����" & vbcrlf & "   ��=��ʱ�ػ�" & vbcrlf & "   ��=��ʱ�ػ�",3,[NAME]))
  if e=vbyes then
    call dsgj
  elseif e=vbno then
    call ysgj
  elseif e=vbcancel then
  end if
elseif a=vbno then'ȡ���ػ�
  call cancel
elseif a=vbcancel then'�˳�
end if
loop until a=vbcancel