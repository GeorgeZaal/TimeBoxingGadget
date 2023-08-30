unit LoadIni;

interface
function  OpenParam (f_name, paramA, ch : string) : string;
function  SearchNoSting (f_name, paramA, ch : string) : integer;
procedure Re_Name(f_name, ch, paramA, paramB : string);
procedure Ch_Param(f_name, ch, paramA, paramB : string);

implementation

{OpenParam(f_name - имя открываемого файла
           paramA - первая часть выражения
           ch     - разделительный символ) : string;}
function OpenParam (f_name, paramA, ch : string) : string;
var
   f : textfile;
   tmp, ok : string;
   i, j : integer;
begin
   assignfile(f, f_name);
   reset(f);
while not eof(f) do
begin
readln(f, tmp);
if (pos(paramA, tmp) = 1) and (pos(ch, tmp) > 1) then
begin
for i := 1 to Length(tmp) do
   if tmp[i] = ch then
      begin
      for j := i + 1 to Length(tmp) do
         ok := ok + tmp[j];
      break;
      end;
While ok[1] = ' ' do Delete(ok, 1, 1);
While ok[Length(ok)] = ' ' do Delete(ok, Length(ok), 1);
end;
end;
   closefile(f);
   OpenParam := ok;
end;

{SearchNoSting(f_name - имя открываемого файла
               paramA - первая часть выражения
               ch     - разделительный символ) : integer;
			   ...
Функция выводит номер строки с указанной первой половиной выражения,
нумерация начинается с нуля (это первая строка)}
function SearchNoSting (f_name, paramA, ch : string) : integer;
var
   f : textfile;
   tmp : string;
   ok : integer;
begin
   ok := 0;
   assignfile(f, f_name);
   reset(f);
while not eof(f) do
begin
readln(f, tmp);
if (pos(paramA, tmp) = 1) and (pos(ch, tmp) > 1) then
   break
   else ok := ok + 1;
end;
   closefile(f);
   SearchNoSting := ok;
end;

{Re_Name(f_name - имя открываемого файла
         ch     - разделительный символ
         paramA - первая часть выражения
         paramB - то, на что её надо заменить);
		 ...
Заменять будет по маске: [paramA] [ch]}
procedure Re_Name(f_name, ch, paramA, paramB : string);
var
   f : textfile;
   tmp, super_tmp : string;
   i, count, ch_count  : integer;
   Base: array of string;
begin
// Обнуление счётчиков:
   count := 0;
   ch_count := 0;
   super_tmp := '';
// Открытие файла для поиска нужной строки:
   assignfile(f, f_name);
   reset(f);
while not eof(f) do
begin
count := count + 1;
readln(f, tmp);
SetLength(Base, Count);
Base[Count-1] := tmp;
if (pos(paramA, tmp) = 1) and (tmp[Length(paramA) + 2] = ch) then
   begin
   for i := 1 to Length(tmp) do
   if tmp[i] = ch then
      begin
      Delete(tmp, 1, i - 1);
      tmp := paramB + ' ' + tmp;
      ch_count := count;
      super_tmp := tmp;
      break;
      end;
   end;
end;
closefile(f);
// Если строка была заменена, то запись массива в файл:
if super_tmp <> '' then
   begin
   base[ch_count-1] := super_tmp;
   rewrite(f);
   for i := 0 to (Length(base)-1) do
      writeLn(f, base[i]);
   closefile(f);
   end;
end;

{Ch_Param(f_name - имя открываемого файла
          ch     - разделительный символ
          paramA - первая часть выражения
          paramB - то, на что надо заменить вторую часть);
		  ...
Заменять будет по маске: [ch] [paramB]}
procedure Ch_Param(f_name, ch, paramA, paramB : string);
var
   f : textfile;
   tmp, super_tmp : string;
   i, count, ch_count  : integer;
   Base: array of string;
begin
// Обнуление счётчиков:
   count := 0;
   ch_count := 0;
   super_tmp := '';
// Открытие файла для поиска нужной строки:
   assignfile(f, f_name);
   reset(f);
while not eof(f) do
begin
count := count + 1;
readln(f, tmp);
SetLength(Base, Count);
Base[Count-1] := tmp;
if (pos(paramA, tmp) = 1) and (tmp[Length(paramA) + 2] = ch) then
   begin
   for i := 1 to Length(tmp) do
   if tmp[i] = ch then
      begin
      Delete(tmp, i+2, length(tmp)-i);
      tmp := tmp + paramB;
      ch_count := count;
      super_tmp := tmp;
      break;
      end;
   end;
end;
closefile(f);
// Если строка была заменена, то запись массива в файл:
if super_tmp <> '' then
   begin
   base[ch_count-1] := super_tmp;
   rewrite(f);
   for i := 0 to (Length(base)-1) do
      writeLn(f, base[i]);
   closefile(f);
   end;
end;

end.