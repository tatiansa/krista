unit Unit1; 

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs,
  StdCtrls, EditBtn, DbCtrls, ComCtrls, IniPropStorage, uib,
  uibdataset, db, BufDataset, dbf, zipper, IniFiles, LConvEncoding;

type

  { TForm1 }

  TForm1 = class(TForm)
      Button1: TButton;
      ComboBox1: TComboBox;
      Datasource1: TDatasource;
      Datasource2: TDatasource;
      Datasource3: TDatasource;
      Datasource4: TDatasource;
      Datasource5: TDatasource;
      DateEdit1: TDateEdit;
      DateEdit2: TDateEdit;
      Dbf1: TDbf;
      Dbf2: TDbf;
      Dbf3: TDbf;
      DirectoryEdit1: TDirectoryEdit;
      Edit1: TEdit;
      Edit2: TEdit;
      Edit3: TEdit;
      FileNameEdit1: TFileNameEdit;
      IniPropStorage1: TIniPropStorage;
      Label1: TLabel;
      Label2: TLabel;
      Label3: TLabel;
      Label4: TLabel;
      Label5: TLabel;
      Label6: TLabel;
      Label7: TLabel;
      ProgressBar1: TProgressBar;
      StatusBar1: TStatusBar;
    UIBDataBase1: TUIBDataBase;
    UIBDataSet1: TUIBDataSet;
    UIBDataSet2: TUIBDataSet;
    UIBTransaction1: TUIBTransaction;
    UIBTransaction2: TUIBTransaction;
    procedure Button1Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject);
    procedure UnloadPLP(d1,d2:String;savepath:AnsiString;config:TIniFile);
    procedure UnloadPBS(d1,d2:String;savepath:AnsiString;config:TIniFile);
    procedure UnloadAGR(d1,d2:String;savepath:AnsiString;config:TIniFile);
    procedure UnloadBND(d1,d2:String;savepath:AnsiString;config:TIniFile);
    procedure WriteFkr(str:String);
    procedure CollectOrg(where:String);

  private
    { private declarations }
  public
    { public declarations }
  end; 

var
  Form1: TForm1;

implementation

{$R *.lfm}

{ TForm1 }
// конвертация строки из windows-1251 в oem-866
function AnsiToConsole(str:string):string;
begin
     result:=ConvertEncoding(str, 'cp1251', 'cp866');
end;

// упаковка в zip архив
procedure PackToZip(savepath,pref:AnsiString; files: array of string);
  var zipfile:string;
  var zip: TZipper;
  var Zfiles: TZipFileEntries;
  var i,len: Integer;
begin
     zipfile:=savePath+pref+'_'+FormatDateTime('yyyymmdd',Date())+'.zip';
     if FileExists(zipfile)=True then
        DeleteFile(zipfile);

     Zfiles := TZipFileEntries.Create(TZipFileEntry);
     len:=Length(files)-1;
     for i:=0 to len do
     begin
          Zfiles.AddFileEntry(savePath+files[i], files[i]);
     end;

     zip := TZipper.Create;
     zip.FileName:=zipfile;
     zip.ZipFiles(Zfiles);

     zip.Free;
     Zfiles.Free;

     for i:=0 to len do
     begin
          DeleteFile(savePath+files[i]);
     end;
end;

// обработчик кнопки "выгрузить"
procedure TForm1.Button1Click(Sender: TObject);
  var d1,d2:string;
  var savepath:String;
  var date1,date2:TDateTime;
  var err,unloadindex:Integer;
  var config:TIniFile;
begin
  try
     err:=1;

     unloadindex:=ComboBox1.ItemIndex;

     date1:=DateEdit1.Date;
     date2:=DateEdit2.Date;
     d1:=FormatDateTime('YYYYMMDD',date1);
     d2:=FormatDateTime('YYYYMMDD',date2);

     UIBDataBase1.Connected:=False;
     UIBDataBase1.DatabaseName:=Utf8ToAnsi(FileNameEdit1.FileName);
     UIBDataBase1.UserName:=Edit1.Text;
     UIBDataBase1.PassWord:=Edit2.Text;
     err:=2;
     UIBDataBase1.Connected:=True;
     err:=1;

     savePath:=DirectoryEdit1.Directory;
     if not DirectoryExists(savePath) then
        begin
          if not ForceDirectories(savepath) then
             begin
               ShowMessage('Не удалось создать путь для выгрузки.');
               Exit;
             end;
        end;
     if not(RightStr(savePath,1)='\') then savePath+='\';
     savepath:=Utf8ToAnsi(savepath);

     err:=3;
     if not FileExists('urmload.cfg') then
        begin
          ShowMessage('Отсутствует файл конфигурации');
          Exit;
		end;

     config := TIniFile.Create('urmload.cfg');

     StatusBar1.Panels.Items[0].Text:='Обработка данных';
     StatusBar1.Refresh;
     err:=0;
     case unloadindex of
       0:UnloadPLP(d1,d2,savepath,config);
       1:UnloadPBS(d1,d2,savepath,config);
       2:UnloadAGR(d1,d2,savepath,config);
       3:UnloadBND(d1,d2,savepath,config);
     end;

     UIBDataSet1.Close;
     StatusBar1.Panels.Items[0].Text:='Выполнено';
  except
     StatusBar1.Panels.Items[0].Text:='Ошибка';
     case err of
       1:ShowMessage('Произошла неопознанная ошибка, проверьте параметры');
       2:ShowMessage('Не удалось подключиться к базе данных. Проверьте логин, пароль, путь к базе');
       3:ShowMessage('Не удалось прочитать файл конфигурации');
     end;
  end;
end;

// инициализация начального состояния
procedure TForm1.FormActivate(Sender: TObject);
  var strdate:string;
begin
     strdate:=IniPropStorage1.StoredValue['Date1'];
     if not(strdate=EmptyStr) then DateEdit1.Date:=StrToDate(strdate);
     strdate:=IniPropStorage1.StoredValue['Date2'];
     if not(strdate=EmptyStr) then DateEdit2.Date:=StrToDate(strdate);
end;

// сохранение состояния
procedure TForm1.FormClose(Sender: TObject);
begin
     IniPropStorage1.StoredValue['Date1']:=DateToStr(DateEdit1.Date);
     IniPropStorage1.StoredValue['Date2']:=DateToStr(DateEdit2.Date);
end;

// выгрузка платежных поручений
procedure TForm1.UnloadPLP(d1,d2,savepath:String;config:TIniFile);
  var org_idx, fkr_idx, idx, fkrid, ls, org_select: string;
  var inselect, outselect: String;
  var err, perc, s, mode, s_lenth: Integer;
  var dbffiles, select: array of string;
begin
  try
     err:=1;
     org_idx := '/';
     fkr_idx := '/';
     ls:=Edit3.Text;
     inselect := config.ReadString('plp','incoming',EmptyStr);
     outselect := config.ReadString('plp','outgoing',EmptyStr);
     mode := config.ReadInteger('main','mode',1);

     if mode = 1
     then
        begin
          SetLength(select,1);
          select[0]:=
	     'select '+
                     'a.id, a.docnumber, a.documentdate, a.paydate, '+
                     'a.acceptdate, a.note, a.nds, '+
                     'a.sourcefacialacc_cls as ent_ls, a.taxnote, '+
                     'b.sourcekfsr as divsn, kesr_.code as sourcekesr, '+
                     'b.sourcemeanstype as refbu, b.credit, '+
                     'c.org_ref as dest_org, c.bank_ref, c.acc as dest_rs, '+
                     'd.mfo as dest_mfo, d.cor as dest_cor, '+
                     'e.code as grbs, '+
                     'f.code as targt, '+
                     'g.code as tarst, '+
                     'i.inn as ent_inn, i.name as ent_name, '+
                     'i.shortname as ent_sname, i.inn20 as ent_kpp, '+
                     'h.acc as ent_rs, h.service_acc_ref, '+
                     'k.mfo as ent_mfo, k.cor as ent_cor '+
             'from '+
                   'facialfincaption a, facialfindetail b, org_accounts c, '+
                   'banks d, kvsr e, kcsr f, kvr g, facialacc_cls j, '+
                   'organizations i, org_accounts h, banks k, kesr kesr_ '+
             'where '+
                    'a.reject_cls is null and '+
                    'b.recordindex=a.id and '+
                    'c.id=a.destaccount and '+
                    'd.id=c.bank_ref and '+
                    'e.id=b.sourcekvsr and '+
                    'f.id=b.sourcekcsr and '+
                    'g.id=b.sourcekvr and '+
                    'h.id=a.sourceaccount and '+
                    'j.id=a.sourcefacialacc_cls and '+
                    'i.id=j.org_ref and '+
                    'k.id=h.bank_ref and '+
                    'kesr_.id=b.sourcekesr and '+
                    'a.acceptdate>='+d1+' and a.acceptdate<='+d2;
          if not(outselect=EmptyStr) then
             select[0] += ' and ' + outselect;

        end
     else
        begin
          SetLength(select,2);
          // исходящие платежи
          select[0]:=
             'select '+
                     'a.id, a.docnumber, a.documentdate, a.paydate, '+
                     'a.acceptdate, a.note, a.nds, a.credit, '+
                     'a.sourcefacialacc_cls as ent_ls, a.taxnote, '+
                     'b.sourcekfsr as divsn, kesr_.code as sourcekesr, '+
                     'b.sourcemeanstype as refbu, '+
                     'c.org_ref as dest_org, c.bank_ref, c.acc as dest_rs, '+
                     'd.mfo as dest_mfo, d.cor as dest_cor, '+
                     'e.code as grbs, '+
                     'f.code as targt, '+
                     'g.code as tarst, '+
                     'i.inn as ent_inn, i.name as ent_name, '+
                     'i.shortname as ent_sname, i.inn20 as ent_kpp, '+
                     'h.acc as ent_rs, h.service_acc_ref, '+
                     'k.mfo as ent_mfo, k.cor as ent_cor '+
             'from '+
                   'facialfincaption a, facialfindetail b, org_accounts c, '+
                   'banks d, kvsr e, kcsr f, kvr g, facialacc_cls j, '+
                   'organizations i, org_accounts h, banks k, kesr kesr_ '+
             'where '+
                    'a.reject_cls is null and '+
                    'b.recordindex=a.id and '+
                    'c.id=a.destaccount and '+
                    'd.id=c.bank_ref and '+
                    'e.id=b.sourcekvsr and '+
                    'f.id=b.sourcekcsr and '+
                    'g.id=b.sourcekvr and '+
                    'h.id=a.sourceaccount and '+
                    'j.id=a.sourcefacialacc_cls and '+
                    'i.id=j.org_ref and '+
                    'k.id=h.bank_ref and '+
                    'kesr_.id=b.sourcekesr and '+
                    'a.acceptdate>='+d1+' and a.acceptdate<='+d2;
          if not(outselect=EmptyStr) then
             select[0] += ' and ' + outselect;

          // входящие платежи
          select[1]:=
             'select '+
                     'a.id, a.docnumber, a.documentdate, a.paydate, '+
                     'a.acceptdate, a.note, a.nds, a.credit, '+
                     'a.destfacialacc_cls as ent_ls, a.taxnote, '+
                     'b.sourcekfsr as divsn, kesr_.code as sourcekesr, '+
                     'b.destmeanstype as refbu, '+
                     'c.org_ref as dest_org, c.bank_ref, c.acc as dest_rs, '+
                     'd.mfo as dest_mfo, d.cor as dest_cor, '+
                     'e.code as grbs, '+
                     'f.code as targt, '+
                     'g.code as tarst, '+
                     'i.inn as ent_inn, i.name as ent_name, '+
                     'i.shortname as ent_sname, i.inn20 as ent_kpp, '+
                     'h.acc as ent_rs, h.service_acc_ref, '+
                     'k.mfo as ent_mfo, k.cor as ent_cor '+
             'from '+
                   'facialfincaption a, facialfindetail b, org_accounts c, '+
                   'banks d, kvsr e, kcsr f, kvr g, facialacc_cls j, '+
                   'organizations i, org_accounts h, banks k, kesr kesr_ '+
             'where '+
                    'a.reject_cls is null and '+
                    'b.recordindex=a.id and '+
                    'c.id=a.sourceaccount and '+
                    'd.id=c.bank_ref and '+
                    'e.id=b.sourcekvsr and '+
                    'f.id=b.sourcekcsr and '+
                    'g.id=b.sourcekvr and '+
                    'h.id=a.destaccount and '+
                    'j.id=a.destfacialacc_cls and '+
                    'i.id=j.org_ref and '+
                    'k.id=h.bank_ref and '+
                    'kesr_.id=b.sourcekesr and '+
                    'a.acceptdate>='+d1+' and a.acceptdate<='+d2;
          if not(inselect=EmptyStr) then
             select[1] += ' and ' + inselect;
        end;

     s_lenth:=Length(select);

     if not(ls=EmptyStr) then
        for s:=0 to s_lenth-1 do
            select[s] += ' and a.destfacialacc_cls='+ls;

     err:=2;
     Dbf1.Close;
     Dbf1.TableName:=savePath+'plp_main.dbf';
     Dbf1.FieldDefs.clear;

     Dbf1.FieldDefs.Add('id'        ,ftString,15);
     Dbf1.FieldDefs.Add('ent_inn'   ,ftString,12);
     Dbf1.FieldDefs.Add('ent_kpp'   ,ftString,9);
     Dbf1.FieldDefs.Add('ent_sname' ,ftString,1024);
     Dbf1.FieldDefs.Add('ent_name'  ,ftString,1024);
     Dbf1.FieldDefs.Add('ent_ls'    ,ftString,25);
     Dbf1.FieldDefs.Add('ent_mfo'   ,ftString,9);
     Dbf1.FieldDefs.Add('ent_cor'   ,ftString,20);
     Dbf1.FieldDefs.Add('ent_rs'    ,ftString,20);
     Dbf1.FieldDefs.Add('docnumber' ,ftString,50);
     Dbf1.FieldDefs.Add('docdate'   ,ftString,8);
     Dbf1.FieldDefs.Add('credit'    ,ftFloat,15);
     Dbf1.FieldDefs[Dbf1.FieldDefs.Count-1].Precision := 2;
     Dbf1.FieldDefs.Add('paydate'   ,ftString,8);
     Dbf1.FieldDefs.Add('acceptdate',ftString,8);
     Dbf1.FieldDefs.Add('kazn_ls'   ,ftString,20);
     Dbf1.FieldDefs.Add('note'      ,ftString,1024);
     Dbf1.FieldDefs.Add('dest_org'  ,ftString,15);
     Dbf1.FieldDefs.Add('dest_rs'   ,ftString,20);
     Dbf1.FieldDefs.Add('dest_mfo'  ,ftString,9);
     Dbf1.FieldDefs.Add('dest_cor'  ,ftString,20);
     Dbf1.FieldDefs.Add('nds'       ,ftString,100);
     Dbf1.FieldDefs.Add('taxnote'   ,ftString,150);
     Dbf1.FieldDefs.Add('fkrid'     ,ftString,30);
     Dbf1.FieldDefs.Add('sourcekesr',ftString,10);
     Dbf1.FieldDefs.Add('refbu'     ,ftString,5);
     Dbf1.FieldDefs.Add('agrid'     ,ftstring,15);
     Dbf1.FieldDefs.Add('buhpaycls' ,ftFloat ,2);

     if FileExists(savePath+'plp_main.dbf')=True then
        DeleteFile(savePath+'plp_main.dbf');

     Dbf1.CreateTable;
     Dbf1.Exclusive:=True;
     Dbf1.Open;

     Dbf2.Close;
     Dbf2.TableName:=savePath+'plp_org.dbf';
     Dbf2.FieldDefs.Clear;

     Dbf2.FieldDefs.Add('id'       ,ftString,15);
     Dbf2.FieldDefs.Add('inn'      ,ftString,12);
     Dbf2.FieldDefs.Add('kpp'      ,ftString,9);
     Dbf2.FieldDefs.Add('shortname',ftString,1024);
     Dbf2.FieldDefs.Add('name'     ,ftString,1024);
     Dbf2.FieldDefs.Add('okato'    ,ftString,11);

     if FileExists(savePath+'plp_org.dbf')=True then
        DeleteFile(savePath+'plp_org.dbf');

     Dbf2.CreateTable;
     Dbf2.Exclusive:=True;
     Dbf2.Open;

     Dbf3.Close;
     Dbf3.TableName:=savePath+'plp_fkr.dbf';
     Dbf3.FieldDefs.Clear;

     Dbf3.FieldDefs.Add('id',ftString,30);
     Dbf3.FieldDefs.Add('grbs',ftString,3);
     Dbf3.FieldDefs.Add('divsn',ftString,4);
     Dbf3.FieldDefs.Add('targt',ftString,7);
     Dbf3.FieldDefs.Add('tarst',ftString,3);

     if FileExists(savePath+'plp_fkr.dbf')=True then
        DeleteFile(savePath+'plp_fkr.dbf');

     Dbf3.CreateTable;
     Dbf3.Exclusive:=True;
     Dbf3.Open;
     org_select:='';
     for s:=0 to s_lenth-1 do
     begin
          UIBDataSet1.SQL.Clear;
          UIBDataSet1.SQL.Add(select[s]);
          UIBDataSet1.Open;
          ProgressBar1.Min:=1;
          UIBDataSet1.Last;
          ProgressBar1.Max:=UIBDataSet1.RecordCount;
          perc:=1;
          err:=3;

          UIBDataSet1.First;
          while not UIBDataSet1.EOF do
          begin
               Dbf1.Insert;
               Dbf1.FieldByName('id').AsString:=UIBDataSet1.
                    FieldByName('id').AsString;
               Dbf1.FieldByName('ent_inn').AsString:=UIBDataSet1.
                    FieldByName('ent_inn').AsString;
               Dbf1.FieldByName('ent_kpp').AsString:=UIBDataSet1.
                    FieldByName('ent_kpp').AsString;
               Dbf1.FieldByName('ent_name').AsString:=AnsiToConsole(UIBDataSet1.
                    FieldByName('ent_name').AsString);
               Dbf1.FieldByName('ent_sname').AsString:=AnsiToConsole(UIBDataSet1.
                    FieldByName('ent_sname').AsString);
               Dbf1.FieldByName('ent_ls').AsString:=UIBDataSet1.
                    FieldByName('ent_ls').AsString;

               if UIBDataSet1.FieldByName('service_acc_ref').AsString=EmptyStr
               then
               begin
		            Dbf1.FieldByName('ent_mfo').AsString:=UIBDataSet1.
		                 FieldByName('ent_mfo').AsString;
		            Dbf1.FieldByName('ent_cor').AsString:=UIBDataSet1.
		                 FieldByName('ent_cor').AsString;
		            Dbf1.FieldByName('ent_rs').AsString:=UIBDataSet1.
		                 FieldByName('ent_rs').AsString;
               end
               else
               begin
                    UIBDataSet2.Close;
                    UIBDataSet2.SQL.Clear;
                    UIBDataSet2.SQL.Add('select a.acc, b.mfo, b.cor from org_accounts a, banks b where a.id=');
                    UIBDataSet2.SQL.Add(UIBDataSet1.FieldByName('service_acc_ref').AsString);
                    UIBDataSet2.SQL.Add(' and b.id=a.bank_ref');
                    UIBDataSet2.Open;
                    Dbf1.FieldByName('kazn_ls').AsString:=UIBDataSet1.
                         FieldByName('ent_rs').AsString;
                    Dbf1.FieldByName('ent_mfo').AsString:=UIBDataSet2.
                         FieldByName('mfo').AsString;
                    Dbf1.FieldByName('ent_cor').AsString:=UIBDataSet2.
                         FieldByName('cor').AsString;
                    Dbf1.FieldByName('ent_rs').AsString:=UIBDataSet2.
                         FieldByName('acc').AsString;
               end;

               Dbf1.FieldByName('docnumber').AsString:=AnsiToConsole(UIBDataSet1.
                    FieldByName('docnumber').AsString);
               Dbf1.FieldByName('docdate').AsString:=UIBDataSet1.
                    FieldByName('documentdate').AsString;
	           Dbf1.FieldByName('credit').AsFloat:=UIBDataSet1.
                    FieldByName('credit').AsFloat;

               if UIBDataSet1.FieldByName('paydate').AsString=EmptyStr then
                   Dbf1.FieldByName('paydate').AsString:=UIBDataSet1.
                        FieldByName('documentdate').AsString
               else
	               Dbf1.FieldByName('paydate').AsString:=UIBDataSet1.
                        FieldByName('paydate').AsString;

               Dbf1.FieldByName('acceptdate').AsString:=UIBDataSet1.
                    FieldByName('acceptdate').AsString;
               Dbf1.FieldByName('note').AsString:=AnsiToConsole(UIBDataSet1.
                    FieldByName('note').AsString);
               Dbf1.FieldByName('dest_org').AsString:=UIBDataSet1.
                    FieldByName('dest_org').AsString;
               Dbf1.FieldByName('dest_rs').AsString:=AnsiToConsole(UIBDataSet1.
                    FieldByName('dest_rs').AsString);
               Dbf1.FieldByName('dest_mfo').AsString:=UIBDataSet1.
                    FieldByName('dest_mfo').AsString;
               Dbf1.FieldByName('dest_cor').AsString:=UIBDataSet1.
                    FieldByName('dest_cor').AsString;
               Dbf1.FieldByName('nds').AsString:=AnsiToConsole(UIBDataSet1.
                    FieldByName('nds').AsString);
               Dbf1.FieldByName('taxnote').AsString:=AnsiToConsole(UIBDataSet1.
                    FieldByName('taxnote').AsString);
               fkrid:=RightStr('000' +Trim(UIBDataSet1.FieldByName('grbs' ).AsString),3)+'.';
               fkrid+=RightStr('0000'+Trim(UIBDataSet1.FieldByName('divsn').AsString),4)+'.';
               fkrid+=RightStr('0000000'+Trim(UIBDataSet1.FieldByName('targt').AsString),7)+'.';
               fkrid+=RightStr('000' +Trim(UIBDataSet1.FieldByName('tarst').AsString),3);
               Dbf1.FieldByName('fkrid').AsString:=fkrid;
               Dbf1.FieldByName('sourcekesr').AsString:=UIBDataSet1.
                    FieldByName('sourcekesr').AsString;
               Dbf1.FieldByName('refbu').AsString:=UIBDataSet1.
                    FieldByName('refbu').AsString;

               UIBDataSet2.close;
               UIBDataSet2.SQL.clear;
               UIBDataSet2.SQL.Add('select ags.recordindex ');
               UIBDataSet2.SQL.Add('from facialfindetail ffd, paymentschedule ps, agreementsteps ags ');
               UIBDataSet2.SQL.Add('where');
               UIBDataSet2.SQL.Add(  'ffd.recordindex='+UIBDataSet1.FieldByName('id').AsString+'and ');
               UIBDataSet2.SQL.Add(  'ffd.sourcepromise=ps.anumber and ');
               UIBDataSet2.SQL.Add(  'ps.recordindex=ags.id');
               UIBDataSet2.Open;
               Dbf1.FieldByName('agrid').AsString:=UIBDataSet2.FieldByName('recordindex').AsString;

               if s=0 then
                   Dbf1.FieldByName('buhpaycls').AsFloat:=1
               else
                   Dbf1.FieldByName('buhpaycls').AsFloat:=0;

               Dbf1.Post;

               if pos('/'+fkrid+'/',fkr_idx)=0 then
               begin
                    fkr_idx+=fkrid+'/';
                    WriteFkr(fkrid);
               end;

               idx := UIBDataSet1.FieldByName('dest_org').AsString;
               if pos('/'+idx+'/',org_idx)=0 then
               begin
                    org_idx+=idx+'/';
                    org_select+=idx+','
               end;

               UIBDataSet1.Next;
               perc+=1;
               ProgressBar1.Position:=perc;
               ProgressBar1.Refresh;
          end;
     end;
     if org_select <> EmptyStr then
     begin
          org_select := LeftStr(org_select, Length(org_select)-1);
          org_select := 'id in ('+org_select+')';
          CollectOrg(org_select);
     end;
     dbf1.Close;
     dbf2.Close;
     dbf3.Close;
     UIBDataSet2.Close;

     err:=4;
     SetLength(dbffiles,3);
     dbffiles[0] := 'plp_main.dbf';
     dbffiles[1] := 'plp_org.dbf';
     dbffiles[2] := 'plp_fkr.dbf';
     PackToZip(savepath, 'plp', dbffiles);

     err:=0;
  except
     StatusBar1.Panels.Items[0].Text:='Ошибка';
     case err of
       1:ShowMessage('Произошла неопознанная ошибка, проверьте параметры');
       2:ShowMessage('Ошибка создания файлов выгрузки');
       3:ShowMessage('Ошибка заполнения файлов выгрузки');
       4:ShowMessage('Ошибка создания архива');
     end;
  end;
end;

// сметные назначения
procedure TForm1.UnloadPBS(d1,d2,savepath:String;config:TIniFile);
  var err,perc:Integer;
  var select,str,fkr_idx,ls,dopselect:string;
  var dbffiles: array of string;
begin
  try
     err:=1;
     dopselect:=config.ReadString('pbs','config',EmptyStr);
     select:='select '+
                 'a.id, a.dat, a.anumber, a.docdat, a.docnumber, a.note, '+
                 'a.facialacccls as ent_ls, a.orgref as rasp_id, '+
                 'b.kfsr as divsn, b.kesr as kosgu, b.summayear1 as summ, '+
                 'b.org_ref as ent_id, b.meanstype as meanstype, '+
                 'c.name as rasp_name, '+
                 'g.name as ent_name, g.inn as ent_inn, '+
                 'g.shortname as ent_sname, g.inn20 as ent_kpp, '+
                 'd.code as grbs, '+
                 'e.code as targt, '+
                 'f.code as tarst '+
             'from '+
                 'budnotify a, '+
                 'budgetdata b, '+
                 'organizations c, '+
                 'kvsr d, '+
                 'kcsr e, '+
                 'kvr f, '+
                 'organizations g '+
             'where '+
                 'a.rejectnote is null and '+
                 'a.rejectcls is null and '+
                 'a.id=b.recordindex and '+
                 'd.id=b.kvsr and '+
                 'e.id=b.kcsr and '+
                 'f.id=b.kvr and '+
                 'a.orgref=c.id and '+
                 'b.org_ref=g.id and '+
                 'a.dat>='+d1+' and '+
                 'a.dat<='+d2;
     if not(dopselect = EmptyStr) then
        select += ' and ' + dopselect;

     ls:=Edit3.Text;
     if not(ls=EmptyStr) then select+=' and b.facialacc_cls='+ls;

     UIBDataSet1.SQL.Clear;
     UIBDataSet1.SQL.Add(select);
     UIBDataSet1.Open;

     err:=2;
     Dbf1.Close;
     Dbf1.TableName:=savePath+'pbs_main.dbf';
     Dbf1.FieldDefs.clear;

     Dbf1.FieldDefs.Add('id'       ,ftString,15 );
     Dbf1.FieldDefs.Add('ent_inn'  ,ftString,12 );
     Dbf1.FieldDefs.Add('ent_kpp'  ,ftString,9);
     Dbf1.FieldDefs.Add('ent_sname',ftString,1024);
     Dbf1.FieldDefs.Add('ent_name' ,ftString,1024);
     Dbf1.FieldDefs.Add('ent_ls'   ,ftString,25 );
     Dbf1.FieldDefs.Add('dat'      ,ftString,8  );
     Dbf1.FieldDefs.Add('anumber'  ,ftString,20 );
     Dbf1.FieldDefs.Add('docdat'   ,ftString,8  );
     Dbf1.FieldDefs.Add('docnumber',ftString,50 );
     Dbf1.FieldDefs.Add('note'     ,ftString,255);
     Dbf1.FieldDefs.Add('rasp_name',ftString,255);
     Dbf1.FieldDefs.Add('fkrid'   ,ftString,30 );
     Dbf1.FieldDefs.Add('kosgu'    ,ftString,10 );
     Dbf1.FieldDefs.Add('summ'     ,ftFloat ,15 );
     Dbf1.FieldDefs[Dbf1.FieldDefs.Count-1].Precision := 2;
     Dbf1.FieldDefs.Add('kvd'     ,ftFloat ,15 );

     if FileExists(savePath+'pbs_main.dbf')=True then
        DeleteFile(savePath+'pbs_main.dbf');

     Dbf1.CreateTable;
     Dbf1.Exclusive:=True;
     Dbf1.Open;

     Dbf3.Close;
     Dbf3.TableName:=savePath+'pbs_fkr.dbf';
     Dbf3.FieldDefs.Clear;

     Dbf3.FieldDefs.Add('id',ftString,30);
     Dbf3.FieldDefs.Add('grbs',ftString,3);
     Dbf3.FieldDefs.Add('divsn',ftString,4);
     Dbf3.FieldDefs.Add('targt',ftString,7);
     Dbf3.FieldDefs.Add('tarst',ftString,3);

     if FileExists(savePath+'pbs_fkr.dbf')=True then
        DeleteFile(savePath+'pbs_fkr.dbf');

     Dbf3.CreateTable;
     Dbf3.Exclusive:=True;
     Dbf3.Open;

     ProgressBar1.Min:=1;
     UIBDataSet1.Last;
     ProgressBar1.Max:=UIBDataSet1.RecordCount;
     perc:=1;
     err:=3;
     UIBDataSet1.First;
     select:=EmptyStr;
     fkr_idx:='/';

     while not UIBDataSet1.EOF do
     begin
       Dbf1.Insert;
       Dbf1.FieldByName('id').AsString:=UIBDataSet1.FieldByName('id').AsString;
       Dbf1.FieldByName('ent_inn').AsString:=UIBDataSet1.FieldByName('ent_inn').AsString;
       Dbf1.FieldByName('ent_kpp').AsString:=UIBDataSet1.FieldByName('ent_kpp').AsString;
       Dbf1.FieldByName('ent_sname').AsString:=AnsiToConsole(UIBDataSet1.FieldByName('ent_sname').AsString);
       Dbf1.FieldByName('ent_name').AsString:=AnsiToConsole(UIBDataSet1.FieldByName('ent_name').AsString);
       Dbf1.FieldByName('ent_ls').AsString:=UIBDataSet1.FieldByName('ent_ls').AsString;
       Dbf1.FieldByName('dat').AsString:=UIBDataSet1.FieldByName('dat').AsString;
       Dbf1.FieldByName('anumber').AsString:=AnsiToConsole(UIBDataSet1.FieldByName('anumber').AsString);
       Dbf1.FieldByName('docdat').AsString:=UIBDataSet1.FieldByName('docdat').AsString;
       Dbf1.FieldByName('docnumber').AsString:=AnsiToConsole(UIBDataSet1.FieldByName('docnumber').AsString);
       Dbf1.FieldByName('note').AsString:=AnsiToConsole(UIBDataSet1.FieldByName('note').AsString);
       Dbf1.FieldByName('rasp_name').AsString:=AnsiToConsole(UIBDataSet1.FieldByName('rasp_name').AsString);
       str:=RightStr('000' +Trim(UIBDataSet1.FieldByName('grbs' ).AsString),3)+'.';
       str+=RightStr('0000'+Trim(UIBDataSet1.FieldByName('divsn').AsString),4)+'.';
       str+=RightStr('0000000'+Trim(UIBDataSet1.FieldByName('targt').AsString),7)+'.';
       str+=RightStr('000' +Trim(UIBDataSet1.FieldByName('tarst').AsString),3);
       Dbf1.FieldByName('fkrid').AsString:=str;
       Dbf1.FieldByName('kosgu').AsString:=UIBDataSet1.FieldByName('kosgu').AsString;
       Dbf1.FieldByName('summ').AsFloat:=UIBDataSet1.FieldByName('summ').AsFloat;
       Dbf1.FieldByName('kvd').AsFloat:=UIBDataSet1.FieldByName('meanstype').AsFloat;
       Dbf1.Post;

       if pos('/'+str+'/',fkr_idx)=0 then
       begin
         fkr_idx+=str+'/';
         WriteFkr(str)
       end;

       UIBDataSet1.Next;
       perc+=1;
       ProgressBar1.Position:=perc;
       ProgressBar1.Refresh;
     end;

     dbf1.Close;
     dbf2.Close;
     dbf3.Close;

     err:=4;
     SetLength(dbffiles,2);
     dbffiles[0] := 'pbs_main.dbf';
     dbffiles[1] := 'pbs_fkr.dbf';
     PackToZip(savepath, 'pbs', dbffiles);

     err:=0;

  except
     StatusBar1.Panels.Items[0].Text:='Ошибка';
     case err of
       1:ShowMessage('Произошла неопознанная ошибка, проверьте параметры');
       2:ShowMessage('Ошибка создания файлов выгрузки');
       3:ShowMessage('Ошибка заполнения базы документов или базы кбк');
       4:ShowMessage('Ошибка создания архива');
     end;
  end;
end;

// реестр обязательств
procedure TForm1.UnloadAGR(d1,d2,savepath:String; config:TIniFile);
  var err,perc,i,mode:integer;
  var select,fkr_idx,org_idx,str,ls,part,dopselect,addselect:string;
  var aSelect,dbffiles:array of string;
  var Dbf4: TDbf;
begin
  try
     err:=1;
     addselect:=config.ReadString('agr','config',EmptyStr);
     mode := config.ReadInteger('main','mode',1);
     SetLength(aSelect,2);
     aSelect[0]:='select '+
                     'a.id, a.agreementtype, a.docnumber, a.agreementdate, '+
                     'a.agreementbegindate, a.agreementenddate, a.executer_ref, '+
                     'a.purportdoc, a.progindex, a.adjustmentdocnumber, '+
                     'a.reestrnumber, a.agreementsumma, '+
                     'b.acceptdate, b.kfsr as divsn, kesr_.code as kosgu, '+
                     'b.month01, b.month02, b.month03, b.month04, b.month05, '+
                     'b.month06, b.month07, b.month08, b.month09, b.month10, '+
                     'b.month11, b.month12, b.parentnumber as parid, '+
                     'b.meanstype, b.summa, '+
                     'd.code as grbs, '+
                     'e.code as targt, '+
                     'f.code as tarst, '+
                     'g.inn as ent_inn, g.name as ent_name, '+
                     'g.shortname as ent_sname, g.inn20 as ent_kpp, '+
                     'i.acc as ex_rs, '+
                     'h.mfo as ex_mfo, h.cor as ex_cor '+
                 'from '+
                     'agreements a, paymentschedule b, kvsr d, kcsr e, kvr f, '+
                     'organizations g, banks h, org_accounts i, kesr kesr_ '+
                 'where '+
                     'a.rejectcause is null and a.rejectcls is null and '+
                     'a.id=b.agreementref and b.kvsr=d.id and b.kcsr=e.id and '+
                     'b.kvr=f.id and a.client_ref=g.id and '+
                     'a.executeraccref=i.id and i.bank_ref=h.id and '+
                     'kesr_.id = b.kesr and ' +
                     'a.acceptdate>='+d1+' and a.acceptdate<='+d2;

     aSelect[1]:='select '+
                     'a.id, a.agreementtype, a.docnumber, a.agreementdate, '+
                     'a.agreementbegindate, a.agreementenddate, a.executer_ref, '+
                     'a.purportdoc, a.progindex, a.adjustmentdocnumber, '+
                     'a.reestrnumber, a.agreementsumma, '+
                     'b.acceptdate, b.kfsr as divsn, kesr_.code as kosgu, '+
                     'b.month01, b.month02, b.month03, b.month04, b.month05, '+
                     'b.month06, b.month07, b.month08, b.month09, b.month10, '+
                     'b.month11, b.month12, b.parentnumber as parid, '+
                     'b.meanstype, b.summa, '+
                     'd.code as grbs, '+
                     'e.code as targt, '+
                     'f.code as tarst, '+
                     'g.inn as ent_inn, g.name as ent_name, '+
                     'g.shortname as ent_sname, g.inn20 as ent_kpp '+
                 'from '+
                     'agreements a, paymentschedule b, kvsr d, kcsr e, kvr f, '+
                     'organizations g, kesr kesr_ '+
                 'where '+
                     'a.rejectcause is null and a.rejectcls is null and '+
                     'a.id=b.agreementref and b.kvsr=d.id and b.kcsr=e.id and '+
                     'b.kvr=f.id and a.client_ref=g.id and '+
                     'a.executeraccref is null and '+
                     'kesr_.id = b.kesr and ' +
                     'a.acceptdate>='+d1+' and a.acceptdate<='+d2;

     ls:=Edit3.Text;
     if not(ls=EmptyStr) then
     begin
        aSelect[0]+=' and a.facialacc_cls='+ls;
        aSelect[1]+=' and a.facialacc_cls='+ls;
	 end;


     dopselect:='select a.id as agr_id, '+
                    'e.id as est_id, e.amount, e.summa, '+
                    't.name as tdo_name, '+
                    'case when okdp.id is null then 0 else okdp.sourcecode end as sourcecode, '+
                    't.okpd2 as okpd2_code, '+
                    'm.id as msm_id, m.name as msm_name, m.shortname as msm_shortname '+
                'from '+
                    'estimate e '+
                    'inner join agreements as a on a.id=e.recordindex '+
                    'inner join tenderobjects as t on t.id=e.productcls '+
                    'inner join measurementcls as m on m.id=t.measurementcls '+
                    'left outer join okdp on okdp.id=t.okdp '+
                'where '+
                    'a.rejectcause is null and '+
                    'a.rejectcls is null and '+
                    'a.acceptdate>='+d1+' and a.acceptdate<='+d2;

     if not(addselect = EmptyStr) then
     begin
       aSelect[0] += ' and ' + addselect;
       aSelect[1] += ' and ' + addselect;
       dopselect  += ' and ' + addselect;
	 end;

     err:=2;
     Dbf1.Close;
     Dbf1.TableName:=savePath+'agr_main.dbf';
     Dbf1.FieldDefs.clear;

     Dbf1.FieldDefs.Add('id'        ,ftString,15);
     Dbf1.FieldDefs.Add('parid'     ,ftString,15);
     Dbf1.FieldDefs.Add('ent_inn'   ,ftString,12);
     Dbf1.FieldDefs.Add('ent_kpp'   ,ftString,9);
     Dbf1.FieldDefs.Add('ent_sname' ,ftString,1024);
     Dbf1.FieldDefs.Add('ent_name'  ,ftString,1024);
     Dbf1.FieldDefs.Add('agrtype'   ,ftString,1);
     Dbf1.FieldDefs.Add('docnumber' ,ftString,50);
     Dbf1.FieldDefs.Add('agrdate'   ,ftString,8);
     Dbf1.FieldDefs.Add('agrbegdate',ftString,8);
     Dbf1.FieldDefs.Add('agrenddate',ftString,8);
     Dbf1.FieldDefs.Add('adjdocnum' ,ftString,50);
     Dbf1.FieldDefs.Add('reestrnum' ,ftString,20);
     Dbf1.FieldDefs.Add('executer'  ,ftString,15);
     Dbf1.FieldDefs.Add('ex_rs'     ,ftString,20);
     Dbf1.FieldDefs.Add('ex_mfo'    ,ftString,9);
     Dbf1.FieldDefs.Add('ex_cor'    ,ftString,20);
     Dbf1.FieldDefs.Add('fkr'       ,ftString,30);
     Dbf1.FieldDefs.Add('acceptdate',ftString,8);
     Dbf1.FieldDefs.Add('kosgu'     ,ftString,8);
     Dbf1.FieldDefs.Add('month01'   ,ftFloat,15);
     Dbf1.FieldDefs[Dbf1.FieldDefs.Count-1].Precision := 2;
     Dbf1.FieldDefs.Add('month02'   ,ftFloat,15);
     Dbf1.FieldDefs[Dbf1.FieldDefs.Count-1].Precision := 2;
     Dbf1.FieldDefs.Add('month03'   ,ftFloat,15);
     Dbf1.FieldDefs[Dbf1.FieldDefs.Count-1].Precision := 2;
     Dbf1.FieldDefs.Add('month04'   ,ftFloat,15);
     Dbf1.FieldDefs[Dbf1.FieldDefs.Count-1].Precision := 2;
     Dbf1.FieldDefs.Add('month05'   ,ftFloat,15);
     Dbf1.FieldDefs[Dbf1.FieldDefs.Count-1].Precision := 2;
     Dbf1.FieldDefs.Add('month06'   ,ftFloat,15);
     Dbf1.FieldDefs[Dbf1.FieldDefs.Count-1].Precision := 2;
     Dbf1.FieldDefs.Add('month07'   ,ftFloat,15);
     Dbf1.FieldDefs[Dbf1.FieldDefs.Count-1].Precision := 2;
     Dbf1.FieldDefs.Add('month08'   ,ftFloat,15);
     Dbf1.FieldDefs[Dbf1.FieldDefs.Count-1].Precision := 2;
     Dbf1.FieldDefs.Add('month09'   ,ftFloat,15);
     Dbf1.FieldDefs[Dbf1.FieldDefs.Count-1].Precision := 2;
     Dbf1.FieldDefs.Add('month10'   ,ftFloat,15);
     Dbf1.FieldDefs[Dbf1.FieldDefs.Count-1].Precision := 2;
     Dbf1.FieldDefs.Add('month11'   ,ftFloat,15);
     Dbf1.FieldDefs[Dbf1.FieldDefs.Count-1].Precision := 2;
     Dbf1.FieldDefs.Add('month12'   ,ftFloat,15);
     Dbf1.FieldDefs[Dbf1.FieldDefs.Count-1].Precision := 2;
     Dbf1.FieldDefs.Add('summ'  ,ftFloat,15);
     Dbf1.FieldDefs[Dbf1.FieldDefs.Count-1].Precision := 2;
     Dbf1.FieldDefs.Add('refbu'     ,ftString,5);
     Dbf1.FieldDefs.Add('purportdoc',ftString,255);
     Dbf1.FieldDefs.Add('docindex'  ,ftString,3);

     if FileExists(savePath+'agr_main.dbf')=True then
        DeleteFile(savePath+'agr_main.dbf');

     Dbf1.CreateTable;
     Dbf1.Exclusive:=True;
     Dbf1.Open;

     Dbf2.Close;
     Dbf2.TableName:=savePath+'agr_org.dbf';
     Dbf2.FieldDefs.Clear;

     Dbf2.FieldDefs.Add('id'       ,ftString,15);
     Dbf2.FieldDefs.Add('inn'      ,ftString,12);
     Dbf2.FieldDefs.Add('kpp'      ,ftString,9);
     Dbf2.FieldDefs.Add('name'     ,ftString,1024);
     Dbf2.FieldDefs.Add('shortname',ftString,1024);
     Dbf2.FieldDefs.Add('okato'    ,ftString,11);

     if FileExists(savePath+'agr_org.dbf')=True then
        DeleteFile(savePath+'agr_org.dbf');

     Dbf2.CreateTable;
     Dbf2.Exclusive:=True;
     Dbf2.Open;

     Dbf3.Close;
     Dbf3.TableName:=savePath+'agr_fkr.dbf';
     Dbf3.FieldDefs.Clear;

     Dbf3.FieldDefs.Add('id'   ,ftString,30);
     Dbf3.FieldDefs.Add('grbs' ,ftString,3);
     Dbf3.FieldDefs.Add('divsn',ftString,4);
     Dbf3.FieldDefs.Add('targt',ftString,7);
     Dbf3.FieldDefs.Add('tarst',ftString,3);

     if FileExists(savePath+'agr_fkr.dbf')=True then
        DeleteFile(savePath+'agr_fkr.dbf');

     Dbf3.CreateTable;
     Dbf3.Exclusive:=True;
     Dbf3.Open;

     Dbf4:=TDbf.Create(nil);
     Dbf4.Close;
     Dbf4.TableName:=savePath+'agr_est.dbf';
     Dbf4.FieldDefs.Clear;

     Dbf4.FieldDefs.Add('id',ftString,20);
     Dbf4.FieldDefs.Add('recordidx',ftString,20);
     Dbf4.FieldDefs.Add('amount',ftFloat,17);
     Dbf4.FieldDefs.Items[2].Precision:=4;
     Dbf4.FieldDefs.Add('summa',ftFloat,15);
     Dbf4.FieldDefs.Items[3].Precision:=2;
     Dbf4.FieldDefs.Add('name',ftWideString,1024);
     Dbf4.FieldDefs.Add('okdp_code',ftString,20);
     Dbf4.FieldDefs.Add('okpd2_code',ftString,20);
     Dbf4.FieldDefs.Add('msm_id',ftString,20);
     Dbf4.FieldDefs.Add('msm_name',ftWideString,1024);
     Dbf4.FieldDefs.Add('msm_shortn',ftString,50);

     if FileExists(savePath+'agr_est.dbf')=True then
        DeleteFile(savePath+'agr_est.dbf');

     Dbf4.CreateTable;
     Dbf4.Exclusive:=True;
     Dbf4.Open;

     fkr_idx:='/';
     org_idx:='/';
     select:=EmptyStr;

     for i:=0 to 1 do
     begin
          UIBDataSet1.SQL.Clear;
          UIBDataSet1.SQL.Add(aSelect[i]);
          UIBDataSet1.Open;
          ProgressBar1.Min:=1;
          UIBDataSet1.Last;
          ProgressBar1.Max:=UIBDataSet1.RecordCount;
          perc:=1;
          err:=3;
          UIBDataSet1.First;
          while not UIBDataSet1.EOF do
          begin
               Dbf1.Insert;
               Dbf1.FieldByName('id').AsString:=UIBDataSet1.FieldByName('id').
                    AsString;
               if not(UIBDataSet1.FieldByName('parid').AsString=EmptyStr) then
               begin
                    UIBDataSet2.Close;
                    UIBDataSet2.SQL.Clear;
                    UIBDataSet2.SQL.Add(
                        'select agreementref from paymentschedule where anumber='
                    );
                    UIBDataSet2.SQL.Add(UIBDataSet1.FieldByName('parid').AsString);
                    UIBDataSet2.Open;
                    Dbf1.FieldByName('parid').AsString:=UIBDataSet2.
                        FieldByName('agreementref').AsString;
               end;
               Dbf1.FieldByName('ent_inn').AsString:=UIBDataSet1.
                    FieldByName('ent_inn').AsString;
               Dbf1.FieldByName('ent_kpp').AsString:=UIBDataSet1.
                    FieldByName('ent_kpp').AsString;
               Dbf1.FieldByName('ent_sname').AsString:=AnsiToConsole(UIBDataSet1.
                    FieldByName('ent_sname').AsString);
               Dbf1.FieldByName('ent_name').AsString:=AnsiToConsole(UIBDataSet1.
                    FieldByName('ent_name').AsString);
               Dbf1.FieldByName('agrtype').AsString:=UIBDataSet1.
                    FieldByName('agreementtype').AsString;
               Dbf1.FieldByName('docnumber').AsString:=AnsiToConsole(UIBDataSet1.
                    FieldByName('docnumber').AsString);
               Dbf1.FieldByName('agrdate').AsString:=UIBDataSet1.
                    FieldByName('agreementdate').AsString;
               Dbf1.FieldByName('agrbegdate').AsString:=UIBDataSet1.
                    FieldByName('agreementbegindate').AsString;
               Dbf1.FieldByName('agrenddate').AsString:=UIBDataSet1.
                    FieldByName('agreementenddate').AsString;
               Dbf1.FieldByName('acceptdate').AsString:=UIBDataSet1.
                    FieldByName('acceptdate').AsString;
               Dbf1.FieldByName('adjdocnum').AsString:=UIBDataSet1.
                    FieldByName('adjustmentdocnumber').AsString;
               Dbf1.FieldByName('reestrnum').AsString:=UIBDataSet1.
                    FieldByName('reestrnumber').AsString;
               Dbf1.FieldByName('executer').AsString:=UIBDataSet1.
                    FieldByName('executer_ref').AsString;
               if i=0 then
               begin
                    Dbf1.FieldByName('ex_rs').AsString:=UIBDataSet1.
                        FieldByName('ex_rs').AsString;
                    Dbf1.FieldByName('ex_mfo').AsString:=UIBDataSet1.
                        FieldByName('ex_mfo').AsString;
                    Dbf1.FieldByName('ex_cor').AsString:=UIBDataSet1.
                        FieldByName('ex_cor').AsString;
               end;
               if mode = 1 then
                  Dbf1.FieldByName('summ').AsFloat:=UIBDataSet1.
                       FieldByName('summa').AsFloat
               else
                  Dbf1.FieldByName('summ').AsFloat:=UIBDataSet1.
                       FieldByName('agreementsumma').AsFloat;

               Dbf1.FieldByName('kosgu').AsString:=UIBDataSet1.
                    FieldByName('kosgu').AsString;
               Dbf1.FieldByName('month01').AsFloat:=UIBDataSet1.
                    FieldByName('month01').AsFloat;
               Dbf1.FieldByName('month02').AsFloat:=UIBDataSet1.
                    FieldByName('month02').AsFloat;
               Dbf1.FieldByName('month03').AsFloat:=UIBDataSet1.
                    FieldByName('month03').AsFloat;
               Dbf1.FieldByName('month04').AsFloat:=UIBDataSet1.
                    FieldByName('month04').AsFloat;
               Dbf1.FieldByName('month05').AsFloat:=UIBDataSet1.
                    FieldByName('month05').AsFloat;
               Dbf1.FieldByName('month06').AsFloat:=UIBDataSet1.
                    FieldByName('month06').AsFloat;
               Dbf1.FieldByName('month07').AsFloat:=UIBDataSet1.
                    FieldByName('month07').AsFloat;
               Dbf1.FieldByName('month08').AsFloat:=UIBDataSet1.
                    FieldByName('month08').AsFloat;
               Dbf1.FieldByName('month09').AsFloat:=UIBDataSet1.
                    FieldByName('month09').AsFloat;
               Dbf1.FieldByName('month10').AsFloat:=UIBDataSet1.
                    FieldByName('month10').AsFloat;
               Dbf1.FieldByName('month11').AsFloat:=UIBDataSet1.
                    FieldByName('month11').AsFloat;
               Dbf1.FieldByName('month12').AsFloat:=UIBDataSet1.
                    FieldByName('month12').AsFloat;
               Dbf1.FieldByName('purportdoc').AsString:=AnsiToConsole(UIBDataSet1.
                    FieldByName('purportdoc').AsString);
               Dbf1.FieldByName('docindex').AsString:=UIBDataSet1.
                    FieldByName('progindex').AsString;

               part:='000'+Trim(UIBDataSet1.FieldByName('grbs').AsString);
               str:=RightStr(part,3)+'.';
               part:='0000'+Trim(UIBDataSet1.FieldByName('divsn').AsString);
               str+=RightStr(part,4)+'.';
               part:='0000000'+Trim(UIBDataSet1.FieldByName('targt').AsString);
               str+=RightStr(part,7)+'.';
               part:='000'+Trim(UIBDataSet1.FieldByName('tarst').AsString);
               str+=RightStr(part,3);
               Dbf1.FieldByName('fkr').AsString:=str;

               Dbf1.FieldByName('refbu').AsString:=UIBDataSet1.
                    FieldByName('meanstype').AsString;
               Dbf1.Post;

               if pos('/'+str+'/',fkr_idx)=0 then
               begin
                    fkr_idx+=str+'/';
                    WriteFkr(str);
               end;

               str:=UIBDataSet1.FieldByName('executer_ref').AsString;
               if pos('/'+str+'/',org_idx)=0 then
               begin
                    org_idx+=str+'/';
                    if not(select=EmptyStr) then select+=' or ';
                    select+='id='+str
               end;

               UIBDataSet1.Next;
               perc+=1;
               ProgressBar1.Position:=perc;
               ProgressBar1.Refresh;
          end;
     end;

     err:=6;
     UIBDataSet1.SQL.Clear;
     UIBDataSet1.SQL.Add(dopselect);
     UIBDataSet1.Open;
     UIBDataSet1.Last;
     ProgressBar1.Max:=UIBDataSet1.RecordCount;
     perc:=1;
     UIBDataSet1.First;
     while not UIBDataSet1.EOF do
     begin
          Dbf4.Insert;
          Dbf4.FieldByName('id').AsString:=UIBDataSet1.FieldByName('est_id').AsString;
          Dbf4.FieldByName('recordidx').AsString:=UIBDataSet1.FieldByName('agr_id').AsString;
          Dbf4.FieldByName('amount').AsString:=UIBDataSet1.FieldByName('amount').AsString;
          Dbf4.FieldByName('summa').AsString:=UIBDataSet1.FieldByName('summa').AsString;
          Dbf4.FieldByName('name').AsString:=AnsiToConsole(UIBDataSet1.FieldByName('tdo_name').AsString);
          Dbf4.FieldByName('okdp_code').AsString:=UIBDataSet1.FieldByName('sourcecode').AsString;
          Dbf4.FieldByName('okpd2_code').AsString:=UIBDataSet1.FieldByName('okpd2_code').AsString;
          Dbf4.FieldByName('msm_id').AsString:=UIBDataSet1.FieldByName('msm_id').AsString;
          Dbf4.FieldByName('msm_name').AsString:=AnsiToConsole(UIBDataSet1.FieldByName('msm_name').AsString);
          Dbf4.FieldByName('msm_shortn').AsString:=AnsiToConsole(UIBDataSet1.FieldByName('msm_shortname').AsString);
          Dbf4.Post;
          UIBDataSet1.Next;
          perc+=1;
          ProgressBar1.Position:=perc;
          ProgressBar1.Refresh;
		 end;

     err:=4;
     CollectOrg(select);

     Dbf1.Close;
     Dbf2.Close;
     Dbf3.Close;
     Dbf4.Close;

     err:=5;
     SetLength(dbffiles,4);
     dbffiles[0] := 'agr_main.dbf';
     dbffiles[1] := 'agr_org.dbf';
     dbffiles[2] := 'agr_fkr.dbf';
     dbffiles[3] := 'agr_est.dbf';
     PackToZip(savepath, 'arg', dbffiles);

     err:=0;

  except
     StatusBar1.Panels.Items[0].Text:='Ошибка';
     case err of
       1:ShowMessage('Произошла неопознанная ошибка, проверьте параметры');
       2:ShowMessage('Ошибка создания файлов выгрузки');
       3:ShowMessage('Ошибка заполнения базы документов или базы кбк');
       4:ShowMessage('Ошибка заполнения базы организаций');
       5:ShowMessage('Ошибка создания архива');
       6:ShowMessage('Ошибка заполнения agr_est.dbf');
     end;
  end;
end;


// сохранение расшифровки кбк
procedure TForm1.WriteFkr(str:string);
begin
     Dbf3.Insert;
     Dbf3.FieldByName('id').AsString:=str;
     Dbf3.FieldByName('grbs').AsString:=UIBDataSet1.FieldByName('grbs').AsString;
     Dbf3.FieldByName('divsn').AsString:=UIBDataSet1.FieldByName('divsn').AsString;
     Dbf3.FieldByName('targt').AsString:=UIBDataSet1.FieldByName('targt').AsString;
     Dbf3.FieldByName('tarst').AsString:=UIBDataSet1.FieldByName('tarst').AsString;
     Dbf3.Post;
end;

// сохранение информации по контрагентам
procedure TForm1.CollectOrg(where:string);
  var select, inn_field_value:string;
  var inn_field_length: integer;
begin
     if where=EmptyStr then Exit;
     select:='select id, inn, inn20 as kpp, name, shortname, okato from organizations where '+where;
     UIBDataSet1.Close;
     UIBDataSet1.SQL.Clear;
     UIBDataSet1.SQL.Add(select);
     UIBDataSet1.Open;
     UIBDataSet1.First;
     while not UIBDataSet1.EOF do
     begin
        Dbf2.Insert;
        Dbf2.FieldByName('id').AsString:=UIBDataSet1.FieldByName('id').AsString;
        inn_field_value:=UIBDataSet1.FieldByName('inn').AsString;
        inn_field_length:=length(UIBDataSet1.FieldByName('inn').AsString);
        if (inn_field_length=9) or (inn_field_length = 11) then
        begin
                inn_field_value:= '0' + UIBDataSet1.FieldByName('inn').AsString;
        end;
        Dbf2.FieldByName('inn').AsString:=inn_field_value;
        Dbf2.FieldByName('kpp').AsString:=UIBDataSet1.FieldByName('kpp').AsString;
        Dbf2.FieldByName('name').AsString:=AnsiToConsole(UIBDataSet1.FieldByName('name').AsString);
        Dbf2.FieldByName('shortname').AsString:=AnsiToConsole(UIBDataSet1.FieldByName('shortname').AsString);
        Dbf2.FieldByName('okato').AsString:=AnsiToConsole(UIBDataSet1.FieldByName('okato').AsString);
        Dbf2.Post;
        UIBDataSet1.Next;
     end;
end;

// прочие финансовые документы
procedure TForm1.UnloadBND(d1,d2,savepath:String; config:TIniFile);
  var select,ls,org_idx,fkr_idx,str,addselect:string;
  var err,perc:integer;
  var dbffiles: array of string;
begin
  try
     err:=1;
     addselect:=config.ReadString('bnd','config',EmptyStr);
     select:='select '+
                 'a.acceptdate, '+
                 'b.id, b.docnum, b.docdate, b.org_ref as k_id, '+
                 'b.inn as k_inn, b.note, b.credit, b.debit, '+
                 'b.facialacc_cls as ent_ls, b.clstype, b.meanstype, '+
                 'c.mfo as k_mfo, c.cor as k_cor, '+
                 'd.acc as k_rs, '+
                 'e.kdvalue, '+
                 'f.finsourcevalue, '+
                 'h.inn as ent_inn, h.inn20 as ent_kpp, '+
                 'h.shortname as ent_sname, h.name as ent_name '+
             'from '+
                   'quotestitle a, '+
                   'incomes32 b, '+
                   'org_accounts d, '+
                   'kd e, '+
                   'innerfinsource f, '+
                   'facialacc_cls g, '+
                   'organizations h, '+
                   'banks c '+
             'where '+
                   'a.rejectcls is null and '+
                   'a.id=b.recordindex and '+
                   'b.accountref=d.id and '+
                   'd.bank_ref=c.id and '+
                   'b.kd=e.id and '+
                   'b.ifs=f.id and '+
                   'b.facialacc_cls=g.id and '+
                   'g.org_ref=h.id and '+
                   'a.acceptdate>='+d1+' and '+
                   'a.acceptdate<='+d2;

     if not(addselect = EmptyStr) then
        select += ' and ' + addselect;

     ls:=Edit3.Text;
     if not(ls=EmptyStr) then select+=' and b.facialacc_cls='+ls;

     UIBDataSet1.SQL.Clear;
     UIBDataSet1.SQL.Add(select);
     UIBDataSet1.Open;

     err:=2;

     dbf1.Close;
     dbf1.TableName:=savepath+'bnd_main.dbf';
     dbf1.FieldDefs.Clear;

     dbf1.FieldDefs.Add('id',ftString,15);
     dbf1.FieldDefs.Add('ent_inn',ftString,12);
     dbf1.FieldDefs.Add('ent_kpp',ftString,9);
     dbf1.FieldDefs.Add('ent_sname',ftstring,1024);
     dbf1.FieldDefs.Add('ent_name',ftstring,1024);
     dbf1.FieldDefs.add('ent_ls',ftString,25);
     dbf1.FieldDefs.add('k_id',ftstring,15);
     dbf1.FieldDefs.add('k_rs',ftString,20);
     dbf1.FieldDefs.add('k_mfo',ftString,9);
     dbf1.FieldDefs.add('k_cor',ftString,20);
     dbf1.FieldDefs.add('docnum',ftString,50);
     dbf1.FieldDefs.add('docdate',ftString,8);
     dbf1.FieldDefs.add('acceptdate',ftstring,8);
     dbf1.FieldDefs.Add('credit',ftFloat,15);
     dbf1.FieldDefs[Dbf1.FieldDefs.Count-1].Precision := 2;
     dbf1.FieldDefs.Add('debit',ftFloat,15);
     dbf1.FieldDefs[Dbf1.FieldDefs.Count-1].Precision := 2;
     dbf1.FieldDefs.add('note',ftstring,255);
     dbf1.FieldDefs.Add('clstype',ftstring,1);
     dbf1.FieldDefs.add('kd',ftstring,20);
     dbf1.FieldDefs.add('ifs',ftstring,20);
     dbf1.FieldDefs.add('kvd',ftstring,20);

     if FileExists(savePath+'bnd_main.dbf')=True then
        DeleteFile(savePath+'bnd_main.dbf');

     Dbf1.CreateTable;
     Dbf1.Exclusive:=True;
     Dbf1.Open;

     Dbf2.Close;
     Dbf2.TableName:=savePath+'bnd_org.dbf';
     Dbf2.FieldDefs.Clear;

     Dbf2.FieldDefs.Add('id'       ,ftString,15);
     Dbf2.FieldDefs.Add('inn'      ,ftString,12);
     Dbf2.FieldDefs.Add('kpp'      ,ftString,9);
     Dbf2.FieldDefs.Add('name'     ,ftString,1024);
     Dbf2.FieldDefs.Add('shortname',ftString,1024);
     Dbf2.FieldDefs.Add('okato'    ,ftString,11);

     if FileExists(savePath+'bnd_org.dbf')=True then
        DeleteFile(savePath+'bnd_org.dbf');

     Dbf2.CreateTable;
     Dbf2.Exclusive:=True;
     Dbf2.Open;

     Dbf3.Close;
     Dbf3.TableName:=savePath+'bnd_fkr.dbf';
     Dbf3.FieldDefs.Clear;

     Dbf3.FieldDefs.Add('type' ,ftInteger,1);
     Dbf3.FieldDefs.Add('stat' ,ftString,20);

     if FileExists(savePath+'bnd_fkr.dbf')=True then
        DeleteFile(savePath+'bnd_fkr.dbf');

     Dbf3.CreateTable;
     Dbf3.Exclusive:=True;
     Dbf3.Open;

     ProgressBar1.Min:=1;
     UIBDataSet1.Last;
     ProgressBar1.Max:=UIBDataSet1.RecordCount;

     perc:=1;
     err:=3;
     org_idx:='/';
     fkr_idx:='/00000000000000000000/'; //таким образом исключаем кбк
     select:=EmptyStr;
     UIBDataSet1.First;
     while not UIBDataSet1.EOF do
     begin
       Dbf1.Insert;
       Dbf1.FieldByName('id').AsString:=UIBDataSet1.FieldByName('id').AsString;
       Dbf1.FieldByName('ent_inn').AsString:=UIBDataSet1.FieldByName('ent_inn').AsString;
       Dbf1.FieldByName('ent_kpp').AsString:=UIBDataSet1.FieldByName('ent_kpp').AsString;
       Dbf1.FieldByName('ent_sname').AsString:=AnsiToConsole(UIBDataSet1.FieldByName('ent_sname').AsString);
       Dbf1.FieldByName('ent_name').AsString:=AnsiToConsole(UIBDataSet1.FieldByName('ent_name').AsString);
       Dbf1.FieldByName('ent_ls').AsString:=AnsiToConsole(UIBDataSet1.FieldByName('ent_ls').AsString);
       Dbf1.FieldByName('k_id').AsString:=UIBDataSet1.FieldByName('k_id').AsString;
       Dbf1.FieldByName('k_rs').AsString:=UIBDataSet1.FieldByName('k_rs').AsString;
       Dbf1.FieldByName('k_mfo').AsString:=UIBDataSet1.FieldByName('k_mfo').AsString;
       Dbf1.FieldByName('k_cor').AsString:=UIBDataSet1.FieldByName('k_cor').AsString;
       Dbf1.FieldByName('docnum').AsString:=AnsiToConsole(UIBDataSet1.FieldByName('docnum').AsString);
       Dbf1.FieldByName('docdate').AsString:=UIBDataSet1.FieldByName('docdate').AsString;
       Dbf1.FieldByName('acceptdate').AsString:=UIBDataSet1.FieldByName('acceptdate').AsString;
       Dbf1.FieldByName('note').AsString:=AnsiToConsole(UIBDataSet1.FieldByName('note').AsString);
       Dbf1.FieldByName('clstype').AsString:=UIBDataSet1.FieldByName('clstype').AsString;
       Dbf1.FieldByName('debit').AsFloat:=UIBDataSet1.FieldByName('debit').AsFloat;
       Dbf1.FieldByName('credit').AsFloat:=UIBDataSet1.FieldByName('credit').AsFloat;
       Dbf1.FieldByName('kd').AsString:=UIBDataSet1.FieldByName('kdvalue').AsString;
       Dbf1.FieldByName('ifs').AsString:=UIBDataSet1.FieldByName('finsourcevalue').AsString;
       Dbf1.FieldByName('kvd').AsFloat:=UIBDataSet1.FieldByName('meanstype').AsFloat;
       Dbf1.Post;

       str:=UIBDataSet1.FieldByName('kdvalue').AsString;
       if pos('/'+str+'/',fkr_idx)=0 then
       begin
         fkr_idx+=str+'/';
         dbf3.Insert;
         dbf3.FieldByName('type').AsInteger:=1;
         dbf3.FieldByName('stat').AsString:=str;
         dbf3.Post;
       end;

       str:=UIBDataSet1.FieldByName('finsourcevalue').AsString;
       if pos('/'+str+'/',fkr_idx)=0 then
       begin
         fkr_idx+=str+'/';
         dbf3.Insert;
         dbf3.FieldByName('type').AsInteger:=2;
         dbf3.FieldByName('stat').AsString:=str;
         dbf3.post;
       end;

       str:=UIBDataSet1.FieldByName('k_id').AsString;
       if pos('/'+str+'/',org_idx)=0 then
       begin
         org_idx+=str+'/';
         if not(select=EmptyStr) then select+=' or ';
         select+='id='+str
       end;

       UIBDataSet1.Next;
       perc+=1;
       ProgressBar1.Position:=perc;
       ProgressBar1.Refresh;
     end;

     err:=4;
     CollectOrg(select);

     Dbf1.Close;
     Dbf2.Close;
     Dbf3.Close;

     err:=5;
     SetLength(dbffiles,3);
     dbffiles[0] := 'bnd_main.dbf';
     dbffiles[1] := 'bnd_org.dbf';
     dbffiles[2] := 'bnd_fkr.dbf';
     PackToZip(savepath, 'bnd', dbffiles);

     err:=0;

  except
     StatusBar1.Panels.Items[0].Text:='Ошибка';
     case err of
       1:ShowMessage('Произошла неопознанная ошибка, проверьте параметры');
       2:ShowMessage('Ошибка создания файлов выгрузки');
       3:ShowMessage('Ошибка заполнения базы документов');
       4:ShowMessage('Ошибка заполнения базы организаций');
       5:ShowMessage('Ошибка создания архива');
     end;

  end;
end;

end.
