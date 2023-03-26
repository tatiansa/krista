program project1;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms, Unit1, uiblaz, dbflaz, memdslaz;

{$R *.res}

begin
  Application.Title:='URMload';
  Application.Initialize;
  Application.Run;
end.

