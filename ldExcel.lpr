program ldExcel;

{$MODE Delphi}



uses
  Forms, Interfaces,
  ldExcelApp in 'ldExcelApp.pas' {FldExcel};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TFldExcel, FldExcel);
  Application.Run;
end.
