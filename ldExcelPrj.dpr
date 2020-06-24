program ldExcelPrj;



uses
  Forms,
  ldExcel in 'ldExcel.pas' {FldExcel};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TFldExcel, FldExcel);
  Application.Run;
end.
