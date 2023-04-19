program ExcelFilterExport;

uses
  Vcl.Forms,
  frmExcelFilterExportU in 'frmExcelFilterExportU.pas' {frmExcelFilterExport};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmExcelFilterExport, frmExcelFilterExport);
  Application.Run;
end.
