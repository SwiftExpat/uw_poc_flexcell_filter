unit frmExcelFilterExportU;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.TMSFNCTypes, Vcl.TMSFNCUtils, Vcl.TMSFNCGraphics, Vcl.TMSFNCGraphicsTypes,
  Vcl.StdCtrls, Vcl.TMSFNCPageControl, Vcl.TMSFNCCustomControl, Vcl.TMSFNCTabSet, Vcl.TMSFNCGridCell,
  Vcl.TMSFNCGridOptions, Vcl.TMSFNCToolBar, Vcl.TMSFNCCustomScrollControl, Vcl.TMSFNCGridData, Vcl.TMSFNCCustomGrid,
  Vcl.TMSFNCGrid, Vcl.FlexCel.Core, FlexCel.XlsAdapter, Vcl.TMSFNCStatusBar;

type
  {
    Question, is "last worked on" a Date value that needs to be sorted?
    If it needs to sort in Delhpi, it should be parsed or read from xlsx as a date.
  }

  {
    Depending on how much functionality is desired, TMSFNCUExcelAdapter, would be a good unit to review and decide what to implement.
    Reviewing the list of units in the xlsAdapter folder would quickly identify possible goals, ie Hyperlinks and notes on cells.
  }

  /// <summary>API class to wrap common functions of Flexcell and make them Delphi control friendly. </summary>
  /// <remarks>Unit Test this api. Included in form only for easy review, should be separte unit to share easily.</remarks>
  /// <returns>Items returned based on zero index for Delphi.</returns>
  TXLSXSource = class
  strict private
    FExcelFile: TExcelFile;
    FFileName: string;
  public
    constructor Create(AFileName: string);
    destructor Destroy; override;
    /// <summary>API Access to the Excel File </summary>
    /// <returns>TExcel file for raw access to the API</returns>
    property ExcelFile: TExcelFile read FExcelFile;
    /// <summary> /\/\ </summary>
    /// <remarks> /\/\ </remarks>
    /// <returns> /\/\ </returns>
    procedure ReadTabs(ATabList: TStringList);
    /// <summary>Returns column count for the active sheet</summary>
    /// <returns>Int value of the column count</returns>
    function ColCount: Integer;
    /// <summary>Returns row count for the active sheet</summary>
    /// <returns>Int value of the row count </returns>
    function RowCount: Integer;
    /// <summary>Returns the cell value at column & row specified</summary>
    /// <remarks>need to determine data only or cell formulas</remarks>
    /// <returns>String value of the cell in excel</returns>
    { TODO : Adjust based on cell type, see notes in function }
    function CellValue(const ARow, ACol: Integer): string;
    /// <summary>Allows selecting the sheet if more than one available</summary>
    /// <remarks>Implemented based on 0 index in Delphi</remarks>
    { TODO : Convert to property to query active/ selected sheet. query should be read & select should be write }
    procedure SelectSheet(AZeroIndex: Integer);
  end;

  TfrmExcelFilterExport = class(TForm)
    pcMain: TTMSFNCPageControl;
    TMSFNCPageControl1Page0: TTMSFNCPageControlContainer;
    TMSFNCPageControl1Page1: TTMSFNCPageControlContainer;
    Memo1: TMemo;
    sbMain: TTMSFNCStatusBar;
    TMSFNCToolBar1: TTMSFNCToolBar;
    DataGrid: TTMSFNCGrid;
    btnFileOpen: TTMSFNCToolBarButton;
    btnSave: TTMSFNCToolBarButton;
    TMSFNCToolBarSeparator1: TTMSFNCToolBarSeparator;
    cbSheets: TTMSFNCToolBarItemPicker;
    procedure btnFileOpenClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure cbSheetsItemSelected(Sender: TObject; AItemIndex: Integer);
    procedure FormShow(Sender: TObject);
  strict private
    FXLSXSource: TXLSXSource;
    FProgress: TTMSFNCStatusBarPanel;
  private
    procedure StatusMessage(AMessage: string);
    procedure LogMsg(AMessage: string);
    function ExportFileName(const AFileExt: string): string;
    function FilterString(const AFileExt: string): string;
    /// <summary>Call back method for the FNC save file dialog</summary>
    /// <remarks>Calls FNC Datagrid export which has a built in csv export</remarks>
    procedure SaveCSV(const AFile: string; const AResult: boolean);
    /// <summary>Call back method for the FNC select file dialog</summary>
    /// <remarks>Calls FNC Datagrid export which has a built in csv export</remarks>
    procedure OpenXLSX(const AFile: string; const AResult: boolean);
    /// <summary>Loads dropdown with a list of the sheets in the workbook </summary>
    procedure LoadTabList;
    /// <summary>Loads the data to the grid</summary>
    /// <remarks>? is Last worked on a date value?</remarks>
    procedure LoadGridData;
  public
    { Public declarations }
  end;

var
  frmExcelFilterExport: TfrmExcelFilterExport;

implementation

{$R *.dfm}

uses System.IOUtils, System.DateUtils;

{ TfrmExcelFilterExport }

procedure TfrmExcelFilterExport.btnFileOpenClick(Sender: TObject);
const
  ext_xlsx = 'xlsx';
var
  fn: string;
  fr: boolean;
begin
  fr := TTMSFNCUtils.SelectFile(fn, 'c:\temp', FilterString(ext_xlsx), OpenXLSX);
  if fr then
    LogMsg('File Opened @ ' + fn)
  else
    LogMsg('Failed export to ' + fn)
end;

procedure TfrmExcelFilterExport.OpenXLSX(const AFile: string; const AResult: boolean);
begin
  if AResult = false then
    exit;
  try
    if FXLSXSource <> nil then
      FXLSXSource.Free;
    FXLSXSource := TXLSXSource.Create(AFile);
    LoadTabList;
    LoadGridData;
  except
    on E: Exception do
      LogMsg(E.message);
  end;
end;

procedure TfrmExcelFilterExport.SaveCSV(const AFile: string; const AResult: boolean);
begin
  if AResult = false then
    exit;
  try
    DataGrid.SaveToCsv(AFile);
    { TODO -oCSV -cDataFormat : This does not save column names to the CSV, just the data. File format might change solution }
  except
    on E: Exception do
      LogMsg(E.message);
  end;
end;

procedure TfrmExcelFilterExport.btnSaveClick(Sender: TObject);
const
  ext_csv = 'csv';
var
  fn: string;
  fr: boolean;
begin
  fn := ExportFileName(ext_csv);
  fr := TTMSFNCUtils.SaveFile(fn, FilterString(ext_csv), SaveCSV);
  if fr then
    StatusMessage('Exported @ ' + fn)
  else
    LogMsg('Failed export to ' + fn)
end;

procedure TfrmExcelFilterExport.LoadGridData;
var
  r, c, j, k, s: Integer;
  st, et: TDateTime;
begin
  DataGrid.BeginUpdate;
  try
    st := now;
    DataGrid.Clear;
    r := FXLSXSource.RowCount;
    c := FXLSXSource.ColCount;
    DataGrid.ColumnCount := c;
    DataGrid.RowCount := r;
    FProgress.Visible := true;

    for j := 1 to r do
    begin
      LogMsg('Loading row #' + j.ToString);
      s := round(((j / r) * 100));
      if s > FProgress.Progress.Position + 2 then
      begin // avoid delay on the progress bar, repaint is slow
        sbMain.BeginUpdate;
        FProgress.Progress.Position := round(((j / r) * 100));
        sbMain.EndUpdate;
        sbMain.Repaint;
      end;
      for k := 1 to c do
      begin
        // LogMsg('Load Value r=' + j.ToString + ' c=' + k.ToString);
        DataGrid.Cells[(k - 1), (j - 1)] := FXLSXSource.CellValue(j, k);
        { TODO :
          Update API method to handle the one based index. Loops starting at 1 are harder to review.
          This part of the code is Delphi, make the API convert! }
      end;
    end;
    et := now;
    StatusMessage('Grid filled in ' + MillisecondsBetween(st, et).ToString + ' ms');
    DataGrid.AutoSizeColumns(false, 8);
    et := now;
    StatusMessage('Grid Sized & Filled ' + MillisecondsBetween(st, et).ToString + ' ms');
  finally
    DataGrid.EndUpdate;
    FProgress.Visible := false;
  end;
end;

procedure TfrmExcelFilterExport.LoadTabList;
begin
  cbSheets.Items.Clear;
  FXLSXSource.ReadTabs(cbSheets.Items);
  { TODO : Does not query the api for the selected sheet. Extend API and set based on active sheet }
  if cbSheets.Items.Count > 0 then
    cbSheets.SelectedItemIndex := 0;
end;

procedure TfrmExcelFilterExport.cbSheetsItemSelected(Sender: TObject; AItemIndex: Integer);
begin
  if FXLSXSource = nil then
    exit;

  FXLSXSource.SelectSheet(AItemIndex); // API will handle 1 based index
  LoadGridData;
end;

function TfrmExcelFilterExport.ExportFileName(const AFileExt: string): string;
const
  fmt_datetime = 'yyyy_mm_dd_hh_nn_ss';
begin
  result := TPath.Combine(TTMSFNCUtils.GetTempPath, 'GridExport_');
  result := result + FormatDateTime(fmt_datetime, now) + '.' + AFileExt;
end;

function TfrmExcelFilterExport.FilterString(const AFileExt: string): string;
begin
  result := AFileExt.ToUpper + ' Files (*.' + AFileExt.ToLower + ')|*.' + AFileExt.ToUpper
end;

procedure TfrmExcelFilterExport.FormDestroy(Sender: TObject);
begin
  FreeAndNil(FXLSXSource);
end;

procedure TfrmExcelFilterExport.FormShow(Sender: TObject);
begin
  if FProgress = nil then
    FProgress := sbMain.Panels.Items[1];
end;

procedure TfrmExcelFilterExport.LogMsg(AMessage: string);
begin
{$IFNDEF RELEASE}
  Memo1.Lines.Add(AMessage)
{$ENDIF}
end;

procedure TfrmExcelFilterExport.StatusMessage(AMessage: string);
begin
  sbMain.Panels.Items[0].Text := AMessage;
  Memo1.Lines.Add('Status Updated :  ' + AMessage);
end;

{ TXLSXSource }

function TXLSXSource.CellValue(const ARow, ACol: Integer): string;
var
  v: TCellValue;
begin
  try
    v := FExcelFile.GetCellValue(ARow, ACol);
    case v.ValueType of
      TCellValueType.Empty:
        result := '';
    else
      result := FExcelFile.GetCellValue(ARow, ACol).ToString;
    end;
    { TODO -oXLSX -cSelect Data : This returns all values as strings. Review the data if dates are included and determine how to parse }
    { GetCellValue returns a TCellValue which can be used in a case statement to
      form the demo code in UReadingFiles, implement something like this:
      function TFReadingFiles.FormatValue(const v: TCellValue; const Row, Col: integer): String; }
  except
    on E: Exception do
      result := 'nil';
  end;
end;

function TXLSXSource.ColCount: Integer;
begin
  if FExcelFile <> nil then
    result := FExcelFile.ColCount
  else
    result := 0;
end;

constructor TXLSXSource.Create(AFileName: string);
begin
  FFileName := AFileName;
  if TFile.Exists(FFileName) then
  begin
    FExcelFile := TXlsFile.Create(false);
    FExcelFile.Open(FFileName);
  end
  else
    raise Exception.Create('File not found');

end;

destructor TXLSXSource.Destroy;
begin
  FExcelFile.Free;
  inherited;
end;

procedure TXLSXSource.ReadTabs(ATabList: TStringList);
var
  i: Integer;
begin
  ATabList.Clear;
  for i := 1 to FExcelFile.SheetCount do
  begin
    ATabList.Add(FExcelFile.GetSheetName(i));
  end;
end;

function TXLSXSource.RowCount: Integer;
begin
  if FExcelFile <> nil then
    result := FExcelFile.RowCount
  else
    result := 0;
end;

procedure TXLSXSource.SelectSheet(AZeroIndex: Integer);
begin
  { TODO -oXLSX -cSelect Data : Add some range checking here }
  FExcelFile.ActiveSheet := AZeroIndex + 1;
end;

end.
