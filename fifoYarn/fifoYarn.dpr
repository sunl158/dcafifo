program fifoYarn;

uses
  FastMM4,
  FastCode,
  FastMove,
  RtlVclOptimize,
  ControlResizeBugFix,
  SysUtils,
  Forms,
  Windows,
  MainF in 'MainF.pas' {FrmMain},
  aboutf in 'aboutf.pas' {aboutForm: TTntForm},
  cpicker in 'cpicker.pas' {CPickerForm: TTntForm},
  ChoixImpF in 'ChoixImpF.pas' {ChoixImpForm: TTntForm},
  DataColorF in 'DataColorF.pas' {DataColorForm},
  fancyf in 'fancyf.pas' {FancyForm: TTntForm},
  RepeatF in 'RepeatF.pas' {RepeatForm: TTntForm},
  ScanWizardF in 'ScanWizardF.pas' {ScanForm: TTntForm},
  SlubF in 'SlubF.pas' {SlubForm},
  SpaceDF in 'SpaceDF.pas' {SpaceDyeForm},
  PrefF in 'PrefF.pas' {PrefFrom: TTntForm},
  ThumbF in 'ThumbF.pas' {ThumbForm: TTntForm},
  ADjustF in 'ADjustF.pas' {ADjustForm},
  DataF in 'DataF.pas' {DataForm: TTntForm},
  Datas in '..\Common\Datas.pas' {DM: TDataModule},
  DobbyWF in '..\Common\DobbyWF.pas' {DobbyWizardFrom: TTntForm},
  simulf in '..\YX-Weave New\simulf.pas' {SimulForm: TTntForm},
  selimpf in '..\YX-Weave New\selimpf.pas' {SelImp: TTntForm},
  libF in 'libF.pas' {LibraryForm: TTntForm},
  YarnUnit in '..\Common\YarnUnit.pas',
  FindColorF in 'FindColorF.pas' {FindColorForm: TTntForm};

{$R *.res}

begin   
  Application.MainFormOnTaskBar := True;
  ShowWindow(Application.Handle, SW_HIDE);
  SetWindowLong(Application.Handle, GWL_EXSTYLE,
  GetWindowLong(Application.Handle, GWL_EXSTYLE) and not WS_EX_APPWINDOW);
  AboutForm := TAboutForm.Create(Application);
  try
    AboutForm.Show;
    AboutForm.Refresh;
    Application.Initialize;
    Application.Title := 'fifo-Yarn';
  Application.CreateForm(TFrmMain, FrmMain);
  Application.CreateForm(TDM, DM);
  AboutForm.Update;
  finally
    AboutForm.Close;
  end;
  Application.Run;

end.

