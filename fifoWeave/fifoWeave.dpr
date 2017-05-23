Program fifoWeave;
{$SetPEFlags $20}

uses
  FastMM4,
  FastCode,
  FastMove,
  RtlVclOptimize,
  SysUtils,
  Forms,
  Windows,
  Datas in '..\Common\Datas.pas' {DM: TDataModule},
  varianteF in 'varianteF.pas' {VarianteForm: TTntForm},
  cpicker in 'cpicker.pas' {CPickerForm: TTntForm},
  ChoixImpColorF in 'ChoixImpColorF.pas' {ChoixImpFormColor: TTntForm},
  DataColorF in 'DataColorF.pas' {DataColorForm},
  aboutf in 'aboutf.pas' {aboutForm: TTntForm},
  simulf in 'simulf.pas' {SimulForm: TTntForm},
  DesignF in 'DesignF.pas' {DesignForm: TTntForm},
  DataF in 'DataF.pas' {DataForm: TTntForm},
  EntretienF in 'EntretienF.pas' {EntretienForm: TTntForm},
  BarCodeF in 'BarCodeF.pas' {BarcodeForm: TTntForm},
  wrepeatf in 'wrepeatf.pas' {WeaveRepeatform: TTntForm},
  KEYBD in 'KEYBD.pas' {Clavier: TTntForm},
  mecanicF in 'mecanicF.pas' {MecanicForm: TTntForm},
  Sim3DF in 'Sim3DF.pas' {Sim3DForm: TTntForm},
  varsimul in 'varsimul.pas' {varsimulForm: TTntForm},
  HugeRepeatF in 'HugeRepeatF.pas' {HugeRepeatForm: TTntForm},
  CustPreviewF in 'CustPreviewF.pas' {CustPreviewForm: TTntForm},
  MultiImpF in 'MultiImpF.pas' {MultiImpForm: TTntForm},
  MultiFabricF in 'MultiFabricF.pas' {MultiFabricForm: TTntForm},
  selimpf in 'selimpf.pas' {SelImp: TTntForm},
  ColorweaveConverterF in 'ColorweaveConverterF.pas' {ColorweaveConverter: TTntForm},
  TechDobbyF in 'TechDobbyF.pas' {TechDobbyForm},
  KEYBD2 in 'KEYBD2.pas' {RentrageClavierForm: TTntForm},
  DraftWF in 'DraftWF.pas' {DraftWizardForm: TTntForm},
  RentrageF in 'RentrageF.pas' {RentrageForm: TTntForm},
  InsertF in 'InsertF.pas' {InsertForm: TTntForm},
  SelectSelF in 'SelectSelF.pas' {SelectSelForm: TTntForm},
  DisposeF in 'DisposeF.pas' {DisposeForm: TTntForm},
  analysef in 'analysef.pas' {AnalyseForm: TTntForm},
  analysF in 'analysF.pas' {AnalysForm: TTntForm},
  Image2DataF in 'Image2DataF.pas' {Image2DataForm: TTntForm},
  centrageF in 'centrageF.pas' {centrageform: TTntForm},
  DobbySelectF in 'DobbySelectF.pas' {DobbySelectForm},
  randomF in 'randomF.pas' {randomform: TTntForm},
  ScanF in 'ScanF.pas' {ScanForm: TTntForm},
  BlanketF in 'BlanketF.pas' {BlanketForm},
  SurimpDimF in 'SurimpDimF.pas' {SurimpDimForm: TTntForm},
  DenimF in 'DenimF.pas' {DenimForm: TTntForm},
  CrinkleF in 'CrinkleF.pas' {CrinkleForm: TTntForm},
  importCF in 'importCF.pas' {ImportCardFrom: TTntForm},
  FloppyF in 'FloppyF.pas' {FloppyForm: TTntForm},
  classic in 'classic.pas' {classicform: TTntForm},
  EncantrageF in 'EncantrageF.pas' {EncantrageForm: TTntForm},
  ChoixImpF in 'ChoixImpF.pas' {ChoixImpForm: TTntForm},
  CatalogF in 'CatalogF.pas' {CatalogForm: TTntForm},
  SearchF in 'SearchF.pas' {SearchForm: TTntForm},
  TodoF in 'TodoF.pas' {ToDoForm: TTntForm},
  ChoixImpFBlanket in 'ChoixImpFBlanket.pas' {ChoixImpFormBlanket: TTntForm},
  BlanketNameF in 'BlanketNameF.pas' {BlanketNameForm},
  jacqwizard2F in 'jacqwizard2F.pas' {JacqWizardForm2: TTntForm},
  PiquageClavierF in 'PiquageClavierF.pas' {PiquageClavierForm: TTntForm},
  ProductionDobby in '..\Common\ProductionDobby.pas',
  ProductionJacquard in '..\Common\ProductionJacquard.pas',
  melangeurF in '..\Common\melangeurF.pas' {MelangeurForm: TTntForm},
  machines in '..\Common\machines.pas' {MachineForm: TTntForm},
  PrefF in '..\Common\PrefF.pas' {PrefForm: TTntForm},
  YarnUnit in '..\Common\YarnUnit.pas',
  VirtualYarnF in 'VirtualYarnF.pas' {VirtualYarnform},
  flipWF in 'flipWF.pas' {FlipWeaveForm: TTntForm},
  FullRepeatF in 'FullRepeatF.pas' {FullRepeatForm},
  RegSeqF in 'RegSeqF.pas' {RegulatorSequenceform},
  DobbyWF in '..\Common\DobbyWF.pas' {DobbyWizardFrom: TTntForm},
  JacquardWF in 'JacquardWF.pas' {NewJacForm: TTntForm},
  ClavierJacF in 'ClavierJacF.pas' {ClavierJacForm: TTntForm},
  MappingLibrary_Intf in '..\Common\MappingLibrary_Intf.pas',
  CylindreF in 'CylindreF.pas' {Cylindreform},
  PicanolContF in 'PicanolContF.pas' {PicanolControlerForm},
  boucleF in 'boucleF.pas' {BoucleForm: TTntForm},
  EpongeF in 'EpongeF.pas' {EpongeForm: TTntForm},
  LeonardoF in 'LeonardoF.pas' {LeonardoForm},
  ColletageStaubF in 'ColletageStaubF.pas' {ColletageStaubliForm},
  CustLattageF in 'CustLattageF.pas' {CustomLattageForm: TTntForm},
  LattageClavierF in 'LattageClavierF.pas' {LattageClavierForm: TTntForm},
  StatJacqF in 'StatJacqF.pas' {StatForm: TTntForm},
  FMain in 'FMain.pas' {MainForm},
  CompoundlistF in 'CompoundlistF.pas' {CompoundListForm},
  correctF in 'correctF.pas' {CorrectForm: TTntForm},
  FillWeaveUnit in 'FillWeaveUnit.pas',
  Fabric3DF in 'Fabric3DF.pas' {Fabric3DForm},
  DepavageF in 'DepavageF.pas' {DepavageForm},
  DB1ImportF in 'DB1ImportF.pas' {DB1Importer},
  ADjustF in 'ADjustF.pas' {ADjustForm},
  ColorwayWizardF in 'ColorwayWizardF.pas' {ColorwayWizardForm},
  FindColorF in 'FindColorF.pas' {FindColorForm},
  DegradeF in 'DegradeF.pas' {DegradeForm},
  InsertMECF in 'InsertMECF.pas' {InsertMECForm: TTntForm},
  TracBarF in 'C2007\SPTBXLib\Source\TracBarF.pas' {TrackBarForm: TTntForm},
  RedistLattageF in 'RedistLattageF.pas' {RedistLattageForm},
  JacqF in 'JacqF.pas' {JacquardFrom},
  RemoveitemFromMECF in 'RemoveitemFromMECF.pas' {RemoveitemFromMECForm},
  MECService_Impl in '..\Common\MEC\MECService_Impl.pas' {MECService: TRORemoteDataModule},
  MECLibrary_Intf in '..\Common\MEC\MECLibrary_Intf.pas',
  MECLibrary_Invk in '..\Common\MEC\MECLibrary_Invk.pas',
  GradientF in 'GradientF.pas' {GradientForm},
  pegplanF in 'pegplanF.pas' {PegPlanForm},
  FunctionsF in 'FunctionsF.pas' {FunctionsForm},
  DegSeqF in 'DegSeqF.pas' {DegSeqForm},
  DobImportF in 'DobImportF.pas' {DobImportForm},
  ThumbF in 'ThumbF.pas' {Thumbform},
  InsertDobbyF in 'InsertDobbyF.pas' {InsertDobbyForm},
  cartonClavier in 'cartonClavier.pas' {CartonClavierForm},
  Weave3DlibUnit in 'Weave3DlibUnit.pas',
  Visualweave in 'Simul 3D - Sans dll\Visualweave.pas',
  ExchangeShaftF in 'ExchangeShaftF.pas' {ExchangeShaftForm};

{$R *.res}


Begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'fifo-Weave';
  ShowWindow(Application.Handle, SW_HIDE);
  SetWindowLong(Application.Handle, GWL_EXSTYLE,
  GetWindowLong(Application.Handle, GWL_EXSTYLE) and not WS_EX_APPWINDOW);
  AboutForm := TAboutForm.Create(Application);
  Try
    AboutForm.Show;
    AboutForm.Refresh;
    Application.CreateForm(TMainForm, MainForm);
  Application.CreateForm(TTrackBarForm, TrackBarForm);
  Application.CreateForm(TDM, DM);
  AboutForm.Update;
  Finally
    AboutForm.Close;
    AboutForm := Nil;
  End;
  Application.Run;

End.

