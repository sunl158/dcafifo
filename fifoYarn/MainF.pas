Unit MainF;

Interface
{$I Hasp.inc}

Uses
  hyiedefs, CMProfile, CMProfile_types, Dib, Gr32, Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Dialogs, AdvOfficeStatusBar, Math, pngImage, tntforms,
  AdvOfficeStatusBarStylers, AdvToolBar, AdvToolBarStylers, AdvGlowButton, Menus, AdvMenus,
  ImgList, AdvOfficeHint, AdvMenuStylers, StdCtrls, ExtCtrls, Couleur, resample,
  ActnList, TntActnList, AdvAlertWindow, AdvGrid, frmctrllink, WinSkinData,
  AdvPreviewMenu, AdvPreviewMenuStylers, SpTBXItem, TB2Item, TBX, TBXAthenTheme, TB2Dock,
  TBXDkPanels, SpTBXDkPanels, SpTBXEditors, AdvShapeButton, UnitSpin,
  TntStdCtrls, AdvOfficeButtons, Grids, BaseGrid, Mask,
  Spin, RXSpin, AdvGroupBox, AdvStyleIF, imageenproc, imageenio,
  ProgressDialog, TntDialogs, ieopensavedlg, ExtDlgs, siComp, siLang_Tnt,
  SMEEngine, ExportDS, SMEWiz, TB2ExtItems, TBXExtItems, TB2Toolbar,
  TntExtCtrls, ZLibEX, animatedcursor, PngImageList, goChecker, ZBitmap, ZDataBitmap,
  IceLicense, IceLicenseCompDate, ProtectHL, ieview, imageenview, imageen,
  MCoPRecentMenu, rxToolEdit, rxCurrEdit, ShFolder, TntMenus, YarnUnit, jpeg,
  MDIWallp, SpTBXSkins, SpTBXControls, PegtopFileDialogs;

Const
  vers = 21;

Type
  TFrmMain = Class(TAdvToolBarForm)
    AdvToolBarOfficeStyler1: TAdvToolBarOfficeStyler;
    AdvToolBarPager1: TAdvToolBarPager;
    AdvPage1: TAdvPage;
    AdvPage2: TAdvPage;
    AdvToolBar1: TAdvToolBar;
    ImprimeBtn: TAdvGlowButton;
    AdvToolBar4: TAdvToolBar;
    ArtificialPage: TAdvPage;
    AdvPage4: TAdvPage;
    ChineBtn: TAdvGlowButton;
    MoulineBtn: TAdvGlowButton;
    MulticoBtn: TAdvGlowButton;
    AdvGlowButton3: TAdvGlowButton;
    AdvGlowButton4: TAdvGlowButton;
    SimulBtn: TAdvGlowButton;
    AdvOfficeHint1: TAdvOfficeHint;
    AdvMenuOfficeStyler1: TAdvMenuOfficeStyler;
    AdvShapeButton2: TAdvShapeButton;
    AdvPreviewMenuOfficeStyler1: TAdvPreviewMenuOfficeStyler;
    AdvQuickAccessToolBar1: TAdvQuickAccessToolBar;
    AdvGlowButton13: TAdvGlowButton;
    JaspeBtn: TAdvGlowButton;
    TextureBtn: TAdvGlowButton;
    FancyBtn: TAdvGlowButton;
    SpaceDyeBtn: TAdvGlowButton;
    SlubBtn: TAdvGlowButton;
    AdvPage5: TAdvPage;
    AdvToolBar12: TAdvToolBar;
    RG_sens_fil: TAdvOfficeRadioGroup;
    AdvToolBarSeparator1: TAdvToolBarSeparator;
    Label8: TTntLabel;
    fil_titre: TRxSpinEdit;
    TntLabel2: TTntLabel;
    TitreCb: TTntComboBox;
    SkinData1: TSkinData;
    Label3: TTntLabel;
    fil_torsion: TSpinEdit;
    Label7: TTntLabel;
    fil_densite: TRxSpinEdit;
    Label13: TTntLabel;
    fil_prix: TCurrencyEdit;
    AdvToolBar2: TAdvToolBar;
    AdvToolBar3: TAdvToolBar;
    Label25: TTntLabel;
    PilositeEdit: TSpinEdit;
    Label26: TTntLabel;
    DensPoilEdit: TSpinEdit;
    TntRadioGroup1: TAdvGroupBox;
    Image1: TImage;
    ContrastEdit: TSpinEdit;
    CompressReliefCB: TAdvOfficeCheckBox;
    TntGroupBox2: TTntGroupBox;
    Image2: TImage;
    RondeurEdit: TSpinEdit;
    CompressRondeurCB: TAdvOfficeCheckBox;
    AdvToolBar5: TAdvToolBar;
    DelavableCB: TAdvOfficeCheckBox;
    TntLabel3: TTntLabel;
    PrelavageEdit: TSpinEdit;
    AdvToolBar6: TAdvToolBar;
    Label2: TTntLabel;
    fil_irregularite_Longueur: TSpinEdit;
    Label15: TTntLabel;
    F_Irregularite: TSpinEdit;
    AdvToolBar13: TAdvToolBar;
    Label23: TTntLabel;
    DiametreEdit: TSpinEdit;
    ChineToolbar: TAdvToolBar;
    Label12: TTntLabel;
    Label10: TTntLabel;
    FibreCB: TAdvOfficeCheckBox;
    Edit_TailleCarreau: TSpinEdit;
    EspacementEdit: TSpinEdit;
    MulticoToolbar: TAdvToolBar;
    Label4: TTntLabel;
    label5: TTntLabel;
    Label6: TTntLabel;
    multico_random_max: TSpinEdit;
    multico_rapport: TSpinEdit;
    space_rapport: TSpinEdit;
    AdvToolBar8: TAdvToolBar;
    Label16: TTntLabel;
    RangeEdit: TSpinEdit;
    GrilleMatiere: TAdvStringGrid;
    MatiereCb: TTntComboBox;
    AdvOfficeStatusBarOfficeStyler1: TAdvOfficeStatusBarOfficeStyler;
    SB: TAdvOfficeStatusBar;
    AdvAlertWindow1: TAdvAlertWindow;
    ActionList2: TTntActionList;
    ActColorsClose: TTntAction;
    ColorsWindow: TSpTBXDockablePanel;
    SpTBXLabelItem16: TSpTBXLabelItem;
    SpTBXRightAlignSpacerItem10: TSpTBXRightAlignSpacerItem;
    SpTBXItem26: TSpTBXItem;
    ColorListBox: TListBox;
    PD: TPrintDialog;
    ImageEnProc1: TImageEnProc;
    sdfil: TSaveImageEnDialog;
    SDImage: TSaveImageEnDialog;
    OPD: TOpenPictureDialog;
    IO: TImageEnIO;
    odfil: TOpenImageEnDialog;
    psd: TPrinterSetupDialog;
    ODImage: TOpenImageEnDialog;
    ody2t: TTntOpenDialog;
    sdy2t: TTntSaveDialog;
    spb: TProgressDialog;
    SMEWizardDlg1: TSMEWizardDlg;
    SMEVirtualDataEngine1: TSMEVirtualDataEngine;
    siLangmain: TsiLang_Tnt;
    PngImageList1: TPngImageList;
    TntActionList1: TTntActionList;
    ActQuit: TTntAction;
    ActSimul: TTntAction;
    ActDatas: TTntAction;
    ActDobbyWizard: TTntAction;
    ActChine: TTntAction;
    ActMouline: TTntAction;
    ActMultico: TTntAction;
    ActImprime: TTntAction;
    ActJaspe: TTntAction;
    ActTexture: TTntAction;
    ActFancy: TTntAction;
    ActSpaceDye: TTntAction;
    ActSlub: TTntAction;
    AdvGlowMenuButton1: TAdvGlowMenuButton;
    PMZoom: TTntPopupMenu;
    N301: TTntMenuItem;
    N1001: TTntMenuItem;
    N2001: TTntMenuItem;
    N3001: TTntMenuItem;
    N4001: TTntMenuItem;
    N6001: TTntMenuItem;
    N8001: TTntMenuItem;
    ImageSmall: TPngImageList;
    imageListMenu: TPngImageList;
    ActionListMenu: TTntActionList;
    ActOpen: TTntAction;
    ActQuitMedium: TTntAction;
    ActSave: TTntAction;
    ActSaveAS: TTntAction;
    ActBrowse: TTntAction;
    ActPrinterSetup: TTntAction;
    ActPrint: TTntAction;
    ActExportDatas: TTntAction;
    ActImportDatas: TTntAction;
    ActCreateGamme: TTntAction;
    ActCheckUpdate: TTntAction;
    ActAdjustColor: TTntAction;
    ActionListSmall: TTntActionList;
    ActQuitSmall: TTntAction;
    ActPref: TTntAction;
    ActSimulSmall: TTntAction;
    goChecker1: TgoChecker;
    ActTexturedWizard: TTntAction;
    ActFancyWizard: TTntAction;
    ActSlubWizard: TTntAction;
    ActSpaceDyeWizard: TTntAction;
    ActShowColorGrid: TTntAction;
    AdvGlowButton2: TAdvGlowButton;
    N501: TTntMenuItem;
    ActShowPalette: TTntAction;
    ShowPaletteBtn: TAdvGlowButton;
    ActfindColor: TTntAction;
    SpTBXItem1: TSpTBXItem;
    ImageListWizard: TPngImageList;
    TBImageList1: TImageList;
    IL1: TIceLicenseCompDate;
    Timer1: TTimer;
    ActExportImage: TTntAction;
    ActOpenPalette: TTntAction;
    ODPalette: TTntOpenDialog;
    SpTBXItem2: TSpTBXItem;
    ImageListFlag: TPngImageList;
    Actionlistflag: TTntActionList;
    French: TTntAction;
    English: TTntAction;
    German: TTntAction;
    Spanish: TTntAction;
    Turkish: TTntAction;
    Chinese: TTntAction;
    YarnToolbar: TAdvDockPanel;
    AdvToolBar7: TAdvToolBar;
    ImageEn1: TImageEn;
    ActSaveAsSmall: TTntAction;
    AdvGlowButton1: TAdvGlowButton;
    AdvPreviewMenu1: TAdvPreviewMenu;
    ActOpenfilSmall: TTntAction;
    AdvGlowMenuButton2: TAdvGlowMenuButton;
    HistoryMenu: TTntPopupMenu;
    RecentMenu1: TRecentMenu;
    SpTBXStatusBar2: TSpTBXStatusBar;
    LibraryFilename: TSpTBXLabelItem;
    SpTBXStatusBar1: TSpTBXStatusBar;
    ColorCountTxt: TSpTBXLabelItem;
    AdvToolBar10: TAdvToolBar;
    DobbyWizard: TAdvGlowButton;
    MDIWallpaper1: TMDIWallpaper;
    SpTBXMultiDock1: TSpTBXMultiDock;
    ColorGridToolbar: TSpTBXDockablePanel;
    ColorGrid: TAdvStringGrid;
    SpTBXPanel1: TSpTBXPanel;
    SpTBXLabel1: TSpTBXLabel;
    fil_nbcolor: TSpTBXSpinEdit;
    colorpanel: TPanel;
    PegtopExtendedOpenDialog1: TPegtopExtendedOpenDialog;
    PopupLanguage: TSpTBXPopupMenu;
    SpTBXItem3: TSpTBXItem;
    SpTBXItem4: TSpTBXItem;
    SpTBXItem5: TSpTBXItem;
    SpTBXItem6: TSpTBXItem;
    SpTBXItem7: TSpTBXItem;
    SpTBXItem8: TSpTBXItem;
    RetorsEdit: TRxSpinEdit;
    ActUpdate: TTntAction;
    Procedure AdvToolBar2OptionClick(Sender: TObject; ClientPoint,
      ScreenPoint: TPoint);
    Procedure FormCreate(Sender: TObject);
    Procedure AdvGlowButton24Click(Sender: TObject);
    Procedure AdvPreviewMenu1ButtonClick(Sender: TObject; ButtonIndex: Integer);
    Procedure Showtoolbarontop1Click(Sender: TObject);
    Procedure Showtoolbarbelow1Click(Sender: TObject);
    Procedure ColorGridGetDisplText(Sender: TObject; ACol, ARow: Integer;
      Var Value: String);
    Procedure ColorGridGetEditorType(Sender: TObject; ACol, ARow: Integer;
      Var AEditor: TEditorType);
    Procedure GrilleMatiereGetEditorType(Sender: TObject; ACol,
      ARow: Integer; Var AEditor: TEditorType);
    Procedure RangeEditChange(Sender: TObject);
    Procedure ActColorsCloseExecute(Sender: TObject);
    Procedure ColorListBoxDragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; Var Accept: Boolean);
    Procedure ColorListBoxDrawItem(Control: TWinControl; Index: Integer;
      Rect: TRect; State: TOwnerDrawState);
    Procedure ColorListBoxMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    Procedure ColorGridDragDrop(Sender, Source: TObject; X, Y: Integer);
    Procedure ColorGridDragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; Var Accept: Boolean);
    Procedure FormClose(Sender: TObject; Var Action: TCloseAction);
    Procedure ContrastEditChange(Sender: TObject);
    Procedure FormCloseQuery(Sender: TObject; Var CanClose: Boolean);
    Procedure ActQuitExecute(Sender: TObject);
    Procedure ActSimulExecute(Sender: TObject);
    Procedure ActDobbyWizardExecute(Sender: TObject);
    Procedure ActChineExecute(Sender: TObject);
    Procedure N301Click(Sender: TObject);
    Procedure ActOpenExecute(Sender: TObject);
    Procedure ActSaveExecute(Sender: TObject);
    Procedure ActSaveASExecute(Sender: TObject);
    Procedure ActPrinterSetupExecute(Sender: TObject);
    Procedure ActPrintExecute(Sender: TObject);
    Procedure ActSimulSmallUpdate(Sender: TObject);
    Procedure ActPrefExecute(Sender: TObject);
    Procedure ActBrowseExecute(Sender: TObject);
    Procedure ActAdjustColorExecute(Sender: TObject);
    Procedure ActCreateGammeExecute(Sender: TObject);
    Procedure ActCheckUpdateExecute(Sender: TObject);
    Procedure ActExportDatasExecute(Sender: TObject);
    Procedure ActDatasExecute(Sender: TObject);
    Procedure goChecker1AbortDownload(Sender: TObject);
    Procedure goChecker1Available(VersionInfo, DateBuildInfo: String);
    Procedure goChecker1BeforeUpdate(Var RunUpdate: Boolean);
    Procedure goChecker1DownloadError(ErrorMsg: String;
      StatusCode: Integer);
    Procedure goChecker1NotOnline(Sender: TObject);
    Procedure goChecker1NoUpdate(Sender: TObject);
    Procedure goChecker1UpdateError(ErrorMsg: String);
    Procedure SMEVirtualDataEngine1Count(Sender: TObject;
      Var Count: Integer);
    Procedure SMEVirtualDataEngine1FillColumns(Sender: TObject);
    Procedure SMEVirtualDataEngine1GetValue(Sender: TObject;
      Column: TSMEColumn; Var Value: Variant);
    Procedure ActTexturedWizardExecute(Sender: TObject);
    Procedure ActFancyWizardExecute(Sender: TObject);
    Procedure ActSlubWizardExecute(Sender: TObject);
    Procedure ActSpaceDyeWizardExecute(Sender: TObject);
    Procedure ActShowColorGridExecute(Sender: TObject);
    Procedure ActShowPaletteExecute(Sender: TObject);
    Procedure ColorListBoxStartDrag(Sender: TObject;
      Var DragObject: TDragObject);
    Procedure ActfindColorExecute(Sender: TObject);
    Procedure ActfindColorUpdate(Sender: TObject);
    Procedure IL1ExeNotEncrypted(Sender: TObject);
    Procedure IL1LicenseFull(ExpiresInfo: Integer;
      ExtraLicenseInfo: String);
    Procedure IL1NoneLicense(Sender: TObject);
    Procedure Timer1Timer(Sender: TObject);
    Procedure ActExportImageExecute(Sender: TObject);
    Procedure ActExportDatasUpdate(Sender: TObject);
    Procedure ActAdjustColorUpdate(Sender: TObject);
    Procedure ActImportDatasUpdate(Sender: TObject);
    Procedure ActExportImageUpdate(Sender: TObject);
    Procedure ActOpenPaletteExecute(Sender: TObject);
    Procedure FrenchExecute(Sender: TObject);
    Procedure ColorGridDblClickCell(Sender: TObject; ARow, ACol: Integer);
    Procedure ActSaveAsSmallUpdate(Sender: TObject);
    Procedure RecentMenu1Click(Sender: TMenuItem; FileName: TFileName);
    Procedure siLangmainChangeLanguage(Sender: TObject);
    Procedure FormControlEditLink1GetEditorValue(Sender: TObject;
      Grid: TAdvStringGrid; Var AValue: String);
    Procedure FormControlEditLink1SetEditorFocus(Sender: TObject;
      Grid: TAdvStringGrid; AControl: TWinControl);
    Procedure FormControlEditLink1SetEditorValue(Sender: TObject;
      Grid: TAdvStringGrid; AValue: String);
    Procedure fil_nbcolorValueChanged(Sender: TObject);
    Procedure ColorGridToolbarClose(Sender: TObject);
    Procedure ActUpdateExecute(Sender: TObject);
    Procedure Button1Click(Sender: TObject);
  Private
    { Private declarations }
    stopgamme: boolean;
    CentreTop, CentreBottom, NBExec: integer;
    Procedure LimitRelief;
    Procedure LimitRondeur;
    Procedure OnAppIdle(Sender: TObject; Var Done: Boolean);
  Protected
  Public
    { Public declarations }
    Procedure UpdateTypeDeFil;
    Procedure CreerFil;
    Procedure SetProgress(Position: integer);
    Procedure IncProgress;
    Procedure EndProgress;
    Procedure DoShowMessage(TheMessage: WideString);
    Procedure ConvertDibToRGB(Var ADib: TDibImage);
    Procedure ConvertDibToDisplay(Var ADib: TDibImage);
    Procedure ConvertRGBToDib(Var ADib: TDibImage);
    Procedure SauverFilToStream(FS: TStream);
    Procedure SauverFil(filename: String);
    Procedure DoFabricSimulation(Sciseaux: boolean);
    Procedure OpenFil(fname: String);
    Function GetTypeDeFilName(index: integer): String;
    Procedure DoAboutProgress(NewPosition: Integer);

  End;

Var
  FrmMain: TFrmMain;
  TypeDeFil, PreviousTypeDeFil: integer;
  DragColor: TColor;
  DragName: String;
  DragValue: EncodedLabValue;
  UseUnicode: boolean;

  UseCompressedStream, DoNotConfirmClose, ValidLicence: boolean;
  user_maxundo: integer;
  TempDir: String;
  ID_HL: longint;

  DensityUnit: integer;                 //0 cm, 1 inch

  UsePrintingServer, HasChamatexPluggin: boolean;
  Spoolfolder: String;
  CacheThumb, NoNeedToSimulForThumb: boolean;

  //
  HSVList: TList;                       //La liste de la bibli de couleur;
  //si on sauve le fil dans un fichier ou dans un memorystream
  SavingYarnToFile: boolean;
  chargement, LoginDone: boolean;

  mainformCreated, FirstSimul: boolean;
  AppFolder, PasDePub: String;
  prelavage, TheZoom, DentVideLg: integer;
  DentVide: Array[0..14999] Of boolean;

  ExportFolder: String;

  //Données technique
  CodeMatiere, CodeCouleur, Description: String;
  CodeBarre: String;
  DesignationCouleur: String;
  CodeRecetteTeinture: String;
  DesignationRecetteTeinture: String;
  TitrageCommercial: String;
  CodeComposition: String;
  DesignationComposition: String;
  CodeFammile: String;
  DesignationFamille: String;
  FilTeint: boolean;
  Observations: String;
  CodeUniteStockage: String;
  DesignationUniteStockage: String;
  CodeStructure: String;
  DesignationStructure: String;
  ObservationTechnique: String;
  TypeArticle: String;

Procedure PBar(pos, max: integer);

Implementation

Uses inifiles, ResampleGDIPlus, cpicker, Datas, uVistaFuncs, fancyf, RepeatF, ScanWizardF, SlubF, SpaceDF, DobbyWF,
  PrefF, ThumbF, ADjustF, libF,  DataF, aboutf, selimpf, registry, ShlObj;

{$R *.dfm}
{$R YX-Yarn.RES}

Procedure PBar(pos, max: integer);
Var
  tmp: integer;
Begin
  If Max = 0 Then
    exit;
  tmp := GetTickCount;
  If tmp - tps > 1000 Then
    Begin
      DM.SetProgress(pos * 100 Div max);
      tps := tmp;
    End;
  application.ProcessMessages;
End;

Function DefaultSaveLocation: String;
Var
  P: PChar;
Begin
  P := Nil;
  Try
    P := AllocMem(MAX_PATH);
    If SHGetFolderPath(0, CSIDL_APPDATA, 0, 0, P) = S_OK Then
      Result := P
    Else
      Result := GetCurrentDir;
  Finally
    FreeMem(P);
  End;
End;

Function DocumentFolder: String;
Var
  P: PChar;
Begin
  P := Nil;
  Try
    P := AllocMem(MAX_PATH);
    If SHGetFolderPath(0, CSIDL_PERSONAL, 0, 0, P) = S_OK Then
      Result := P
    Else
      Result := GetCurrentDir;
  Finally
    FreeMem(P);
  End;
End;

Procedure TFrmMain.FormCreate(Sender: TObject);
Var
  i, j, exec: integer;
Begin
  Try

    DossierPrefYarn := '';
    DossierIni := '';
    If IsWindowsVista Then
      Begin
        DossierIni := DefaultSaveLocation + '\fifo';
        If Not Directoryexists(DossierIni) Then
          ForceDirectories(DossierIni);
        DossierIni := DossierIni + '\';

        DossierPrefYarn := DefaultSaveLocation + '\fifo\fifo-Yarn';
        If Not Directoryexists(DossierPrefYarn) Then
          ForceDirectories(DossierPrefYarn);
        DossierPrefYarn := DossierPrefYarn + '\';
      End;
    UseUnicode := FileExists(DossierIni + 'unicode.txt');
    screen.Cursor := CrWorkingBGAqua;
    NBExec := 0;
    If IsWindowsVista Then
      Begin
        SkinData1.Active := false;
        SetDesktopIconFonts(Self.Font);
      End;
    SkinManager.SetSkin(SkinManager.SkinsList[9]);
    TBXSetTheme('Athen');
    ValidLicence := false;
    IL1.Run;
    AppFolder := extractFilePath(ParamStr(0));

    AdvToolBarOfficeStyler1.Style := bsOffice2007Silver;
    AdvPreviewMenuOfficeStyler1.Style := psOffice2007Silver;
    GrilleMatiere.SetComponentStyle(tsOffice2007Silver);
    ColorGrid.SetComponentStyle(tsOffice2007Silver);

    If Not ValidLicence Then
      Begin
        application.Terminate;
        exit;
      End;

    If IL1.LicenseType = ltNone Then
      Begin
        application.Terminate;
        exit;
      End;
    Timer1.Enabled := true;

    If IsSoftIce95Loaded Or IsSoftIceNTLoaded Then
      Application.Terminate;
    InitializeKey(IL1);

{$IFDEF HASP4}
    If Not CheckHasp(IL1) Then
      Begin
        FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_42' (* 'Please connect the hardware protection key to the computer!!' *)));
        Application.Terminate;
        exit;
      End;

    If Not Checkfifo(IL1) Then
      Begin
        FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_43' (* 'Contact fifo to use this software' *)));
        Application.Terminate;
        exit;
      End;

    If Not CheckYarn(IL1) Then
      Begin
        FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_44' (* 'Contact fifo to be able to use YX-WEAVE' *)));
        Application.Terminate;
        exit;
      End;

    exec := CheckNumberOfExecutionYarn(IL1);

    goChecker1.InfoURL := IL1.SecureStrings[2];

{$ELSE}

    If Not CheckYxendis_HL(IL1) Then
      Begin
        FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_43' (* 'Contact fifo to use this software' *)));
        Application.Terminate;
        exit;
      End;

    ID_HL := GetID_HL(IL1);
    NomClient := GetUserName_HL(IL1);
    AboutForm.GetBuildInfo;

    LoginDone := LoginYarn_HL(IL1);

    If Not LoginDone Then
      Begin
        FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_46' (* 'Contact fifo to be able to use fifo-Yarn' *)));
        Application.Terminate;
        exit;
      End;

    If Not CheckYarn_HL(IL1) Then
      Begin
        FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_46' (* 'Contact fifo to be able to use fifo-Yarn' *)));
        Application.Terminate;
        exit;
      End;

    If Not IsYarnMaintenanceValid(IL1) Then
      Begin
        FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_48' (* 'Your maintenance contract has expired, you cannot use this release !' *)));
        Application.Terminate;
        exit;
      End;

    exec := CheckNumberOfExecution_HL(IL1);

    goChecker1.InfoURL := IL1.SecureStrings[3];
{$ENDIF}

    If (exec < 20) And (exec > 0) Then
      FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_49' (* 'Contact fifo to update the key, you can only execute this software ' *)) + inttostr(exec + 1) + siLangmain.GetTextOrDefault('IDS_50' (* ' times more !' *)));
    If (exec = 0) Then
      FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_51' (* 'Contact fifo to update the key, this is the last time you can use this software !' *)));

    AboutForm.GetBuildInfo;
    YarnBitmap := TDibimage.Create(True);
    MulticoToolbar.Visible := false;
    FirstSimul := true;

    ColorsWindow.floating := true;
    ColorsWindow.Parent.Left := 600;
    ColorsWindow.Parent.Top := 338;
    ColorGridToolbar.Options.RightAlignSpacer.FontSettings.Style := [];
    ColorGridToolbar.Options.RightAlignSpacer.FontSettings.Color := $0076C1FF;

    ColorGrid.SpinEdit.MinValue := 1;
    ColorGrid.SpinEdit.MaxValue := 100;

    Imageen1.Cursor := CrPan;

    TempDir := GetTempDirectory;
    YarnStream := TMemorystream.Create;

    user_maxundo := 16;
    TheZoom := 100;

    SlubTitle := 10;
    NBSectionSlub := 2;
    StoneContrast := 4;
    AdvToolBarPager1.ActivePageIndex := 0;

    NBSection := 4;
    For i := 0 To 20 Do
      Begin
        SectionColor[i].rgbRed := 127;
        SectionColor[i].rgbGreen := 127;
        SectionColor[i].rgbBlue := 255;
        SectionColor[i].rgbReserved := 255;
        SectionSize[i] := 10;

        If odd(i) Then
          SectionSlubSize[i] := 150
        Else
          SectionSlubSize[i] := 30;
        sectionSlubPentG[i] := 75;
        sectionSlubPentD[i] := 75;
      End;

    SlubColor.b := 99;
    SlubColor.a := 120;
    SlubColor.L := 96;

    TypeDeFil := 0;
    PreviousTypeDeFil := 0;
    PercentContrast := 25;

    PasDePub := GetSysDir + '\NoAdvertissement.dat';

    For j := 1 To 10 Do
      Begin
        ColorGrid.Cells[2, j] := '1';
        tabpred[i] := 1;
      End;

    nbcolorfil := Trunc(fil_nbcolor.Value);

    CentreTop := 0;
    CentreBottom := 19;
    Ytop := 0;
    YBottom := 19;

    HSVList := TList.Create;

    GrilleMatiere.WideCells[0, 0] := siLangmain.GetTextOrDefault('IDS_114' (* 'Qtité' *));
    GrilleMatiere.WideCells[1, 0] := siLangmain.GetTextOrDefault('IDS_115' (* 'Matiere' *));

    GrilleMatiere.cells[0, 1] := '100';
    GrilleMatiere.cells[1, 1] := 'Coton';

    GrilleMatiere.ComboBox.Clear;
    For i := 0 To MatiereCB.Items.count - 1 Do
      GrilleMatiere.Combobox.Items.Add(MatiereCB.Items[i]);

    opd.InitialDir := ExpandFileName(DataFolder + '\' + YarnFolder);
    sdfil.InitialDir := ExpandFileName(DataFolder + '\' + YarnFolder);
    RecentMenu1.MakeRecentMenus;
    Application.OnIdle := OnAppIdle;
  Finally
    screen.Cursor := CrDefault;
  End;
End;

Procedure TFrmMain.OnAppIdle(Sender: TObject; Var Done: Boolean);
Begin
  MoulineBtn.Down := TTNTAction(MoulineBtn.Action).Checked;
  ChineBtn.Down := TTNTAction(ChineBtn.Action).Checked;
  MulticoBtn.Down := TTNTAction(MulticoBtn.Action).Checked;
  ImprimeBtn.Down := TTNTAction(ImprimeBtn.Action).Checked;
  JaspeBtn.Down := TTNTAction(JaspeBtn.Action).Checked;
  TextureBtn.Down := TTNTAction(TextureBtn.Action).Checked;
  FancyBtn.Down := TTNTAction(FancyBtn.Action).Checked;
  SpaceDyeBtn.Down := TTNTAction(SpaceDyeBtn.Action).Checked;
  SlubBtn.Down := TTNTAction(SlubBtn.Action).Checked;
End;

Procedure TFrmMain.Showtoolbarbelow1Click(Sender: TObject);
Begin
  AdvToolBarPager1.ShowQATBelow := true;
End;

Procedure TFrmMain.Showtoolbarontop1Click(Sender: TObject);
Begin
  AdvToolBarPager1.ShowQATBelow := false;
End;

{*------------------------------------------------------------------------------
  Change language of software
  @param Sender   ParameterDescription
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure TFrmMain.siLangmainChangeLanguage(Sender: TObject);
Var
  i, j: integer;
  Mylist: TStrings;
Begin

  If siLangmain.ActiveLanguage = 8 Then
    Begin
      Colorgrid.FixedFont.Name := 'Simsun';
      Colorgrid.FixedFont.Charset := GB2312_CHARSET;
      GrilleMatiere.FixedFont.Name := 'Simsun';
      GrilleMatiere.FixedFont.Charset := GB2312_CHARSET;
    End
  Else
    Begin
      Colorgrid.FixedFont.Name := 'Tahoma';
      Colorgrid.FixedFont.Charset := ANSI_CHARSET;
      GrilleMatiere.FixedFont.Name := 'Tahoma';
      GrilleMatiere.FixedFont.Charset := ANSI_CHARSET;
    End;

  GrilleMatiere.ComboBox.Clear;
  For i := 0 To MatiereCB.Items.count - 1 Do
    GrilleMatiere.Combobox.Items.Add(MatiereCB.Items[i]);

  Colorgrid.wideCells[0, 0] := siLangmain.GetTextOrDefaultW('IDS_12' (* 'N°' *));
  Colorgrid.wideCells[1, 0] := siLangmain.GetTextOrDefaultW('IDS_16' (* 'Couleur' *));
  Colorgrid.wideCells[2, 0] := siLangmain.GetTextOrDefaultW('IDS_35' (* 'Predominance' *));
  GrilleMatiere.wideCells[0, 0] := siLangmain.GetTextOrDefaultW('IDS_36' (* 'Qtité (%)' *));
  GrilleMatiere.wideCells[1, 0] := siLangmain.GetTextOrDefaultW('IDS_115' (* 'Matiere' *));

  For i := 0 To AdvPreviewMenu1.MenuItems.Count - 1 Do
    Begin
      If TADVPreviewMenuItem(AdvPreviewMenu1.MenuItems[i]).Action <> Nil Then
        Begin
          TADVPreviewMenuItem(AdvPreviewMenu1.MenuItems[i]).WideCaption := TTNTAction(TADVPreviewMenuItem(AdvPreviewMenu1.MenuItems[i]).Action).Caption;
          TADVPreviewMenuItem(AdvPreviewMenu1.MenuItems[i]).Caption := '';
        End;
      For j := 0 To TADVPreviewMenuItem(AdvPreviewMenu1.MenuItems[i]).SubItems.Count - 1 Do
        Begin
          If TADVPreviewSubMenuItem(AdvPreviewMenu1.MenuItems[i].SubItems[j]).Action <> Nil Then
            Begin
              TADVPreviewSubMenuItem(AdvPreviewMenu1.MenuItems[i].SubItems[j]).WideTitle := TTNTAction(TADVPreviewSubMenuItem(AdvPreviewMenu1.MenuItems[i].SubItems[j]).Action).Caption;
              TADVPreviewSubMenuItem(AdvPreviewMenu1.MenuItems[i].SubItems[j]).Title := '';
            End;
        End;

    End;

  For i := 0 To AdvPreviewMenu1.SubMenuItems.Count - 1 Do
    Begin
      AdvPreviewMenu1.SubMenuItems[i].WideTitle := TTNTAction(TADVPreviewSubMenuItem(AdvPreviewMenu1.SubMenuItems[i]).Action).Caption;
      AdvPreviewMenu1.SubMenuItems[i].Title := '';
    End;

  If UseUnicode And (ComponentCount > 0) Then
    Begin
      For i := 0 To componentcount - 1 Do
        If Components[i] Is TAdvGlowButton Then
          Begin
            If (TAdvGlowButton(Components[i]).Action <> Nil) And TAdvGlowButton(Components[i]).ShowCaption Then
              Begin
                TAdvGlowButton(Components[i]).Caption := '';
                TAdvGlowButton(Components[i]).WideCaption := TTNTAction(TAdvGlowButton(Components[i]).Action).Caption;
              End;
          End;

      AdvToolBar1.Caption := '';
      AdvToolBar1.WideCaption := siLangmain.GetTextOrDefaultW('IDS_18' (* 'Type de fil' *));
      AdvToolBar4.Caption := '';
      AdvToolBar4.WideCaption := siLangmain.GetTextOrDefaultW('IDS_20' (* 'Actions' *));
      AdvToolBar12.Caption := '';
      AdvToolBar12.WideCaption := siLangmain.GetTextOrDefaultW('IDS_21' (* 'Epaisseur du fil' *));
      AdvToolBar8.Caption := '';
      AdvToolBar8.WideCaption := siLangmain.GetTextOrDefaultW('IDS_22' (* 'Composition' *));
      AdvToolBar2.Caption := '';
      AdvToolBar2.WideCaption := siLangmain.GetTextOrDefaultW('IDS_23' (* 'Volume' *));
      AdvToolBar3.Caption := '';
      AdvToolBar3.WideCaption := siLangmain.GetTextOrDefaultW('IDS_24' (* 'Pilosité' *));
      AdvToolBar5.Caption := '';
      AdvToolBar5.WideCaption := siLangmain.GetTextOrDefaultW('IDS_25' (* 'Aptitude au délavage' *));
      AdvToolBar6.Caption := '';
      AdvToolBar6.WideCaption := siLangmain.GetTextOrDefaultW('IDS_26' (* 'Anti-moirrage' *));
      AdvToolBar13.Caption := '';
      AdvToolBar13.WideCaption := siLangmain.GetTextOrDefaultW('IDS_27' (* 'Effet irrégulier' *));
      ChineToolbar.Caption := '';
      ChineToolbar.WideCaption := siLangmain.GetTextOrDefaultW('IDS_28' (* 'Chinés' *));
      MulticoToolbar.Caption := '';
      MulticoToolbar.WideCaption := siLangmain.GetTextOrDefaultW('IDS_29' (* 'Multicouleurs / Imprimé' *));
      AdvToolBar10.Caption := '';
      AdvToolBar10.WideCaption := siLangmain.GetTextOrDefaultW('IDS_41' (* 'Tissage' *));
      ColorGridToolbar.Caption := siLangmain.GetTextOrDefaultW('IDS_30' (* 'Coloris' *));

      AdvPage2.Caption := '';
      AdvPage2.WideCaption := siLangmain.GetTextOrDefaultW('IDS_31' (* 'Interface principale et simulation' *));
      AdvPage5.Caption := '';
      AdvPage5.WideCaption := siLangmain.GetTextOrDefaultW('IDS_32' (* 'Propriétés' *));
      AdvPage1.Caption := '';
      AdvPage1.WideCaption := siLangmain.GetTextOrDefaultW('IDS_33' (* 'Aspect' *));
      ArtificialPage.Caption := '';
      ArtificialPage.WideCaption := siLangmain.GetTextOrDefaultW('IDS_34' (* 'Moteur artificiel' *));
      AdvPage4.Caption := '';
      AdvPage4.WideCaption := siLangmain.GetTextOrDefaultW('IDS_17' (* 'Simulation' *));
    End;
End;

Procedure TFrmMain.AdvGlowButton24Click(Sender: TObject);
Begin
  ShowMessage('Search GoTo function');
End;

Procedure TFrmMain.AdvPreviewMenu1ButtonClick(Sender: TObject; ButtonIndex: Integer);
Begin
  Case ButtonIndex Of
    0: Close;
    1:
      Begin
        AdvPreviewMenu1.HideMenu;
        ShowMessage('Handle options here');
      End;
  End;
End;

Procedure TFrmMain.AdvToolBar2OptionClick(Sender: TObject; ClientPoint,
  ScreenPoint: TPoint);
Begin
  //ShowMessage('Toolbar options can be made available here');
End;

Procedure TFrmMain.Button1Click(Sender: TObject);
Var
  a, b: integer;
Begin
  a := 5;
  b := 0;
  showmessage(inttostr(a Div b));
End;

{*------------------------------------------------------------------------------
  Update yarn type on GUI
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure TFrmMain.UpdateTypeDeFil;
Begin
  ArtificialPage.TabVisible := TypeDeFil In [1, 2, 3];
  ChineToolbar.Visible := TypeDeFil = 1;
  MulticoToolbar.Visible := TypeDeFil In [2, 3];

  Case TypeDeFil Of                     // type de fil
    0: ActMouline.Checked := true;
    1: Actchine.Checked := true;
    2: ActMultico.Checked := true;
    3: ActImprime.Checked := true;
    4: ActJaspe.Checked := true;
    5: ActTexture.Checked := true;
    6: ActFancy.Checked := true;
    7: ActSpaceDye.Checked := true;
    8: ActSlub.Checked := true;
  End;

End;

Procedure TFrmMain.ColorGridGetDisplText(Sender: TObject; ACol,
  ARow: Integer; Var Value: String);
Begin
  If ACol > 0 Then
    exit;
  If (ACol = 0) And (ARow > 0) Then
    Value := 'Col ' + inttostr(Arow);
End;

Procedure TFrmMain.ColorGridGetEditorType(Sender: TObject; ACol,
  ARow: Integer; Var AEditor: TEditorType);
Begin
  If (Arow > 0) And (Acol = 2) Then
    Begin
      AEditor := edSpinEdit;
    End;

  {if (Arow > 0) and (Acol = 1) then
    AEditor := edEditBtn;   }
End;

Procedure TFrmMain.ColorGridToolbarClose(Sender: TObject);
Begin
  ActShowColorGrid.checked := false;
End;

Procedure TFrmMain.GrilleMatiereGetEditorType(Sender: TObject; ACol,
  ARow: Integer; Var AEditor: TEditorType);
Begin
  If (ACol > 0) And (ARow > 0) Then
    Begin
      AEditor := edComboEdit;
    End;
End;

Procedure TFrmMain.RangeEditChange(Sender: TObject);
Begin
  GrilleMatiere.RowCount := RangeEdit.Value + 1;
  If GrilleMatiere.RowCount < 2 Then
    GrilleMatiere.RowCount := 2;
End;

Procedure TFrmMain.SetProgress(Position: integer);
Begin
  If visible Then
    sb.Panels[0].Progress.Position := Position;
  If sb.Panels[0].Progress.Position = 0 Then
    EndProgress;
  application.ProcessMessages;
End;

Procedure TFrmMain.IncProgress;
Begin
  If visible Then
    sb.Panels[0].Progress.Position := sb.Panels[0].Progress.Position + 1;
  If sb.Panels[0].Progress.Position = 0 Then
    EndProgress;
  application.ProcessMessages;
End;

Procedure TFrmMain.EndProgress;
Begin
  If visible Then
    sb.Panels[0].Progress.Position := 0;
  application.ProcessMessages;
End;

Procedure TFrmMain.DoShowMessage(TheMessage: WideString);
Begin
  ShowMessage(TheMessage);
End;

Procedure TFrmMain.ActColorsCloseExecute(Sender: TObject);
Begin
  ColorsWindow.Visible := false;
  //ShowPaletteBtn.Down:=False;
  ActShowPalette.Checked := false;
End;

Procedure TFrmMain.ColorListBoxDragOver(Sender, Source: TObject; X,
  Y: Integer; State: TDragState; Var Accept: Boolean);
Begin
  Accept := Source = ColorListBox;
End;

Procedure TFrmMain.ColorListBoxDrawItem(Control: TWinControl;
  Index: Integer; Rect: TRect; State: TOwnerDrawState);
Var
  ll, aa, bb, r1, g1, b1: byte;
  CNVS: TCanvas;
Begin
  CNVS := ColorListBox.Canvas;
  Cnvs.Brush.Color := Color;
  Cnvs.FillRect(Rect);
  If index = ColorListBox.ItemIndex Then
    Cnvs.Font.Color := ClRed
  Else
    Cnvs.Font.Color := ClBlue;
  Cnvs.TextOut(Rect.Left + 2, Rect.Top + 2, ColorListBox.Items[index]);
  If Index < ColorListBox.Items.Count Then
    Begin
      LL := L2Fix(TColorLib(ListeCouleur.Items[index]).L);
      AA := ab2Fix(TColorLib(ListeCouleur.Items[index]).a);
      BB := ab2Fix(TColorLib(ListeCouleur.Items[index]).b);
      hDisplayProfile.GetRGBFromLab(LL, AA, BB, r1, g1, b1);
      Cnvs.Brush.Color := TColor(rgb(r1, g1, b1));
      Cnvs.Rectangle(Rect.Left + 2, Rect.top + 16, Rect.right - 4, Rect.bottom - 2);
    End;
End;

Procedure TFrmMain.ColorListBoxMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
Begin
  ColorListBox.BeginDrag(false, 10);
End;

Procedure TFrmMain.ColorGridDragDrop(Sender, Source: TObject; X,
  Y: Integer);
Var
  ACol, ARow: integer;
Begin
  If Source = ColorListbox Then
    Begin
      ColorGrid.MouseToCell(x, y, Acol, Arow);
      If (ACol = 1) And (Arow > 0) Then
        Begin
          ColorGrid.Colors[Acol, Arow] := DragColor;
          ColorGrid.FontColors[Acol, Arow] := clwhite Xor DragColor;
          ColorGrid.Cells[Acol, Arow] := DragName;
          NomCouleur[Arow - 1] := DragName;
          tabcolor[Arow - 1].moyen.rgbRed := DragValue.b;
          tabcolor[Arow - 1].moyen.rgbGreen := DragValue.a;
          tabcolor[Arow - 1].moyen.rgbBlue := DragValue.L;
          tabcolor[Arow - 1].moyen.rgbReserved := 255;
          tabcolor[Arow - 1].clair := Eclaircir(tabcolor[Arow - 1].moyen, PercentContrast);
          tabcolor[Arow - 1].fonce := Assombrir(tabcolor[Arow - 1].moyen, PercentContrast);
          ActSimul.Enabled := true;
        End;
    End;
End;

Procedure TFrmMain.ColorGridDragOver(Sender, Source: TObject; X,
  Y: Integer; State: TDragState; Var Accept: Boolean);
Var
  ACol, ARow: integer;
Begin
  If Source = ColorListbox Then
    Begin
      ColorGrid.MouseToCell(x, y, Acol, Arow);
      If (ACol = 1) And (Arow > 0) Then
        Begin
          Accept := True;
        End;
    End;
End;

Procedure TFrmMain.FormClose(Sender: TObject; Var Action: TCloseAction);
Var
  i: integer;
Begin
  If YarnBitmap <> Nil Then
    FreeAndNil(YarnBitmap);

  For i := 0 To HSVList.Count - 1 Do
    TColorLib(HSVList.Items[i]).Free;   // libération de la mémoire
  // pour chaque élément de la liste
//Voir remarque sur la notion de transtype plus haut dans le code
  If HSVList <> Nil Then
    FreeAndNil(HSVList);

{$IFDEF HASPHL}
  If LoginDone Then
    DoLogOut_HL(IL1);
{$ENDIF}
End;

{*------------------------------------------------------------------------------
  Apply printer profile
  @param ADib   ParameterDescription
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure TfrmMain.ConvertDibToRGB(Var ADib: TDibImage);
Var
  Bits: pointer;
Begin
  If ADib.IsImageAssigned Then
    Try
      Bits := ADib.Bits;
      TCMProfile.ConvertInRGBFromLab(Bits, ADib.Width * ADib.Height, 4, hPPaperProfile);
    Except
      FrmMain.siLangmain.ShowMessage('Profile to Profile conversion Failed');
    End;
End;

{*------------------------------------------------------------------------------
  Apply screen profile
  @param ADib   ParameterDescription
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure TfrmMain.ConvertDibToDisplay(Var ADib: TDibImage);
Var
  Bits: pointer;
Begin
  If ADib.IsImageAssigned Then
    Try
      Bits := ADib.Bits;
      TCMProfile.ConvertInRGBFromLab(Bits, ADib.Width * ADib.Height, 4, hDisplayProfile);
    Except
      FrmMain.siLangmain.ShowMessage('Profile to Profile conversion Failed');
    End;
End;

Procedure TfrmMain.ConvertRGBToDib(Var ADib: TDibImage);
Var
  Bits: pointer;
Begin
  If ADib.IsImageAssigned Then
    Try
      Bits := ADib.Bits;
      TCMProfile.ConvertInLabFromRGB(Bits, ADib.Width * ADib.Height, 4, hPPaperProfile);
    Except
      FrmMain.siLangmain.ShowMessage('Profile to Profile conversion Failed');
    End;
End;

{*------------------------------------------------------------------------------
  set right proportion of yarn elements
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure ReproportionnePredominance;
Var
  MinV, Maxv, i, j: integer;
Begin
  Minv := 10000;
  MaxV := 0;
  For i := 0 To nbcolorfil - 1 Do
    Begin
      If tabpred[i] > Maxv Then
        MaxV := tabpred[i];
      If tabpred[i] < Minv Then
        Minv := tabpred[i];
    End;
  If MaxV < 15 Then
    exit;
  For i := 0 To nbcolorfil - 1 Do
    Begin
      If tabpred[i] > 0 Then
        Begin
          tabpred[i] := Round(tabpred[i] * 15 / MaxV);
          If tabpred[i] = 0 Then
            tabpred[i] := 1;
        End;
    End;
End;

{*------------------------------------------------------------------------------
  Create yarn
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure TfrmMain.CreerFil;
Var
  epaisseur, i, j, k, l, colorlength, sens, typ, icoul, rap_multico, rap_space, taille, som_tabpred, tc: integer;
  diametre, diam_calc, longueur, long_calc, torsion, resol: Single;
  imgpetite, imggrande, imgfiltmp, tmpDib, tmp2: TDIBimage;
  p32, p1, p2: PScanline32;
  p24: PScanline24;
  tabnbalea: Array Of integer;
  TitreNm: extended;
  Largeur, facteur: integer;            //la largeur du fil en largeur x 20 (taille dans laquelle le fil sera sauvé)
  bmp: TBitmap;
  Melangevariable: TDibImage;
  count, LargeurEnPlus, MaxWidth, Section: integer;
  MemStream: TMemorystream;
  Couleur: TColor;
Begin
  Try
    Screen.Cursor := CrAniGlobe;
    If TypeDeFil < 0 Then
      Begin
        TypeDeFil := 0;                 //securité si rien n'est selectionné
        PreviousTypeDeFil := TypeDeFil;
        UpdateTypeDeFil;
      End;
    nbcolorfil := Trunc(fil_nbcolor.Value);
    For i := 0 To nbcolorfil - 1 Do
      tabpred[i] := StrToIntDef(ColorGrid.Cells[2, i + 1], 1);
    //pour eviter d'avoir des raports de chinés trop long
    ReproportionnePredominance;
    //Limatation du relief pour les coloris extremes
    LimitRelief;

    PercentContrast := ContrastEdit.value;
    SquareSize := Edit_taillecarreau.Value;

    //la largeur en plus du fil pour l'irregulatité
    LargeurEnPlus := DiametreEdit.Value * 2;

    ZoomValue := TheZoom / 100;
    resol := 200 / ResolutionW * ZoomValue;
    //le titre en Nm
    Case titreCb.ItemIndex Of
      0: TitreNm := fil_titre.value;    //Nm
      1: TitreNm := 1000 / fil_titre.value; //Tex
      2: TitreNm := 10000 / fil_titre.value; //Tex
      3: TitreNm := 9000 / fil_titre.value; //Denier
      4: TitreNm := fil_titre.value * 61 / 36; //Ne
    End;                                //of case
    //On tient compte du nombre de bruns
    TitreNm := TitreNm / RetorsEdit.value;
    diametre := sqrt(1000 / TitreNm) / fil_densite.Value * resol;
    longueur := 1000 / fil_torsion.value * resol;
    diam_calc := 5 * diametre;
    long_calc := 5 * longueur;
    torsion := long_calc / (diam_calc * 1.5);

    rap_multico := StrToInt(multico_rapport.text);
    rap_space := StrToInt(space_rapport.text);
    sens := RG_sens_fil.ItemIndex;

    {// si il n'y a qu'une couleur alors alors on fait comme si ç'était un chiné
    if fil_nbcolor.Value = 1 then
      TypeDeFil := 1;
    application.ProcessMessages;   }

    typ := TypeDeFil;

    If (typ <> 5) And (typ <> 6) And (typ <> 8) Then
      Begin
        CentreTop := 0;
        CentreBottom := 19;
      End
    Else
      Begin
        CentreTop := Min(YTop, YBottom);
        CentreBottom := Max(YTop, YBottom);
      End;

    If imgfil <> Nil Then
      FreeAndNil(imgfil);
    imgfil := TDIBimage.Create;
    imgfiltmp := TDIBimage.Create;
    If (typ < 5) Then
      Begin
        If imgfilExport <> Nil Then
          FreeAndNil(imgfilExport);
        imgfilexport := TDIBimage.Create;
      End;

    nbcoloralea := 0;
    For i := 0 To nbcolorfil - 1 Do
      inc(nbcoloralea, tabpred[i]);
    SetLength(tabcoloralea, nbcoloralea);
    k := 0;
    For i := 0 To nbcolorfil - 1 Do
      For j := 0 To tabpred[i] - 1 Do
        Begin
          tabcoloralea[k] := tabcolor[i];
          inc(k);
        End;
    som_tabpred := k;

    LargeurEnPlus := 0;

    Case typ Of

      0:                                // mouliné
        Begin
          taille := round(long_calc * nbcolorfil);
          Largeur := round(taille * (20 + LargeurEnPlus) / diam_calc);

          //sil il ya une torsion
          If (sens < 2) Then
            Begin
              imgfilexport.NewDIB(Largeur, (20 + LargeurEnPlus), 32, 0, 0, 0, false, Nil);
              i := 0;
              k := 0;
              While i < imgfilexport.Width Do
                Begin
                  //pb.Position := i * 100 div imgfilexport.width;
                  icoul := k Mod nbcolorfil;
                  CreateBasicYarnShape(imgfilexport, i, round(Largeur * tabpred[icoul] / som_tabpred), (20 + LargeurEnPlus), icoul, torsion, trunc(fil_nbcolor.Value));
                  inc(i, round(Largeur * tabpred[icoul] / som_tabpred));
                  inc(k);
                End;
            End
          Else
            Begin
              imgfilexport.NewDIB(100, (20 + LargeurEnPlus), 32, 0, 0, 0, false, Nil);
              imgfilexport.FillRect32s(-2, -2, imgfilexport.Width, imgfilexport.Height, TabColor[0].Moyen);
            End;
          //on stocke le fil sans torsion au cas ou on veut l'utiliser pour creer un fil mouliné de bouts chinés
          If FilSansRelief <> Nil Then
            FreeAndNil(FilSansRelief);
          FilSansRelief := TDibImage.Create;
          FilSansRelief.CopyFromDib(imgfilexport);
          ActExportImage.enabled := True;
        End;

      1:                                // mélangé
        Begin
          PreparePredominantlyMelange(EspacementEdit.Value);

          taille := round(long_calc);
          Largeur := round(taille * (20 + LargeurEnPlus) / diam_calc);
          imgfilexport.NewDIB(Largeur * NBSectionMelange, (20 + LargeurEnPlus), 32, 0, 0, 0, false, Nil);
          //sil il ya une torsion
          If (sens < 2) Then
            Begin
              //on dessine les portions de mouliné avec pour chaque portion ses predominance
              i := 0;
              count := -1;
              While i < imgfilexport.Width Do
                Begin
                  inc(count);
                  PredominantMelange(count);
                  //pb.Position := i * 100 div imgfilexport.width;
                  If FibreCB.Checked Then
                    CreateMelangeYarn(imgfilexport, i, Largeur, (20 + LargeurEnPlus), sens, SquareSize, torsion)
                  Else
                    CreateMelangeYarn2(imgfilexport, i, Largeur, (20 + LargeurEnPlus), sens, SquareSize, torsion);
                  inc(i, Largeur);
                End;
            End
          Else
            imgfilexport.FillRect32s(-2, -2, imgfilexport.Width, imgfilexport.Height, TabColor[0].Moyen);
          //on stocke le fil sans torsion au cas ou on veut l'utiliser pour creer un fil mouliné de bouts chinés
          If FilSansRelief <> Nil Then
            FreeAndNil(FilSansRelief);
          FilSansRelief := TDibImage.Create;
          FilSansRelief.CopyFromDib(imgfilexport);
          ActExportImage.enabled := True;

          // on applique les torsions si il y en a
          If rg_sens_fil.ItemIndex < 2 Then
            Begin
              i := 0;
              While i < imgfilexport.Width Do
                Begin
                  //pb.Position := i * 100 div imgfilexport.width;
                  TwistMelange(imgfilexport, i, Largeur, (20 + LargeurEnPlus), sens, torsion);
                  inc(i, Largeur);
                End;
            End;
          SetLength(Melangetabpred, 0, 0);
          sb.Panels[0].Progress.Position := 0;
        End;

      2:                                // multico
        Begin
          SetLength(tabnbalea, nbcolorfil * rap_multico);
          k := 0;
          For i := 0 To nbcolorfil * rap_multico - 1 Do
            Begin
              tabnbalea[i] := Random(multico_random_max.value) + 1;
              inc(k, tabnbalea[i] * tabpred[i Mod nbcolorfil]);
            End;

          taille := 3 * round(k * long_calc);
          Largeur := round(taille * (20 + LargeurEnPlus) / diam_calc);
          imgfilexport.NewDIB(Largeur, (20 + LargeurEnPlus), 32, 0, 0, 0, false, Nil);
          //sil il ya une torsion
          If (sens < 2) Then
            Begin
              i := 0;

              While i < imgfilexport.Width Do
                Begin
                  //sb.Panels[0].Progress.Position := i * 100 div imgfilexport.width;
                  For k := 0 To nbcolorfil * rap_multico - 1 Do
                    Begin
                      icoul := k Mod nbcolorfil;
                      l := 0;

                      While l < tabnbalea[k] * tabpred[icoul] * long_calc Do
                        Begin
                          CreateBasicYarnShape(imgfilexport, i + l, round(long_calc * tabpred[icoul] * (20 + LargeurEnPlus) / diam_calc), (20 + LargeurEnPlus), icoul, torsion, trunc(fil_nbcolor.Value));
                          inc(l, round(tabpred[icoul] * long_calc * (20 + LargeurEnPlus) / diam_calc));
                        End;
                      inc(i, l);
                    End;
                End;
            End
          Else
            imgfilexport.FillRect32s(-2, -2, imgfilexport.Width, imgfilexport.Height, TabColor[0].Moyen);
          //on stocke le fil sans torsion au cas ou on veut l'utiliser pour creer un fil mouliné de bouts chinés
          If FilSansRelief <> Nil Then
            FreeAndNil(FilSansRelief);
          FilSansRelief := TDibImage.Create;
          FilSansRelief.CopyFromDib(imgfilexport);
          ActExportImage.enabled := True;
        End;

      3:                                // space
        Begin
          k := 0;
          For i := 0 To nbcolorfil - 1 Do
            inc(k, rap_space * tabpred[i] + 1);
          taille := round(k * long_calc);
          Largeur := round(taille * (20 + LargeurEnPlus) / diam_calc);
          imgfilexport.NewDIB(Largeur, (20 + LargeurEnPlus), 32, 0, 0, 0, false, Nil);

          //sil il ya une torsion
          If (sens < 2) Then
            Begin
              i := 0;
              k := 0;

              While i < imgfilexport.Width Do
                Begin
                  //sb.Panels[0].Progress.Position := i * 100 div imgfilexport.width;
                  icoul := k Mod nbcolorfil;

                  If sens = 2 Then
                    Begin
                      CreateStraightSpacedYarn(imgfilexport, i, round(rap_space * long_calc * (20 + LargeurEnPlus) / diam_calc * tabpred[icoul]), (20 + LargeurEnPlus), icoul);
                      CreateSpacedYarnStraightTransition(imgfilexport, i + round(rap_space * long_calc * (20 + LargeurEnPlus) / diam_calc * tabpred[icoul]), round(long_calc * (20 + LargeurEnPlus) /
                        diam_calc),
                        (20 + LargeurEnPlus), icoul, (k + 1) Mod nbcolorfil)
                    End
                  Else
                    Begin
                      For j := 1 To rap_space * tabpred[icoul] Do
                        CreateBasicYarnShape(imgfilexport, i + round(long_calc * (20 + LargeurEnPlus) / diam_calc * (j - 1)), round(long_calc * (20 + LargeurEnPlus) / diam_calc), (20 + LargeurEnPlus),
                          icoul, torsion, trunc(fil_nbcolor.Value));
                      CreateSpacedYarnTransition(imgfilexport, i + round(rap_space * long_calc * (20 + LargeurEnPlus) / diam_calc * tabpred[icoul]), round(long_calc * (20 + LargeurEnPlus) / diam_calc), (20
                        +
                        LargeurEnPlus), icoul, (k + 1) Mod nbcolorfil, torsion);
                    End;

                  inc(k);
                  inc(i, round((rap_space * tabpred[icoul] + 1) * long_calc * (20 + LargeurEnPlus) / diam_calc));
                End;
            End
          Else
            imgfilexport.FillRect32s(-2, -2, imgfilexport.Width, imgfilexport.Height, TabColor[0].Moyen);

          //on stocke le fil sans torsion au cas ou on veut l'utiliser pour creer un fil mouliné de bouts chinés
          If FilSansRelief <> Nil Then
            FreeAndNil(FilSansRelief);
          FilSansRelief := TDibImage.Create;
          FilSansRelief.CopyFromDib(imgfilexport);
          ActExportImage.enabled := True;
        End;

      4:                                // mouliné / Chine
        Begin
          MaxWidth := 0;
          For i := 0 To nbcolorfil - 1 Do
            Begin
              If ImageFil[i] = Nil Then
                Begin
                  FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_0' (* 'Affecter toutes les images de fil avant de lancer la création du fil composé' *)));
                  exit;
                End;
              If ImageFil[i].Width > MaxWidth Then
                MaxWidth := ImageFil[i].Width;
            End;
          taille := round(long_calc * nbcolorfil);
          Largeur := round(taille * (20 + LargeurEnPlus) / diam_calc);
          Section := Largeur;

          If Largeur < MaxWidth Then
            Largeur := Round(MaxWidth / Largeur) * Largeur;
          imgfilexport.NewDIB(Largeur, (20 + LargeurEnPlus), 32, 0, 0, 0, false, Nil);

          i := 0;
          k := 0;

          For i := 0 To nbcolorfil - 1 Do
            Begin
              If ImageTwisted[i] <> Nil Then
                FreeAndNil(ImageTwisted[i]);
              ImageTwisted[i] := TDibImage.Create;
              ImageTwisted[i].CopyFromDib(ImageFil[i]);
              //Si la torsion finale est Z alors on txiste la torsion initiale
              If {(sens = 1) and }(SensFil[i] = 1) Then
                ImageTwisted[i].Flip;
            End;

          While i < imgfilexport.Width Do
            Begin
              //sb.Panels[0].Progress.Position := i * 100 div imgfilexport.width;
              icoul := k Mod nbcolorfil;
              CreerFilMoulineChine(imgfilexport, i, round(Section * tabpred[icoul] / som_tabpred), (20 + LargeurEnPlus), icoul, torsion);
              inc(i, round(Section * tabpred[icoul] / som_tabpred));
              inc(k);
            End;
          //on stocke le fil sans torsion au cas ou on veut l'utiliser pour creer un fil mouliné de bouts chinés
          If FilSansRelief <> Nil Then
            FreeAndNil(FilSansRelief);
          FilSansRelief := TDibImage.Create;
          FilSansRelief.CopyFromDib(imgfilexport);
          ActExportImage.enabled := True;
        End;

      5, 6:                             // fantaisie
        Begin
          If imgfilexport = Nil Then
            imgfilexport := TDibImage.Create;
          If imgfilfantaisieimport <> Nil Then
            Begin
              If imgfilfantaisieimport.IsImageAssigned Then
                imgfilexport.CopyFromDib(imgfilfantaisieimport) // si on a ouvert un .fil avec un fil fantaisie
              Else
                Begin
                  FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_1' (* 'Concevoir le fil fantaisie avant de lancer la simulation' *)));
                  exit;
                End;
            End
          Else                          {// sinon l'utilisateur choisi une image de fil a charger}
            If OPD.execute Then
              Begin
                imgfiltmp.LoadFromFile(OPD.filename, 0);
                imgfilexport.NewDIB(imgfiltmp.width, imgfiltmp.height, 32, 0, 0, 0, false, Nil);
                For j := 0 To imgfilexport.Height - 1 Do
                  Begin
                    p24 := imgfiltmp.ScanLine[j];
                    p32 := imgfilexport.ScanLine[j];
                    For i := 0 To imgfilexport.Width - 1 Do
                      Begin
                        p32[i].rgbRed := p24[i].Red;
                        p32[i].rgbGreen := p24[i].Green;
                        p32[i].rgbBlue := p24[i].Blue;
                        If (p24[i].Red = 0) And (p24[i].Green = 0) And (p24[i].Blue = 0) Then
                          p32[i].rgbReserved := 0
                        Else
                          p32[i].rgbReserved := 255;
                      End;
                  End;
                ConvertRGBToDib(imgfilexport);
                If sourceDib <> Nil Then
                  freeAndNil(sourceDib);
                SourceDib := TDibImage.Create;
                SourceDib.CopyFromDib(imgfilexport);
                taille := imgfilexport.Width;
              End
            Else
              exit;
          If imgfilexport = Nil Then
            Begin
              FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_1' (* 'Concevoir le fil fantaisie avant de lancer la simulation' *)));
              exit;
            End
          Else
            Begin
              If Not imgfilexport.IsImageAssigned Then
                Begin
                  FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_1' (* 'Concevoir le fil fantaisie avant de lancer la simulation' *)));
                  exit;
                End;
            End;
          //on stocke le fil sans torsion au cas ou on veut l'utiliser pour creer un fil mouliné de bouts chinés
          If FilSansRelief <> Nil Then
            FreeAndNil(FilSansRelief);
          FilSansRelief := TDibImage.Create;
          FilSansRelief.CopyFromDib(imgfilexport);
          ActExportImage.enabled := True;
        End;                            // end case 4

      7, 8:
        If imgfilexport = Nil Then
          Begin
            FrmMain.siLangmain.ShowMessage('Concevoir le fil space dye avant de lancer la simulation');
            exit;
          End

    End;                                // endcase

    If (sens = 1) And (typ < 5) Then    // sens Z
      ChangeYarnDirection(imgfilexport);

    //effet de gradient
    If (typ < 5) Then
      Begin
        LimitRondeur;
        GradientYarn(imgfilexport, rondeurEdit.Value, GradientYarnFromCenter);
      End;

    {if LargeurEnPlus > 0 then
      OndulationFil(imgfilexport);}

    //on va irregulariser le fil dans les cas des moulinées et des chines
    If ((typ = 0) Or (typ = 1)) And (fil_irregularite_Longueur.Value > 0) Then
      Begin
        facteur := F_Irregularite.Value;
        MakeYarnIrregulated(imgfilexport, fil_irregularite_Longueur.Value, facteur);
        If FilSansRelief <> Nil Then
          MakeYarnIrregulated(FilSansRelief, fil_irregularite_Longueur.Value, facteur);
        F_Irregularite.Value := facteur;
      End;

    //On va donner un effet stone washed au fil
    If (typ <> 8) And DelavableCB.Checked And (PrelavageEdit.Value > 0) Then
      Begin
        FadeYarn(imgfilexport, PrelavageEdit.Value);
      End;

    If (typ < 5) And (nbcolorfil > 1) Then
      Begin
        Try
          tmp2 := TDibImage.Create;
          tmp2.SetSize32(imgfilexport.Width, imgfilexport.Height);
          If (typ <> 1) Then
            SpreadYarnr(imgfilexport, tmp2, 6, 6)
          Else
            SpreadYarnr(imgfilexport, tmp2, 10, 10);
          imgfilexport.CopyFromDib(tmp2);

        Finally
          tmp2.Free;
        End;
      End;

    Try
      imageEn1.BackGround := Color;
      tmpDib := TDibImage.Create;
      tmpDib.CopyFromDib(imgfilexport);

      ConvertDibToDisplay(tmpDib);
      MemStream := TMemorystream.Create;
      MemStream.SetSize(tmpDib.ImageFileSize);
      tmpDib.SaveToStream(MemStream, true);
      MemStream.Position := 0;
      imageEn1.IO.LoadFromStream(MemStream);
      imageEn1.FitToHeight;
      YarnToolbar.Visible := true;
    Finally
      If tmpDib <> Nil Then
        FreeAndnil(tmpDib);
      If MemStream <> Nil Then
        FreeAndnil(MemStream);
    End;

    Couleur := GetMediumColor(imgfilexport, Trunc(fil_nbcolor.Value), TypeDeFil); //couleur RGB moyen du fil
    colorpanel.Color := Couleur;

    ActSaveAs.Enabled := true;
    ActSave.Enabled := true;
    ActPrint.Enabled := true;
    ActAdjustColor.Enabled := true;

  Finally
    Screen.cursor := CrDefault;
  End;
End;

{*------------------------------------------------------------------------------
  limit yarn relief
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure TfrmMain.LimitRelief;
Var
  i, PercentContrast: integer;
Begin
  If CompressReliefCB.Checked And (TypeDeFil < 4) Then
    Begin
      PercentContrast := ContrastEdit.value;

      //Relief max d'eclaircissement
      For i := 0 To (nbcolorfil - 1) Do
        Begin
          If 255 - tabcolor[i].moyen.rgbBlue < PercentContrast Then
            PercentContrast := 255 - tabcolor[i].moyen.rgbBlue;
        End;

      //Relief max d'assombrissement
      For i := 0 To (nbcolorfil - 1) Do
        Begin
          If tabcolor[i].moyen.rgbBlue < PercentContrast Then
            PercentContrast := tabcolor[i].moyen.rgbBlue;
        End;

      If ContrastEdit.value <> PercentContrast Then
        Begin
          ContrastEdit.value := PercentContrast;
          ContrastEditChange(Nil);
        End;
    End;
End;

{*------------------------------------------------------------------------------
  limit roundness
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure TfrmMain.LimitRondeur;
Var
  i, j, PercentRondeur: integer;
Begin
  If CompressRondeurCB.Checked And (TypeDeFil < 4) Then
    Begin
      PercentRondeur := RondeurEdit.value;

      //Relief max d'eclaircissement
      For i := 0 To (nbcolorfil - 1) Do
        Begin
          //Clair
          If 255 - tabcolor[i].Clair.rgbBlue < PercentRondeur Then
            PercentRondeur := 255 - tabcolor[i].Clair.rgbBlue;
          //Moyen
          If 255 - tabcolor[i].Moyen.rgbBlue < PercentRondeur Then
            PercentRondeur := 255 - tabcolor[i].moyen.rgbBlue;
          //Foncé
          If 255 - tabcolor[i].Fonce.rgbBlue < PercentRondeur Then
            PercentRondeur := 255 - tabcolor[i].Fonce.rgbBlue;
        End;

      //Relief max d'assombrissement
      For i := 0 To (nbcolorfil - 1) Do
        Begin
          //Clair
          If tabcolor[i].Clair.rgbBlue < PercentRondeur Then
            PercentRondeur := tabcolor[i].Clair.rgbBlue;
          //Moyen
          If tabcolor[i].moyen.rgbBlue < PercentRondeur Then
            PercentRondeur := tabcolor[i].moyen.rgbBlue;
          //Foncé
          If tabcolor[i].Fonce.rgbBlue < PercentRondeur Then
            PercentRondeur := tabcolor[i].Fonce.rgbBlue;
        End;

      If RondeurEdit.value <> PercentRondeur Then
        RondeurEdit.value := PercentRondeur;
    End;
End;

Procedure TFrmMain.ContrastEditChange(Sender: TObject);
Var
  i: integer;
Begin
  If ContrastEdit.Text <> '' Then
    Begin
      PercentContrast := ContrastEdit.value;
      For i := 0 To Trunc(fil_nbcolor.Value) - 1 Do
        Begin                           // pour chaque couleur on affecte les valeurs resultantes du contraste
          tabcolor[i].clair := Eclaircir(tabcolor[i].moyen, PercentContrast); // on calcule les couleurs + clair
          tabcolor[i].fonce := Assombrir(tabcolor[i].moyen, PercentContrast); // et + fonce
        End;
    End;
End;

// ===================================================================================================================================================
// ===================================================================================================================================================

{*------------------------------------------------------------------------------
  Save yarn to stream
  @param FS   ParameterDescription
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure TfrmMain.SauverFilToStream(FS: TStream);
Var
  i, a: integer;
  j: single;
  c: currency;
  b, version, LL, aa, bb: byte;
  s: String;
  New, fibre, b2: boolean;
  filHeader: TfilHeader;
  Couleur: Tcolor;
  bmp: TBitmap;
  FancyDib: TDibImage;
  CoulLab: EncodedLabValue;
  SimulBitmap: TBitmap;
  tmpDib: TDibImage;

  Procedure SauverTexte(Texte: String);
  Begin
    s := Texte;
    i := length(s);
    fs.WriteBuffer(i, sizeof(i));
    fs.WriteBuffer(s[1], i);
  End;

Begin
  Try
    SavingYarnToFile := true;
    SimulBitmap := TBitmap.Create;
    SimulBitmap.PIXELFORMAT := Pf32bit;
    SimulBitmap.Width := 320;
    SimulBitmap.Height := 320;
    //Magic Number
   // write header
    With filHeader Do
      Begin
        Magic := 'FILFILEFORMAT';
        Width := 123;
        Height := 123;
      End;

    fs.Write(filHeader, sizeof(TfilHeader));

    version := vers;                    //n° de version
    fs.WriteBuffer(version, 1);
    New := true;                        // tag de nouveau tissu
    fs.WriteBuffer(New, sizeof(New));
    Couleur := GetMediumColor(imgfilexport, Trunc(fil_nbcolor.Value), TypeDeFil); //couleur RGB moyen du fil

    colorpanel.Color := Couleur;
    fs.WriteBuffer(Couleur, sizeof(Couleur));
    //Depuis la version 11
    GetMediumColorInLab(imgfilexport, LL, aa, bb, Trunc(frmMain.fil_nbcolor.Value), TypeDeFil);
    CoulLab.L := LL;
    CoulLab.a := aa;
    CoulLab.b := bb;
    fs.WriteBuffer(CoulLab, sizeof(CoulLab));

    //on va sauver une icone representant l'echantillon de tissu
    If Not NoNeedToSimulForThumb Then
      Begin
        Try
          CreatingThumb := true;
          Try
            TypeDeSimulation := 6;
            SimulDPI := 100;
            DM.DoSimulation(10, 0, 0, true, 0, 0, false);
          Except
            Try
              DM.Create2DThumb;
              //FinalSimuler(YarnSize, 0, 0, true, 0, 0, false);
            Except
              If thumb <> Nil Then
                FreeAndNil(Thumb);
              SiLangShowCleverMessage('Création de la vignette impossible');
            End;
          End;
        Finally
          Creatingthumb := false;
        End;

        If thumb <> Nil Then
          Begin
            If thumb.IsImageAssigned Then
              Begin
                Try
                  tmpDib := TDibImage.Create;
                  bmp := Tbitmap.Create;
                  tmpDib.NewDIB(32, 32, 32, 0, 0, 0, false, Nil);
                  NewResampleDibToDib(Thumb, tmpDib, true);
                  tmpDib.SaveToBitmap(bmp);
                  bmp.Canvas.Brush.Color := clWhite;
                  bmp.Canvas.Pen.Color := clWhite;
                  bmp.Canvas.Polygon([Point(bmp.Width, 0), Point(bmp.Width - 8, 0),
                    Point(bmp.Width - 1, 7), Point(bmp.Width - 1, 0)]);
                  bmp.Canvas.Brush.Color := rgb(81, 151, 227);
                  bmp.Canvas.Pen.Color := rgb(81, 151, 227);
                  bmp.Canvas.Polygon([Point(bmp.Width - 8, 0), Point(bmp.Width - 1,
                      7),
                    Point(bmp.Width - 8, 7), Point(bmp.Width - 8, 0)]);
                  bmp.SaveToStream(fs);
                Finally
                  bmp.Free;
                  tmpDib.Free;
                End;
                Try
                  tmpDib := TDibImage.Create;
                  tmpDib.SetSize24(Thumb.Width, Thumb.Height);
                  tmpDib.Draw(Thumb);
                  tmpDib.SaveToStream(fs, true);
                Finally
                  tmpDib.Free;
                End;
              End
            Else
              Begin
                Try
                  tmpDib := TDibImage.Create;
                  tmpDib.LoadFromBitmap(Preview32);
                  tmpDib.SaveToStream(fs, true);
                  tmpDib.FreeDib;
                  tmpDib.LoadFromBitmap(Preview123);
                  tmpDib.SaveToStream(fs, true);
                Finally
                  tmpDib.Free;
                End;
              End;
          End
        Else
          Begin
            Try
              tmpDib := TDibImage.Create;
              tmpDib.LoadFromBitmap(Preview32);
              tmpDib.SaveToStream(fs, true);
              tmpDib.FreeDib;
              tmpDib.LoadFromBitmap(Preview123);
              tmpDib.SaveToStream(fs, true);
            Finally
              tmpDib.free;
            End;
          End;
        If thumb <> Nil Then
          FreeAndNil(Thumb);
      End
    Else
      Begin
        If thumb <> Nil Then
          FreeAndNil(Thumb);
        Try
          tmpDib := TDibImage.Create;
          tmpDib.SetSize32(32, 32);
          tmpDib.SaveToStream(fs, true);
          tmpDib.FreeDib;
          tmpDib.SetSize32(123, 123);
          tmpDib.SaveToStream(fs, true);
        Finally
          If tmpDib <> Nil Then
            FreeAndNil(tmpDib);
        End;
      End;

    //Depuis la version 19 on dit si le fil est compressé ou non (pas compressé pour la simul
    fs.WriteBuffer(NoNeedToSimulForThumb, 1);

    //Code Matiere
    s := CodeMatiere;
    i := length(s);
    fs.WriteBuffer(i, sizeof(i));
    fs.WriteBuffer(s[1], i);
    //Code Couleur
    s := CodeCouleur;
    i := length(s);
    fs.WriteBuffer(i, sizeof(i));
    fs.WriteBuffer(s[1], i);
    //Description
    s := Description;
    i := length(s);
    fs.WriteBuffer(i, sizeof(i));
    fs.WriteBuffer(s[1], i);

    b := TypeDeFil;                     // type de fil
    fs.WriteBuffer(b, 1);
    b := RG_sens_fil.ItemIndex;         // sens du fil
    fs.WriteBuffer(b, 1);
    j := fil_titre.value;               // titre du fil
    fs.WriteBuffer(j, sizeof(j));
    i := TitreCb.ItemIndex;             // unite du fil
    fs.WriteBuffer(i, sizeof(i));
    //Depuis la version 20
    i := Trunc(RetorsEdit.value);       // nombre de bruns
    fs.WriteBuffer(i, sizeof(i));
    //
    j := fil_densite.value;             // coef matiere
    fs.WriteBuffer(j, sizeof(j));
    i := fil_torsion.value;             // torsion du fil
    fs.WriteBuffer(i, sizeof(i));
    i := fil_irregularite_Longueur.value; // irregularite du fil en longueur
    fs.WriteBuffer(i, sizeof(i));

    i := F_Irregularite.value;          // facteur d'irregularite du fil
    fs.WriteBuffer(i, sizeof(i));

    i := DiametreEdit.value;            // irregularite du fil en largueur
    fs.WriteBuffer(i, sizeof(i));

    i := 0;                             //ondulation.value;              // ondulation du fil
    fs.WriteBuffer(i, sizeof(i));

    i := ContrastEdit.value;            // relief du fil
    fs.WriteBuffer(i, sizeof(i));
    i := RangeEdit.value;               // nombre de rangee de la compo du fil
    fs.WriteBuffer(i, sizeof(i));
    For a := 0 To RangeEdit.value - 1 Do
      Begin
        i := strtointdef(GrilleMatiere.cells[0, a + 1], 100); // % de la matiere
        fs.WriteBuffer(i, sizeof(i));
        s := GrilleMatiere.cells[1, a + 1];
        i := length(s);
        fs.WriteBuffer(i, sizeof(i));
        fs.WriteBuffer(s[1], i);
      End;

    c := fil_prix.value;                // prix du fil
    fs.WriteBuffer(c, sizeof(c));

    i := RondeurEdit.value;             // rondeur du fil
    fs.WriteBuffer(i, sizeof(i));

    i := DiametreEdit.value;            // variation de diametre du fil
    fs.WriteBuffer(i, sizeof(i));

    i := PilositeEdit.value;            // longueur des poils du fil
    fs.WriteBuffer(i, sizeof(i));

    i := DensPoilEdit.value;            // densite des poils du fil
    fs.WriteBuffer(i, sizeof(i));

    //depuis la version 10 on sauve la position du milieu
    fs.WriteBuffer(CentreTop, sizeof(CentreTop));
    fs.WriteBuffer(CentreBottom, sizeof(CentreBottom));

    b := space_rapport.Value;           // rapport du space
    fs.WriteBuffer(b, 1);
    b := multico_rapport.Value;         // temps de repetition du multico
    fs.WriteBuffer(b, 1);
    b := multico_random_max.Value;      // longueur aleatoire maximum du multico
    fs.WriteBuffer(b, 1);

    b2 := DelavableCB.Checked;          // fil delavable
    fs.Write(b2, sizeof(b2));

    b := Trunc(fil_nbcolor.Value);      // nonbre de couleur
    fs.WriteBuffer(b, sizeof(byte));

    For i := 0 To Trunc(fil_nbcolor.Value) - 1 Do
      Begin                             // pour chaque couleur
        b := tabcolor[i].moyen.rgbRed;  // composante rouge
        fs.WriteBuffer(b, 1);
        b := tabcolor[i].moyen.rgbGreen; // composante verte
        fs.WriteBuffer(b, 1);
        b := tabcolor[i].moyen.rgbBlue; // composante bleue
        fs.WriteBuffer(b, 1);
        b := tabpred[i];                // predominance
        fs.WriteBuffer(b, 1);
        //Nom de la couleur depuis la version 18
        SauverTexte(NomCouleur[i]);
      End;

    If TypeDeFil = 1 Then
      i := Edit_TailleCarreau.Value
    Else
      i := 1;
    fs.WriteBuffer(i, sizeof(i));       // si melange on sauve la taille des carreaux sinon champs libre

    If TypeDeFil = 1 Then
      i := EspacementEdit.Value
    Else
      i := 1;
    fs.WriteBuffer(i, sizeof(i));       // si melange on sauve l'espacement du melange

    fibre := fibrecb.Checked;           //si les fibres du melangé sont orientées
    fs.WriteBuffer(fibre, sizeof(fibre));

    fs.WriteBuffer(i, sizeof(i));       // 4 champs libre de la taille d'un entier
    fs.WriteBuffer(i, sizeof(i));
    fs.WriteBuffer(i, sizeof(i));
    fs.WriteBuffer(i, sizeof(i));

    //Donnée technique depuis la version 15
    SauverTexte(CodeBarre);
    SauverTexte(DesignationCouleur);
    SauverTexte(CodeRecetteTeinture);
    SauverTexte(DesignationRecetteTeinture);
    SauverTexte(TitrageCommercial);
    SauverTexte(CodeComposition);
    SauverTexte(DesignationComposition);
    SauverTexte(CodeFammile);
    SauverTexte(DesignationFamille);
    SauverTexte(Observations);
    SauverTexte(CodeUniteStockage);
    SauverTexte(DesignationUniteStockage);
    SauverTexte(CodeStructure);
    SauverTexte(DesignationStructure);
    SauverTexte(ObservationTechnique);
    SauverTexte(TypeArticle);
    fs.WriteBuffer(FilTeint, sizeof(FilTeint));
    ////////////////////

//    if SavingYarnToFile then
    If NoNeedToSimulForThumb Then
      Begin                             //on passe directement la bitmap
        YarnBitmap.CopyFromDib(imgfilexport);
        StoreYarnOnDisk := false;
      End
    Else
      SaveBitmapToCompressedStream(imgfilexport, fs, NoNeedToSimulForThumb);

    // si c'est juste pour la simul on ne sauve pas le reste
    If NoNeedToSimulForThumb Then
      exit;

    {    else
          begin
            TMemorystream(fs).SetSize(fs.Size + imgfilexport.ImageFileSize);
            imgfilexport.SaveToStream(fs, true); // l'image !
          end;    }

        //si c'est un mouliné composé de chiné alors on sauve les chines
    If TypeDeFil = 4 Then
      Begin
        For i := 0 To Trunc(fil_nbcolor.Value) - 1 Do //if ImageFil[i] <> nil then
          Begin
            //je fais cette procedure à la con sinon ça sauve mal mais me demande pas pourquoi (tableau de Dib)
            Try
              If SavingYarnToFile Then
                //SaveDibToPNGStream(ImageFil[i], fs)  avant version 14
                SaveBitmapToCompressedStream(ImageFil[i], fs)
              Else
                Begin
                  bmp := Tbitmap.Create;
                  ImageFil[i].SaveToBitmap(bmp);
                  bmp.SaveToStream(fs);
                End;
            Finally
              If Not SavingYarnToFile Then
                bmp.free;
            End;
            //ImageFil[i].SaveToStream(fs, true);
          End;
      End;

    //si le fil est un fil scanné on sauve aussi son image origianle rgb, pour d'eventuelle modification ulterieure
    If (TypeDeFil = 5) Or (TypeDeFil = 6) Then
      If SavingYarnToFile Then
        SaveBitmapToCompressedStream(SourceDib, fs)
          //SaveDibToPNGStream(SourceDib, fs) avant version 14
      Else
        Begin
          TMemorystream(fs).SetSize(fs.Size + SourceDib.ImageFileSize);
          SourceDib.SaveToStream(fs, true); // l'image !
        End;

    //depuis la version 14
    If TypeDeFil = 7 Then
      Begin
        fs.writebuffer(NBSection, sizeof(NBSection));
        For i := 0 To NBSection - 1 Do
          Begin
            fs.writebuffer(SectionColor[i], sizeof(SectionColor[i]));
            fs.writebuffer(SectionSize[i], sizeof(SectionSize[i]));
            //Depuis la version 18
            SauverTexte(SectionNomCouleur[i]);
          End;
      End;

    //depuis la version 16
    If TypeDeFil = 8 Then
      Begin
        fs.writebuffer(SlubTitle, sizeof(SlubTitle));
        fs.writebuffer(NBSectionSlub, sizeof(NBSectionSlub));
        fs.writebuffer(StoneContrast, Sizeof(StoneContrast));
        For i := 0 To NBSectionSlub - 1 Do
          Begin
            fs.writebuffer(SectionSlubSize[i], sizeof(SectionSlubSize[i]));
            fs.writebuffer(sectionSlubPentG[i], sizeof(sectionSlubPentG[i]));
            fs.writebuffer(sectionSlubPentD[i], sizeof(sectionSlubPentD[i]));
          End;
        //depuis la version 21
        For i := 0 To NBSectionSlub - 1 Do
          Begin
            fs.writebuffer(SlubSectionColor[i], sizeof(SlubSectionColor[i]));
            SauverTexte(SlubSectionNomCouleur[i]);
          End;
      End;

    //depuis la version 17
    prelavage := prelavageEdit.Value;
    fs.writebuffer(Prelavage, sizeof(Prelavage));
  Finally
    If SimulBitmap <> Nil Then
      FreeAndNil(SimulBitmap);
  End;
End;

// ===================================================================================================================================================
// ===================================================================================================================================================

{*------------------------------------------------------------------------------
  Save yarn to file
  @param filename   ParameterDescription
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure TfrmMain.SauverFil(filename: String);
Var
  fs: tfilestream;
  i: integer;
Begin
  Try
    Screen.Cursor := CrHourGlass;
    //on va activer la compression maximum des fichiers tiffs

    DoFabricSimulation(true);
    Try
      fs := TFileStream.Create(FileName, fmCreate Or fmShareExclusive);
    Except
      Showmessage('Cannot create the file ' + Filename);
      exit;
    End;
    nbcolorfil := Trunc(fil_nbcolor.Value);
    For i := 0 To nbcolorfil - 1 Do
      tabpred[i] := StrToIntDef(ColorGrid.Cells[2, i + 1], 1);
    SauverFilToStream(fs);
  Finally
    Screen.Cursor := CrDefault;
    SavingYarnToFile := false;
    If fs <> Nil Then
      FreeAndNil(fs);
  End;
  SB.panels[2].Text := extractfilename(filename);
  //Caption := 'fifo-Yarn by fifo (www.fifo.com) | ' + extractfilename(FileName);
 // TBMRUList1.Add(FileName);
End;

{*------------------------------------------------------------------------------
  launch fabric simulation based on this yarn
  @param Sciseaux   ParameterDescription
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure TfrmMain.DoFabricSimulation(Sciseaux: boolean); //Relief en param
Begin
  YarnStream.clear;
  Try
    NoNeedToSimulForThumb := true;
    SauverFilToStream(YarnStream);
    YarnStream.Position := 0;
    If Not loadDataFilFromStream('Yarn Simulation', YarnStream, TbYarnDataWarp[0, 0], WarpColor[0, 0]) Then
      Showmessage('Fil impossible à simuler');

    If Not loadDataFilFromStream('Yarn Simulation', YarnStream, TbYarnDataWeft[0, 0], WeftColor[0, 0]) Then
      Showmessage('Fil impossible à simuler');
    UseCrantage := Sciseaux;
  Finally
    NoNeedToSimulForThumb := false;
  End;
End;

Procedure TFrmMain.FormCloseQuery(Sender: TObject; Var CanClose: Boolean);
Var
  texte: String;
Begin
  If DoNotConfirmClose Then
    Begin
      Canclose := true;
      Exit;
    End;
  If ColorLibModified Then
    Begin
      //if siLangDispatcher1.ActiveLanguage=2 then texte:='Do you want to quit Without saving your color library ?';
      {if siLangDispatcher1.ActiveLanguage=1 then}texte := siLangmain.GetTextOrDefault('IDS_72' (* 'Voulez vous Quitter l''application sans enregistrer la bibliotheque de couleur?' *));
      If FrmMain.siLangmain.MessageDlg(texte, mtConfirmation,
        [mbOk, mbCancel], 0) = mrCancel Then
        CanClose := false
      Else
        Canclose := true;
      exit;
    End;

  //if siLangDispatcher1.ActiveLanguage=2 then texte:='Do you want to quit this software ?';
  {if siLangDispatcher1.ActiveLanguage=1 then}texte := siLangmain.GetTextOrDefault('IDS_73' (* 'Voulez vous Quitter l''application ?' *));
  If FrmMain.siLangmain.MessageDlg(texte, mtConfirmation,
    [mbOk, mbCancel], 0) = mrCancel Then
    CanClose := false
  Else
    Begin
      Canclose := true;
    End;
End;

Procedure TFrmMain.FormControlEditLink1GetEditorValue(Sender: TObject;
  Grid: TAdvStringGrid; Var AValue: String);
Begin
  AValue := MatiereCB.Text;
End;

Procedure TFrmMain.FormControlEditLink1SetEditorFocus(Sender: TObject;
  Grid: TAdvStringGrid; AControl: TWinControl);
Begin
  AControl.SetFocus;
End;

Procedure TFrmMain.FormControlEditLink1SetEditorValue(Sender: TObject;
  Grid: TAdvStringGrid; AValue: String);
Begin
  MatiereCB.Text := AValue;
End;

Procedure TFrmMain.ActQuitExecute(Sender: TObject);
Begin
  close;
End;

Procedure TFrmMain.ActSimulExecute(Sender: TObject);
Var
  tc: integer;
Begin
  If FirstSimul Then
    ActDobbyWizardExecute(Sender);
  If FirstSimul Then
    exit;
  Try
    CreerFil;
    Screen.Cursor := CrAniGlobe;
    DoFabricSimulation(TheZoom = 100);
    Screen.Cursor := CrAniGlobe;
    SimulAborted := false;
    TypeDeSimulation := 0;
    SimulDPI := DisplayDPI;
    DM.DoSimulation(YarnSize, 0, 0, true, StartWarp, StartWeft, false);
  Finally
    Screen.Cursor := CrDefault;
  End;
End;

{*------------------------------------------------------------------------------
  Set fabric parameters
  @param Sender   ParameterDescription
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure TFrmMain.ActDobbyWizardExecute(Sender: TObject);
Var
  i, j: integer;
  NewString: String;
Begin
  DobbyWizardFrom := TDobbyWizardFrom.Create(self);
  If DobbyWizardFrom.Showmodal <> Mrok Then
    exit;
  DataFolder2Use := DataFolder;
  YarnFolder2Use := YarnFolder;
  SetLength(MatCroisement, 0, 0);
  SetLength(MatCroisement, 0);
  IsJacquard := False;
  MatrixLg := nbShafts;
  MatrixHt := nbShafts;
  nbvar := 1;
  WarpPatternLg := 1;
  WeftPatternHT := 1;
  VarCourante := 0;
  NBWarp2use := 1;
  NBWeft2use := 1;
  DentingLg := 4;
  TbDenting[0] := 0;
  TbDenting[1] := 0;
  TbDenting[2] := 1;
  TbDenting[3] := 1;
  DM.SetRepeatDenting;

  TowelMode := False;
  ZoneEpongeHT := 1;
  ZoneEponge[0] := false;
  FrappeLongueHT := 1;
  FrappeLongue[0] := false;
  HauteurBoucleHT := 1;
  HauteurBoucle[0] := 0;

  RegulateurHT := 1;
  NEWRegulator[0] := 0;
  DM.SetRepeatRegulator;
  For i := 0 To MatrixLg - 1 Do
    Drafting[i] := i Mod nbShafts;
  //MissDentLg := 1;
  For i := 0 To 14999 Do
    NewDentsVides[i] := 0;
  TbWarpPattern[0] := 0;
  TbWeftPattern[0] := 0;

  DM.CalculDensiteWarpFini;
  DM.CalculDensiteWeftFini;
  FirstSimul := false;
End;

{*------------------------------------------------------------------------------
  melange yarn
  @param Sender   ParameterDescription
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure TFrmMain.ActChineExecute(Sender: TObject);
Begin
  If TAction(Sender).Tag In [5..8] Then
    Begin
      ActMouline.Checked := false;
      Actchine.Checked := False;
      ActMultico.Checked := False;
      ActImprime.Checked := False;
      ActJaspe.Checked := False;
      ActTexture.Checked := False;
      ActFancy.Checked := False;
      ActSpaceDye.Checked := False;
      ActSlub.Checked := False;
      UpdateTypeDeFil;
      Showmessage('Impossible de modifier directement cette propriété, vous devez passez par les assistants pour le faire !');
      exit;
    End;
  TAction(Sender).Checked := True;
  TypeDeFil := TAction(Sender).Tag;
  PreviousTypeDeFil := TypeDeFil;
  UpdateTypeDeFil;
End;

Procedure TFrmMain.N301Click(Sender: TObject);
Begin
  TheZoom := TMenuitem(Sender).Tag;
  TMenuitem(Sender).Checked := true;
End;

Procedure TFrmMain.ActOpenExecute(Sender: TObject);
Begin
  odfil.InitialDir := ExpandFileName(DataFolder + '\' + YarnFolder);
  If odfil.Execute Then
    OpenFil(odfil.filename);
End;

{*------------------------------------------------------------------------------
  open yarns
  @param fname   ParameterDescription
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure TfrmMain.OpenFil(fname: String);
Var
  i, lg, ht, a: integer;
  b, version: byte;
  fs: tfilestream;
  j: single;
  c: currency;
  s: String;
  new, bo, YarnNotCompressed: boolean;
  Couleur: Tcolor;
  bmp: Tbitmap;
  filHeader: TfilHeader;
  fibre: boolean;
  LL, aa, BB, R1, G1, B1: byte;
  CoulLab: EncodedLabValue;
  Col, F: Array[0..9] Of TColor;

  Procedure ReadText(Var texte: String);
  Begin
    fs.ReadBuffer(i, sizeof(i));
    s := '';
    setlength(s, i);
    fs.ReadBuffer(s[1], i);
    texte := s;
  End;

  Procedure TexteAVide(Var texte: String);
  Begin
    texte := '';
  End;
Begin
  CompressReliefCB.Checked := false;
  CompressRondeurCB.Checked := false;
  Screen.Cursor := CrHourGlass;
  PercentContrast := ContrastEdit.value;
  chargement := true;
  nomfil := fname;
  sdfil.filename := nomfil;
  SB.panels[2].Text := extractfilename(fname);
  //Caption := siLangmain.GetTextOrDefaultW('IDS_62' (* 'Création de fils - ' *)) + nomfil;
  //TBMRUList1.Add(nomfil);
  Try
    fs := TFileStream.Create(nomfil, fmOpenRead Or fmShareDenyWrite);

    //Magic Number
    fs.Read(filHeader, sizeof(TfilHeader));
    If filHeader.Magic <> filMagic Then
      Begin
        FrmMain.siLangmain.ShowMessage('This is not an fifo-Yarns File');
        fs.Free;
        exit;
      End;

    fs.ReadBuffer(version, 1);          //n° de version
    If version > vers Then
      Begin
        FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_64' (* 'Ce fil a été créé avec une version plus récente de fifo-Yarns' *)));
        FreeandNil(FS);
        exit;
      End;

    fs.ReadBuffer(new, sizeof(new));    // tag nouveau fil ou modifié
    If version > 2 Then
      Begin
        fs.ReadBuffer(Couleur, sizeof(couleur));
        //Depuis la version 11
        If version > 10 Then
          fs.ReadBuffer(CoulLab, sizeof(CoulLab));
        Try
          bmp := TBitmap.Create;
          bmp.LoadFromStream(fs);
          bmp.LoadFromStream(fs);
        Finally
          bmp.Free;
        End;
      End;

    If version > 18 Then
      fs.ReadBuffer(YarnNotCompressed, 1)
    Else
      YarnNotCompressed := false;

    If version > 7 Then
      Begin
        //Code Matiere
        fs.ReadBuffer(i, sizeof(i));
        s := '';
        setlength(s, i);
        fs.ReadBuffer(s[1], i);
        CodeMatiere := s;
        //Code Couleur
        fs.ReadBuffer(i, sizeof(i));
        s := '';
        setlength(s, i);
        fs.ReadBuffer(s[1], i);
        CodeCouleur := s;
        //Description
        fs.ReadBuffer(i, sizeof(i));
        s := '';
        setlength(s, i);
        fs.ReadBuffer(s[1], i);
        Description := s;
      End;

    fs.ReadBuffer(b, 1);
    TypeDeFil := b;
    PreviousTypeDeFil := TypeDeFil;
    UpdateTypeDeFil;
    Case TypeDeFil Of                   // type de fil
      0: ActMouline.Checked := true;
      1: Actchine.Checked := true;
      2: ActMultico.Checked := true;
      3: ActImprime.Checked := true;
      4: ActJaspe.Checked := true;
      5: ActTexture.Checked := true;
      6: ActFancy.Checked := true;
      7: ActSpaceDye.Checked := true;
      8: ActSlub.Checked := true;
    End;
    fs.ReadBuffer(b, 1);
    RG_sens_fil.ItemIndex := b;         // sens du fil
    fs.ReadBuffer(j, sizeof(j));
    fil_titre.value := j;               // titre du fil
    fs.ReadBuffer(i, sizeof(i));
    TitreCb.ItemIndex := i;             // unite du fil
    //Depuis la version 20
    If version > 19 Then
      Begin
        //Depuis la version 20   Nombre de bruns
        fs.ReadBuffer(i, sizeof(i));
        RetorsEdit.value := i;
        //
      End
    Else
      RetorsEdit.value := 1;
    //
    fs.ReadBuffer(j, sizeof(j));
    fil_densite.value := j;             // coef matiere
    fs.ReadBuffer(i, sizeof(integer));
    fil_torsion.Value := i;             // torsion du fil
    fs.ReadBuffer(i, sizeof(integer));
    fil_irregularite_Longueur.Value := i; // irregularite du fil

    fs.ReadBuffer(i, sizeof(i));
    F_Irregularite.value := i;          // facteur d'irregularite du fil

    If version > 3 Then
      Begin
        fs.ReadBuffer(i, sizeof(i));
        DiametreEdit.value := i;        // irregularite du fil en largueur

        fs.ReadBuffer(i, sizeof(i));
        //ondulation.value := i;          // ondulation du fil
      End;

    fs.ReadBuffer(i, sizeof(i));
    ContrastEdit.value := i;            // relief du fil
    fs.ReadBuffer(i, sizeof(i));
    RangeEdit.value := i;               // nombre de rangee de la compo du fil
    GrilleMatiere.RowCount := i + 1;

    For a := 0 To RangeEdit.value - 1 Do
      Begin
        fs.ReadBuffer(i, sizeof(i));
        GrilleMatiere.cells[0, a + 1] := inttostr(i); // % de la matiere
        fs.ReadBuffer(i, sizeof(i));
        s := '';
        setlength(s, i);
        fs.ReadBuffer(s[1], i);
        GrilleMatiere.cells[1, a + 1] := s; //matiere
      End;

    fs.ReadBuffer(c, sizeof(c));
    fil_prix.Value := c;                // prix du fil

    If version > 6 Then
      Begin
        fs.ReadBuffer(i, sizeof(i));    // rondeur du fil
        RondeurEdit.value := i;

        fs.ReadBuffer(i, sizeof(i));    // variation de diametre du fil
        DiametreEdit.value := i;

        fs.ReadBuffer(i, sizeof(i));    // longueur des poils du fil
        PilositeEdit.value := i;

        fs.ReadBuffer(i, sizeof(i));    // densite des poils du fil
        DensPoilEdit.value := i;
      End
    Else
      Begin
        RondeurEdit.value := 0;
        DiametreEdit.value := 0;
        PilositeEdit.value := 0;
        DensPoilEdit.value := 0;
      End;

    //depuis la version 10 on sauve la position du milieu
    If version > 9 Then
      Begin
        fs.ReadBuffer(CentreTop, sizeof(CentreTop));
        fs.ReadBuffer(CentreBottom, sizeof(CentreBottom));
      End
    Else
      Begin
        CentreTop := 0;
        CentreBottom := 19;
      End;

    YBottom := CentreBottom;
    YTop := CentreTop;

    fs.ReadBuffer(b, 1);
    space_rapport.Value := b;           // rapport du space
    fs.ReadBuffer(b, 1);
    multico_rapport.Value := b;         // temps de repetition du multico
    fs.ReadBuffer(b, 1);
    multico_random_max.Value := b;      // longueur aleatoire maximum du multico

    If version > 8 Then
      Begin
        // fil delavable
        fs.ReadBuffer(bo, sizeof(bo));
        DelavableCB.Checked := bo;
      End
    Else
      DelavableCB.Checked := false;

    fs.ReadBuffer(b, 1);                // nombre de couleurs
    fil_nbcolor.Value := b;
    ColorGrid.RowCount := b + 1;

    For i := 0 To Trunc(fil_nbcolor.Value) - 1 Do
      Begin                             // pour chaque couleur
        fs.ReadBuffer(b, 1);
        tabcolor[i].moyen.rgbRed := b;  // composante a
        fs.ReadBuffer(b, 1);
        tabcolor[i].moyen.rgbGreen := b; // composante a
        fs.ReadBuffer(b, 1);
        tabcolor[i].moyen.rgbBlue := b; // composante L
        tabcolor[i].moyen.rgbReserved := 255;
        LABValue.b := tabcolor[i].moyen.rgbRed;
        LABValue.a := tabcolor[i].moyen.rgbGreen;
        LABValue.L := tabcolor[i].moyen.rgbBlue;
        //cmsDoTransform(EcranTransform, @LABValue, @EncodedRGB, 1);
        LL := LABValue.L;
        AA := LABValue.a;
        BB := LABValue.b;

        hDisplayProfile.GetRGBFromLab(LL, AA, BB, r1, g1, b1);
        Col[i] := RGB(R1, G1, B1);
        F[i] := clwhite Xor RGB(R1, G1, B1);
        fs.ReadBuffer(b, 1);
        tabpred[i] := b;                // predominance
        tabcolor[i].clair := Eclaircir(tabcolor[i].moyen, PercentContrast); // on calcule les couleurs + clair
        tabcolor[i].fonce := Assombrir(tabcolor[i].moyen, PercentContrast); // et + fonce

        //Depuis la version 18
        If version > 17 Then
          ReadText(s)
        Else
          TexteAVide(s);
        NomCouleur[i] := s;
      End;

    For i := 0 To ColorGrid.RowCount - 2 Do
      Begin
        ColorGrid.Cells[2, i + 1] := inttostr(tabpred[i]);
        ColorGrid.Colors[1, i + 1] := Col[i];
        ColorGrid.FontColors[1, i + 1] := F[i];
        ColorGrid.Cells[1, i + 1] := NomCouleur[i];
      End;

    //if RG_type_fil.ItemIndex = 1 then begin
    fs.ReadBuffer(i, sizeof(i));        // si melange, le premier champs libre est recupere
    Edit_TailleCarreau.Value := i;      // (c'est la taille des carreaux du melange)
    If version > 1 Then
      fs.ReadBuffer(i, sizeof(i))
    Else
      i := 1;                           // si melange, on recupere l'espacement
    EspacementEdit.Value := i;
    //end;

    If version > 5 Then
      Begin
        fs.ReadBuffer(fibre, sizeof(fibre));
        fibrecb.Checked := fibre;       //si les fibres du melangé sont orientées
      End;

    fs.ReadBuffer(i, sizeof(i));        // 4 champs libre de la taille d'un entier
    fs.ReadBuffer(i, sizeof(i));
    fs.ReadBuffer(i, sizeof(i));
    fs.ReadBuffer(i, sizeof(i));

    //Donnée technique depuis la version 15
    If version > 14 Then
      Begin
        ReadText(CodeBarre);
        ReadText(DesignationCouleur);
        ReadText(CodeRecetteTeinture);
        ReadText(DesignationRecetteTeinture);
        ReadText(TitrageCommercial);
        ReadText(CodeComposition);
        ReadText(DesignationComposition);
        ReadText(CodeFammile);
        ReadText(DesignationFamille);
        ReadText(Observations);
        ReadText(CodeUniteStockage);
        ReadText(DesignationUniteStockage);
        ReadText(CodeStructure);
        ReadText(DesignationStructure);
        ReadText(ObservationTechnique);
        ReadText(TypeArticle);
        fs.ReadBuffer(FilTeint, sizeof(FilTeint));
      End
    Else
      Begin
        TexteAVide(CodeBarre);
        TexteAVide(DesignationCouleur);
        TexteAVide(CodeRecetteTeinture);
        TexteAVide(DesignationRecetteTeinture);
        TexteAVide(TitrageCommercial);
        TexteAVide(CodeComposition);
        TexteAVide(DesignationComposition);
        TexteAVide(CodeFammile);
        TexteAVide(DesignationFamille);
        TexteAVide(Observations);
        TexteAVide(CodeUniteStockage);
        TexteAVide(DesignationUniteStockage);
        TexteAVide(CodeStructure);
        TexteAVide(DesignationStructure);
        TexteAVide(ObservationTechnique);
        TexteAVide(TypeArticle);
        FilTeint := false;
      End;
    ////////////////////

      //si c'est un mouliné composé de chiné alors on sauve les chines
    If TypeDeFil = 4 Then
      Begin

        If imgfilexport <> Nil Then
          FreeAndNil(imgfilexport);
        imgfilexport := TDIBimage.Create;
        If version > 4 Then
          Begin
            If version < 14 Then
              LoadDibFromPNGStream(imgfilexport, Fs)
            Else
              LoadBitmapFromCompressedStream(imgfilexport, TStream(Fs), YarnNotCompressed);
          End
        Else
          imgfilexport.LoadFromStream(fs, 0); //L'image du fil

        For i := 0 To trunc(fil_nbcolor.Value) - 1 Do
          Begin
            If ImageFil[i] <> Nil Then
              FreeAndNil(ImageFil[i]);
            ImageFil[i] := TDibImage.Create;
            If fs.position <> (fs.Size - 1) Then
              If version > 4 Then
                Begin
                  If version < 14 Then
                    LoadDibFromPNGStream(ImageFil[i], Fs)
                  Else
                    LoadBitmapFromCompressedStream(ImageFil[i], TStream(Fs), YarnNotCompressed);
                End
              Else
                ImageFil[i].LoadFromStream(fs, 0);
          End;
      End;

    //si le fil est un fil scanné on lit aussi son image origianle rgb, pour d'eventuelle modification ulterieure
    If TypeDeFil > 4 Then
      Begin

        If imgfilfantaisieimport <> Nil Then
          FreeAndNil(imgfilfantaisieimport);

        If TypeDeFil < 7 Then
          Begin
            imgfilfantaisieimport := TDIBimage.Create;
            If version > 4 Then
              Begin
                If version < 14 Then
                  LoadDibFromPNGStream(imgfilfantaisieimport, Fs)
                Else
                  LoadBitmapFromCompressedStream(imgfilfantaisieimport, TStream(Fs), YarnNotCompressed);
              End
            Else
              imgfilfantaisieimport.LoadFromStream(fs, 0); //L'image du fil

            If SourceDib <> Nil Then
              FreeAndNil(SourceDib);
            SourceDib := TDibImage.Create;
            If version > 4 Then
              Begin
                If version < 14 Then
                  LoadDibFromPNGStream(SourceDib, Fs)
                Else
                  LoadBitmapFromCompressedStream(SourceDib, TStream(Fs), YarnNotCompressed);
              End
            Else
              SourceDib.LoadFromStream(fs, 0); // l'image de la texture
            // on passe le fil en Lab
            If version < 12 Then
              ConvertRGBToDib(SourceDib);
          End
        Else                            //Space dye
          Begin
            If imgfilexport = Nil Then
              imgfilexport := TDibImage.Create;
            LoadBitmapFromCompressedStream(imgfilexport, TStream(Fs), YarnNotCompressed);
          End;
      End;

    //depuis la version 14
    If TypeDeFil = 7 Then
      Begin
        fs.readbuffer(NBSection, sizeof(NBSection));
        For i := 0 To NBSection - 1 Do
          Begin
            fs.readbuffer(SectionColor[i], sizeof(SectionColor[i]));
            fs.readbuffer(SectionSize[i], sizeof(SectionSize[i]));
            //Depuis la version 18
            If Version > 17 Then
              ReadText(SectionNomCouleur[i])
            Else
              TexteAVide(SectionNomCouleur[i]);
          End;
      End;

    //depuis la version 16
    If TypeDeFil = 8 Then
      Begin
        fs.readbuffer(SlubTitle, sizeof(SlubTitle));
        fs.readbuffer(NBSectionSlub, sizeof(NBSectionSlub));
        fs.readbuffer(StoneContrast, Sizeof(StoneContrast));
        For i := 0 To NBSectionSlub - 1 Do
          Begin
            fs.readbuffer(SectionSlubSize[i], sizeof(SectionSlubSize[i]));
            fs.readbuffer(sectionSlubPentG[i], sizeof(sectionSlubPentG[i]));
            fs.readbuffer(sectionSlubPentD[i], sizeof(sectionSlubPentD[i]));
          End;
        If version > 20 Then
          Begin
            For i := 0 To NBSectionSlub - 1 Do
              Begin
                fs.readbuffer(SlubSectionColor[i], sizeof(SlubSectionColor[i]));
                ReadText(SlubSectionNomCouleur[i]);
              End;
          End
        Else
          Begin
            For i := 0 To NBSectionSlub - 1 Do
              Begin
                SlubSectionColor[i].rgbRed := tabcolor[0].moyen.rgbRed;
                SlubSectionColor[i].rgbGreen := tabcolor[0].moyen.rgbGreen;
                SlubSectionColor[i].rgbBlue := tabcolor[0].moyen.rgbBlue;
                SlubSectionNomCouleur[i] := '';
              End;
          End;
      End;

    //depuis la version 17
    If Version > 16 Then
      fs.readbuffer(Prelavage, sizeof(Prelavage))
    Else
      prelavage := 0;
    PrelavageEdit.value := prelavage;

  Finally
    If fs <> Nil Then
      FreeAndNil(fs);
  End;
  UpdateTypeDeFil;
  SB.panels[2].Text := extractfilename(fname);
  RecentMenu1.AddFile(fname);
  Screen.Cursor := CrDefault;
  ActSimul.Enabled := true;
  ActSaveAs.Enabled := true;
  ActSave.Enabled := true;
  ActPrint.Enabled := true;
  CreerFil;

End;

Procedure TFrmMain.ActSaveExecute(Sender: TObject);
Begin
  If sdfil.FileName <> '' Then
    SauverFil(sdfil.FileName)
  Else
    ActSaveAsExecute(Sender);
End;

Procedure TFrmMain.ActSaveASExecute(Sender: TObject);
Begin
  sdfil.InitialDir := ExpandFileName(DataFolder + '\' + YarnFolder);
  If sdfil.Execute Then
    SauverFil(sdfil.filename);
End;

Procedure TFrmMain.ActPrinterSetupExecute(Sender: TObject);
Begin
  If psd.Execute Then
    DM.DoAutoProfile;
End;

Procedure TFrmMain.ActPrintExecute(Sender: TObject);
Begin
  If Not pd.Execute Then
    exit;
  If FirstSimul Then
    ActDobbyWizardExecute(Sender);
  If FirstSimul Then
    exit;

  SelImp := TSelImp.Create(self);
  If SelImp.ShowModal = MrOk Then
    Begin
      SimulAborted := false;
      TypeDeSimulation := 13;
      CreerFil;
      SimulDPI := PrinterDPI;
      DM.DoSimulation(YarnSize, EndImpX - StartImpX, EndImpY - StartImpY, true, 0, 0, false);
    End;
  varcourante := VC;
End;

Procedure TFrmMain.ActSimulSmallUpdate(Sender: TObject);
Begin
  ActSimulSmall.Enabled := ActSimul.Enabled;
End;

Procedure TFrmMain.ActPrefExecute(Sender: TObject);
Begin
  PrefFrom := TPrefFrom.Create(self);
  PrefFrom.ShowModal;
End;

Procedure TFrmMain.ActBrowseExecute(Sender: TObject);
Begin
  If ThumbForm = Nil Then
    ThumbForm := TThumbForm.Create(self);
  ThumbForm.BringToFront;
End;

{*------------------------------------------------------------------------------
  Adjust colors
  @param Sender   ParameterDescription
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure TFrmMain.ActAdjustColorExecute(Sender: TObject);
Var
  Bits: pointer;
  ADib: TDibImage;
  i, count, PercentContrast: integer;
  LL, aa, BB, R1, G1, B1: byte;
Begin
  ADjustForm := TADjustForm.Create(self);
  If ADjustForm.ShowModal = MrOk Then
    Begin
      Case TypeDeFil Of
        0..3:
          Begin
            PercentContrast := ContrastEdit.value;
            Count := Trunc(fil_nbcolor.value);
            For i := 0 To Count - 1 Do
              Begin
                tabcolor[i].moyen := TCMProfile.CompenseDeriv(tabcolor[i].moyen, AdjustLab);
                tabcolor[i].moyen.rgbReserved := 255;
                LL := tabcolor[i].moyen.rgbBlue;
                AA := tabcolor[i].moyen.rgbGreen;
                BB := tabcolor[i].moyen.rgbRed;
                hDisplayProfile.GetRGBFromLab(LL, AA, BB, r1, g1, b1);
                ColorGrid.Colors[1, i + 1] := RGB(R1, G1, B1);
                ColorGrid.FontColors[1, i + 1] := clwhite Xor RGB(R1, G1, B1);
                tabcolor[i].clair := Eclaircir(tabcolor[i].moyen, PercentContrast);
                tabcolor[i].fonce := Assombrir(tabcolor[i].moyen, PercentContrast);
              End;
          End;
        4:
          Begin
            Count := Trunc(fil_nbcolor.value);
            For i := 0 To Count - 1 Do
              Begin
                If ImageFil[i] <> Nil Then
                  Begin
                    ADib := ImageFil[i];
                    Bits := ADib.Bits;
                    TCMProfile.CorrectionDeriv(Bits, ADib.Width * ADib.Height, 4, AdjustLab);
                  End;
              End;
          End;
        5, 6, 7:
          Begin
            If imgfilfantaisieimport <> Nil Then
              Begin
                ADib := imgfilfantaisieimport;
                Bits := ADib.Bits;
                TCMProfile.CorrectionDeriv(Bits, ADib.Width * ADib.Height, 4, AdjustLab);
              End;
            If imgfilexport <> Nil Then
              Begin
                ADib := imgfilexport;
                Bits := ADib.Bits;
                TCMProfile.CorrectionDeriv(Bits, ADib.Width * ADib.Height, 4, AdjustLab);
              End;
          End;
      End;
      ActSimulExecute(Sender);
    End;
End;

{*------------------------------------------------------------------------------
  Create yarn color library
  @param Sender   ParameterDescription
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure TFrmMain.ActCreateGammeExecute(Sender: TObject);
Var
  r, g, b, ri, gi, bi, i, j, count, t: integer;
  changed: boolean;
  LL, aa, BB, R1, G1, B1: byte;
Begin
  LibYarnName := 'Test ';
  LibraryForm := TLibraryForm.Create(self);
  If LibraryForm.ShowModal <> MrOk Then
    exit;
  If LibYarnName = '' Then
    exit;

  application.ProcessMessages;
  PercentContrast := ContrastEdit.value;
  i := 0;                               //RG_choix_color.ItemIndex;
  count := ListeCouleur.Count;
  stopgamme := false;
  spb.MaxProgress := count - 1;
  spb.Execute;
  t := 0;
  For j := 0 To count - 1 Do
    If CheckLibYarn[j] Then
      Begin
        inc(t);
        EncodedLab.L := L2Fix(TColorLib(ListeCouleur.Items[j]).L);
        EncodedLab.a := ab2Fix(TColorLib(ListeCouleur.Items[j]).a);
        EncodedLab.b := ab2Fix(TColorLib(ListeCouleur.Items[j]).b);

        tabcolor[i].moyen.rgbRed := EncodedLab.b;
        tabcolor[i].moyen.rgbGreen := EncodedLab.a;
        tabcolor[i].moyen.rgbBlue := EncodedLab.L;
        tabcolor[i].moyen.rgbReserved := 255;

        LL := EncodedLab.L;
        AA := EncodedLab.a;
        BB := EncodedLab.b;

        hDisplayProfile.GetRGBFromLab(LL, AA, BB, r1, g1, b1);

        //cmsDoTransform(EcranTransform, @LABValue, @EncodedRGB, 1);
        ColorGrid.Colors[1, i + 1] := RGB(R1, G1, B1);
        ColorGrid.FontColors[1, i + 1] := clwhite Xor RGB(R1, G1, B1);
        tabcolor[i].clair := Eclaircir(tabcolor[i].moyen, PercentContrast);
        tabcolor[i].fonce := Assombrir(tabcolor[i].moyen, PercentContrast);
        ActSimul.Enabled := true;
        //sb.Panels[1].text := TColorLib(ListeCouleur.Items[j]).Name;
        sb.Panels[1].text := inttostr(t + 1) + ' of ' + inttostr(CheckLibCount);

        ActSimulExecute(Sender);
        application.ProcessMessages;
        If Prefix = 0 Then
          sdfil.FileName := TrimRight(ExpandFileName(DataFolder + '\' + YarnFolder) + '\' + LibYarnName + ' ' + TColorLib(ListeCouleur.Items[j]).Name) + '.fil'
        Else
          sdfil.FileName := TrimRight(ExpandFileName(DataFolder + '\' + YarnFolder) + '\' + TColorLib(ListeCouleur.Items[j]).Name + ' ' + LibYarnName) + '.fil';
        Try
          ActSaveExecute(Sender);
        Except
          Showmessage('Cannot save ' + sdfil.FileName);
        End;
        spb.UpdateProgress(j, 'Creating ' + inttostr(j + 1) + ' of ' + inttostr(Count + 1), 'Yarn ' + extractfilename(sdfil.FileName));
        If spb.Canceled Then
          Begin
            exit;
          End;
      End;
  SetLength(CheckLibYarn, 0);
  spb.UpdateProgress(Count - 1, 'Conversion done !', inttostr(t) + ' yarns converted !');
End;

Procedure TFrmMain.ActCheckUpdateExecute(Sender: TObject);
Begin
  If goChecker1.IsOnline Then
    goChecker1.CheckUpdate;
End;

Procedure TFrmMain.ActExportDatasExecute(Sender: TObject);
Begin
  SMEWizardDlg1.Execute;
End;

Procedure TFrmMain.ActDatasExecute(Sender: TObject);
Begin
  DataForm := TDataForm.Create(Self);
  DataForm.ShowModal;
End;

Procedure TFrmMain.goChecker1AbortDownload(Sender: TObject);
Begin
  FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_214' (* 'Download is aborted!' *)));
End;

Procedure TFrmMain.goChecker1Available(VersionInfo, DateBuildInfo: String);
Var
  s: String;
Begin
  If FrmMain.siLangmain.MessageDlg(#13 + siLangmain.GetTextOrDefault('IDS_215' (* 'New version is available ' *)) + VersionInfo + siLangmain.GetTextOrDefault('IDS_216' (* ' date build : ' *)) + DateBuildInfo +
    #13 + siLangmain.GetTextOrDefault('IDS_217' (* 'Would you like to download now ?' *)), mtConfirmation, [MbYes, MbNo], 0) = mrYes Then
    Begin
      // Download the update installer, select the folder where you want the update will be stored.
      s := ExtractFilePath(ParamStr(0)) + '\Updates';
      If Not DirectoryExists(s) Then
        CreateDir(s);

      goChecker1.SetDestFolder(ExtractFilePath(ParamStr(0)));

      goChecker1.DownloadUpdate(S);
    End;
End;

Procedure TFrmMain.goChecker1BeforeUpdate(Var RunUpdate: Boolean);
Begin
  // This event is called before the update installer run, you can choose if you want to run update
  // after download using RunUpdate var.

  If FrmMain.siLangmain.MessageDlg(#13 + siLangmain.GetTextOrDefault('IDS_219' (* 'Would you like to run the update now ?' *)), MtConfirmation, [MbYes, MbNo], 0) = mrYes Then
    Begin
      RunUpdate := True;
      DoNotConfirmClose := true;
      Close;
    End
  Else
    RunUpdate := False;
End;

Procedure TFrmMain.goChecker1DownloadError(ErrorMsg: String;
  StatusCode: Integer);
Begin
  FrmMain.siLangmain.ShowMessage(ErrorMsg);
End;

Procedure TFrmMain.goChecker1NotOnline(Sender: TObject);
Begin
  FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_220' (* 'No Internet Connection found, impossible to download INF File from the web server' *)));

End;

Procedure TFrmMain.goChecker1NoUpdate(Sender: TObject);
Begin
  FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_221' (* 'No Update is available!' *)));

End;

Procedure TFrmMain.goChecker1UpdateError(ErrorMsg: String);
Begin
  FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_222' (* 'An error occured while trying to update ' *)) + ErrorMsg);

End;

Procedure TFrmMain.SMEVirtualDataEngine1Count(Sender: TObject;
  Var Count: Integer);
Begin
  {we must say how many rows we want to export}
  Count := 1;
End;

{*------------------------------------------------------------------------------
  Yqrn export
  @param Sender   ParameterDescription
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure TFrmMain.SMEVirtualDataEngine1FillColumns(Sender: TObject);
Begin
  {we must define columns which will be exported.
  As alternative you can define a same Columns directly in TSMExportToXLS.Columns

  IMPORTANT:
  Must be defined at least one column}

  SMEWizardDlg1.Columns.Clear;

  {add first virtual column}
  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Name';
      Title.Caption := 'Name';
      DataType := ctString;
      Width := 100
    End;

  {add second virtual column}
  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Type';
      Title.Caption := 'Type';
      DataType := ctString;
      Width := 40
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Twist Direction';
      Title.Caption := 'Twist Direction';
      DataType := ctString;
      Width := 20
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Count';
      Title.Caption := 'Count';
      DataType := ctDouble;
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Count Unit';
      Title.Caption := 'Count Unit';
      DataType := ctString;
      Width := 10;
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Price';
      Title.Caption := 'Price';
      DataType := ctCurrency;
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'NB Colors';
      Title.Caption := 'NB Colors';
      DataType := ctInteger;
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Material Code';
      Title.Caption := 'Material Code';
      DataType := ctString;
      Width := 40
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Color Code';
      Title.Caption := 'Color Code';
      DataType := ctString;
      Width := 40
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Description';
      Title.Caption := 'Description';
      DataType := ctString;
      Width := 100
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Twist';
      Title.Caption := 'Twist';
      DataType := ctInteger;
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Composition';
      Title.Caption := 'Composition';
      DataType := ctString;
      Width := 100
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Lab color';
      Title.Caption := 'Lab color';
      DataType := ctInteger;
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'BarCode';
      Title.Caption := 'BarCode';
      DataType := ctString;
      Width := 100
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Color Description';
      Title.Caption := 'Color Description';
      DataType := ctString;
      Width := 100
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Colour Recipe Code';
      Title.Caption := 'Colour Recipe Code';
      DataType := ctString;
      Width := 100
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Colour Recipe Designation';
      Title.Caption := 'Colour Recipe Designation';
      DataType := ctString;
      Width := 100
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Commercial Yarn Count';
      Title.Caption := 'Commercial Yarn Count';
      DataType := ctString;
      Width := 100
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Composition Code';
      Title.Caption := 'Composition Code';
      DataType := ctString;
      Width := 100
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Composition Designation';
      Title.Caption := 'Composition Designation';
      DataType := ctString;
      Width := 100
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Family Code';
      Title.Caption := 'Family Code';
      DataType := ctString;
      Width := 100
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Family Designation';
      Title.Caption := 'Family Designation';
      DataType := ctString;
      Width := 100
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Yarn Dyed';
      Title.Caption := 'Yarn Dyed';
      DataType := ctBoolean;
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Observations';
      Title.Caption := 'Observations';
      DataType := ctString;
      Width := 100
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Storage Unit Code';
      Title.Caption := 'Storage Unit Code';
      DataType := ctString;
      Width := 100
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Storage Unit Designation';
      Title.Caption := 'Storage Unit Designation';
      DataType := ctString;
      Width := 100
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Structure Designation';
      Title.Caption := 'Structure Designation';
      DataType := ctString;
      Width := 100
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Technical Observation';
      Title.Caption := 'Technical Observation';
      DataType := ctString;
      Width := 100
    End;

  With SMEWizardDlg1.Columns.Add Do
    Begin
      FieldName := 'Yarn Type';
      Title.Caption := 'Yarn Type';
      DataType := ctString;
      Width := 100
    End;
End;

{*------------------------------------------------------------------------------
  Get Yarn compo
  @return String   ResultDescription  
------------------------------------------------------------------------------*}
Function GetCompoFil: String;
Var
  i: integer;
  s: String;
Begin
  s := '';
  With frmMain Do
    Begin
      For i := 1 To RangeEdit.Value Do
        Begin
          If i < RangeEdit.Value Then
            s := s + GrilleMatiere.Cells[0, i] + ' ' + GrilleMatiere.Cells[1, i] + '|'
          Else
            s := s + GrilleMatiere.Cells[0, i] + ' ' + GrilleMatiere.Cells[1, i];
        End;
    End;
  Result := s;
End;


{*------------------------------------------------------------------------------
  Set fiels values for export
  @param Sender   ParameterDescription
  @param Column   ParameterDescription
  @param Value   ParameterDescription
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure TFrmMain.SMEVirtualDataEngine1GetValue(Sender: TObject;
  Column: TSMEColumn; Var Value: Variant);
Var
  s: String;
Begin
  s := ExtractFileName(nomfil);
  {here we must return a value for current row for Column}
  If Assigned(Column) Then
    Begin
      If (Column.Index = 0) Then
        Value := copy(S, 1, length(S) - 4)
      Else
        If (Column.Index = 1) Then
          Value := GetTypeDeFilName(TypeDeFil)
        Else
          If (Column.Index = 2) Then
            Value := RG_sens_fil.Items[RG_sens_fil.ItemIndex]
          Else
            If (Column.Index = 3) Then
              Value := fil_titre.Value
            Else
              If (Column.Index = 4) Then
                Value := TitreCb.Items[TitreCb.ItemIndex]
              Else
                If (Column.Index = 5) Then
                  Value := fil_prix.Value
                Else
                  If (Column.Index = 6) Then
                    Value := fil_nbcolor.Value
                  Else
                    If (Column.Index = 7) Then
                      Value := CodeMatiere
                    Else
                      If (Column.Index = 8) Then
                        Value := CodeCouleur
                      Else
                        If (Column.Index = 9) Then
                          Value := description
                        Else
                          If (Column.Index = 10) Then
                            Value := fil_torsion.Value
                          Else
                            If (Column.Index = 11) Then
                              Value := GetCompoFil
                            Else
                              If (Column.Index = 12) Then
                                Begin
                                  Value := ClWhite;
                                  If (imgfilExport <> Nil) And (imgfilExport.IsImageAssigned) Then
                                    Value := GetMediumColor(imgfilExport, Trunc(fil_nbcolor.Value), TypeDeFil);
                                End
                              Else
                                If (Column.Index = 13) Then
                                  Value := CodeBarre
                                Else
                                  If (Column.Index = 14) Then
                                    Value := DesignationCouleur
                                  Else
                                    If (Column.Index = 15) Then
                                      Value := CodeRecetteTeinture
                                    Else
                                      If (Column.Index = 16) Then
                                        Value := DesignationRecetteTeinture
                                      Else
                                        If (Column.Index = 17) Then
                                          Value := TitrageCommercial
                                        Else
                                          If (Column.Index = 18) Then
                                            Value := CodeComposition
                                          Else
                                            If (Column.Index = 19) Then
                                              Value := DesignationComposition
                                            Else
                                              If (Column.Index = 20) Then
                                                Value := CodeFammile
                                              Else
                                                If (Column.Index = 21) Then
                                                  Value := DesignationFamille
                                                Else
                                                  If (Column.Index = 22) Then
                                                    Value := FilTeint
                                                  Else
                                                    If (Column.Index = 23) Then
                                                      Value := Observations
                                                    Else
                                                      If (Column.Index = 24) Then
                                                        Value := CodeUniteStockage
                                                      Else
                                                        If (Column.Index = 25) Then
                                                          Value := DesignationUniteStockage
                                                        Else
                                                          If (Column.Index = 26) Then
                                                            Value := CodeStructure
                                                          Else
                                                            If (Column.Index = 27) Then
                                                              Value := DesignationStructure
                                                            Else
                                                              If (Column.Index = 28) Then
                                                                Value := ObservationTechnique
                                                              Else
                                                                If (Column.Index = 29) Then
                                                                  Value := TypeArticle
                                                                Else
                                                                  Value := '3'
    End;
End;

{*------------------------------------------------------------------------------
  get yarn type
  @param index   ParameterDescription
  @return String   ResultDescription  
------------------------------------------------------------------------------*}
Function TFrmMain.GetTypeDeFilName(index: integer): String;
Begin
  Case index Of
    0: Result := 'Mouliné';
    1: Result := 'Chiné';
    2: Result := 'Multicouleurs';
    3: Result := 'Imprimé';
    4: Result := 'Jaspé';
    5: Result := 'Scannés';
    6: Result := 'Fantaisie';
    7: Result := 'Space dye';
    8: Result := 'Flammés';
  End;
End;

{*------------------------------------------------------------------------------
  Textured yarn wizard
  @param Sender   ParameterDescription
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure TFrmMain.ActTexturedWizardExecute(Sender: TObject);
Var
  imgpetite, imggrande, imgfiltmp, tmpDib: TDIBimage;
  resol: Single;
  i, j, k, zoom, diametre, diam_calc: integer;
  TitreNm: extended;
  bmp: tbitmap;
Begin
  ScanForm := TScanForm.create(self);
  ScanForm.fil_titre.Value := fil_titre.Value;
  ScanForm.RetorsEdit.Value := RetorsEdit.Value;
  ScanForm.TitreCb.ItemIndex := TitreCb.ItemIndex;
  ScanForm.fil_densite.value := fil_densite.value;
  ScanForm.fil_torsion.Value := fil_torsion.Value;
  ScanForm.ContrastEdit.Value := ContrastEdit.Value;
  ScanForm.fil_irregularite.Value := fil_irregularite_Longueur.Value;
  ScanForm.F_Irregularite.Value := F_Irregularite.Value;
  ScanForm.fil_prix.Value := fil_prix.Value;
  ScanForm.RG_sens_fil.ItemIndex := RG_sens_fil.ItemIndex;
  If ScanForm.ShowModal = mrok Then
    Begin
      fil_titre.Value := ScanForm.fil_titre.Value;
      RetorsEdit.Value := ScanForm.RetorsEdit.Value;
      TitreCb.ItemIndex := ScanForm.TitreCb.ItemIndex;
      fil_densite.value := ScanForm.fil_densite.value;
      fil_torsion.Value := ScanForm.fil_torsion.Value;
      ContrastEdit.Value := ScanForm.ContrastEdit.Value;
      fil_irregularite_Longueur.Value := ScanForm.fil_irregularite.Value;
      F_Irregularite.Value := ScanForm.F_Irregularite.Value;
      fil_prix.Value := ScanForm.fil_prix.Value;
      RG_sens_fil.ItemIndex := ScanForm.RG_sens_fil.ItemIndex;
      TypeDeFil := 5;
      PreviousTypeDeFil := TypeDeFil;
      UpdateTypeDeFil;
      ActSimulExecute(Sender);
    End;
  ScanForm.Free;
End;

{*------------------------------------------------------------------------------
  Fancy yarn wizard
  @param Sender   ParameterDescription
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure TFrmMain.ActFancyWizardExecute(Sender: TObject);
Var
  ms: TMemorystream;
  ADib: TDibImage;
Begin
  If ActShowColorGrid.Checked Then
    Begin
      ActShowColorGrid.Checked := false;
      ColorGridToolbar.Visible := ActShowColorGrid.Checked;
    End;

  If FancyForm = Nil Then
    FancyForm := TFancyForm.Create(self);
  FancyForm.BringToFront;
  If SourceDib <> Nil Then
    Begin
      Try
        ADib := TDibImage.Create;
        ADib.CopyFromDib(SourceDib);
        //on passe le SourceDib en RGB
        ConvertDibToRGB(ADib);
        ms := TMemorystream.Create;
        ms.SetSize(ADib.ImageFileSize);
        ADib.SaveToStream(ms, true);
        ms.Position := 0;

        With FancyForm Do
          Begin
            YTop := CentreTop;
            YBottom := CentreBottom;
            YarnBitmap := TZDataBitmap.Create;
            LoadFromStream(YarnBitmap, ms);

            z.Image.BitMap := YarnBitmap;

            CrayonCheckLayer.Active := true;
            SaveUndo;
            SaveRedo;
          End;

      Finally
        If ms <> Nil Then
          FreeAndNil(ms);
        If ADib <> Nil Then
          FreeAndNil(ADib);
      End;
    End;
End;

{*------------------------------------------------------------------------------
  Slub yarn wizard
  @param Sender   ParameterDescription
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure TFrmMain.ActSlubWizardExecute(Sender: TObject);
Begin
  If ActShowColorGrid.Checked Then
    Begin
      ActShowColorGrid.Checked := false;
      ColorGridToolbar.Visible := ActShowColorGrid.Checked;
    End;
  If SlubForm = Nil Then
    SlubForm := TSlubForm.Create(self);
  SlubForm.BringToFront;
End;

{*------------------------------------------------------------------------------
  Space dye wizard
  @param Sender   ParameterDescription
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure TFrmMain.ActSpaceDyeWizardExecute(Sender: TObject);
Begin
  If SpaceDyeForm = Nil Then
    SpaceDyeForm := TSpaceDyeForm.Create(self);
  SpaceDyeForm.BringToFront;
End;

Procedure TFrmMain.ActShowColorGridExecute(Sender: TObject);
Begin
  ActShowColorGrid.Checked := Not ActShowColorGrid.Checked;
  ColorGridToolbar.Visible := ActShowColorGrid.Checked;
End;

Procedure TFrmMain.ActShowPaletteExecute(Sender: TObject);
Begin
  ActShowPalette.Checked := Not ActShowPalette.Checked;
  ColorsWindow.Visible := ActShowPalette.Checked;
End;

Procedure TFrmMain.ColorListBoxStartDrag(Sender: TObject;
  Var DragObject: TDragObject);
Var
  LL, AA, BB, R1, G1, B1: byte;
Begin
  If ColorListBox.Items.Count < 1 Then
    exit;
  DragColorBox := ColorListBox.ItemIndex;

  If (DragColorBox >= ListeCouleur.Count) Or (DragColorBox < 0) Then
    exit;

  LL := L2Fix(TColorLib(ListeCouleur.Items[DragColorBox]).L);
  AA := ab2Fix(TColorLib(ListeCouleur.Items[DragColorBox]).a);
  BB := ab2Fix(TColorLib(ListeCouleur.Items[DragColorBox]).b);

  DragValue.L := LL;
  DragValue.a := AA;
  DragValue.b := BB;

  hDisplayProfile.GetRGBFromLab(LL, AA, BB, r1, g1, b1);
  DragColor := TColor(rgb(r1, g1, b1));
  DragName := TColorLib(ListeCouleur.Items[DragColorBox]).Name;
  CreateMyCustomDragCursor(DragColor);
  ColorListBox.DragCursor := crMyDragCursor;
End;

Function FindNearestString(s: String; strings: TStrings): integer;
Var
  i: integer;
Begin
  result := -1;
  For i := 0 To Strings.Count - 1 Do
    Begin
      If POS(lowercase(s), lowercase(strings[i])) > 0 Then
        Begin
          Result := i;
          break;
        End;
    End;
End;

{*------------------------------------------------------------------------------
  Search for color
  @param Sender   ParameterDescription
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure TFrmMain.ActfindColorExecute(Sender: TObject);
Var
  NewString: String;
  i: integer;
Begin
  If FrmMain.siLangMain.InputQuery('Recherche de coloris', 'Nom : ', NewString) Then
    Begin
      i := FindNearestString(NewString, ColorListBox.Items);
      If i < 0 Then
        Showmessage('Pas de nom de coloris ressemblant !')
      Else
        Begin
          ColorListBox.TopIndex := i;
          ColorListBox.ItemIndex := i;
        End;
    End;
End;

Procedure TFrmMain.ActfindColorUpdate(Sender: TObject);
Begin
  ActFindColor.Enabled := ColorListBox.Items.Count > 0;
End;

Procedure TFrmMain.IL1ExeNotEncrypted(Sender: TObject);
Begin
  ValidLicence := False;
  FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_212' (* 'Crypting problem' *)));
  application.Terminate;
End;

Procedure TFrmMain.IL1LicenseFull(ExpiresInfo: Integer;
  ExtraLicenseInfo: String);
Begin
  ValidLicence := true;
End;

Procedure TFrmMain.IL1NoneLicense(Sender: TObject);
Begin
  If NBExec > 0 Then
    Begin
      application.terminate;
      exit;
    End;
  If FileExists(GetSysDir + '\fifo-Yarn Licence.ldf') Then
    Begin
      inc(NBExec);
      IL1.LoadLicenseFromFile(GetSysDir + '\fifo-Yarn Licence.ldf');
      IL1.CheckLicense;
      If IL1.LicenseType = ltNone Then
        Begin
          Showmessage('Fichier Licence Invalide !');
          Application.Terminate;
        End;
      If ValidLicence Then
        Begin
          Showmessage('Licence file found, the application will stop, please restart it !');
          Application.Terminate;
        End;
    End
  Else
    Begin
      Showmessage('Fichier Licence non trouvé !');
      application.terminate;
    End;
End;

Procedure TFrmMain.Timer1Timer(Sender: TObject);
Begin
{$IFDEF HASP4}
  If Not CheckHasp(IL1) Then
    Begin
      FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_42' (* 'Please connect the hardware protection key to the computer!!' *)));
      Application.Terminate;
    End;

  If Not Checkfifo(IL1) Then
    Begin
      FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_43' (* 'Contact fifo to use this software' *)));
      Application.Terminate;
    End;

  If Not CheckYarn(IL1) Then
    Begin
      FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_44' (* 'Contact fifo to be able to use YX-WEAVE' *)));
      Application.Terminate;
    End;

{$ELSE}

  If ID_HL <> VerifyID_HL(IL1) Then
    Begin
      FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_209' (* 'You have Changed the key, see you next time  :-)  !!!' *)));
      Application.Terminate;
    End;

  If Not CheckYarn_HL(IL1) Then
    Begin
      FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_44' (* 'Contact fifo to be able to use YX-WEAVE' *)));
      Application.Terminate;
    End;

  If Not IsYarnMaintenanceValid(IL1) Then
    Begin
      FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_48' (* 'Your maintenance contract has expired, you cannot use this release !' *)));
      Application.Terminate;
      exit;
    End;
{$ENDIF}

End;

{*------------------------------------------------------------------------------
  Export yarn image
  @param Sender   ParameterDescription
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure TFrmMain.ActExportImageExecute(Sender: TObject);
Begin
  sdy2t.FileName := '';
  If (FilSansRelief <> Nil) Then
    If sdy2t.execute Then
      Begin
        //if CheckBox1.checked then ConvertDibToRGB(FilSansRelief);
        SaveTwistedFiber(sdy2t.Filename, RG_sens_fil.ItemIndex);
        //FilSansRelief.SaveToFile(SDImage.FileName);
      End;
End;

Procedure TFrmMain.ActExportDatasUpdate(Sender: TObject);
Begin
  ActExportDatas.Enabled := ActSaveAs.Enabled;
End;

Procedure TFrmMain.ActAdjustColorUpdate(Sender: TObject);
Begin
  ActAdjustColor.Enabled := ActSave.Enabled;
End;

Procedure TFrmMain.ActImportDatasUpdate(Sender: TObject);
Begin
  ActImportDatas.Enabled := {HasChamatexPluggin and} ActSimul.Enabled;
End;

Procedure TFrmMain.ActExportImageUpdate(Sender: TObject);
Begin
  ActExportImage.enabled := FilSansRelief <> Nil;
End;

{*------------------------------------------------------------------------------
  open color palette
  @param Sender   ParameterDescription
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure TFrmMain.ActOpenPaletteExecute(Sender: TObject);
Var
  fifoIni: TIniFile;
  i: integer;
Begin
  fifoIni := TIniFile.Create(DossierIni + 'fifo.ini');
  If ODPalette.Execute Then
    Begin
      If ListeCouleur.Count > 0 Then
        For i := ListeCouleur.Count - 1 Downto 0 Do
          Begin
            TColorLib(ListeCouleur.Items[i]).Free;
            ListeCouleur.Delete(i);
          End;
      OpenColorLib(ODPalette.Filename);
      DM.RefreshColorLibrary;
      colorListBox.Clear;
      For i := 0 To ListeCouleur.Count - 1 Do
        colorListBox.Items.Append(TColorLib(ListeCouleur.Items[i]).Name);
      fifoIni.WriteString('Color Lib', 'Name', ODPalette.Filename);
      fifoIni.free;
    End;
End;

Procedure TFrmMain.FrenchExecute(Sender: TObject);
Var
  ini: TInifile;
Begin
  TAction(sender).Checked := true;
  ini := TIniFile.Create(DossierIni + 'fifo.ini');
  ini.WriteString('Langue', 'Langage', IntToStr(TTBXItem(sender).Tag));
  ini.free;
  If TAction(sender).Tag > 6 Then
    DM.siLangDispatcher1.ActiveLanguage := 2; //on met de l'English avant le turque pour les mots pas encore traduit
  DM.siLangDispatcher1.ActiveLanguage := TAction(sender).Tag;
  Case DM.siLangDispatcher1.ActiveLanguage Of
    0, 1:
      Begin
        SetFrenchResources;
      End;
    2:
      Begin
        SetEnglishResources;
      End;
    3:
      Begin
        SetGermanResources;
      End;
    5:
      Begin
        SetSpanishResources;
      End;
    7:
      Begin
        SetTurkishResources;
      End;
    8:
      Begin
        SetChineseResources;
      End;
  End;
End;

Procedure TFrmMain.ColorGridDblClickCell(Sender: TObject; ARow,
  ACol: Integer);
Var
  r, g, b, ri, gi, bi, i: integer;
  changed: boolean;
  nom: String;
  LL, aa, BB, R1, G1, B1: byte;
Begin
  If (Arow < 1) Or (ACol <> 1) Then
    exit;
  If TypeDeFil <> 4 Then
    Begin
      PercentContrast := ContrastEdit.value;
      i := Arow - 1;
      LABValue.b := tabcolor[i].moyen.rgbRed;
      LABValue.a := tabcolor[i].moyen.rgbGreen;
      LABValue.L := tabcolor[i].moyen.rgbBlue;
      cpickerform := Tcpickerform.create(self);
      cpickerform.Edit1.text := ColorGrid.Cells[Acol, Arow];
      If cpickerform.showmodal = mrOk Then
        Begin
          tabcolor[i].moyen.rgbRed := LABValue.b;
          tabcolor[i].moyen.rgbGreen := LABValue.a;
          tabcolor[i].moyen.rgbBlue := LABValue.L;
          tabcolor[i].moyen.rgbReserved := 255;

          LL := LABValue.L;
          AA := LABValue.a;
          BB := LABValue.b;

          hDisplayProfile.GetRGBFromLab(LL, AA, BB, r1, g1, b1);

          ColorGrid.Colors[Acol, Arow] := RGB(R1, G1, B1);
          ColorGrid.FontColors[Acol, Arow] := clwhite Xor RGB(R1, G1, B1);
          ColorGrid.Cells[Acol, Arow] := DragName;
          NomCouleur[i] := DragName;

          tabcolor[i].clair := Eclaircir(tabcolor[i].moyen, PercentContrast);
          tabcolor[i].fonce := Assombrir(tabcolor[i].moyen, PercentContrast);
          ActSimul.Enabled := true;
        End;
    End
  Else
    Begin
      PercentContrast := ContrastEdit.value;
      i := Arow - 1;
      If ody2t.Execute Then
        Begin
          If ImageFil[i] <> Nil Then
            FreeAndNil(ImageFil[i]);
          ImageFil[i] := TDibImage.Create;
          OpenTwistedFiber(ody2t.FileName, i);
          ColorGrid.Colors[Acol, Arow] := GetMediumColor(ImageFil[i], Trunc(fil_nbcolor.Value), TypeDeFil);
          ColorGrid.FontColors[Acol, Arow] := clwhite Xor ColorGrid.Colors[Acol, Arow];
          nom := ExtractfileName(ody2t.FileName);
          ColorGrid.Cells[Acol, Arow] := copy(nom, 1, length(nom) - 4);
          NomCouleur[Arow - 1] := ColorGrid.Cells[Acol, Arow];
          ActSimul.Enabled := true;
        End;
    End;
End;

Procedure TFrmMain.fil_nbcolorValueChanged(Sender: TObject);
Var
  NBR: integer;
Begin
  NBR := Trunc(fil_nbcolor.Value);
  If NBR < 0 Then
    NBR := 0;
  If NBR > 10 Then
    NBR := 10;
  ColorGrid.RowCount := NBR + 1;
  Case NBR Of
    1:
      If RG_sens_fil.Items.Count < 3 Then
        RG_sens_fil.Items.Add(siLangmain.GetTextOrDefaultW('IDS_9' (* 'Sans' *)));
    Else
      Begin
        If RG_sens_fil.Items.Count > 2 Then
          RG_sens_fil.Items.Delete(2);
        If (RG_sens_fil.ItemIndex < 0) Or (RG_sens_fil.ItemIndex > 1) Then
          RG_sens_fil.ItemIndex := 0;
      End;
  End;
End;

Procedure TFrmMain.ActSaveAsSmallUpdate(Sender: TObject);
Begin
  ActSaveAsSmall.Enabled := ActSaveAs.Enabled;
End;

Procedure TFrmMain.DoAboutProgress(NewPosition: Integer);
Begin
  If AboutForm <> Nil Then
    Begin
      AboutForm.Progress := NewPosition;
    End
  Else
    SetProgress(NewPosition);
End;

Procedure TFrmMain.RecentMenu1Click(Sender: TMenuItem;
  FileName: TFileName);
Begin
  OpenFil(Filename);
End;

Procedure TfrmMain.ActUpdateExecute(Sender: TObject);
Var
  i: integer;
Begin
  odfil.InitialDir := ExpandFileName(DataFolder + '\' + YarnFolder);
  odfil.Options := [ofAllowMultiSelect];
  If odfil.Execute Then
    For I := 0 To odfil.files.count - 1 Do
      Begin
        Try
          OpenFil(odfil.files[I]);
          SauverFil(odfil.files[I]);
        Except
          FrmMain.siLangmain.ShowMessage(siLangmain.GetTextOrDefault('IDS_77' (* 'Probleme sur le fil ' *)) + odfil.files[I]);
        End;
      End;
  odfil.Options := [];
End;

{*------------------------------------------------------------------------------
  Register file extension on Windows
  @return None   ResultDescription  
------------------------------------------------------------------------------*}
Procedure RegisterFile;
Begin
  With TRegistry.Create Do
    Try
      RootKey := HKEY_CLASSES_ROOT;
      OpenKey('.fil', True);
      WriteString('', 'fifo-Yarn file');
      CloseKey;
      OpenKey('fifo-Yarn file', True);
      WriteString('', 'fifo-Yarn file');
      OpenKey('DefaultIcon', True);
      WriteString('', ParamStr(0) + ',0');
      CloseKey;
      OpenKey('fifo-Yarn file\Shell\Open', True);
      WriteString('', '&Open');
      OpenKey('Command', True);
      WriteString('', ParamStr(0) + ' "%1"');
      CloseKey;

      SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, Nil, Nil);
    Finally
      Free;
    End;
End;

Initialization
  MsgDlgFormClass := TTNTForm;
  MsgDlgLabelClass := TTntLabel;
  MsgDlgEditClass := TTntEdit;
  MsgDlgButtonClass := TTntButton;

  RegisterFile;

End.
