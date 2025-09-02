// Von Klassen-Assistent automatisch erstellte IDispatch-Kapselungsklasse(n). 
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Adjustments 

class Adjustments : public COleDispatchDriver
{
public:
	Adjustments() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Adjustments(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Adjustments(const Adjustments& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	float GetItem(long Index);
	void SetItem(long Index, float newValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse CalloutFormat 

class CalloutFormat : public COleDispatchDriver
{
public:
	CalloutFormat() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	CalloutFormat(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CalloutFormat(const CalloutFormat& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void AutomaticLength();
	void CustomDrop(float Drop);
	void CustomLength(float Length);
	void PresetDrop(long DropType);
	long GetAccent();
	void SetAccent(long nNewValue);
	long GetAngle();
	void SetAngle(long nNewValue);
	long GetAutoAttach();
	void SetAutoAttach(long nNewValue);
	long GetAutoLength();
	long GetBorder();
	void SetBorder(long nNewValue);
	float GetDrop();
	long GetDropType();
	float GetGap();
	void SetGap(float newValue);
	float GetLength();
	long GetType();
	void SetType(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ColorFormat 

class ColorFormat : public COleDispatchDriver
{
public:
	ColorFormat() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ColorFormat(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ColorFormat(const ColorFormat& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetRgb();
	void SetRgb(long nNewValue);
	long GetSchemeColor();
	void SetSchemeColor(long nNewValue);
	long GetType();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse LineFormat 

class LineFormat : public COleDispatchDriver
{
public:
	LineFormat() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	LineFormat(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	LineFormat(const LineFormat& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBackColor();
	void SetBackColor(LPDISPATCH newValue);
	long GetBeginArrowheadLength();
	void SetBeginArrowheadLength(long nNewValue);
	long GetBeginArrowheadStyle();
	void SetBeginArrowheadStyle(long nNewValue);
	long GetBeginArrowheadWidth();
	void SetBeginArrowheadWidth(long nNewValue);
	long GetDashStyle();
	void SetDashStyle(long nNewValue);
	long GetEndArrowheadLength();
	void SetEndArrowheadLength(long nNewValue);
	long GetEndArrowheadStyle();
	void SetEndArrowheadStyle(long nNewValue);
	long GetEndArrowheadWidth();
	void SetEndArrowheadWidth(long nNewValue);
	LPDISPATCH GetForeColor();
	void SetForeColor(LPDISPATCH newValue);
	long GetPattern();
	void SetPattern(long nNewValue);
	long GetStyle();
	void SetStyle(long nNewValue);
	float GetTransparency();
	void SetTransparency(float newValue);
	long GetVisible();
	void SetVisible(long nNewValue);
	float GetWeight();
	void SetWeight(float newValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ShapeNode 

class ShapeNode : public COleDispatchDriver
{
public:
	ShapeNode() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ShapeNode(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ShapeNode(const ShapeNode& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetEditingType();
	VARIANT GetPoints();
	long GetSegmentType();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ShapeNodes 

class ShapeNodes : public COleDispatchDriver
{
public:
	ShapeNodes() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ShapeNodes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ShapeNodes(const ShapeNodes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN Get_NewEnum();
	void Delete(long Index);
	void Insert(long Index, long SegmentType, long EditingType, float X1, float Y1, float X2, float Y2, float X3, float Y3);
	void SetEditingType(long Index, long EditingType);
	void SetPosition(long Index, float X1, float Y1);
	void SetSegmentType(long Index, long SegmentType);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse PictureFormat 

class PictureFormat : public COleDispatchDriver
{
public:
	PictureFormat() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	PictureFormat(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	PictureFormat(const PictureFormat& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void IncrementBrightness(float Increment);
	void IncrementContrast(float Increment);
	float GetBrightness();
	void SetBrightness(float newValue);
	long GetColorType();
	void SetColorType(long nNewValue);
	float GetContrast();
	void SetContrast(float newValue);
	float GetCropBottom();
	void SetCropBottom(float newValue);
	float GetCropLeft();
	void SetCropLeft(float newValue);
	float GetCropRight();
	void SetCropRight(float newValue);
	float GetCropTop();
	void SetCropTop(float newValue);
	long GetTransparencyColor();
	void SetTransparencyColor(long nNewValue);
	long GetTransparentBackground();
	void SetTransparentBackground(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ShadowFormat 

class ShadowFormat : public COleDispatchDriver
{
public:
	ShadowFormat() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ShadowFormat(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ShadowFormat(const ShadowFormat& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void IncrementOffsetX(float Increment);
	void IncrementOffsetY(float Increment);
	LPDISPATCH GetForeColor();
	void SetForeColor(LPDISPATCH newValue);
	long GetObscured();
	void SetObscured(long nNewValue);
	float GetOffsetX();
	void SetOffsetX(float newValue);
	float GetOffsetY();
	void SetOffsetY(float newValue);
	float GetTransparency();
	void SetTransparency(float newValue);
	long GetType();
	void SetType(long nNewValue);
	long GetVisible();
	void SetVisible(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse TextEffectFormat 

class TextEffectFormat : public COleDispatchDriver
{
public:
	TextEffectFormat() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	TextEffectFormat(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	TextEffectFormat(const TextEffectFormat& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void ToggleVerticalText();
	long GetAlignment();
	void SetAlignment(long nNewValue);
	long GetFontBold();
	void SetFontBold(long nNewValue);
	long GetFontItalic();
	void SetFontItalic(long nNewValue);
	CString GetFontName();
	void SetFontName(LPCTSTR lpszNewValue);
	float GetFontSize();
	void SetFontSize(float newValue);
	long GetKernedPairs();
	void SetKernedPairs(long nNewValue);
	long GetNormalizedHeight();
	void SetNormalizedHeight(long nNewValue);
	long GetPresetShape();
	void SetPresetShape(long nNewValue);
	long GetPresetTextEffect();
	void SetPresetTextEffect(long nNewValue);
	long GetRotatedChars();
	void SetRotatedChars(long nNewValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	float GetTracking();
	void SetTracking(float newValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ThreeDFormat 

class ThreeDFormat : public COleDispatchDriver
{
public:
	ThreeDFormat() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ThreeDFormat(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ThreeDFormat(const ThreeDFormat& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void IncrementRotationX(float Increment);
	void IncrementRotationY(float Increment);
	void ResetRotation();
	void SetThreeDFormat(long PresetThreeDFormat);
	void SetExtrusionDirection(long PresetExtrusionDirection);
	float GetDepth();
	void SetDepth(float newValue);
	LPDISPATCH GetExtrusionColor();
	long GetExtrusionColorType();
	void SetExtrusionColorType(long nNewValue);
	long GetPerspective();
	void SetPerspective(long nNewValue);
	long GetPresetExtrusionDirection();
	long GetPresetLightingDirection();
	void SetPresetLightingDirection(long nNewValue);
	long GetPresetLightingSoftness();
	void SetPresetLightingSoftness(long nNewValue);
	long GetPresetMaterial();
	void SetPresetMaterial(long nNewValue);
	long GetPresetThreeDFormat();
	float GetRotationX();
	void SetRotationX(float newValue);
	float GetRotationY();
	void SetRotationY(float newValue);
	long GetVisible();
	void SetVisible(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse FillFormat 

class FillFormat : public COleDispatchDriver
{
public:
	FillFormat() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	FillFormat(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	FillFormat(const FillFormat& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void OneColorGradient(long Style, long Variant, float Degree);
	void Patterned(long Pattern);
	void PresetGradient(long Style, long Variant, long PresetGradientType);
	void PresetTextured(long PresetTexture);
	void Solid();
	void TwoColorGradient(long Style, long Variant);
	void UserPicture(LPCTSTR PictureFile);
	void UserTextured(LPCTSTR TextureFile);
	LPDISPATCH GetBackColor();
	void SetBackColor(LPDISPATCH newValue);
	LPDISPATCH GetForeColor();
	void SetForeColor(LPDISPATCH newValue);
	long GetGradientColorType();
	float GetGradientDegree();
	long GetGradientStyle();
	long GetGradientVariant();
	long GetPattern();
	long GetPresetGradientType();
	long GetPresetTexture();
	CString GetTextureName();
	long GetTextureType();
	float GetTransparency();
	void SetTransparency(float newValue);
	long GetType();
	long GetVisible();
	void SetVisible(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse _Application 

class _Application : public COleDispatchDriver
{
public:
	_Application() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	_Application(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	_Application(const _Application& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetActiveCell();
	LPDISPATCH GetActiveChart();
	CString GetActivePrinter();
	void SetActivePrinter(LPCTSTR lpszNewValue);
	LPDISPATCH GetActiveSheet();
	LPDISPATCH GetActiveWindow();
	LPDISPATCH GetActiveWorkbook();
	LPDISPATCH GetAddIns();
	LPDISPATCH GetAssistant();
	void Calculate();
	LPDISPATCH GetCells();
	LPDISPATCH GetCharts();
	LPDISPATCH GetColumns();
	LPDISPATCH GetCommandBars();
	long GetDDEAppReturnCode();
	void DDEExecute(long Channel, LPCTSTR String);
	long DDEInitiate(LPCTSTR App, LPCTSTR Topic);
	void DDEPoke(long Channel, const VARIANT& Item, const VARIANT& Data);
	VARIANT DDERequest(long Channel, LPCTSTR Item);
	void DDETerminate(long Channel);
	VARIANT Evaluate(const VARIANT& Name);
	VARIANT _Evaluate(const VARIANT& Name);
	VARIANT ExecuteExcel4Macro(LPCTSTR String);
	LPDISPATCH Intersect(LPDISPATCH Arg1, LPDISPATCH Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	LPDISPATCH GetNames();
	LPDISPATCH GetRange(const VARIANT& Cell1, const VARIANT& Cell2);
	LPDISPATCH GetRows();
	VARIANT Run(const VARIANT& Macro, const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, 
		const VARIANT& Arg10, const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, 
		const VARIANT& Arg20, const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, 
		const VARIANT& Arg30);
	VARIANT _Run2(const VARIANT& Macro, const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, 
		const VARIANT& Arg10, const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, 
		const VARIANT& Arg20, const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, 
		const VARIANT& Arg30);
	LPDISPATCH GetSelection();
	void SendKeys(const VARIANT& Keys, const VARIANT& Wait);
	LPDISPATCH GetSheets();
	LPDISPATCH GetThisWorkbook();
	LPDISPATCH Union(LPDISPATCH Arg1, LPDISPATCH Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, const VARIANT& Arg11, 
		const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, const VARIANT& Arg21, 
		const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	LPDISPATCH _Application::Union(LPDISPATCH Arg1, LPDISPATCH Arg2);
	LPDISPATCH GetWindows();
	LPDISPATCH GetWorkbooks();
	LPDISPATCH GetWorksheetFunction();
	LPDISPATCH GetWorksheets();
	LPDISPATCH GetExcel4IntlMacroSheets();
	LPDISPATCH GetExcel4MacroSheets();
	void ActivateMicrosoftApp(long Index);
	void AddChartAutoFormat(const VARIANT& Chart, LPCTSTR Name, const VARIANT& Description);
	void AddCustomList(const VARIANT& ListArray, const VARIANT& ByRow);
	BOOL GetAlertBeforeOverwriting();
	void SetAlertBeforeOverwriting(BOOL bNewValue);
	CString GetAltStartupPath();
	void SetAltStartupPath(LPCTSTR lpszNewValue);
	BOOL GetAskToUpdateLinks();
	void SetAskToUpdateLinks(BOOL bNewValue);
	BOOL GetEnableAnimations();
	void SetEnableAnimations(BOOL bNewValue);
	LPDISPATCH GetAutoCorrect();
	long GetBuild();
	BOOL GetCalculateBeforeSave();
	void SetCalculateBeforeSave(BOOL bNewValue);
	long GetCalculation();
	void SetCalculation(long nNewValue);
	VARIANT GetCaller(const VARIANT& Index);
	BOOL GetCanPlaySounds();
	BOOL GetCanRecordSounds();
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	BOOL GetCellDragAndDrop();
	void SetCellDragAndDrop(BOOL bNewValue);
	double CentimetersToPoints(double Centimeters);
	BOOL CheckSpelling(LPCTSTR Word, const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase);
	VARIANT GetClipboardFormats(const VARIANT& Index);
	BOOL GetDisplayClipboardWindow();
	void SetDisplayClipboardWindow(BOOL bNewValue);
	long GetCommandUnderlines();
	void SetCommandUnderlines(long nNewValue);
	BOOL GetConstrainNumeric();
	void SetConstrainNumeric(BOOL bNewValue);
	VARIANT ConvertFormula(const VARIANT& Formula, long FromReferenceStyle, const VARIANT& ToReferenceStyle, const VARIANT& ToAbsolute, const VARIANT& RelativeTo);
	BOOL GetCopyObjectsWithCells();
	void SetCopyObjectsWithCells(BOOL bNewValue);
	long GetCursor();
	void SetCursor(long nNewValue);
	long GetCustomListCount();
	long GetCutCopyMode();
	void SetCutCopyMode(long nNewValue);
	long GetDataEntryMode();
	void SetDataEntryMode(long nNewValue);
	CString Get_Default();
	CString GetDefaultFilePath();
	void SetDefaultFilePath(LPCTSTR lpszNewValue);
	void DeleteChartAutoFormat(LPCTSTR Name);
	void DeleteCustomList(long ListNum);
	LPDISPATCH GetDialogs();
	BOOL GetDisplayAlerts();
	void SetDisplayAlerts(BOOL bNewValue);
	BOOL GetDisplayFormulaBar();
	void SetDisplayFormulaBar(BOOL bNewValue);
	BOOL GetDisplayFullScreen();
	void SetDisplayFullScreen(BOOL bNewValue);
	BOOL GetDisplayNoteIndicator();
	void SetDisplayNoteIndicator(BOOL bNewValue);
	long GetDisplayCommentIndicator();
	void SetDisplayCommentIndicator(long nNewValue);
	BOOL GetDisplayExcel4Menus();
	void SetDisplayExcel4Menus(BOOL bNewValue);
	BOOL GetDisplayRecentFiles();
	void SetDisplayRecentFiles(BOOL bNewValue);
	BOOL GetDisplayScrollBars();
	void SetDisplayScrollBars(BOOL bNewValue);
	BOOL GetDisplayStatusBar();
	void SetDisplayStatusBar(BOOL bNewValue);
	void DoubleClick();
	BOOL GetEditDirectlyInCell();
	void SetEditDirectlyInCell(BOOL bNewValue);
	BOOL GetEnableAutoComplete();
	void SetEnableAutoComplete(BOOL bNewValue);
	long GetEnableCancelKey();
	void SetEnableCancelKey(long nNewValue);
	BOOL GetEnableSound();
	void SetEnableSound(BOOL bNewValue);
	VARIANT GetFileConverters(const VARIANT& Index1, const VARIANT& Index2);
	LPDISPATCH GetFileSearch();
	LPDISPATCH GetFileFind();
	void FindFile();
	BOOL GetFixedDecimal();
	void SetFixedDecimal(BOOL bNewValue);
	long GetFixedDecimalPlaces();
	void SetFixedDecimalPlaces(long nNewValue);
	VARIANT GetCustomListContents(long ListNum);
	long GetCustomListNum(const VARIANT& ListArray);
	VARIANT GetOpenFilename(const VARIANT& FileFilter, const VARIANT& FilterIndex, const VARIANT& Title, const VARIANT& ButtonText, const VARIANT& MultiSelect);
	VARIANT GetSaveAsFilename(const VARIANT& InitialFilename, const VARIANT& FileFilter, const VARIANT& FilterIndex, const VARIANT& Title, const VARIANT& ButtonText);
	void Goto(const VARIANT& Reference, const VARIANT& Scroll);
	double GetHeight();
	void SetHeight(double newValue);
	void Help(const VARIANT& HelpFile, const VARIANT& HelpContextID);
	BOOL GetIgnoreRemoteRequests();
	void SetIgnoreRemoteRequests(BOOL bNewValue);
	double InchesToPoints(double Inches);
	VARIANT InputBox(LPCTSTR Prompt, const VARIANT& Title, const VARIANT& Default, const VARIANT& Left, const VARIANT& Top, const VARIANT& HelpFile, const VARIANT& HelpContextID, const VARIANT& Type);
	BOOL GetInteractive();
	void SetInteractive(BOOL bNewValue);
	VARIANT GetInternational(const VARIANT& Index);
	BOOL GetIteration();
	void SetIteration(BOOL bNewValue);
	double GetLeft();
	void SetLeft(double newValue);
	CString GetLibraryPath();
	void MacroOptions(const VARIANT& Macro, const VARIANT& Description, const VARIANT& HasMenu, const VARIANT& MenuText, const VARIANT& HasShortcutKey, const VARIANT& ShortcutKey, const VARIANT& Category, const VARIANT& StatusBar, 
		const VARIANT& HelpContextID, const VARIANT& HelpFile);
	void MailLogoff();
	void MailLogon(const VARIANT& Name, const VARIANT& Password, const VARIANT& DownloadNewMail);
	VARIANT GetMailSession();
	long GetMailSystem();
	BOOL GetMathCoprocessorAvailable();
	double GetMaxChange();
	void SetMaxChange(double newValue);
	long GetMaxIterations();
	void SetMaxIterations(long nNewValue);
	long GetMemoryFree();
	long GetMemoryTotal();
	long GetMemoryUsed();
	BOOL GetMouseAvailable();
	BOOL GetMoveAfterReturn();
	void SetMoveAfterReturn(BOOL bNewValue);
	long GetMoveAfterReturnDirection();
	void SetMoveAfterReturnDirection(long nNewValue);
	LPDISPATCH GetRecentFiles();
	CString GetName();
	LPDISPATCH NextLetter();
	CString GetNetworkTemplatesPath();
	LPDISPATCH GetODBCErrors();
	long GetODBCTimeout();
	void SetODBCTimeout(long nNewValue);
	void OnKey(LPCTSTR Key, const VARIANT& Procedure);
	void OnRepeat(LPCTSTR Text, LPCTSTR Procedure);
	void OnTime(const VARIANT& EarliestTime, LPCTSTR Procedure, const VARIANT& LatestTime, const VARIANT& Schedule);
	void OnUndo(LPCTSTR Text, LPCTSTR Procedure);
	CString GetOnWindow();
	void SetOnWindow(LPCTSTR lpszNewValue);
	CString GetOperatingSystem();
	CString GetOrganizationName();
	CString GetPath();
	CString GetPathSeparator();
	VARIANT GetPreviousSelections(const VARIANT& Index);
	BOOL GetPivotTableSelection();
	void SetPivotTableSelection(BOOL bNewValue);
	BOOL GetPromptForSummaryInfo();
	void SetPromptForSummaryInfo(BOOL bNewValue);
	void Quit();
	void RecordMacro(const VARIANT& BasicCode, const VARIANT& XlmCode);
	BOOL GetRecordRelative();
	long GetReferenceStyle();
	void SetReferenceStyle(long nNewValue);
	VARIANT GetRegisteredFunctions(const VARIANT& Index1, const VARIANT& Index2);
	BOOL RegisterXLL(LPCTSTR Filename);
	void Repeat();
	BOOL GetRollZoom();
	void SetRollZoom(BOOL bNewValue);
	void SaveWorkspace(const VARIANT& Filename);
	BOOL GetScreenUpdating();
	void SetScreenUpdating(BOOL bNewValue);
	void SetDefaultChart(const VARIANT& FormatName, const VARIANT& Gallery);
	long GetSheetsInNewWorkbook();
	void SetSheetsInNewWorkbook(long nNewValue);
	BOOL GetShowChartTipNames();
	void SetShowChartTipNames(BOOL bNewValue);
	BOOL GetShowChartTipValues();
	void SetShowChartTipValues(BOOL bNewValue);
	CString GetStandardFont();
	void SetStandardFont(LPCTSTR lpszNewValue);
	double GetStandardFontSize();
	void SetStandardFontSize(double newValue);
	CString GetStartupPath();
	VARIANT GetStatusBar();
	void SetStatusBar(const VARIANT& newValue);
	CString GetTemplatesPath();
	BOOL GetShowToolTips();
	void SetShowToolTips(BOOL bNewValue);
	double GetTop();
	void SetTop(double newValue);
	long GetDefaultSaveFormat();
	void SetDefaultSaveFormat(long nNewValue);
	CString GetTransitionMenuKey();
	void SetTransitionMenuKey(LPCTSTR lpszNewValue);
	long GetTransitionMenuKeyAction();
	void SetTransitionMenuKeyAction(long nNewValue);
	BOOL GetTransitionNavigKeys();
	void SetTransitionNavigKeys(BOOL bNewValue);
	void Undo();
	double GetUsableHeight();
	double GetUsableWidth();
	BOOL GetUserControl();
	void SetUserControl(BOOL bNewValue);
	CString GetUserName_();
	void SetUserName(LPCTSTR lpszNewValue);
	CString GetValue();
	LPDISPATCH GetVbe();
	CString GetVersion();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	void Volatile(const VARIANT& Volatile);
	void Wait(const VARIANT& Time);
	double GetWidth();
	void SetWidth(double newValue);
	BOOL GetWindowsForPens();
	long GetWindowState();
	void SetWindowState(long nNewValue);
	long GetUILanguage();
	void SetUILanguage(long nNewValue);
	long GetDefaultSheetDirection();
	void SetDefaultSheetDirection(long nNewValue);
	long GetCursorMovement();
	void SetCursorMovement(long nNewValue);
	long GetControlCharacters();
	void SetControlCharacters(long nNewValue);
	BOOL GetEnableEvents();
	void SetEnableEvents(BOOL bNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse _Chart 

class _Chart : public COleDispatchDriver
{
public:
	_Chart() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	_Chart(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	_Chart(const _Chart& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void Activate();
	void Copy(const VARIANT& Before, const VARIANT& After);
	void Delete();
	CString GetCodeName();
	CString Get_CodeName();
	void Set_CodeName(LPCTSTR lpszNewValue);
	long GetIndex();
	void Move(const VARIANT& Before, const VARIANT& After);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	LPDISPATCH GetNext();
	LPDISPATCH GetPageSetup();
	LPDISPATCH GetPrevious();
	void PrintOut(const VARIANT& From, const VARIANT& To, const VARIANT& Copies, const VARIANT& Preview, const VARIANT& ActivePrinter, const VARIANT& PrintToFile, const VARIANT& Collate);
	void PrintPreview(const VARIANT& EnableChanges);
	void Protect(const VARIANT& Password, const VARIANT& DrawingObjects, const VARIANT& Contents, const VARIANT& Scenarios, const VARIANT& UserInterfaceOnly);
	BOOL GetProtectContents();
	BOOL GetProtectDrawingObjects();
	BOOL GetProtectionMode();
	void SaveAs(LPCTSTR Filename, const VARIANT& FileFormat, const VARIANT& Password, const VARIANT& WriteResPassword, const VARIANT& ReadOnlyRecommended, const VARIANT& CreateBackup, const VARIANT& AddToMru, const VARIANT& TextCodepage, 
		const VARIANT& TextVisualLayout);
	void Select(const VARIANT& Replace);
	void Unprotect(const VARIANT& Password);
	long GetVisible();
	void SetVisible(long nNewValue);
	LPDISPATCH GetShapes();
	void ApplyDataLabels(long Type, const VARIANT& LegendKey, const VARIANT& AutoText, const VARIANT& HasLeaderLines);
	LPDISPATCH GetArea3DGroup();
	LPDISPATCH AreaGroups(const VARIANT& Index);
	BOOL GetAutoScaling();
	void SetAutoScaling(BOOL bNewValue);
	LPDISPATCH Axes(const VARIANT& Type, long AxisGroup);
	void SetBackgroundPicture(LPCTSTR Filename);
	LPDISPATCH GetBar3DGroup();
	LPDISPATCH BarGroups(const VARIANT& Index);
	LPDISPATCH GetChartArea();
	LPDISPATCH ChartGroups(const VARIANT& Index);
	LPDISPATCH ChartObjects(const VARIANT& Index);
	LPDISPATCH GetChartTitle();
	void ChartWizard(const VARIANT& Source, const VARIANT& Gallery, const VARIANT& Format, const VARIANT& PlotBy, const VARIANT& CategoryLabels, const VARIANT& SeriesLabels, const VARIANT& HasLegend, const VARIANT& Title, 
		const VARIANT& CategoryTitle, const VARIANT& ValueTitle, const VARIANT& ExtraTitle);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	LPDISPATCH GetColumn3DGroup();
	LPDISPATCH ColumnGroups(const VARIANT& Index);
	void CopyPicture(long Appearance, long Format, long Size);
	LPDISPATCH GetCorners();
	void CreatePublisher(const VARIANT& Edition, long Appearance, long Size, const VARIANT& ContainsPICT, const VARIANT& ContainsBIFF, const VARIANT& ContainsRTF, const VARIANT& ContainsVALU);
	LPDISPATCH GetDataTable();
	long GetDepthPercent();
	void SetDepthPercent(long nNewValue);
	void Deselect();
	long GetDisplayBlanksAs();
	void SetDisplayBlanksAs(long nNewValue);
	LPDISPATCH DoughnutGroups(const VARIANT& Index);
	long GetElevation();
	void SetElevation(long nNewValue);
	VARIANT Evaluate(const VARIANT& Name);
	VARIANT _Evaluate(const VARIANT& Name);
	LPDISPATCH GetFloor();
	long GetGapDepth();
	void SetGapDepth(long nNewValue);
	VARIANT GetHasAxis(const VARIANT& Index1, const VARIANT& Index2);
	void SetHasAxis(const VARIANT& Index1, const VARIANT& Index2, const VARIANT& newValue);
	BOOL GetHasDataTable();
	void SetHasDataTable(BOOL bNewValue);
	BOOL GetHasLegend();
	void SetHasLegend(BOOL bNewValue);
	BOOL GetHasTitle();
	void SetHasTitle(BOOL bNewValue);
	long GetHeightPercent();
	void SetHeightPercent(long nNewValue);
	LPDISPATCH GetHyperlinks();
	LPDISPATCH GetLegend();
	LPDISPATCH GetLine3DGroup();
	LPDISPATCH LineGroups(const VARIANT& Index);
	LPDISPATCH Location(long Where, const VARIANT& Name);
	LPDISPATCH OLEObjects(const VARIANT& Index);
	void Paste(const VARIANT& Type);
	long GetPerspective();
	void SetPerspective(long nNewValue);
	LPDISPATCH GetPie3DGroup();
	LPDISPATCH PieGroups(const VARIANT& Index);
	LPDISPATCH GetPlotArea();
	BOOL GetPlotVisibleOnly();
	void SetPlotVisibleOnly(BOOL bNewValue);
	LPDISPATCH RadarGroups(const VARIANT& Index);
	VARIANT GetRightAngleAxes();
	void SetRightAngleAxes(const VARIANT& newValue);
	VARIANT GetRotation();
	void SetRotation(const VARIANT& newValue);
	LPDISPATCH SeriesCollection(const VARIANT& Index);
	BOOL GetSizeWithWindow();
	void SetSizeWithWindow(BOOL bNewValue);
	BOOL GetShowWindow();
	void SetShowWindow(BOOL bNewValue);
	LPDISPATCH GetSurfaceGroup();
	long GetChartType();
	void SetChartType(long nNewValue);
	void ApplyCustomType(long ChartType, const VARIANT& TypeName);
	LPDISPATCH GetWalls();
	BOOL GetWallsAndGridlines2D();
	void SetWallsAndGridlines2D(BOOL bNewValue);
	LPDISPATCH XYGroups(const VARIANT& Index);
	long GetBarShape();
	void SetBarShape(long nNewValue);
	long GetPlotBy();
	void SetPlotBy(long nNewValue);
	BOOL GetProtectFormatting();
	void SetProtectFormatting(BOOL bNewValue);
	BOOL GetProtectData();
	void SetProtectData(BOOL bNewValue);
	BOOL GetProtectGoalSeek();
	void SetProtectGoalSeek(BOOL bNewValue);
	BOOL GetProtectSelection();
	void SetProtectSelection(BOOL bNewValue);
	void GetChartElement(long X, long Y, long* ElementID, long* Arg1, long* Arg2);
	void SetSourceData(LPDISPATCH Source, const VARIANT& PlotBy);
	BOOL Export(LPCTSTR Filename, const VARIANT& FilterName, const VARIANT& Interactive);
	void Refresh();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Sheets 

class Sheets : public COleDispatchDriver
{
public:
	Sheets() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Sheets(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Sheets(const Sheets& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH Add(const VARIANT& Before, const VARIANT& After, const VARIANT& Count, const VARIANT& Type);
	void Copy(const VARIANT& Before, const VARIANT& After);
	long GetCount();
	void Delete();
	void FillAcrossSheets(LPDISPATCH Range, long Type);
	LPDISPATCH GetItem(const VARIANT& Index);
	void Move(const VARIANT& Before, const VARIANT& After);
	LPUNKNOWN Get_NewEnum();
	void PrintOut(const VARIANT& From, const VARIANT& To, const VARIANT& Copies, const VARIANT& Preview, const VARIANT& ActivePrinter, const VARIANT& PrintToFile, const VARIANT& Collate);
	void PrintPreview(const VARIANT& EnableChanges);
	void Select(const VARIANT& Replace);
	LPDISPATCH GetHPageBreaks();
	LPDISPATCH GetVPageBreaks();
	VARIANT GetVisible();
	void SetVisible(const VARIANT& newValue);
	LPDISPATCH Get_Default(const VARIANT& Index);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse _Worksheet 

class _Worksheet : public COleDispatchDriver
{
public:
	_Worksheet() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	_Worksheet(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	_Worksheet(const _Worksheet& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void Activate();
	void Copy(const VARIANT& Before, const VARIANT& After);
	void Delete();
	CString GetCodeName();
	CString Get_CodeName();
	void Set_CodeName(LPCTSTR lpszNewValue);
	long GetIndex();
	void Move(const VARIANT& Before, const VARIANT& After);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	LPDISPATCH GetNext();
	LPDISPATCH GetPageSetup();
	LPDISPATCH GetPrevious();
	void PrintOut(const VARIANT& From, const VARIANT& To, const VARIANT& Copies, const VARIANT& Preview, const VARIANT& ActivePrinter, const VARIANT& PrintToFile, const VARIANT& Collate);
	void PrintPreview(const VARIANT& EnableChanges);
	void Protect(const VARIANT& Password, const VARIANT& DrawingObjects, const VARIANT& Contents, const VARIANT& Scenarios, const VARIANT& UserInterfaceOnly);
	BOOL GetProtectContents();
	BOOL GetProtectDrawingObjects();
	BOOL GetProtectionMode();
	BOOL GetProtectScenarios();
	void SaveAs(LPCTSTR Filename, const VARIANT& FileFormat, const VARIANT& Password, const VARIANT& WriteResPassword, const VARIANT& ReadOnlyRecommended, const VARIANT& CreateBackup, const VARIANT& AddToMru, const VARIANT& TextCodepage, 
		const VARIANT& TextVisualLayout);
	void Select(const VARIANT& Replace);
	void Unprotect(const VARIANT& Password);
	long GetVisible();
	void SetVisible(long nNewValue);
	LPDISPATCH GetShapes();
	BOOL GetTransitionExpEval();
	void SetTransitionExpEval(BOOL bNewValue);
	BOOL GetAutoFilterMode();
	void SetAutoFilterMode(BOOL bNewValue);
	void SetBackgroundPicture(LPCTSTR Filename);
	void Calculate();
	BOOL GetEnableCalculation();
	void SetEnableCalculation(BOOL bNewValue);
	LPDISPATCH GetCells();
	LPDISPATCH ChartObjects(const VARIANT& Index);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	LPDISPATCH GetCircularReference();
	void ClearArrows();
	LPDISPATCH GetColumns();
	long GetConsolidationFunction();
	VARIANT GetConsolidationOptions();
	VARIANT GetConsolidationSources();
	BOOL GetEnableAutoFilter();
	void SetEnableAutoFilter(BOOL bNewValue);
	long GetEnableSelection();
	void SetEnableSelection(long nNewValue);
	BOOL GetEnableOutlining();
	void SetEnableOutlining(BOOL bNewValue);
	BOOL GetEnablePivotTable();
	void SetEnablePivotTable(BOOL bNewValue);
	VARIANT Evaluate(const VARIANT& Name);
	VARIANT _Evaluate(const VARIANT& Name);
	BOOL GetFilterMode();
	void ResetAllPageBreaks();
	LPDISPATCH GetNames();
	LPDISPATCH OLEObjects(const VARIANT& Index);
	LPDISPATCH GetOutline();
	void Paste(const VARIANT& Destination, const VARIANT& Link);
	void Paste();
	void PasteSpecial_New(VARIANT& Format, VARIANT& Link, VARIANT& DisplayAsIcon, VARIANT& IconFileName, VARIANT& IconIndex, VARIANT& IconLabel, VARIANT& NoHTMLFormatting);
	void PasteSpecial_New();
	void PasteSpecial(const VARIANT& Format, const VARIANT& Link, const VARIANT& DisplayAsIcon, const VARIANT& IconFileName, const VARIANT& IconIndex, const VARIANT& IconLabel);
	void PasteSpecial();
	void PasteSpecial(const VARIANT& Format);
	LPDISPATCH PivotTables(const VARIANT& Index);
	LPDISPATCH PivotTableWizard(const VARIANT& SourceType, const VARIANT& SourceData, const VARIANT& TableDestination, const VARIANT& TableName, const VARIANT& RowGrand, const VARIANT& ColumnGrand, const VARIANT& SaveData, 
		const VARIANT& HasAutoFormat, const VARIANT& AutoPage, const VARIANT& Reserved, const VARIANT& BackgroundQuery, const VARIANT& OptimizeCache, const VARIANT& PageFieldOrder, const VARIANT& PageFieldWrapCount, const VARIANT& ReadData, 
		const VARIANT& Connection);
	LPDISPATCH GetRange(const VARIANT& Cell1, const VARIANT& Cell2);
	LPDISPATCH GetRows();
	LPDISPATCH Scenarios(const VARIANT& Index);
	CString GetScrollArea();
	void SetScrollArea(LPCTSTR lpszNewValue);
	void ShowAllData();
	void ShowDataForm();
	double GetStandardHeight();
	double GetStandardWidth();
	void SetStandardWidth(double newValue);
	BOOL GetTransitionFormEntry();
	void SetTransitionFormEntry(BOOL bNewValue);
	long GetType();
	LPDISPATCH GetUsedRange();
	LPDISPATCH GetHPageBreaks();
	LPDISPATCH GetVPageBreaks();
	LPDISPATCH GetQueryTables();
	BOOL GetDisplayPageBreaks();
	void SetDisplayPageBreaks(BOOL bNewValue);
	LPDISPATCH GetComments();
	LPDISPATCH GetHyperlinks();
	void ClearCircles();
	void CircleInvalid();
	LPDISPATCH GetAutoFilter();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse _Global 

class _Global : public COleDispatchDriver
{
public:
	_Global() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	_Global(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	_Global(const _Global& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetActiveCell();
	LPDISPATCH GetActiveChart();
	CString GetActivePrinter();
	void SetActivePrinter(LPCTSTR lpszNewValue);
	LPDISPATCH GetActiveSheet();
	LPDISPATCH GetActiveWindow();
	LPDISPATCH GetActiveWorkbook();
	LPDISPATCH GetAddIns();
	LPDISPATCH GetAssistant();
	void Calculate();
	LPDISPATCH GetCells();
	LPDISPATCH GetCharts();
	LPDISPATCH GetColumns();
	LPDISPATCH GetCommandBars();
	long GetDDEAppReturnCode();
	void DDEExecute(long Channel, LPCTSTR String);
	long DDEInitiate(LPCTSTR App, LPCTSTR Topic);
	void DDEPoke(long Channel, const VARIANT& Item, const VARIANT& Data);
	VARIANT DDERequest(long Channel, LPCTSTR Item);
	void DDETerminate(long Channel);
	VARIANT Evaluate(const VARIANT& Name);
	VARIANT _Evaluate(const VARIANT& Name);
	VARIANT ExecuteExcel4Macro(LPCTSTR String);
	LPDISPATCH Intersect(LPDISPATCH Arg1, LPDISPATCH Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	LPDISPATCH GetNames();
	LPDISPATCH GetRange(const VARIANT& Cell1, const VARIANT& Cell2);
	LPDISPATCH GetRows();
	VARIANT Run(const VARIANT& Macro, const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, 
		const VARIANT& Arg10, const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, 
		const VARIANT& Arg20, const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, 
		const VARIANT& Arg30);
	VARIANT _Run2(const VARIANT& Macro, const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, 
		const VARIANT& Arg10, const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, 
		const VARIANT& Arg20, const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, 
		const VARIANT& Arg30);
	LPDISPATCH GetSelection();
	void SendKeys(const VARIANT& Keys, const VARIANT& Wait);
	LPDISPATCH GetSheets();
	LPDISPATCH GetThisWorkbook();
	LPDISPATCH Union(LPDISPATCH Arg1, LPDISPATCH Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, const VARIANT& Arg11, 
		const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, const VARIANT& Arg21, 
		const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	LPDISPATCH GetWindows();
	LPDISPATCH GetWorkbooks();
	LPDISPATCH GetWorksheetFunction();
	LPDISPATCH GetWorksheets();
	LPDISPATCH GetExcel4IntlMacroSheets();
	LPDISPATCH GetExcel4MacroSheets();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse _Workbook 

class _Workbook : public COleDispatchDriver
{
public:
	_Workbook() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	_Workbook(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	_Workbook(const _Workbook& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	BOOL GetAcceptLabelsInFormulas();
	void SetAcceptLabelsInFormulas(BOOL bNewValue);
	void Activate();
	LPDISPATCH GetActiveChart();
	LPDISPATCH GetActiveSheet();
	long GetAutoUpdateFrequency();
	void SetAutoUpdateFrequency(long nNewValue);
	BOOL GetAutoUpdateSaveChanges();
	void SetAutoUpdateSaveChanges(BOOL bNewValue);
	long GetChangeHistoryDuration();
	void SetChangeHistoryDuration(long nNewValue);
	LPDISPATCH GetBuiltinDocumentProperties();
	void ChangeFileAccess(long Mode, const VARIANT& WritePassword, const VARIANT& Notify);
	void ChangeLink(LPCTSTR Name, LPCTSTR NewName, long Type);
	LPDISPATCH GetCharts();
	void Close(const VARIANT& SaveChanges, const VARIANT& Filename, const VARIANT& RouteWorkbook);
	CString GetCodeName();
	CString Get_CodeName();
	void Set_CodeName(LPCTSTR lpszNewValue);
	VARIANT GetColors(const VARIANT& Index);
	void SetColors(const VARIANT& Index, const VARIANT& newValue);
	LPDISPATCH GetCommandBars();
	long GetConflictResolution();
	void SetConflictResolution(long nNewValue);
	LPDISPATCH GetContainer();
	BOOL GetCreateBackup();
	LPDISPATCH GetCustomDocumentProperties();
	BOOL GetDate1904();
	void SetDate1904(BOOL bNewValue);
	void DeleteNumberFormat(LPCTSTR NumberFormat);
	long GetDisplayDrawingObjects();
	void SetDisplayDrawingObjects(long nNewValue);
	BOOL ExclusiveAccess();
	long GetFileFormat();
	void ForwardMailer();
	CString GetFullName();
	BOOL GetHasPassword();
	BOOL GetHasRoutingSlip();
	void SetHasRoutingSlip(BOOL bNewValue);
	BOOL GetIsAddin();
	void SetIsAddin(BOOL bNewValue);
	VARIANT LinkInfo(LPCTSTR Name, long LinkInfo, const VARIANT& Type, const VARIANT& EditionRef);
	VARIANT LinkSources(const VARIANT& Type);
	LPDISPATCH GetMailer();
	void MergeWorkbook(const VARIANT& Filename);
	BOOL GetMultiUserEditing();
	CString GetName();
	LPDISPATCH GetNames();
	LPDISPATCH NewWindow();
	void OpenLinks(LPCTSTR Name, const VARIANT& ReadOnly, const VARIANT& Type);
	CString GetPath();
	BOOL GetPersonalViewListSettings();
	void SetPersonalViewListSettings(BOOL bNewValue);
	BOOL GetPersonalViewPrintSettings();
	void SetPersonalViewPrintSettings(BOOL bNewValue);
	LPDISPATCH PivotCaches();
	void Post(const VARIANT& DestName);
	BOOL GetPrecisionAsDisplayed();
	void SetPrecisionAsDisplayed(BOOL bNewValue);
	void PrintOut(const VARIANT& From, const VARIANT& To, const VARIANT& Copies, const VARIANT& Preview, const VARIANT& ActivePrinter, const VARIANT& PrintToFile, const VARIANT& Collate);
	void PrintPreview(const VARIANT& EnableChanges);
	void Protect(const VARIANT& Password, const VARIANT& Structure, const VARIANT& Windows);
	void ProtectSharing(const VARIANT& Filename, const VARIANT& Password, const VARIANT& WriteResPassword, const VARIANT& ReadOnlyRecommended, const VARIANT& CreateBackup, const VARIANT& SharingPassword);
	BOOL GetProtectStructure();
	BOOL GetProtectWindows();
	BOOL GetReadOnly();
	BOOL GetReadOnlyRecommended();
	void RefreshAll();
	void Reply();
	void ReplyAll();
	void RemoveUser(long Index);
	long GetRevisionNumber();
	void Route();
	BOOL GetRouted();
	LPDISPATCH GetRoutingSlip();
	void RunAutoMacros(long Which);
	void Save();
	void SaveAs(const VARIANT& Filename, const VARIANT& FileFormat, const VARIANT& Password, const VARIANT& WriteResPassword, const VARIANT& ReadOnlyRecommended, const VARIANT& CreateBackup, long AccessMode, const VARIANT& ConflictResolution, 
		const VARIANT& AddToMru, const VARIANT& TextCodepage, const VARIANT& TextVisualLayout);
	void SaveCopyAs(const VARIANT& Filename);
	BOOL GetSaved();
	void SetSaved(BOOL bNewValue);
	BOOL GetSaveLinkValues();
	void SetSaveLinkValues(BOOL bNewValue);
	void SendMail(const VARIANT& Recipients, const VARIANT& Subject, const VARIANT& ReturnReceipt);
	void SendMailer(const VARIANT& FileFormat, long Priority);
	void SetLinkOnData(LPCTSTR Name, const VARIANT& Procedure);
	LPDISPATCH GetSheets();
	BOOL GetShowConflictHistory();
	void SetShowConflictHistory(BOOL bNewValue);
	LPDISPATCH GetStyles();
	void Unprotect(const VARIANT& Password);
	void UnprotectSharing(const VARIANT& SharingPassword);
	void UpdateFromFile();
	void UpdateLink(const VARIANT& Name, const VARIANT& Type);
	BOOL GetUpdateRemoteReferences();
	void SetUpdateRemoteReferences(BOOL bNewValue);
	VARIANT GetUserStatus();
	LPDISPATCH GetCustomViews();
	LPDISPATCH GetWindows();
	LPDISPATCH GetWorksheets();
	BOOL GetWriteReserved();
	CString GetWriteReservedBy();
	LPDISPATCH GetExcel4IntlMacroSheets();
	LPDISPATCH GetExcel4MacroSheets();
	BOOL GetTemplateRemoveExtData();
	void SetTemplateRemoveExtData(BOOL bNewValue);
	void HighlightChangesOptions(const VARIANT& When, const VARIANT& Who, const VARIANT& Where);
	BOOL GetHighlightChangesOnScreen();
	void SetHighlightChangesOnScreen(BOOL bNewValue);
	BOOL GetKeepChangeHistory();
	void SetKeepChangeHistory(BOOL bNewValue);
	BOOL GetListChangesOnNewSheet();
	void SetListChangesOnNewSheet(BOOL bNewValue);
	void PurgeChangeHistoryNow(long Days, const VARIANT& SharingPassword);
	void AcceptAllChanges(const VARIANT& When, const VARIANT& Who, const VARIANT& Where);
	void RejectAllChanges(const VARIANT& When, const VARIANT& Who, const VARIANT& Where);
	void ResetColors();
	LPDISPATCH GetVBProject();
	void FollowHyperlink(LPCTSTR Address, const VARIANT& SubAddress, const VARIANT& NewWindow, const VARIANT& AddHistory, const VARIANT& ExtraInfo, const VARIANT& Method, const VARIANT& HeaderInfo);
	void AddToFavorites();
	BOOL GetIsInplace();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Workbooks 

class Workbooks : public COleDispatchDriver
{
public:
	Workbooks() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Workbooks(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Workbooks(const Workbooks& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH Add(const VARIANT& Template);
	LPDISPATCH Add();
	void Close();
	long GetCount();
	LPDISPATCH GetItem(const VARIANT& Index);
	LPUNKNOWN Get_NewEnum();
	LPDISPATCH Open(LPCTSTR Filename, const VARIANT& UpdateLinks, const VARIANT& ReadOnly, const VARIANT& Format, const VARIANT& Password, const VARIANT& WriteResPassword, const VARIANT& IgnoreReadOnlyRecommended, const VARIANT& Origin, 
		const VARIANT& Delimiter, const VARIANT& Editable, const VARIANT& Notify, const VARIANT& Converter, const VARIANT& AddToMru);
	void OpenText(LPCTSTR Filename, const VARIANT& Origin, const VARIANT& StartRow, const VARIANT& DataType, long TextQualifier, const VARIANT& ConsecutiveDelimiter, const VARIANT& Tab, const VARIANT& Semicolon, const VARIANT& Comma, 
		const VARIANT& Space, const VARIANT& Other, const VARIANT& OtherChar, const VARIANT& FieldInfo, const VARIANT& TextVisualLayout);
	LPDISPATCH Get_Default(const VARIANT& Index);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Font 

class Font : public COleDispatchDriver
{
public:
	Font() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Font(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Font(const Font& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	VARIANT GetBackground();
	void SetBackground(const VARIANT& newValue);
	VARIANT GetBold();
	void SetBold(const VARIANT& newValue);
	VARIANT GetColor();
	void SetColor(const VARIANT& newValue);
	VARIANT GetColorIndex();
	void SetColorIndex(const VARIANT& newValue);
	VARIANT GetFontStyle();
	void SetFontStyle(const VARIANT& newValue);
	VARIANT GetItalic();
	void SetItalic(const VARIANT& newValue);
	VARIANT GetName();
	void SetName(const VARIANT& newValue);
	VARIANT GetOutlineFont();
	void SetOutlineFont(const VARIANT& newValue);
	VARIANT GetShadow();
	void SetShadow(const VARIANT& newValue);
	VARIANT GetSize();
	void SetSize(const VARIANT& newValue);
	VARIANT GetStrikethrough();
	void SetStrikethrough(const VARIANT& newValue);
	VARIANT GetSubscript();
	void SetSubscript(const VARIANT& newValue);
	VARIANT GetSuperscript();
	void SetSuperscript(const VARIANT& newValue);
	VARIANT GetUnderline();
	void SetUnderline(const VARIANT& newValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Window 

class Window : public COleDispatchDriver
{
public:
	Window() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Window(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Window(const Window& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void Activate();
	void ActivateNext();
	void ActivatePrevious();
	LPDISPATCH GetActiveCell();
	LPDISPATCH GetActiveChart();
	LPDISPATCH GetActivePane();
	LPDISPATCH GetActiveSheet();
	VARIANT GetCaption();
	void SetCaption(const VARIANT& newValue);
	void Close(const VARIANT& SaveChanges, const VARIANT& Filename, const VARIANT& RouteWorkbook);
	BOOL GetDisplayFormulas();
	void SetDisplayFormulas(BOOL bNewValue);
	BOOL GetDisplayGridlines();
	void SetDisplayGridlines(BOOL bNewValue);
	BOOL GetDisplayHeadings();
	void SetDisplayHeadings(BOOL bNewValue);
	BOOL GetDisplayHorizontalScrollBar();
	void SetDisplayHorizontalScrollBar(BOOL bNewValue);
	BOOL GetDisplayOutline();
	void SetDisplayOutline(BOOL bNewValue);
	BOOL GetDisplayRightToLeft();
	void SetDisplayRightToLeft(BOOL bNewValue);
	BOOL GetDisplayVerticalScrollBar();
	void SetDisplayVerticalScrollBar(BOOL bNewValue);
	BOOL GetDisplayWorkbookTabs();
	void SetDisplayWorkbookTabs(BOOL bNewValue);
	BOOL GetDisplayZeros();
	void SetDisplayZeros(BOOL bNewValue);
	BOOL GetEnableResize();
	void SetEnableResize(BOOL bNewValue);
	BOOL GetFreezePanes();
	void SetFreezePanes(BOOL bNewValue);
	long GetGridlineColor();
	void SetGridlineColor(long nNewValue);
	long GetGridlineColorIndex();
	void SetGridlineColorIndex(long nNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	long GetIndex();
	void LargeScroll(const VARIANT& Down, const VARIANT& Up, const VARIANT& ToRight, const VARIANT& ToLeft);
	double GetLeft();
	void SetLeft(double newValue);
	LPDISPATCH NewWindow();
	CString GetOnWindow();
	void SetOnWindow(LPCTSTR lpszNewValue);
	LPDISPATCH GetPanes();
	void PrintOut(const VARIANT& From, const VARIANT& To, const VARIANT& Copies, const VARIANT& Preview, const VARIANT& ActivePrinter, const VARIANT& PrintToFile, const VARIANT& Collate);
	void PrintPreview(const VARIANT& EnableChanges);
	LPDISPATCH GetRangeSelection();
	long GetScrollColumn();
	void SetScrollColumn(long nNewValue);
	long GetScrollRow();
	void SetScrollRow(long nNewValue);
	void ScrollWorkbookTabs(const VARIANT& Sheets, const VARIANT& Position);
	LPDISPATCH GetSelectedSheets();
	LPDISPATCH GetSelection();
	void SmallScroll(const VARIANT& Down, const VARIANT& Up, const VARIANT& ToRight, const VARIANT& ToLeft);
	BOOL GetSplit();
	void SetSplit(BOOL bNewValue);
	long GetSplitColumn();
	void SetSplitColumn(long nNewValue);
	double GetSplitHorizontal();
	void SetSplitHorizontal(double newValue);
	long GetSplitRow();
	void SetSplitRow(long nNewValue);
	double GetSplitVertical();
	void SetSplitVertical(double newValue);
	double GetTabRatio();
	void SetTabRatio(double newValue);
	double GetTop();
	void SetTop(double newValue);
	long GetType();
	double GetUsableHeight();
	double GetUsableWidth();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	LPDISPATCH GetVisibleRange();
	double GetWidth();
	void SetWidth(double newValue);
	long GetWindowNumber();
	long GetWindowState();
	void SetWindowState(long nNewValue);
	VARIANT GetZoom();
	void SetZoom(const VARIANT& newValue);
	long GetView();
	void SetView(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Windows 

class Windows : public COleDispatchDriver
{
public:
	Windows() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Windows(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Windows(const Windows& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void Arrange(long ArrangeStyle, const VARIANT& ActiveWorkbook, const VARIANT& SyncHorizontal, const VARIANT& SyncVertical);
	long GetCount();
	LPDISPATCH GetItem(const VARIANT& Index);
	LPUNKNOWN Get_NewEnum();
	LPDISPATCH Get_Default(const VARIANT& Index);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse AppEvents 

class AppEvents : public COleDispatchDriver
{
public:
	AppEvents() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	AppEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	AppEvents(const AppEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	void NewWorkbook(LPDISPATCH Wb);
	void SheetSelectionChange(LPDISPATCH Sh, LPDISPATCH Target);
	void SheetBeforeDoubleClick(LPDISPATCH Sh, LPDISPATCH Target, BOOL* Cancel);
	void SheetBeforeRightClick(LPDISPATCH Sh, LPDISPATCH Target, BOOL* Cancel);
	void SheetActivate(LPDISPATCH Sh);
	void SheetDeactivate(LPDISPATCH Sh);
	void SheetCalculate(LPDISPATCH Sh);
	void SheetChange(LPDISPATCH Sh, LPDISPATCH Target);
	void WorkbookOpen(LPDISPATCH Wb);
	void WorkbookActivate(LPDISPATCH Wb);
	void WorkbookDeactivate(LPDISPATCH Wb);
	void WorkbookBeforeClose(LPDISPATCH Wb, BOOL* Cancel);
	void WorkbookBeforeSave(LPDISPATCH Wb, BOOL SaveAsUI, BOOL* Cancel);
	void WorkbookBeforePrint(LPDISPATCH Wb, BOOL* Cancel);
	void WorkbookNewSheet(LPDISPATCH Wb, LPDISPATCH Sh);
	void WorkbookAddinInstall(LPDISPATCH Wb);
	void WorkbookAddinUninstall(LPDISPATCH Wb);
	void WindowResize(LPDISPATCH Wb, LPDISPATCH Wn);
	void WindowActivate(LPDISPATCH Wb, LPDISPATCH Wn);
	void WindowDeactivate(LPDISPATCH Wb, LPDISPATCH Wn);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse WorksheetFunction 

class WorksheetFunction : public COleDispatchDriver
{
public:
	WorksheetFunction() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	WorksheetFunction(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	WorksheetFunction(const WorksheetFunction& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	double Count(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	BOOL IsNA(const VARIANT& Arg1);
	BOOL IsError(const VARIANT& Arg1);
	double Sum(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	double Average(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	double Min(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	double Max(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	double Npv(double Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, const VARIANT& Arg11, 
		const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, const VARIANT& Arg21, 
		const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	double StDev(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	CString Dollar(double Arg1, const VARIANT& Arg2);
	CString Fixed(double Arg1, const VARIANT& Arg2, const VARIANT& Arg3);
	double Pi();
	double Ln(double Arg1);
	double Log10(double Arg1);
	double Round(double Arg1, double Arg2);
	VARIANT Lookup(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3);
	VARIANT Index(const VARIANT& Arg1, double Arg2, const VARIANT& Arg3, const VARIANT& Arg4);
	CString Rept(LPCTSTR Arg1, double Arg2);
	BOOL And(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, const VARIANT& Arg11, 
		const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, const VARIANT& Arg21, 
		const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	BOOL Or(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, const VARIANT& Arg11, 
		const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, const VARIANT& Arg21, 
		const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	double DCount(LPDISPATCH Arg1, const VARIANT& Arg2, const VARIANT& Arg3);
	double DSum(LPDISPATCH Arg1, const VARIANT& Arg2, const VARIANT& Arg3);
	double DAverage(LPDISPATCH Arg1, const VARIANT& Arg2, const VARIANT& Arg3);
	double DMin(LPDISPATCH Arg1, const VARIANT& Arg2, const VARIANT& Arg3);
	double DMax(LPDISPATCH Arg1, const VARIANT& Arg2, const VARIANT& Arg3);
	double DStDev(LPDISPATCH Arg1, const VARIANT& Arg2, const VARIANT& Arg3);
	double Var(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	double DVar(LPDISPATCH Arg1, const VARIANT& Arg2, const VARIANT& Arg3);
	CString Text(const VARIANT& Arg1, LPCTSTR Arg2);
	VARIANT LinEst(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4);
	VARIANT Trend(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4);
	VARIANT LogEst(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4);
	VARIANT Growth(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4);
	double Pv(double Arg1, double Arg2, double Arg3, const VARIANT& Arg4, const VARIANT& Arg5);
	double Fv(double Arg1, double Arg2, double Arg3, const VARIANT& Arg4, const VARIANT& Arg5);
	double NPer(double Arg1, double Arg2, double Arg3, const VARIANT& Arg4, const VARIANT& Arg5);
	double Pmt(double Arg1, double Arg2, double Arg3, const VARIANT& Arg4, const VARIANT& Arg5);
	double Rate(double Arg1, double Arg2, double Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6);
	double MIrr(const VARIANT& Arg1, double Arg2, double Arg3);
	double Irr(const VARIANT& Arg1, const VARIANT& Arg2);
	double Match(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3);
	double Weekday(const VARIANT& Arg1, const VARIANT& Arg2);
	double Search(LPCTSTR Arg1, LPCTSTR Arg2, const VARIANT& Arg3);
	VARIANT Transpose(const VARIANT& Arg1);
	double Atan2(double Arg1, double Arg2);
	double Asin(double Arg1);
	double Acos(double Arg1);
	VARIANT Choose(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	VARIANT HLookup(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4);
	VARIANT VLookup(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4);
	double Log(double Arg1, const VARIANT& Arg2);
	CString Proper(LPCTSTR Arg1);
	CString Trim(LPCTSTR Arg1);
	CString Replace(LPCTSTR Arg1, double Arg2, double Arg3, LPCTSTR Arg4);
	CString Substitute(LPCTSTR Arg1, LPCTSTR Arg2, LPCTSTR Arg3, const VARIANT& Arg4);
	double Find(LPCTSTR Arg1, LPCTSTR Arg2, const VARIANT& Arg3);
	BOOL IsErr(const VARIANT& Arg1);
	BOOL IsText(const VARIANT& Arg1);
	BOOL IsNumber(const VARIANT& Arg1);
	double Sln(double Arg1, double Arg2, double Arg3);
	double Syd(double Arg1, double Arg2, double Arg3, double Arg4);
	double Ddb(double Arg1, double Arg2, double Arg3, double Arg4, const VARIANT& Arg5);
	CString Clean(LPCTSTR Arg1);
	double MDeterm(const VARIANT& Arg1);
	VARIANT MInverse(const VARIANT& Arg1);
	VARIANT MMult(const VARIANT& Arg1, const VARIANT& Arg2);
	double Ipmt(double Arg1, double Arg2, double Arg3, double Arg4, const VARIANT& Arg5, const VARIANT& Arg6);
	double Ppmt(double Arg1, double Arg2, double Arg3, double Arg4, const VARIANT& Arg5, const VARIANT& Arg6);
	double CountA(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	double Product(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	double Fact(double Arg1);
	double DProduct(LPDISPATCH Arg1, const VARIANT& Arg2, const VARIANT& Arg3);
	BOOL IsNonText(const VARIANT& Arg1);
	double StDevP(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	double VarP(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	double DStDevP(LPDISPATCH Arg1, const VARIANT& Arg2, const VARIANT& Arg3);
	double DVarP(LPDISPATCH Arg1, const VARIANT& Arg2, const VARIANT& Arg3);
	BOOL IsLogical(const VARIANT& Arg1);
	double DCountA(LPDISPATCH Arg1, const VARIANT& Arg2, const VARIANT& Arg3);
	CString USDollar(double Arg1, double Arg2);
	double FindB(LPCTSTR Arg1, LPCTSTR Arg2, const VARIANT& Arg3);
	double SearchB(LPCTSTR Arg1, LPCTSTR Arg2, const VARIANT& Arg3);
	CString ReplaceB(LPCTSTR Arg1, double Arg2, double Arg3, LPCTSTR Arg4);
	double RoundUp(double Arg1, double Arg2);
	double RoundDown(double Arg1, double Arg2);
	double Rank(double Arg1, LPDISPATCH Arg2, const VARIANT& Arg3);
	double Days360(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3);
	double Vdb(double Arg1, double Arg2, double Arg3, double Arg4, double Arg5, const VARIANT& Arg6, const VARIANT& Arg7);
	double Median(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	double SumProduct(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	double Sinh(double Arg1);
	double Cosh(double Arg1);
	double Tanh(double Arg1);
	double Asinh(double Arg1);
	double Acosh(double Arg1);
	double Atanh(double Arg1);
	VARIANT DGet(LPDISPATCH Arg1, const VARIANT& Arg2, const VARIANT& Arg3);
	double Db(double Arg1, double Arg2, double Arg3, double Arg4, const VARIANT& Arg5);
	VARIANT Frequency(const VARIANT& Arg1, const VARIANT& Arg2);
	double AveDev(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	double BetaDist(double Arg1, double Arg2, double Arg3, const VARIANT& Arg4, const VARIANT& Arg5);
	double GammaLn(double Arg1);
	double BetaInv(double Arg1, double Arg2, double Arg3, const VARIANT& Arg4, const VARIANT& Arg5);
	double BinomDist(double Arg1, double Arg2, double Arg3, BOOL Arg4);
	double ChiDist(double Arg1, double Arg2);
	double ChiInv(double Arg1, double Arg2);
	double Combin(double Arg1, double Arg2);
	double Confidence(double Arg1, double Arg2, double Arg3);
	double CritBinom(double Arg1, double Arg2, double Arg3);
	double Even(double Arg1);
	double ExponDist(double Arg1, double Arg2, BOOL Arg3);
	double FDist(double Arg1, double Arg2, double Arg3);
	double FInv(double Arg1, double Arg2, double Arg3);
	double Fisher(double Arg1);
	double FisherInv(double Arg1);
	double Floor(double Arg1, double Arg2);
	double GammaDist(double Arg1, double Arg2, double Arg3, BOOL Arg4);
	double GammaInv(double Arg1, double Arg2, double Arg3);
	double Ceiling(double Arg1, double Arg2);
	double HypGeomDist(double Arg1, double Arg2, double Arg3, double Arg4);
	double LogNormDist(double Arg1, double Arg2, double Arg3);
	double LogInv(double Arg1, double Arg2, double Arg3);
	double NegBinomDist(double Arg1, double Arg2, double Arg3);
	double NormDist(double Arg1, double Arg2, double Arg3, BOOL Arg4);
	double NormSDist(double Arg1);
	double NormInv(double Arg1, double Arg2, double Arg3);
	double NormSInv(double Arg1);
	double Standardize(double Arg1, double Arg2, double Arg3);
	double Odd(double Arg1);
	double Permut(double Arg1, double Arg2);
	double Poisson(double Arg1, double Arg2, BOOL Arg3);
	double TDist(double Arg1, double Arg2, double Arg3);
	double Weibull(double Arg1, double Arg2, double Arg3, BOOL Arg4);
	double SumXMY2(const VARIANT& Arg1, const VARIANT& Arg2);
	double SumX2MY2(const VARIANT& Arg1, const VARIANT& Arg2);
	double SumX2PY2(const VARIANT& Arg1, const VARIANT& Arg2);
	double ChiTest(const VARIANT& Arg1, const VARIANT& Arg2);
	double Correl(const VARIANT& Arg1, const VARIANT& Arg2);
	double Covar(const VARIANT& Arg1, const VARIANT& Arg2);
	double Forecast(double Arg1, const VARIANT& Arg2, const VARIANT& Arg3);
	double FTest(const VARIANT& Arg1, const VARIANT& Arg2);
	double Intercept(const VARIANT& Arg1, const VARIANT& Arg2);
	double Pearson(const VARIANT& Arg1, const VARIANT& Arg2);
	double RSq(const VARIANT& Arg1, const VARIANT& Arg2);
	double StEyx(const VARIANT& Arg1, const VARIANT& Arg2);
	double Slope(const VARIANT& Arg1, const VARIANT& Arg2);
	double TTest(const VARIANT& Arg1, const VARIANT& Arg2, double Arg3, double Arg4);
	double Prob(const VARIANT& Arg1, const VARIANT& Arg2, double Arg3, const VARIANT& Arg4);
	double DevSq(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	double GeoMean(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	double HarMean(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	double SumSq(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	double Kurt(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	double Skew(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	double ZTest(const VARIANT& Arg1, double Arg2, const VARIANT& Arg3);
	double Large(const VARIANT& Arg1, double Arg2);
	double Small(const VARIANT& Arg1, double Arg2);
	double Quartile(const VARIANT& Arg1, double Arg2);
	double Percentile(const VARIANT& Arg1, double Arg2);
	double PercentRank(const VARIANT& Arg1, double Arg2, const VARIANT& Arg3);
	double Mode(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	double TrimMean(const VARIANT& Arg1, double Arg2);
	double TInv(double Arg1, double Arg2);
	double Power(double Arg1, double Arg2);
	double Radians(double Arg1);
	double Degrees(double Arg1);
	double Subtotal(double Arg1, LPDISPATCH Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, const VARIANT& Arg11, 
		const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, const VARIANT& Arg21, 
		const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	double SumIf(LPDISPATCH Arg1, const VARIANT& Arg2, const VARIANT& Arg3);
	double CountIf(LPDISPATCH Arg1, const VARIANT& Arg2);
	double CountBlank(LPDISPATCH Arg1);
	double Ispmt(double Arg1, double Arg2, double Arg3, double Arg4);
	CString Roman(double Arg1, const VARIANT& Arg2);
	CString Asc(LPCTSTR Arg1);
	CString Dbcs(LPCTSTR Arg1);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Range 

class Range : public COleDispatchDriver
{
public:
	Range() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Range(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Range(const Range& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void Activate();
	VARIANT GetAddIndent();
	void SetAddIndent(const VARIANT& newValue);
	CString GetAddress(const VARIANT& RowAbsolute, const VARIANT& ColumnAbsolute, long ReferenceStyle, const VARIANT& External, const VARIANT& RelativeTo);
	CString GetAddressLocal(const VARIANT& RowAbsolute, const VARIANT& ColumnAbsolute, long ReferenceStyle, const VARIANT& External, const VARIANT& RelativeTo);
	void AdvancedFilter(long Action, const VARIANT& CriteriaRange, const VARIANT& CopyToRange, const VARIANT& Unique);
	void ApplyNames(const VARIANT& Names, const VARIANT& IgnoreRelativeAbsolute, const VARIANT& UseRowColumnNames, const VARIANT& OmitColumn, const VARIANT& OmitRow, long Order, const VARIANT& AppendLast);
	void ApplyOutlineStyles();
	LPDISPATCH GetAreas();
	CString AutoComplete(LPCTSTR String);
	void AutoFill(LPDISPATCH Destination, long Type);
	void AutoFilter(const VARIANT& Field, const VARIANT& Criteria1, long Operator, const VARIANT& Criteria2, const VARIANT& VisibleDropDown);
	void AutoFit();
	void AutoFormat(long Format, const VARIANT& Number, const VARIANT& Font, const VARIANT& Alignment, const VARIANT& Border, const VARIANT& Pattern, const VARIANT& Width);
	void AutoOutline();
	void BorderAround(const VARIANT& LineStyle, long Weight, long ColorIndex, const VARIANT& Color);
	LPDISPATCH GetBorders();
	void Calculate();
	LPDISPATCH GetCells();
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	void Clear();
	void ClearContents();
	void ClearFormats();
	void ClearNotes();
	void ClearOutline();
	long GetColumn();
	LPDISPATCH ColumnDifferences(const VARIANT& Comparison);
	LPDISPATCH GetColumns();
	VARIANT GetColumnWidth();
	void SetColumnWidth(const VARIANT& newValue);
	void Consolidate(const VARIANT& Sources, const VARIANT& Function, const VARIANT& TopRow, const VARIANT& LeftColumn, const VARIANT& CreateLinks);
	void Copy(const VARIANT& Destination);
	long CopyFromRecordset(LPUNKNOWN Data, const VARIANT& MaxRows, const VARIANT& MaxColumns);
	void CopyPicture(long Appearance, long Format);
	long GetCount();
	void CreateNames(const VARIANT& Top, const VARIANT& Left, const VARIANT& Bottom, const VARIANT& Right);
	void CreatePublisher(const VARIANT& Edition, long Appearance, const VARIANT& ContainsPICT, const VARIANT& ContainsBIFF, const VARIANT& ContainsRTF, const VARIANT& ContainsVALU);
	LPDISPATCH GetCurrentArray();
	LPDISPATCH GetCurrentRegion();
	void Cut(const VARIANT& Destination);
	void DataSeries(const VARIANT& Rowcol, long Type, long Date, const VARIANT& Step, const VARIANT& Stop, const VARIANT& Trend);
	VARIANT Get_Default(const VARIANT& RowIndex, const VARIANT& ColumnIndex);
	void Set_Default(const VARIANT& RowIndex, const VARIANT& ColumnIndex, const VARIANT& newValue);
	void Delete(const VARIANT& Shift);
	LPDISPATCH GetDependents();
	VARIANT DialogBox_();
	LPDISPATCH GetDirectDependents();
	LPDISPATCH GetDirectPrecedents();
	VARIANT EditionOptions(long Type, long Option, const VARIANT& Name, const VARIANT& Reference, long Appearance, long ChartSize, const VARIANT& Format);
	LPDISPATCH GetEnd(long Direction);
	LPDISPATCH GetEntireColumn();
	LPDISPATCH GetEntireRow();
	void FillDown();
	void FillLeft();
	void FillRight();
	void FillUp();
	LPDISPATCH Find(const VARIANT& What, const VARIANT& After, const VARIANT& LookIn, const VARIANT& LookAt, const VARIANT& SearchOrder, long SearchDirection, const VARIANT& MatchCase, const VARIANT& MatchByte, 
		const VARIANT& MatchControlCharacters, const VARIANT& MatchDiacritics, const VARIANT& MatchKashida, const VARIANT& MatchAlefHamza);
	LPDISPATCH FindNext(const VARIANT& After);
	LPDISPATCH FindPrevious(const VARIANT& After);
	LPDISPATCH GetFont();
	VARIANT GetFormula();
	void SetFormula(const VARIANT& newValue);
	VARIANT GetFormulaArray();
	void SetFormulaArray(const VARIANT& newValue);
	long GetFormulaLabel();
	void SetFormulaLabel(long nNewValue);
	VARIANT GetFormulaHidden();
	void SetFormulaHidden(const VARIANT& newValue);
	VARIANT GetFormulaLocal();
	void SetFormulaLocal(const VARIANT& newValue);
	VARIANT GetFormulaR1C1();
	void SetFormulaR1C1(const VARIANT& newValue);
	VARIANT GetFormulaR1C1Local();
	void SetFormulaR1C1Local(const VARIANT& newValue);
	void FunctionWizard();
	BOOL GoalSeek(const VARIANT& Goal, LPDISPATCH ChangingCell);
	VARIANT Group(const VARIANT& Start, const VARIANT& End, const VARIANT& By, const VARIANT& Periods);
	VARIANT GetHasArray();
	VARIANT GetHasFormula();
	VARIANT GetHeight();
	VARIANT GetHidden();
	void SetHidden(const VARIANT& newValue);
	VARIANT GetHorizontalAlignment();
	void SetHorizontalAlignment(const VARIANT& newValue);
	VARIANT GetIndentLevel();
	void SetIndentLevel(const VARIANT& newValue);
	void InsertIndent(long InsertAmount);
	void Insert(const VARIANT& Shift);
	LPDISPATCH GetInterior();
	VARIANT GetItem(const VARIANT& RowIndex, const VARIANT& ColumnIndex);
	void SetItem(const VARIANT& RowIndex, const VARIANT& ColumnIndex, const VARIANT& newValue);
	void Justify();
	VARIANT GetLeft();
	long GetListHeaderRows();
	void ListNames();
	long GetLocationInTable();
	VARIANT GetLocked();
	void SetLocked(const VARIANT& newValue);
	void Merge(const VARIANT& Across);
	void UnMerge();
	LPDISPATCH GetMergeArea();
	VARIANT GetMergeCells();
	void SetMergeCells(const VARIANT& newValue);
	VARIANT GetName();
	void SetName(const VARIANT& newValue);
	void NavigateArrow(const VARIANT& TowardPrecedent, const VARIANT& ArrowNumber, const VARIANT& LinkNumber);
	LPUNKNOWN Get_NewEnum();
	LPDISPATCH GetNext();
	CString NoteText(const VARIANT& Text, const VARIANT& Start, const VARIANT& Length);
	VARIANT GetNumberFormat();
	void SetNumberFormat(const VARIANT& newValue);
	VARIANT GetNumberFormatLocal();
	void SetNumberFormatLocal(const VARIANT& newValue);
	LPDISPATCH GetOffset(const VARIANT& RowOffset, const VARIANT& ColumnOffset);
	VARIANT GetOrientation();
	void SetOrientation(const VARIANT& newValue);
	VARIANT GetOutlineLevel();
	void SetOutlineLevel(const VARIANT& newValue);
	long GetPageBreak();
	void SetPageBreak(long nNewValue);
	void Parse(const VARIANT& ParseLine, const VARIANT& Destination);
	void PasteSpecial(long Paste, long Operation, const VARIANT& SkipBlanks, const VARIANT& Transpose);
	void PasteSpecial();
	LPDISPATCH GetPivotField();
	LPDISPATCH GetPivotItem();
	LPDISPATCH GetPivotTable();
	LPDISPATCH GetPrecedents();
	VARIANT GetPrefixCharacter();
	LPDISPATCH GetPrevious();
	void PrintOut(const VARIANT& From, const VARIANT& To, const VARIANT& Copies, const VARIANT& Preview, const VARIANT& ActivePrinter, const VARIANT& PrintToFile, const VARIANT& Collate);
	void PrintPreview(const VARIANT& EnableChanges);
	LPDISPATCH GetQueryTable();
	LPDISPATCH GetRange(const VARIANT& Cell1, const VARIANT& Cell2);
	void RemoveSubtotal();
	BOOL Replace(const VARIANT& What, const VARIANT& Replacement, const VARIANT& LookAt, const VARIANT& SearchOrder, const VARIANT& MatchCase, const VARIANT& MatchByte, const VARIANT& MatchControlCharacters, const VARIANT& MatchDiacritics, 
		const VARIANT& MatchKashida, const VARIANT& MatchAlefHamza);
	LPDISPATCH GetResize(const VARIANT& RowSize, const VARIANT& ColumnSize);
	long GetRow();
	LPDISPATCH RowDifferences(const VARIANT& Comparison);
	VARIANT GetRowHeight();
	void SetRowHeight(const VARIANT& newValue);
	LPDISPATCH GetRows();
	VARIANT Run(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
	void Select();
	void Show();
	void ShowDependents(const VARIANT& Remove);
	VARIANT GetShowDetail();
	void SetShowDetail(const VARIANT& newValue);
	void ShowErrors();
	void ShowPrecedents(const VARIANT& Remove);
	VARIANT GetShrinkToFit();
	void SetShrinkToFit(const VARIANT& newValue);
	void Sort(const VARIANT& Key1, long Order1, const VARIANT& Key2, const VARIANT& Type, long Order2, const VARIANT& Key3, long Order3, long Header, const VARIANT& OrderCustom, const VARIANT& MatchCase, long Orientation, long SortMethod, 
		const VARIANT& IgnoreControlCharacters, const VARIANT& IgnoreDiacritics, const VARIANT& IgnoreKashida);
	void SortSpecial(long SortMethod, const VARIANT& Key1, long Order1, const VARIANT& Type, const VARIANT& Key2, long Order2, const VARIANT& Key3, long Order3, long Header, const VARIANT& OrderCustom, const VARIANT& MatchCase, long Orientation);
	LPDISPATCH GetSoundNote();
	LPDISPATCH SpecialCells(long Type, const VARIANT& Value);
	VARIANT GetStyle();
	void SetStyle(const VARIANT& newValue);
	void SubscribeTo(LPCTSTR Edition, long Format);
	void Subtotal(long GroupBy, long Function, const VARIANT& TotalList, const VARIANT& Replace, const VARIANT& PageBreaks, long SummaryBelowData);
	VARIANT GetSummary();
	void Table(const VARIANT& RowInput, const VARIANT& ColumnInput);
	VARIANT GetText();
	void TextToColumns(const VARIANT& Destination, long DataType, long TextQualifier, const VARIANT& ConsecutiveDelimiter, const VARIANT& Tab, const VARIANT& Semicolon, const VARIANT& Comma, const VARIANT& Space, const VARIANT& Other, 
		const VARIANT& OtherChar, const VARIANT& FieldInfo);
	VARIANT GetTop();
	void Ungroup();
	VARIANT GetUseStandardHeight();
	void SetUseStandardHeight(const VARIANT& newValue);
	VARIANT GetUseStandardWidth();
	void SetUseStandardWidth(const VARIANT& newValue);
	LPDISPATCH GetValidation();
	VARIANT GetValue();
	void SetValue(const VARIANT& newValue);
	VARIANT GetValue2();
	void SetValue2(const VARIANT& newValue);
	VARIANT GetVerticalAlignment();
	void SetVerticalAlignment(const VARIANT& newValue);
	VARIANT GetWidth();
	LPDISPATCH GetWorksheet();
	VARIANT GetWrapText();
	void SetWrapText(const VARIANT& newValue);
	LPDISPATCH AddComment(const VARIANT& Text);
	LPDISPATCH GetComment();
	void ClearComments();
	LPDISPATCH GetPhonetic();
	LPDISPATCH GetFormatConditions();
	long GetReadingOrder();
	void SetReadingOrder(long nNewValue);
	LPDISPATCH GetHyperlinks();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ChartEvents 

class ChartEvents : public COleDispatchDriver
{
public:
	ChartEvents() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ChartEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ChartEvents(const ChartEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	void Activate();
	void Deactivate();
	void Resize();
	void MouseDown(long Button, long Shift, long X, long Y);
	void MouseUp(long Button, long Shift, long X, long Y);
	void MouseMove(long Button, long Shift, long X, long Y);
	void BeforeRightClick(BOOL* Cancel);
	void DragPlot();
	void DragOver();
	void BeforeDoubleClick(long ElementID, long Arg1, long Arg2, BOOL* Cancel);
	void Select(long ElementID, long Arg1, long Arg2);
	void SeriesChange(long SeriesIndex, long PointIndex);
	void Calculate();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse VPageBreak 

class VPageBreak : public COleDispatchDriver
{
public:
	VPageBreak() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	VPageBreak(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	VPageBreak(const VPageBreak& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void Delete();
	void DragOff(long Direction, long RegionIndex);
	long GetType();
	void SetType(long nNewValue);
	long GetExtent();
	LPDISPATCH GetLocation();
	void SetRefLocation(LPDISPATCH newValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse HPageBreak 

class HPageBreak : public COleDispatchDriver
{
public:
	HPageBreak() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	HPageBreak(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	HPageBreak(const HPageBreak& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void Delete();
	void DragOff(long Direction, long RegionIndex);
	long GetType();
	void SetType(long nNewValue);
	long GetExtent();
	LPDISPATCH GetLocation();
	void SetRefLocation(LPDISPATCH newValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse HPageBreaks 

class HPageBreaks : public COleDispatchDriver
{
public:
	HPageBreaks() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	HPageBreaks(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	HPageBreaks(const HPageBreaks& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH GetItem(long Index);
	LPDISPATCH Get_Default(long Index);
	LPUNKNOWN Get_NewEnum();
	LPDISPATCH Add(LPDISPATCH Before);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse VPageBreaks 

class VPageBreaks : public COleDispatchDriver
{
public:
	VPageBreaks() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	VPageBreaks(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	VPageBreaks(const VPageBreaks& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH GetItem(long Index);
	LPDISPATCH Get_Default(long Index);
	LPUNKNOWN Get_NewEnum();
	LPDISPATCH Add(LPDISPATCH Before);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse RecentFile 

class RecentFile : public COleDispatchDriver
{
public:
	RecentFile() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	RecentFile(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	RecentFile(const RecentFile& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetName();
	CString GetPath();
	long GetIndex();
	LPDISPATCH Open();
	void Delete();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse RecentFiles 

class RecentFiles : public COleDispatchDriver
{
public:
	RecentFiles() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	RecentFiles(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	RecentFiles(const RecentFiles& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetMaximum();
	void SetMaximum(long nNewValue);
	long GetCount();
	LPDISPATCH GetItem(long Index);
	LPDISPATCH Get_Default(long Index);
	LPDISPATCH Add(LPCTSTR Name);
	LPUNKNOWN Get_NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse DocEvents 

class DocEvents : public COleDispatchDriver
{
public:
	DocEvents() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	DocEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	DocEvents(const DocEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	void SelectionChange(LPDISPATCH Target);
	void BeforeDoubleClick(LPDISPATCH Target, BOOL* Cancel);
	void BeforeRightClick(LPDISPATCH Target, BOOL* Cancel);
	void Activate();
	void Deactivate();
	void Calculate();
	void Change(LPDISPATCH Target);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Style 

class Style : public COleDispatchDriver
{
public:
	Style() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Style(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Style(const Style& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	BOOL GetAddIndent();
	void SetAddIndent(BOOL bNewValue);
	BOOL GetBuiltIn();
	LPDISPATCH GetBorders();
	void Delete();
	LPDISPATCH GetFont();
	BOOL GetFormulaHidden();
	void SetFormulaHidden(BOOL bNewValue);
	long GetHorizontalAlignment();
	void SetHorizontalAlignment(long nNewValue);
	BOOL GetIncludeAlignment();
	void SetIncludeAlignment(BOOL bNewValue);
	BOOL GetIncludeBorder();
	void SetIncludeBorder(BOOL bNewValue);
	BOOL GetIncludeFont();
	void SetIncludeFont(BOOL bNewValue);
	BOOL GetIncludeNumber();
	void SetIncludeNumber(BOOL bNewValue);
	BOOL GetIncludePatterns();
	void SetIncludePatterns(BOOL bNewValue);
	BOOL GetIncludeProtection();
	void SetIncludeProtection(BOOL bNewValue);
	long GetIndentLevel();
	void SetIndentLevel(long nNewValue);
	LPDISPATCH GetInterior();
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	VARIANT GetMergeCells();
	void SetMergeCells(const VARIANT& newValue);
	CString GetName();
	CString GetNameLocal();
	CString GetNumberFormat();
	void SetNumberFormat(LPCTSTR lpszNewValue);
	CString GetNumberFormatLocal();
	void SetNumberFormatLocal(LPCTSTR lpszNewValue);
	long GetOrientation();
	void SetOrientation(long nNewValue);
	BOOL GetShrinkToFit();
	void SetShrinkToFit(BOOL bNewValue);
	CString GetValue();
	long GetVerticalAlignment();
	void SetVerticalAlignment(long nNewValue);
	BOOL GetWrapText();
	void SetWrapText(BOOL bNewValue);
	CString Get_Default();
	long GetReadingOrder();
	void SetReadingOrder(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Styles 

class Styles : public COleDispatchDriver
{
public:
	Styles() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Styles(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Styles(const Styles& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH Add(LPCTSTR Name, const VARIANT& BasedOn);
	long GetCount();
	LPDISPATCH GetItem(const VARIANT& Index);
	void Merge(const VARIANT& Workbook);
	LPUNKNOWN Get_NewEnum();
	LPDISPATCH Get_Default(const VARIANT& Index);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Borders 

class Borders : public COleDispatchDriver
{
public:
	Borders() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Borders(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Borders(const Borders& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	VARIANT GetColor();
	void SetColor(const VARIANT& newValue);
	VARIANT GetColorIndex();
	void SetColorIndex(const VARIANT& newValue);
	long GetCount();
	LPDISPATCH GetItem(long Index);
	VARIANT GetLineStyle();
	void SetLineStyle(const VARIANT& newValue);
	LPUNKNOWN Get_NewEnum();
	VARIANT GetValue();
	void SetValue(const VARIANT& newValue);
	VARIANT GetWeight();
	void SetWeight(const VARIANT& newValue);
	LPDISPATCH Get_Default(long Index);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse AddIn 

class AddIn : public COleDispatchDriver
{
public:
	AddIn() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	AddIn(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	AddIn(const AddIn& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetFullName();
	BOOL GetInstalled();
	void SetInstalled(BOOL bNewValue);
	CString GetName();
	CString GetPath();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse AddIns 

class AddIns : public COleDispatchDriver
{
public:
	AddIns() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	AddIns(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	AddIns(const AddIns& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH Add(LPCTSTR Filename, const VARIANT& CopyFile);
	long GetCount();
	LPDISPATCH GetItem(const VARIANT& Index);
	LPUNKNOWN Get_NewEnum();
	LPDISPATCH Get_Default(const VARIANT& Index);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Toolbar 

class Toolbar : public COleDispatchDriver
{
public:
	Toolbar() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Toolbar(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Toolbar(const Toolbar& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	BOOL GetBuiltIn();
	void Delete();
	long GetHeight();
	void SetHeight(long nNewValue);
	long GetLeft();
	void SetLeft(long nNewValue);
	CString GetName();
	long GetPosition();
	void SetPosition(long nNewValue);
	long GetProtection();
	void SetProtection(long nNewValue);
	void Reset();
	LPDISPATCH GetToolbarButtons();
	long GetTop();
	void SetTop(long nNewValue);
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	long GetWidth();
	void SetWidth(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Toolbars 

class Toolbars : public COleDispatchDriver
{
public:
	Toolbars() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Toolbars(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Toolbars(const Toolbars& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH Add(const VARIANT& Name);
	long GetCount();
	LPDISPATCH Get_Default(const VARIANT& Index);
	LPDISPATCH GetItem(const VARIANT& Index);
	LPUNKNOWN Get_NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ToolbarButton 

class ToolbarButton : public COleDispatchDriver
{
public:
	ToolbarButton() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ToolbarButton(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ToolbarButton(const ToolbarButton& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	BOOL GetBuiltIn();
	BOOL GetBuiltInFace();
	void SetBuiltInFace(BOOL bNewValue);
	void Copy(LPDISPATCH Toolbar, long Before);
	void CopyFace();
	void Delete();
	void Edit();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	long GetHelpContextID();
	void SetHelpContextID(long nNewValue);
	CString GetHelpFile();
	void SetHelpFile(LPCTSTR lpszNewValue);
	long GetId();
	BOOL GetIsGap();
	void Move(LPDISPATCH Toolbar, long Before);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	CString GetOnAction();
	void SetOnAction(LPCTSTR lpszNewValue);
	void PasteFace();
	BOOL GetPushed();
	void SetPushed(BOOL bNewValue);
	void Reset();
	CString GetStatusBar();
	void SetStatusBar(LPCTSTR lpszNewValue);
	long GetWidth();
	void SetWidth(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ToolbarButtons 

class ToolbarButtons : public COleDispatchDriver
{
public:
	ToolbarButtons() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ToolbarButtons(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ToolbarButtons(const ToolbarButtons& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH Add(const VARIANT& Button, const VARIANT& Before, const VARIANT& OnAction, const VARIANT& Pushed, const VARIANT& Enabled, const VARIANT& StatusBar, const VARIANT& HelpFile, const VARIANT& HelpContextID);
	long GetCount();
	LPDISPATCH GetItem(long Index);
	LPUNKNOWN Get_NewEnum();
	LPDISPATCH Get_Default(long Index);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Areas 

class Areas : public COleDispatchDriver
{
public:
	Areas() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Areas(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Areas(const Areas& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH GetItem(long Index);
	LPUNKNOWN Get_NewEnum();
	LPDISPATCH Get_Default(long Index);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse WorkbookEvents 

class WorkbookEvents : public COleDispatchDriver
{
public:
	WorkbookEvents() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	WorkbookEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	WorkbookEvents(const WorkbookEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	void Open();
	void Activate();
	void Deactivate();
	void BeforeClose(BOOL* Cancel);
	void BeforeSave(BOOL SaveAsUI, BOOL* Cancel);
	void BeforePrint(BOOL* Cancel);
	void NewSheet(LPDISPATCH Sh);
	void AddinInstall();
	void AddinUninstall();
	void WindowResize(LPDISPATCH Wn);
	void WindowActivate(LPDISPATCH Wn);
	void WindowDeactivate(LPDISPATCH Wn);
	void SheetSelectionChange(LPDISPATCH Sh, LPDISPATCH Target);
	void SheetBeforeDoubleClick(LPDISPATCH Sh, LPDISPATCH Target, BOOL* Cancel);
	void SheetBeforeRightClick(LPDISPATCH Sh, LPDISPATCH Target, BOOL* Cancel);
	void SheetActivate(LPDISPATCH Sh);
	void SheetDeactivate(LPDISPATCH Sh);
	void SheetCalculate(LPDISPATCH Sh);
	void SheetChange(LPDISPATCH Sh, LPDISPATCH Target);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse MenuBars 

class MenuBars : public COleDispatchDriver
{
public:
	MenuBars() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	MenuBars(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	MenuBars(const MenuBars& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH Add(const VARIANT& Name);
	long GetCount();
	LPDISPATCH Get_Default(const VARIANT& Index);
	LPDISPATCH GetItem(const VARIANT& Index);
	LPUNKNOWN Get_NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse MenuBar 

class MenuBar : public COleDispatchDriver
{
public:
	MenuBar() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	MenuBar(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	MenuBar(const MenuBar& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void Activate();
	BOOL GetBuiltIn();
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	void Delete();
	long GetIndex();
	LPDISPATCH GetMenus();
	void Reset();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Menus 

class Menus : public COleDispatchDriver
{
public:
	Menus() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Menus(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Menus(const Menus& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH Add(LPCTSTR Caption, const VARIANT& Before, const VARIANT& Restore);
	long GetCount();
	LPDISPATCH Get_Default(const VARIANT& Index);
	LPDISPATCH GetItem(const VARIANT& Index);
	LPUNKNOWN Get_NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Menu 

class Menu : public COleDispatchDriver
{
public:
	Menu() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Menu(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Menu(const Menu& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	void Delete();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	long GetIndex();
	LPDISPATCH GetMenuItems();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse MenuItems 

class MenuItems : public COleDispatchDriver
{
public:
	MenuItems() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	MenuItems(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	MenuItems(const MenuItems& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH Add(LPCTSTR Caption, const VARIANT& OnAction, const VARIANT& ShortcutKey, const VARIANT& Before, const VARIANT& Restore, const VARIANT& StatusBar, const VARIANT& HelpFile, const VARIANT& HelpContextID);
	LPDISPATCH AddMenu(LPCTSTR Caption, const VARIANT& Before, const VARIANT& Restore);
	long GetCount();
	LPDISPATCH Get_Default(const VARIANT& Index);
	LPDISPATCH GetItem(const VARIANT& Index);
	LPUNKNOWN Get_NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse MenuItem 

class MenuItem : public COleDispatchDriver
{
public:
	MenuItem() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	MenuItem(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	MenuItem(const MenuItem& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	BOOL GetChecked();
	void SetChecked(BOOL bNewValue);
	void Delete();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	long GetHelpContextID();
	void SetHelpContextID(long nNewValue);
	CString GetHelpFile();
	void SetHelpFile(LPCTSTR lpszNewValue);
	long GetIndex();
	CString GetOnAction();
	void SetOnAction(LPCTSTR lpszNewValue);
	CString GetStatusBar();
	void SetStatusBar(LPCTSTR lpszNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Charts 

class Charts : public COleDispatchDriver
{
public:
	Charts() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Charts(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Charts(const Charts& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH Add(const VARIANT& Before, const VARIANT& After, const VARIANT& Count);
	void Copy(const VARIANT& Before, const VARIANT& After);
	long GetCount();
	void Delete();
	LPDISPATCH GetItem(const VARIANT& Index);
	void Move(const VARIANT& Before, const VARIANT& After);
	LPUNKNOWN Get_NewEnum();
	void PrintOut(const VARIANT& From, const VARIANT& To, const VARIANT& Copies, const VARIANT& Preview, const VARIANT& ActivePrinter, const VARIANT& PrintToFile, const VARIANT& Collate);
	void PrintPreview(const VARIANT& EnableChanges);
	void Select(const VARIANT& Replace);
	LPDISPATCH GetHPageBreaks();
	LPDISPATCH GetVPageBreaks();
	VARIANT GetVisible();
	void SetVisible(const VARIANT& newValue);
	LPDISPATCH Get_Default(const VARIANT& Index);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse DrawingObjects 

class DrawingObjects : public COleDispatchDriver
{
public:
	DrawingObjects() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	DrawingObjects(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	DrawingObjects(const DrawingObjects& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	VARIANT GetAccelerator();
	void SetAccelerator(const VARIANT& newValue);
	BOOL GetAddIndent();
	void SetAddIndent(BOOL bNewValue);
	void AddItem(const VARIANT& Text, const VARIANT& Index);
	VARIANT GetArrowHeadLength();
	void SetArrowHeadLength(const VARIANT& newValue);
	VARIANT GetArrowHeadStyle();
	void SetArrowHeadStyle(const VARIANT& newValue);
	VARIANT GetArrowHeadWidth();
	void SetArrowHeadWidth(const VARIANT& newValue);
	BOOL GetAutoSize();
	void SetAutoSize(BOOL bNewValue);
	LPDISPATCH GetBorder();
	BOOL GetCancelButton();
	void SetCancelButton(BOOL bNewValue);
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	long Get_Default();
	void Set_Default(long nNewValue);
	BOOL GetDefaultButton();
	void SetDefaultButton(BOOL bNewValue);
	BOOL GetDismissButton();
	void SetDismissButton(BOOL bNewValue);
	BOOL GetDisplay3DShading();
	void SetDisplay3DShading(BOOL bNewValue);
	BOOL GetDisplayVerticalScrollBar();
	void SetDisplayVerticalScrollBar(BOOL bNewValue);
	long GetDropDownLines();
	void SetDropDownLines(long nNewValue);
	LPDISPATCH GetFont();
	BOOL GetHelpButton();
	void SetHelpButton(BOOL bNewValue);
	VARIANT GetHorizontalAlignment();
	void SetHorizontalAlignment(const VARIANT& newValue);
	long GetInputType();
	void SetInputType(long nNewValue);
	LPDISPATCH GetInterior();
	long GetLargeChange();
	void SetLargeChange(long nNewValue);
	CString GetLinkedCell();
	void SetLinkedCell(LPCTSTR lpszNewValue);
	VARIANT List(const VARIANT& Index);
	CString GetListFillRange();
	void SetListFillRange(LPCTSTR lpszNewValue);
	long GetListIndex();
	void SetListIndex(long nNewValue);
	BOOL GetLockedText();
	void SetLockedText(BOOL bNewValue);
	long GetMax();
	void SetMax(long nNewValue);
	long GetMin();
	void SetMin(long nNewValue);
	BOOL GetMultiLine();
	void SetMultiLine(BOOL bNewValue);
	BOOL GetMultiSelect();
	void SetMultiSelect(BOOL bNewValue);
	VARIANT GetOrientation();
	void SetOrientation(const VARIANT& newValue);
	VARIANT GetPhoneticAccelerator();
	void SetPhoneticAccelerator(const VARIANT& newValue);
	void RemoveAllItems();
	void RemoveItem(long Index, const VARIANT& Count);
	void Reshape(long Vertex, const VARIANT& Insert, const VARIANT& Left, const VARIANT& Top);
	BOOL GetRoundedCorners();
	void SetRoundedCorners(BOOL bNewValue);
	VARIANT Selected(const VARIANT& Index);
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
	long GetSmallChange();
	void SetSmallChange(long nNewValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	LPDISPATCH Ungroup();
	long GetValue();
	void SetValue(long nNewValue);
	VARIANT GetVerticalAlignment();
	void SetVerticalAlignment(const VARIANT& newValue);
	VARIANT Vertices(const VARIANT& Index1, const VARIANT& Index2);
	long GetReadingOrder();
	void SetReadingOrder(long nNewValue);
	long GetCount();
	LPDISPATCH Item(const VARIANT& Index);
	LPDISPATCH Group();
	void LinkCombo(const VARIANT& Link);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse PivotCache 

class PivotCache : public COleDispatchDriver
{
public:
	PivotCache() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	PivotCache(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	PivotCache(const PivotCache& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	BOOL GetBackgroundQuery();
	void SetBackgroundQuery(BOOL bNewValue);
	VARIANT GetConnection();
	void SetConnection(const VARIANT& newValue);
	BOOL GetEnableRefresh();
	void SetEnableRefresh(BOOL bNewValue);
	long GetIndex();
	long GetMemoryUsed();
	BOOL GetOptimizeCache();
	void SetOptimizeCache(BOOL bNewValue);
	long GetRecordCount();
	void Refresh();
	DATE GetRefreshDate();
	CString GetRefreshName();
	BOOL GetRefreshOnFileOpen();
	void SetRefreshOnFileOpen(BOOL bNewValue);
	VARIANT GetSql();
	void SetSql(const VARIANT& newValue);
	BOOL GetSavePassword();
	void SetSavePassword(BOOL bNewValue);
	VARIANT GetSourceData();
	void SetSourceData(const VARIANT& newValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse PivotCaches 

class PivotCaches : public COleDispatchDriver
{
public:
	PivotCaches() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	PivotCaches(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	PivotCaches(const PivotCaches& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Item(const VARIANT& Index);
	LPDISPATCH Get_Default(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse PivotFormula 

class PivotFormula : public COleDispatchDriver
{
public:
	PivotFormula() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	PivotFormula(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	PivotFormula(const PivotFormula& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void Delete();
	CString Get_Default();
	void Set_Default(LPCTSTR lpszNewValue);
	CString GetFormula();
	void SetFormula(LPCTSTR lpszNewValue);
	CString GetValue();
	void SetValue(LPCTSTR lpszNewValue);
	long GetIndex();
	void SetIndex(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse PivotFormulas 

class PivotFormulas : public COleDispatchDriver
{
public:
	PivotFormulas() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	PivotFormulas(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	PivotFormulas(const PivotFormulas& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Add(LPCTSTR Formula);
	LPDISPATCH Item(const VARIANT& Index);
	LPDISPATCH Get_Default(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse PivotTable 

class PivotTable : public COleDispatchDriver
{
public:
	PivotTable() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	PivotTable(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	PivotTable(const PivotTable& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void AddFields(const VARIANT& RowFields, const VARIANT& ColumnFields, const VARIANT& PageFields, const VARIANT& AddToTable);
	LPDISPATCH GetColumnFields(const VARIANT& Index);
	BOOL GetColumnGrand();
	void SetColumnGrand(BOOL bNewValue);
	LPDISPATCH GetColumnRange();
	void ShowPages(const VARIANT& PageField);
	LPDISPATCH GetDataBodyRange();
	LPDISPATCH GetDataFields(const VARIANT& Index);
	LPDISPATCH GetDataLabelRange();
	CString Get_Default();
	void Set_Default(LPCTSTR lpszNewValue);
	BOOL GetHasAutoFormat();
	void SetHasAutoFormat(BOOL bNewValue);
	LPDISPATCH GetHiddenFields(const VARIANT& Index);
	CString GetInnerDetail();
	void SetInnerDetail(LPCTSTR lpszNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	LPDISPATCH GetPageFields(const VARIANT& Index);
	LPDISPATCH GetPageRange();
	LPDISPATCH GetPageRangeCells();
	LPDISPATCH PivotFields(const VARIANT& Index);
	DATE GetRefreshDate();
	CString GetRefreshName();
	BOOL RefreshTable();
	LPDISPATCH GetRowFields(const VARIANT& Index);
	BOOL GetRowGrand();
	void SetRowGrand(BOOL bNewValue);
	LPDISPATCH GetRowRange();
	BOOL GetSaveData();
	void SetSaveData(BOOL bNewValue);
	VARIANT GetSourceData();
	void SetSourceData(const VARIANT& newValue);
	LPDISPATCH GetTableRange1();
	LPDISPATCH GetTableRange2();
	CString GetValue();
	void SetValue(LPCTSTR lpszNewValue);
	LPDISPATCH GetVisibleFields(const VARIANT& Index);
	long GetCacheIndex();
	void SetCacheIndex(long nNewValue);
	LPDISPATCH CalculatedFields();
	BOOL GetDisplayErrorString();
	void SetDisplayErrorString(BOOL bNewValue);
	BOOL GetDisplayNullString();
	void SetDisplayNullString(BOOL bNewValue);
	BOOL GetEnableDrilldown();
	void SetEnableDrilldown(BOOL bNewValue);
	BOOL GetEnableFieldDialog();
	void SetEnableFieldDialog(BOOL bNewValue);
	BOOL GetEnableWizard();
	void SetEnableWizard(BOOL bNewValue);
	CString GetErrorString();
	void SetErrorString(LPCTSTR lpszNewValue);
	double GetData(LPCTSTR Name);
	void ListFormulas();
	BOOL GetManualUpdate();
	void SetManualUpdate(BOOL bNewValue);
	BOOL GetMergeLabels();
	void SetMergeLabels(BOOL bNewValue);
	CString GetNullString();
	void SetNullString(LPCTSTR lpszNewValue);
	LPDISPATCH PivotCache();
	LPDISPATCH PivotFormulas();
	void PivotTableWizard(const VARIANT& SourceType, const VARIANT& SourceData, const VARIANT& TableDestination, const VARIANT& TableName, const VARIANT& RowGrand, const VARIANT& ColumnGrand, const VARIANT& SaveData, const VARIANT& HasAutoFormat, 
		const VARIANT& AutoPage, const VARIANT& Reserved, const VARIANT& BackgroundQuery, const VARIANT& OptimizeCache, const VARIANT& PageFieldOrder, const VARIANT& PageFieldWrapCount, const VARIANT& ReadData, const VARIANT& Connection);
	BOOL GetSubtotalHiddenPageItems();
	void SetSubtotalHiddenPageItems(BOOL bNewValue);
	long GetPageFieldOrder();
	void SetPageFieldOrder(long nNewValue);
	CString GetPageFieldStyle();
	void SetPageFieldStyle(LPCTSTR lpszNewValue);
	long GetPageFieldWrapCount();
	void SetPageFieldWrapCount(long nNewValue);
	BOOL GetPreserveFormatting();
	void SetPreserveFormatting(BOOL bNewValue);
	void PivotSelect(LPCTSTR Name, long Mode);
	CString GetPivotSelection();
	void SetPivotSelection(LPCTSTR lpszNewValue);
	long GetSelectionMode();
	void SetSelectionMode(long nNewValue);
	CString GetTableStyle();
	void SetTableStyle(LPCTSTR lpszNewValue);
	CString GetTag();
	void SetTag(LPCTSTR lpszNewValue);
	void Update();
	CString GetVacatedStyle();
	void SetVacatedStyle(LPCTSTR lpszNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse PivotTables 

class PivotTables : public COleDispatchDriver
{
public:
	PivotTables() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	PivotTables(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	PivotTables(const PivotTables& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse PivotField 

class PivotField : public COleDispatchDriver
{
public:
	PivotField() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	PivotField(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	PivotField(const PivotField& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCalculation();
	void SetCalculation(long nNewValue);
	LPDISPATCH GetChildField();
	VARIANT GetChildItems(const VARIANT& Index);
	CString GetCurrentPage();
	void SetCurrentPage(LPCTSTR lpszNewValue);
	LPDISPATCH GetDataRange();
	long GetDataType();
	CString Get_Default();
	void Set_Default(LPCTSTR lpszNewValue);
	long GetFunction();
	void SetFunction(long nNewValue);
	VARIANT GetGroupLevel();
	VARIANT GetHiddenItems(const VARIANT& Index);
	LPDISPATCH GetLabelRange();
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	CString GetNumberFormat();
	void SetNumberFormat(LPCTSTR lpszNewValue);
	long GetOrientation();
	void SetOrientation(long nNewValue);
	BOOL GetShowAllItems();
	void SetShowAllItems(BOOL bNewValue);
	LPDISPATCH GetParentField();
	VARIANT GetParentItems(const VARIANT& Index);
	VARIANT PivotItems(const VARIANT& Index);
	VARIANT GetPosition();
	void SetPosition(const VARIANT& newValue);
	CString GetSourceName();
	VARIANT Subtotals(const VARIANT& Index);
	VARIANT GetBaseField();
	void SetBaseField(const VARIANT& newValue);
	VARIANT GetBaseItem();
	void SetBaseItem(const VARIANT& newValue);
	VARIANT GetTotalLevels();
	CString GetValue();
	void SetValue(LPCTSTR lpszNewValue);
	VARIANT GetVisibleItems(const VARIANT& Index);
	LPDISPATCH CalculatedItems();
	void Delete();
	BOOL GetDragToColumn();
	void SetDragToColumn(BOOL bNewValue);
	BOOL GetDragToHide();
	void SetDragToHide(BOOL bNewValue);
	BOOL GetDragToPage();
	void SetDragToPage(BOOL bNewValue);
	BOOL GetDragToRow();
	void SetDragToRow(BOOL bNewValue);
	CString GetFormula();
	void SetFormula(LPCTSTR lpszNewValue);
	BOOL GetIsCalculated();
	long GetMemoryUsed();
	BOOL GetServerBased();
	void SetServerBased(BOOL bNewValue);
	void AutoSort(long Order, LPCTSTR Field);
	void AutoShow(long Type, long Range, long Count, LPCTSTR Field);
	long GetAutoSortOrder();
	CString GetAutoSortField();
	long GetAutoShowType();
	long GetAutoShowRange();
	long GetAutoShowCount();
	CString GetAutoShowField();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse PivotFields 

class PivotFields : public COleDispatchDriver
{
public:
	PivotFields() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	PivotFields(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	PivotFields(const PivotFields& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse CalculatedFields 

class CalculatedFields : public COleDispatchDriver
{
public:
	CalculatedFields() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	CalculatedFields(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CalculatedFields(const CalculatedFields& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Add(LPCTSTR Name, LPCTSTR Formula);
	LPDISPATCH Item(const VARIANT& Index);
	LPDISPATCH Get_Default(const VARIANT& Field);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse PivotItem 

class PivotItem : public COleDispatchDriver
{
public:
	PivotItem() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	PivotItem(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	PivotItem(const PivotItem& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	VARIANT GetChildItems(const VARIANT& Index);
	LPDISPATCH GetDataRange();
	CString Get_Default();
	void Set_Default(LPCTSTR lpszNewValue);
	LPDISPATCH GetLabelRange();
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	LPDISPATCH GetParentItem();
	BOOL GetParentShowDetail();
	long GetPosition();
	void SetPosition(long nNewValue);
	BOOL GetShowDetail();
	void SetShowDetail(BOOL bNewValue);
	VARIANT GetSourceName();
	CString GetValue();
	void SetValue(LPCTSTR lpszNewValue);
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	void Delete();
	BOOL GetIsCalculated();
	long GetRecordCount();
	CString GetFormula();
	void SetFormula(LPCTSTR lpszNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse PivotItems 

class PivotItems : public COleDispatchDriver
{
public:
	PivotItems() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	PivotItems(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	PivotItems(const PivotItems& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void Add(LPCTSTR Name);
	long GetCount();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse CalculatedItems 

class CalculatedItems : public COleDispatchDriver
{
public:
	CalculatedItems() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	CalculatedItems(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CalculatedItems(const CalculatedItems& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Add(LPCTSTR Name, LPCTSTR Formula);
	LPDISPATCH Item(const VARIANT& Index);
	LPDISPATCH Get_Default(const VARIANT& Field);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Characters 

class Characters : public COleDispatchDriver
{
public:
	Characters() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Characters(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Characters(const Characters& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	long GetCount();
	void Delete();
	LPDISPATCH GetFont();
	void Insert(LPCTSTR String);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	CString GetPhoneticCharacters();
	void SetPhoneticCharacters(LPCTSTR lpszNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Dialogs 

class Dialogs : public COleDispatchDriver
{
public:
	Dialogs() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Dialogs(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Dialogs(const Dialogs& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH GetItem(long Index);
	LPDISPATCH Get_Default(long Index);
	LPUNKNOWN Get_NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Dialog 

class Dialog : public COleDispatchDriver
{
public:
	Dialog() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Dialog(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Dialog(const Dialog& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	BOOL Show(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, const VARIANT& Arg11, 
		const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, const VARIANT& Arg21, 
		const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse SoundNote 

class SoundNote : public COleDispatchDriver
{
public:
	SoundNote() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	SoundNote(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	SoundNote(const SoundNote& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void Delete();
	void Import(LPCTSTR Filename);
	void Play();
	void Record();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Button 

class Button : public COleDispatchDriver
{
public:
	Button() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Button(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Button(const Button& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBottomRightCell();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	long GetIndex();
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	LPDISPATCH GetTopLeftCell();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	BOOL GetAddIndent();
	void SetAddIndent(BOOL bNewValue);
	VARIANT GetAutoScaleFont();
	void SetAutoScaleFont(const VARIANT& newValue);
	BOOL GetAutoSize();
	void SetAutoSize(BOOL bNewValue);
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	LPDISPATCH GetFont();
	CString GetFormula();
	void SetFormula(LPCTSTR lpszNewValue);
	VARIANT GetHorizontalAlignment();
	void SetHorizontalAlignment(const VARIANT& newValue);
	BOOL GetLockedText();
	void SetLockedText(BOOL bNewValue);
	VARIANT GetOrientation();
	void SetOrientation(const VARIANT& newValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	VARIANT GetVerticalAlignment();
	void SetVerticalAlignment(const VARIANT& newValue);
	long GetReadingOrder();
	void SetReadingOrder(long nNewValue);
	VARIANT GetAccelerator();
	void SetAccelerator(const VARIANT& newValue);
	BOOL GetCancelButton();
	void SetCancelButton(BOOL bNewValue);
	BOOL GetDefaultButton();
	void SetDefaultButton(BOOL bNewValue);
	BOOL GetDismissButton();
	void SetDismissButton(BOOL bNewValue);
	BOOL GetHelpButton();
	void SetHelpButton(BOOL bNewValue);
	VARIANT GetPhoneticAccelerator();
	void SetPhoneticAccelerator(const VARIANT& newValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Buttons 

class Buttons : public COleDispatchDriver
{
public:
	Buttons() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Buttons(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Buttons(const Buttons& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	BOOL GetAddIndent();
	void SetAddIndent(BOOL bNewValue);
	VARIANT GetAutoScaleFont();
	void SetAutoScaleFont(const VARIANT& newValue);
	BOOL GetAutoSize();
	void SetAutoSize(BOOL bNewValue);
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	LPDISPATCH GetFont();
	CString GetFormula();
	void SetFormula(LPCTSTR lpszNewValue);
	VARIANT GetHorizontalAlignment();
	void SetHorizontalAlignment(const VARIANT& newValue);
	BOOL GetLockedText();
	void SetLockedText(BOOL bNewValue);
	VARIANT GetOrientation();
	void SetOrientation(const VARIANT& newValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	VARIANT GetVerticalAlignment();
	void SetVerticalAlignment(const VARIANT& newValue);
	long GetReadingOrder();
	void SetReadingOrder(long nNewValue);
	VARIANT GetAccelerator();
	void SetAccelerator(const VARIANT& newValue);
	BOOL GetCancelButton();
	void SetCancelButton(BOOL bNewValue);
	BOOL GetDefaultButton();
	void SetDefaultButton(BOOL bNewValue);
	BOOL GetDismissButton();
	void SetDismissButton(BOOL bNewValue);
	BOOL GetHelpButton();
	void SetHelpButton(BOOL bNewValue);
	VARIANT GetPhoneticAccelerator();
	void SetPhoneticAccelerator(const VARIANT& newValue);
	LPDISPATCH Add(double Left, double Top, double Width, double Height);
	long GetCount();
	LPDISPATCH Group();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse CheckBox 

class CheckBox : public COleDispatchDriver
{
public:
	CheckBox() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	CheckBox(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CheckBox(const CheckBox& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBottomRightCell();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	long GetIndex();
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	LPDISPATCH GetTopLeftCell();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	BOOL GetLockedText();
	void SetLockedText(BOOL bNewValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	VARIANT GetAccelerator();
	void SetAccelerator(const VARIANT& newValue);
	LPDISPATCH GetBorder();
	long Get_Default();
	void Set_Default(long nNewValue);
	BOOL GetDisplay3DShading();
	void SetDisplay3DShading(BOOL bNewValue);
	LPDISPATCH GetInterior();
	CString GetLinkedCell();
	void SetLinkedCell(LPCTSTR lpszNewValue);
	VARIANT GetPhoneticAccelerator();
	void SetPhoneticAccelerator(const VARIANT& newValue);
	VARIANT GetValue();
	void SetValue(const VARIANT& newValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse CheckBoxes 

class CheckBoxes : public COleDispatchDriver
{
public:
	CheckBoxes() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	CheckBoxes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CheckBoxes(const CheckBoxes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	BOOL GetLockedText();
	void SetLockedText(BOOL bNewValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	VARIANT GetAccelerator();
	void SetAccelerator(const VARIANT& newValue);
	LPDISPATCH GetBorder();
	long Get_Default();
	void Set_Default(long nNewValue);
	BOOL GetDisplay3DShading();
	void SetDisplay3DShading(BOOL bNewValue);
	LPDISPATCH GetInterior();
	CString GetLinkedCell();
	void SetLinkedCell(LPCTSTR lpszNewValue);
	VARIANT GetPhoneticAccelerator();
	void SetPhoneticAccelerator(const VARIANT& newValue);
	VARIANT GetValue();
	void SetValue(const VARIANT& newValue);
	LPDISPATCH Add(double Left, double Top, double Width, double Height);
	long GetCount();
	LPDISPATCH Group();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse OptionButton 

class OptionButton : public COleDispatchDriver
{
public:
	OptionButton() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	OptionButton(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	OptionButton(const OptionButton& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBottomRightCell();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	long GetIndex();
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	LPDISPATCH GetTopLeftCell();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	BOOL GetLockedText();
	void SetLockedText(BOOL bNewValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	VARIANT GetAccelerator();
	void SetAccelerator(const VARIANT& newValue);
	LPDISPATCH GetBorder();
	long Get_Default();
	void Set_Default(long nNewValue);
	BOOL GetDisplay3DShading();
	void SetDisplay3DShading(BOOL bNewValue);
	LPDISPATCH GetInterior();
	CString GetLinkedCell();
	void SetLinkedCell(LPCTSTR lpszNewValue);
	VARIANT GetPhoneticAccelerator();
	void SetPhoneticAccelerator(const VARIANT& newValue);
	VARIANT GetValue();
	void SetValue(const VARIANT& newValue);
	LPDISPATCH GetGroupBox();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse OptionButtons 

class OptionButtons : public COleDispatchDriver
{
public:
	OptionButtons() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	OptionButtons(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	OptionButtons(const OptionButtons& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	BOOL GetLockedText();
	void SetLockedText(BOOL bNewValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	VARIANT GetAccelerator();
	void SetAccelerator(const VARIANT& newValue);
	LPDISPATCH GetBorder();
	long Get_Default();
	void Set_Default(long nNewValue);
	BOOL GetDisplay3DShading();
	void SetDisplay3DShading(BOOL bNewValue);
	LPDISPATCH GetInterior();
	CString GetLinkedCell();
	void SetLinkedCell(LPCTSTR lpszNewValue);
	VARIANT GetPhoneticAccelerator();
	void SetPhoneticAccelerator(const VARIANT& newValue);
	VARIANT GetValue();
	void SetValue(const VARIANT& newValue);
	LPDISPATCH GetGroupBox();
	LPDISPATCH Add(double Left, double Top, double Width, double Height);
	long GetCount();
	LPDISPATCH Group();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse EditBox 

class EditBox : public COleDispatchDriver
{
public:
	EditBox() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	EditBox(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	EditBox(const EditBox& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBottomRightCell();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	long GetIndex();
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	LPDISPATCH GetTopLeftCell();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	BOOL GetLockedText();
	void SetLockedText(BOOL bNewValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	BOOL GetDisplayVerticalScrollBar();
	void SetDisplayVerticalScrollBar(BOOL bNewValue);
	long GetInputType();
	void SetInputType(long nNewValue);
	CString GetLinkedObject();
	BOOL GetMultiLine();
	void SetMultiLine(BOOL bNewValue);
	BOOL GetPasswordEdit();
	void SetPasswordEdit(BOOL bNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse EditBoxes 

class EditBoxes : public COleDispatchDriver
{
public:
	EditBoxes() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	EditBoxes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	EditBoxes(const EditBoxes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	BOOL GetLockedText();
	void SetLockedText(BOOL bNewValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	BOOL GetDisplayVerticalScrollBar();
	void SetDisplayVerticalScrollBar(BOOL bNewValue);
	long GetInputType();
	void SetInputType(long nNewValue);
	BOOL GetMultiLine();
	void SetMultiLine(BOOL bNewValue);
	BOOL GetPasswordEdit();
	void SetPasswordEdit(BOOL bNewValue);
	LPDISPATCH Add(double Left, double Top, double Width, double Height);
	long GetCount();
	LPDISPATCH Group();
	VARIANT Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ScrollBar 

class ScrollBar : public COleDispatchDriver
{
public:
	ScrollBar() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ScrollBar(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ScrollBar(const ScrollBar& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBottomRightCell();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	long GetIndex();
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	LPDISPATCH GetTopLeftCell();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	long Get_Default();
	void Set_Default(long nNewValue);
	BOOL GetDisplay3DShading();
	void SetDisplay3DShading(BOOL bNewValue);
	CString GetLinkedCell();
	void SetLinkedCell(LPCTSTR lpszNewValue);
	long GetMax();
	void SetMax(long nNewValue);
	long GetMin();
	void SetMin(long nNewValue);
	long GetSmallChange();
	void SetSmallChange(long nNewValue);
	long GetValue();
	void SetValue(long nNewValue);
	long GetLargeChange();
	void SetLargeChange(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ScrollBars 

class ScrollBars : public COleDispatchDriver
{
public:
	ScrollBars() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ScrollBars(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ScrollBars(const ScrollBars& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	long Get_Default();
	void Set_Default(long nNewValue);
	BOOL GetDisplay3DShading();
	void SetDisplay3DShading(BOOL bNewValue);
	CString GetLinkedCell();
	void SetLinkedCell(LPCTSTR lpszNewValue);
	long GetMax();
	void SetMax(long nNewValue);
	long GetMin();
	void SetMin(long nNewValue);
	long GetSmallChange();
	void SetSmallChange(long nNewValue);
	long GetValue();
	void SetValue(long nNewValue);
	long GetLargeChange();
	void SetLargeChange(long nNewValue);
	LPDISPATCH Add(double Left, double Top, double Width, double Height);
	long GetCount();
	LPDISPATCH Group();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ListBox 

class ListBox : public COleDispatchDriver
{
public:
	ListBox() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ListBox(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ListBox(const ListBox& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBottomRightCell();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	long GetIndex();
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	LPDISPATCH GetTopLeftCell();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	void AddItem(const VARIANT& Text, const VARIANT& Index);
	BOOL GetDisplay3DShading();
	void SetDisplay3DShading(BOOL bNewValue);
	long Get_Default();
	void Set_Default(long nNewValue);
	CString GetLinkedCell();
	void SetLinkedCell(LPCTSTR lpszNewValue);
	VARIANT GetLinkedObject();
	VARIANT List(const VARIANT& Index);
	long GetListCount();
	CString GetListFillRange();
	void SetListFillRange(LPCTSTR lpszNewValue);
	long GetListIndex();
	void SetListIndex(long nNewValue);
	long GetMultiSelect();
	void SetMultiSelect(long nNewValue);
	void RemoveAllItems();
	void RemoveItem(long Index, const VARIANT& Count);
	VARIANT Selected(const VARIANT& Index);
	long GetValue();
	void SetValue(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ListBoxes 

class ListBoxes : public COleDispatchDriver
{
public:
	ListBoxes() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ListBoxes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ListBoxes(const ListBoxes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	void AddItem(const VARIANT& Text, const VARIANT& Index);
	BOOL GetDisplay3DShading();
	void SetDisplay3DShading(BOOL bNewValue);
	long Get_Default();
	void Set_Default(long nNewValue);
	CString GetLinkedCell();
	void SetLinkedCell(LPCTSTR lpszNewValue);
	VARIANT List(const VARIANT& Index);
	CString GetListFillRange();
	void SetListFillRange(LPCTSTR lpszNewValue);
	long GetListIndex();
	void SetListIndex(long nNewValue);
	long GetMultiSelect();
	void SetMultiSelect(long nNewValue);
	void RemoveAllItems();
	void RemoveItem(long Index, const VARIANT& Count);
	VARIANT Selected(const VARIANT& Index);
	long GetValue();
	void SetValue(long nNewValue);
	LPDISPATCH Add(double Left, double Top, double Width, double Height);
	long GetCount();
	LPDISPATCH Group();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse GroupBox 

class GroupBox : public COleDispatchDriver
{
public:
	GroupBox() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	GroupBox(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	GroupBox(const GroupBox& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBottomRightCell();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	long GetIndex();
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	LPDISPATCH GetTopLeftCell();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	BOOL GetLockedText();
	void SetLockedText(BOOL bNewValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	VARIANT GetAccelerator();
	void SetAccelerator(const VARIANT& newValue);
	BOOL GetDisplay3DShading();
	void SetDisplay3DShading(BOOL bNewValue);
	VARIANT GetPhoneticAccelerator();
	void SetPhoneticAccelerator(const VARIANT& newValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse GroupBoxes 

class GroupBoxes : public COleDispatchDriver
{
public:
	GroupBoxes() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	GroupBoxes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	GroupBoxes(const GroupBoxes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	BOOL GetLockedText();
	void SetLockedText(BOOL bNewValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	VARIANT GetAccelerator();
	void SetAccelerator(const VARIANT& newValue);
	BOOL GetDisplay3DShading();
	void SetDisplay3DShading(BOOL bNewValue);
	VARIANT GetPhoneticAccelerator();
	void SetPhoneticAccelerator(const VARIANT& newValue);
	LPDISPATCH Add(double Left, double Top, double Width, double Height);
	long GetCount();
	LPDISPATCH Group();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse DropDown 

class DropDown : public COleDispatchDriver
{
public:
	DropDown() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	DropDown(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	DropDown(const DropDown& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBottomRightCell();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	long GetIndex();
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	LPDISPATCH GetTopLeftCell();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	void AddItem(const VARIANT& Text, const VARIANT& Index);
	BOOL GetDisplay3DShading();
	void SetDisplay3DShading(BOOL bNewValue);
	long Get_Default();
	void Set_Default(long nNewValue);
	CString GetLinkedCell();
	void SetLinkedCell(LPCTSTR lpszNewValue);
	VARIANT GetLinkedObject();
	VARIANT List(const VARIANT& Index);
	long GetListCount();
	CString GetListFillRange();
	void SetListFillRange(LPCTSTR lpszNewValue);
	long GetListIndex();
	void SetListIndex(long nNewValue);
	void RemoveAllItems();
	void RemoveItem(long Index, const VARIANT& Count);
	VARIANT Selected(const VARIANT& Index);
	long GetValue();
	void SetValue(long nNewValue);
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	long GetDropDownLines();
	void SetDropDownLines(long nNewValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse DropDowns 

class DropDowns : public COleDispatchDriver
{
public:
	DropDowns() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	DropDowns(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	DropDowns(const DropDowns& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	void AddItem(const VARIANT& Text, const VARIANT& Index);
	BOOL GetDisplay3DShading();
	void SetDisplay3DShading(BOOL bNewValue);
	long Get_Default();
	void Set_Default(long nNewValue);
	CString GetLinkedCell();
	void SetLinkedCell(LPCTSTR lpszNewValue);
	VARIANT List(const VARIANT& Index);
	CString GetListFillRange();
	void SetListFillRange(LPCTSTR lpszNewValue);
	long GetListIndex();
	void SetListIndex(long nNewValue);
	void RemoveAllItems();
	void RemoveItem(long Index, const VARIANT& Count);
	VARIANT Selected(const VARIANT& Index);
	long GetValue();
	void SetValue(long nNewValue);
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	long GetDropDownLines();
	void SetDropDownLines(long nNewValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	LPDISPATCH Add(double Left, double Top, double Width, double Height, const VARIANT& Editable);
	long GetCount();
	LPDISPATCH Group();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Spinner 

class Spinner : public COleDispatchDriver
{
public:
	Spinner() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Spinner(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Spinner(const Spinner& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBottomRightCell();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	long GetIndex();
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	LPDISPATCH GetTopLeftCell();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	long Get_Default();
	void Set_Default(long nNewValue);
	BOOL GetDisplay3DShading();
	void SetDisplay3DShading(BOOL bNewValue);
	CString GetLinkedCell();
	void SetLinkedCell(LPCTSTR lpszNewValue);
	long GetMax();
	void SetMax(long nNewValue);
	long GetMin();
	void SetMin(long nNewValue);
	long GetSmallChange();
	void SetSmallChange(long nNewValue);
	long GetValue();
	void SetValue(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Spinners 

class Spinners : public COleDispatchDriver
{
public:
	Spinners() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Spinners(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Spinners(const Spinners& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	long Get_Default();
	void Set_Default(long nNewValue);
	BOOL GetDisplay3DShading();
	void SetDisplay3DShading(BOOL bNewValue);
	CString GetLinkedCell();
	void SetLinkedCell(LPCTSTR lpszNewValue);
	long GetMax();
	void SetMax(long nNewValue);
	long GetMin();
	void SetMin(long nNewValue);
	long GetSmallChange();
	void SetSmallChange(long nNewValue);
	long GetValue();
	void SetValue(long nNewValue);
	LPDISPATCH Add(double Left, double Top, double Width, double Height);
	long GetCount();
	LPDISPATCH Group();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse DialogFrame 

class DialogFrame : public COleDispatchDriver
{
public:
	DialogFrame() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	DialogFrame(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	DialogFrame(const DialogFrame& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void CopyPicture(long Appearance, long Format);
	double GetHeight();
	void SetHeight(double newValue);
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	void Select(const VARIANT& Replace);
	double GetTop();
	void SetTop(double newValue);
	double GetWidth();
	void SetWidth(double newValue);
	LPDISPATCH GetShapeRange();
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	BOOL GetLockedText();
	void SetLockedText(BOOL bNewValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Label 

class Label : public COleDispatchDriver
{
public:
	Label() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Label(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Label(const Label& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBottomRightCell();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	long GetIndex();
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	LPDISPATCH GetTopLeftCell();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	BOOL GetLockedText();
	void SetLockedText(BOOL bNewValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	VARIANT GetAccelerator();
	void SetAccelerator(const VARIANT& newValue);
	VARIANT GetPhoneticAccelerator();
	void SetPhoneticAccelerator(const VARIANT& newValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Labels 

class Labels : public COleDispatchDriver
{
public:
	Labels() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Labels(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Labels(const Labels& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	BOOL GetLockedText();
	void SetLockedText(BOOL bNewValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	VARIANT GetAccelerator();
	void SetAccelerator(const VARIANT& newValue);
	VARIANT GetPhoneticAccelerator();
	void SetPhoneticAccelerator(const VARIANT& newValue);
	LPDISPATCH Add(double Left, double Top, double Width, double Height);
	long GetCount();
	LPDISPATCH Group();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Panes 

class Panes : public COleDispatchDriver
{
public:
	Panes() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Panes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Panes(const Panes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH GetItem(long Index);
	LPDISPATCH Get_Default(long Index);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Pane 

class Pane : public COleDispatchDriver
{
public:
	Pane() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Pane(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Pane(const Pane& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void Activate();
	long GetIndex();
	void LargeScroll(const VARIANT& Down, const VARIANT& Up, const VARIANT& ToRight, const VARIANT& ToLeft);
	long GetScrollColumn();
	void SetScrollColumn(long nNewValue);
	long GetScrollRow();
	void SetScrollRow(long nNewValue);
	void SmallScroll(const VARIANT& Down, const VARIANT& Up, const VARIANT& ToRight, const VARIANT& ToLeft);
	LPDISPATCH GetVisibleRange();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Scenarios 

class Scenarios : public COleDispatchDriver
{
public:
	Scenarios() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Scenarios(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Scenarios(const Scenarios& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH Add(LPCTSTR Name, const VARIANT& ChangingCells, const VARIANT& Values, const VARIANT& Comment, const VARIANT& Locked, const VARIANT& Hidden);
	long GetCount();
	void CreateSummary(long ReportType, const VARIANT& ResultCells);
	LPDISPATCH Item(const VARIANT& Index);
	void Merge(const VARIANT& Source);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Scenario 

class Scenario : public COleDispatchDriver
{
public:
	Scenario() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Scenario(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Scenario(const Scenario& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void ChangeScenario(const VARIANT& ChangingCells, const VARIANT& Values);
	LPDISPATCH GetChangingCells();
	CString GetComment();
	void SetComment(LPCTSTR lpszNewValue);
	void Delete();
	BOOL GetHidden();
	void SetHidden(BOOL bNewValue);
	long GetIndex();
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	void Show();
	VARIANT GetValues(const VARIANT& Index);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse GroupObject 

class GroupObject : public COleDispatchDriver
{
public:
	GroupObject() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	GroupObject(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	GroupObject(const GroupObject& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBottomRightCell();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	long GetIndex();
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	LPDISPATCH GetTopLeftCell();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	BOOL GetAddIndent();
	void SetAddIndent(BOOL bNewValue);
	VARIANT GetArrowHeadLength();
	void SetArrowHeadLength(const VARIANT& newValue);
	VARIANT GetArrowHeadStyle();
	void SetArrowHeadStyle(const VARIANT& newValue);
	VARIANT GetArrowHeadWidth();
	void SetArrowHeadWidth(const VARIANT& newValue);
	BOOL GetAutoSize();
	void SetAutoSize(BOOL bNewValue);
	LPDISPATCH GetBorder();
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	long Get_Default();
	void Set_Default(long nNewValue);
	LPDISPATCH GetFont();
	VARIANT GetHorizontalAlignment();
	void SetHorizontalAlignment(const VARIANT& newValue);
	LPDISPATCH GetInterior();
	VARIANT GetOrientation();
	void SetOrientation(const VARIANT& newValue);
	BOOL GetRoundedCorners();
	void SetRoundedCorners(BOOL bNewValue);
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
	LPDISPATCH Ungroup();
	VARIANT GetVerticalAlignment();
	void SetVerticalAlignment(const VARIANT& newValue);
	long GetReadingOrder();
	void SetReadingOrder(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse GroupObjects 

class GroupObjects : public COleDispatchDriver
{
public:
	GroupObjects() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	GroupObjects(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	GroupObjects(const GroupObjects& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	BOOL GetAddIndent();
	void SetAddIndent(BOOL bNewValue);
	VARIANT GetArrowHeadLength();
	void SetArrowHeadLength(const VARIANT& newValue);
	VARIANT GetArrowHeadStyle();
	void SetArrowHeadStyle(const VARIANT& newValue);
	VARIANT GetArrowHeadWidth();
	void SetArrowHeadWidth(const VARIANT& newValue);
	BOOL GetAutoSize();
	void SetAutoSize(BOOL bNewValue);
	LPDISPATCH GetBorder();
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	long Get_Default();
	void Set_Default(long nNewValue);
	LPDISPATCH GetFont();
	VARIANT GetHorizontalAlignment();
	void SetHorizontalAlignment(const VARIANT& newValue);
	LPDISPATCH GetInterior();
	VARIANT GetOrientation();
	void SetOrientation(const VARIANT& newValue);
	BOOL GetRoundedCorners();
	void SetRoundedCorners(BOOL bNewValue);
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
	LPDISPATCH Ungroup();
	VARIANT GetVerticalAlignment();
	void SetVerticalAlignment(const VARIANT& newValue);
	long GetReadingOrder();
	void SetReadingOrder(long nNewValue);
	long GetCount();
	LPDISPATCH Group();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Line 

class Line : public COleDispatchDriver
{
public:
	Line() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Line(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Line(const Line& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBottomRightCell();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	long GetIndex();
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	LPDISPATCH GetTopLeftCell();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	VARIANT GetArrowHeadLength();
	void SetArrowHeadLength(const VARIANT& newValue);
	VARIANT GetArrowHeadStyle();
	void SetArrowHeadStyle(const VARIANT& newValue);
	VARIANT GetArrowHeadWidth();
	void SetArrowHeadWidth(const VARIANT& newValue);
	LPDISPATCH GetBorder();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Lines 

class Lines : public COleDispatchDriver
{
public:
	Lines() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Lines(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Lines(const Lines& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	VARIANT GetArrowHeadLength();
	void SetArrowHeadLength(const VARIANT& newValue);
	VARIANT GetArrowHeadStyle();
	void SetArrowHeadStyle(const VARIANT& newValue);
	VARIANT GetArrowHeadWidth();
	void SetArrowHeadWidth(const VARIANT& newValue);
	LPDISPATCH GetBorder();
	LPDISPATCH Add(double X1, double Y1, double X2, double Y2);
	long GetCount();
	LPDISPATCH Group();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Rectangle 

class Rectangle : public COleDispatchDriver
{
public:
	Rectangle() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Rectangle(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Rectangle(const Rectangle& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBottomRightCell();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	long GetIndex();
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	LPDISPATCH GetTopLeftCell();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	BOOL GetAddIndent();
	void SetAddIndent(BOOL bNewValue);
	VARIANT GetAutoScaleFont();
	void SetAutoScaleFont(const VARIANT& newValue);
	BOOL GetAutoSize();
	void SetAutoSize(BOOL bNewValue);
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	LPDISPATCH GetFont();
	CString GetFormula();
	void SetFormula(LPCTSTR lpszNewValue);
	VARIANT GetHorizontalAlignment();
	void SetHorizontalAlignment(const VARIANT& newValue);
	BOOL GetLockedText();
	void SetLockedText(BOOL bNewValue);
	VARIANT GetOrientation();
	void SetOrientation(const VARIANT& newValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	VARIANT GetVerticalAlignment();
	void SetVerticalAlignment(const VARIANT& newValue);
	long GetReadingOrder();
	void SetReadingOrder(long nNewValue);
	LPDISPATCH GetBorder();
	LPDISPATCH GetInterior();
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
	BOOL GetRoundedCorners();
	void SetRoundedCorners(BOOL bNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Rectangles 

class Rectangles : public COleDispatchDriver
{
public:
	Rectangles() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Rectangles(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Rectangles(const Rectangles& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	BOOL GetAddIndent();
	void SetAddIndent(BOOL bNewValue);
	VARIANT GetAutoScaleFont();
	void SetAutoScaleFont(const VARIANT& newValue);
	BOOL GetAutoSize();
	void SetAutoSize(BOOL bNewValue);
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	LPDISPATCH GetFont();
	CString GetFormula();
	void SetFormula(LPCTSTR lpszNewValue);
	VARIANT GetHorizontalAlignment();
	void SetHorizontalAlignment(const VARIANT& newValue);
	BOOL GetLockedText();
	void SetLockedText(BOOL bNewValue);
	VARIANT GetOrientation();
	void SetOrientation(const VARIANT& newValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	VARIANT GetVerticalAlignment();
	void SetVerticalAlignment(const VARIANT& newValue);
	long GetReadingOrder();
	void SetReadingOrder(long nNewValue);
	LPDISPATCH GetBorder();
	LPDISPATCH GetInterior();
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
	BOOL GetRoundedCorners();
	void SetRoundedCorners(BOOL bNewValue);
	LPDISPATCH Add(double Left, double Top, double Width, double Height);
	long GetCount();
	LPDISPATCH Group();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Oval 

class Oval : public COleDispatchDriver
{
public:
	Oval() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Oval(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Oval(const Oval& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBottomRightCell();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	long GetIndex();
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	LPDISPATCH GetTopLeftCell();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	BOOL GetAddIndent();
	void SetAddIndent(BOOL bNewValue);
	VARIANT GetAutoScaleFont();
	void SetAutoScaleFont(const VARIANT& newValue);
	BOOL GetAutoSize();
	void SetAutoSize(BOOL bNewValue);
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	LPDISPATCH GetFont();
	CString GetFormula();
	void SetFormula(LPCTSTR lpszNewValue);
	VARIANT GetHorizontalAlignment();
	void SetHorizontalAlignment(const VARIANT& newValue);
	BOOL GetLockedText();
	void SetLockedText(BOOL bNewValue);
	VARIANT GetOrientation();
	void SetOrientation(const VARIANT& newValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	VARIANT GetVerticalAlignment();
	void SetVerticalAlignment(const VARIANT& newValue);
	long GetReadingOrder();
	void SetReadingOrder(long nNewValue);
	LPDISPATCH GetBorder();
	LPDISPATCH GetInterior();
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Ovals 

class Ovals : public COleDispatchDriver
{
public:
	Ovals() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Ovals(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Ovals(const Ovals& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	BOOL GetAddIndent();
	void SetAddIndent(BOOL bNewValue);
	VARIANT GetAutoScaleFont();
	void SetAutoScaleFont(const VARIANT& newValue);
	BOOL GetAutoSize();
	void SetAutoSize(BOOL bNewValue);
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	LPDISPATCH GetFont();
	CString GetFormula();
	void SetFormula(LPCTSTR lpszNewValue);
	VARIANT GetHorizontalAlignment();
	void SetHorizontalAlignment(const VARIANT& newValue);
	BOOL GetLockedText();
	void SetLockedText(BOOL bNewValue);
	VARIANT GetOrientation();
	void SetOrientation(const VARIANT& newValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	VARIANT GetVerticalAlignment();
	void SetVerticalAlignment(const VARIANT& newValue);
	long GetReadingOrder();
	void SetReadingOrder(long nNewValue);
	LPDISPATCH GetBorder();
	LPDISPATCH GetInterior();
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
	LPDISPATCH Add(double Left, double Top, double Width, double Height);
	long GetCount();
	LPDISPATCH Group();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Arc 

class Arc : public COleDispatchDriver
{
public:
	Arc() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Arc(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Arc(const Arc& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBottomRightCell();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	long GetIndex();
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	LPDISPATCH GetTopLeftCell();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	BOOL GetAddIndent();
	void SetAddIndent(BOOL bNewValue);
	VARIANT GetAutoScaleFont();
	void SetAutoScaleFont(const VARIANT& newValue);
	BOOL GetAutoSize();
	void SetAutoSize(BOOL bNewValue);
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	LPDISPATCH GetFont();
	CString GetFormula();
	void SetFormula(LPCTSTR lpszNewValue);
	VARIANT GetHorizontalAlignment();
	void SetHorizontalAlignment(const VARIANT& newValue);
	BOOL GetLockedText();
	void SetLockedText(BOOL bNewValue);
	VARIANT GetOrientation();
	void SetOrientation(const VARIANT& newValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	VARIANT GetVerticalAlignment();
	void SetVerticalAlignment(const VARIANT& newValue);
	long GetReadingOrder();
	void SetReadingOrder(long nNewValue);
	LPDISPATCH GetBorder();
	LPDISPATCH GetInterior();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Arcs 

class Arcs : public COleDispatchDriver
{
public:
	Arcs() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Arcs(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Arcs(const Arcs& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	BOOL GetAddIndent();
	void SetAddIndent(BOOL bNewValue);
	VARIANT GetAutoScaleFont();
	void SetAutoScaleFont(const VARIANT& newValue);
	BOOL GetAutoSize();
	void SetAutoSize(BOOL bNewValue);
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	LPDISPATCH GetFont();
	CString GetFormula();
	void SetFormula(LPCTSTR lpszNewValue);
	VARIANT GetHorizontalAlignment();
	void SetHorizontalAlignment(const VARIANT& newValue);
	BOOL GetLockedText();
	void SetLockedText(BOOL bNewValue);
	VARIANT GetOrientation();
	void SetOrientation(const VARIANT& newValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	VARIANT GetVerticalAlignment();
	void SetVerticalAlignment(const VARIANT& newValue);
	long GetReadingOrder();
	void SetReadingOrder(long nNewValue);
	LPDISPATCH GetBorder();
	LPDISPATCH GetInterior();
	LPDISPATCH Add(double X1, double Y1, double X2, double Y2);
	long GetCount();
	LPDISPATCH Group();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse OLEObjectEvents 

class OLEObjectEvents : public COleDispatchDriver
{
public:
	OLEObjectEvents() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	OLEObjectEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	OLEObjectEvents(const OLEObjectEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	void GotFocus();
	void LostFocus();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse _OLEObject 

class _OLEObject : public COleDispatchDriver
{
public:
	_OLEObject() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	_OLEObject(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	_OLEObject(const _OLEObject& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBottomRightCell();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	long GetIndex();
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	LPDISPATCH GetTopLeftCell();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	LPDISPATCH GetBorder();
	LPDISPATCH GetInterior();
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
	void Activate();
	BOOL GetAutoLoad();
	void SetAutoLoad(BOOL bNewValue);
	BOOL GetAutoUpdate();
	void SetAutoUpdate(BOOL bNewValue);
	LPDISPATCH GetObject();
	VARIANT GetOLEType();
	CString GetSourceName();
	void SetSourceName(LPCTSTR lpszNewValue);
	void Update();
	void Verb(long Verb);
	CString GetLinkedCell();
	void SetLinkedCell(LPCTSTR lpszNewValue);
	CString GetListFillRange();
	void SetListFillRange(LPCTSTR lpszNewValue);
	CString GetProgId();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse OLEObjects 

class OLEObjects : public COleDispatchDriver
{
public:
	OLEObjects() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	OLEObjects(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	OLEObjects(const OLEObjects& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	LPDISPATCH GetBorder();
	LPDISPATCH GetInterior();
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
	BOOL GetAutoLoad();
	void SetAutoLoad(BOOL bNewValue);
	CString GetSourceName();
	void SetSourceName(LPCTSTR lpszNewValue);
	LPDISPATCH Add(const VARIANT& ClassType, const VARIANT& Filename, const VARIANT& Link, const VARIANT& DisplayAsIcon, const VARIANT& IconFileName, const VARIANT& IconIndex, const VARIANT& IconLabel, const VARIANT& Left, const VARIANT& Top, 
		const VARIANT& Width, const VARIANT& Height);
	long GetCount();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse TextBox 

class TextBox : public COleDispatchDriver
{
public:
	TextBox() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	TextBox(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	TextBox(const TextBox& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBottomRightCell();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	long GetIndex();
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	LPDISPATCH GetTopLeftCell();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	BOOL GetAddIndent();
	void SetAddIndent(BOOL bNewValue);
	VARIANT GetAutoScaleFont();
	void SetAutoScaleFont(const VARIANT& newValue);
	BOOL GetAutoSize();
	void SetAutoSize(BOOL bNewValue);
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	LPDISPATCH GetFont();
	CString GetFormula();
	void SetFormula(LPCTSTR lpszNewValue);
	VARIANT GetHorizontalAlignment();
	void SetHorizontalAlignment(const VARIANT& newValue);
	BOOL GetLockedText();
	void SetLockedText(BOOL bNewValue);
	VARIANT GetOrientation();
	void SetOrientation(const VARIANT& newValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	VARIANT GetVerticalAlignment();
	void SetVerticalAlignment(const VARIANT& newValue);
	long GetReadingOrder();
	void SetReadingOrder(long nNewValue);
	LPDISPATCH GetBorder();
	LPDISPATCH GetInterior();
	BOOL GetRoundedCorners();
	void SetRoundedCorners(BOOL bNewValue);
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse TextBoxes 

class TextBoxes : public COleDispatchDriver
{
public:
	TextBoxes() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	TextBoxes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	TextBoxes(const TextBoxes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	BOOL GetAddIndent();
	void SetAddIndent(BOOL bNewValue);
	VARIANT GetAutoScaleFont();
	void SetAutoScaleFont(const VARIANT& newValue);
	BOOL GetAutoSize();
	void SetAutoSize(BOOL bNewValue);
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	LPDISPATCH GetFont();
	CString GetFormula();
	void SetFormula(LPCTSTR lpszNewValue);
	VARIANT GetHorizontalAlignment();
	void SetHorizontalAlignment(const VARIANT& newValue);
	BOOL GetLockedText();
	void SetLockedText(BOOL bNewValue);
	VARIANT GetOrientation();
	void SetOrientation(const VARIANT& newValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	VARIANT GetVerticalAlignment();
	void SetVerticalAlignment(const VARIANT& newValue);
	long GetReadingOrder();
	void SetReadingOrder(long nNewValue);
	LPDISPATCH GetBorder();
	LPDISPATCH GetInterior();
	BOOL GetRoundedCorners();
	void SetRoundedCorners(BOOL bNewValue);
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
	LPDISPATCH Add(double Left, double Top, double Width, double Height);
	long GetCount();
	LPDISPATCH Group();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Picture 

class Picture : public COleDispatchDriver
{
public:
	Picture() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Picture(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Picture(const Picture& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBottomRightCell();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	long GetIndex();
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	LPDISPATCH GetTopLeftCell();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	LPDISPATCH GetBorder();
	LPDISPATCH GetInterior();
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
	CString GetFormula();
	void SetFormula(LPCTSTR lpszNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Pictures 

class Pictures : public COleDispatchDriver
{
public:
	Pictures() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Pictures(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Pictures(const Pictures& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	LPDISPATCH GetBorder();
	LPDISPATCH GetInterior();
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
	CString GetFormula();
	void SetFormula(LPCTSTR lpszNewValue);
	LPDISPATCH Add(double Left, double Top, double Width, double Height);
	long GetCount();
	LPDISPATCH Group();
	LPDISPATCH Insert(LPCTSTR Filename, const VARIANT& Converter);
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
	LPDISPATCH Paste(const VARIANT& Link);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Drawing 

class Drawing : public COleDispatchDriver
{
public:
	Drawing() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Drawing(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Drawing(const Drawing& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBottomRightCell();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	long GetIndex();
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	LPDISPATCH GetTopLeftCell();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	BOOL GetAddIndent();
	void SetAddIndent(BOOL bNewValue);
	VARIANT GetAutoScaleFont();
	void SetAutoScaleFont(const VARIANT& newValue);
	BOOL GetAutoSize();
	void SetAutoSize(BOOL bNewValue);
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	LPDISPATCH GetFont();
	CString GetFormula();
	void SetFormula(LPCTSTR lpszNewValue);
	VARIANT GetHorizontalAlignment();
	void SetHorizontalAlignment(const VARIANT& newValue);
	BOOL GetLockedText();
	void SetLockedText(BOOL bNewValue);
	VARIANT GetOrientation();
	void SetOrientation(const VARIANT& newValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	VARIANT GetVerticalAlignment();
	void SetVerticalAlignment(const VARIANT& newValue);
	long GetReadingOrder();
	void SetReadingOrder(long nNewValue);
	LPDISPATCH GetBorder();
	LPDISPATCH GetInterior();
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
	void AddVertex(double Left, double Top);
	void Reshape(long Vertex, BOOL Insert, const VARIANT& Left, const VARIANT& Top);
	VARIANT GetVertices(const VARIANT& Index1, const VARIANT& Index2);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Drawings 

class Drawings : public COleDispatchDriver
{
public:
	Drawings() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Drawings(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Drawings(const Drawings& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	BOOL GetAddIndent();
	void SetAddIndent(BOOL bNewValue);
	VARIANT GetAutoScaleFont();
	void SetAutoScaleFont(const VARIANT& newValue);
	BOOL GetAutoSize();
	void SetAutoSize(BOOL bNewValue);
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	LPDISPATCH GetFont();
	CString GetFormula();
	void SetFormula(LPCTSTR lpszNewValue);
	VARIANT GetHorizontalAlignment();
	void SetHorizontalAlignment(const VARIANT& newValue);
	BOOL GetLockedText();
	void SetLockedText(BOOL bNewValue);
	VARIANT GetOrientation();
	void SetOrientation(const VARIANT& newValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	VARIANT GetVerticalAlignment();
	void SetVerticalAlignment(const VARIANT& newValue);
	long GetReadingOrder();
	void SetReadingOrder(long nNewValue);
	LPDISPATCH GetBorder();
	LPDISPATCH GetInterior();
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
	void Reshape(long Vertex, BOOL Insert, const VARIANT& Left, const VARIANT& Top);
	LPDISPATCH Add(double X1, double Y1, double X2, double Y2, BOOL Closed);
	long GetCount();
	LPDISPATCH Group();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse RoutingSlip 

class RoutingSlip : public COleDispatchDriver
{
public:
	RoutingSlip() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	RoutingSlip(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	RoutingSlip(const RoutingSlip& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetDelivery();
	void SetDelivery(long nNewValue);
	VARIANT GetMessage();
	void SetMessage(const VARIANT& newValue);
	VARIANT Recipients(const VARIANT& Index);
	void Reset();
	BOOL GetReturnWhenDone();
	void SetReturnWhenDone(BOOL bNewValue);
	long GetStatus();
	VARIANT GetSubject();
	void SetSubject(const VARIANT& newValue);
	BOOL GetTrackStatus();
	void SetTrackStatus(BOOL bNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Outline 

class Outline : public COleDispatchDriver
{
public:
	Outline() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Outline(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Outline(const Outline& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	BOOL GetAutomaticStyles();
	void SetAutomaticStyles(BOOL bNewValue);
	void ShowLevels(const VARIANT& RowLevels, const VARIANT& ColumnLevels);
	long GetSummaryColumn();
	void SetSummaryColumn(long nNewValue);
	long GetSummaryRow();
	void SetSummaryRow(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Module 

class Module : public COleDispatchDriver
{
public:
	Module() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Module(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Module(const Module& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void Activate();
	void Copy(const VARIANT& Before, const VARIANT& After);
	void Delete();
	CString GetCodeName();
	CString Get_CodeName();
	void Set_CodeName(LPCTSTR lpszNewValue);
	long GetIndex();
	void Move(const VARIANT& Before, const VARIANT& After);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	LPDISPATCH GetNext();
	LPDISPATCH GetPageSetup();
	LPDISPATCH GetPrevious();
	void PrintOut(const VARIANT& From, const VARIANT& To, const VARIANT& Copies, const VARIANT& Preview, const VARIANT& ActivePrinter, const VARIANT& PrintToFile, const VARIANT& Collate);
	void Protect(const VARIANT& Password, const VARIANT& DrawingObjects, const VARIANT& Contents, const VARIANT& Scenarios, const VARIANT& UserInterfaceOnly);
	BOOL GetProtectContents();
	BOOL GetProtectionMode();
	void SaveAs(LPCTSTR Filename, const VARIANT& FileFormat, const VARIANT& Password, const VARIANT& WriteResPassword, const VARIANT& ReadOnlyRecommended, const VARIANT& CreateBackup, const VARIANT& AddToMru, const VARIANT& TextCodepage, 
		const VARIANT& TextVisualLayout);
	void Select(const VARIANT& Replace);
	void Unprotect(const VARIANT& Password);
	long GetVisible();
	void SetVisible(long nNewValue);
	LPDISPATCH GetShapes();
	void InsertFile(const VARIANT& Filename, const VARIANT& Merge);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Modules 

class Modules : public COleDispatchDriver
{
public:
	Modules() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Modules(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Modules(const Modules& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH Add(const VARIANT& Before, const VARIANT& After, const VARIANT& Count);
	void Copy(const VARIANT& Before, const VARIANT& After);
	long GetCount();
	void Delete();
	LPDISPATCH GetItem(const VARIANT& Index);
	void Move(const VARIANT& Before, const VARIANT& After);
	LPUNKNOWN Get_NewEnum();
	void PrintOut(const VARIANT& From, const VARIANT& To, const VARIANT& Copies, const VARIANT& Preview, const VARIANT& ActivePrinter, const VARIANT& PrintToFile, const VARIANT& Collate);
	void Select(const VARIANT& Replace);
	LPDISPATCH GetHPageBreaks();
	LPDISPATCH GetVPageBreaks();
	VARIANT GetVisible();
	void SetVisible(const VARIANT& newValue);
	LPDISPATCH Get_Default(const VARIANT& Index);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse DialogSheet 

class DialogSheet : public COleDispatchDriver
{
public:
	DialogSheet() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	DialogSheet(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	DialogSheet(const DialogSheet& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void Activate();
	void Copy(const VARIANT& Before, const VARIANT& After);
	void Delete();
	CString GetCodeName();
	CString Get_CodeName();
	void Set_CodeName(LPCTSTR lpszNewValue);
	long GetIndex();
	void Move(const VARIANT& Before, const VARIANT& After);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	LPDISPATCH GetNext();
	LPDISPATCH GetPageSetup();
	LPDISPATCH GetPrevious();
	void PrintOut(const VARIANT& From, const VARIANT& To, const VARIANT& Copies, const VARIANT& Preview, const VARIANT& ActivePrinter, const VARIANT& PrintToFile, const VARIANT& Collate);
	void PrintPreview(const VARIANT& EnableChanges);
	void Protect(const VARIANT& Password, const VARIANT& DrawingObjects, const VARIANT& Contents, const VARIANT& Scenarios, const VARIANT& UserInterfaceOnly);
	BOOL GetProtectContents();
	BOOL GetProtectDrawingObjects();
	BOOL GetProtectionMode();
	BOOL GetProtectScenarios();
	void SaveAs(LPCTSTR Filename, const VARIANT& FileFormat, const VARIANT& Password, const VARIANT& WriteResPassword, const VARIANT& ReadOnlyRecommended, const VARIANT& CreateBackup, const VARIANT& AddToMru, const VARIANT& TextCodepage, 
		const VARIANT& TextVisualLayout);
	void Select(const VARIANT& Replace);
	void Unprotect(const VARIANT& Password);
	long GetVisible();
	void SetVisible(long nNewValue);
	LPDISPATCH GetShapes();
	BOOL GetEnableCalculation();
	void SetEnableCalculation(BOOL bNewValue);
	LPDISPATCH ChartObjects(const VARIANT& Index);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript);
	BOOL GetEnableAutoFilter();
	void SetEnableAutoFilter(BOOL bNewValue);
	long GetEnableSelection();
	void SetEnableSelection(long nNewValue);
	BOOL GetEnableOutlining();
	void SetEnableOutlining(BOOL bNewValue);
	BOOL GetEnablePivotTable();
	void SetEnablePivotTable(BOOL bNewValue);
	VARIANT Evaluate(const VARIANT& Name);
	VARIANT _Evaluate(const VARIANT& Name);
	void ResetAllPageBreaks();
	LPDISPATCH GetNames();
	LPDISPATCH OLEObjects(const VARIANT& Index);
	void Paste(const VARIANT& Destination, const VARIANT& Link);
	void PasteSpecial(const VARIANT& Format, const VARIANT& Link, const VARIANT& DisplayAsIcon, const VARIANT& IconFileName, const VARIANT& IconIndex, const VARIANT& IconLabel);
	CString GetScrollArea();
	void SetScrollArea(LPCTSTR lpszNewValue);
	LPDISPATCH GetHPageBreaks();
	LPDISPATCH GetVPageBreaks();
	LPDISPATCH GetQueryTables();
	BOOL GetDisplayPageBreaks();
	void SetDisplayPageBreaks(BOOL bNewValue);
	LPDISPATCH GetComments();
	LPDISPATCH GetHyperlinks();
	void ClearCircles();
	void CircleInvalid();
	LPDISPATCH GetAutoFilter();
	VARIANT GetDefaultButton();
	void SetDefaultButton(const VARIANT& newValue);
	VARIANT GetFocus();
	void SetFocus(const VARIANT& newValue);
	BOOL Hide(const VARIANT& Cancel);
	BOOL Show();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse DialogSheets 

class DialogSheets : public COleDispatchDriver
{
public:
	DialogSheets() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	DialogSheets(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	DialogSheets(const DialogSheets& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH Add(const VARIANT& Before, const VARIANT& After, const VARIANT& Count);
	void Copy(const VARIANT& Before, const VARIANT& After);
	long GetCount();
	void Delete();
	LPDISPATCH GetItem(const VARIANT& Index);
	void Move(const VARIANT& Before, const VARIANT& After);
	LPUNKNOWN Get_NewEnum();
	void PrintOut(const VARIANT& From, const VARIANT& To, const VARIANT& Copies, const VARIANT& Preview, const VARIANT& ActivePrinter, const VARIANT& PrintToFile, const VARIANT& Collate);
	void PrintPreview(const VARIANT& EnableChanges);
	void Select(const VARIANT& Replace);
	LPDISPATCH GetHPageBreaks();
	LPDISPATCH GetVPageBreaks();
	VARIANT GetVisible();
	void SetVisible(const VARIANT& newValue);
	LPDISPATCH Get_Default(const VARIANT& Index);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Worksheets 

class Worksheets : public COleDispatchDriver
{
public:
	Worksheets() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Worksheets(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Worksheets(const Worksheets& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH Add(const VARIANT& Before, const VARIANT& After, const VARIANT& Count, const VARIANT& Type);
	void Copy(const VARIANT& Before, const VARIANT& After);
	long GetCount();
	void Delete();
	void FillAcrossSheets(LPDISPATCH Range, long Type);
	LPDISPATCH GetItem(const VARIANT& Index);
	void Move(const VARIANT& Before, const VARIANT& After);
	LPUNKNOWN Get_NewEnum();
	void PrintOut(const VARIANT& From, const VARIANT& To, const VARIANT& Copies, const VARIANT& Preview, const VARIANT& ActivePrinter, const VARIANT& PrintToFile, const VARIANT& Collate);
	void PrintPreview(const VARIANT& EnableChanges);
	void Select(const VARIANT& Replace);
	LPDISPATCH GetHPageBreaks();
	LPDISPATCH GetVPageBreaks();
	VARIANT GetVisible();
	void SetVisible(const VARIANT& newValue);
	LPDISPATCH Get_Default(const VARIANT& Index);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse PageSetup 

class PageSetup : public COleDispatchDriver
{
public:
	PageSetup() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	PageSetup(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	PageSetup(const PageSetup& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	BOOL GetBlackAndWhite();
	void SetBlackAndWhite(BOOL bNewValue);
	double GetBottomMargin();
	void SetBottomMargin(double newValue);
	CString GetCenterFooter();
	void SetCenterFooter(LPCTSTR lpszNewValue);
	CString GetCenterHeader();
	void SetCenterHeader(LPCTSTR lpszNewValue);
	BOOL GetCenterHorizontally();
	void SetCenterHorizontally(BOOL bNewValue);
	BOOL GetCenterVertically();
	void SetCenterVertically(BOOL bNewValue);
	long GetChartSize();
	void SetChartSize(long nNewValue);
	BOOL GetDraft();
	void SetDraft(BOOL bNewValue);
	long GetFirstPageNumber();
	void SetFirstPageNumber(long nNewValue);
	VARIANT GetFitToPagesTall();
	void SetFitToPagesTall(const VARIANT& newValue);
	VARIANT GetFitToPagesWide();
	void SetFitToPagesWide(const VARIANT& newValue);
	double GetFooterMargin();
	void SetFooterMargin(double newValue);
	double GetHeaderMargin();
	void SetHeaderMargin(double newValue);
	CString GetLeftFooter();
	void SetLeftFooter(LPCTSTR lpszNewValue);
	CString GetLeftHeader();
	void SetLeftHeader(LPCTSTR lpszNewValue);
	double GetLeftMargin();
	void SetLeftMargin(double newValue);
	long GetOrder();
	void SetOrder(long nNewValue);
	long GetOrientation();
	void SetOrientation(long nNewValue);
	long GetPaperSize();
	void SetPaperSize(long nNewValue);
	CString GetPrintArea();
	void SetPrintArea(LPCTSTR lpszNewValue);
	BOOL GetPrintGridlines();
	void SetPrintGridlines(BOOL bNewValue);
	BOOL GetPrintHeadings();
	void SetPrintHeadings(BOOL bNewValue);
	BOOL GetPrintNotes();
	void SetPrintNotes(BOOL bNewValue);
	VARIANT PrintQuality(const VARIANT& Index);
	CString GetPrintTitleColumns();
	void SetPrintTitleColumns(LPCTSTR lpszNewValue);
	CString GetPrintTitleRows();
	void SetPrintTitleRows(LPCTSTR lpszNewValue);
	CString GetRightFooter();
	void SetRightFooter(LPCTSTR lpszNewValue);
	CString GetRightHeader();
	void SetRightHeader(LPCTSTR lpszNewValue);
	double GetRightMargin();
	void SetRightMargin(double newValue);
	double GetTopMargin();
	void SetTopMargin(double newValue);
	VARIANT GetZoom();
	void SetZoom(const VARIANT& newValue);
	long GetPrintComments();
	void SetPrintComments(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Names 

class Names : public COleDispatchDriver
{
public:
	Names() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Names(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Names(const Names& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH Add(const VARIANT& Name, const VARIANT& RefersTo, const VARIANT& Visible, const VARIANT& MacroType, const VARIANT& ShortcutKey, const VARIANT& Category, const VARIANT& NameLocal, const VARIANT& RefersToLocal, 
		const VARIANT& CategoryLocal, const VARIANT& RefersToR1C1, const VARIANT& RefersToR1C1Local);
	LPDISPATCH Item(const VARIANT& Index, const VARIANT& IndexLocal, const VARIANT& RefersTo);
	LPDISPATCH _Default(const VARIANT& Index, const VARIANT& IndexLocal, const VARIANT& RefersTo);
	long GetCount();
	LPUNKNOWN Get_NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Name 

class Name : public COleDispatchDriver
{
public:
	Name() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Name(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Name(const Name& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString Get_Default();
	long GetIndex();
	CString GetCategory();
	void SetCategory(LPCTSTR lpszNewValue);
	CString GetCategoryLocal();
	void SetCategoryLocal(LPCTSTR lpszNewValue);
	void Delete();
	long GetMacroType();
	void SetMacroType(long nNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	VARIANT GetRefersTo();
	void SetRefersTo(const VARIANT& newValue);
	CString GetShortcutKey();
	void SetShortcutKey(LPCTSTR lpszNewValue);
	CString GetValue();
	void SetValue(LPCTSTR lpszNewValue);
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	CString GetNameLocal();
	void SetNameLocal(LPCTSTR lpszNewValue);
	VARIANT GetRefersToLocal();
	void SetRefersToLocal(const VARIANT& newValue);
	VARIANT GetRefersToR1C1();
	void SetRefersToR1C1(const VARIANT& newValue);
	VARIANT GetRefersToR1C1Local();
	void SetRefersToR1C1Local(const VARIANT& newValue);
	LPDISPATCH GetRefersToRange();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ChartObject 

class ChartObject : public COleDispatchDriver
{
public:
	ChartObject() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ChartObject(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ChartObject(const ChartObject& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBottomRightCell();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	long GetIndex();
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	LPDISPATCH GetTopLeftCell();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	long GetZOrder();
	LPDISPATCH GetShapeRange();
	void Activate();
	LPDISPATCH GetChart();
	BOOL GetProtectChartObject();
	void SetProtectChartObject(BOOL bNewValue);
	BOOL GetRoundedCorners();
	void SetRoundedCorners(BOOL bNewValue);
	LPDISPATCH GetBorder();
	LPDISPATCH GetInterior();
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ChartObjects 

class ChartObjects : public COleDispatchDriver
{
public:
	ChartObjects() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ChartObjects(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ChartObjects(const ChartObjects& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void BringToFront();
	void Copy();
	void CopyPicture(long Appearance, long Format);
	void Cut();
	void Delete();
	LPDISPATCH Duplicate();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	double GetHeight();
	void SetHeight(double newValue);
	double GetLeft();
	void SetLeft(double newValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	VARIANT GetPlacement();
	void SetPlacement(const VARIANT& newValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	void Select(const VARIANT& Replace);
	void SendToBack();
	double GetTop();
	void SetTop(double newValue);
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	double GetWidth();
	void SetWidth(double newValue);
	LPDISPATCH GetShapeRange();
	BOOL GetRoundedCorners();
	void SetRoundedCorners(BOOL bNewValue);
	LPDISPATCH GetBorder();
	LPDISPATCH GetInterior();
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
	LPDISPATCH Add(double Left, double Top, double Width, double Height);
	long GetCount();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Mailer 

class Mailer : public COleDispatchDriver
{
public:
	Mailer() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Mailer(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Mailer(const Mailer& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	VARIANT GetBCCRecipients();
	void SetBCCRecipients(const VARIANT& newValue);
	VARIANT GetCCRecipients();
	void SetCCRecipients(const VARIANT& newValue);
	VARIANT GetEnclosures();
	void SetEnclosures(const VARIANT& newValue);
	BOOL GetReceived();
	DATE GetSendDateTime();
	CString GetSender();
	CString GetSubject();
	void SetSubject(LPCTSTR lpszNewValue);
	VARIANT GetToRecipients();
	void SetToRecipients(const VARIANT& newValue);
	VARIANT GetWhichAddress();
	void SetWhichAddress(const VARIANT& newValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse CustomViews 

class CustomViews : public COleDispatchDriver
{
public:
	CustomViews() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	CustomViews(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CustomViews(const CustomViews& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Item(const VARIANT& ViewName);
	LPDISPATCH Add(LPCTSTR ViewName, const VARIANT& PrintSettings, const VARIANT& RowColSettings);
	LPDISPATCH Get_Default(const VARIANT& ViewName);
	LPUNKNOWN Get_NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse CustomView 

class CustomView : public COleDispatchDriver
{
public:
	CustomView() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	CustomView(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CustomView(const CustomView& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetName();
	BOOL GetPrintSettings();
	BOOL GetRowColSettings();
	void Show();
	void Delete();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse FormatConditions 

class FormatConditions : public COleDispatchDriver
{
public:
	FormatConditions() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	FormatConditions(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	FormatConditions(const FormatConditions& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Item(const VARIANT& Index);
	LPDISPATCH Add(long Type, const VARIANT& Operator, const VARIANT& Formula1, const VARIANT& Formula2);
	LPDISPATCH Get_Default(const VARIANT& Index);
	LPUNKNOWN Get_NewEnum();
	void Delete();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse FormatCondition 

class FormatCondition : public COleDispatchDriver
{
public:
	FormatCondition() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	FormatCondition(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	FormatCondition(const FormatCondition& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void Modify(long Type, const VARIANT& Operator, const VARIANT& Formula1, const VARIANT& Formula2);
	long GetType();
	long GetOperator();
	CString GetFormula1();
	CString GetFormula2();
	LPDISPATCH GetInterior();
	LPDISPATCH GetBorders();
	LPDISPATCH GetFont();
	void Delete();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Comments 

class Comments : public COleDispatchDriver
{
public:
	Comments() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Comments(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Comments(const Comments& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Item(long Index);
	LPDISPATCH Get_Default(long Index);
	LPUNKNOWN Get_NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Comment 

class Comment : public COleDispatchDriver
{
public:
	Comment() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Comment(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Comment(const Comment& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetAuthor();
	LPDISPATCH GetShape();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	CString Text(const VARIANT& Text, const VARIANT& Start, const VARIANT& Overwrite);
	void Delete();
	LPDISPATCH Next();
	LPDISPATCH Previous();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse RefreshEvents 

class RefreshEvents : public COleDispatchDriver
{
public:
	RefreshEvents() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	RefreshEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	RefreshEvents(const RefreshEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	void BeforeRefresh(BOOL* Cancel);
	void AfterRefresh(BOOL Success);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse _QueryTable 

class _QueryTable : public COleDispatchDriver
{
public:
	_QueryTable() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	_QueryTable(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	_QueryTable(const _QueryTable& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	BOOL GetFieldNames();
	void SetFieldNames(BOOL bNewValue);
	BOOL GetRowNumbers();
	void SetRowNumbers(BOOL bNewValue);
	BOOL GetFillAdjacentFormulas();
	void SetFillAdjacentFormulas(BOOL bNewValue);
	BOOL GetHasAutoFormat();
	void SetHasAutoFormat(BOOL bNewValue);
	BOOL GetRefreshOnFileOpen();
	void SetRefreshOnFileOpen(BOOL bNewValue);
	BOOL GetRefreshing();
	BOOL GetFetchedRowOverflow();
	BOOL GetBackgroundQuery();
	void SetBackgroundQuery(BOOL bNewValue);
	void CancelRefresh();
	long GetRefreshStyle();
	void SetRefreshStyle(long nNewValue);
	BOOL GetEnableRefresh();
	void SetEnableRefresh(BOOL bNewValue);
	BOOL GetSavePassword();
	void SetSavePassword(BOOL bNewValue);
	LPDISPATCH GetDestination();
	VARIANT GetConnection();
	void SetConnection(const VARIANT& newValue);
	VARIANT GetSql();
	void SetSql(const VARIANT& newValue);
	CString GetPostText();
	void SetPostText(LPCTSTR lpszNewValue);
	LPDISPATCH GetResultRange();
	void Delete();
	BOOL Refresh(const VARIANT& BackgroundQuery);
	LPDISPATCH GetParameters();
	LPDISPATCH GetRecordset();
	void SetRefRecordset(LPDISPATCH newValue);
	BOOL GetSaveData();
	void SetSaveData(BOOL bNewValue);
	BOOL GetTablesOnlyFromHTML();
	void SetTablesOnlyFromHTML(BOOL bNewValue);
	BOOL GetEnableEditing();
	void SetEnableEditing(BOOL bNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse QueryTables 

class QueryTables : public COleDispatchDriver
{
public:
	QueryTables() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	QueryTables(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	QueryTables(const QueryTables& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Add(const VARIANT& Connection, LPDISPATCH Destination, const VARIANT& Sql);
	LPDISPATCH Item(const VARIANT& Index);
	LPDISPATCH Get_Default(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Parameter 

class Parameter : public COleDispatchDriver
{
public:
	Parameter() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Parameter(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Parameter(const Parameter& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetDataType();
	void SetDataType(long nNewValue);
	long GetType();
	CString GetPromptString();
	VARIANT GetValue();
	LPDISPATCH GetSourceRange();
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	void SetParam(long Type, const VARIANT& Value);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Parameters 

class Parameters : public COleDispatchDriver
{
public:
	Parameters() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Parameters(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Parameters(const Parameters& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH Add(LPCTSTR Name, const VARIANT& iDataType);
	long GetCount();
	LPDISPATCH Item(const VARIANT& Index);
	LPDISPATCH Get_Default(const VARIANT& Index);
	void Delete();
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ODBCError 

class ODBCError : public COleDispatchDriver
{
public:
	ODBCError() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ODBCError(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ODBCError(const ODBCError& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetSqlState();
	CString GetErrorString();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ODBCErrors 

class ODBCErrors : public COleDispatchDriver
{
public:
	ODBCErrors() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ODBCErrors(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ODBCErrors(const ODBCErrors& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Item(long Index);
	LPDISPATCH Get_Default(long Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Validation 

class Validation : public COleDispatchDriver
{
public:
	Validation() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Validation(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Validation(const Validation& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void Add(long Type, const VARIANT& AlertStyle, const VARIANT& Operator, const VARIANT& Formula1, const VARIANT& Formula2);
	long GetAlertStyle();
	BOOL GetIgnoreBlank();
	void SetIgnoreBlank(BOOL bNewValue);
	long GetIMEMode();
	void SetIMEMode(long nNewValue);
	BOOL GetInCellDropdown();
	void SetInCellDropdown(BOOL bNewValue);
	void Delete();
	CString GetErrorMessage();
	void SetErrorMessage(LPCTSTR lpszNewValue);
	CString GetErrorTitle();
	void SetErrorTitle(LPCTSTR lpszNewValue);
	CString GetInputMessage();
	void SetInputMessage(LPCTSTR lpszNewValue);
	CString GetInputTitle();
	void SetInputTitle(LPCTSTR lpszNewValue);
	CString GetFormula1();
	CString GetFormula2();
	void Modify(const VARIANT& Type, const VARIANT& AlertStyle, const VARIANT& Operator, const VARIANT& Formula1, const VARIANT& Formula2);
	long GetOperator();
	BOOL GetShowError();
	void SetShowError(BOOL bNewValue);
	BOOL GetShowInput();
	void SetShowInput(BOOL bNewValue);
	long GetType();
	BOOL GetValue();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Hyperlinks 

class Hyperlinks : public COleDispatchDriver
{
public:
	Hyperlinks() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Hyperlinks(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Hyperlinks(const Hyperlinks& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH Add(LPDISPATCH Anchor, LPCTSTR Address, const VARIANT& SubAddress);
	long GetCount();
	LPDISPATCH GetItem(const VARIANT& Index);
	LPDISPATCH Get_Default(const VARIANT& Index);
	LPUNKNOWN Get_NewEnum();
	void Delete();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Hyperlink 

class Hyperlink : public COleDispatchDriver
{
public:
	Hyperlink() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Hyperlink(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Hyperlink(const Hyperlink& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetName();
	LPDISPATCH GetRange();
	LPDISPATCH GetShape();
	CString GetSubAddress();
	void SetSubAddress(LPCTSTR lpszNewValue);
	CString GetAddress();
	void SetAddress(LPCTSTR lpszNewValue);
	long GetType();
	void AddToFavorites();
	void Delete();
	void Follow(const VARIANT& NewWindow, const VARIANT& AddHistory, const VARIANT& ExtraInfo, const VARIANT& Method, const VARIANT& HeaderInfo);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse AutoFilter 

class AutoFilter : public COleDispatchDriver
{
public:
	AutoFilter() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	AutoFilter(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	AutoFilter(const AutoFilter& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetRange();
	LPDISPATCH GetFilters();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Filters 

class Filters : public COleDispatchDriver
{
public:
	Filters() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Filters(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Filters(const Filters& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Get_Default(long Index);
	LPDISPATCH GetItem(long Index);
	LPUNKNOWN Get_NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Filter 

class Filter : public COleDispatchDriver
{
public:
	Filter() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Filter(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Filter(const Filter& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	BOOL GetOn();
	VARIANT GetCriteria1();
	long GetOperator();
	VARIANT GetCriteria2();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse AutoCorrect 

class AutoCorrect : public COleDispatchDriver
{
public:
	AutoCorrect() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	AutoCorrect(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	AutoCorrect(const AutoCorrect& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void AddReplacement(LPCTSTR What, LPCTSTR Replacement);
	BOOL GetCapitalizeNamesOfDays();
	void SetCapitalizeNamesOfDays(BOOL bNewValue);
	void DeleteReplacement(LPCTSTR What);
	VARIANT ReplacementList(const VARIANT& Index);
	BOOL GetReplaceText();
	void SetReplaceText(BOOL bNewValue);
	BOOL GetTwoInitialCapitals();
	void SetTwoInitialCapitals(BOOL bNewValue);
	BOOL GetCorrectSentenceCap();
	void SetCorrectSentenceCap(BOOL bNewValue);
	BOOL GetCorrectCapsLock();
	void SetCorrectCapsLock(BOOL bNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Border 

class Border : public COleDispatchDriver
{
public:
	Border() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Border(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Border(const Border& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	VARIANT GetColor();
	void SetColor(const VARIANT& newValue);
	VARIANT GetColorIndex();
	void SetColorIndex(const VARIANT& newValue);
	VARIANT GetLineStyle();
	void SetLineStyle(const VARIANT& newValue);
	VARIANT GetWeight();
	void SetWeight(const VARIANT& newValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Interior 

class Interior : public COleDispatchDriver
{
public:
	Interior() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Interior(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Interior(const Interior& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	VARIANT GetColor();
	void SetColor(const VARIANT& newValue);
	VARIANT GetColorIndex();
	void SetColorIndex(const VARIANT& newValue);
	VARIANT GetInvertIfNegative();
	void SetInvertIfNegative(const VARIANT& newValue);
	VARIANT GetPattern();
	void SetPattern(const VARIANT& newValue);
	VARIANT GetPatternColor();
	void SetPatternColor(const VARIANT& newValue);
	VARIANT GetPatternColorIndex();
	void SetPatternColorIndex(const VARIANT& newValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ChartFillFormat 

class ChartFillFormat : public COleDispatchDriver
{
public:
	ChartFillFormat() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ChartFillFormat(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ChartFillFormat(const ChartFillFormat& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void OneColorGradient(long Style, long Variant, float Degree);
	void TwoColorGradient(long Style, long Variant);
	void PresetTextured(long PresetTexture);
	void Solid();
	void Patterned(long Pattern);
	void UserPicture(const VARIANT& PictureFile, const VARIANT& PictureFormat, const VARIANT& PictureStackUnit, const VARIANT& PicturePlacement);
	void UserTextured(LPCTSTR TextureFile);
	void PresetGradient(long Style, long Variant, long PresetGradientType);
	LPDISPATCH GetBackColor();
	LPDISPATCH GetForeColor();
	long GetGradientColorType();
	float GetGradientDegree();
	long GetGradientStyle();
	long GetGradientVariant();
	long GetPattern();
	long GetPresetGradientType();
	long GetPresetTexture();
	CString GetTextureName();
	long GetTextureType();
	long GetType();
	long GetVisible();
	void SetVisible(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ChartColorFormat 

class ChartColorFormat : public COleDispatchDriver
{
public:
	ChartColorFormat() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ChartColorFormat(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ChartColorFormat(const ChartColorFormat& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetSchemeColor();
	void SetSchemeColor(long nNewValue);
	long GetRgb();
	long Get_Default();
	long GetType();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Axis 

class Axis : public COleDispatchDriver
{
public:
	Axis() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Axis(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Axis(const Axis& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	BOOL GetAxisBetweenCategories();
	void SetAxisBetweenCategories(BOOL bNewValue);
	long GetAxisGroup();
	LPDISPATCH GetAxisTitle();
	LPDISPATCH GetBorder();
	VARIANT GetCategoryNames();
	void SetCategoryNames(const VARIANT& newValue);
	long GetCrosses();
	void SetCrosses(long nNewValue);
	double GetCrossesAt();
	void SetCrossesAt(double newValue);
	void Delete();
	BOOL GetHasMajorGridlines();
	void SetHasMajorGridlines(BOOL bNewValue);
	BOOL GetHasMinorGridlines();
	void SetHasMinorGridlines(BOOL bNewValue);
	BOOL GetHasTitle();
	void SetHasTitle(BOOL bNewValue);
	LPDISPATCH GetMajorGridlines();
	long GetMajorTickMark();
	void SetMajorTickMark(long nNewValue);
	double GetMajorUnit();
	void SetMajorUnit(double newValue);
	BOOL GetMajorUnitIsAuto();
	void SetMajorUnitIsAuto(BOOL bNewValue);
	double GetMaximumScale();
	void SetMaximumScale(double newValue);
	BOOL GetMaximumScaleIsAuto();
	void SetMaximumScaleIsAuto(BOOL bNewValue);
	double GetMinimumScale();
	void SetMinimumScale(double newValue);
	BOOL GetMinimumScaleIsAuto();
	void SetMinimumScaleIsAuto(BOOL bNewValue);
	LPDISPATCH GetMinorGridlines();
	long GetMinorTickMark();
	void SetMinorTickMark(long nNewValue);
	double GetMinorUnit();
	void SetMinorUnit(double newValue);
	BOOL GetMinorUnitIsAuto();
	void SetMinorUnitIsAuto(BOOL bNewValue);
	BOOL GetReversePlotOrder();
	void SetReversePlotOrder(BOOL bNewValue);
	long GetScaleType();
	void SetScaleType(long nNewValue);
	void Select();
	long GetTickLabelPosition();
	void SetTickLabelPosition(long nNewValue);
	LPDISPATCH GetTickLabels();
	long GetTickLabelSpacing();
	void SetTickLabelSpacing(long nNewValue);
	long GetTickMarkSpacing();
	void SetTickMarkSpacing(long nNewValue);
	long GetType();
	void SetType(long nNewValue);
	long GetBaseUnit();
	void SetBaseUnit(long nNewValue);
	BOOL GetBaseUnitIsAuto();
	void SetBaseUnitIsAuto(BOOL bNewValue);
	long GetMajorUnitScale();
	void SetMajorUnitScale(long nNewValue);
	long GetMinorUnitScale();
	void SetMinorUnitScale(long nNewValue);
	long GetCategoryType();
	void SetCategoryType(long nNewValue);
	double GetLeft();
	double GetTop();
	double GetWidth();
	double GetHeight();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ChartTitle 

class ChartTitle : public COleDispatchDriver
{
public:
	ChartTitle() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ChartTitle(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ChartTitle(const ChartTitle& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetName();
	void Select();
	LPDISPATCH GetBorder();
	void Delete();
	LPDISPATCH GetInterior();
	LPDISPATCH GetFill();
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	LPDISPATCH GetFont();
	VARIANT GetHorizontalAlignment();
	void SetHorizontalAlignment(const VARIANT& newValue);
	double GetLeft();
	void SetLeft(double newValue);
	VARIANT GetOrientation();
	void SetOrientation(const VARIANT& newValue);
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	double GetTop();
	void SetTop(double newValue);
	VARIANT GetVerticalAlignment();
	void SetVerticalAlignment(const VARIANT& newValue);
	long GetReadingOrder();
	void SetReadingOrder(long nNewValue);
	VARIANT GetAutoScaleFont();
	void SetAutoScaleFont(const VARIANT& newValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse AxisTitle 

class AxisTitle : public COleDispatchDriver
{
public:
	AxisTitle() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	AxisTitle(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	AxisTitle(const AxisTitle& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetName();
	void Select();
	LPDISPATCH GetBorder();
	void Delete();
	LPDISPATCH GetInterior();
	LPDISPATCH GetFill();
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	LPDISPATCH GetFont();
	VARIANT GetHorizontalAlignment();
	void SetHorizontalAlignment(const VARIANT& newValue);
	double GetLeft();
	void SetLeft(double newValue);
	VARIANT GetOrientation();
	void SetOrientation(const VARIANT& newValue);
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	double GetTop();
	void SetTop(double newValue);
	VARIANT GetVerticalAlignment();
	void SetVerticalAlignment(const VARIANT& newValue);
	long GetReadingOrder();
	void SetReadingOrder(long nNewValue);
	VARIANT GetAutoScaleFont();
	void SetAutoScaleFont(const VARIANT& newValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ChartGroup 

class ChartGroup : public COleDispatchDriver
{
public:
	ChartGroup() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ChartGroup(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ChartGroup(const ChartGroup& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetAxisGroup();
	void SetAxisGroup(long nNewValue);
	long GetDoughnutHoleSize();
	void SetDoughnutHoleSize(long nNewValue);
	LPDISPATCH GetDownBars();
	LPDISPATCH GetDropLines();
	long GetFirstSliceAngle();
	void SetFirstSliceAngle(long nNewValue);
	long GetGapWidth();
	void SetGapWidth(long nNewValue);
	BOOL GetHasDropLines();
	void SetHasDropLines(BOOL bNewValue);
	BOOL GetHasHiLoLines();
	void SetHasHiLoLines(BOOL bNewValue);
	BOOL GetHasRadarAxisLabels();
	void SetHasRadarAxisLabels(BOOL bNewValue);
	BOOL GetHasSeriesLines();
	void SetHasSeriesLines(BOOL bNewValue);
	BOOL GetHasUpDownBars();
	void SetHasUpDownBars(BOOL bNewValue);
	LPDISPATCH GetHiLoLines();
	long GetIndex();
	long GetOverlap();
	void SetOverlap(long nNewValue);
	LPDISPATCH GetRadarAxisLabels();
	LPDISPATCH SeriesCollection(const VARIANT& Index);
	LPDISPATCH GetSeriesLines();
	LPDISPATCH GetUpBars();
	BOOL GetVaryByCategories();
	void SetVaryByCategories(BOOL bNewValue);
	long GetSizeRepresents();
	void SetSizeRepresents(long nNewValue);
	long GetBubbleScale();
	void SetBubbleScale(long nNewValue);
	BOOL GetShowNegativeBubbles();
	void SetShowNegativeBubbles(BOOL bNewValue);
	long GetSplitType();
	void SetSplitType(long nNewValue);
	VARIANT GetSplitValue();
	void SetSplitValue(const VARIANT& newValue);
	long GetSecondPlotSize();
	void SetSecondPlotSize(long nNewValue);
	BOOL GetHas3DShading();
	void SetHas3DShading(BOOL bNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ChartGroups 

class ChartGroups : public COleDispatchDriver
{
public:
	ChartGroups() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ChartGroups(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ChartGroups(const ChartGroups& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Axes 

class Axes : public COleDispatchDriver
{
public:
	Axes() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Axes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Axes(const Axes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Item(long Type, long AxisGroup);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Points 

class Points : public COleDispatchDriver
{
public:
	Points() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Points(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Points(const Points& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Item(long Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Point 

class Point : public COleDispatchDriver
{
public:
	Point() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Point(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Point(const Point& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void ApplyDataLabels(long Type, const VARIANT& LegendKey, const VARIANT& AutoText);
	LPDISPATCH GetBorder();
	void ClearFormats();
	void Copy();
	LPDISPATCH GetDataLabel();
	void Delete();
	long GetExplosion();
	void SetExplosion(long nNewValue);
	BOOL GetHasDataLabel();
	void SetHasDataLabel(BOOL bNewValue);
	LPDISPATCH GetInterior();
	BOOL GetInvertIfNegative();
	void SetInvertIfNegative(BOOL bNewValue);
	long GetMarkerBackgroundColor();
	void SetMarkerBackgroundColor(long nNewValue);
	long GetMarkerBackgroundColorIndex();
	void SetMarkerBackgroundColorIndex(long nNewValue);
	long GetMarkerForegroundColor();
	void SetMarkerForegroundColor(long nNewValue);
	long GetMarkerForegroundColorIndex();
	void SetMarkerForegroundColorIndex(long nNewValue);
	long GetMarkerSize();
	void SetMarkerSize(long nNewValue);
	long GetMarkerStyle();
	void SetMarkerStyle(long nNewValue);
	void Paste();
	long GetPictureType();
	void SetPictureType(long nNewValue);
	long GetPictureUnit();
	void SetPictureUnit(long nNewValue);
	void Select();
	BOOL GetApplyPictToSides();
	void SetApplyPictToSides(BOOL bNewValue);
	BOOL GetApplyPictToFront();
	void SetApplyPictToFront(BOOL bNewValue);
	BOOL GetApplyPictToEnd();
	void SetApplyPictToEnd(BOOL bNewValue);
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
	BOOL GetSecondaryPlot();
	void SetSecondaryPlot(BOOL bNewValue);
	LPDISPATCH GetFill();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Series 

class Series : public COleDispatchDriver
{
public:
	Series() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Series(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Series(const Series& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void ApplyDataLabels(long Type, const VARIANT& LegendKey, const VARIANT& AutoText, const VARIANT& HasLeaderLines);
	long GetAxisGroup();
	void SetAxisGroup(long nNewValue);
	LPDISPATCH GetBorder();
	void ClearFormats();
	void Copy();
	LPDISPATCH DataLabels(const VARIANT& Index);
	void Delete();
	void ErrorBar(long Direction, long Include, long Type, const VARIANT& Amount, const VARIANT& MinusValues);
	LPDISPATCH GetErrorBars();
	long GetExplosion();
	void SetExplosion(long nNewValue);
	CString GetFormula();
	void SetFormula(LPCTSTR lpszNewValue);
	CString GetFormulaLocal();
	void SetFormulaLocal(LPCTSTR lpszNewValue);
	CString GetFormulaR1C1();
	void SetFormulaR1C1(LPCTSTR lpszNewValue);
	CString GetFormulaR1C1Local();
	void SetFormulaR1C1Local(LPCTSTR lpszNewValue);
	BOOL GetHasDataLabels();
	void SetHasDataLabels(BOOL bNewValue);
	BOOL GetHasErrorBars();
	void SetHasErrorBars(BOOL bNewValue);
	LPDISPATCH GetInterior();
	LPDISPATCH GetFill();
	BOOL GetInvertIfNegative();
	void SetInvertIfNegative(BOOL bNewValue);
	long GetMarkerBackgroundColor();
	void SetMarkerBackgroundColor(long nNewValue);
	long GetMarkerBackgroundColorIndex();
	void SetMarkerBackgroundColorIndex(long nNewValue);
	long GetMarkerForegroundColor();
	void SetMarkerForegroundColor(long nNewValue);
	long GetMarkerForegroundColorIndex();
	void SetMarkerForegroundColorIndex(long nNewValue);
	long GetMarkerSize();
	void SetMarkerSize(long nNewValue);
	long GetMarkerStyle();
	void SetMarkerStyle(long nNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	void Paste();
	long GetPictureType();
	void SetPictureType(long nNewValue);
	long GetPictureUnit();
	void SetPictureUnit(long nNewValue);
	long GetPlotOrder();
	void SetPlotOrder(long nNewValue);
	LPDISPATCH Points(const VARIANT& Index);
	void Select();
	BOOL GetSmooth();
	void SetSmooth(BOOL bNewValue);
	LPDISPATCH Trendlines(const VARIANT& Index);
	long GetType();
	void SetType(long nNewValue);
	long GetChartType();
	void SetChartType(long nNewValue);
	void ApplyCustomType(long ChartType);
	VARIANT GetValues();
	void SetValues(const VARIANT& newValue);
	VARIANT GetXValues();
	void SetXValues(const VARIANT& newValue);
	VARIANT GetBubbleSizes();
	void SetBubbleSizes(const VARIANT& newValue);
	long GetBarShape();
	void SetBarShape(long nNewValue);
	BOOL GetApplyPictToSides();
	void SetApplyPictToSides(BOOL bNewValue);
	BOOL GetApplyPictToFront();
	void SetApplyPictToFront(BOOL bNewValue);
	BOOL GetApplyPictToEnd();
	void SetApplyPictToEnd(BOOL bNewValue);
	BOOL GetHas3DEffect();
	void SetHas3DEffect(BOOL bNewValue);
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
	BOOL GetHasLeaderLines();
	void SetHasLeaderLines(BOOL bNewValue);
	LPDISPATCH GetLeaderLines();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse SeriesCollection 

class SeriesCollection : public COleDispatchDriver
{
public:
	SeriesCollection() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	SeriesCollection(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	SeriesCollection(const SeriesCollection& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH Add(const VARIANT& Source, long Rowcol, const VARIANT& SeriesLabels, const VARIANT& CategoryLabels, const VARIANT& Replace);
	long GetCount();
	void Extend(const VARIANT& Source, const VARIANT& Rowcol, const VARIANT& CategoryLabels);
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
	void Paste(long Rowcol, const VARIANT& SeriesLabels, const VARIANT& CategoryLabels, const VARIANT& Replace, const VARIANT& NewSeries);
	LPDISPATCH NewSeries();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse DataLabel 

class DataLabel : public COleDispatchDriver
{
public:
	DataLabel() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	DataLabel(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	DataLabel(const DataLabel& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetName();
	void Select();
	LPDISPATCH GetBorder();
	void Delete();
	LPDISPATCH GetInterior();
	LPDISPATCH GetFill();
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	LPDISPATCH GetCharacters(const VARIANT& Start, const VARIANT& Length);
	LPDISPATCH GetFont();
	VARIANT GetHorizontalAlignment();
	void SetHorizontalAlignment(const VARIANT& newValue);
	double GetLeft();
	void SetLeft(double newValue);
	VARIANT GetOrientation();
	void SetOrientation(const VARIANT& newValue);
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	double GetTop();
	void SetTop(double newValue);
	VARIANT GetVerticalAlignment();
	void SetVerticalAlignment(const VARIANT& newValue);
	long GetReadingOrder();
	void SetReadingOrder(long nNewValue);
	VARIANT GetAutoScaleFont();
	void SetAutoScaleFont(const VARIANT& newValue);
	BOOL GetAutoText();
	void SetAutoText(BOOL bNewValue);
	CString GetNumberFormat();
	void SetNumberFormat(LPCTSTR lpszNewValue);
	BOOL GetNumberFormatLinked();
	void SetNumberFormatLinked(BOOL bNewValue);
	BOOL GetShowLegendKey();
	void SetShowLegendKey(BOOL bNewValue);
	VARIANT GetType();
	void SetType(const VARIANT& newValue);
	long GetPosition();
	void SetPosition(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse DataLabels 

class DataLabels : public COleDispatchDriver
{
public:
	DataLabels() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	DataLabels(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	DataLabels(const DataLabels& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetName();
	void Select();
	LPDISPATCH GetBorder();
	void Delete();
	LPDISPATCH GetInterior();
	LPDISPATCH GetFill();
	LPDISPATCH GetFont();
	VARIANT GetHorizontalAlignment();
	void SetHorizontalAlignment(const VARIANT& newValue);
	VARIANT GetOrientation();
	void SetOrientation(const VARIANT& newValue);
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
	VARIANT GetVerticalAlignment();
	void SetVerticalAlignment(const VARIANT& newValue);
	long GetReadingOrder();
	void SetReadingOrder(long nNewValue);
	VARIANT GetAutoScaleFont();
	void SetAutoScaleFont(const VARIANT& newValue);
	BOOL GetAutoText();
	void SetAutoText(BOOL bNewValue);
	CString GetNumberFormat();
	void SetNumberFormat(LPCTSTR lpszNewValue);
	BOOL GetNumberFormatLinked();
	void SetNumberFormatLinked(BOOL bNewValue);
	BOOL GetShowLegendKey();
	void SetShowLegendKey(BOOL bNewValue);
	VARIANT GetType();
	void SetType(const VARIANT& newValue);
	long GetPosition();
	void SetPosition(long nNewValue);
	long GetCount();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse LegendEntry 

class LegendEntry : public COleDispatchDriver
{
public:
	LegendEntry() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	LegendEntry(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	LegendEntry(const LegendEntry& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void Delete();
	LPDISPATCH GetFont();
	long GetIndex();
	LPDISPATCH GetLegendKey();
	void Select();
	VARIANT GetAutoScaleFont();
	void SetAutoScaleFont(const VARIANT& newValue);
	double GetLeft();
	double GetTop();
	double GetWidth();
	double GetHeight();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse LegendEntries 

class LegendEntries : public COleDispatchDriver
{
public:
	LegendEntries() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	LegendEntries(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	LegendEntries(const LegendEntries& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse LegendKey 

class LegendKey : public COleDispatchDriver
{
public:
	LegendKey() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	LegendKey(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	LegendKey(const LegendKey& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBorder();
	void ClearFormats();
	void Delete();
	LPDISPATCH GetInterior();
	LPDISPATCH GetFill();
	BOOL GetInvertIfNegative();
	void SetInvertIfNegative(BOOL bNewValue);
	long GetMarkerBackgroundColor();
	void SetMarkerBackgroundColor(long nNewValue);
	long GetMarkerBackgroundColorIndex();
	void SetMarkerBackgroundColorIndex(long nNewValue);
	long GetMarkerForegroundColor();
	void SetMarkerForegroundColor(long nNewValue);
	long GetMarkerForegroundColorIndex();
	void SetMarkerForegroundColorIndex(long nNewValue);
	long GetMarkerSize();
	void SetMarkerSize(long nNewValue);
	long GetMarkerStyle();
	void SetMarkerStyle(long nNewValue);
	long GetPictureType();
	void SetPictureType(long nNewValue);
	long GetPictureUnit();
	void SetPictureUnit(long nNewValue);
	void Select();
	BOOL GetSmooth();
	void SetSmooth(BOOL bNewValue);
	double GetLeft();
	double GetTop();
	double GetWidth();
	double GetHeight();
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Trendlines 

class Trendlines : public COleDispatchDriver
{
public:
	Trendlines() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Trendlines(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Trendlines(const Trendlines& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH Add(long Type, const VARIANT& Order, const VARIANT& Period, const VARIANT& Forward, const VARIANT& Backward, const VARIANT& Intercept, const VARIANT& DisplayEquation, const VARIANT& DisplayRSquared, const VARIANT& Name);
	long GetCount();
	LPDISPATCH Item(const VARIANT& Index);
	LPUNKNOWN _NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Trendline 

class Trendline : public COleDispatchDriver
{
public:
	Trendline() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Trendline(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Trendline(const Trendline& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetBackward();
	void SetBackward(long nNewValue);
	LPDISPATCH GetBorder();
	void ClearFormats();
	LPDISPATCH GetDataLabel();
	void Delete();
	BOOL GetDisplayEquation();
	void SetDisplayEquation(BOOL bNewValue);
	BOOL GetDisplayRSquared();
	void SetDisplayRSquared(BOOL bNewValue);
	long GetForward();
	void SetForward(long nNewValue);
	long GetIndex();
	double GetIntercept();
	void SetIntercept(double newValue);
	BOOL GetInterceptIsAuto();
	void SetInterceptIsAuto(BOOL bNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	BOOL GetNameIsAuto();
	void SetNameIsAuto(BOOL bNewValue);
	long GetOrder();
	void SetOrder(long nNewValue);
	long GetPeriod();
	void SetPeriod(long nNewValue);
	void Select();
	long GetType();
	void SetType(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Corners 

class Corners : public COleDispatchDriver
{
public:
	Corners() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Corners(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Corners(const Corners& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetName();
	void Select();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse SeriesLines 

class SeriesLines : public COleDispatchDriver
{
public:
	SeriesLines() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	SeriesLines(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	SeriesLines(const SeriesLines& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetName();
	void Select();
	LPDISPATCH GetBorder();
	void Delete();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse HiLoLines 

class HiLoLines : public COleDispatchDriver
{
public:
	HiLoLines() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	HiLoLines(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	HiLoLines(const HiLoLines& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetName();
	void Select();
	LPDISPATCH GetBorder();
	void Delete();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Gridlines 

class Gridlines : public COleDispatchDriver
{
public:
	Gridlines() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Gridlines(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Gridlines(const Gridlines& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetName();
	void Select();
	LPDISPATCH GetBorder();
	void Delete();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse DropLines 

class DropLines : public COleDispatchDriver
{
public:
	DropLines() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	DropLines(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	DropLines(const DropLines& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetName();
	void Select();
	LPDISPATCH GetBorder();
	void Delete();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse LeaderLines 

class LeaderLines : public COleDispatchDriver
{
public:
	LeaderLines() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	LeaderLines(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	LeaderLines(const LeaderLines& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBorder();
	void Delete();
	void Select();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse UpBars 

class UpBars : public COleDispatchDriver
{
public:
	UpBars() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	UpBars(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	UpBars(const UpBars& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetName();
	void Select();
	LPDISPATCH GetBorder();
	void Delete();
	LPDISPATCH GetInterior();
	LPDISPATCH GetFill();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse DownBars 

class DownBars : public COleDispatchDriver
{
public:
	DownBars() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	DownBars(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	DownBars(const DownBars& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetName();
	void Select();
	LPDISPATCH GetBorder();
	void Delete();
	LPDISPATCH GetInterior();
	LPDISPATCH GetFill();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Floor 

class Floor : public COleDispatchDriver
{
public:
	Floor() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Floor(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Floor(const Floor& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetName();
	void Select();
	LPDISPATCH GetBorder();
	void ClearFormats();
	LPDISPATCH GetInterior();
	LPDISPATCH GetFill();
	VARIANT GetPictureType();
	void SetPictureType(const VARIANT& newValue);
	void Paste();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Walls 

class Walls : public COleDispatchDriver
{
public:
	Walls() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Walls(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Walls(const Walls& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetName();
	void Select();
	LPDISPATCH GetBorder();
	void ClearFormats();
	LPDISPATCH GetInterior();
	LPDISPATCH GetFill();
	VARIANT GetPictureType();
	void SetPictureType(const VARIANT& newValue);
	void Paste();
	VARIANT GetPictureUnit();
	void SetPictureUnit(const VARIANT& newValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse TickLabels 

class TickLabels : public COleDispatchDriver
{
public:
	TickLabels() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	TickLabels(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	TickLabels(const TickLabels& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void Delete();
	LPDISPATCH GetFont();
	CString GetName();
	CString GetNumberFormat();
	void SetNumberFormat(LPCTSTR lpszNewValue);
	BOOL GetNumberFormatLinked();
	void SetNumberFormatLinked(BOOL bNewValue);
	long GetOrientation();
	void SetOrientation(long nNewValue);
	void Select();
	long GetReadingOrder();
	void SetReadingOrder(long nNewValue);
	VARIANT GetAutoScaleFont();
	void SetAutoScaleFont(const VARIANT& newValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse PlotArea 

class PlotArea : public COleDispatchDriver
{
public:
	PlotArea() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	PlotArea(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	PlotArea(const PlotArea& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetName();
	void Select();
	LPDISPATCH GetBorder();
	void ClearFormats();
	double GetHeight();
	void SetHeight(double newValue);
	LPDISPATCH GetInterior();
	LPDISPATCH GetFill();
	double GetLeft();
	void SetLeft(double newValue);
	double GetTop();
	void SetTop(double newValue);
	double GetWidth();
	void SetWidth(double newValue);
	double GetInsideLeft();
	double GetInsideTop();
	double GetInsideWidth();
	double GetInsideHeight();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ChartArea 

class ChartArea : public COleDispatchDriver
{
public:
	ChartArea() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ChartArea(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ChartArea(const ChartArea& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetName();
	void Select();
	LPDISPATCH GetBorder();
	void Clear();
	void ClearContents();
	void Copy();
	LPDISPATCH GetFont();
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
	void ClearFormats();
	double GetHeight();
	void SetHeight(double newValue);
	LPDISPATCH GetInterior();
	LPDISPATCH GetFill();
	double GetLeft();
	void SetLeft(double newValue);
	double GetTop();
	void SetTop(double newValue);
	double GetWidth();
	void SetWidth(double newValue);
	VARIANT GetAutoScaleFont();
	void SetAutoScaleFont(const VARIANT& newValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Legend 

class Legend : public COleDispatchDriver
{
public:
	Legend() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Legend(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Legend(const Legend& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetName();
	void Select();
	LPDISPATCH GetBorder();
	void Delete();
	LPDISPATCH GetFont();
	LPDISPATCH LegendEntries(const VARIANT& Index);
	long GetPosition();
	void SetPosition(long nNewValue);
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
	void Clear();
	double GetHeight();
	void SetHeight(double newValue);
	LPDISPATCH GetInterior();
	LPDISPATCH GetFill();
	double GetLeft();
	void SetLeft(double newValue);
	double GetTop();
	void SetTop(double newValue);
	double GetWidth();
	void SetWidth(double newValue);
	VARIANT GetAutoScaleFont();
	void SetAutoScaleFont(const VARIANT& newValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ErrorBars 

class ErrorBars : public COleDispatchDriver
{
public:
	ErrorBars() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ErrorBars(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ErrorBars(const ErrorBars& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetName();
	void Select();
	LPDISPATCH GetBorder();
	void Delete();
	void ClearFormats();
	long GetEndStyle();
	void SetEndStyle(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse DataTable 

class DataTable : public COleDispatchDriver
{
public:
	DataTable() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	DataTable(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	DataTable(const DataTable& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	BOOL GetShowLegendKey();
	void SetShowLegendKey(BOOL bNewValue);
	BOOL GetHasBorderHorizontal();
	void SetHasBorderHorizontal(BOOL bNewValue);
	BOOL GetHasBorderVertical();
	void SetHasBorderVertical(BOOL bNewValue);
	BOOL GetHasBorderOutline();
	void SetHasBorderOutline(BOOL bNewValue);
	LPDISPATCH GetBorder();
	LPDISPATCH GetFont();
	void Select();
	void Delete();
	VARIANT GetAutoScaleFont();
	void SetAutoScaleFont(const VARIANT& newValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Phonetic 

class Phonetic : public COleDispatchDriver
{
public:
	Phonetic() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Phonetic(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Phonetic(const Phonetic& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	long GetCharacterType();
	void SetCharacterType(long nNewValue);
	long GetAlignment();
	void SetAlignment(long nNewValue);
	LPDISPATCH GetFont();
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Shape 

class Shape : public COleDispatchDriver
{
public:
	Shape() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Shape(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Shape(const Shape& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void Apply();
	void Delete();
	LPDISPATCH Duplicate();
	void Flip(long FlipCmd);
	void IncrementLeft(float Increment);
	void IncrementRotation(float Increment);
	void IncrementTop(float Increment);
	void PickUp();
	void RerouteConnections();
	void ScaleHeight(float Factor, long RelativeToOriginalSize, const VARIANT& Scale);
	void ScaleWidth(float Factor, long RelativeToOriginalSize, const VARIANT& Scale);
	void Select(const VARIANT& Replace);
	void SetShapesDefaultProperties();
	LPDISPATCH Ungroup();
	void ZOrder(long ZOrderCmd);
	LPDISPATCH GetAdjustments();
	LPDISPATCH GetTextFrame();
	long GetAutoShapeType();
	void SetAutoShapeType(long nNewValue);
	LPDISPATCH GetCallout();
	long GetConnectionSiteCount();
	long GetConnector();
	LPDISPATCH GetConnectorFormat();
	LPDISPATCH GetFill();
	LPDISPATCH GetGroupItems();
	float GetHeight();
	void SetHeight(float newValue);
	long GetHorizontalFlip();
	float GetLeft();
	void SetLeft(float newValue);
	LPDISPATCH GetLine();
	long GetLockAspectRatio();
	void SetLockAspectRatio(long nNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	LPDISPATCH GetNodes();
	float GetRotation();
	void SetRotation(float newValue);
	LPDISPATCH GetPictureFormat();
	LPDISPATCH GetShadow();
	LPDISPATCH GetTextEffect();
	LPDISPATCH GetThreeD();
	float GetTop();
	void SetTop(float newValue);
	long GetType();
	long GetVerticalFlip();
	VARIANT GetVertices();
	long GetVisible();
	void SetVisible(long nNewValue);
	float GetWidth();
	void SetWidth(float newValue);
	long GetZOrderPosition();
	LPDISPATCH GetHyperlink();
	long GetBlackWhiteMode();
	void SetBlackWhiteMode(long nNewValue);
	CString GetOnAction();
	void SetOnAction(LPCTSTR lpszNewValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	LPDISPATCH GetTopLeftCell();
	LPDISPATCH GetBottomRightCell();
	long GetPlacement();
	void SetPlacement(long nNewValue);
	void Copy();
	void Cut();
	void CopyPicture(const VARIANT& Appearance, const VARIANT& Format);
	LPDISPATCH GetControlFormat();
	LPDISPATCH GetLinkFormat();
	LPDISPATCH GetOLEFormat();
	long GetFormControlType();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse Shapes 

class Shapes : public COleDispatchDriver
{
public:
	Shapes() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	Shapes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Shapes(const Shapes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Item(const VARIANT& Index);
	LPDISPATCH _Default(const VARIANT& Index);
	LPUNKNOWN Get_NewEnum();
	LPDISPATCH AddCallout(long Type, float Left, float Top, float Width, float Height);
	LPDISPATCH AddConnector(long Type, float BeginX, float BeginY, float EndX, float EndY);
	LPDISPATCH AddCurve(const VARIANT& SafeArrayOfPoints);
	LPDISPATCH AddLabel(long Orientation, float Left, float Top, float Width, float Height);
	LPDISPATCH AddLine(float BeginX, float BeginY, float EndX, float EndY);
	LPDISPATCH AddPicture(LPCTSTR Filename, long LinkToFile, long SaveWithDocument, float Left, float Top, float Width, float Height);
	LPDISPATCH AddPolyline(const VARIANT& SafeArrayOfPoints);
	LPDISPATCH AddShape(long Type, float Left, float Top, float Width, float Height);
	LPDISPATCH AddTextEffect(long PresetTextEffect, LPCTSTR Text, LPCTSTR FontName, float FontSize, long FontBold, long FontItalic, float Left, float Top);
	LPDISPATCH AddTextbox(long Orientation, float Left, float Top, float Width, float Height);
	LPDISPATCH BuildFreeform(long EditingType, float X1, float Y1);
	LPDISPATCH GetRange(const VARIANT& Index);
	void SelectAll();
	LPDISPATCH AddFormControl(long Type, long Left, long Top, long Width, long Height);
	LPDISPATCH AddOLEObject(const VARIANT& ClassType, const VARIANT& Filename, const VARIANT& Link, const VARIANT& DisplayAsIcon, const VARIANT& IconFileName, const VARIANT& IconIndex, const VARIANT& IconLabel, const VARIANT& Left, 
		const VARIANT& Top, const VARIANT& Width, const VARIANT& Height);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ShapeRange 

class ShapeRange : public COleDispatchDriver
{
public:
	ShapeRange() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ShapeRange(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ShapeRange(const ShapeRange& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Item(const VARIANT& Index);
	LPDISPATCH _Default(const VARIANT& Index);
	LPUNKNOWN Get_NewEnum();
	void Align(long AlignCmd, long RelativeTo);
	void Apply();
	void Delete();
	void Distribute(long DistributeCmd, long RelativeTo);
	LPDISPATCH Duplicate();
	void Flip(long FlipCmd);
	void IncrementLeft(float Increment);
	void IncrementRotation(float Increment);
	void IncrementTop(float Increment);
	LPDISPATCH Group();
	void PickUp();
	void RerouteConnections();
	LPDISPATCH Regroup();
	void ScaleHeight(float Factor, long RelativeToOriginalSize, const VARIANT& Scale);
	void ScaleWidth(float Factor, long RelativeToOriginalSize, const VARIANT& Scale);
	void Select(const VARIANT& Replace);
	void SetShapesDefaultProperties();
	LPDISPATCH Ungroup();
	void ZOrder(long ZOrderCmd);
	LPDISPATCH GetAdjustments();
	LPDISPATCH GetTextFrame();
	long GetAutoShapeType();
	void SetAutoShapeType(long nNewValue);
	LPDISPATCH GetCallout();
	long GetConnectionSiteCount();
	long GetConnector();
	LPDISPATCH GetConnectorFormat();
	LPDISPATCH GetFill();
	LPDISPATCH GetGroupItems();
	float GetHeight();
	void SetHeight(float newValue);
	long GetHorizontalFlip();
	float GetLeft();
	void SetLeft(float newValue);
	LPDISPATCH GetLine();
	long GetLockAspectRatio();
	void SetLockAspectRatio(long nNewValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	LPDISPATCH GetNodes();
	float GetRotation();
	void SetRotation(float newValue);
	LPDISPATCH GetPictureFormat();
	LPDISPATCH GetShadow();
	LPDISPATCH GetTextEffect();
	LPDISPATCH GetThreeD();
	float GetTop();
	void SetTop(float newValue);
	long GetType();
	long GetVerticalFlip();
	VARIANT GetVertices();
	long GetVisible();
	void SetVisible(long nNewValue);
	float GetWidth();
	void SetWidth(float newValue);
	long GetZOrderPosition();
	long GetBlackWhiteMode();
	void SetBlackWhiteMode(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse GroupShapes 

class GroupShapes : public COleDispatchDriver
{
public:
	GroupShapes() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	GroupShapes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	GroupShapes(const GroupShapes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Item(const VARIANT& Index);
	LPDISPATCH _Default(const VARIANT& Index);
	LPUNKNOWN Get_NewEnum();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse TextFrame 

class TextFrame : public COleDispatchDriver
{
public:
	TextFrame() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	TextFrame(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	TextFrame(const TextFrame& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	float GetMarginBottom();
	void SetMarginBottom(float newValue);
	float GetMarginLeft();
	void SetMarginLeft(float newValue);
	float GetMarginRight();
	void SetMarginRight(float newValue);
	float GetMarginTop();
	void SetMarginTop(float newValue);
	long GetOrientation();
	void SetOrientation(long nNewValue);
	LPDISPATCH Characters(const VARIANT& Start, const VARIANT& Length);
	long GetHorizontalAlignment();
	void SetHorizontalAlignment(long nNewValue);
	long GetVerticalAlignment();
	void SetVerticalAlignment(long nNewValue);
	BOOL GetAutoSize();
	void SetAutoSize(BOOL bNewValue);
	long GetReadingOrder();
	void SetReadingOrder(long nNewValue);
	BOOL GetAutoMargins();
	void SetAutoMargins(BOOL bNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ConnectorFormat 

class ConnectorFormat : public COleDispatchDriver
{
public:
	ConnectorFormat() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ConnectorFormat(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ConnectorFormat(const ConnectorFormat& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void BeginConnect(LPDISPATCH ConnectedShape, long ConnectionSite);
	void BeginDisconnect();
	void EndConnect(LPDISPATCH ConnectedShape, long ConnectionSite);
	void EndDisconnect();
	long GetBeginConnected();
	LPDISPATCH GetBeginConnectedShape();
	long GetBeginConnectionSite();
	long GetEndConnected();
	LPDISPATCH GetEndConnectedShape();
	long GetEndConnectionSite();
	long GetType();
	void SetType(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse FreeformBuilder 

class FreeformBuilder : public COleDispatchDriver
{
public:
	FreeformBuilder() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	FreeformBuilder(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	FreeformBuilder(const FreeformBuilder& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void AddNodes(long SegmentType, long EditingType, float X1, float Y1, const VARIANT& X2, const VARIANT& Y2, const VARIANT& X3, const VARIANT& Y3);
	LPDISPATCH ConvertToShape();
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse ControlFormat 

class ControlFormat : public COleDispatchDriver
{
public:
	ControlFormat() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	ControlFormat(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ControlFormat(const ControlFormat& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void AddItem(LPCTSTR Text, const VARIANT& Index);
	void RemoveAllItems();
	void RemoveItem(long Index, const VARIANT& Count);
	long GetDropDownLines();
	void SetDropDownLines(long nNewValue);
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	long GetLargeChange();
	void SetLargeChange(long nNewValue);
	CString GetLinkedCell();
	void SetLinkedCell(LPCTSTR lpszNewValue);
	VARIANT List(const VARIANT& Index);
	long GetListCount();
	void SetListCount(long nNewValue);
	CString GetListFillRange();
	void SetListFillRange(LPCTSTR lpszNewValue);
	long GetListIndex();
	void SetListIndex(long nNewValue);
	BOOL GetLockedText();
	void SetLockedText(BOOL bNewValue);
	long GetMax();
	void SetMax(long nNewValue);
	long GetMin();
	void SetMin(long nNewValue);
	long GetMultiSelect();
	void SetMultiSelect(long nNewValue);
	BOOL GetPrintObject();
	void SetPrintObject(BOOL bNewValue);
	long GetSmallChange();
	void SetSmallChange(long nNewValue);
	long Get_Default();
	void Set_Default(long nNewValue);
	long GetValue();
	void SetValue(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse OLEFormat 

class OLEFormat : public COleDispatchDriver
{
public:
	OLEFormat() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	OLEFormat(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	OLEFormat(const OLEFormat& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void Activate();
	LPDISPATCH GetObject();
	CString GetProgId();
	void Verb(const VARIANT& Verb);
};
/////////////////////////////////////////////////////////////////////////////
// Wrapper-Klasse LinkFormat 

class LinkFormat : public COleDispatchDriver
{
public:
	LinkFormat() {}		// Ruft den Standardkonstruktor für COleDispatchDriver auf
	LinkFormat(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	LinkFormat(const LinkFormat& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attribute
public:

// Operationen
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	BOOL GetAutoUpdate();
	void SetAutoUpdate(BOOL bNewValue);
	BOOL GetLocked();
	void SetLocked(BOOL bNewValue);
	void Update();
};
