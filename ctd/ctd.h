// ctd.h

#ifndef _INC_CTD
#define _INC_CTD

// Pragmas
#pragma warning(disable:4786)

// Includes
#ifdef _AFXDLL
#include "stdafx.h"
#else
#include <windows.h>
#endif
#include <stdio.h>
#include <tchar.h>
#include "tstring.h"

// Definitionen je nach TD-Version
#ifdef CTD1x
#define UDV_PACK 1

#ifndef NUMBERSIZE
	#define NUMBERSIZE				12
#endif // NUMBERSIZE
#ifndef DATETIMESIZE
	#define DATETIMESIZE			12
#endif // DATETIMESIZE

#define CLSNAME_CHECKBOX			_T("Button")
#define CLSNAME_COLUMN				_T("Centura:Ch:Column")
#define CLSNAME_COLUMN_TBW			_T("Centura:Ta:Column")
#define CLSNAME_CONTAINER			_T("Centura:Ch:Container")
#define CLSNAME_CONTAINER_TBW		_T("Centura:Ta:Container")
#define CLSNAME_DROPDOWN			_T(":DropDown")
#define CLSNAME_DROPDOWNBTN			_T(":DropDownBtn")
#define CLSNAME_LISTCLIP			_T("Centura:Ch:ListClip")
#define CLSNAME_LISTCLIP_TBW		_T("Centura:Ta:ListClip")
#endif // CTD1x

#ifdef CTD2x
#define UDV_PACK 1

#ifndef NUMBERSIZE
	#define NUMBERSIZE				24
#endif // NUMBERSIZE
#ifndef DATETIMESIZE
	#define DATETIMESIZE			24
#endif // DATETIMESIZE

#define CLSNAME_CHECKBOX			_T("Button")
#define CLSNAME_COLUMN				_T("Centura:Ch:Column")
#define CLSNAME_COLUMN_TBW			_T("Centura:Ta:Column")
#define CLSNAME_CONTAINER			_T("Centura:Ch:Container")
#define CLSNAME_CONTAINER_TBW		_T("Centura:Ta:Container")
#define CLSNAME_DROPDOWN			_T("Centura:DropDown")
#define CLSNAME_DROPDOWNBTN			_T("Centura:DropDownBtn")
#define CLSNAME_LISTCLIP			_T("Centura:Ch:ListClip")
#define CLSNAME_LISTCLIP_TBW		_T("Centura:Ta:ListClip")
#endif // CTD2x

#ifdef CTD3x
#define UDV_PACK 1

#ifndef NUMBERSIZE
	#define NUMBERSIZE				24
#endif // NUMBERSIZE
#ifndef DATETIMESIZE
	#define DATETIMESIZE			24
#endif // DATETIMESIZE

#define CLSNAME_CHECKBOX			_T("Button")
#define CLSNAME_COLUMN				_T("Gupta:Chil:Column")
#define CLSNAME_COLUMN_TBW			_T("Gupta:Tabl:Column")
#define CLSNAME_CONTAINER			_T("Gupta:Chil:Container")
#define CLSNAME_CONTAINER_TBW		_T("Gupta:Tabl:Container")
#define CLSNAME_DROPDOWN			_T("Gupta:DropDown")
#define CLSNAME_DROPDOWNBTN			_T("Gupta:DropDownBtn")
#define CLSNAME_LISTCLIP			_T("Gupta:Chil:ListClip")
#define CLSNAME_LISTCLIP_TBW		_T("Gupta:Tabl:ListClip")
#endif // CTD3x

#ifdef CTD4x
#define UDV_PACK 1

#ifndef NUMBERSIZE
	#define NUMBERSIZE				24
#endif // NUMBERSIZE
#ifndef DATETIMESIZE
	#define DATETIMESIZE			24
#endif // DATETIMESIZE

#define CLSNAME_CHECKBOX			_T("Button")
#define CLSNAME_COLUMN				_T("Gupta:Chil:Column")
#define CLSNAME_COLUMN_TBW			_T("Gupta:Tabl:Column")
#define CLSNAME_CONTAINER			_T("Gupta:Chil:Container")
#define CLSNAME_CONTAINER_TBW		_T("Gupta:Tabl:Container")
#define CLSNAME_DROPDOWN			_T("Gupta:DropDown")
#define CLSNAME_DROPDOWNBTN			_T("Gupta:DropDownBtn")
#define CLSNAME_LISTCLIP			_T("Gupta:Chil:ListClip")
#define CLSNAME_LISTCLIP_TBW		_T("Gupta:Tabl:ListClip")
#endif // CTD4x

#ifdef CTD5x
#define UDV_PACK 1

#ifndef NUMBERSIZE
	#define NUMBERSIZE				24
#endif // NUMBERSIZE
#ifndef DATETIMESIZE
	#define DATETIMESIZE			24
#endif // DATETIMESIZE

#define CLSNAME_CHECKBOX			_T("Button")
#define CLSNAME_COLUMN				_T("Gupta:ChildTable:Column")
#define CLSNAME_COLUMN_TBW			_T("Gupta:Table:Column")
#define CLSNAME_CONTAINER			_T("Gupta:ChildTable:Container")
#define CLSNAME_CONTAINER_TBW		_T("Gupta:Table:Container")
#define CLSNAME_DROPDOWN			_T("Gupta:DropDown")
#define CLSNAME_DROPDOWNBTN			_T("Gupta:DropDownBtn")
#define CLSNAME_LISTCLIP			_T("Gupta:ChildTable:ListClip")
#define CLSNAME_LISTCLIP_TBW		_T("Gupta:Table:ListClip")

#define TD_SOLID_TBL_LINES

#define TDTHEMES
#define THEME_Default					0
#define THEME_Office2000				1
#define THEME_OfficeXP					2
#define THEME_Office2003				3
#define THEME_Office2003NoThemes		4
#define THEME_Studio2005				5
#define THEME_Studio2008				6
#define THEME_NativeXP					7
#define THEME_Office2007_R1				8
#define THEME_Office2007_R2_LunaBlue	9
#define THEME_Office2007_R2_Obsidian	10
#define THEME_Office2007_R2_Silver		11
#define THEME_Office2007_R3_LunaBlue	12
#define THEME_Office2007_R3_Obsidian	13
#define THEME_Office2007_R3_Silver		14

#endif // CTD5x

#ifdef CTD6x
#define UDV_PACK 1

#ifndef NUMBERSIZE
	#define NUMBERSIZE				24
#endif // NUMBERSIZE
#ifndef DATETIMESIZE
	#define DATETIMESIZE			24
#endif // DATETIMESIZE

#define CLSNAME_CHECKBOX			_T("Button")
#define CLSNAME_COLUMN				_T("Gupta:ChildTable:Column")
#define CLSNAME_COLUMN_TBW			_T("Gupta:Table:Column")
#define CLSNAME_CONTAINER			_T("Gupta:ChildTable:Container")
#define CLSNAME_CONTAINER_TBW		_T("Gupta:Table:Container")
#define CLSNAME_DROPDOWN			_T("Gupta:DropDown")
#define CLSNAME_DROPDOWNBTN			_T("Gupta:DropDownBtn")
#define CLSNAME_LISTCLIP			_T("Gupta:ChildTable:ListClip")
#define CLSNAME_LISTCLIP_TBW		_T("Gupta:Table:ListClip")

#define TD_SOLID_TBL_LINES

#define TDTHEMES
#define TDTHEMES_OFFICE2010
#define THEME_Default					0
#define THEME_Office2000				1
#define THEME_OfficeXP					2
#define THEME_Office2003				3
#define THEME_Office2003NoThemes		4
#define THEME_Studio2005				5
#define THEME_Studio2008				6
#define THEME_NativeXP					7
#define THEME_Office2007_R1				8
#define THEME_Office2007_R2_LunaBlue	9
#define THEME_Office2007_R2_Obsidian	10
#define THEME_Office2007_R2_Silver		11
#define THEME_Office2007_R3_LunaBlue	12
#define THEME_Office2007_R3_Obsidian	13
#define THEME_Office2007_R3_Silver		14
#define THEME_Office2010_R1				15
#define THEME_Office2010_R2_Blue		16
#define THEME_Office2010_R2_Silver		17
#define THEME_Office2010_R2_Black		18

#endif // CTD6x

#ifdef TD63
#define TDTHEMES_OFFICE2013
#define THEME_Office2013				19

#define TD_POPUP_EDIT_ALWAYS_SHOWN
#endif // TD63

#ifdef CTD7x
#define UDV_PACK 1

#ifndef NUMBERSIZE
#define NUMBERSIZE				24
#endif // NUMBERSIZE
#ifndef DATETIMESIZE
#define DATETIMESIZE			24
#endif // DATETIMESIZE

#define CLSNAME_CHECKBOX			_T("Button")
#define CLSNAME_COLUMN				_T("Gupta:ChildTable:Column")
#define CLSNAME_COLUMN_TBW			_T("Gupta:Table:Column")
#define CLSNAME_CONTAINER			_T("Gupta:ChildTable:Container")
#define CLSNAME_CONTAINER_TBW		_T("Gupta:Table:Container")
#define CLSNAME_DROPDOWN			_T("Gupta:DropDown")
#define CLSNAME_DROPDOWNBTN			_T("Gupta:DropDownBtn")
#define CLSNAME_LISTCLIP			_T("Gupta:ChildTable:ListClip")
#define CLSNAME_LISTCLIP_TBW		_T("Gupta:Table:ListClip")

#define TD_SOLID_TBL_LINES

#define TDTHEMES
#define TDTHEMES_OFFICE2010
#define TDTHEMES_OFFICE2013
#define THEME_Default					0
#define THEME_Office2000				1
#define THEME_OfficeXP					2
#define THEME_Office2003				3
#define THEME_Office2003NoThemes		4
#define THEME_Studio2005				5
#define THEME_Studio2008				6
#define THEME_NativeXP					7
#define THEME_Office2007_R1				8
#define THEME_Office2007_R2_LunaBlue	9
#define THEME_Office2007_R2_Obsidian	10
#define THEME_Office2007_R2_Silver		11
#define THEME_Office2007_R3_LunaBlue	12
#define THEME_Office2007_R3_Obsidian	13
#define THEME_Office2007_R3_Silver		14
#define THEME_Office2010_R1				15
#define THEME_Office2010_R2_Blue		16
#define THEME_Office2010_R2_Silver		17
#define THEME_Office2010_R2_Black		18
#define THEME_Office2013				19

#define TD_POPUP_EDIT_ALWAYS_SHOWN

#endif // CTD7x

#ifdef TD72
#define TDTHEMES_OFFICE2016
#define THEME_Office2016_R2_Metro		21
#endif // TD72

#ifdef TD73
#define TDTHEMES_OFFICE2016
#define THEME_Office2016_R2_Metro		21
#endif // TD73

#ifdef TD74
#define TDTHEMES_OFFICE2016
#define THEME_Office2016_R2_Metro		21
#endif // TD74

#ifdef TD75
#define TDTHEMES_OFFICE2016
#define THEME_Office2016_R2_Metro		21
#endif // TD75

// UDV-Größe
#ifdef _WIN64
#define UDVSIZE							24
#else
#define UDVSIZE							12
#endif

// UDV-Offset
#ifdef _WIN64
#define UDV_OFFSET_SIZE					8
#elif !defined CTD1x
#define UDV_OFFSET_SIZE					4
#endif

// Maske Date/Time-Länge
#define DTLENGTHMASK					0x0F

#ifdef _WIN64
#define HITEM_SEGMENT64
#endif

// Outline
#ifndef _DEF_HITEM
#define _DEF_HITEM
DECLARE_HANDLE(HITEM);
#endif // _DEF_HITEM

#ifndef _DEF_HOUTLINE
#define _DEF_HOUTLINE
DECLARE_HANDLE(HOUTLINE);
#endif // _DEF_HOUTLINE

// Tabellen - Konstanten
#define TBL_Flag_SizableCols		1
#define TBL_Flag_MovableCols		2
#define TBL_Flag_ShowVScroll		4
#define TBL_Flag_SelectableCols		0x80
#define TBL_Flag_HScrollByCols		0x100
#define TBL_Flag_GrayedHeaders		0x200
#define TBL_Flag_ShowWaitCursor		0x400
#define TBL_Flag_SuppressRowLines	0x800
#define TBL_Flag_SingleSelection	0x1000
#define TBL_Flag_SuppressLastColLine	0x2000
#define TBL_Flag_EditLeftJustify	0x4000
#define TBL_Flag_HideColumnHeaders	0x10000

#ifdef TBL_Error
#undef TBL_Error
#endif // TBL_Error
#define TBL_Error					0x7FFF1591

#ifdef TBL_MaxRow
#undef TBL_MaxRow
#endif // TBL_MaxRow
#define TBL_MaxRow					0x7FFF1590

#ifdef TBL_MinRow
#undef TBL_MinRow
#endif // TBL_MinRow
#define TBL_MinRow					-2147423632

#define TBL_MinSplitRow				-2147423630
#define TBL_MaxSplitRow				-1
#define TBL_AutoScroll				0
#define TBL_ScrollBottom			2
#define TBL_ScrollTop				1
#define TBL_NoAdjust				0
#define TBL_Adjust					1
#define TBL_RowFetched				1
#define TBL_NoMoreRows				0
#define TBL_RowDeleted				2
#define TBL_SortIncreasing			1
#define TBL_SortDecreasing			0
#define TBL_ScrollTop				1
#define TBL_ScrollBottom			2
#define ROW_InCache					0x0001
#define ROW_Edited					0x0004
#define ROW_New						0x0008
#define ROW_Selected				0x0002
#define ROW_Hidden					0x2000
#define ROW_HideMarks				0x0040
#define ROW_MarkDeleted				0x0020
#define ROW_UnusedFlag1				0x0080
#define ROW_UnusedFlag2				0x0100
#define ROW_UserFlag1				0x0080
#define ROW_UserFlag2				0x0100
#define ROW_UserFlag3				0x0200
#define ROW_UserFlag4				0x0400
#define ROW_UserFlag5				0x0800
#define TBL_RowHdr_MarkEdits		1
#define TBL_RowHdr_Sizable			4
#define TBL_RowHdr_Visible			8
#define TBL_RowHdr_ShareColor		16
#define TBL_RowHdr_AllowRowSizing	0x0020
#define TBL_YOverNormalRows			0x04
#define TBL_YOverSplitBar			0x08
#define TBL_YOverSplitRows			0x02

// Spalten - Konstanten
#define COL_GetID					0
#define COL_GetPos					1
#define COL_CheckBox_IgnoreCase     0x0001
#define COL_CellType_Undefined		0
#define COL_CellType_Standard		1
#define COL_CellType_CheckBox		2
#define COL_CellType_DropDownList	3
#define COL_CellType_PopupEdit		4
#define COL_DropDownList_Sorted		0x0001
#define COL_Visible					0x0001
#define COL_MultilineCell			0x00800000
#define COL_Editable				0x002
#define COL_Selected				4
#define COL_CenterJustify			0x100
#define COL_RightJustify			0x80
#define COL_Invisible				0x0010
#define COL_IndicateOverflow		0x1000
#define COL_Dynamic					0x008
#define COL_DimTitle				0x020

// PIC-Konstanten
#ifndef PIC_FormatBitmap
#define PIC_FormatBitmap			1
#endif // PIC_FormatBitmap

#ifndef PIC_FormatIcon
#define PIC_FormatIcon				2
#endif // PIC_FormatIcon

#ifndef PIC_FormatObject
#define PIC_FormatObject			3
#endif // PIC_FormatObject

#define PIC_ImageTypeNone			0
#define PIC_ImageTypeBMP			1
#define PIC_ImageTypeICON			2
#define PIC_ImageTypeWMF			3
#define PIC_ImageTypeTIFF			4
#define PIC_ImageTypePCX			5
#define PIC_ImageTypeGIF			6
#define PIC_ImageTypeJPEG			7

// SalColorGet/Set - Konstanten
#define COLOR_IndexCellText			100
#define COLOR_IndexTransparent		101
#define COLOR_IndexWindow			5
#define COLOR_IndexWindowText		8

// SalGetDataType - Konstanten
#define DT_Boolean					0x0001
#define DT_DateTime					0x0002
#define DT_Number					0x0003
#define DT_String					0x0005
#define DT_LongString				0x0007

// SalGetMaxDataLength - Konstanten
#define DW_Default					-1

// SalGetType - Kosntanten
#define TYPE_Any					0x7fffffff
#define TYPE_FormWindow				0x0001
#define TYPE_TableWindow			0x0002
#define TYPE_DialogBox				0x0004
#define TYPE_DataField				0x0008
#define TYPE_MultilineText			0x0010
#define TYPE_PushButton				0x0020
#define TYPE_RadioButton			0x0040
#define TYPE_CheckBox				0x0080
#define TYPE_GroupBox				0x0100
#define TYPE_HorizScrollBar			0x0200
#define TYPE_VertScrollBar			0x0400
#define TYPE_BkgdText				0x0800
#define TYPE_TableColumn			0x1000
#define TYPE_ListBox				0x2000
#define TYPE_ComboBox				0x4000
#define TYPE_Line					0x8000
#define TYPE_Frame					0x00010000L
#define TYPE_Pict					0x00020000L
#define TYPE_ChildTable				0x00080000L
#define TYPE_MDIWindow				0x00100000L
#define TYPE_FormToolBar			0x00200000L
#define TYPE_SpinField				0x01000000L
#define TYPE_OptButton				0x02000000L
#define TYPE_CustControl			0x04000000L
#define TYPE_ActiveX				0x08000000L
#define TYPE_AccFrame				0x10000000L
#define TYPE_Other					0x40000000L

// SalFontGet/Set - Konstanten
#define FONT_EnhNormal 0

#ifndef FONT_EnhDefault
#define FONT_EnhDefault 1
#endif // FONT_EnhDefault

#ifndef FONT_EnhItalic
#define FONT_EnhItalic 2
#endif // FONT_EnhItalic

#ifndef FONT_EnhUnderline
#define FONT_EnhUnderline 4
#endif // FONT_EnhUnderline

#ifndef FONT_EnhBold
#define FONT_EnhBold 8
#endif // FONT_EnhBold

#ifndef FONT_EnhStrikeOut
#define FONT_EnhStrikeOut 16
#endif // FONT_EnhStrikeOut

// SalTblObjectsFromPoint - Konstanten
#define TBL_XOverRowHeader			0x0100
#define TBL_YOverColumnHeader		0x01

// Validate-Konstanten
#define VALIDATE_Cancel				0
#define VALIDATE_Ok					1
#define VALIDATE_OkClearFlag		2

// Nachrichten
#define SAM_AnyEdit					0x2004
#define SAM_AppExit					0x2002
#define SAM_AppStartup				0x2001
#define SAM_CacheFull				0x200B
#define SAM_Click					0x2006
#define SAM_Close					0x0010
#define SAM_ContextMenu				0x2048
#define SAM_CountRows				0x2011
#define SAM_Create					0x1001
#define SAM_Destroy					0x1002
#define SAM_DragCanAutoStart		0x2033
#define SAM_DoubleClick				0x2009
#define SAM_EndCellTab				0x2018
#define SAM_FetchRow				0x2007
#define SAM_FieldEdit				0x2005
#define SAM_KillFocus				0x1008
#define SAM_ReportFetchInit			0x200D
#define SAM_ReportFetchNext			0x200E
#define SAM_ReportFinish			0x2010
#define SAM_ReportStart				0x200F
#define SAM_RowHeaderClick			0x2014
#define SAM_RowHeaderDoubleClick	0x2015
#define SAM_RowValidate				0x201D
#define SAM_ScrollBar				0x200A
#define SAM_SetFocus				0x1007
#define SAM_SqlError				0x2003
#define SAM_Timer					0x0113
#define SAM_User					0x4000
#define SAM_Validate				0x200C

// Interne Tabellen-Nachrichten
#define TIM_ClearSelection (WM_USER + 12)
#define TIM_DeleteRow (WM_USER + 13)
#define TIM_FetchRow (WM_USER + 14)
#define TIM_InsertRow (WM_USER + 17)
#define TIM_KillEdit (WM_USER + 18)
#define TIM_KillFocusFrame (WM_USER + 19)
#define TIM_Reset (WM_USER + 25)
#define TIM_QueryVisRange (WM_USER + 26)
#define TIM_Scroll (WM_USER + 27)
#define TIM_SetContext (WM_USER + 28)
#define TIM_SetContextOnly (WM_USER + 29)
#define TIM_SetFocusCell (WM_USER + 30)
#define TIM_SetFocusRow (WM_USER + 31)
#define TIM_SetRowFlags (WM_USER + 32)
#define TIM_SetRange (WM_USER + 33)
#define TIM_GetColumnText (WM_USER + 40)
#define TIM_QueryColumnWidth (WM_USER + 43)
#define TIM_QueryColumnFlags (WM_USER + 49)
#define TIM_GetMetrics ( WM_USER + 54 )
#define TIM_GetColumnTextLength ( WM_USER + 57 )
#define TIM_DefineSplitWindow ( WM_USER + 64 )
#define TIM_SetTableFlags ( WM_USER + 70 )
#define TIM_QueryTableFlags ( WM_USER + 71 )
#define TIM_SetLockedColumns ( WM_USER + 83 )
#define TIM_QueryPixelOrigin ( WM_USER + 87 )
#define TIM_QueryLastRow ( WM_USER + 92 )
#define TIM_DrawRowFocus ( WM_USER + 94 )
#define TIM_DestroyColumns ( WM_USER + 98 )
#define TIM_SwapRows (WM_USER + 108)
#define TIM_EditModeChange (WM_USER + 112)
#define TIM_GetExtEdit (WM_USER + 119)
#define TIM_SetLinesPerRow (WM_USER + 120)
#define TIM_QueryColumnCellType (WM_USER + 123)
#define TIM_PosChanged (WM_USER + 128)
#define TIM_ObjectsFromPoint (WM_USER + 132)
#define TIM_RowSetContext (WM_USER + 7223)

// Interne Spalten-Nachrichten
#define CIM_SetText (WM_USER + 38)
#define CIM_SetPos (WM_USER + 46)
#define CIM_SetWidth (WM_USER + 47)
#define CIM_SetFlags (WM_USER + 50)
#define CIM_SetTextColor (WM_USER + 68)
#define CIM_GetTextColor (WM_USER + 69)
#define CIM_SetEditType (WM_USER + 117)
#define CIM_GetExtEdit (WM_USER + 119)
#define CIM_DefineCellType (WM_USER + 122)
#define CIM_QueryCellType (WM_USER + 123)
#define CIM_QueryCellTypeLines (WM_USER + 125)
#define CIM_DefineCellTypeParamLast (WM_USER + 126)
#define CIM_DefineCellTypeParamMiddle (WM_USER + 129)

// Typedefs Datentypen
typedef HANDLE HARRAY;
typedef ULONG_PTR HSTRING;
typedef DWORD_PTR TEMPLATE;

typedef HSTRING *LPHSTRING;
typedef HWND *LPHWND;

#if defined(_WIN64)
typedef LONGLONG CTDRESID;
typedef LONGLONG TDLONG;		//Verwendung, wenn Parameter vom Team Developer in Strukturen verwendet werden.
#else
typedef DWORD CTDRESID;
typedef LONG TDLONG;
#endif

#ifdef _WIN64
#define CLASS_UDV_OFFSET_SIZE  8  // size of offset pointer
#else
#define CLASS_UDV_OFFSET_SIZE  4  // size of offset pointer
#endif

// Strukturen
#ifndef CBTYPE_H
typedef struct
{
  BYTE numLength;
  BYTE numValue[NUMBERSIZE];
}
NUMBER, *LPNUMBER;

typedef struct
{
  BYTE DATETIME_Length;
  BYTE DATETIME_Buffer[DATETIMESIZE];
}
DATETIME, *LPDATETIME;


typedef struct
{
  BYTE opaque[UDVSIZE];
}
HUDV, *LPHUDV;

#ifdef _WIN64
#define INT_TO_STR  _i64toa_s
#define STR_TO_INT  _atoi64
#else
#define INT_TO_STR  _itoa_s
#define STR_TO_INT  atol
#endif

#endif // CBTYPE_H

// Interne Strukturen
typedef struct
{
	TDLONG lRow;
	BOOL bSet;
} IS_SETROWFLAGS, *LPIS_SETROWFLAGS;

// Typedefs Funktionen
typedef BOOL (WINAPI *SALARRAYDIMCOUNT)(HARRAY, LPLONG);
typedef BOOL (WINAPI *SALARRAYGETLOWERBOUND)(HARRAY, INT, LPLONG);
typedef BOOL (WINAPI *SALARRAYGETUPPERBOUND)(HARRAY, INT, LPLONG);
typedef BOOL (WINAPI *SALARRAYISEMPTY)(HARRAY);
typedef DWORD (WINAPI *SALCOLORGET)(HWND, INT);

typedef INT (WINAPI *SALCOMPILEANDEVALUATE)(LPTSTR, LPINT, LPINT, LPNUMBER, LPHSTRING, LPDATETIME, LPHWND, BOOL, LPTSTR);
#if defined(TDASCII)
typedef INT (WINAPI *SALCOMPILEANDEVALUATEA)(LPSTR, LPINT, LPINT, LPNUMBER, LPHSTRING, LPDATETIME, LPHWND, BOOL, LPSTR);
#endif

typedef HSTRING (WINAPI *SALCONTEXTCURRENT)(void);
typedef INT (WINAPI *SALDATETOSTR)(DATETIME, LPHSTRING);
typedef BOOL (WINAPI *SALDISABLEWINDOW)(HWND);
typedef BOOL (WINAPI *SALDRAGDROPSTART)(HWND);
typedef BOOL (WINAPI *SALDRAGDROPSTOP)(void);
typedef BOOL (WINAPI *SALENABLEWINDOW)(HWND);
typedef BOOL (WINAPI *SALFMTFIELDTOSTR)(HWND, LPHSTRING, BOOL);
typedef BOOL (WINAPI *SALFMTGETINPUTMASK)(HWND, LPHSTRING);
typedef BOOL (WINAPI *SALFMTGETPICTURE)(HWND, LPHSTRING);
typedef INT  (WINAPI *SALFORMUNITSTOPIXELS)(HWND, NUMBER, BOOL);
typedef INT (WINAPI *SALGETDATATYPE)(HWND);
typedef BOOL (WINAPI *SALGETITEMNAME)(HWND, LPHSTRING);
typedef INT (WINAPI *SALGETMAXDATALENGTH)(HWND);
typedef LONG (WINAPI *SALGETTYPE)(HWND);
typedef WORD (WINAPI *SALGETVERSION)();
typedef HSTRING (WINAPI *SALHSTRINGUNREF)(HSTRING);
typedef BOOL (WINAPI *SALINVALIDATEWINDOW)(HWND);
typedef BOOL (WINAPI *SALISWINDOWENABLED)(HWND);
typedef BOOL (WINAPI *SALISWINDOWVISIBLE)(HWND);
typedef INT (WINAPI *SALLISTADD)(HWND, LPTSTR);
typedef BOOL (WINAPI *SALLISTCLEAR)(HWND);
typedef INT (WINAPI *SALLISTQUERYCOUNT)(HWND);
typedef INT (WINAPI *SALLISTQUERYSELECTION)(HWND);
typedef INT (WINAPI *SALLISTQUERYTEXT)(HWND, INT, LPHSTRING);
typedef BOOL (WINAPI *SALLISTSETSELECT)(HWND, INT);
typedef HSTRING (WINAPI *SALNUMBERTOHSTRING) (DWORD);
typedef HWND (WINAPI *SALPARENTWINDOW)(HWND);
typedef long (WINAPI *SALPICGETIMAGE)(HWND, LPHSTRING, LPINT);
typedef long (WINAPI *SALPICGETSTRING)(HWND, INT, LPHSTRING);
typedef NUMBER (WINAPI *SALPIXELSTOFORMUNITS)(HWND, int, BOOL);
typedef WORD (WINAPI *SALSENDVALIDATEMSG)();
typedef BOOL (WINAPI *SALSETWINDOWLOC)(HWND, NUMBER, NUMBER);
typedef BOOL (WINAPI *SALSETWINDOWSIZE)(HWND, NUMBER, NUMBER);

typedef BOOL (WINAPI *SALSTATUSSETTEXT)(HWND, LPTSTR);

typedef DWORD(WINAPI *SALFONTGET)(HWND, HSTRING, LPINT, LPDWORD);

#if defined(TDASCII)
typedef BOOL (WINAPI *SALSTATUSSETTEXTA)(HWND, LPSTR);
#endif

typedef BOOL (WINAPI *SALSTRISVALIDNUMBER)(LPTSTR);
#if defined(TDASCII)
typedef BOOL (WINAPI *SALSTRISVALIDNUMBERA)(LPSTR);
#endif

typedef DATETIME (WINAPI *SALSTRTODATE)(LPTSTR);
#if defined(TDASCII)
typedef DATETIME (WINAPI *SALSTRTODATEA)(LPSTR);
#endif

typedef NUMBER (WINAPI *SALSTRTONUMBER)(LPTSTR);
#if defined(TDASCII)
typedef NUMBER (WINAPI *SALSTRTONUMBERA)(LPSTR);
#endif

typedef BOOL (WINAPI *SALUPDATEWINDOW)(HWND);

typedef BOOL (WINAPI *SALWINDOWGETPROPERTY)(HWND, LPTSTR, LPHSTRING);
#if defined(TDASCII)
typedef BOOL (WINAPI *SALWINDOWGETPROPERTYA)(HWND, LPSTR, LPHSTRING);
#endif

typedef BOOL (WINAPI *SALVALIDATESET)(HWND, BOOL, LONG);

typedef BOOL (WINAPI *SALTBLANYROWS)(HWND, WORD, WORD);
typedef BOOL (WINAPI *SALTBLCLEARSELECTION)(HWND);
typedef BOOL (WINAPI *SALTBLDEFINECHECKBOXCOLUMN)(HWND, DWORD, LPTSTR, LPTSTR);
typedef BOOL (WINAPI *SALTBLDEFINEDROPDOWNLISTCOLUMN)(HWND, DWORD, INT);
typedef BOOL (WINAPI *SALTBLDEFINEPOPUPEDITCOLUMN)(HWND, DWORD, INT);
typedef BOOL (WINAPI *SALTBLDEFINEROWHEADER)(HWND, LPTSTR, INT, WORD, HWND);
typedef BOOL (WINAPI *SALTBLDEFINESPLITWINDOW)(HWND, INT, BOOL);
typedef BOOL (WINAPI *SALTBLDELETEROW)(HWND, LONG, WORD);
typedef INT (WINAPI *SALTBLFETCHROW)(HWND, LONG);
typedef BOOL (WINAPI *SALTBLFINDNEXTROW)(HWND, LPLONG, WORD, WORD);
typedef BOOL (WINAPI *SALTBLFINDPREVROW)(HWND, LPLONG, WORD, WORD);
typedef BOOL (WINAPI *SALTBLGETCOLUMNTEXT)(HWND, UINT, LPHSTRING);
typedef int  (WINAPI *SALTBLGETCOLUMNTITLE)(HWND, LPHSTRING, int);
typedef HWND (WINAPI *SALTBLGETCOLUMNWINDOW)(HWND, int, WORD);
typedef LONG (WINAPI *SALTBLINSERTROW)(HWND, LONG);
typedef BOOL (WINAPI *SALTBLKILLEDIT)(HWND);
typedef BOOL (WINAPI *SALTBLKILLFOCUS)(HWND);
typedef BOOL (WINAPI *SALTBLOBJECTSFROMPOINT)(HWND, int, int, LPLONG, LPHWND, LPDWORD);
typedef BOOL (WINAPI *SALTBLQUERYCHECKBOXCOLUMN)(HWND, LPDWORD, LPHSTRING, LPHSTRING);
typedef BOOL (WINAPI *SALTBLQUERYCOLUMNCELLTYPE)(HWND, LPINT);
typedef BOOL (WINAPI *SALTBLQUERYCOLUMNFLAGS)(HWND, DWORD);
typedef int (WINAPI *SALTBLQUERYCOLUMNID)(HWND);
typedef int (WINAPI *SALTBLQUERYCOLUMNPOS)(HWND);
typedef BOOL (WINAPI *SALTBLQUERYCOLUMNWIDTH)(HWND, LPNUMBER);
typedef long (WINAPI *SALTBLQUERYCONTEXT)(HWND);
typedef BOOL (WINAPI *SALTBLQUERYDROPDOWNLISTCOLUMN)(HWND, LPDWORD, LPINT);
typedef BOOL (WINAPI *SALQUERYFIELDEDIT)(HWND);
typedef BOOL (WINAPI *SALTBLQUERYFOCUS)(HWND, LPLONG, LPHWND);
typedef BOOL (WINAPI *SALTBLQUERYLINESPERROW)(HWND, LPINT);
typedef INT  (WINAPI *SALTBLQUERYLOCKEDCOLUMNS)(HWND);
typedef BOOL (WINAPI *SALTBLQUERYPOPUPEDITCOLUMN)(HWND, LPDWORD, LPINT);
typedef BOOL (WINAPI *SALTBLQUERYROWFLAGS)(HWND, LONG, WORD);
typedef BOOL (WINAPI *SALTBLQUERYROWHEADER)(HWND, LPHSTRING, INT, LPINT, LPWORD, LPHWND);
typedef BOOL (WINAPI *SALTBLQUERYSCROLL)(HWND, LPLONG, LPLONG, LPLONG);
typedef BOOL (WINAPI *SALTBLQUERYSPLITWINDOW)(HWND, LPINT, LPBOOL);
typedef BOOL (WINAPI *SALTBLQUERYTABLEFLAGS)(HWND, WORD);
typedef BOOL (WINAPI *SALTBLQUERYVISIBLERANGE)(HWND, LPLONG, LPLONG);
typedef BOOL (WINAPI *SALTBLSCROLL)(HWND, LONG, HWND, WORD);
typedef BOOL (WINAPI *SALTBLSETCOLUMNFLAGS)(HWND, DWORD, BOOL);

typedef BOOL (WINAPI *SALTBLSETCOLUMNTEXT)(HWND, UINT, LPTSTR);
#if defined(TDASCII)
typedef BOOL (WINAPI *SALTBLSETCOLUMNTEXTA)(HWND, UINT, LPSTR);
#endif

typedef BOOL (WINAPI *SALTBLSETCOLUMNTITLE)(HWND, LPTSTR);
#if defined(TDASCII)
typedef BOOL (WINAPI *SALTBLSETCOLUMNTITLEA)(HWND, LPSTR);
#endif

typedef BOOL (WINAPI *SALTBLSETCOLUMNWIDTH)(HWND, NUMBER);
typedef BOOL (WINAPI *SALTBLSETCONTEXT)(HWND, long);
typedef BOOL (WINAPI *SALTBLSETFOCUSCELL)(HWND, long, HWND, int, int);
typedef BOOL (WINAPI *SALTBLSETFOCUSROW)(HWND, long);
typedef BOOL (WINAPI *SALTBLSETLINESPERROW)(HWND, int);
typedef BOOL (WINAPI *SALTBLSETROWFLAGS)(HWND, LONG, WORD, BOOL);
typedef BOOL (WINAPI *SALTBLSETTABLEFLAGS)(HWND, WORD, BOOL);

#ifdef TDTHEMES
typedef int (WINAPI *SALTHEMEGET)(void);
#endif

typedef BOOL (WINAPI *SWINCVTDOUBLETONUMBER)(double, LPNUMBER);
typedef BOOL (WINAPI *SWINCVTINTTONUMBER)(int, LPNUMBER);
typedef BOOL (WINAPI *SWINCVTLONGTONUMBER)(long, LPNUMBER);
typedef BOOL (WINAPI *SWINCVTLONGLONGTONUMBER)(LONGLONG, LPNUMBER);
typedef BOOL (WINAPI *SWINCVTWORDTONUMBER)(WORD, LPNUMBER);
typedef BOOL (WINAPI *SWINCVTNUMBERTODOUBLE)(LPNUMBER, double FAR *);
typedef BOOL (WINAPI *SWINCVTNUMBERTOINT)(LPNUMBER, LPINT);
typedef BOOL (WINAPI *SWINCVTNUMBERTOLONG)(LPNUMBER, LPLONG);
#ifdef _WIN64
typedef BOOL(WINAPI *SWINCVTNUMBERTOLONGLONG)(LPNUMBER, PLONGLONG);
#endif
typedef BOOL (WINAPI *SWINCVTNUMBERTOULONG)(LPNUMBER, LPDWORD);
#ifdef _WIN64
typedef BOOL(WINAPI *SWINCVTNUMBERTOULONGLONG)(LPNUMBER, PULONGLONG);
#endif
typedef BOOL (WINAPI *SWINCVTNUMBERTOWORD)(LPNUMBER, LPWORD);
typedef BOOL (WINAPI *SWINCVTULONGTONUMBER)(ULONG, LPNUMBER);
typedef BOOL (WINAPI *SWININITLPHSTRINGPARAM)(LPHSTRING, LONG);
typedef BOOL (__cdecl *SWINMDARRAYGETHANDLE)(HARRAY, LPHANDLE, long,...);
typedef BOOL (__cdecl *SWINMDARRAYGETHSTRING)(HARRAY, LPHSTRING, LONG,...);
typedef BOOL (__cdecl *SWINMDARRAYGETNUMBER)(HARRAY, LPNUMBER, LONG,...);
typedef BOOL (__cdecl *SWINMDARRAYGETUDV)(HARRAY, LPHUDV, long,...);
typedef BOOL (__cdecl *SWINMDARRAYPUTHANDLE)(HARRAY, HANDLE, LONG,...);
typedef BOOL (__cdecl *SWINMDARRAYPUTHSTRING)(HARRAY, HSTRING, LONG,...);
typedef BOOL (__cdecl *SWINMDARRAYPUTNUMBER)(HARRAY, LPNUMBER, LONG,...);

typedef LPTSTR (WINAPI *SWINSTRINGGETBUFFER)(HSTRING, LPLONG);
#if defined(TDASCII)
typedef LPSTR (WINAPI *SWINSTRINGGETBUFFERA)(HSTRING, LPLONG);
#endif

typedef void(WINAPI *SWINHSTRINGUNLOCK)(HSTRING);

typedef LPSTR (WINAPI *SWINUDVDEREF)(HUDV);

typedef long (WINAPI *SALLOGLINE)(LPTSTR, LPTSTR );
#if defined(TDASCII)
typedef long (WINAPI *SALLOGLINEA)(LPSTR, LPSTR );
#endif

typedef CTDRESID (WINAPI *SALRESID)(LPCTSTR);
#if defined(TDASCII)
typedef DWORD (WINAPI *SALRESIDA)(LPCSTR);
#endif

typedef BOOL (WINAPI *SALRESLOAD)(CTDRESID, LPHSTRING);
typedef BOOL(WINAPI *SALWAITCURSOR)(BOOL);

// Klassen
class CCtd
{
private:
	// Variablen
	HINSTANCE m_hInst;
	BOOL m_Initialized;

	SALPICGETIMAGE SalPicGetImage;
public:
	// Variablen	
	TCHAR m_szDLL[MAX_PATH];

	// Funktionsvariablen
	SALARRAYDIMCOUNT ArrayDimCount;
	SALARRAYGETLOWERBOUND ArrayGetLowerBound;
	SALARRAYGETUPPERBOUND ArrayGetUpperBound;
	SALARRAYISEMPTY ArrayIsEmpty;
	SALCOLORGET ColorGet;
	SALCOMPILEANDEVALUATE CompileAndEvaluate;
#if defined(TDASCII)
	SALCOMPILEANDEVALUATEA CompileAndEvaluateA;
#endif
	SALCONTEXTCURRENT ContextCurrent;
	SALDATETOSTR DateToStr;
	SALDISABLEWINDOW DisableWindow;
	SALDRAGDROPSTART DragDropStart;
	SALDRAGDROPSTOP DragDropStop;
	SALENABLEWINDOW EnableWindow;
	SALFMTFIELDTOSTR FmtFieldToStr;
	SALFMTGETPICTURE FmtGetPicture;
	SALFMTGETINPUTMASK FmtGetInputMask;
	SALFORMUNITSTOPIXELS FormUnitsToPixels;
	SALGETDATATYPE GetDataType;
	SALGETITEMNAME GetItemName;
	SALGETMAXDATALENGTH GetMaxDataLength;
	SALGETTYPE GetType;
	SALGETVERSION GetVersion;
	SALHSTRINGUNREF HStringUnRef;
	SALINVALIDATEWINDOW InvalidateWindow;
	SALISWINDOWENABLED IsWindowEnabled;
	SALISWINDOWVISIBLE IsWindowVisible;
	SALLISTADD ListAdd;
	SALLISTCLEAR ListClear;
	SALLISTQUERYCOUNT ListQueryCount;
	SALLISTQUERYSELECTION ListQuerySelection;
	SALLISTQUERYTEXT ListQueryText;
	SALLISTSETSELECT ListSetSelect;
	SALNUMBERTOHSTRING NumberToHString;
	SALPARENTWINDOW ParentWindow;
	SALPICGETSTRING PicGetString;
	SALPIXELSTOFORMUNITS PixelsToFormUnits;	
	SALQUERYFIELDEDIT QueryFieldEdit;
	SALSENDVALIDATEMSG SendValidateMsg;
	SALSETWINDOWSIZE SetWindowSize;
	SALSETWINDOWLOC SetWindowLoc;
	SALSTATUSSETTEXT StatusSetText;
	SALFONTGET FontGet;

#ifdef TDTHEMES
	SALTHEMEGET ThemeGet;
#endif
	SALVALIDATESET ValidateSet;
#if defined(TDASCII)
	SALSTATUSSETTEXTA StatusSetTextA;
#endif
	SALSTRISVALIDNUMBER StrIsValidNumber;
#if defined(TDASCII)
	SALSTRISVALIDNUMBERA StrIsValidNumberA;
#endif
	SALSTRTODATE StrToDate;
#if defined(TDASCII)
	SALSTRTODATEA StrToDateA;
#endif
	SALSTRTONUMBER StrToNumber;
#if defined(TDASCII)
	SALSTRTONUMBERA StrToNumberA;
#endif
	SALUPDATEWINDOW UpdateWindow;
	SALWINDOWGETPROPERTY WindowGetProperty;
#if defined(TDASCII)
	SALWINDOWGETPROPERTYA WindowGetPropertyA;
#endif
	SALTBLANYROWS TblAnyRows;
	SALTBLCLEARSELECTION TblClearSelection;
	SALTBLDEFINECHECKBOXCOLUMN TblDefineCheckBoxColumn;
	SALTBLDEFINEDROPDOWNLISTCOLUMN TblDefineDropDownListColumn;
	SALTBLDEFINEPOPUPEDITCOLUMN TblDefinePopupEditColumn;
	SALTBLDEFINEROWHEADER TblDefineRowHeader;
	SALTBLDEFINESPLITWINDOW TblDefineSplitWindow;
	SALTBLDELETEROW TblDeleteRow;
	SALTBLFETCHROW TblFetchRow;
	SALTBLFINDNEXTROW TblFindNextRow;
	SALTBLFINDPREVROW TblFindPrevRow;
	SALTBLGETCOLUMNTEXT TblGetColumnText;
	SALTBLGETCOLUMNTITLE TblGetColumnTitle;
	SALTBLGETCOLUMNWINDOW TblGetColumnWindow;
	SALTBLINSERTROW TblInsertRow;
	SALTBLKILLEDIT TblKillEdit;
	SALTBLKILLFOCUS TblKillFocus;
	SALTBLOBJECTSFROMPOINT TblObjectsFromPoint;
	SALTBLQUERYCHECKBOXCOLUMN TblQueryCheckBoxColumn;
	SALTBLQUERYCOLUMNCELLTYPE TblQueryColumnCellType;
	SALTBLQUERYCOLUMNFLAGS TblQueryColumnFlags;
	SALTBLQUERYCOLUMNID TblQueryColumnID;
	SALTBLQUERYCOLUMNPOS TblQueryColumnPos;
	SALTBLQUERYCOLUMNWIDTH TblQueryColumnWidth;
	SALTBLQUERYCONTEXT TblQueryContext;
	SALTBLQUERYDROPDOWNLISTCOLUMN TblQueryDropDownListColumn;
	SALTBLQUERYFOCUS TblQueryFocus;
	SALTBLQUERYLINESPERROW TblQueryLinesPerRow;
	SALTBLQUERYLOCKEDCOLUMNS TblQueryLockedColumns;
	SALTBLQUERYPOPUPEDITCOLUMN TblQueryPopupEditColumn;
	SALTBLQUERYROWFLAGS TblQueryRowFlags;
	SALTBLQUERYROWHEADER TblQueryRowHeader;
	SALTBLQUERYSCROLL TblQueryScroll;
	SALTBLQUERYSPLITWINDOW TblQuerySplitWindow;
	SALTBLQUERYTABLEFLAGS TblQueryTableFlags;
	SALTBLQUERYVISIBLERANGE TblQueryVisibleRange;
	SALTBLSCROLL TblScroll;
	SALTBLSETCOLUMNFLAGS TblSetColumnFlags;
	SALTBLSETCOLUMNWIDTH TblSetColumnWidth;
	SALTBLSETCOLUMNTEXT TblSetColumnText;
#if defined(TDASCII)
	SALTBLSETCOLUMNTEXTA TblSetColumnTextA;
#endif
	SALTBLSETCOLUMNTITLE TblSetColumnTitle;
#if defined(TDASCII)
	SALTBLSETCOLUMNTITLEA TblSetColumnTitleA;
#endif
	SALTBLSETCONTEXT TblSetContext;
	SALTBLSETFOCUSCELL TblSetFocusCell;
	SALTBLSETFOCUSROW TblSetFocusRow;
	SALTBLSETLINESPERROW TblSetLinesPerRow;
	SALTBLSETROWFLAGS TblSetRowFlags;
	SALTBLSETTABLEFLAGS TblSetTableFlags;

	SWINCVTDOUBLETONUMBER CvtDoubleToNumber;
	SWINCVTINTTONUMBER CvtIntToNumber;
	SWINCVTLONGTONUMBER CvtLongToNumber;
	SWINCVTLONGLONGTONUMBER CvtLongLongToNumber;
	SWINCVTWORDTONUMBER CvtWordToNumber;
	SWINCVTNUMBERTODOUBLE CvtNumberToDouble;
	SWINCVTNUMBERTOINT CvtNumberToInt;
	SWINCVTNUMBERTOLONG CvtNumberToLong;
#ifdef _WIN64
	SWINCVTNUMBERTOLONGLONG CvtNumberToLongLong;
#endif
	SWINCVTNUMBERTOULONG CvtNumberToULong;
#ifdef _WIN64
	SWINCVTNUMBERTOULONGLONG CvtNumberToULongLong;
#endif
	SWINCVTNUMBERTOWORD CvtNumberToWord;
	SWINCVTULONGTONUMBER CvtULongToNumber;
	SWINHSTRINGUNLOCK HStringUnlock;
	SWININITLPHSTRINGPARAM InitLPHSTRINGParam;
	SWINMDARRAYGETHANDLE MDArrayGetHandle;
	SWINMDARRAYGETHSTRING MDArrayGetHString;
	SWINMDARRAYGETNUMBER MDArrayGetNumber;
	SWINMDARRAYGETUDV MDArrayGetUdv;
	SWINMDARRAYPUTHANDLE MDArrayPutHandle;
	SWINMDARRAYPUTHSTRING MDArrayPutHString;
	SWINMDARRAYPUTNUMBER MDArrayPutNumber;
	SWINSTRINGGETBUFFER StringGetBuffer;
#if defined(TDASCII)
	SWINSTRINGGETBUFFERA StringGetBufferA;
#endif
	SWINUDVDEREF UdvDeref;

	SALLOGLINE LogLine;
#if defined(TDASCII)
	SALLOGLINEA LogLineA;
#endif
	SALRESID ResId;
#if defined(TDASCII)
	SALRESIDA ResIdA;
#endif
	SALRESLOAD ResLoad;
	SALWAITCURSOR WaitCursor;

	// Konstruktor
	CCtd( void );
	// Destruktor
	~CCtd( void );

	// Funktionen
	long BufLenFromStrLen( long lStrLen ){ return ( lStrLen * sizeof(TCHAR) ) + sizeof(TCHAR);}	
	int HStrToInt( HSTRING hs ){return NumToInt( HStrToNum( hs ) );}
	LPTSTR HStrToLPTSTR( HSTRING hs ){long lLen; return StringGetBuffer( hs, &lLen );}
	NUMBER HStrToNum( HSTRING hs );	
	ULONG HStrToULong( HSTRING hs ){return NumToULong( HStrToNum( hs ) );}
	BOOL HStrToTStr(HSTRING hs, tstring &ts, BOOL bUnref = FALSE);
	void HStrUnlock(HSTRING hs);
	BOOL IsPresent( ) { return m_Initialized; };
	NUMBER NumGetNull( );
	BOOL NumIsEqual( NUMBER &n1, NUMBER &n2 );
	BOOL NumToBool( NUMBER &n );
	double NumToDouble( NUMBER &n );
	int NumToInt( NUMBER &n );
	long NumToLong( NUMBER &n );
#ifdef _WIN64
	LONGLONG NumToLongLong(NUMBER &n);
#endif
	ULONG NumToULong( NUMBER &n );
#ifdef _WIN64
	ULONGLONG NumToULongLong(NUMBER &n);
#endif
	long PicGetImage( HWND hWndPic, LPHSTRING phsStr, LPINT piType );
	long StrLenFromBufLen( long lBufLen ){ return ( lBufLen / sizeof(TCHAR) ) - 1;}
	int TblGetCellType( HWND hWndCol );
	long TblGetFocusRow( HWND hWndTbl );
	BOOL TblIsRow( HWND hWndTbl, long lRow );
	BOOL TblSetContextEx( HWND hWndTbl, long lRow );
	BOOL TblSetContextOnly( HWND hWndTbl, long lRow ) { return SendMessage( hWndTbl, TIM_SetContextOnly, 0, lRow ); };
	long TblSetFlagsRowRange( HWND hWndTbl, long lRowFrom, long lRowTo, WORD wFlagsOn, WORD wFlagsOff, WORD wFlagsSet, BOOL bSet );
};

#endif // _INC_CTD