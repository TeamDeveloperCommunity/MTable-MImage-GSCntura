// ooconst.h

#ifndef _INC_OOCONST
#define _INC_OOCONST

// Konstanten
// Cell Flags
#define CF_All CF_Value | CF_Datetime | CF_String | CF_Annotation | CF_Formula | CF_HardAttr | CF_Styles | CF_Objects | CF_EditAttr | CF_Formatted
// Font Weight
#define FW_Bold 150.0f
#define FW_Normal 100.0f
// Strikeout
#define SO_None 0
#define SO_Single 1
// Underline
#define UL_None 0
#define UL_Single 1
// Border Line Type
#define BL_Top 0
#define BL_Left 1
#define BL_Right 2
#define BL_Bottom 3
#define BL_Horizontal 4
#define BL_Vertical 5
#define BL_TopLeft_BottomRight 6
#define BL_BottomLeft_TopRight 7

// Enums
enum CellFlags
{
	CF_Value = 1,
	CF_Datetime = 2,
	CF_String = 4,
	CF_Annotation = 8,
	CF_Formula = 16,
	CF_HardAttr = 32,
	CF_Styles = 64,
	CF_Objects = 128,
	CF_EditAttr = 256,
	CF_Formatted = 512

};

enum CellHoriJustify
{
	CHJ_Standard,
	CHJ_Left,
	CHJ_Center,
	CHJ_Right,
	CHJ_Block,
	CHJ_Repeat
};

enum CellVertJustify
{
	CVJ_Standard,
	CVJ_Top,
	CVJ_Center,
	CVJ_Bottom
};

enum FontSlant
{
	FS_None,
	FS_Oblique,
	FS_Italic,
	FS_Dontknow,
	FS_ReverseOblique,
	FS_ReverseItalic
};

#endif // _INC_OOCONST