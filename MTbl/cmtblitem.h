// cmtblitem.h

#ifndef _INC_CMTBLITEM
#define _INC_CMTBLITEM

#ifdef _AFXDLL
#include "stdafx.h"
#else
#include <windows.h>
#endif
#include "cmtblrow.h"
#include "ctd.h"

// Items
#define MTBL_ITEM_UNDEF 0
#define MTBL_ITEM_CELL 1
#define MTBL_ITEM_COLUMN 2
#define MTBL_ITEM_LAST_LOCKED_COLUMN 8
#define MTBL_ITEM_COLUMNS 20
#define MTBL_ITEM_COLHDR 3
#define MTBL_ITEM_COLHDRGRP 4
#define MTBL_ITEM_ROW 5
#define MTBL_ITEM_ROWS 50
#define MTBL_ITEM_ROWHDR 6
#define MTBL_ITEM_CORNER 7

// Item parts
#define MTBL_PART_UNDEF 0
#define MTBL_PART_TEXT 1
#define MTBL_PART_IMAGE 2
#define MTBL_PART_IMAGE_EXP 3
#define MTBL_PART_BTN 4
#define MTBL_PART_DDBTN 5
#define MTBL_PART_NODE 6

// MTBLITEMUDV
#pragma pack(1)
typedef struct MTBLITEMUDV
{
#ifndef CTD1x
	BYTE bOffset[CLASS_UDV_OFFSET_SIZE];
#endif
	NUMBER Type;
	HWND WndHandle;
	NUMBER Id;
	NUMBER Part;
}
MTBLITEMUDV, *LPMTBLITEMUDV;
#pragma pack()


// CMTblItem
class CMTblItem
{
public:
	long Type;
	HWND WndHandle;
	LPVOID Id;
	long Part;

	// Kontruktoren
	CMTblItem();

	// Destruktor
	CMTblItem::~CMTblItem();

	// Operatoren
	CMTblItem& operator=(CMTblItem &p);
	BOOL operator==(CMTblItem &p);

	// Funktionen
	void Init();

	void DefineAsCell(HWND hWndCol, LPVOID pRow);
	void DefineAsCell(HWND hWndCol, long lRowNr);
	void DefineAsCellBtn(HWND hWndCol, LPVOID pRow);
	void DefineAsCellBtn(HWND hWndCol, long lRowNr);
	void DefineAsCellDDBtn(HWND hWndCol, LPVOID pRow);
	void DefineAsCellDDBtn(HWND hWndCol, long lRowNr);
	void DefineAsCellNode(HWND hWndCol, LPVOID pRow);
	void DefineAsCellNode(HWND hWndCol, long lRowNr);
	void DefineAsCol(HWND hWndCol);
	void DefineAsCols(HWND hWndTbl);
	void DefineAsColHdr(HWND hWndCol);
	void DefineAsColHdrBtn(HWND hWndCol);
	void DefineAsColHdrGrp(HWND hWndTbl, LPVOID pGrp);
	void DefineAsCorner(HWND hWndTbl);
	void DefineAsLastLockedCol(HWND hWndTbl);
	void DefineAsRow(HWND hWndTbl, LPVOID pRow);
	void DefineAsRow(HWND hWndTbl, long lRowNr);
	void DefineAsRows(HWND hWndTbl);
	void DefineAsRowHdr(HWND hWndTbl, LPVOID pRow);
	void DefineAsRowHdr(HWND hWndTbl, long lRowNr);
	void DefineAsRowHdrNode(HWND hWndTbl, LPVOID pRow);
	void DefineAsRowHdrNode(HWND hWndTbl, long lRowNr);

	BOOL FromUdv(HUDV hUdv);
	BOOL ToUdv(HUDV hUdv);

	long GetRowNr();
	CMTblRow* GetRowPtr();
	HWND GetTableHandle();

	BOOL IsWndHandleCol() { return IsCell() || IsCol() || IsColHdr(); }

	BOOL IsCell() { return Type == MTBL_ITEM_CELL; }
	BOOL IsCol() { return Type == MTBL_ITEM_COLUMN; }
	BOOL IsCols() { return Type == MTBL_ITEM_COLUMNS; }
	BOOL IsColHdr() { return Type == MTBL_ITEM_COLHDR; }
	BOOL IsColHdrGrp() { return Type == MTBL_ITEM_COLHDRGRP; }
	BOOL IsCorner() { return Type == MTBL_ITEM_CORNER; }
	BOOL IsLastLockedCol() { return Type == MTBL_ITEM_LAST_LOCKED_COLUMN; }
	BOOL IsRow() { return Type == MTBL_ITEM_ROW; }
	BOOL IsRows() { return Type == MTBL_ITEM_ROWS; }
	BOOL IsRowHdr() { return Type == MTBL_ITEM_ROWHDR; }

	BOOL IsRealCell() { return IsCell() && WndHandle && Id; }
	BOOL IsRealColHdr() { return IsColHdr() && WndHandle; }
	BOOL IsRealRow() { return IsRow() && Id; }
	BOOL IsRealRowHdr() { return IsRowHdr() && Id; }
};

typedef CMTblItem MTBLITEM;
typedef CMTblItem* LPMTBLITEM;
typedef list<MTBLITEM> ItemList;

#endif // _INC_CMTBLITEM