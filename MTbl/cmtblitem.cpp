// cmtblitem.cpp
#include "cmtblitem.h"

extern CCtd Ctd;

// Konstruktor
CMTblItem::CMTblItem()
{
	this->Init();
}

// Destruktor
CMTblItem::~CMTblItem()
{

}

// Operatoren
CMTblItem& CMTblItem::operator=(CMTblItem &p)
{
	if (&p != this)
	{
		this->Type = p.Type;
		this->WndHandle = p.WndHandle;
		this->Id = p.Id;
		this->Part = p.Part;
	}

	return *this;
}

BOOL CMTblItem::operator==(CMTblItem &p)
{
	if (this->Type != p.Type) return FALSE;
	if (this->Part != p.Part) return FALSE;
	if (this->WndHandle != p.WndHandle) return FALSE;
	if (this->Id != p.Id) return FALSE;
	return TRUE;
}

// Funktionen
void CMTblItem::DefineAsCell(HWND hWndCol, LPVOID pRow)
{
	this->Type = MTBL_ITEM_CELL;
	this->WndHandle = hWndCol;
	this->Id = pRow;
	this->Part = MTBL_PART_UNDEF;
}

void CMTblItem::DefineAsCell(HWND hWndCol, long lRowNr)
{
	this->Type = MTBL_ITEM_CELL;
	this->WndHandle = hWndCol;
	this->Id = (LPVOID)SendMessage(Ctd.ParentWindow(hWndCol), EM_GETLINE, 0, lRowNr); // MTblGetRowID
	this->Part = MTBL_PART_UNDEF;
}

void CMTblItem::DefineAsCellBtn(HWND hWndCol, LPVOID pRow)
{
	DefineAsCell(hWndCol, pRow);
	this->Part = MTBL_PART_BTN;
}

void CMTblItem::DefineAsCellBtn(HWND hWndCol, long lRowNr)
{
	DefineAsCell(hWndCol, lRowNr);
	this->Part = MTBL_PART_BTN;
}

void CMTblItem::DefineAsCellDDBtn(HWND hWndCol, LPVOID pRow)
{
	DefineAsCell(hWndCol, pRow);
	this->Part = MTBL_PART_DDBTN;
}

void CMTblItem::DefineAsCellDDBtn(HWND hWndCol, long lRowNr)
{
	DefineAsCell(hWndCol, lRowNr);
	this->Part = MTBL_PART_DDBTN;
}

void CMTblItem::DefineAsCellNode(HWND hWndCol, LPVOID pRow)
{
	DefineAsCell(hWndCol, pRow);
	this->Part = MTBL_PART_NODE;
}

void CMTblItem::DefineAsCellNode(HWND hWndCol, long lRowNr)
{
	DefineAsCell(hWndCol, lRowNr);
	this->Part = MTBL_PART_NODE;
}

void CMTblItem::DefineAsCol(HWND hWndCol)
{
	this->Type = MTBL_ITEM_COLUMN;
	this->WndHandle = hWndCol;
	this->Id = NULL;
	this->Part = MTBL_PART_UNDEF;
}

void CMTblItem::DefineAsCols(HWND hWndTbl)
{
	this->Type = MTBL_ITEM_COLUMNS;
	this->WndHandle = hWndTbl;
	this->Id = NULL;
	this->Part = MTBL_PART_UNDEF;
}

void CMTblItem::DefineAsColHdr(HWND hWndCol)
{
	this->Type = MTBL_ITEM_COLHDR;
	this->WndHandle = hWndCol;
	this->Id = NULL;
	this->Part = MTBL_PART_UNDEF;
}

void CMTblItem::DefineAsColHdrBtn(HWND hWndCol)
{
	this->Type = MTBL_ITEM_COLHDR;
	this->WndHandle = hWndCol;
	this->Id = NULL;
	this->Part = MTBL_PART_BTN;
}

void CMTblItem::DefineAsColHdrGrp(HWND hWndTbl, LPVOID pGrp)
{
	this->Type = MTBL_ITEM_COLHDRGRP;
	this->WndHandle = hWndTbl;
	this->Id = pGrp;
	this->Part = MTBL_PART_UNDEF;
}

void CMTblItem::DefineAsCorner(HWND hWndTbl)
{
	this->Type = MTBL_ITEM_CORNER;
	this->WndHandle = hWndTbl;
	this->Id = NULL;
	this->Part = MTBL_PART_UNDEF;
}

void CMTblItem::DefineAsLastLockedCol(HWND hWndTbl)
{
	this->Type = MTBL_ITEM_LAST_LOCKED_COLUMN;
	this->WndHandle = hWndTbl;
	this->Id = NULL;
	this->Part = MTBL_PART_UNDEF;
}

void CMTblItem::DefineAsRow(HWND hWndTbl, LPVOID pRow)
{
	this->Type = MTBL_ITEM_ROW;
	this->WndHandle = hWndTbl;
	this->Id = pRow;
	this->Part = MTBL_PART_UNDEF;
}

void CMTblItem::DefineAsRow(HWND hWndTbl, long lRowNr)
{
	this->Type = MTBL_ITEM_ROW;
	this->WndHandle = hWndTbl;
	this->Id = (LPVOID)SendMessage(hWndTbl, EM_GETLINE, 0, lRowNr); // MTblGetRowID
	this->Part = MTBL_PART_UNDEF;
}

void CMTblItem::DefineAsRows(HWND hWndTbl)
{
	this->Type = MTBL_ITEM_ROWS;
	this->WndHandle = hWndTbl;
	this->Id = NULL;
	this->Part = MTBL_PART_UNDEF;
}

void CMTblItem::DefineAsRowHdr(HWND hWndTbl, LPVOID pRow)
{
	this->Type = MTBL_ITEM_ROWHDR;
	this->WndHandle = hWndTbl;
	this->Id = pRow;
	this->Part = MTBL_PART_UNDEF;
}

void CMTblItem::DefineAsRowHdr(HWND hWndTbl, long lRowNr)
{
	this->Type = MTBL_ITEM_ROWHDR;
	this->WndHandle = hWndTbl;
	this->Id = (LPVOID)SendMessage(hWndTbl, EM_GETLINE, 0, lRowNr); // MTblGetRowID
	this->Part = MTBL_PART_UNDEF;
}

void CMTblItem::DefineAsRowHdrNode(HWND hWndTbl, LPVOID pRow)
{
	DefineAsRowHdr(hWndTbl, pRow);
	this->Part = MTBL_PART_NODE;
}

void CMTblItem::DefineAsRowHdrNode(HWND hWndTbl, long lRowNr)
{
	DefineAsRowHdr(hWndTbl, lRowNr);
	this->Part = MTBL_PART_NODE;
}

BOOL CMTblItem::FromUdv(HUDV hUdv)
{
	// UDV dereferenzieren
	LPMTBLITEMUDV pUdv = (LPMTBLITEMUDV)Ctd.UdvDeref(hUdv);
	if (!pUdv)
		return FALSE;

	// Werte setzen
	this->Type = Ctd.NumToLong(pUdv->Type);
	this->WndHandle = pUdv->WndHandle;
#ifdef _WIN64
	this->Id = (LPVOID)Ctd.NumToLongLong(pUdv->Id);
#else
	this->Id = (LPVOID)Ctd.NumToLong(pUdv->Id);
#endif
	this->Part = Ctd.NumToLong(pUdv->Part);

	// Fertig
	return TRUE;
}

long CMTblItem::GetRowNr()
{
	CMTblRow* pRow = this->GetRowPtr();
	return pRow ? pRow->GetNr() : TBL_Error;
}

CMTblRow* CMTblItem::GetRowPtr()
{
	CMTblRow* pRow = NULL;
	switch (this->Type)
	{
	case MTBL_ITEM_CELL:
	case MTBL_ITEM_ROW:
	case MTBL_ITEM_ROWHDR:
		pRow = (CMTblRow*)this->Id;
		break;
	}

	return pRow;
}


HWND CMTblItem::GetTableHandle()
{
	if (this->IsWndHandleCol())
		return Ctd.ParentWindow(this->WndHandle);
	else
		return this->WndHandle;
}

void CMTblItem::Init()
{
	this->Type = MTBL_ITEM_UNDEF;
	this->WndHandle = NULL;
	this->Id = NULL;
	this->Part = MTBL_PART_UNDEF;
}

BOOL CMTblItem::ToUdv(HUDV hUdv)
{
	// UDV dereferenzieren
	LPMTBLITEMUDV pUdv = (LPMTBLITEMUDV)Ctd.UdvDeref(hUdv);
	if (!pUdv)
		return FALSE;

	// Werte setzen
	Ctd.CvtLongToNumber(this->Type, &pUdv->Type);
	pUdv->WndHandle = this->WndHandle;
#ifdef _WIN64 
	Ctd.CvtLongLongToNumber((LONGLONG)this->Id, &pUdv->Id);
#else
	Ctd.CvtLongToNumber((long)this->Id, &pUdv->Id);
#endif
	Ctd.CvtLongToNumber(this->Part, &pUdv->Part);

	// Fertig
	return TRUE;
}
