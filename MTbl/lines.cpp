// lines.cpp

#include "lines.h"

extern CCtd Ctd;

// Konstruktor
CMTblLineDef::CMTblLineDef()
{
	this->Init();
}

CMTblLineDef::CMTblLineDef(CMTblLineDef& ld)
{
	*this = ld;
}

// Destruktor
CMTblLineDef::~CMTblLineDef()
{

}

// Operatoren
CMTblLineDef & CMTblLineDef::operator=(CMTblLineDef &ld)
{
	this->Style = ld.Style;
	this->Color = ld.Color;
	this->Thickness = ld.Thickness;

	return *this;
}

BOOL CMTblLineDef::operator==(CMTblLineDef &ld)
{
	if (this->Style != ld.Style)
		return FALSE;

	if (this->Color != ld.Color)
		return FALSE;

	if (this->Thickness != ld.Thickness)
		return FALSE;

	return TRUE;
}

void CMTblLineDef::operator+=(CMTblLineDef &ld)
{
	if (ld.Style != MTLS_UNDEF)
		Style = ld.Style;

	if (ld.Color != MTBL_COLOR_UNDEF)
		Color = ld.Color;

	if (ld.Thickness != MTBL_LINE_UNDEF_THICKNESS)
		Thickness = ld.Thickness;
}

// Funktionen
BOOL CMTblLineDef::AnyUndef()
{
	return Style != MTLS_UNDEF || Color != MTBL_COLOR_UNDEF || Thickness != MTBL_LINE_UNDEF_THICKNESS;
}

BOOL CMTblLineDef::FromUdv(HUDV hUdv)
{
	// UDV dereferenzieren
	LPMTBLLINEDEFUDV pUdv = (LPMTBLLINEDEFUDV)Ctd.UdvDeref(hUdv);
	if (!pUdv)
		return FALSE;

	// Werte setzen
	this->Style = Ctd.NumToInt(pUdv->Style);
	this->Color = Ctd.NumToULong(pUdv->Color);
	this->Thickness = Ctd.NumToInt(pUdv->Thickness);

	// Fertig
	return TRUE;
}

void CMTblLineDef::Init()
{
	this->Style = MTLS_UNDEF;
	this->Color = MTBL_COLOR_UNDEF;
	this->Thickness = MTBL_LINE_UNDEF_THICKNESS;
}


BOOL CMTblLineDef::operator!=(CMTblLineDef &ld)
{
	return !(*this == ld);
}

void CMTblLineDef::ReplaceUndefinedFrom(CMTblLineDef &ld)
{
	if (this->Style == MTLS_UNDEF && ld.Style != MTLS_UNDEF)
		this->Style = ld.Style;

	if (this->Color == MTBL_COLOR_UNDEF && ld.Color != MTBL_COLOR_UNDEF)
		this->Color = ld.Color;

	if (this->Thickness == MTBL_LINE_UNDEF_THICKNESS && ld.Thickness != MTBL_LINE_UNDEF_THICKNESS)
		this->Thickness = ld.Thickness;
}

BOOL CMTblLineDef::ToUdv(HUDV hUdv)
{
	// UDV dereferenzieren
	LPMTBLLINEDEFUDV pUdv = (LPMTBLLINEDEFUDV)Ctd.UdvDeref(hUdv);
	if (!pUdv)
		return FALSE;

	// Werte setzen
	Ctd.CvtIntToNumber(this->Style, &pUdv->Style);
	Ctd.CvtULongToNumber((ULONG)this->Color, &pUdv->Color);
	Ctd.CvtIntToNumber(this->Thickness, &pUdv->Thickness);

	// Fertig
	return TRUE;
}