// cmtblfont.cpp

// Includes
#include "cmtblfont.h"

// Konstanten initialisieren
const tstring CMTblFont::UNDEF_NAME = _T("");
const USHORT CMTblFont::UNDEF_SIZE = USHRT_MAX;
const BYTE CMTblFont::UNDEF_ENH = UCHAR_MAX;

// Standard-Konstruktor
CMTblFont::CMTblFont( ) : m_sName(UNDEF_NAME),m_nSize(UNDEF_SIZE), m_nEnh(UNDEF_ENH)
{
	m_Bits.NameUndef = m_Bits.SizeUndef = m_Bits.EnhUndef = 1;
}

// Destruktor
CMTblFont::~CMTblFont( )
{

}

// Operator =
CMTblFont & CMTblFont::operator =( CMTblFont & f )
{
	if ( &f != this )
	{
		m_sName = f.m_sName;
		m_Bits.NameUndef = f.m_Bits.NameUndef;
		m_nSize = f.m_nSize;
		m_Bits.SizeUndef = f.m_Bits.SizeUndef;
		m_nEnh = f.m_nEnh;
		m_Bits.EnhUndef = f.m_Bits.EnhUndef;
	}

	return *this;
}

// Operator ==
BOOL CMTblFont::operator ==( CMTblFont & f )
{
	return m_sName == f.m_sName &&
		   m_nSize == f.m_nSize &&
		   m_nEnh == f.m_nEnh;
}

// Operator !=
BOOL CMTblFont::operator !=( CMTblFont & f )
{
	return !operator==( f );
}


// AddEnh
// Fügt einen Stil hinzu
BOOL CMTblFont::AddEnh( long lEnh )
{
	if ( m_nEnh == UNDEF_ENH )
	{
		m_nEnh = lEnh;
		m_Bits.EnhUndef = 0;
	}
	else
		m_nEnh |= lEnh;
	return TRUE;
}

// Create
// Erzeugt den Font

HFONT CMTblFont::Create( HDC hDC )
{
	if ( !hDC ) return NULL;

	LOGFONT lf;

	// Logfont-Struktur initialisieren
	lf.lfHeight = 0;
	lf.lfWidth = 0;
	lf.lfEscapement = 0;
	lf.lfOrientation = 0;
	lf.lfWeight = FW_NORMAL;
	lf.lfItalic = FALSE;
	lf.lfUnderline = FALSE;
	lf.lfStrikeOut = FALSE;
	lf.lfCharSet = DEFAULT_CHARSET;
	lf.lfOutPrecision = OUT_DEFAULT_PRECIS;
	lf.lfClipPrecision = CLIP_DEFAULT_PRECIS;
	lf.lfQuality = DEFAULT_QUALITY;
	lf.lfPitchAndFamily = DEFAULT_PITCH | FF_DONTCARE;
	lf.lfFaceName[0] = 0;

	// Richtiges Charset ermitteln
	/*LCID lcid = GetThreadLocale( );
	HKL hkl = GetKeyboardLayout( 0 );
	WORD wLangID = LOWORD( hkl );
	if ( wLangID != LANGIDFROMLCID( lcid ) )
		lcid = MAKELCID( wLangID, SORTIDFROMLCID( lcid ) );

	int iLen = GetLocaleInfo( lcid, LOCALE_IDEFAULTANSICODEPAGE, NULL, 0 );
	if ( iLen > 0 )
	{
		LPTSTR lpsCodePage = new TCHAR[iLen];
		if ( GetLocaleInfo( lcid, LOCALE_IDEFAULTANSICODEPAGE, lpsCodePage, iLen ) )
		{
			DWORD dwCodePage = _ttol( lpsCodePage );
			CHARSETINFO ci;
			if ( TranslateCharsetInfo( (DWORD*)dwCodePage, &ci, TCI_SRCCODEPAGE ) )
				lf.lfCharSet = ci.ciCharset;
		}
		delete []lpsCodePage;
	}*/

	// Werte setzen
	if ( !m_Bits.NameUndef )
	{
		LPCTSTR lpsName = m_sName.c_str( );
		if ( lpsName )
			lstrcpy( lf.lfFaceName, lpsName );
	}

	if ( !m_Bits.SizeUndef )
		lf.lfHeight = -MulDiv( m_nSize, GetDeviceCaps( hDC, LOGPIXELSY ), 72 );

	if ( !m_Bits.EnhUndef )
	{
		if ( m_nEnh & FONT_EnhBold )
			lf.lfWeight = FW_BOLD;
		if ( m_nEnh & FONT_EnhItalic )
			lf.lfItalic = TRUE;
		if ( m_nEnh & FONT_EnhUnderline )
			lf.lfUnderline = TRUE;
		if ( m_nEnh & FONT_EnhStrikeOut )
			lf.lfStrikeOut = TRUE;
	}

	return CreateFontIndirect( &lf );
}

// Get
// Liefert den Font
BOOL CMTblFont::Get( LPHSTRING phsName, LPLONG plSize, LPLONG plEnh )
{
	BOOL bOk = TRUE;

	long lLen = Ctd.BufLenFromStrLen( m_sName.size( ) );
	if ( !Ctd.InitLPHSTRINGParam( phsName, lLen ) )
		return FALSE;
	LPTSTR lpsName = Ctd.StringGetBuffer( *phsName, &lLen );
	if (lpsName)
		lstrcpy(lpsName, m_sName.c_str());
	else
		bOk = FALSE;
	Ctd.HStrUnlock(*phsName);
	
	if (bOk)
	{
		*plSize = m_nSize;
		*plEnh = m_nEnh;
	}	

	return bOk;
}

// Get
// Liefert den Font
BOOL CMTblFont::Get( tstring & sName, long & lSize, long & lEnh )
{
	sName = m_sName;
	lSize = m_nSize;
	lEnh = m_nEnh;
	return TRUE;
}

// Set
// Setzt den Font
BOOL CMTblFont::Set( LPTSTR lpsName, long lSize, long lEnh )
{
	if ( lpsName && lstrlen( lpsName ) > LF_FACESIZE - 1 ) return FALSE;

	m_sName = ( lpsName ? lpsName : UNDEF_NAME );
	m_Bits.NameUndef = ( m_sName == UNDEF_NAME );

	m_nSize = lSize;
	m_Bits.SizeUndef = ( m_nSize == UNDEF_SIZE );

	m_nEnh = lEnh;
	m_Bits.EnhUndef = ( m_nEnh == UNDEF_ENH );

	return TRUE;
}

// Set
// Setzt den Font anhand eines Font-Handles
BOOL CMTblFont::Set( HDC hDC, HFONT hFont )
{
	if ( !( hDC ) )
		return FALSE;

	LOGFONT lf;
	if ( !hFont )
		hFont = (HFONT)GetStockObject( SYSTEM_FONT );
	if ( !GetObject( hFont, sizeof( lf ), &lf ) ) return FALSE;

	m_sName = lf.lfFaceName;
	m_Bits.NameUndef = ( m_sName == UNDEF_NAME );

	m_nSize = abs( MulDiv( lf.lfHeight, 72, GetDeviceCaps( hDC, LOGPIXELSY ) ) );
	m_Bits.SizeUndef = ( m_nSize == UNDEF_SIZE );

	m_nEnh = 0;
	m_nEnh |= ( lf.lfWeight >= FW_BOLD ? FONT_EnhBold : 0 );
	m_nEnh |= ( lf.lfItalic ? FONT_EnhItalic : 0 );
	m_nEnh |= ( lf.lfStrikeOut ? FONT_EnhStrikeOut : 0 );
	m_nEnh |= ( lf.lfUnderline ? FONT_EnhUnderline : 0 );
	m_Bits.EnhUndef = ( m_nEnh == UNDEF_ENH );

	return TRUE;
}

// SubstUndef
// Ersetzt die undefinierten Einstellungen
BOOL CMTblFont::SubstUndef( CMTblFont & f )
{
	if ( &f != this )
	{
		if ( m_Bits.NameUndef && !f.m_Bits.NameUndef)
		{
			m_sName = f.m_sName;
			m_Bits.NameUndef = 0;
		}

		if ( m_Bits.SizeUndef && !f.m_Bits.SizeUndef )
		{
			m_nSize = f.m_nSize;
			m_Bits.SizeUndef = 0;
		}

		if ( m_Bits.EnhUndef && !f.m_Bits.EnhUndef )
		{
			m_nEnh = f.m_nEnh;
			m_Bits.EnhUndef = 0;
		}
	}

	return TRUE;
}