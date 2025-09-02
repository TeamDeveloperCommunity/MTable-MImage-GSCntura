// cmtblcoltag.cpp

#include "cmtblcoltag.h"

// Kontruktor
CMTblColTag::CMTblColTag( )
{
	m_dwBackColor = m_dwBackColorTbl = m_dwTextColor = m_dwTextColorTbl = CMTblColor::UNDEF;
	m_lColSpan = m_lRowSpan = m_lWidth = 0;
}

// Destruktor
CMTblColTag::~CMTblColTag( )
{

}

// Member-Funktionen

// GetHTMLText
void CMTblColTag::GetHTMLText( tstring &sText )
{
	if ( sText.empty( ) )
	{
		sText = _T("<p>&#160;</p>");
		return;
	}
	else
	{
		BOOL bDBCSSecByte = FALSE;
		int i;
		tstring sModText = _T("<p>");
		tstring::iterator it, itBegin = sText.begin( ), itEnd = sText.end( );
		for ( it = itBegin; it != itEnd; ++it )
		{
			if ( IsDBCSLeadByte( *it ) )
			{
				bDBCSSecByte = TRUE;
				sModText.append( 1, *it );
			}
			else if ( bDBCSSecByte )
			{
				bDBCSSecByte = FALSE;
				sModText.append( 1, *it );
			}
			else if ( !m_bUseSpecialCharConversion )
			{
				sModText.append( 1, *it );
			}
			else
			{
				switch ( *it )
				{
				case _T(' '):
					if ( it != itBegin && *(it-1) == _T(' ') )
					{
						sModText.append( _T("&#160;") );
					}
					else
						sModText.append( 1, _T(' ') );
					break;

#ifdef UNICODE
				// Schmales Leerzeichen, Zeilenumbruch erlaubt
				case 8201:					
					sModText.append( L"&#8201;" );
					break;

				// Schmales Leerzeichen, Zeilenumbruch nicht erlaubt
				case 8239:					
					sModText.append( L"&#8239;" );
					break;
#endif

				case _T('ä'):
					sModText.append( _T("&#228;") );
					break;
				case _T('Ä'):
					sModText.append( _T("&#196;") );
					break;
				case _T('ö'):
					sModText.append( _T("&#246;") );
					break;
				case _T('Ö'):
					sModText.append( _T("&#214;") );
					break;
				case _T('ü'):
					sModText.append( _T("&#252;") );
					break;
				case _T('Ü'):
					sModText.append( _T("&#220;") );
					break;
				case _T('ß'):
					sModText.append( _T("&#223;") );
					break;
				case _T('<'):
					sModText.append( _T("&#60;") );
					break;
				case _T('>'):
					sModText.append( _T("&#62;") );
					break;
				case _T('&'):
					sModText.append( _T("&#38;") );
					break;
				case _T('"'):
					sModText.append( _T("&#34;") );
					break;

				case _T('\t'):
					for ( i = 0; i < 8; ++i )
						sModText.append( _T("&#160;") );
					break;
				case _T('\r'):
					sModText.append( _T("<br>") );
					break;
				case _T('\n'):
					break;

				default:
					sModText.append( 1, *it );
				}
			}
		}

		sModText.append( _T("</p>") );
		sText = sModText;
	}
}

// GetString
void CMTblColTag::GetString( tstring &s )
{
	TCHAR szBuf[255];

	// Der Anfang
	s = _T("<td");

	// Breite
	if ( m_lWidth > 0 )
	{
		wsprintf( szBuf, _T(" width=\"%i\""), m_lWidth );
		s += szBuf;
	}

	// Höhe
	if ( m_lHeight > 0 )
	{
		wsprintf( szBuf, _T(" height=\"%i\""), m_lHeight );
		s += szBuf;
	}

	// Alignments
	if ( !m_sAlign.empty( ) )
	{
		s += _T(" align=\"");
		s += m_sAlign;
		s += _T("\"");
	}

	if ( !m_sVAlign.empty( ) )
	{
		s += _T(" valign=\"");
		s += m_sVAlign;
		s += _T("\"");
	}

	// Spannings
	if ( m_lColSpan > 0 )
	{
		wsprintf( szBuf, _T(" colspan=\"%i\""), m_lColSpan );
		s += szBuf;
	}

	if ( m_lRowSpan > 0 )
	{
		wsprintf( szBuf, _T(" rowspan=\"%i\""), m_lRowSpan );
		s += szBuf;
	}

	// CSS-Stile
	if ( m_Font != m_FontTbl || m_dwTextColor != m_dwTextColorTbl || m_dwBackColor != m_dwBackColorTbl || m_lTextIndent > 0 )
	{
		s += _T(" style=\"");

		// Text-Einzug
		if ( m_lTextIndent > 0 )
		{
			wsprintf( szBuf, _T("text-indent: %ipx;"), m_lTextIndent );
			s += szBuf;
		}

		// Font
		// Name
		if ( !m_Font.IsNameUndef( ) && m_Font.GetName( ) != m_FontTbl.GetName( ) )
		{
			s += _T("font-family: ");
			s += m_Font.GetName( );
			s += _T(";");
		}

		// Größe
		if ( !m_Font.IsSizeUndef( ) && m_Font.GetSize( ) != m_FontTbl.GetSize( ) )
		{
			wsprintf( szBuf, _T("font-size: %ipt;"), m_Font.GetSize( ) );
			s += szBuf;
		}

		// Enhancements
		if ( !m_Font.IsEnhUndef( ) && m_Font.GetEnh( ) != m_FontTbl.GetEnh( ) )
		{
			if ( m_Font.GetEnh( ) & FONT_EnhItalic )
				s +=  _T("font-style: italic;");

			if ( m_Font.GetEnh( ) & FONT_EnhBold )
				s += _T("font-weight: bold;");

			BOOL bUnderline = m_Font.GetEnh( ) & FONT_EnhUnderline;
			BOOL bStrikeOut = m_Font.GetEnh( ) & FONT_EnhStrikeOut;
			if ( bUnderline || bStrikeOut )
			{
				s += _T("text-decoration:");
				if ( bUnderline )
					s += _T(" underline");
				if ( bStrikeOut )
					s += _T(" line-through");
				s += _T(";");
			}
		}

		// Farben
		// Text
		if ( m_dwTextColor != CMTblColor::UNDEF && m_dwTextColor != m_dwTextColorTbl )
		{
			wsprintf( szBuf, _T("color: #%02X%02X%02X;"), GetRValue( m_dwTextColor ), GetGValue( m_dwTextColor ), GetBValue( m_dwTextColor ) );
			s += szBuf;			
		}

		// Hintergrund
		if ( m_dwBackColor != CMTblColor::UNDEF && m_dwBackColor != m_dwBackColorTbl )
		{
			wsprintf( szBuf, _T("background-color: #%02X%02X%02X;"), GetRValue( m_dwBackColor ), GetGValue( m_dwBackColor ), GetBValue( m_dwBackColor ) );
			s += szBuf;			
		}

		// Abschließendes Anführungszeichen
		s += _T("\"");
	}

	// Abschließende spitze Klammer
	s += _T(">");

	// Text
	tstring sHTMLText = m_sText;
	GetHTMLText( sHTMLText );

	s += sHTMLText;

	// Abschluß
	s+= _T("</td>");

}

void CMTblColTag::Init( )
{
	m_bUseSpecialCharConversion = TRUE;

	m_Font.SetUndef( );
	m_dwBackColor = m_dwTextColor = CMTblColor::UNDEF;
	m_lColSpan = m_lRowSpan = m_lTextIndent = m_lWidth = m_lHeight = 0;
	m_sAlign = m_sText = m_sVAlign = _T("");
}

void CMTblColTag::UseSpecialCharConversion( BOOL bUse )
{
	m_bUseSpecialCharConversion = bUse;
}
