// merge.cpp
#pragma warning( disable: 4786 )

// Incudes
#include "merge.h"
#include "ctd.h"

// ********************
// ****MTBLMERGEROW****
// ********************

void MTBLMERGEROW::CopyTo( MTBLMERGEROW &mr )
{
	mr.iMergedRows = iMergedRows;
	mr.lRowFrom = lRowFrom;
	mr.lRowTo = lRowTo;
	mr.lRowVisPosFrom = lRowVisPosFrom;
	mr.lRowVisPosTo = lRowVisPosTo;
}

void MTBLMERGEROW::Init( )
{
	iMergedRows = 0;

	lRowFrom = lRowTo = TBL_Error;

	lRowVisPosFrom = lRowVisPosTo = -1;
}