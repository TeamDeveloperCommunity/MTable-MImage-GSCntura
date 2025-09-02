/* $Header: /Phoenix/sqlwin/SRC/H/swcc.h 1     11/04/97 5:35p Jlacerda $ */
#ifndef __SWCC__
#define __SWCC__
/*
    swcc.h
    Centura Builder custom control external interface.
    Copyright (c) 1996 Centura Software Corporation,  All Rights Reserved.
    Written by Jim Tierney

    This file defines data types, constants, macros, and function calls
    that are used by Centura Builder and custom control dynamic libraries
    to define class specific attributes of a custom control instance.
*/

/*
    This structure is used to passed custom control defined settings
    between Centura Builder and the custom control library.
*/
typedef struct tagSWCCSETTINGS
{
  DWORD dwLength;
  HGLOBAL hMem;
} SWCCSETTINGS;
typedef SWCCSETTINGS FAR *LPSWCCSETTINGS;


/*
   One or more of these flags, combined with the '|' operator
   are returned to Centura Builder to indicate how Centura Builder should
   behave after calling the dynamic library's edit function.
*/
#define SWCC_NEWSTYLE       0x0001
#define SWCC_NEWSETTINGS    0x0002
#define SWCC_NOCHANGE       0x0000
#define SWCC_PAINTWINDOW    0x0004
#define SWCC_RECREATEWINDOW 0x0008
#define SWCC_ERROR          0xffff

/*
    This is the definition of the custom control edit function that
    Centura Builder looks for in the custom control's dynamic library.
    The purpose of this function is to define a window style and/or
    custom control settings specific to a custom control class.

    This function is called by Centura Builder when a Centura Builder user
    selects an edit command from the customizer.  If Centura Builder can't find
    a function with this name in the dynamic library, the edit command
    will be disabled.

    This edit function should
    display a modal dialog box that allows the user to change the
    custom control window style and/or custom control settings.  When
    control returns to Centura Builder, information defined by the user will
    be stored in the Centura Builder outline.

    Input values:
      hWndDialogParent - This window handle should be the parent window
                         of the modal dialog created by this function.
      lpszClass  - This is the name of the custom control class
                   being edited.
      hWndCustom - This is the window handle of the custom control.
      lpSettings - Indicates the current custom control defined settings.
                   dwLength - Indicates the length in bytes of the
                             currently defined settings.  This element
                             will be zero if no settings have been defined.
                   hMem    - A handle to global memory containing the
                             current settings.  This will always be a valid
                             handle even if dwLength is zero.  This
                             memory block will always contain at least
                             as many bytes as indicated by dwLength.
      lpdwStyle -  Indicates the current window style bits.  These
                   are the bits passed to CreateWindow.  The
                   following style bits are controlled by Centura Builder
                   and not included in dwStyle:
                   WS_CHILD
                   WS_VSCROLL
                   WS_HSCROLL
                   WS_TABSTOP
                   WS_GROUP
                   WS_BORDER
                   WS_VISIBLE

    Output values:
      lpSettings - Indicates the new custom control defined settings.
                   dwLength - Indicates the length in bytes of the
                             new settings.  This element is set to
                             zero if no settings are defined.
                   hMem    - A handle to global memory containing the
                             new settings.  The dynamic library may
                             reallocate this memory block if necessary to
                             increase its size.  Centura Builder will store
                             the data in this memory block as part of this
                             outline when SWCC_NEWSETTINGS is
                             returned to Centura Builder.
      lpdwStyle - Indicates the new window style bits.  These
                  Centura Builder will replace the current window
                  style with this value when
                  SWCC_STYLECHANGE is returned
                  to Centura Builder.  The following style bits
                  will be ignored because they are controlled
                  by Centura Builder:
                  WS_CHILD
                  WS_VSCROLL
                  WS_HSCROLL
                  WS_TABSTOP
                  WS_GROUP
                  WS_BORDER
                  WS_VISIBLED
    Return value:
      The return value is one or more of these flags, combined with
      the '|' operator.  They indicate whether the dynamic library has
      changed the settings and/or style and how Centura Builder should update
      the custom control window.

      SWCC_NEWSTYLE  - The style has changed.
      SWCC_NEWSETTINGS - The settings have changed.
      SWCC_NOCHANGE - Neither the style or settings has changed.
      SWCC_PAINTWINDOW - Centura Builder needs to repaint the custom
                                   control window, if any.
      SWCC_RECREATEWINDOW - Centura Builder needs to recreate the
                                      custom control window, if any.
      SWCC_ERROR - An error occured that prevented the
                            edit function from succeeding (e.g.; out of
                            memory).

*/

#ifdef MACINTOSH
WORD 
SW41CustControlEditor(
#elif WIN32
		      WORD FAR PASCAL SW41CustControlEditor(
#elif WIN16
		      WORD FAR PASCAL _export _loadds SW41CustControlEditor(
#endif
						const HWND hWndDialogParent,
						      const LPSTR lpszClass,
						      const HWND hWndCustom,
						  LPSWCCSETTINGS lpSettings,
							 LPDWORD lpdwStyle);
/*
    IMPORTANT:  This function name must appear in
    the dynamic libraries .DEF file in order for Centura Builder to find it.
*/


/*
    A pointer to this structure is passed in the lParam of the
    WM_CREATE and WM_NCCREATE messages to a custom control.
    Use the macros below to access the elements of this structure.
*/
typedef struct tagSWCCCREATESTRUCT
{
  int nReserved;
  UINT fUserMode:1;		/* User mode or design mode */
  UINT fUnused:15;		/* Unused bits */
  LPSWCCSETTINGS lpSettings;	/* Custom control settings */
} SWCCCREATESTRUCT;
typedef SWCCCREATESTRUCT FAR *LPSWCCCREATESTRUCT;

/*
    The following macros are used to extract information from the lParam
    value passed with WM_CREATE and WM_NCCREATE.  These macros must only
    be used when processing these messages.
*/

/*
    This macro is used to retrieve the custom control settings.
    This macro must only be used while processing the WM_CREATE or
    WM_NCCREATE messages.
    The return value is a pointer to a structure.  This pointer and
    the contents of the structure are only valid whild processing the
    these messages.  The structure contains the following items:

    lpSettings -
                 dwLength - Indicates the length in bytes of the
                           settings.  This element is set to
                           zero if no settings are defined.
                 hMem    - This is zero if no settings are defined.
                           Otherwise it is a handle to global memory
                           containing the settings defined for this
                           custom control.
*/

#define SWCCGETCREATESETTINGS( lparam )   \
   (((LPSWCCCREATESTRUCT)((LPCREATESTRUCT)lparam)->lpCreateParams)->lpSettings)

/*
    This macro is used to retrieve the mode of the control control.
    This macro must only be used while processing the WM_CREATE or
    WM_NCCREATE messages.
    The return value is a BOOLEAN value.  It is TRUE if this control
    is created in user mode.  FALSE if it is created in design mode.
*/

#define SWCCGETCREATEMODE( lparam )   (BOOL) \
   (((LPSWCCCREATESTRUCT)((LPCREATESTRUCT)lparam)->lpCreateParams)->fUserMode)

#endif				/* __SWCC__ */

/*
 * $History: swcc.h $
 * 
 * *****************  Version 1  *****************
 * User: Jlacerda     Date: 11/04/97   Time: 5:35p
 * Created in $/Phoenix/sqlwin/SRC/H
 * 
 * *****************  Version 2  *****************
 * User: Dking        Date: 6/18/96    Time: 7:55p
 * Branched in $/Patriot/sqlwin/SRC/H
 * Stinger project branched for Patriot
 * 
 * *****************  Version 3  *****************
 * User: Mkenrich     Date: 2/03/96    Time: 2:35p
 * Updated in $/WIN32/SQLWIN/SRC/H
 * cosmetic names changes BITTS 53513-53516
 * 
 * *****************  Version 2  *****************
 * User: Tmcparla     Date: 2/06/95    Time: 4:37p
 * Updated in $/SQLWIN/SRC/H
 *
 */
