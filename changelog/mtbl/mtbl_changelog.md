## V 3.5.0
- General<br>
  Development environment changed from VC 2013 to VC 2022.<br>
  Please note that from this version on the Visual C++ 2022 Redistributable Package is required.

## V 3.4.1
### Fixes
- x64: MTblGetItem always returns FALSE or causes a crash under certain conditions

## V 3.4.0
- General<br>
   Ready for TD 7.5
	 
- Row merging
  - New table flag MTBL_FLAG_NO_USER_ROW_FLAG_REPL<br>
    If set, the row flags ROW_UserFlag1, ROW_UserFlag2, ROW_UserFlag3, ROW_UserFlag4 and ROW_UserFlag5 are not automatically replicated within merged rows.

### Fixes
- Fixed x64 issue (application crashes when M!Table is initialized, since TD 7.4.5)

## V 3.3.0
- General<br>
   Ready for TD 7.4

- Item lines<br>
   Lines can now be defined individually for cells, columns, column headers, column header groups, rows, row headers and the corner.
  - New functions
    - MTblDefineItemLine
    - MTblQueryItemLine
    - MTblQueryEffItemLine

  - New classes
    - fcMTblLineDef (line definition)

  - New constants
    - MTLT_HORZ (line type horizontal)
    - MTLT_VERT (line type vertical)
    - MTBL_LINE_UNDEF_THICKNESS (undefined line thickness)

- Selection darkening<br>
   In addition to inversion and fixed selection colors, another option for visualizing selections.<br>
   The background, text and tree line colors are darkened by a certain amount.
  - New functions:
    - MTblSetSelectionDarkening
    - MTblGetSelectionDarkening

- Deprecated functions<br>
   The following functions will no longer be developed and may not work properly anymore in this or following releases:
  - MTblExportToHTML
  - MTblExportToOOCalc

### Fixes
- Printing: Word wrap is not considered (since version 3.0.0)

## V 3.2.1
- New copy functionality<br>
   When [STRG]+[C] or [STRG]+[INS] is pressed, the selected cells are copied to clipboard.

- New message MTM_Copy<br>
   This message is sent when [STRG]+[C] or [STRG]+[INS] is pressed.<br>
   To avoid copying, return MTBL_NODEFAULTACTION.

### Fixes
- Under certain conditions, the application crashes when an edit control is created

- Under certain conditions, the application crashes when SalTblReset is called while a tooltip is visible

- Under certain conditions, MTblExportToExcel computes a wrong font size ( e.g., table font size is 10, calculated Excel font size is 9 )

## V 3.2.0
- General<br>
   Ready for TD 7.3

- New function MTblGetEditHandle<br>
   Gets the window handle of the current edit control.

### Fixes
- Fixed another multithreading issue regarding strings

## V 3.1.0
- General<br>
   Ready for TD 7.2

### Fixes
- Fixed Excel export issue (Excel sometimes throws PasteSpecial exception when exporting via clipboard)

- Fixed x64 issue (random crashes espeacially on Win7 systems)

## V 3.0.2
### Fixes
- Fixed multithreading issue regarding strings (TD 7.1)

## V 3.0.1
- New row flags:
  - MTBL_ROW_BLOCK_BEGIN
  - MTBL_ROW_BLOCK_END<br>
   Can be used to keep rows together on one page when printing.

### Fixes
- M!Table crashes when edit mode is ended with ESC and ROW_Selected or ROW_Hidden is set on SAM_Validate

- Under certain conditions, popup edits are editable even if the column is disabled ( only TD 6.3 or higher )

- Under certain conditions, tooltips are displayed on several monitors, e.g. one part on the left monitor and one part on the right monitor

- Under certain conditions, the vertical scrollbars aren't visible when the table has split rows and/or the Office 2013 theme is used

## V 3.0.0
- General<br>
   Ready for TD 7.1<br>
   Ready for 64 Bit (TD 7.0 and TD 7.1)<br>
   Print functionality was migrated from Delphi to C++ and is now included in mtblXX.dll, hence mtblprXX.dll is no longer needed

- Support for theme "Office 2013"

### Fixes
- The vertical scrollbar is not always shown even if TBL_Flag_ShowVScroll is set

- In newer TD versions (e.g. 6.3), the vertical scrollbar isn't visible when variable row heights are activated

## V 2.1.5
- General<br>
   Ready for TD 7.0 (32bit)

### Fixes
- Bad Excel export performance when a table with very long cell texts is exported

- Under certain conditions, the vertical scrollbar isn't visible when variable row heights are activated

- Under certain conditions, pressing TAB in a Table Window sets the input focus to a cell which was disabled with MTblDisableRow or MTblDisableCell

- Under certain conditions, checkboxes are not toggled if different checkboxes are clicked quickly in succession

## V 2.1.4
- General<br>
   Ready for TD 6.3<br>
   Please note that from this version on the Visual C++ 2013 Redistributable Package (x86) is required

- MTblObjFromPoint<br>
   Sets the flag MTOFP_OVERIMAGE for column headers.

   New constants:
  - MTOFP_OVERTEXT ( same value as MTOFP_OVERCELLTEXT )
  - MTOFP_OVERIMAGE ( same value as MTOFP_OVERCELLIMAGE )

   

### Fixes
- Wrong vertical text alignment in column header groups with Flag MTBL_CHG_FLAG_NOCOLHEADERS.

- If the type of a cell is changed more than once with MTblDefineCellType, the new cell type may not be applied immediately.

## V 2.1.3
- General<br>
   Ready for TD 6.2

- Tree
  - New tree flag MTBL_TREE_FLAG_AUTO_NORM_HIER<br>
    If set, the hierarchy is normalized automatically if MTBL_ROW_HIDDEN is set/removed or rows are deleted.<br>
    Descendant rows are automatically moved in the hierarchy below the first visible ancestor row if a row was deleted or hidden.<br>
    If a row was shown, the row itself and the descendant rows are automatically moved in the hierarchy below the first visible ancestor row, on the basis of the original hierarchy.

  - New message MTM_QueryAutoNormHierarchy<br>
    If MTBL_TREE_FLAG_AUTO_NORM_HIER is set, this message is sent just before a row is hidden, shown or deleted.<br>
    To avoid normalizing, return MTBL_NODEFAULTACTION.

  - New function MTblGetOrigParentRow<br>
    Gets a row's original parent row.

### Fixes
- Cell button tips are not shown if the cell is in edit mode.<br>
   This bug exists only since 2.1.2.

- In expanded rows, M!Table doesn't use the image set with MTblSetCellImage if no expanded image is set with MTblSetCellImageExp.<br>
   This bug exists only since 2.1.2.

- MTblExportToExcel(Ex): One row more than necessary is cleared before inserting the data which leads to a blank row after the last exported row

- The table continuously receives WM_MOUSEMOVE messages, even if the mouse is not moving.<br>
   This bug exists only since 2.1.2.

- Copying the text of a read-only cell with keyboard shortcuts (e.g. [CTRL] + [C]) doesn't work

- Gradient row headers are not painted correctly at the lower edge if TBL_Flag_SuppressRowLines is set

## V 2.1.2 Patch 1
### Fixes
- Returning TRUE on MTM_DragCanAutoStart doesn't start the drag drop mode anymore

## V 2.1.2
- Item concept<br>
   With this version, the item concept has been introduced.<br>
   An item is a functional class which identifies a table object (e.g. a cell) or a part of a table object (e.g. a cell image).<br>
   Most of the new functions in this release are using this new item class.

   It's planned to replace item specific functions with generic item functions in the next releases.<br>
   For example, the functions MTblSetCellBackColor, MTblSetColHdrGrpBackColor, MTblSetColumnBackColor, MTblSetColumnHdrBackColor, MTblSetCornerBackColor,<br>
   MTblSetRowBackColor and MTblSetRowHdrBackColor would be replaced by a single function: MTblSetItemBackColor.<br>
   Such replacements would lead to a clearer API.<br>
   Don't worry, the old functions would continue to exist (for backward compatibility).

   Related new classes:
  - fcMTblItem<br>
     Item definition

- Enter/Leave-Notification<br>
   A new feature which sends messages when the mouse pointer has entered or left a specific table item (e.g. a cell).
   
   Related new extended messages:
  - MTM_MouseEnterItem<br>
     This message is sent when the mouse pointer has entered a table item.

  - MTM_MouseLeaveItem<br>
     This message is sent when the mouse pointer has left a table item.

   Related new functions:
  - MTblGetItem<br>
     Gets the data of an item.

- Highlighting<br>
   A new feature which enables you to highlight table items (e.g. cells).<br>
   For example, items can be highlighted with another background color or another image.<br>
   Together with the new Enter/Leave-Notification you can easily implement "hot tracking".

   Related new functions:
  - MTblDefineHighlighting<br>
     Defines the highlighting settings of an item type.

  - MTblQueryHighlighting<br>
     Queries the highlighting settings of an item type.

  - MTblSetHighlighted<br>
     Sets or removes the highlighted state of an item.

  - MTblIsHighlighted<br>
     Determines whether an item is highlighted.

  - MTblGetEffHighlighting<br>
     Gets the effective highlighting settings of an item.

   Related new classes:
  - fcMTblHiLiDef<br>
     Highlighting definition.


- Tooltips
  - New tip type MTBL_TIP_COLHDR_BTN ( for the new column header buttons )

  - MTblSetTipFlags<br>
    New flag MTBL_TIP_FLAG_GRADIENT<br>
    If set, the tooltip is drawn with a gradient background, beginning at the top with white and ending at the bottom with the current background color.

  - Tooltips aren't dsiplayed anymore while the mouse input is captured.<br>
    For example, this is the case
    - while selecting rows or cells
    - while sizing a row or a column
    - while moving a column
    - while clicking a button


- Tree
  - Node images<br>
    Now it's possible to replace the default (outdated) +/- symbols with any images you like.<br>
    Related new functions:
    - MTblSetNodeImages
    - MTblGetNodeImages


- Buttons
  - Column header buttons<br>
    Now it's possible to define buttons for column headers.

  - New functions
    - MTblDefineItemButton<br>
     Let you define buttons for cells, columns and column headers.<br>
     Replaces the functions MTblDefineButtonColumn and MTblDefineButtonCell.

    - MTblQueryItemButton<br>
     Let you query button definitions for cells, columns and column headers.<br>
     Replaces the functions MTblQueryButtonColumn and MTblQueryButtonCell.

  - New class fcMTblBtnDef<br>
    Button definition

  - New flags
    - MTBL_CELL_FLAG_BUTTONS_PERMANENT
    - MTBL_COL_FLAG_BUTTONS_PERMANENT.

 
- New extended messages
  - MTM_LinesPerRowChanged<br>
    This message is posted when the user has changed the number of lines per row with the mouse.

  - MTM_RowSized<br>
    This message is posted when the user has changed the height of a row with the mouse.

- New row flags
  - MTBL_ROW_HIDDEN<br>
    The row is hidden and is - in contrast to ROW_Hidden - NOT shown if the parent row is being expanded.

- fcMTblODI
  - New instance variables HWndPart and NPart<br>
    Part of the item to draw

### Fixes
- Mousewheel scrolling for listboxes doesn't work from TD version 5.1 on.

- The cursor doesn't change to an arrow when a popup menu was created with SalTrackPopupMenu

- Setting the column title with MTblSetColumnTitle/SalTblSetColumnTitle corrupts the row/cell focus when the number of text lines have changed

- If MTBL_FLAG_NO_COLUMN_HEADER or MTBL_FLAG_VARIABLE_ROW_HEIGHT is set and cell mode isn't active:<br>
   SAM_Click isn't sent when clicking on a disabled cell or on the area to the right of the last column

- Poor performance when setting ROW_Selected with SalTblSetFlagsAnyRows for a very large number of rows

- If a new row is inserted with SalTblInsertRow, alternating row background colors and/or permanently visible buttons get not visible unless the row is repainted due to another action

## V 2.1.1
- General
  - M!Table doesn't throw an error anymore when any of the default languages can't be registered while loading mtblxx.dll

- New table flags
  - MTBL_FLAG_MOVE_INP_FOCUS_UD_EX<br>
    If set and a multiline cell is in editing mode:<br>
    If the caret is in the first line and the up-arrow is pressed, the input focus is moved up (to the next editable cell above)<br>
    If the caret is in the last line and the down-arrow is pressed, the input focus is moved down (to the next editable cell below )

### Fixes
- If a cell is in editing mode, pressing [Del] deletes 2 characters ( rather than 1 ).<br>
   This bug exists only since 2.1.0.

- If MTBL_FLAG_NO_COLUMN_HEADER or MTBL_FLAG_VARIABLE_ROW_HEIGHT is set:<br>
   The new value of a cell might not be displayed immediatley, e.g. by a statement like "Set tbl1.col1 = "New value".<br>
   This bug exists only since 2.1.0.

- If an edit control isn't autosized, it has only one visible line of text in editing mode, even if "Lines Per Row" is > 1 and "Word Wrap" is enabled.<br>
   This bug exists only since 2.1.0.

- The lower part of an autosized multiline edit control is possibly invisible, thus it doesn't scroll to invisible text lines.<br>
   This bug exists only since 2.1.0.

## V 2.1.0
- General
  - Ready for TD 6.1 (Win32)
  - Significantly improved performance, particularly with many columns and/or merged cells

- Variable row height<br>
   By popular demand, each row may have a variable height now!!

   Related new functions:<br>
   MTblSetRowHeight, MTblGetRowHeight, MTblGetEffRowHeight

   Related new table flags:<br>
   MTBL_FLAG_VARIABLE_ROW_HEIGHT

- Merging rows<br>
   Added the possibility to merge rows. Merged rows are behaving and appearing like one single row.

   Related new functions:<br>
   MTblSetRowMerge, MTblGetRowMerge, MTblGetMergeRow, MTblGetLastMergeRow

- Hidden column header<br>
   The column header may be invisible now

   Related new table flags:<br>
   MTBL_FLAG_NO_COLUMN_HEADER

- Autosizing edit<br>
   The size of the edit control in editing mode may be automatically adapted now.

   Related new flags:<br>
   MTBL_CELL_FLAG_AUTOSIZE_EDIT<br>
   MTBL_COL_FLAG_AUTOSIZE_EDIT<br>
   MTBL_ROW_AUTOSIZE_EDIT

- New functions
  - MTblGetSepRow<br>
    Returns, if the passed coordinates are over a row header separator, the number of the row the separator belongs to.

- New extended messages
  - MTM_EditHeightChanged
  - MTM_RowHdrSepLBtnDown
  - MTM_RowHdrSepLBtnUp
  - MTM_RowHdrSepLBtnDblClk
  - MTM_RowHdrSepRBtnDown
  - MTM_RowHdrSepRBtnUp
  - MTM_RowHdrSepRBtnDblClk

- MTblAutoSizeRows<br>
   If the flag MTBL_FLAG_VARIABLE_ROW_HEIGHT is set, the row heights are adjusted ( instead of setting lines per row ).
  - New flag MTASR_CONTEXTROW_ONLY<br>
     If the flag MTBL_FLAG_VARIABLE_ROW_HEIGHT is set, only the height of the current context row is adjusted.<br>
     If the flag MTBL_FLAG_VARIABLE_ROW_HEIGHT is not set, only the the current context row is used to determine the best number of lines per row.

- MTblObjFromPoint
  - New flag MTOFP_OVERROWHDRSEP<br>
     Is set if the mouse is over a row header separator

- MTblSetExtMsgsOptions
  - New option MTEM_ROWHDRSEP_EXCLUSIVE<br>
    MTM_RowHdrSep* messages are exclusive, MTM_RowHdr* messages are not posted.

### Fixes
- MTblGetRowID returns 0 if no M!Table function for the row and/or a cell of the row was called before

- If a column is only partly visble and the cursor is at the right edge of the visible column area:<br>
   MTblGetSepCol wrongly returns a separator column and MTblObjFromPoint wrongly sets the flag MTOFP_OVERCOLHDRSEP

- Double click on a row header expands the row even if row sizing is allowed and the cursor is over a row header separator ( row sizing cursor is displayed )

- Under certain conditions, permanently visible buttons are partly inverted

- When the table is in cell mode and [PageDown] or [PageUp] is pressed, the focus frame and selection of the new focus cell isn't displayed

- Under certain conditions, the check state of checkbox cells isn't displayed correctly in formatted columns

## V 2.0.2
- Cell mode
  - New message MTM_QueryNoCellSelection<br>
    This message is sent when a cell is about to get selected due to a keyboard action, a mouse click or calling MTblSetCellFlags.<br>
    If a value != 0 is returned, the cell doesn't get selected.

- Extended messages
  - New functions
    - MTblSetExtMsgsOptions<br>
     Sets or clears options for extended messages.

    - MTblQueryExtMsgsOptions<br>
     Queries options for extended messages.

  - New options
    - MTEM_CORNERSEP_EXCLUSIVE<br>
     MTM_CornerSep* messages are exclusive, MTM_Corner* and MTM_ColHdr* messages are not posted.
    - MTEM_COLHDRSEP_EXCLUSIVE<br>
     MTM_ColHdrSep* messages are exclusive, MTM_ColHdr* messages are not posted. This option is set by default.

- MTblExportToExcel(Ex)
  - New excel flag MTE_EXCEL_NO_CLIPBOARD<br>
    If set, the clipboard isn't used to set cell data in Excel.<br>
    Instead, the cell data is set directly by "Range.SetFormulaLocal".<br>
    Setting the cell data directly is somewhat slower, but it avoids that - as some users reported - sometimes "Worksheet.PasteSpecial" throws an exception.

  - New excel flag MTE_EXCEL_NO_AUTO_FIT_COL<br>
    If set, the Excel method "AutoFit" isn't performed on columns

  - New excel flag MTE_EXCEL_NO_AUTO_FIT_ROW<br>
    If set, the Excel method "AutoFit" isn't performed on rows

  - New message MTM_QueryExcelValue<br>
    This message is sent once for every exported cell while exporting to Excel.<br>
    If you return a value ( a string converted to a number by SalHStringToNumber ),<br>
    M!Table uses the returned value rather than determining the value of appropriate cell.

  - New message MTM_QueryExcelNumberFormat<br>
    This message is sent once for every exported column and cell while exporting to Excel.<br>
    If you return a format ( a format string converted to a number by SalHStringToNumber ),<br>
    M!Table sets the appropriate cell(s) in Excel to the specified number format before pasting the cell data.<br>
    If a format for a numeric column is returned, M!Table exports all cells of the column unformatted.

- New message MTM_QueryExcelColWidth<br>
    This message is sent once for every exported column while exporting to Excel.<br>
    If you return a value != 0, M!Table sets the width of the appropriate column in Excel.<br>
    The desired value needs to be multiplied by 100 (to be able to use 2 fractional digits).

- New message MTM_QueryExcelRowHeight<br>
    This message is sent once for every exported row while exporting to Excel.<br>
    If you return a value != 0, M!Table sets the height of the appropriate row in Excel.<br>
    The desired value needs to be multiplied by 100 (to be able to use 2 fractional digits).

  - Automatic calculation<br>
    For performance reasons, automatic calculation in Excel is switched off during the export.<br>
    The original setting is restored after the export.

- MTblObjFromPoint
  - New flag MTOFP_OVERCOLHDRSEP<br>
    Is set if the mouse is over a column header separator.

  - New flag MTOFP_OVERCORNERSEP<br>
    Is set if the mouse is over the vertical corner separator.

- New functions
  - MTblGetSepCol<br>
    Returns, if the passed coordinates are over a column header separator, the handle of the column the separator belongs to.

- New extended messages
  - MTM_CornerSepLBtnDown
  - MTM_CornerSepLBtnUp
  - MTM_CornerSepLBtnDblClk
  - MTM_CornerSepRBtnDown
  - MTM_CornerSepRBtnUp
  - MTM_CornerSepRBtnDblClk

### Fixes
- When a cell's edit alignment differs from the column's text justification,<br>
   the column receives SAM_AnyEdit when the input focus moves to another cell or the editing mode ends.<br>
   This defect was introduced in version 2.0.1.

- MTblExportToHTML<br>
   Missing semicolon in conversion of special character " (double quotation mark)

- MTblExportToExcel(Ex)<br>
   Fixed locale problem with cell data (when transferred via clipboard)

## V 2.0.1
- General
  - Ready for TD 6.0 (Win32)

- Windows Themes<br>
   M!Table now provides a better support for Windows Themes.<br>
   All child controls ( buttons, drop down buttons, checkboxes etc. ) are drawn according to the current Windows Theme.<br>
   Furthermore, the default colors of the current Windows Theme are used.<br>
   You can disable this behaviour by setting the (new) table flag MTBL_FLAG_IGNORE_WIN_THEME.

- Buttons
  - New button flag MTBL_BTN_LEFT<br>
    If set, the button is left aligned
  - New button flag MTBL_BTN_NO_BKGND<br>
    If set, the button is drawn without background
  - New button flag MTBL_BTN_FLAT<br>
    If set (and no Windows Theme is used), the button is drawn with a flat border
  - Added "3D effect" for M!Table buttons<br>
    If a M!Table button is clicked, text and image move 1 pixel to the right/bottom (if no Windows Theme is used or MTBL_BTN_NO_BKGND is set)

- Word Wrap<br>
   If a cell has more than 1 row because of merging and the column flag COL_MultilineCell is set, <br>
   the text is displayed with automatic line breaks and the editing mode is multiline.

- New table flags
  - MTBL_FLAG_IGNORE_WIN_THEME<br>
    If set, M!Table ignores the current Windows Theme.<br>
    All child controls ( buttons, drop down buttons, checkboxes etc. ) are drawn according to the classic style.<br>
    Furthermore, the default colors of the classic style are used.
  - MTBL_FLAG_BUTTONS_FLAT<br>
    If set (and no Windows Theme is used), all M!Table and drop down buttons in the table are drawn with a flat border

### Fixes
- Tooltip of permanent buttons doesn't change if the mouse is moved pretty fast from one button to another.

- If a Windows Theme is active, checkboxes in editing mode are not drawn according to the Windows Theme.

- MTblAutoSizeColumn calculates the width of the drop down button into the column width even when MTBL_FLAG_BUTTONS_OUTSIDE_CELL is set.

- When clicking an editable cell and SalMessageBox is called on SAM_Validate,<br>
   the mouse input remains captured by the edit control of the clicked cell unless one presses and releases the left mouse button again.

- Focus is not moving to the previous column when pressing [SHIFT] + [TAB] when SalMessageBox is called on SAM_Validate.

## V 2.0.0
- Localization<br>
   Added support for custom languages (language files).

   Hence, M!Table itself also uses language files for the default languages:<br>
   German => mtbllang.de<br>
   English => mtbllang.en

   Please note that you have to distribute these files with your application, otherwise M!Table will not function properly.

  - New functions:
    - MTblRegisterLanguage<br>
      Registers a language ID and its related language file
    - MTblGetLanguageFile<br>
      Retrieves the name of the language file related to a language ID
    - MTblGetLanguageString<br>
      Retrieves a language-specific string

  - New constants
    - MTE_ERR_INVALID_LANG<br>
      The MTblExport* functions return this value if the specified language is invalid

    - MTP_ERR_INVALID_LANG<br>
      MTblPrint returns this value if the specified language is invalid 
      

- Tooltips<br>
   The tooltip window is now created when it's needed and destroyed if it's no longer needed.<br>
   This saves window ressources in applications with many tables.

### Fixes
- Drop down lists<br>
   If a cell is near the bottom of the form the list appears at the top.<br>
   It seems that this problem occurs since TD 5.2 SP1.

   Under certain conditions, SalListQuerySelection returns LB_Err.<br>
   It seems that this problem occurs since TD 5.x.

## V 1.9.9
- General
  - Initial Memory consumption<br>
    The initial memory consumption caused by subclassing was reduced dramatically (from ~3664 KB to ~196 KB).<br>
    This was achieved by using dynamic instead of static memory allocation for (16384) column objects.<br>
    However, the effective memory consumption will grow the more column objects are created.<br>
    But because most tables don't have more than 200 columns, also the effective memory consumption will be reduced dramatically!

- Tree
  - New tree flag MTBL_TREE_FLAG_INDENT_NONE<br>
    If set, rows on level 0 are not indented and have no nodes.

- Buttons<br>
   Buttons ( M!Table buttons and drop down buttons ) now can be shown permanently by setting the new table flag MTBL_FLAG_BUTTONS_PERMANENT.<br>
   If permanent buttons are switched on by the table flag, they can be switched off<br>
   on column level by setting the flag MTBL_COL_FLAG_HIDE_PERM_BTN<br>
   or on cell level by setting the flag  MTBL_CELL_FLAG_HIDE_PERM_BTN.

   Permanent buttons can be hidden in non editable cells ( disabled or read only ) by setting the new table flag MTBL_FLAG_HIDE_PERM_BTN_NE.<br>
   If permanent buttons are hidden in non editable cells by the table flag, they can be still shown<br>
   on column level by setting the flag MTBL_COL_FLAG_SHOW_PERM_BTN_NE<br>
   or on cell level by setting the flag  MTBL_CELL_FLAG_SHOW_PERM_BTN_NE.

- Cell mode<br>
   New flag MTBL_CM_FLAG_NO_SELECTION<br>
   If set, cells are never selected.

   New message MTM_QueryNoFocusOnCell<br>
   This message is sent when the cell focus is about to change due to a keyboad action or a mouse click.<br>
   If a value != 0 is returned, the cell doesn't receive the focus.

- Dropdown lists<br>
   New table flags MTBL_FLAG_USE_TBL_FONT_DDL.<br>
   If this flag is set, the table font is used for dropdown lists instead of the effective cell font.

   From now the following VisList*Value functions can be used:<br>
   VisListAddValue, VisListFindValue, VisListGetValue, VisListInsertValue and VisListSetValue.

- Popup edits<br>
   New table flags MTBL_FLAG_USE_TBL_FONT_PE.<br>
   If this flag is set, the table font is used for popup edits instead of the effective cell font.

- Hit test<br>
   New M!Table flags MTblObjFromPoint:<br>
   MTOFP_OVERBTN ( is set when the mouse is over a permanently visible M!Table button )<br>
   MTOFP_OVERDDBTN ( is set when the mouse is over a permanently visible drop down button )

- Cell images<br>
   MTblSetCellImage/MTblSetCellImageExp: New flag MTSI_NOTEXTADJ<br>
   If this flag is set, the cell text will not be adjusted because of the cell image.

- Table Property Editor<br>
   Added the following properties:
  - "Use Table Font For Drop Down Lists" ( in Category "(General)" - "Font Handling" )
  - "Use Table Font For Popup Edits" ( in Category "(General)" - "Font Handling" )
  - "Row Flag Images" ( in category "Rows" )
  - "Indent No Rows On Level 0" ( in category "Tree" )
  - "Show" ( in category "Buttons", group "Permanent" )
  - "Hide In Non editable Cells" ( in category "Buttons", group "Permanent" )
  - "No selection" ( in category "Cell Mode" )

- Column Property Editor<br>
   Added the following properties:
  - "Hide" ( in category "Button", group "Permanent" )
  - "Show In Non Editable Cells" ( in category "Button", group "Permanent" )

### Fixes
- Buttons<br>
   When the button hotkey is pressed, MTM_BtnClick is sent even though if the button is disabled.

- MTblDefineCellType<br>
   If a cell of a drop down list column is defined as drop down list cell with MTblDefineCellType,<br>
   under certain conditions the wrong list values are displayed for cells which have no own cell type definition.

- MTblExportToExcel / MTblExportToExcelEx:<br>
   Under certain conditions, an "unknown OLE dispatch exception" occurs at the end of the export process (when M!Table calls the Excel function "AutoFit").

- MTblSetRowFlagImage:<br>
   ROW_user* flags are now more significant than the standard flags ( ROW_MarkDeleted, ROW_New, ROW_Edited ).<br>
   Thus, MTblSetRowFlagImage now behaves like VisTblSetRowPicture.

- Row header title is left aligned, but must be centered.

## V 1.9.8
- General
  - Ready for TD 5.2

- Property Editors<br>
   Property editors for table and column properties which can be integrated in the IDE using the Quick Object Interface.

- Improved Help<br>
   M!Table now ships with a Windows help file ( mtbl.chm ) which replaces the old Manual folder.<br>
   Furthermore, a help launcher tool is provided which can be intergrated in the IDE using the Tools interface.

- Excel Export
  - New function MTblExportToExcelEx:<br>
    Allows you ( in addition to MTblExportToExcel ) to
    - load a template file before export
    - save the Workbook as a file after export
    - set the Worksheet name

  - New class fcMTblExcelExpParams ( needed for MTblExportToExcelEx )

  - New Excel flags:
    - MTE_EXCEL_HIDE_INSTANCE
    - MTE_EXCEL_CLOSE_INSTANCE
    - MTE_EXCEL_CLOSE_WORKBOOK

- New messages
  - MTM_BeforeEnterEditMode<br>
    Is sent before the table enters the editing mode

### Fixes
- MTblExportToOOCalc doesn't export any cell texts along with the Unicode versions of Team Developer ( 5 or higher )

- MTblAutoSizeColumn includes the drop down button width only when the column type is COL_CellType_DropDownList

- M!Table crashes when "Discardable" is "No" and a row insertion exceeds "Max Rows In Memory"

## V 1.9.7
- General
  - TD-Themes<br>
    M!Table now uses default colors according to the TD-Themes

  - Performance<br>
    Images without transparency are painted much faster now

- New functions
  - MTblExportToOOCalc<br>
    Exports a table to Open Office Calc

- New messages
  - MTM_SetCursor<br>
    Replaces WM_SETCURSOR

- New table flags
  - MTBL_FLAG_IGNORE_TD_THEME_COLORS<br>
    If set, M!Table uses the "old fashioned" default colors

- New constants
  - MTblExportToOOCalc
    - MTE_OOCALC_NEW_WORKBOOK
    - MTE_OOCALC_NEW_WORKSHEET
    - MTE_OOCALC_CURRENT_POS
    - MTE_OOCALC_STRING_COLS_AS_TEXT
    - MTE_OOCALC_NULLS_AS_SPACE

### Fixes
- The table doesn't enter the editing mode when ROW_Selected is set using SalTblSetRowFlags or SalTblSetFlagsAnyRows in the SAM_Click message handler

- Under certain conditions, clicking the column header or column separator causes an infinite validation loop.

- When a string column is sorted with flag MTS_DT_DATETIME, actually identical values are evaluated slightly different.<br>
   This leads to an "unstable" sort order which is particularly noticed when sorting multiple columns.

- Very bad performance when adding a large number of items with SalListAdd or SalListInsert.<br>
  This problem exists since V1.9.5.

- If the table is in editing mode and a disabled cell or row is clicked,<br>
  the table scrolls to the row with the focus cell if the focus cell isn't in the visible range.

  - SAM_AnyEdit is sent recursive when calling SalTblSetFocusCell inside the SAM_AnyEdit message handler.

  - SAM_RowValidate is sent twice when one clicks into a different row.<br>
   No problems when one uses the keyboard.

  - MTblCopyRows* functions become very slow if more than 2 MB of data is copied

  - A checkbox cell retrieves SAM_Click AFTER it's value has changed. This is not the default behaviour.

  - A cell in a checkbox column, which defined cell type is not COL_CellType_CheckBox:<br>
     a) can't receive the edit focus when the cell is read only<br>
     b) doesn't display the complete text tooltip

## V 1.9.6
- New messages
  - MTM_DragCanAutoStart<br>
    Because SAM_DragCanAutoStart doesn't work correctly anymore since M!Table 1.9.0, this message is introduced.

- New table flags
  - MTBL_FLAG_GRADIENT_HEADERS<br>
    Draws all headers with a gradient background, beginning at the top with white and ending at the bottom with the current background color

### Fixes
- When a table receives the focus after it lost the focus in editing mode, the row focus frame isn't painted correctly

- Under certain conditions, multiple SAM_Validate messages are sent

- The column header text isn't drawn in disabled state when the undocumented column flag 0x020 ( COL_DimTitle ) is set

- Under certain conditions, SalTblDestroyColumns causes a crash

- When cell mode is active and the table is in editing mode and a "column size click" is made, the cell which was in editing mode isn't selected

- If the table is in editing mode and SalTblSetColumnWidth is called, the row whose cell was in editing mode is drawn selected even if MTBL_FLAG_SUPPRESS_ROW_SELECTION is set

## V 1.9.5
- General
  - Checkbox cells
    - Can be merged now
    - Can have a button now
    - *TXTALIGN* flags now controlling the position of the checkbox
    - Flag MTBL_CELL_FLAG_HIDDEN_TEXT now hides the checkbox
    - MTblObjFromPoint now sets the flag MTOFP_OVER_CELLTEXT when the given position is within a checkbox
    - MTM_ShowCellTextTip is now sent when the mouse is over a checkbox
    - An image now affects the position of the checkbox like a text were affected
  - Checkbox columns
    - Can be the tree column now
       
- New functions
  - MTblDefineCellType / MTblQueryCellType<br>
     Defines/Queries the type definition ( Standard, PopupEdit, DropDownList, Checkbox ) of a cell

  - MTblGetEffCellIndent<br>
     Gets a cell's effective indent in pixels ( = [cell indent level] * [indent per level] )

  - MTblGetCellType<br>
     Returns the type of a cell

  - MTblSetCellIndentLevel / MTblGetCellIndentLevel<br>
     Sets/Gets a cell's indent level

  - MTblSetDropDownListContext / MTblQueryDropDownListContext<br>
     Sets/Queries the current drop down list context

  - MTblSetHeadersBackColor / MTblGetHeadersBackColor<br>
     Sets/Gets the default background color of headers ( column headers, row headers and the corner )

  - MTblSetIndentPerLevel / MTblGetIndentPerLevel<br>
     Sets/Gets the indent per level in pixels

- New tree flags
  - MTBL_TREE_FLAG_FLAT_STRUCT<br>
     Rows of all levels have the same indent, which results in a "flat" tree structure

- New cell flags
  - MTBL_CELL_FLAG_CLIP_NONEDIT<br>
     The nonedit area ( e.g. an image ) is clipped and remains visible in editing mode

- New column flags
  - MTBL_COL_FLAG_CLIP_NONEDIT<br>
     The nonedit area ( e.g. an image ) is clipped and remains visible in editing mode

- New row flags
  - MTBL_ROW_CLIP_NONEDIT<br>
     The nonedit area ( e.g. an image ) is clipped and remains visible in editing mode

- MTblExportToHTML
  - When a column has the flag MTBL_COL_FLAG_EXPORT_AS_TEXT, the text of a cell is exported "as it is" ( no replacement of characters with HTML special characters )
  - Added the style "border-collapse: collapse" to the HTML-table which leads to a "effective" border width of 1 pixel

### Fixes
- SalTblSortRows is very slow, no problems with MTblSort*

- MTblExportToExcel always fetches all rows of a table, even if this isn't necessary because of a row flag restriction

- MTblExportToExcel now uses the correct worksheet constraints for Excel 2007 ( max. 16384 columns and max. 1048576 rows )

- Fixed fetch order problem when setting ROW_Hidden on SAM_FetchRow while the table is populated with SalTblSetRange

- Selections are drawn incorrectly when setting row flags ROW_Selected or ROW_Hidden while the table is in editing mode

- Checkboxes have the wrong size if the value for "DPI" in the Windows display settings is not the default (=96 DPI)

- Cell mode: Entering the editing mode automatically sometimes doesn't work with merged cells

- MTblAutoSizeColumn with flag MTASC_BTNSPACE inludes the button width even if table flag MTBL_FLAG_BUTTONS_OUTSIDE_CELL is set

- Selections are drawn incorrectly after setting the cell type with SalTblDefine*Column

- MTblObjFromPoint: Flags MTOFP_OVERCELLIMAGE and/or MTOFP_OVERCELLTEXT not set with checkbox cells

  - MTblExportToHTML wrongfully uses the unit "pt" instead of "px" for the CSS-style "text-indent"

- Under certain conditions, the colors set with MTblSetAltRowBackColors are not visualized

- Definition of MTM_CHAR is missing in mtbl.apl

- When the cell mode is active an one scrolls up with the arrow keys, sometimes fragments of the "original" row focus frame are visible

- When a disabled cell is clicked while the table is in editing mode, the row whose cell was in editing mode is incorrectly drawn with a selection

- Some GDI objects for dotted column and/or row lines are not destroyed when a table is destroyed

## V 1.9.0 Patch 1
### Fixes
- TD 5.1 only (UNICODE): MTblExportToHTML now writes a "Byte Order Mark" at the beginning of the file ( 0xFFFE ) because<br>
   the Internet Explorer doesn't display chinese characters correctly without it. No problems with Firefox.

- Readonly cells of a column with input mask are editable anyway

- When navigating through the table with the keyboard and the editing mode is killed because the "end" of the table is reached,<br>
   buttons with keyboard accelerator [ESC] receive SAM_Click.

## V 1.9.0
- General
  - Ready for TD 5.1

- New messages
  - MTM_QueryMinColWidth<br>
   This message is sent when M!Table have to determine the min. column width, e.g. when MTblAutoSizeColumn is called

  - MTM_QueryMinRowHeight<br>
   This message is sent when M!Table have to determine the min. row height, e.g. when MTblAutoSizeRows is called

## V 1.9.0 Beta
- General
  - M!Table implements an API-Hook for TABLIxx.DLL.<br>
   This is mainly done to get rid of the focus frame and selections drawn by the table control.<br>
   The result is an absolutely flicker free table!

- New functions
  - MTblAutoSizeRows<br>
   Adjusts the number of lines per row according to the cell with the highest contents

  - MTblCalcRowHeight<br>
   Calculates the row height in pixels for a given number of lines per row

  - MTblClearCellSelection<br>
   Clears the current cell selection

  - MTblDefineCellMode / MTblQueryCellMode<br>
   Defines/Queries the cell mode.<br>
   In cell mode, a single cell has the focus instead of a whole row.<br>
   Furthermore it's possible to select individual cells.

  - MTblDefineFocusFrame / MTblQueryFocusFrame<br>
   Defines/Queries the appearance of the focus frame.

  - MTblSetFocusCell / MTblQueryFocusCell<br>
   Sets/Queries the current focus cell when the table is in cell mode

- New messages
  - MTM_FocusCellChanged<br>
   This message is sent when the table is in cell mode and the focus cell has changed

  - MTM_QueryBestCellHeight<br>
   This message is sent when M!Table have to determine the best cell height, e.g. when MTblAutoSizeRows is called

  - MTM_QueryMaxRowHeight<br>
   This message is sent when M!Table have to determine the max. row height, e.g. when MTblAutoSizeRows is called

- MTblAutoSizeColumn
  - New flag MTASC_INCREASE_ONLY<br>
   If set, the column width is only adjusted when the calculated best width exceeds the current width

### Fixes
- The upper part of the focus frame may not be erased when the table scrolled down<br>
   because the [Down] key was pressed

- When the table loses the focus while it is in editing mode, the focus frame may not be erased

- When the table font is changed, the focus frame isn't updated

- When the cell mode is deactivated with MTblDefineCellMode,<br>
   fragments of the "old" focus frame are still visible.

- MTM_FocusCellChanged was sent when the internal focus cell changed.<br>
   Now the message is sent when the effective focus cell ( the cell with the frame ) changed.

- MTblQueryFocusCell may not return the effective focus cell ( the cell with the frame )<br>
   if the internal focus cell is part of a cell merge.

- If the table is in cell mode, while navigating with the arrow keys the focus cell may "jump"<br>
   from the split row area to the normal row area and vice versa.

- When the row height is greater than the height of the split area, split rows are not displayed<br>
   correctly, in particular after scrolling the split area.<br>
   This defect was caused by a bug in SalTblQuerySplitWindow, which sometimes delivers 0 in "nRowsLowerHalf",<br>
   even though the table is splitted.

## V 1.9.0 Beta
- Screen flickers when the column flag COL_Selected is set/removed with SalTblSetColumnFlags<br>
   but the flag wasn't effectively changed. The flickering only occurs when:
  - A *NOSELINV flag is set
  - A selection color is set
  - The flag MTBL_FLAG_FULL_OVERLAP_SELECTION is set

- Print preview:<br>
   Even if language MTP_LNG_ENGLISH is set, the cancel button of the scale dialog has the text "Abbrechen" instead of "Cancel"

- MTblSetCellBackColor / MTblSetColumnBackColor / MTblSetRowBackColor:<br>
   Even if MTSC_REDRAW is specified, the background color of a check box cell isn't changed when the cell is in editing mode.

- Dotted lines have always a white "background", regardless of the table background color

## V 1.8.3
- General
  - Ready for TD 4.2

- New functions
  - MTblIsListOpen<br>
   Determines whether the list of a drop down column is currently open

### Fixes
- Listbox sometimes "jumps" to the top of the screen

## V 1.8.2
### Fixes
- In a numeric sort, 0 is treated greater than values <= 0.001 ( e.g in an ascending sort, 0 appears AFTER 0.001 )

- M!Table causes a memory leak every time the table is painted

- Rows may missing when they are disabled on SAM_FetchRow[Done] and the flag MTBL_FLAG_SUPPRESS_FOCUS is set

- All normal cells right beside a vertically merged cell are printed with the text of the cell in the first merged row

- MTM_QueryMaxColWidth doesn't work anymore ( since 1.8.0 )

## V 1.8.1
- MTblSort*
  - Sorting is stable now.<br>
   That means if several rows have the same sort value(s), they keep their position among themselves.

- New table flags
  - MTBL_FLAG_ADAPT_LIST_WIDTH<br>
   The width of drop down lists are adapted so that the widest list entry is completely visible.

  - MTBL_FLAG_SORT_KEEP_ROWMERGE<br>
   When sorting, rows which are part of merged cells keep their position among themselves.

  - MTBL_FLAG_SORT_RESTORE_TREE<br>
   When sorting, the tree is automatically restored if any parent and/or child rows were affected.

### Fixes
- If the position of a column is changed with SalTblSetColumnPos and the extended messages aren't enabled,<br>
   the internal cell merge data isn't updated which leads to unpredictable results if any cell of the column is <br>
   part of a cell merge.

- If the text color of a merge cell is set with SalColorSet, the merged cells aren't redrawn.

- Under certain conditions, MTblGetMergeCell wrongfully returns a merge cell for hidden cells

- MTblSortTree with MTST_BOTTOMUP doesn't work if the table was populated with SalTblSetRange and the <br>
   last row in the scroll range has not been fetched yet.

- Under certain conditions, MTblObjFromPoint fails when the table has no vertical scrollbar.<br>
   Furthermore, all functionality that uses MTblObjFromPoint internally ( e.g. Hyperlinks, Tooltips ) may also not work properly.<br>
   This defect was introduced in 1.8.0 :(.

- The image flag MTSI_NOSELINV ( introduced in 1.8.0 ) doesn't work properly.

## V 1.8.0
- General
  - From now on there's one set of DLL's per TD version.<br>
   The TD version number is appended to the DLL name,<br>
   e.g. mtbl41.dll, mtblee41.dll ( Excel export ) and mtblpr41.dll ( Printing ) for TD 4.1 ( 2005.1 )<br>
   Futhermore, mtbl.apl and mtbl.app for TD11 have now the correct outline version number ( 4.0.26 )

  - All constants in mtbl.apl were moved to the System section

  - Revised automatic column sizing regarding merged cells:<br>
   M!Table now considers cell merging when the best column width is calculated

- New functions
  - MTblDefineEffCellPropDet / MTblQueryEffCellPropDet<br>
   Defines/Queries how particular effective cell properties ( background color, edit justify, font, text color, text justify, vertical text alignment ) are determined

  - MTblFetchMissingRows<br>
   Fetches all rows which are not fetched yet

  - MTblGetRowVisPos<br>
   Gets a row's position within the visible rows

  - MTblQueryRowFlagImageFlags<br>
   Queries the flags of a row flag image

  - MTblSetCellMergeEx / MTblGetCellMergeEx / MTblGetLastMergedCell / MTblGetMergeCell<br>
   New functions for extended cell merging ( you can merge columns AND rows now )

  - MTblSetAltRowBackColors / MTblGetAltRowBackColors<br>
   Sets/Gets alternating row background colors of normal or split rows

- MTblSetRowFlagImage
  - Special image handle MTBL_HIMG_NODEFAULT<br>
   If passed as image handle, the default row flag image isn't displayed anymore

- New image flags
  - MTSI_NOSELINV<br>
   If set, the image will not be inverted when it's part of a selected item

### Fixes
- Setting/Getting the row header back color of the row headers in the area below the last visible row by<br>
   passing TBL_Error to the appropriate function doesn't work.

- MTblGetEff*Font functions return FALSE when the system font is used.<br>
   This causes among others that printing fails.

- MTM_ColPosChanged is not posted when the user moves a column with the mouse

- MTblPrint: If fcMTblPrintParams.FitCellSize is set, it may occur that the calculated row height isn't sufficient

- MTblPrint: When columns are spread over several pages, the following issues may occur:
  - Different number of rows are printed on pages which should have the same number of rows
  - Rows are missing on some pages
  - Preview: When one is on the first page and the button "Last page" is clicked, the status bar displays "Creating page 2", but nothing happens

- MTblInsertChildRow inserts the new child row directly after the last child row, even if the last child row itself has child rows

## V 1.7.6
- Ready for TD 4.1 ( TD 2005.1 )

- New functions
  - MTblSetColumnHdrFlags / MTblQueryColumnHdrFlags<br>
   You can set/query the following column header flags now:
    - MTBL_COLHDR_FLAG_TXTALIGN_LEFT, MTBL_COLHDR_FLAG_TXTALIGN_RIGHT<br>
     Horizontal text alignment

- New column header group flags
  - MTBL_CHG_FLAG_TXTALIGN_LEFT, MTBL_CHG_FLAG_TXTALIGN_RIGHT<br>
   Horizontal text alignment

- New sort flags ( MImgSort* functions )
  - MTS_STRINGSORT<br>
   If set, strings are sorted with the string sort technique instead of the word sort technique

- New table flags
  - MTBL_FLAG_THUMBTRACK_HSCROLL / MTBL_FLAG_THUMBTRACK_VSCROLL<br>
   Enables horizontal and/or vertical thumbtrack scrolling ( the table is scrolled immediately while the user is dragging the scroll box )

### Fixes
- Under certain conditions, tooltips are displayed at the wrong position on multi monitor systems

- No more rows are printed when the height of a cell exceeds the available total height for cells on one page

- On Win9x only, when the name returned by MTblPrintGetDefPrinterName is used as printer name for MTblPrint,<br>
   the error MTP_ERR_INVALID_PRINTER occurs.

- Column header groups with the flag MTBL_CHG_NOCOLHEADERS are drawn "unselected" when selection colors are used

- Cells of password columns ( = with flag COL_Invisible ) have no margins in editing mode with certain fonts ( e.g. MS Sans Serif ).<br>
   This defect was introduced in 1.7.5 :o(.

## V 1.7.5
- General
  - Up to 35% paint performance gain
  - Improved editing mode for merged cells:
    - The edit area is as wide as the merged cells now - this applies also to popup edits and lists
    - Only the master cell can receive the focus now

- New functions
  - MTblDefineButtonHotkey / MTblQueryButtonHotkey <br>
   You may define/query a hotkey for buttons ( on table level ) now

  - MTblSetMWheelOption / MTblGetMWheelOption<br>
   You may set/get the following options regarding mousewheel scrolling now:<br>
   MTMW_ROWS_PER_NOTCH		Rows to scroll per notch<br>
   MTMW_SPLITROWS_PER_NOTCH	Split-Rows to scroll per notch<br>
   MTMW_PAGE_KEY		The virtual code of the key that, if pressed, affects a page scroll

- New messages
  - MTM_ColPosChanged<br>
   Is posted when the position of a column has changed because the user has moved a column with the mouse or SalTblSetColumnPos was called
  - MTM_PaintDone<br>
   Is sent when a table is painted completely

- New table flags
  - MTBL_FLAG_BUTTONS_OUTSIDE_CELL<br>
   If set, buttons ( drop down and M!Table ) are located right hand outside the cell area
  - MTBL_FLAG_NO_DITHER<br>
   With 8 bit color depth, colors which are not in the VGA palette are replaced with the nearest VGA color instead of dithering them.<br>
   This applies to tooltips, too.

- New row flags
  - MTBL_ROW_EDITALIGN_LEFT, MTBL_ROW_EDITALIGN_CENTER, MTBL_ROW_EDITALIGN_RIGHT<br>
   Edit alignment of a row's cells
  - MTBL_ROW_READ_ONLY<br>
   Makes all cells of a row read only

- New column flags
  - MTBL_COL_FLAG_EDITALIGN_LEFT, MTBL_COL_FLAG_EDITALIGN_CENTER, MTBL_COL_FLAG_EDITALIGN_RIGHT<br>
   Edit alignment of a column's cells
  - MTBL_COL_FLAG_READ_ONLY<br>
   Makes all cells of a column read only

- New cell flags
  - MTBL_CELL_FLAG_EDITALIGN_LEFT, MTBL_CELL_FLAG_EDITALIGN_CENTER, MTBL_CELL_FLAG_EDITALIGN_RIGHT<br>
   Cell edit alignment
  - MTBL_CELL_FLAG_READ_ONLY<br>
   Makes a cell read only

- New messages
  - MTM_LButtonDown<br>
   Is sent to the table when the left mouse button is pressed.<br>
   Use this message instead of WM_LBUTTONDOWN, because it's not always sent to the table.   

- MTblGetCellRectEx
  - New flag MTGR_INCLUDE_MERGED<br>
   If set, the rectangle also includes the merged cells ( if any )

- MTblSort*
  - Enhanced performance
  - Strings are sorted according to the current locale now, so that e.g. German Umlauts are sorted correctly.<br>
   Previous versions did a simple byte by byte comparison.
  - New custom sort type MTS_DT_CUSTOM_STRING which allows custom sort strings.

### Fixes
- When the top of the table is reached by pressing [SHIFT] + [TAB], the application dies with a GPF

- Under certain conditions, when changing to editing mode by clicking in a cell the edit control is grayed

- When calling SalTblFetchRow for a row that isn't fetched yet, M!Table row and/or cell functions called on SAM_FetchRowDone fail and the cells of the fetched row have no text

- When editing the text of a merge-cell and you leave the merge-cell, the merged cells aren't repainted

- When columns are not selectable and a column header is clicked, the row focus is displayed although MTBL_FLAG_SUPPRESS_FOCUS is set

- Cell merging in split rows doesn't work if only cells in split rows are merged

- wParam of MTM_DrawItem is always 0 while printing

- MTblExportToHTML: Doesn't work properly with double byte character sets, e.g. Chinese

- MTblExportToExcel: The flag MTE_EXCEL_STRING_COLS_AS_TEXT doesn't work anymore. This bug was introduced in 1.7.0 - shame on me!

## V 1.7.0
- New functions
  - MTblGetPrevChildRow<br>
   Gets the previous child row

  - MTblGetRowID / MTblGetRowFromID<br>
   M!Table provides unique ID's for rows now.

  - MTblSetColHdrGrpImage / MTblGetColHdrGrpImage / MTblQueryColHdrGrpImageFlags<br>
   You may use images in column header groups now.

  - MTblSetColumnHdrFreeBackColor / MTblGetColumnHdrFreeBackColor / MTblGetEffColumnHdrFreeBackColor<br>
   You may customize the background color of the free column header area ( = area to the right of the last column header ) now

  - MTblSetCornerBackColor / MTblGetCornerBackColor / MTblGetEffCornerBackColor<br>
   You may customize the background color of the corner now

  - MTblSortEx<br>
   Extended sorting with the facility to define the rows to sort with any combination of their position, level and superior row.

- New cell flags
  - MTBL_CELL_FLAG_TXTALIGN_LEFT, MTBL_CELL_FLAG_TXTALIGN_CENTER, MTBL_CELL_FLAG_TXTALIGN_RIGHT<br>
   Horizontal cell text alignment
  - MTBL_CELL_FLAG_TXTALIGN_TOP, MTBL_CELL_FLAG_TXTALIGN_VCENTER, MTBL_CELL_FLAG_TXTALIGN_BOTTOM<br>
   Vertical cell text alignment

- New column flags
  - MTBL_COL_FLAG_TXTALIGN_LEFT, MTBL_COL_FLAG_TXTALIGN_CENTER, MTBL_COL_FLAG_TXTALIGN_RIGHT<br>
   Horizontal text alignment of a column's cells
  - MTBL_COL_FLAG_TXTALIGN_TOP, MTBL_COL_FLAG_TXTALIGN_VCENTER, MTBL_COL_FLAG_TXTALIGN_BOTTOM<br>
   Vertical text alignment of a column's cells
  - MTBL_COL_FLAG_EXPORT_AS_TEXT<br>
   Exports a column's cells as text

- New row flags
  - MTBL_ROW_NO_SORT<br>
   Rows with this flag aren't sorted
  - MTBL_ROW_TXTALIGN_LEFT, MTBL_ROW_TXTALIGN_CENTER, MTBL_ROW_TXTALIGN_RIGHT<br>
   Horizontal text alignment of a row's cells
  - MTBL_ROW_TXTALIGN_TOP, MTBL_ROW_TXTALIGN_VCENTER, MTBL_ROW_TXTALIGN_BOTTOM<br>
   Vertical text alignment of a row's cells

- MTblEnableMWheelScroll
  - When mousewheel scrolling is enabled, the following was added for drop down list columns whose list is open:
    - When the cursor is over the list, the list is scrolled
    - When the cursor is over the cell, the next/previous entry in the list is selected
  - Now 3 rows per notch are scrolled, not only 1

- MTblExportToHTML
  - New HTML export flag
    - MTE_HTML_ROW_HEIGHT<br>
    Sets the current row height of the table in the appropriate tags of the HTML file

### Fixes
- When inserting a row while the table is in editing mode, the extended selections don't work properly

- When printing with fcMTblPrintParams.FitCellSize = TRUE, under certain conditions rows with merged cells are too high

## V 1.6.3 Patch 1
- New functions
  - MTblGetPatchLevel<br>
   Retrieves the M!Table patch level

### Fixes
- When FALSE is returned on SAM_EndCellTab, the app crashes

## V 1.6.3
- Ready for TD 4.0 ( TD 2005 )

- Up to 80% less memory usage!

- New functions
  - MTblExportToHTML<br>
   Exports a table to a HTML file

  - MTblInsertChildRow<br>
   Inserts a child row. Simplifies inserting rows in a tree.

- New messages
  - MTM_QueryMaxColWidth<br>
   This message is posted when M!Table has to determine the maximum column width,<br>
   e.g. when MTblAutoSizeColumn is called

- New table flags
  - MTBL_FLAG_NO_FREE_COL_AREA_LINES<br>
   No lines are painted in the area right of the last visible column

  - MTBL_FLAG_NO_FREE_ROW_AREA_LINES<br>
   No lines are painted in the area below the last visible row

  - MTBL_FLAG_SOLID_ROWHDR<br>
   No horizontal lines and no 3D edges are painted in the row header

  - MTBL_FLAG_SUPPRESS_FOCUS<br>
   Suppresses the row focus frame

- New column flags
  - MTBL_COL_FLAG_NO_COLLINE<br>
   No column line is painted in a column

  - MTBL_COL_FLAG_NO_ROWLINES<br>
   No row lines are painted in a column

- New row flags
  - MTBL_ROW_FLAG_NO_COLLINES<br>
   No column lines are painted in a row

  - MTBL_COL_FLAG_NO_ROWLINE<br>
   No row line is painted in a row

- New cell flags
  - MTBL_CELL_FLAG_NO_COLLINE<br>
   No column line is painted in a cell

  - MTBL_CELL_FLAG_NO_ROWLINE<br>
   No row line is painted in a cell
   

### Fixes
- When printing with fcMTblPrintParams.SpreadCols = TRUE and a merged cell appears on another page than the merge cell,<br>
   no more cells from the merge cell to the end of the page are printed

- MTblPrint doesn't consider invisible lines

- Under certain conditions, extended selections are not displayed properly

- When extended selections are used, hiding rows is quite slow and produces heavy flickering if the table must scroll

- Row hierarchy doesn't work properly anymore after all rows of the table are deleted ( and at least one was<br>
   deleted with TBL_Adjust ) and new rows are inserted. When SalTblReset is called before new rows are inserted,<br>
   there is no problem.

- MTblGetCellRectEx returns 0 for nTop and nBottom if the row isn't in the visible range.<br>
   This bug was introduced in 1.6.2 due to performance optimizing :o(

## V 1.6.2
- New functions
  - MTblCopyRowsAndHeaders<br>
   Copies rows and column headers into the clipboard

  - MTblSetSelectionColors/MTblGetSelectionColors<br>
   You may set/get the background and text color for selections now

  - MTblGetRowSelChanges<br>
   Used to get the changes of the row selection on MTM_RowSelChanged

  - MTblGetColSelChanges<br>
   Used to get the changes of the column selection on MTM_ColSelChanged

- New extended messages
  - MTM_RowSelChanged<br>
   Is sent when the selection of any row has changed
  - MTM_ColSelChanged<br>
   Is sent when the selection of any column has changed

- New table flags
  - MTBL_FLAG_FULL_OVERLAP_SELECTION<br>
   If set, the cells in the overlapping selection of rows and columns are also inverted or, if selection colors<br>
   are set, drawn with the selection colors

- New cell flags
  - MTBL_CELL_FLAG_NOSELINV_IMAGE<br>
   If set, the image of a selected cell isn't inverted
  - MTBL_CELL_FLAG_NOSELINV_TEXT<br>
   If set, the text of a selected cell isn't inverted
  - MTBL_CELL_FLAG_NOSELINV_BKGND<br>
   If set, the background of a selected cell isn't inverted

- New column flags
  - MTBL_COL_FLAG_NOSELINV_IMAGE<br>
   If set, the images of the column's selected cells aren't inverted
  - MTBL_COL_FLAG_NOSELINV_TEXT<br>
   If set, the texts of the column's selected cells aren't inverted
  - MTBL_COL_FLAG_NOSELINV_BKGND<br>
   If set, the backgrounds of the column's selected cells aren't inverted

- New row flags
  - MTBL_ROW_FLAG_NOSELINV_IMAGE<br>
   If set, the images of the row's selected cells aren't inverted
  - MTBL_ROW_FLAG_NOSELINV_TEXT<br>
   If set, the texts of the row's selected cells aren't inverted
  - MTBL_ROW_FLAG_NOSELINV_BKGND<br>
   If set, the backgrounds of the row's selected cells aren't inverted

- New tree flags
  - MTBL_TREE_FLAG_NOSELINV_NODES<br>
   If set, the nodes of the tree column's selected cells aren't inverted
  - MTBL_TREE_FLAG_NOSELINV_LINES<br>
   If set, the tree lines of the tree column's selected cells aren't inverted

- fcMTblODI
  - New members
    - Selected<br>
    Indicates whether an ownerdrawn item is selected ( = TRUE ) or not ( = FALSE )

## V 1.6.1
- Tree column mirrored in the row header
  - MTblObjFromPoint sets the flag MTOFP_OVERNODE now if the mouse is over the mirrored node
  - The row is expanded and/or collapsed now if the mirrored node is clicked

- MTblExportToExcel
  - When a table is exported with the export flag MTE_GRID, the row and column line definitions<br>
   as well as the flag TBL_Flag_SuppressRowLines and the tree flag MTBL_TREE_FLAG_NO_ROWLINES are considered now.

- MTblPrint
  - fcMTblPrintParams
    - New members
      - DocName<br>
     You may specify the document name now
      - NoPageNrsScale<br>
     If TRUE, the automatic page numbers aren't scaled
  - fcMTblPrintLine
    - New members
      - Flags
        - You may set the following flags now:
          - MTPL_NO_IMAGE_SCALE<br>
       If set, the line's images aren't scaled
          - MTPL_NO_TEXT_SCALE<br>
       If set, the line's texts ( incl. positioned texts ) aren't scaled  

- New functions
  - MTblCopyRows<br>
   Copies rows into the clipboard - more sophisticated than SalTblCopyRows, no 64 KB limit

  - MTblSetColHdrGrpFlags / MTblQueryColHdrGrpFlags<br>
   You may set/query the following column header group flags now:
    - MTBL_CHG_FLAG_NOINNERVERTLINES<br>
     If set, the inner vertical lines are omitted
    - MTBL_CHG_FLAG_NOINNERHORZLINE<br>
     If set, the inner horizontal line is omitted
    - MTBL_CHG_FLAG_NOCOLHEADERS<br>
     If set, the column headers are omitted

  - MTblSetExportSubstColor/MTblGetExportSubstColor/MTblClearExportSubstColors/MTblEnumExportSubstColors<br>
   You may substitute colors for exports ( e.g. Excel export ) now

### Fixes
- When a table is populated with TBL_FillAll, under certain conditions<br>
   the SQL error 202 and/or 203 occurs when rows are deleted

- When the flag TBL_Flag_SingleSelection is set, a row is clicked and the mouse is dragged, <br>
   the suppression of row selection ( flag MTBL_FLAG_SUPPRESS_ROW_SELECTION ) doesn't work if <br>
   the mouse isn't released over the focus row

- When the flag TBL_Flag_SingleSelection is set and the table is scrolled up by clicking a row and dragging the <br>
   mouse, the row focus frame isn't painted correctly

- When using themes under Win XP, the check boxes aren't painted correctly

- When setting the text of a merge-cell, the merged cells aren't repainted

- MTblPrint / MTBL_TREE_FLAG_NO_ROWLINES<br>
   Rowlines are printed in the tree column although the flag MTBL_TREE_FLAG_NO_ROWLINES is set

## V 1.6.0
- Fixes the TD3.1 bug #79776 ( column justify flags are removed if any other column flags are set )

- Improved paint performance, particularly with many visible columns and/or merged cells

- Under Windows XP, tooltips have a shadow ( if configured in the XP display settings )

- MTblSetFlags
  - New flag MTBL_FLAG_SUPPRESS_ROW_SELECTION<br>
   Suppresses row selection by user action and/or calls to SalTblSetRowFlags and/or SalTblSetFlagsAnyRows.

- New functions
  - MTblDefineTreeLines/MTblQueryTreeLines<br>
   You may set/query the tree line style and color now.

  - MTblGetEffRowHdrBackColor<br>
   Gets a row header's effective background color ( = currently displayed background color ).

  - MTblGetOwnerDrawnItem<br>
   Fills an fcMTblODI instance with information about an owner drawn item.

  - MTblIsOwnerDrawnItem<br>
   Determines whether an item is ownerdrawn

  - MTblSetColumnFlags / MTblQueryColumnFlags<br>
   You may set/query the following column flags now:
    - MTBL_COL_FLAG_END_ELLIPSE<br>
     Replaces part of the cell texts with ellipses, if necessary, so that the text fits in the cell.

  - MTblSetMsgOffset / MTblGetMsgOffset<br>
   You may set/get the offset of MTM messages now ( MTM_First = SAM_User + offset ).

  - MTblSetOwnerDrawnItem<br>
   Sets an item ( e.g. a cell ) ownerdrawn.

  - MTblSetTreeNodeColors / MTblGetTreeNodeColors<br>
   You may set/get the tree node's outer, inner and background color now.

  - MTblSetTreeFlags / MTblQueryTreeFlags<br>
   You may set/query the following tree flags now:
    - MTBL_TREE_FLAG_INDENT_ALL<br>
     Indents the tree column cells of all level 0 rows, regardless whether the MTBL_ROW_CANEXPAND flag is set or not.
    - MTBL_TREE_FLAG_NO_ROWLINES<br>
     No rowlines are painted in the tree column
    - MTBL_TREE_FLAG_BOTTOM_UP<br>
     The normal row's tree order is bottom-up
    - MTBL_TREE_FLAG_SPLIT_BOTTOM_UP<br>
     The split row's tree order is bottom-up

- New classes
  - fcMTblODI<br>
   Information about an owner drawn item

- New messages
  - MTM_DrawItem<br>
   This message is posted when an ownerdrawn item ( e.g. a cell ) must be drawn.
  - MTM_QueryBestCellWidth<br>
   This message is posted when M!Table has to determine the best cell width.
  - MTM_QueryBestColHdrWidth<br>
   This message is posted when M!Table has to determine the best column header width.

### Fixes
- Focus row rectangle<br>
   Selecting a column while a table is in editing mode leads to a wrong painted focus row rectangle

- Check box cells<br>
   The check box is always painted centered, even if the column justify is left or right.

- MTblSort<br>
   Under certain conditions, the sort result of formatted columns is wrong.<br>
   For example, this occurs when a Date/Time column has the format 'yyyy'.

- MTblSetCellBackColor / MTblSetColumnBackColor / MTblSetRowBackColor:<br>
   Even if MTSC_REDRAW is specified, the background color of the focus cell isn't changed

- Permeable tooltips doesn't work properly

- Transparent tooltips doesn't work if the fade in time is 0

## V 1.5.0 Beta2
- Under certain conditions, MTblAutoSizeColumn makes columns 1 pixel too narrow

- Column flag COL_IndicateOverflow doesn't work

- MTblObjFromPoint doesn't return MTOFP_OVERCELLTEXT although the cursor is over a cell's text.<br>
   This only happens with multiline cells.

## V 1.5.0 Beta1
- MTblObjFromPoint wrongly returns merged-flags ( e.g. MTOFP_OVERCELLTEXT_MERGED ) although the cell isn't merged

- Problems when you access the table on SAM_Destroy ( missing values, GPFs )

- Up to 50% performance loss in message processing after subclassing with MTblSubClass

- MTblPrint prints the text of invisible columns

- MTblAutoSizeColumn doesn't work correctly with invisible columns

## V 1.5.0
- Ready for TD 3.1<br>
  Please note that the extended table GUI isn't supported yet!

- New functions
  - MTblSetRowHdrBackColor / MTblGetRowHdrBackColor<br>
   You may set/get the row header background color now.

- MTblSetFlags
  - New flags:
    - MTBL_FLAG_EXPAND_CRLF<br>
     Expands line breaks ( carriage return / line feed ) in cell texts instead of displaying a special character,<br>
     regardless whether LinesPerRow < 2 and/or a column's WordWrap property is not set.

## V 1.5.0 Beta 2
- Added some documentation in mtbl.apl so that the basic function and/or parameter descriptions<br>
  are displayed along with the active coding assistant.<br>
  Thanks to Gerhard Achrainer for his effort!

- New functions
  - MTblEnableTipType / MTblIsTipTypeEnabled<br>
   Every tip type may be enabled separately now.

  - MTblGetMergeCol<br>
   Retrieves the handle of a merged cell's merge column.

  - MTblSetColumnHdrImage / MTblGetColumnHdrImage / MTblQueryColumnHdrImageFlags<br>
   You may use images in column headers now.

  - MTblSetTipFadeInTime / MTblGetTipFadeInTime<br>
   You may set/get a tip fade in time now.<br>
   Please note that this feature works only with Win2k or greater.

  - MTblSetTipOpacity / MTblGetTipOpacity<br>
   You may set/get a tip opacity ( alpha transparency ) now.<br>
   Please note that this feature works only with Win2k or greater.

- MTblExportToExcel
  - New export flag MTE_SPLIT_ROWS<br>
   If set, split rows ( if any ) are also exported

- MTblPrint
  - fcMTblPrintParams<br>
   New grid types MTP_GT_STANDARD2, MTP_GT_STANDARD3 and MTP_GT_STANDARD4.<br>
   See documentation for further details.
  - fcMTblPrintLine<br>
   New members LeftImage, CenterImage and RightImage

- Tip types
  - New type MTBL_TIP_COMPLETETEXT<br>
   Shows the complete cell text in the tip if the text doesn't fit into the cell

- Tip-Flags
  - New flags
    - MTBL_TIP_FLAG_PERMEABLE<br>
    Makes a tip "permeable"
  - Removed flags
    - MTBL_TIP_FLAG_HIDEONHOVER

## V 1.5.0 Beta 1
- Cell merge<br>
  As long as a merge cell is disabled, either because the column, the row or the cell is disabled, the merged cells are disabled, too.

- New functions
  - MTblSetCellFlags / MTblQueryCellFlags<br>
   You may set/query the following cell flags now:
    - MTBL_CELL_FLAG_HIDDEN_TEXT<br>
     Hides the text of a cell

  - MTblSetPasswordChar / MTblGetPasswordChar<br>
   You may set/get the password character of invisible columns now

  - MTblSetTip* / MTblGetTip*<br>
   You may set/get a tooltip for cells (text, image and button separately), column headers, column header groups, the corner and row headers now

- New messages
  - MTM_ShowCellTip, MTM_ShowCellTextTip, MTM_ShowCellImageTip, MTM_ShowCellButtonTip,<br>
   MTM_ShowColHdrTip, MTM_ShowColHdrGrpTip,<br>
   MTM_ShowCornerTip, MTM_ShowRowHdrTip<br>
   These messages are posted when an appropriate tooltip could be displayed

- MTblDefineTree / MTblQueryTree
  - New parameter "nIndent"<br>
   Let you specify / query the indentation

- MTblExportToExcel
  - New flag MTE_EXCEL_NULLS_AS_SPACE<br>
   If this flag is set, null values ( empty cells ) are exported as a space character.<br>
   This is the proceeding of M!Table versions up to 1.3.3a.

- MTblObjFromPoint
  - New flag MTOFP_OVERMERGEDCELL<br>
    This flag is set when the mouse is over a merged cell
  - Removed flags MTOFP_OVERCELLTEXT_MERGED, MTOFP_OVERCELLIMAGE_MERGED and MTOFP_OVERNODE_MERGED<br>
    Because of the new flag MTOFP_OVERMERGEDCELL these flags are obsolete.

- MTblPrint
  - New error code MTP_ERR_NOT_SUBCLASSED

## V 1.4.0
- New functions:
  - MTblDefineButtonCell / MTblQueryButtonCell<br>
   You may define/query buttons cell by cell now

  - MTblDefineHyperlinkHover / MTblQueryHyperlinkHover<br>
   You may define/query a HOVER effect for hyperlinks now

  - MTblPrintGetDefPrinterName<br>
   Retrieves the name of the current default printer

  - MTblPrintSetupDlg<br>
   Displays the print setup dialog for a printer

  - MTblSetCornerImage / MTblGetCornerImage / MTblQueryCornerImageFlags<br>
   All you need for a corner image

  - MTblQueryCellImageFlags / MTblQueryCellImageExpFlags / MTblQueryEffCellImageFlags<br>
   Queries a cell image's flags

- MTblSet*Image*
  - New flags:
    - MTSI_ALIGN_LEFT, MTSI_ALIGN_HCENTER, MTSI_ALIGN_RIGHT:<br>
     Horizontal alignment
    - MTSI_ALIGN_TOP, MTSI_ALIGN_VCENTER, MTSI_ALIGN_BOTTOM, MTSI_ALIGN_VTEXT<br>
     Vertical alignment
    - MTSI_ALIGN_TILE<br>
     Image is tiled
    - MTSI_NOLEAD_LEFT, MTSI_NOLEAD_TOP<br>
     No left/top leading

- MTblDefineButton*
  - New parameter "nImage"<br>
   Let you specify a button image 
  - New flag MTBL_BTN_DISABLED<br>
   Disables a button

- MTblSort<br>
  Improved sort performance ( three times faster )

- MTblSetFlags
  - New flags:
    - MTBL_FLAG_COLOR_ENTIRE_ROW<br>
     Paints the background color of rows over the entire table instead of painting only in column cells
    - MTBL_FLAG_COLOR_ENTIRE_COLUMN<br>
     Paints the background color of columns over the entire table instead of painting only in row cells

- MTblSetRowFlags / MTblPrint
  - New flag MTBL_ROW_PAGEBREAK<br>
   When MTblPrint finds a row with this flag, the next rows ( if any ) are printed on a new page

- MTblExportToExcel<br>
  New error code MTE_ERR_NOT_SUBCLASSED

## V 1.3.3a
- MTblExportToExcel
  - New flag MTE_EXCEL_STRING_COLS_AS_TEXT<br>
   Some people asked me to make it controllable whether string columns are exported as text.<br>
   Therefore this flag is introduced.

### Fixes
- The string column formats "Uppercase" and "Lowercase" don't work correctly

- GPF when the user scrolls to the last row in a table populated with TBL_FillNormal and the table is in editing mode

- MTblExportToExcel:<br>
   Row flags set on SAM_FetchRowDone are not recognized

- MTblSetCellTextColor / MTblSetColumnTextColor / MTblSetRowTextColor:<br>
   Even if MTSC_REDRAW is specified, the focus cell isn't redrawn

- MTblPrint:<br>
   Option "FitCellSize" doesn't work properly ( column is too narrow ) if the table font isn't a printable font ( e.g. MS Sans Serif ).<br>
   This occurs in particular when the font is bold or italic.

- MTblExportToExcel:<br>
   Cells of long string columns with a length > 254 aren't displayed correctly in Excel ( overflow '#####' )

## V 1.3.3a
- MTblPrint:<br>
   Changes of the default printer, e.g. by a call of SalPrtSetDefault, aren't recognized

- MTblExportToExcel:<br>
   When the table contains date values, the cells after the first date value possibly have the wrong data type

- MTblPrint:<br>
   Option "FitCellSize" doesn't work properly ( only the first row is recognized )

## V 1.3.3
- M!Image
  - M!Image was completely revised!<br>
   Please refer to the M!Image documentation.

- Row hierarchy
  - New functions
    - MTblDeleteDescRows
    - MTblIsRowDescOf

  - New messages
    - MTM_LoadChildRows
    - MTM_ExpandRowDone
    - MTM_CollapseRowDone

- MTblExportToExcel
  - String column cells are now exported as text instead of let Excel determine the data type.<br>
   This prevents e.g. that Excel interpretes "1/2" as "01-Feb".

- MTblSetFlags
  - New flags:
    - MTBL_FLAG_EXPAND_TABS<br>
     Expands tab characters in cell texts instead of displaying a special character.
    - MTBL_FLAG_EXPAND_ROW_ON_DBLCLK<br>
     In version 1.3.2 the automatic row expanding on double click was removed <br>
     and I didn't mentioned it here.<br>
     Use this flag to get the old behaviour.

### Fixes
- Under Win9x, you loose 1% system and gdi resources when a table is destroyed

- MTblPrint fails when printing directly to the printer with the "FitCellSize" parameter.<br>
   Everything works fine when printing from the preview window.<br>
   This defect was introduced in V 1.3.2.

- Dotted row lines are not painted correctly in columns which are wider than the table's client area


## V 1.3.2a
- Focus frame isn't painted correctly when the table is scrolled horizontally through cell tabbing

- Table receives no or not all SAM_FetchRow messages when populated with SalTblSetRange with nMax != TBL_MaxRow

## V 1.3.2
- Column header groups:
  - MTblCreateColHdrGrp / MTblDeleteColHdrGrp / MTblEnumColHdrGrps / MTblGetColHdrGrp / MTblEnumColHdrGrpCols
  - MTblSetColHdrGrpText / MTblGetColHdrGrpText / MTblSetColumnTitle / MTblGetColumnTitle
  - MTblSetColHdrGrpFont / MTblGetColHdrGrpFont / MTblGetEffColHdrGrpFont
  - MTblSetColHdrGrpTextColor / MTblGetColHdrGrpTextColor / MTblGetEffColHdrGrpTextColor
  - MTblSetColHdrGrpBackColor / MTblGetColHdrGrpBackColor / MTblGetEffColHdrGrpBackColor

- MTblExportToExcel:
  - New parameter sExcelStartCell

- MTblPrint:
  - Option FitCellSize:<br>
   If a column's Word Wrap property is set, only the height of the cell is adapted

### Fixes
- A splitted table without any split rows receives SAM_CacheFull when it's painted

- M!Table doesn't work with TD 1.1.x ( Message: "Not all required Centura functions could be loaded." )

- Popup edit and listbox have the table font instead of the cell font

- Sometimes the focus frame isn't painted correctly

## V 1.3.1
- New message MTM_Paint

### Fixes
- M!Table doesn't work under Win95 and WinNT

## V 1.3.0
- Ready for TD 3.0

- Custom fonts:
  - MTblSetCellFont / MTblGetCellFont / MTblGetEffCellFont
  - MTblSetColumnFont / MTblGetColumnFont
  - MTblSetColumnHdrFont / MTblGetColumnHdrFont / MTblGetEffColumnHdrFont
  - MTblSetRowFont / MTblGetRowFont

- Custom text colors:
  - MTblSetCellTextColor / MTblGetCellTextColor / MTblGetEffCellTextColor
  - MTblSetColumnTextColor / MTblGetColumnTextColor
  - MTblSetColumnHdrTextColor / MTblGetColumnHdrTextColor / MTblGetEffColumnHdrTextColor
  - MTblSetRowTextColor / MTblGetRowTextColor

- Custom background colors:
  - MTblSetCellBackColor / MTblGetCellBackColor / MTblGetEffCellBackColor
  - MTblSetColumnBackColor / MTblGetColumnBackColor
  - MTblSetColumnHdrBackColor / MTblGetColumnHdrBackColor / MTblGetEffColumnHdrBackColor
  - MTblSetRowBackColor / MTblGetRowBackColor

- Custom images:
  - New library M!Image
  - MTblSetCellImage / MTblGetCellImage
  - MTblSetCellImageExp / MTblGetCellImageExp
  - MTblSetRowFlagImage / MTblGetRowFlagImage

- Custom lines:
  - MTblDefineColLines / MTblQueryColLines
  - MTblDefineLastLockedColLine / MTblQueryLastLockedColLine
  - MTblDefineRowLines / MTblQueryRowLines

- Cell merging:
  - MtblSetCellMerge / MTblGetCellMerge

- MTblDefineButtonColumn / MTblQueryButtonColumn:
  - New parameter nFlags
  - New functions MTblDefineButtonColumn_121 / MTblQueryButtonColumn_121 for backward compatibility

- MTblObjFromPoint:
  - New flag MTOFP_OVERCELLIMAGE

- New functions MTblGetCellRectEx / MTblGetColHdrRectEx

- New functions MTblSetFlags / MTblQueryFlags

- New function MTBlFindPrevCol

- New function MTblQueryLockedColumns

### Fixes
- Some functions make a table visible, e.g. MTblExpandRow / MTblCollapseRow

## V 1.2.1
- New function MTblSortTree - does a hierarchical sort 

- New function MTblSortRange - sorts a range of rows

- New function MTblGetLastSortDef - retrieves the last sort definition

- New function MTblObjFromPoint - like SalTblObjectsFromPoint, but with enhanced flags

- New function MTblDefineHyperlinkCell - defines a cells hyperlink

- New function MTblQueryHyperlinkCell - retrieves hyperlink definition for a cell

- New functions MTblLockPaint / MTblUnlockPaint

- New message MTM_HyperlinkClick

- MTblExportToExcel now uses the Excel cell type STANDARD instead of TEXT

- On MTM_KeyDown you can now return MTBL_NODEFAULTACTION to prevent default processing

- New flags MTER_BY_USER / MTCR_BY_USER to indicate if the user has expanded / collapsed a row

### Fixes
- After a call to MTblExportToExcel, the text in the clipboard isn't the same as it was before

- If a column has no header text, MTblAutoSizeColumn makes a very huge width for the column

- In TD 2.1, on compiling an app you get the message "Could not load Centura environment".<br>
   This is due to a bug in mtexcexp.dll.

## V 1.2.0
- Ready for TD 2.1

- Excel export - exports your table with one function call to Excel

- Column buttons - displays a custom button in the focus cell of a column

- New sort data type "Custom" - tell M!Table how to sort your data

- New function MTblGetRowLevel

- New function MTblFindNextCol

- New functions MTblExpandRow / MTblCollapseRow

### Fixes
- If you press the up/down arrow in a multiline cell with more than 1 line per row,<br>
   the focus cell is changed instead of moving the cursor in the cell

- Mousewheel scrolling doesn't work when the table is in editing mode

## V 1.1.1
- New message MTM_KeyDown

- New function MTblGetVersion

### Fixes
- Defect #2 is still in CTD2000 version

- Positioned texts aren't printed with CTD 2000

- Columns and the table sometimes don't receive WM_KEYDOWN ( See new message MTM_KEYDOWN )

## V 1.1.0
- Row hierarchy - implements tree view functionality

- Print positioned texts and totals

- Disable rows

- Get the visible rectangle of cells and column headers

- Mousewheel scrolling

- MTM_SplitBar* extended messages

### Fixes
- If you press the up/down arrow in a drop down list column whose list is opened,<br>
   the focus cell is changed instead of selecting entries in the list

- Columns and the table don't receive SAM_Click when you change the focus cell<br>
   with the keyboard ( TAB, UP/DOWN-ARROW )
