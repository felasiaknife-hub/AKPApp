// VSPPGFlex.cpp : Implementation of CVSPPGFlex
#include "stdafx.h"
#include "VSPPG7.h"
#include "VSPPGFlex.h"
#include "PixTwip.h"


/////////////////////////////////////////////////////////////////////////////
// CVSPPGFlex: global info

// color for "hidden" columns
#define COLOR_HIDDEN ((long)GetSysColor(COLOR_APPWORKSPACE))

LPCTSTR g_szAlignmentSettings[] =
{
	"0 - LeftTop",
	"1 - LeftCenter",
	"2 - LeftBottom",
	"3 - CenterTop",
	"4 - CenterCenter",
	"5 - CenterBottom",
	"6 - RightTop",
	"7 - RightCenter",
	"8 - RightBottom",
	"9 - General",
	NULL
};

LPCTSTR g_szFormatSettings[] =
{
	"",
	"#,###.##",
	"#,###",
	"#.##%",
	"Currency",
	"Short Date",
	"Medium Date",
	"Long Date",
	NULL
};

LPCTSTR g_szTypeSettings[] =
{
	"0 - (any)",
	"2 - Short",
	"3 - Long",
	"4 - Single",
	"5 - Double",
	"6 - Currency",
	"7 - Date",
	"8 - String",
	"11 - Boolean",
	NULL
};

LPCTSTR g_szBooleanSettings[] =
{
	"True",
	"False",
	NULL
};

typedef struct tagCOLPROPINFO {
	LPCTSTR  szName;
	LPCTSTR* szList;
	LPCTSTR  szHelp;
} COLPROPINFO;

COLPROPINFO g_ColInfo[] =
{
	NULL,				NULL,					NULL,
	"Header",			NULL,					"Header text. Double-click the cell to edit it on the preview control.",
	"Key",				NULL,					"Key to use with ColIndex to retrieve the column's index.",
	"Width",		    NULL,					"Column width, in twips. You may resize columns on the preview control.",
	"Alignment",	    g_szAlignmentSettings,	"Alignment of cells in scrollable area.",
	"FixedAlignment",	g_szAlignmentSettings,	"Alignment of cells in fixed area.",
	"Format",			g_szFormatSettings,		"Format string (e.g. '#,###.##', type into the preview control to test).",
	"EditMask",			NULL,				    "Editing mask (e.g. '999-9999', type into the preview control to test).",
	"ComboList",	    NULL,					"List for drop-down lists/combos (e.g. 'One|Two|Three').",
	"DataType",			g_szTypeSettings,	    "Set to Boolean to get check-boxes, to Date for proper date sorting.",
	"Hidden",			g_szBooleanSettings,    "Hide this column at run time.",
	"Indent",			NULL,					"Indentation for the column's contents (in twips).",
	NULL,               NULL,                   NULL,
};


/////////////////////////////////////////////////////////////////////////////
// CVSPPGFlex: interface implementation

// prevent ESC key from dismissing dialog while we're editing
HRESULT CVSPPGFlex::TranslateAccelerator(MSG *pMsg)
{
	// disable VK_ESCAPE key while editing grid values
	if (pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		// if editing grid, pass the message to the grid and hide the message from the dialog
		CWindow wnd = (HWND)m_spGrid->EditWindow;
		if (wnd != NULL)
		{
			wnd.SendMessage(WM_KEYDOWN, VK_ESCAPE);
			pMsg->wParam = 0;
			return S_OK;
		}

		// if editing preview grid, pass the message to the grid and hide the message from the dialog
		wnd = (HWND)m_spGridPreview->EditWindow;
		if (wnd != NULL)
		{
			wnd.SendMessage(WM_KEYDOWN, VK_ESCAPE);
			pMsg->wParam = 0;
			return S_OK;
		}
	}

	// use default ATL implementation
	return IPropertyPageImpl<CVSPPGFlex>::TranslateAccelerator(pMsg);
}

// we have an object to edit: initialize preview control with its properties
HRESULT CVSPPGFlex::SetObjects(ULONG nObjects, IUnknown **ppUnk)
{
	// allow default processing
	HRESULT hr = IPropertyPageImpl<CVSPPGFlex>::SetObjects(nObjects, ppUnk);

	// we're clean
	SetDirty(false);

	// no grid yet? we're done
	if (m_spGridPreview == NULL) return hr;
	
	// refresh preview info
	InitGridPreview();

	// done
	return hr;
}

// apply changes to controls
HRESULT CVSPPGFlex::Apply()
{
	// sanity
	if (m_spGridPreview == NULL) return S_OK;

    // build new format string
    _bstr_t strFmt = "(Format)\v";

    // save column info
    _variant_t vVal = m_spGridPreview->Cols;
    strFmt += _bstr_t(vVal);
    strFmt += "\v";
    for (long c = 0; c < m_spGridPreview->Cols; c++) {

		// save ColWidth
		vVal = m_spGridPreview->GetColWidth(c);
        strFmt += _bstr_t(vVal);
		strFmt += "\t";

		// save ColAlignment
		vVal = (short)m_spGridPreview->GetColAlignment(c);
        strFmt += _bstr_t(vVal);
		strFmt += "\t";

		// save ColFixedAlignment
		vVal = (short)m_spGridPreview->GetFixedAlignment(c);
        strFmt += _bstr_t(vVal);
		strFmt += "\t";

		// save ColFormat
		strFmt += m_spGridPreview->GetColFormat(c);
		strFmt += "\t";

		// save ColEditMask
		strFmt += m_spGridPreview->GetColEditMask(c);
		strFmt += "\t";

		// save ColComboList
		strFmt += m_spGridPreview->GetColComboList(c);
		strFmt += "\t";

		// save ColDataType
		vVal = (short)m_spGridPreview->GetColDataType(c);
        strFmt += _bstr_t(vVal);
		strFmt += "\t";

		// save ColHidden
		vVal = m_spGridPreview->GetCell(flexcpBackColor, 0L, c, 0L, c);
        strFmt += ((long)vVal == COLOR_HIDDEN)? _bstr_t("-1"): _bstr_t("0");
		strFmt += "\t";

		// save ColKey
		strFmt += m_spGridPreview->GetColKey(c);
		strFmt += "\t";

		// save ColIndent
		vVal = m_spGridPreview->GetColIndent(c);
        strFmt += _bstr_t(vVal);
		strFmt += "\t";

		// mark column end
		strFmt += "\v";
	}

	// save contents
    strFmt += "(Text)\v";
    long lRows = CountRowsInGrid();
    for (long r = 0; r < lRows; r++) {
        long lCols = CountCellsInRow(r);
        for (long c = 0; c < lCols; c++) {
            strFmt += m_spGridPreview->GetTextMatrix(r, c);
    		strFmt += "\t";
        }
		strFmt += "\v";
    }

	// assign new FormatString to object
	SetCtlProp("Cols", _variant_t(m_spGridPreview->Cols));
	SetCtlProp("FormatString", _variant_t(strFmt));

	// we're clean and done
	SetDirty(false);
	return S_OK;
}


/////////////////////////////////////////////////////////////////////////////
// CVSPPGFlex: message handlers

// initialize dialog
LRESULT CVSPPGFlex::OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	// initialize font
	static FONTDESC _fdDefault = { sizeof(FONTDESC), OLESTR("MS Sans Serif"), FONTSIZE(8),
								   FW_NORMAL, DEFAULT_CHARSET, false, false, false };
	OleCreateFontIndirect(&_fdDefault, IID_IFont, (void**)&m_spFont);
	CComQIPtr<IFontDisp, &IID_IFontDisp> pFont(m_spFont);
	HFONT hFont;
	m_spFont->get_hFont(&hFont);

	// load button icons
	m_hIconClr = (HICON)LoadImage(_Module.GetResourceInstance(), MAKEINTRESOURCE(IDI_CLEAR), IMAGE_ICON, 0, 0, 0);
	m_hIconIns = (HICON)LoadImage(_Module.GetResourceInstance(), MAKEINTRESOURCE(IDI_INS),   IMAGE_ICON, 0, 0, 0);
	m_hIconDel = (HICON)LoadImage(_Module.GetResourceInstance(), MAKEINTRESOURCE(IDI_DEL),   IMAGE_ICON, 0, 0, 0);

	// initialize Clear button
	RECT rc, rcBtn;
	CWindow wndBtn = GetDlgItem(IDC_CLEAR);
    wndBtn.GetClientRect(&rcBtn);
	GetGridRect(&rc, false);
	rc.right = rc.left + rcBtn.right;
	wndBtn.MoveWindow(&rc);
	wndBtn.SendMessage(BM_SETIMAGE, IMAGE_ICON, (LPARAM)m_hIconClr);

	// initialize insert/delete row buttons
	wndBtn = GetDlgItem(IDC_COLINS);
	OffsetRect(&rc, rcBtn.right, 0);
	rc.bottom = rc.top + (rc.bottom - rc.top) / 2;
	wndBtn.MoveWindow(&rc);
	wndBtn.SendMessage(BM_SETIMAGE, IMAGE_ICON, (LPARAM)m_hIconIns);
	wndBtn = GetDlgItem(IDC_COLDEL);
	OffsetRect(&rc, 0, rc.bottom - rc.top);
	wndBtn.MoveWindow(&rc);
	wndBtn.SendMessage(BM_SETIMAGE, IMAGE_ICON, (LPARAM)m_hIconDel);

	// initialize help window
	GetGridRect(&rc, false);
	rc.left = rc.left + 2 * rcBtn.right + BORDER;
	m_wndHelp.Attach(GetDlgItem(IDC_HELPTEXT));
	m_wndHelp.MoveWindow(&rc);
	m_wndHelp.SendMessage(WM_SETFONT, (WPARAM)hFont);

	// create grid controls
	if (S_OK != m_spGrid.CreateInstance(__uuidof(VSFlexGrid)) ||
		S_OK != m_spGridPreview.CreateInstance(__uuidof(VSFlexGrid))) {
		m_wndHelp.SetWindowText(MSG_NOFLEX);
		return -1;
	}

	// initialize ATL hosting
	AtlAxWinInit();

	// create ActiveX control host window for property editor
	GetGridRect(&rc);
	int iSaveRight = rc.right;
	rc.right = (rc.right + rc.left) / 3;
	m_wndHost.Create(m_hWnd, rc, NULL, STYLE_HOST, 0, IDC_GRID);

	// attach FlexGrid control to host window
	CComPtr<IAxWinHostWindow> spHost;
	m_wndHost.QueryHost(&spHost);
	spHost->AttachControl(m_spGrid, m_wndHost);

	// create ActiveX control host window for preview grid
	rc.left = rc.right + BORDER;
	rc.right = iSaveRight;
	m_wndHostPreview.Create(m_hWnd, rc, NULL, STYLE_HOST, 0, IDC_PREVIEW);

	// attach FlexGrid control to host window
	CComPtr<IAxWinHostWindow> spHostPreview;
	m_wndHostPreview.QueryHost(&spHostPreview);
	spHostPreview->AttachControl(m_spGridPreview, m_wndHostPreview);

	// connect events
	AtlAdviseSinkMap(this, true);

	// initialize grid control
	m_spGrid->PutFont(pFont);
	InitGrid();
	
	// initialize preview control
	m_spGridPreview->PutFont(pFont);
	InitGridPreview();
	
	// no double clicks yet
	m_bDblClick = false;

	// all done
	m_wndHost.ShowWindow(SW_SHOW);
	m_wndHostPreview.ShowWindow(SW_SHOW);
    m_wndHostPreview.SetFocus();
	return 0;
}

// done
LRESULT CVSPPGFlex::OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	// release grid controls
	if (m_spGrid) m_spGrid.Release();
	if (m_spGridPreview) m_spGridPreview.Release();

	// destroy button icons
	if (m_hIconClr) DestroyIcon(m_hIconClr);
	if (m_hIconIns) DestroyIcon(m_hIconIns);
	if (m_hIconDel) DestroyIcon(m_hIconDel);

	// allow default processing
	bHandled = false;
	return 0;
}

// show resizing cursor
LRESULT CVSPPGFlex::OnMouseMove(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	// bail out if not over resizing area
	RECT rc;
	GetGridRect(&rc);
	POINT pt = { LOWORD(lParam), HIWORD(lParam) };
	if (!PtInRect(&rc, pt)) return 0;

	// show resizing cursor
	SetCursor(LoadCursor(NULL, MAKEINTRESOURCE(IDC_SIZEWE)));
	return 0;
}

// resize controls
LRESULT CVSPPGFlex::OnMouseDown(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	// bail out if not over resizing area
	RECT rc;
	GetGridRect(&rc);
	POINT pt = { LOWORD(lParam), HIWORD(lParam) };
	if (!PtInRect(&rc, pt)) return 0;

	// prepare to resize
	HDC     hdc = ::GetDC(NULL);
	RECT    rcSize = rc;
	HCURSOR hCur = LoadCursor(NULL, MAKEINTRESOURCE(IDC_SIZEWE));
	BOOL    bErase  = false;
	BOOL    bResize = false;

	// track the mouse until the user releases it
	SetCapture();
	SetCursor(hCur);
    for (BOOL bDone = false; !bDone;) {
	    MSG msg;
		if (PeekMessage(&msg, NULL, WM_KEYFIRST, WM_MOUSELAST, PM_REMOVE)) {
            switch (msg.message) {

                // show resize bar as the mouse moves
                case WM_MOUSEMOVE:

					// show resizing cursor
					SetCursor(hCur);

					// get mouse position, adjust to make sure we're within the resizing area
					POINT ptMouse;
					ptMouse.x = (short)LOWORD(msg.lParam);
					ptMouse.y = (short)HIWORD(msg.lParam);
					if (ptMouse.x < rc.left  + BORDER) ptMouse.x = rc.left  + BORDER;
					if (ptMouse.x > rc.right - BORDER) ptMouse.x = rc.right - BORDER;

					// if the position hasn't changed, break to avoid flicker
					if (bErase && ptMouse.x == pt.x) break;
					pt.x = ptMouse.x;

					// erase old line, show new one
					if (bErase) InvertRect(hdc, &rcSize);
					rcSize = rc;
					rcSize.left = ptMouse.x - BORDER/2;
					rcSize.right = rcSize.left + BORDER;
					ClientToScreen(&rcSize);
					InvertRect(hdc, &rcSize);
					bErase = true;
					break;

                // quit if escape key is hit
                case WM_KEYDOWN:
                    if (msg.wParam == VK_ESCAPE) bDone = true;
                    break;

                // quit when the button is released
                case WM_LBUTTONUP:
                    bDone = true;
					bResize = bErase;
                    break;
            }
        }
    }

    // clean up
	if (bErase) InvertRect(hdc, &rcSize);
	::ReleaseDC(NULL, hdc);
	ReleaseCapture();

	// do the resizing
	if (bResize) {
		GetGridRect(&rc);
		rc.right = pt.x - BORDER / 2;
		m_wndHost.MoveWindow(&rc);
		GetGridRect(&rc);
		rc.left = pt.x + BORDER / 2;
		m_wndHostPreview.MoveWindow(&rc);
	}

	// done
	return 0;
}

// show tooltip
LRESULT CVSPPGFlex::OnSetCursor(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	HWND   hWnd  = (HWND)wParam;
    LPCSTR szTip = "";

    // get help text for Clear button
    if (hWnd == GetDlgItem(IDC_CLEAR)) {
        szTip = "Clear all formatting (cannot undo).";

    // get help text for insert/delete column buttons
    } else if (hWnd == GetDlgItem(IDC_COLINS)) {
        szTip = "Insert a column before the selected column.";
    } else if (hWnd == GetDlgItem(IDC_COLDEL)) {
        szTip = "Delete all selected columns.";

    // get help text for property window
    } else if (hWnd == (HWND)m_spGrid->hWnd) {
        long iRow = m_spGrid->MouseRow;
        long iCol = m_spGrid->MouseCol;
        if (iCol > -1 && iRow > 0)
            szTip = g_ColInfo[iRow].szHelp;

    // get help text for preview window
    } else if (hWnd == (HWND)m_spGridPreview->hWnd) {
        long iRow = m_spGridPreview->MouseRow;
        long iCol = m_spGridPreview->MouseCol;
        if (iRow > -1 && iCol > -1) {
            szTip = (iRow < m_spGridPreview->FixedRows)
                ? "Double-click to edit cell or drag to move column."
                : "Double-click to edit cell (shift-enter inserts line breaks).";
        }
    }

    // build help string
    static TCHAR szHelp[200], szLastHelp[200];
	long   r1, c1, r2, c2;
	m_spGridPreview->GetSelection(&r1, &c1, &r2, &c2);
    if (c1 == c2) wsprintf(szHelp, "Column %d\r\n%s", c1, szTip);
    else		  wsprintf(szHelp, "Columns %d to %d\r\n%s", c1, c2, szTip);

    // display the text if it has changed
    if (_tcscmp(szHelp, szLastHelp))
	    m_wndHelp.SetWindowText(szHelp);

    // save last to avoid flicker
    _tcscpy(szLastHelp, szHelp);

    // done
	return 0;
}

// reset control by setting FormatString to an empty string
LRESULT CVSPPGFlex::OnClickedClear(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
    // make sure user wants to do this
    int iConfirm = MessageBox(
        "Do you really want to clear all row, column, and cell formatting?\n\n"
        "You will not be able to Cancel or Undo this operation.",
        "Please Confirm", MB_YESNO | MB_ICONQUESTION);
    if (iConfirm != IDYES) return 0;

	// clear preview grid
    m_spGridPreview->FormatString = "";
    m_spGridPreview->Clear();
    m_spGridPreview->PutRowHeight(-1, -1);
    m_spGridPreview->PutColWidth(-1, -1);
    m_spGridPreview->PutColHidden(-1, false);
    OnSelChangePV();

    // clear control being edited
	SetCtlProp("FormatString", _variant_t(""));

    // we're really clean now
	SetDirty(false);
	return 0;
}

// insert a column
LRESULT CVSPPGFlex::OnClickedInsert(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	// insert column
    m_spGridPreview->Cols = m_spGridPreview->Cols + 1;
    m_spGridPreview->PutColPosition(m_spGridPreview->Cols - 1, m_spGridPreview->Col);

	// done
	m_wndHostPreview.SetFocus();
	SetDirty(true);
	OnSelChangePV();
	return 0;
}

// delete a column
LRESULT CVSPPGFlex::OnClickedDelete(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	// need at least one column // <<B5>>
	if (m_spGridPreview->Cols < 1) return 0;

	// get selection
	long r1, c1, r2, c2;
	m_spGridPreview->GetSelection(&r1, &c1, &r2, &c2);

    // make sure user wants to do this
	TCHAR szMsg[100];
	if (c1 == c2) wsprintf(szMsg, "Do you really want to delete column %d?", c1);
	else		  wsprintf(szMsg, "Do you really want to delete columns %d through %d?", c1, c2);
    int iConfirm = MessageBox(szMsg, "Please Confirm", MB_YESNO | MB_ICONQUESTION);
    if (iConfirm != IDYES) return 0;
	
	// delete selected columns
	for (long c = c1; c <= c2; c++)
	    m_spGridPreview->PutColPosition(c1, m_spGridPreview->Cols - 1);
	m_spGridPreview->Cols = m_spGridPreview->Cols - (c2 - c1 + 1);	

	// done
	m_wndHostPreview.SetFocus();
	SetDirty(true);
	OnSelChangePV();
	return 0;
}


/////////////////////////////////////////////////////////////////////////////
// CVSPPGFlex: internal methods

// initialize property editor grid
void CVSPPGFlex::InitGrid()
{
	// sanity
	if (m_spGrid == NULL) return;

	// freeze redrawing to avoid flicker
    m_spGrid->Redraw = flexRDNone;							
    
	// initialize grid
	m_spGrid->Rows = 1;								        // start empty
    m_spGrid->Cols = 2;										// property, value
    m_spGrid->PutTextMatrix(0, 0, L"Property");				// column titles
    m_spGrid->PutTextMatrix(0, 1, L"Value");				// column titles
    m_spGrid->PutColAlignment(-1, flexAlignLeftTop);		// align all columns to the left
    m_spGrid->Editable          = flexEDKbd;				// editable, handle double-clicks ourselves
    m_spGrid->AllowSelection    = false;					// select a single cell at a time
    m_spGrid->AllowUserResizing = flexResizeColumns;		// give user freedom
    m_spGrid->ScrollBars        = flexScrollBarVertical;	// can't scroll horizontally
    m_spGrid->FixedCols         = 0;						// to look nice
    m_spGrid->ExtendLastCol     = true;
	m_spGrid->Ellipsis          = flexEllipsisEnd;
    m_spGrid->BackColorBkg      = m_spGrid->GridColor;
    m_spGrid->HighLight         = flexHighlightNever;

	// populate grid
	for (long iRow = 1; g_ColInfo[iRow].szName; iRow++) {
		m_spGrid->Rows = iRow + 1;
		m_spGrid->PutTextMatrix(iRow, 0, g_ColInfo[iRow].szName);
	}
	m_spGrid->PutColWidth(0, m_spGrid->ClientWidth / 2);
	m_spGrid->Select(1, 0);

	// turn redraw back on
    m_spGrid->Redraw = flexRDBuffered;
}

// initialize preview grid
void CVSPPGFlex::InitGridPreview()
{
	// set propperties that are independent of control being edited
	m_spGridPreview->Editable          = flexEDKbdMouse;		// allow editing
	m_spGridPreview->ExplorerBar       = flexExMove;			// user can move columns
	m_spGridPreview->AllowUserResizing = flexResizeColumns;		// user can resize columns
	m_spGridPreview->AllowBigSelection = false;		            // no big selections
	m_spGridPreview->ScrollTrack       = true;					// allow mouse tracking

	// get pointer to control being edited
	// use a dispatch pointer to handle every version (ADO, DAO, Light, UNICODE)
	CComQIPtr<IDispatch, &IID_IDispatch> pDisp(m_ppUnk[0]);
	if (pDisp == NULL) return;
	m_spDisp = pDisp;

	// copy global info from control being edited to preview control
	m_spGridPreview->PutRefFont(GetCtlPropFont());
	m_spGridPreview->BackColor    = GetCtlPropLong("BackColor");
	m_spGridPreview->ForeColor    = GetCtlPropLong("ForeColor");
	m_spGridPreview->Rows         = GetCtlPropLong("Rows");
	m_spGridPreview->Cols         = GetCtlPropLong("Cols");
	m_spGridPreview->FixedRows    = GetCtlPropLong("FixedRows");
	m_spGridPreview->FixedCols    = GetCtlPropLong("FixedCols");
	m_spGridPreview->WordWrap     = (short)GetCtlPropLong("WordWrap");

	// copy format string (to get fixed col headers)
	m_spGridPreview->FormatString = GetCtlPropStr("FormatString");

	// copy column info (in case FormatString missed something, which it shouldn't)
	for (long iCol = 0; iCol < m_spGridPreview->Cols; iCol++) {

		// column properties
		m_spGridPreview->PutColKey        (iCol, GetCtlPropStr  ("ColKey",        iCol));
		m_spGridPreview->PutColWidth      (iCol, GetCtlPropLong ("ColWidth",      iCol));
		m_spGridPreview->PutColFormat     (iCol, GetCtlPropStr  ("ColFormat",     iCol));
		m_spGridPreview->PutColEditMask   (iCol, GetCtlPropStr  ("ColEditMask",   iCol));
		m_spGridPreview->PutColComboList  (iCol, GetCtlPropStr  ("ColComboList",  iCol));
		m_spGridPreview->PutColAlignment  (iCol, (AlignmentSettings)GetCtlPropLong("ColAlignment",  iCol));
		m_spGridPreview->PutFixedAlignment(iCol, (AlignmentSettings)GetCtlPropLong("FixedAlignment",iCol));
		m_spGridPreview->PutColDataType   (iCol, (DataTypeSettings) GetCtlPropLong("ColDataType",   iCol));
		m_spGridPreview->PutColIndent     (iCol, GetCtlPropLong ("ColIndent",  iCol));

		// these require at least one row
		if (m_spGridPreview->Rows > 0) {

			// header text
			m_spGridPreview->PutTextMatrix (0, iCol, GetCtlPropStr("TextMatrix", 0, iCol));

			// we don't hide the preview column: change the color of the top cell instead
			_variant_t vColor = GetCtlPropLong("ColHidden", (iCol))? COLOR_HIDDEN: 0L;
			m_spGridPreview->PutCell(flexcpBackColor, 0L, iCol, 0L, iCol, vColor);
			m_spGridPreview->PutColHidden(iCol, false);
		}
	}

	// initalize selection
	// note: use raw_Select for preview grid because we could have no rows etc
	//       and we don't want to throw any exceptions here.
	m_spGrid->Select(1, 0);
	m_spGridPreview->raw_Select(m_spGridPreview->FixedRows, m_spGridPreview->FixedCols);

	// initialize property window values
	OnSelChangePV();
}

// measure font height
int CVSPPGFlex::GetFontHeight()
{
	TEXTMETRIC tm;
	HFONT      hFont;
	HDC        hdc = GetDC();
	m_spFont->get_hFont(&hFont);
	HFONT hOldFont = (HFONT)SelectObject(hdc, hFont);
	GetTextMetrics(hdc, &tm);
	SelectObject(hdc, hOldFont);
	ReleaseDC(hdc);
	return tm.tmHeight + tm.tmExternalLeading;
}

// get grid rectangle
void CVSPPGFlex::GetGridRect(LPRECT prc, BOOL bGrid)
{
	GetClientRect(prc);
	InflateRect(prc, -BORDER, -BORDER);

    // calculate font height
	int nFontHeight = GetFontHeight();

	// get grid rect (fits both grids)
	if (bGrid)
		prc->bottom -= (nFontHeight * 2 + BORDER);

	// get help window rect
	else
		prc->top = prc->bottom - nFontHeight * 2;
}

// count how many lines there are in a cell
int CVSPPGFlex::CountLinesInCell(long Row, long Col)
{
	// get entry
	_bstr_t strEntry = m_spGridPreview->GetTextMatrix(Row, Col);
	
	// count line breaks
	int nLines = 1;
	for (LPTSTR sz = strEntry; sz && *sz; sz = CharNext(sz))
		if (*sz == '\n') nLines++;

	// return what we got
	return nLines;
}

// count how many cells are in a row (trim empty cells)
int CVSPPGFlex::CountCellsInRow(long Row)
{
    for (long c = m_spGridPreview->Cols - 1; c >= 0; c--) {
        _bstr_t strCell = m_spGridPreview->GetTextMatrix(Row, c);
        if (strCell.length() > 0) return c + 1;
    }
    return 0;
}

// count how many rows are in a grid (trim empty rows)
int CVSPPGFlex::CountRowsInGrid()
{
    for (long r = m_spGridPreview->Rows - 1; r >= 0; r--) {
        if (CountCellsInRow(r)) return r + 1;
    }
    return 0;
}

LPTSTR CVSPPGFlex::InStr(LPTSTR str1, LPTSTR str2)
{
    if (!str1 || !*str1 || !str2 || !*str2) return NULL;
    int l = lstrlen(str2);
    for ( ; *str1; str1++)
        if (!_tcsnicmp(str1, str2, l)) return str1;
    return NULL;
}

// select next item for a cell with a ComboList, return true if successful
bool CVSPPGFlex::SelectNextItem(long Row, long Col)
{
	// get current list, trim combo pipe if any
	USES_CONVERSION;
	LPTSTR szList = W2T(m_spGrid->ComboList);
	if (!szList || !lstrlen(szList)) return false;
	if (*szList == '|') szList++;

    // look for current text in list, fail if not found
    LPTSTR szNext = InStr(szList, W2T(m_spGrid->GetTextMatrix(Row, Col)));
    if (!szNext) return false;
    
    // look for next choice
	szNext = _tcschr(szNext, '|');
	szNext = (szNext)? szNext + 1: szList;
    
    // trim excess
    LPTSTR szEnd = _tcschr(szNext, '|');
	if (szEnd) *szEnd = 0;
    
    // set new entry
    m_spGrid->PutTextMatrix(Row, Col, szNext);
	return true;
}


/////////////////////////////////////////////////////////////////////////////
// CVSPPG: event handlers for property editor

// DblClick: cycle through list items
void _stdcall CVSPPGFlex::OnDblClick()
{
	m_bDblClick = true;
	m_spGrid->EditCell();
}

// KeyDown: search for property
void _stdcall CVSPPGFlex::OnKeyDown(short* KeyCode, short Shift)
{
    // when pushing control+ASCII, look for property
    if (Shift < 2 || *KeyCode < (short)'A' || *KeyCode > (short)'Z') return;
            
    // look from current row down to bottom
	USES_CONVERSION;
	long r;
    for (r = m_spGrid->Row + 1; r < m_spGrid->Rows; r++) {
		LPTSTR sz = W2T(m_spGrid->GetTextMatrix(r, 0));
		if (sz && toupper(*sz) == *KeyCode) {
			m_spGrid->Select(r, 1);
            m_spGrid->ShowCell(r, 1);
            *KeyCode = 0;
			return;
		}
	}
    
    // not found, so look from top down to current - 1
    for (r = m_spGrid->FixedRows; r < m_spGrid->Row; r++) {
		LPTSTR sz = W2T(m_spGrid->GetTextMatrix(r, 0));
		if (sz && toupper(*sz) == *KeyCode) {
			m_spGrid->Select(r, 1);
            m_spGrid->ShowCell(r, 1);
            *KeyCode = 0;
			return;
		}
	}
}

// BeforeRowColChange: only the value column can be selected
void _stdcall CVSPPGFlex::OnBeforeRowColChange(long OldRow, long OldCol, long NewRow, long NewCol, VARIANT_BOOL* Cancel)
{
	if (m_spGrid->Redraw && NewCol != 1) {
		*Cancel = true;
		m_spGrid->Select(NewRow, 1);
	}
}

// BeforeEdit: set ComboList
void _stdcall CVSPPGFlex::OnBeforeEdit(long Row, long Col, VARIANT_BOOL* Cancel)
{
	// prevent editing property name
	if (Col == 0) {
		*Cancel = true;
		return;
	}

	// build ComboList string
	_bstr_t bstrList;
	LPCTSTR* szList = g_ColInfo[Row].szList;
	if (szList) {
		int i;
		for (i = 0; szList[i]; i++) {
			if (i > 0) bstrList += "|";
			bstrList += szList[i];
		}
	}

	// set ComboList string
	m_spGrid->ComboList = bstrList;

    // use automatic double-click for editing text, manual for lists
	m_spGrid->Editable = (bstrList.length())? flexEDKbd: flexEDKbdMouse;
}

// StartEdit: if editing lists with double-click, cycle list values
void _stdcall CVSPPGFlex::OnStartEdit(long Row, long Col, VARIANT_BOOL* Cancel)
{
    // if this is a list, double-clicking selects the next item
    if (m_bDblClick && SelectNextItem(Row, Col)) {
		*Cancel = true;
		OnAfterEdit(Row, Col);
	}

	// done
	m_bDblClick = false;
}

// AfterEdit: apply new settings, mark as dirty
void _stdcall CVSPPGFlex::OnAfterEdit(long Row, long Col)
{
	// get selected range
	long c, r1, c1, r2, c2;
	m_spGridPreview->GetSelection(&r1, &c1, &r2, &c2);

	// get new property value (just edited)
	_bstr_t bstrVal = m_spGrid->GetTextMatrix(Row, 1);
	LPTSTR  szVal   = bstrVal;
	long    iVal    = (szVal && *szVal >= '0' && *szVal <= '9')? _ttol(szVal): -1;

	// apply changes to preview control
	// use raw calls to avoid potential exceptions
	switch (Row) {
		case 1: // Header
			for (c = c1; c <= c2; c++)
				m_spGridPreview->put_TextMatrix(0, c, bstrVal);
			OnAfterEditPV(0, 0); // auto resize fixed cells after editing header
			break;
		case 2: // Key
			for (c = c1; c <= c2; c++)
				m_spGridPreview->put_ColKey(c, bstrVal);
			break;
		case 3: // Width (refuse negative and non-numeric values)
			if (iVal < 0) break;
			for (c = c1; c <= c2; c++)
				m_spGridPreview->put_ColWidth(c, iVal);
			break;
		case 4: // Alignment
			for (c = c1; c <= c2; c++)
				m_spGridPreview->put_ColAlignment(c, (AlignmentSettings)iVal);
			break;
		case 5: // FixedAlignment
			for (c = c1; c <= c2; c++)
				m_spGridPreview->put_FixedAlignment(c, (AlignmentSettings)iVal);
			break;
		case 6: // Format
			for (c = c1; c <= c2; c++)
				m_spGridPreview->put_ColFormat(c, bstrVal);
			break;
		case 7: // EditMask
			for (c = c1; c <= c2; c++)
				m_spGridPreview->put_ColEditMask(c, bstrVal);
			break;
		case 8: // ComboList
			for (c = c1; c <= c2; c++)
				m_spGridPreview->put_ColComboList(c, bstrVal);
			break;
		case 9: // DataType
			for (c = c1; c <= c2; c++)
				m_spGridPreview->put_ColDataType(c, (DataTypeSettings)iVal);
			break;
		case 10: // Hidden (we don't hide the preview column: change the color of the top cell instead)
			for (c = c1; c <= c2; c++) {
				_variant_t vColor = (lstrcmpi(bstrVal, "True") == 0)? COLOR_HIDDEN: 0L;
				m_spGridPreview->PutCell(flexcpBackColor, 0L, c, 0L, c, vColor);
			}
			break;
		case 11: // Indent
			for (c = c1; c <= c2; c++)
				m_spGridPreview->put_ColIndent(c, iVal);
			break;
	}

	// refresh to make sure changes took
	OnSelChangePV();

	// remember we're dirty
	SetDirty(TRUE);
}

/////////////////////////////////////////////////////////////////////////////
// CVSPPG: event handlers for preview grid

// SelChange: update property window values
void _stdcall CVSPPGFlex::OnSelChangePV()
{
	// <<B6>>: get column being edited, bail if there are no columns
	long iCol  = m_spGridPreview->Col;
	long iCols = m_spGridPreview->Cols;
	long iRows = m_spGridPreview->Rows;
	if (iCol < 0 || iCols < 1 || iCol >= iCols) return;

	// update help text
	BOOL bHandled = false;
	OnSetCursor(0, (WPARAM)m_spGridPreview->hWnd, 0, bHandled);

	// update data in property window
	_variant_t vVal;
	long iRow = 1;

	// header text // <<B6>>: if there are fixed rows
	m_spGrid->PutTextMatrix(iRow++, 1, (iRows)? m_spGridPreview->GetTextMatrix(0, iCol): _T(""));

	// column key
	m_spGrid->PutTextMatrix(iRow++, 1, m_spGridPreview->GetColKey(iCol));

	// column width
	vVal = m_spGridPreview->GetColWidth(iCol);
	m_spGrid->PutTextMatrix(iRow++, 1, _bstr_t(vVal));

	// alignment and fixed alignment
	int i = m_spGridPreview->GetColAlignment(iCol);
	m_spGrid->PutTextMatrix(iRow++, 1, _bstr_t(g_szAlignmentSettings[i]));
	i = m_spGridPreview->GetFixedAlignment(iCol);
	m_spGrid->PutTextMatrix(iRow++, 1, _bstr_t(g_szAlignmentSettings[i]));

	// column format, edit mask, and combo list
	m_spGrid->PutTextMatrix(iRow++, 1, m_spGridPreview->GetColFormat(iCol));
	m_spGrid->PutTextMatrix(iRow++, 1, m_spGridPreview->GetColEditMask(iCol));
	m_spGrid->PutTextMatrix(iRow++, 1, m_spGridPreview->GetColComboList(iCol));

	// data type (the numbering in this case is not sequential)
	i = m_spGridPreview->GetColDataType(iCol);
	for (int j = 0; g_szTypeSettings[j]; j++)
		if (_ttol(g_szTypeSettings[j]) == i)
			m_spGrid->PutTextMatrix(iRow++, 1, _bstr_t(g_szTypeSettings[j]));

	// column hidden (we don't hide the preview column: change the color of the top cell instead)
	if (m_spGridPreview->Rows) {
		_variant_t vColor = m_spGridPreview->GetCell(flexcpBackColor, 0L, iCol, 0L, 0L);
		i = V_I4(&vColor)? 0: 1; // if set to non-zero value, we're hidden
		m_spGrid->PutTextMatrix(iRow++, 1, _bstr_t(g_szBooleanSettings[i]));
	}
	
	// column indent
	vVal = m_spGridPreview->GetColIndent(iCol);
	m_spGrid->PutTextMatrix(iRow++, 1, _bstr_t(vVal));
}

// DblClick: start editing fixed cells
void _stdcall CVSPPGFlex::OnDblClickPV()
{
	// get cell that is being pointed to
	long r = m_spGridPreview->MouseRow;
	long c = m_spGridPreview->MouseCol;

	// off the sheet or over scrollable area? bail out
	if (r < 0 || c < 0) return;
	if (r >= m_spGridPreview->FixedRows && c >= m_spGridPreview->FixedCols) return;

	// start editing
	m_spGridPreview->Select(r, c);
	m_spGridPreview->EditCell();
}

// AfterEdit: autosize fixed row height
void _stdcall CVSPPGFlex::OnAfterEditPV(long Row, long Col)
{
	// resize row to accommodate new entry
	long lMaxLines = 1;
	for (long c = 0; c < m_spGridPreview->Cols; c++) {
		long iLines = CountLinesInCell(Row, c);
		if (iLines > lMaxLines) lMaxLines = iLines;
	}

	// resize row to accommodate new entry
	m_spGridPreview->PutRowHeight(Row, -1);
	long lPadding = YPixelsToTwips(3);
	long lSingleLineHeight = m_spGridPreview->GetRowHeight(Row) - lPadding;
	m_spGridPreview->PutRowHeight(Row, (lMaxLines < 2)? -1: lMaxLines * lSingleLineHeight + lPadding);

	// refresh to display new Header text in property window
	OnSelChangePV();

    // remember we're dirty
	SetDirty(TRUE);
}

// AfterUserResize: show new width value
void _stdcall CVSPPGFlex::OnAfterUserResizePV(long Row, long Col)
{
	// show new width value
	OnSelChangePV();

	// if multiple columns are selected, apply new width to all
    if (m_spGridPreview->Col != m_spGridPreview->ColSel) {
        long r1, c1, r2, c2;
        m_spGridPreview->GetSelection(&r1, &c1, &r2, &c2);
        long lWidth = m_spGridPreview->GetColWidth(Col);
        for (long c = c1; c <= c2; c++)
            if (c != Col)
                m_spGridPreview->PutColWidth(c, lWidth);
    }

	// remember we're dirty
	SetDirty(TRUE);
}

// AfterMoveColumn: we're dirty...
void _stdcall CVSPPGFlex::OnAfterMoveColumnPV(long Col, long *Position)
{
	SetDirty(TRUE);
}

