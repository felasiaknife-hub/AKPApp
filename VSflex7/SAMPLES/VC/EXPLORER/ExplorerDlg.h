// ExplorerDlg.h : Declaration of the CExplorerDlg

#ifndef __EXPLORERDLG_H_
#define __EXPLORERDLG_H_

#include "resource.h"       // main symbols


/////////////////////////////////////////////////////////////////////////////
// CWatchWnd: forward messages sent by nested controls to parent.
class CWatchWnd : public CWindowImpl<CWatchWnd>
{
public:
	DECLARE_WND_CLASS_EX(NULL, 0, 0)
	BEGIN_MSG_MAP(CWatchWnd)
		MESSAGE_HANDLER(WM_CTLCOLORSTATIC, OnCtlColor)
	END_MSG_MAP()
	LRESULT OnCtlColor(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		return ::SendMessage(GetParent(), uMsg, wParam, lParam);
	}
};


/////////////////////////////////////////////////////////////////////////////
// CExplorerDlg
class CExplorerDlg : 
	public CAxDialogImpl<CExplorerDlg>,
	public IDispEventImpl<IDC_FGLIST, CExplorerDlg>,
	public IDispEventImpl<IDC_FGTREE, CExplorerDlg>
{
public:
	CExplorerDlg()
	{
	}

	~CExplorerDlg()
	{
	}

	enum { IDD = IDD_EXPLORERDLG };

BEGIN_MSG_MAP(CExplorerDlg)
	MESSAGE_HANDLER(WM_INITDIALOG,				OnInitDialog)
	MESSAGE_HANDLER(WM_SIZE,					OnSize)
	MESSAGE_HANDLER(WM_CLOSE,					OnClose)
	MESSAGE_HANDLER(WM_CTLCOLORSTATIC,			OnCtlColor)
	MESSAGE_HANDLER(WM_LBUTTONDOWN,				OnLButtonDown)
	MESSAGE_HANDLER(WM_MOUSEMOVE,				OnMouseMove)
END_MSG_MAP()
// Handler prototypes:
//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

BEGIN_SINK_MAP(CExplorerDlg)
	//Make sure the Event Handlers have __stdcall calling convention
	SINK_ENTRY(IDC_FGTREE, DISPID_DBLCLICK,		OnDblClickTree)
	SINK_ENTRY(IDC_FGTREE, DISPID_KEYDOWN,		OnKeyDownTree)
	SINK_ENTRY(IDC_FGTREE, 0x2,					OnRowColChangeTree)
	SINK_ENTRY(IDC_FGTREE, 0x13,				OnBeforeCollapseTree)

	SINK_ENTRY(IDC_FGLIST, DISPID_MOUSEMOVE,	OnMouseMoveList)
	SINK_ENTRY(IDC_FGLIST, DISPID_DBLCLICK,		OnDblClickList)
	SINK_ENTRY(IDC_FGLIST, 0x1,					OnSelChangeList)
	SINK_ENTRY(IDC_FGLIST, 0x5,					OnBeforeMouseDownList)
	SINK_ENTRY(IDC_FGLIST, 0x25,				OnOLEStartDragList)
END_SINK_MAP()


/////////////////////////////////////////////////////////////////////////////////////////
//
// Windows message handlers
//
/////////////////////////////////////////////////////////////////////////////////////////

	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		// get pointers to grid controls
		CAxWindow wndTree = GetDlgItem(IDC_FGTREE);
		wndTree.QueryControl(&m_spTree);
		CAxWindow wndList = GetDlgItem(IDC_FGLIST);
		wndList.QueryControl(&m_spList);

		// activate controls and hook up events
		wndList.SetFocus();
		wndTree.SetFocus();
		AtlAdviseSinkMap(this, true);

		// save pointers to all controls
		m_wndTree   = GetDlgItem(IDC_FGTREE);
		m_wndList   = GetDlgItem(IDC_FGLIST);
		m_wndLblDir = GetDlgItem(IDC_LBLDIR);
		m_wndWatch.SubclassWindow(m_wndLblDir);
		m_wndLblSel = GetDlgItem(IDC_LBLSELECTION);

		// initialize controls
		InitControls();
		UpdateLayout();

		// initialize tooltip
		m_ToolTip.Create((HWND)m_spList->hWnd);

		// initialize display
		ShowDrives();

		return 0;
	}

	LRESULT OnSize(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		UpdateLayout();
		return 0;
	}

	LRESULT OnCtlColor(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		HDC  hdcStatic  = (HDC)wParam;
		HWND hwndStatic = (HWND)lParam; 
		SetTextColor(hdcStatic, (hwndStatic == GetDlgItem(IDC_LBLDIR))? 0: 16711680 /* vbBlue */);
		SetBkMode(hdcStatic, TRANSPARENT);
		return (LRESULT)m_brushDir.GetBrushHandle();
	}

    LRESULT OnClose(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		EndDialog(IDOK);
		return 0;
	}

	LRESULT OnMouseMove(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		// bail out if any keys are being pressed
		if (wParam) return 0;

		// check whether the mouse is over the resizing area
		POINT pt = { LOWORD(lParam), HIWORD(lParam) };
		if (PtInRect(&m_rcResizeArea, pt))
			SetCursor(LoadCursor(NULL, IDC_SIZEWE));

		// done
		return 0;
	}
	
	void DrawResizeBar(LPRECT rc)
	{
		ClientToScreen(rc);
		HDC  hdc = ::GetDC(NULL);
		InvertRect(hdc, rc);
		ReleaseDC(hdc);
		ScreenToClient(rc);
	}

	LRESULT OnLButtonDown(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		// check whether the mouse is over the resizing area
		POINT pt = { LOWORD(lParam), HIWORD(lParam) };
		if (!PtInRect(&m_rcResizeArea, pt)) return 0;
		SetCursor(LoadCursor(NULL, IDC_SIZEWE));

		// prepare to work
		RECT rcClient, rcSizeBar;
		GetClientRect(&rcClient);
		InflateRect(&rcClient, -10, -3);
		rcSizeBar = m_rcResizeArea;
		int iBarWidth = rcSizeBar.right - rcSizeBar.left;

		// capture the mouse until the user is done resizing
		SetCapture();
		MSG  msg;
		BOOL bErase = false;
	    for (BOOL bDone = false; !bDone;) {
			if (PeekMessage(&msg, NULL, WM_KEYFIRST, WM_MOUSELAST, PM_REMOVE)) {
			    switch (msg.message) {

					// quit when the button is released or a key is hit
					case WM_KEYDOWN:
					case WM_LBUTTONUP:
						bDone = true;
						break;

				    // handle mouse movement
					case WM_MOUSEMOVE:
						SetCursor(LoadCursor(NULL, IDC_SIZEWE));
						pt.x = LOWORD(msg.lParam);
						pt.y = HIWORD(msg.lParam);  
						if (!PtInRect(&rcClient, pt)) break;
						if (bErase) DrawResizeBar(&rcSizeBar);
						rcSizeBar.left  = pt.x - iBarWidth / 2;
						rcSizeBar.right = rcSizeBar.left + iBarWidth;
						DrawResizeBar(&rcSizeBar);
						bErase = true;
				}
		    }
		}

		// release everything
		if (bErase) DrawResizeBar(&rcSizeBar);
		ReleaseCapture();
		SetCursor(NULL);

		// if the user canceled, we're done
		if (msg.message != WM_LBUTTONUP) return 0;

		// do the resize
		m_wndTree.GetClientRect(&rcClient);
		rcClient.right += rcSizeBar.left - m_rcResizeArea.left;
		UpdateLayout(rcClient.right);

		// done
		return 0;
	}

/////////////////////////////////////////////////////////////////////////////////////////
//
// FlexGrid Tree event handlers
//
/////////////////////////////////////////////////////////////////////////////////////////

	VOID __stdcall OnBeforeCollapseTree(LONG Row, SHORT State, VARIANT_BOOL * Cancel)
	{
		// if we're working, ignore this
		if (m_spTree->Redraw == flexRDNone)
			return;
    
		// don't allow collapsing/expanding groups of rows (with shift-click)
		if (Row < 0) {
			*Cancel = true;
			return;
		}
    
		// make sure we're expanding
		if (State == flexOutlineCollapsed)
			return;
    
		// make sure we have to populate branch
		if (Row >= m_spTree->Rows - 1) return;
		if (m_spTree->GetIsSubtotal(Row) == false) return;
		_bstr_t str = m_spTree->GetTextMatrix(Row + 1, 0);
		if (str.length()) return;
        
		// remove dummy branch
		m_spTree->RemoveItem(Row + 1L);
    
		// populate the branch we're about to open
		ExpandRow(Row);
	}

	// expand/collapse when double-clicking
	VOID __stdcall OnDblClickTree()
	{
		if (m_spTree->GetIsCollapsed(m_spTree->Row) == flexOutlineCollapsed)
			m_spTree->PutIsCollapsed(m_spTree->Row, flexOutlineExpanded);
		else
			m_spTree->PutIsCollapsed(m_spTree->Row, flexOutlineCollapsed);
	}

    // use arrow keys to expand/collapse node
	VOID __stdcall OnKeyDownTree(SHORT * KeyCode, SHORT Shift)
	{
		switch (*KeyCode) {
			case VK_LEFT:
				m_spTree->PutIsCollapsed(m_spTree->Row, flexOutlineCollapsed);
				break;
			case VK_RIGHT:
				m_spTree->PutIsCollapsed(m_spTree->Row, flexOutlineExpanded);
				break;
		}
	}

	VOID __stdcall OnRowColChangeTree()
	{
		UpdateFileList();
	}


/////////////////////////////////////////////////////////////////////////////////////////
//
// FlexGrid file list event handlers
//
/////////////////////////////////////////////////////////////////////////////////////////

	VOID __stdcall OnBeforeMouseDownList(SHORT Button, SHORT Shift, FLOAT X, FLOAT Y, VARIANT_BOOL * Cancel)
	{
	    // use shift-clicks to start OLE drag-drop
		if ((Shift & 1) == 0) return;
    
		// tell control to ignore this click, we'll use it ourselves
		*Cancel = true;
    
		// start drag-drop (set data object in OLEStartDrag event)
		m_spList->OLEDrag();
	}

	VOID __stdcall OnOLEStartDragList(VSDataObject * * Data, LONG * AllowedEffects)
	{
		// get path for currently selected files
		_variant_t vt = m_spTree->GetRowData(m_spTree->Row);
		if (vt.vt != VT_BSTR) return;
		_bstr_t strPath = V_BSTR(&vt);
		strPath += _T("\\");

		// build list of selected files
		_bstr_t strList;
		int iCount = m_spList->SelectedRows;
		for (int i = 0; i < iCount; i++) {
			if (strList.length()) strList += _T("\t");
			strList += strPath;
			strList += m_spList->GetTextMatrix(m_spList->GetSelectedRow(i), 0);
		}
    
		// put selected files into VSDataObject
		IVSDataObjectPtr pDO((LPUNKNOWN)(*Data));
		pDO->SetData(strList, (long)CF_HDROP);
	}

	VOID __stdcall OnMouseMoveList(SHORT Button, SHORT Shift, FLOAT X, FLOAT Y)
	{
		LPTSTR sz = NULL;
        static int i = 0;

        // show what to do with files if mouse is over header
        if (Button == 0 && m_spList->Rows > 1 && m_spList->MouseRow == 0) {
            switch (i) {
                case 0: sz = _T("Click headers to sort columns"); break;
                case 1: sz = _T("Shift-Drag selected files into NotePad or Word"); break;
                case 2: sz = _T("Double-Click files to launch them"); break;
            }

        // if not over header, cycle message
        } else {
            i = (i + 1) % 3;
        }

        // show the tooltip
		m_ToolTip.Update(sz);
	}

	VOID __stdcall OnSelChangeList()
	{
		UpdateFileStats();
	}

	VOID __stdcall OnDblClickList()
	{
		// get path for currently selected files
		_variant_t vt = m_spTree->GetRowData(m_spTree->Row);
		if (vt.vt != VT_BSTR) return;
		_bstr_t strFileName = V_BSTR(&vt);
		strFileName += _T("\\");
		strFileName += m_spList->GetTextMatrix(m_spList->Row, 0);

		// launch file
		USES_CONVERSION;
		ShellExecute(m_hWnd, NULL, W2T(strFileName), NULL, NULL, SW_SHOWDEFAULT);
	}

	
/////////////////////////////////////////////////////////////////////////////////////////
//
// Internal methods
//
/////////////////////////////////////////////////////////////////////////////////////////

	void InitControls()
	{
		// initialize tree control
		//
		m_spTree->Cols = 1;
        m_spTree->Rows = 0;
        m_spTree->FixedCols = 0;
        m_spTree->FixedRows = 0;
        
        // appearance
        m_spTree->GridLines = flexGridNone;
        m_spTree->BackColorBkg = m_spTree->BackColor;
        m_spTree->SheetBorder = m_spTree->BackColor;
        m_spTree->ExtendLastCol = true;
        m_spTree->Redraw = flexRDBuffered;
        m_spTree->OutlineCol = 0;
        m_spTree->OutlineBar = flexOutlineBarCompleteLeaf;
        m_spTree->Ellipsis = flexEllipsisEnd;

        // behavior
        m_spTree->AllowSelection = false;
        m_spTree->HighLight = flexHighlightWithFocus;
        m_spTree->ScrollTrack = true;
        m_spTree->AutoSearch = flexSearchFromCursor;
    
        // create WallPaper picture from file
        CPicHolder pic;
        pic.CreateFromFile(_T("bkg.jpg"));
        if (pic.HasPicture()) {
            m_spTree->putref_WallPaper(pic.GetPictureDispatch(false));
			m_spTree->WallPaperAlignment = flexPicAlignLeftTop;
		}

		// initialize list control
		//
        m_spList->Cols = 3;
        m_spList->Rows = 1;
        m_spList->FixedCols = 0;
        m_spList->FixedRows = 1;
        
        // appearance
        m_spList->BorderStyle = flexBorderNone;
        m_spList->GridLines = flexGridNone;
        m_spList->BackColorBkg = m_spList->BackColor;
        m_spList->SheetBorder = m_spList->BackColor;
        m_spList->ExtendLastCol = true;
        m_spList->SelectionMode = flexSelectionListBox;
        m_spList->FocusRect = flexFocusNone;
        m_spList->AllowUserResizing = flexResizeColumns;
        m_spList->Ellipsis = flexEllipsisEnd;
        
        // behavior
        m_spList->ExplorerBar = flexExSortShow;
        m_spList->HighLight = flexHighlightAlways;
        m_spList->ScrollTrack = true;
        m_spList->AllowUserResizing = flexResizeColumns;
        m_spList->AutoSearch = flexSearchFromCursor;
        
        // content
        m_spList->PutTextMatrix(0, 0, _bstr_t("Name"));
        m_spList->PutTextMatrix(0, 1, _bstr_t("Size (k)"));
        m_spList->PutTextMatrix(0, 2, _bstr_t("Modified"));
        m_spList->PutColAlignment(1, flexAlignRightCenter);
        m_spList->PutColAlignment(2, flexAlignRightCenter);
        m_spList->PutColFormat(1, _bstr_t("#,###.#"));
        m_spList->PutColFormat(2, _bstr_t("Medium Date"));
        m_spList->PutColDataType(2, flexDTDate);
        
	    // initialize dir label control
		//
		m_fntDir.Create("Tahoma", 14, true);
		m_wndLblDir.SetFont(m_fntDir.GetFontHandle());
		m_wndLblDir.SetWindowText(" My Computer");
        
		// initialize selection label control
		//
		m_fntSelection.Create("Tahoma", 8);
		m_wndLblSel.SetFont(m_fntSelection.GetFontHandle());
		m_wndLblSel.SetWindowText("0 Items selected ");

		// initialize background brush for both label controls
		//
		m_brushDir.Create(m_spTree->BackColor);

		// make list and selection label children of the main label control
		//
		m_wndList.SetParent(m_wndLblDir);
		m_wndLblSel.SetParent(m_wndLblDir);
	}

	void UpdateLayout(int iTreeWidth = -1)
	{
		// initialize parameters
		int iBorder = 3;
		RECT rc;
		GetClientRect(&rc);
		InflateRect(&rc, -iBorder, -iBorder);

		// move and position tree control
		RECT rcTree;
		m_wndTree.GetClientRect(&rcTree);
		if (iTreeWidth < 0) iTreeWidth = rcTree.right;
		rcTree = rc;
		rcTree.right = rcTree.left + iTreeWidth;
		m_wndTree.MoveWindow(&rcTree);

		// move and position main label
		RECT rcLblDir = rc;
		rcLblDir.left = rcTree.right + iBorder + 2;
		m_wndLblDir.MoveWindow(&rcLblDir);
    
		// save resizing area
		m_rcResizeArea = rc;
		m_rcResizeArea.left  = rcTree.right;
		m_rcResizeArea.right = rcLblDir.left;

		// move and position list control
		m_wndLblDir.GetClientRect(&rc);
		rc.top += 40;
		m_wndList.MoveWindow(&rc);
		m_spList->PutColWidth(0, m_spList->ClientWidth / 2);
		m_spList->PutColWidth(1, m_spList->ClientWidth / 4);
    
		// move and position selection label
		RECT rcLblSel;
		m_wndLblSel.GetClientRect(&rcLblSel);
		int iHei = rcLblSel.bottom;
		rc.bottom = rc.top;
		rc.top -= iHei;
		m_wndLblSel.MoveWindow(&rc);
	}

	void GetDriveList()
	{
		// get bit mask with all drives
		DWORD dwDrives = GetLogicalDrives();

		// digest bit mask to get drive count and letters
		LPTSTR szDrive = m_szDrives;
		for (int i = 0; i < 26; i++) {
			if (dwDrives & (1 << i)){
				*szDrive = (TCHAR)(i + 'A');
				szDrive++;
			}
		}
		*szDrive = 0;

		// get drive types
		TCHAR sz[4];
		_tcscpy(sz, _T("?:\\"));
		for (i = 0; m_szDrives[i]; i++) {
			*sz = m_szDrives[i];
			m_iDriveTypes[i] = GetDriveType(sz);
		}
	}
   
	void ShowDrives()
	{
        // prepare to work
        m_spTree->Redraw = flexRDNone;
		GetDriveList();
        
        // add computer icon to tree
		CPicHolder pic;
		pic.CreateFromIcon(IDI_COMPUTER);
        m_spTree->Rows = 0;
        m_spTree->AddItem(_bstr_t("My Computer"));
		_variant_t vtPic(pic.GetPictureDispatch(false));
        m_spTree->PutCell(flexcpPicture, 0L, vtMissing, vtMissing, vtMissing, vtPic);
        m_spTree->PutIsSubtotal(0, true);
        
        // add drives to tree
        for (int i = 0; m_szDrives[i]; i++)
		{
			// get drive letter
			TCHAR szBuff[5];
			*szBuff = m_szDrives[i];
			szBuff[1] = 0;
			_tcscat(szBuff, _T(":\\"));

			// get drive name
			TCHAR szVolume[256];
			*szVolume = 0;
			GetVolumeInformation(szBuff, szVolume, 256, NULL, NULL, NULL, NULL, 0);
			_bstr_t strDriveName = szBuff;
			if (*szVolume) {
				_tcscat(szVolume, _T(" (?:)"));
				szVolume[lstrlen(szVolume) - 3] = m_szDrives[i];
				strDriveName = szVolume;
			}

			// get drive picture
			switch (m_iDriveTypes[i]) {
				case DRIVE_REMOTE:
					pic.CreateFromIcon(IDI_DISK_REMOTE);
					break;
				case DRIVE_REMOVABLE:
					pic.CreateFromIcon(IDI_DISK_FLOPPY);
					break;
				case DRIVE_CDROM: 
					pic.CreateFromIcon(IDI_DISK_CDROM);
					break;
				case DRIVE_FIXED:
				default:
					pic.CreateFromIcon(IDI_DISK_FIXED);
			}
			vtPic = pic.GetPictureDispatch(false);

            // add drive node
            m_spTree->AddItem(strDriveName);
			long lRow = m_spTree->Rows - 1;

			// add item info
            m_spTree->PutRowData(lRow, szBuff);
            m_spTree->PutCell(flexcpPicture, lRow, vtMissing, vtMissing, vtMissing, vtPic);
            m_spTree->PutIsSubtotal(lRow, true);
            m_spTree->PutRowOutlineLevel(lRow, 1);
            
            // add hidden node (we'll handle this when the user expands it)
            m_spTree->AddItem(_bstr_t(""));
            m_spTree->PutIsSubtotal(lRow + 1, true);
            m_spTree->PutRowOutlineLevel(lRow + 1, 2);
        }
        
        // start showing all drives and nothing else
        m_spTree->Outline(1);
        m_spTree->Row = 0;
        OnRowColChangeTree();
        
        // ready
        m_spTree->Redraw = flexRDBuffered;
	}

	void ExpandRow(long r)
	{
        // get path for the given row
		_variant_t vPath = m_spTree->GetRowData(r);
		_bstr_t strPath = V_BSTR(&vPath);
    
        // get list of subdirectories
		m_FileList.GetSubDirs(strPath);
        
        // make it fast
        m_spTree->Redraw = flexRDNone;
        SetCursor(LoadCursor(NULL, IDC_WAIT));
    
        // get information for new items
        long iPos = r + 1;
        long iLvl = m_spTree->GetRowOutlineLevel(r) + 1;

        // get folder picture
		CPicHolder pic;
		pic.CreateFromIcon(IDI_FOLDER);
		_variant_t vtPic(pic.GetPictureDispatch(false));

        // list subdirectories
        for (long i = 0; i < m_FileList.GetSize(); i++) {
    
            // add dir
            m_spTree->AddItem(m_FileList.GetFileName(i), iPos);
            m_spTree->PutIsSubtotal(iPos, true);
            m_spTree->PutRowOutlineLevel(iPos, (short)iLvl);
            m_spTree->PutCell(flexcpPicture, iPos, 0L, vtMissing, vtMissing, vtPic);
        
            // save path in rowdata
            m_spTree->PutRowData(iPos, m_FileList.GetFullPath(i));
            iPos++;
        
            // if this directory has subdirs, add dummy item
            // (we'll populate this when it expands)
            if (m_FileList.HasSubDirs(i)) {
				m_spTree->AddItem(_bstr_t(""), iPos);
				m_spTree->PutIsSubtotal(iPos, true);
				m_spTree->PutRowOutlineLevel(iPos, iLvl + 1);
				iPos++;
			}
        }
		
        // collapse the items we just added
        for (i = r + 1; i < iPos; i++) {
            if (m_spTree->GetIsSubtotal(i))
				m_spTree->PutIsCollapsed(i, flexOutlineCollapsed);
		}
    
        // ready to redraw
        SetCursor(NULL);
        m_spTree->Redraw = flexRDBuffered;
        m_spTree->Row = r;
	}

	void UpdateFileList()
	{

		// update main label
		_bstr_t str(" ");
		str += m_spTree->GetText();
		m_wndLblDir.SetWindowText(str);
		m_wndLblSel.SetWindowText(NULL);
    
		// initialize list
		m_spList->Rows = 1;
    
		// get selected path
		str = "";
		_variant_t vt = m_spTree->GetRowData(m_spTree->Row);
		if (vt.vt == VT_BSTR) str = V_BSTR(&vt);
		if (str.length() == 0) return;
    
		// save current row
		long r = m_spTree->Row;
    
		// get files
		m_FileList.GetFiles(str);
		if (m_FileList.GetSize() == 0) return;
    
		// populate the list
        SetCursor(LoadCursor(NULL, IDC_WAIT));
		m_spList->Refresh();
		m_spList->Redraw = flexRDNone;
		for (long i = 0; i < m_FileList.GetSize(); i++)
		{
			str =  m_FileList.GetFileName(i);
			str += "\t";
			str += _bstr_t((m_FileList.GetFileSize(i) / 1024.0));
			str += "\t";
			str += m_FileList.GetFileDate(i);
			m_spList->AddItem(str);
		}
		m_spList->Redraw = flexRDDirect;
        SetCursor(NULL);
    
		// update statistics for this file list
		UpdateFileStats();
	}

	void UpdateFileStats()
	{
		TCHAR szBuff[100];

		// get number of selected files
		long nSel = m_spList->SelectedRows;
    
		// one or none? easy
		if (nSel < 2) {
			wsprintf(szBuff, "%d Files ", m_spList->Rows - 1);
			m_wndLblSel.SetWindowText(szBuff);
			return;
		}
    
		// calculate size for selection > 1 file using new Aggregate property
		// row -1 means "all selected rows", col 1 has the file size
		double dTot = m_spList->GetAggregate(flexSTSum, -1, 1, -1, 1);
		TCHAR  cUnit = 'k';
		if (dTot > 2 * 1024) {
			dTot /= 1024;
			cUnit = 'm';
		}
    
		// show label
		wsprintf(szBuff, "%d Files Selected (%d%c) ", nSel, (int)dTot, cUnit);
		m_wndLblSel.SetWindowText(szBuff);
	}

private:
// Attributes
	// pointers to grid controls
	IVSFlexGridPtr	m_spTree, m_spList;

	// windows for each child control
	CWindow		m_wndTree, m_wndList, m_wndLblDir, m_wndLblSel;
	CWatchWnd	m_wndWatch;

	// CFileList object used to get dirs/files
	CFileList	m_FileList;

	// VB-like tooltip
	CToolTip	m_ToolTip;

	// fonts and brush for label controls
	CGDIFont	m_fntDir, m_fntSelection;
	CGDIBrush	m_brushDir;

	// resizing area (between tree and list)
	RECT		m_rcResizeArea;

	// drive info
	TCHAR		m_szDrives[30];		// all drive letters (e.g.: "ACDFG")
	int			m_iDriveTypes[30];	// drive types (as returned by GetDriveType API)
};

#endif //__EXPLORERDLG_H_
