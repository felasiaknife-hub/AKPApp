// WallPaperDlg.h : Declaration of the CWallPaperDlg

#ifndef __WALLPAPERDLG_H_
#define __WALLPAPERDLG_H_


/////////////////////////////////////////////////////////////////////////////
// CWallPaperDlg
class CWallPaperDlg : 
	public CAxDialogImpl<CWallPaperDlg>
{
public:
	CWallPaperDlg()
	{
	}

	~CWallPaperDlg()
	{
	}

	enum { IDD = IDD_WALLPAPERDLG };

BEGIN_MSG_MAP(CWallPaperDlg)
	MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
	MESSAGE_HANDLER(WM_SIZE, OnSize)
	COMMAND_HANDLER(IDC_COMBO1, CBN_SELENDOK, OnNewAlignment)
	COMMAND_HANDLER(IDC_LIST1, LBN_SELCHANGE, OnNewWallPaper)
	MESSAGE_HANDLER(WM_CLOSE, OnClose)
END_MSG_MAP()
// Handler prototypes:
//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
        // fill alignment list
        static LPTSTR szAlign[] =
        {
            _T("0 - LeftTop"),     
            _T("1 - LeftCenter"),  
            _T("2 - LeftBottom"),  
            _T("3 - CenterTop"),   
            _T("4 - CenterCenter"),
            _T("5 - CenterBottom"),
            _T("6 - RightTop"),    
            _T("7 - RightCenter"), 
            _T("8 - RightBottom"), 
            _T("9 - Stretch"),     
            _T("10 - Tile"),        
            NULL
        };
        for (int i = 0; szAlign[i]; i++)
            SendDlgItemMessage(IDC_COMBO1, CB_ADDSTRING, 0, (LPARAM)szAlign[i]);
        SendDlgItemMessage(IDC_COMBO1, CB_SETCURSEL, 9);

        // fill file list
        static LPTSTR szPicExt[] =
        {
            _T("*.jpg"),     
            _T("*.bmp"),  
            _T("*.gif"),  
            _T("*.wmf"),   
            NULL
        };
        GetModuleFileName(_Module.GetModuleInstance(), m_szPicturePath, 512);
        LPTSTR sz = _tcsrchr(m_szPicturePath, '\\');
        if (sz) *sz = 0;
        _tchdir(m_szPicturePath);
        for (i = 0; szPicExt[i]; i++)
            SendDlgItemMessage(IDC_LIST1, LB_DIR, 0, (LPARAM)szPicExt[i]);

        // no files in our directory? try one level up
        if (SendDlgItemMessage(IDC_LIST1, LB_GETCOUNT) == 0) {
            sz = _tcsrchr(m_szPicturePath, '\\');
            if (sz) *sz = 0;
            _tchdir(m_szPicturePath);
            for (i = 0; szPicExt[i]; i++)
                SendDlgItemMessage(IDC_LIST1, LB_DIR, 0, (LPARAM)szPicExt[i]);
        }

        // get pointer to grid
        CAxWindow wndGrid = GetDlgItem(IDC_VSFLEXGRID1);
        wndGrid.QueryControl(&m_spGrid);

        // make the grid UI-active
        wndGrid.SetFocus();

        // initialize grid
        m_spGrid->put_BackColor(RGB(190, 190, 190));                // gray background
        m_spGrid->put_GridLines(flexGridNone);                      // no grid lines
        m_spGrid->put_Editable(flexEDKbdMouse);                     // editable
        m_spGrid->put_AllowUserResizing(flexResizeBothUniform);     // resizable

		return 1;  // Let the system set the focus
	}

	LRESULT OnClose(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		EndDialog(IDOK);
		return 0;
	}

    LRESULT OnSize(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		// size grid to fit window
        RECT rc, rcGrid;
        CWindow wndGrid = GetDlgItem(IDC_VSFLEXGRID1);
        wndGrid.GetClientRect(&rcGrid);
        wndGrid.ClientToScreen(&rcGrid);
        ScreenToClient(&rcGrid);
        GetClientRect(&rc);
        InflateRect(&rc, -4, -4);
        rcGrid.bottom = rc.bottom;
        rcGrid.right = rc.right;
        wndGrid.MoveWindow(&rcGrid);

        // allow default processing
        bHandled = false;
		return 0;
	}

	LRESULT OnNewAlignment(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
	{
        // apply new alignment
        int iAlign = SendDlgItemMessage(IDC_COMBO1, CB_GETCURSEL);
        m_spGrid->put_WallPaperAlignment((PictureAlignmentSettings)iAlign);

        // done
		return 0;
	}

	LRESULT OnNewWallPaper(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
	{
		// read new wallpaper filename from control
        TCHAR szFileName[512];
        int nIndex = SendDlgItemMessage(IDC_LIST1, LB_GETCURSEL);
        SendDlgItemMessage(IDC_LIST1, LB_GETTEXT, nIndex, (LPARAM)szFileName);
        TCHAR szFileNameFull[512];
        _tcscpy(szFileNameFull, m_szPicturePath);
        _tcscat(szFileNameFull, _T("\\"));
        _tcscat(szFileNameFull, szFileName);

        // create picture from file
        CPicHolder pic;
        pic.CreateFromFile(szFileNameFull);

        // apply new wallpaper
        if (pic.HasPicture())
            m_spGrid->putref_WallPaper(pic.GetPictureDispatch(false));

        // done
		return 0;
	}

    // store pointer to grid control
    IVSFlexGridPtr  m_spGrid;

    // store path where the picture files are located
    TCHAR m_szPicturePath[512];
};

#endif //__WALLPAPERDLG_H_
