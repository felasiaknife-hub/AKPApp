/////////////////////////////////////////////////////////////////////////////
// CToolTip - helper class for implementing VB-like ToolTips
//


#ifndef _CTOOLTIP_H_
#define _CTOOLTIP_H_

// common controls needed for tooltips
#include <commctrl.h>
#pragma comment(lib, "comctl32.lib")

class CToolTip
{
private:
	HWND    m_hWnd;
	HWND    m_hWndParent;
	TCHAR	m_szLastTip[256];

public:
	CToolTip()
	{
		m_hWnd = m_hWndParent = NULL;
		*m_szLastTip = 0;
	}

	~CToolTip()
	{
		if (IsWindow(m_hWnd)) DestroyWindow(m_hWnd);
	}

	// create the tool window
	BOOL Create(HWND hWndParent, HINSTANCE hInst = NULL)
	{
		// get instance handle
		if (!hInst) {
#ifdef _MFC_VER
			hInst = AfxGetResourceHandle();
#endif
#ifdef _ATL_VER
			hInst = _Module.GetResourceInstance();
#endif
		}

		// create window
		INITCOMMONCONTROLSEX icc;
		icc.dwSize = sizeof(icc);
		icc.dwICC  = ICC_BAR_CLASSES;
		InitCommonControlsEx(&icc); 
		if (!m_hWnd) 
			m_hWnd = CreateWindowEx(WS_EX_TOPMOST, TOOLTIPS_CLASS, NULL,
                                    TTS_ALWAYSTIP | TTS_NOPREFIX,
									0, 0, 0, 0, NULL, NULL, hInst, NULL);
		if (!m_hWnd) return false;
		
		// save parent handle to show later
		m_hWndParent = hWndParent;

		// initialize info (we have only one tool)
		TOOLINFO ti;
		ti.cbSize   = sizeof(TOOLINFO);
		ti.uFlags   = TTF_IDISHWND | TTF_SUBCLASS;
		ti.hwnd     = m_hWndParent;
		ti.uId      = (UINT)m_hWndParent;
		ti.hinst    = NULL;
		ti.lpszText = _T("");
		SendMessage(m_hWnd, TTM_ADDTOOL, 0, (LPARAM)&ti);
		SendMessage(m_hWnd, TTM_SETDELAYTIME, TTDT_AUTOMATIC, 200);

		// done
		return true;
	}

	// show the tooltip with the specified text if it has changed since last call
	void Update(LPCTSTR szTip)
	{
		if (_tcscmp(szTip? szTip: _T(""), m_szLastTip))
			Show(szTip);
	}

	void Update(HWND hWndParent, LPCTSTR szTip)
    {
        UpdateParent(hWndParent);
        Update(szTip);
    }

    // show the tooltip with the specified text
	void Show(LPCTSTR szTip)
	{
		// remember last tip
		*m_szLastTip = 0;
		if (szTip) _tcscpy(m_szLastTip, szTip);

		// make sure we have a window
		if (!m_hWnd) return;

		// fill out info
		TOOLINFO ti;
		ti.cbSize   = sizeof(TOOLINFO);
		ti.uFlags   = TTF_IDISHWND | TTF_SUBCLASS;
		ti.hwnd     = m_hWndParent;
		ti.uId      = (UINT)m_hWndParent;
		ti.hinst    = NULL;
		ti.lpszText = (LPTSTR)m_szLastTip;
		
		// update the text
		SendMessage(m_hWnd, TTM_UPDATETIPTEXT, 0, (LPARAM)&ti);

        // hide or show the tooltip
        SendMessage(m_hWnd, (*m_szLastTip)? TTM_UPDATE: TTM_POP, 0, 0);
	}

	void Show(HWND hWndParent, LPCTSTR szTip)
    {
        UpdateParent(hWndParent);
        Show(szTip);
    }

    // make the tooltip belong to a new control
    void UpdateParent(HWND hWndParent)
    {
        // no change? no work.
        if (m_hWndParent == hWndParent) return;

        // remove old tool
		TOOLINFO ti;
		ti.cbSize   = sizeof(TOOLINFO);
		ti.uFlags   = TTF_IDISHWND | TTF_SUBCLASS;
		ti.hwnd     = m_hWndParent;
		ti.uId      = (UINT)m_hWndParent;
		ti.hinst    = NULL;
		ti.lpszText = _T("");
		SendMessage(m_hWnd, TTM_DELTOOL, 0, (LPARAM)&ti);

        // replace with new tool
        m_hWndParent= hWndParent;
		ti.hwnd     = m_hWndParent;
		ti.uId      = (UINT)m_hWndParent;
		SendMessage(m_hWnd, TTM_ADDTOOL, 0, (LPARAM)&ti);

        // reset last message
        *m_szLastTip = 0;
    }

	// hide the tooltip
	void Hide()
	{
		Show(NULL);
	}
};


#endif // _CTOOLTIP_H_
