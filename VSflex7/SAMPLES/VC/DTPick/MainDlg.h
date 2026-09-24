// MainDlg.h : Declaration of the CMainDlg

#ifndef __MAINDLG_H_
#define __MAINDLG_H_

#include "resource.h"       // main symbols


/////////////////////////////////////////////////////////////////////////////
// CMainDlg
class CMainDlg : 
	public CAxDialogImpl<CMainDlg>,
	public IDispEventImpl<IDC_VSFLEXGRID1, CMainDlg>
{
public:
	CMainDlg()
	{
	}

	~CMainDlg()
	{
	}

	enum { IDD = IDD_MAINDLG };


/////////////////////////////////////////////////////////////////////////////
//
// Message Handlers
//
/////////////////////////////////////////////////////////////////////////////
BEGIN_MSG_MAP(CMainDlg)
	MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
	MESSAGE_HANDLER(WM_SIZE,	OnSize)
	MESSAGE_HANDLER(WM_CLOSE,	OnClose)
	MESSAGE_HANDLER(WM_NOTIFY,	OnNotify)
	MESSAGE_HANDLER(WM_COMMAND,	OnCommand)
END_MSG_MAP()

	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		// get pointer to control, activate it, and hook up events
		m_wndGrid = GetDlgItem(IDC_VSFLEXGRID1);
		m_wndGrid.QueryControl(&m_spGrid);
		m_wndGrid.SetFocus();
		AtlAdviseSinkMap(this, true);

		// initialize grid
		m_spGrid->Editable = flexEDKbdMouse;
		m_spGrid->AllowUserResizing = flexResizeBoth;
		m_spGrid->Cols = 4;
		m_spGrid->RowHeightMin = 280;
		m_spGrid->PutColWidth    (0, m_spGrid->GetRowHeight(0));
		m_spGrid->PutColWidth    (1, 2500);
		m_spGrid->PutColWidth    (2, 1500);
		m_spGrid->PutColWidth    (3, 1500);
		m_spGrid->PutTextMatrix  (0, 1, "Name");
		m_spGrid->PutTextMatrix  (0, 2, "Born");
		m_spGrid->PutTextMatrix  (0, 3, "Hired");
		m_spGrid->PutColDataType (2, flexDTDate);
		m_spGrid->PutColDataType (3, flexDTDate);
		m_spGrid->PutColFormat   (2, "Medium Date");
		m_spGrid->PutColFormat   (3, "Medium Date");
		m_spGrid->PutColComboList(2, "Dummy"); // just to show the down-arrow
		m_spGrid->PutColComboList(3, "Dummy"); // just to show the down-arrow
    
		// fill up with dummy data
		for (long r = 1; r < m_spGrid->Rows; r++) {
			TCHAR szBuff[30];
			wsprintf(szBuff, "Employee %d", r);
			m_spGrid->PutTextMatrix(r, 1, szBuff);
			wsprintf(szBuff, "%d/%d/%d", rand() % 12 + 1, rand() % 27 + 1, 1950 + rand() % 20);
			m_spGrid->PutTextMatrix(r, 2, szBuff);
			wsprintf(szBuff, "%d/%d/%d", rand() % 12 + 1, rand() % 27 + 1, 1999 - rand() % 10);
			m_spGrid->PutTextMatrix(r, 3, szBuff);
		}

		// create date/time picker control
		// it is initially a child of the dialog, so the dialog will get the notifications
		// we'll reparent it later to facilitate positioning and improve painting
		INITCOMMONCONTROLSEX icc;
		icc.dwSize = sizeof(icc);
		icc.dwICC  = ICC_DATE_CLASSES;
		InitCommonControlsEx(&icc); 
		RECT rc;
		SetRect(&rc, 0, 0, 0, 0);
		m_dtPick.Create(m_hWnd, rc, NULL, WS_CHILD);

		// done
		return FALSE;
	}

	LRESULT OnSize(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		RECT rc;
		GetClientRect(&rc);
		InflateRect(&rc, -3, -3);
		m_wndGrid.MoveWindow(&rc);
		return 0;
	}

	LRESULT OnClose(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		EndDialog(IDOK);
		return 0;
	}

	// process notifications from DateTimePicker control
	LRESULT OnNotify(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		SYSTEMTIME sysTime;
		LPNMHDR    pNotify = (LPNMHDR)lParam; 
		ATLTRACE("OnNotify: code %d hWnd %d\n", pNotify->code, pNotify->hwndFrom);
		if (pNotify->hwndFrom == m_dtPick) {
			switch (pNotify->code) {

				// update grid value whenever the data changes
				case DTN_DATETIMECHANGE:
					m_dtPick.GetSystemTime(&sysTime);
					m_spGrid->Text = GetDateString(&sysTime);
					break;

			    // hide date picker when user is done with it
				case -8:
					m_dtPick.ShowWindow(SW_HIDE);
					break;
			}
		}
		return 0;
	}

	// if the user hit Enter or Escape, hide the DateTimePicker control
	LRESULT OnCommand(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		if (m_dtPick != NULL && m_dtPick.IsWindowVisible()) {
			WORD wID = LOWORD(wParam);
			switch (wID) {
				case IDCANCEL:
					m_spGrid->Text = m_strSaveDate;
					// FALL THROUGH
				case IDOK:
					m_dtPick.ShowWindow(SW_HIDE);
			}
		}
		return 0;
	}


/////////////////////////////////////////////////////////////////////////////
//
// Event Handlers
//
/////////////////////////////////////////////////////////////////////////////
BEGIN_SINK_MAP(CMainDlg)
	//Make sure the Event Handlers have __stdcall calling convention
	SINK_ENTRY(IDC_VSFLEXGRID1, 0x11, OnBeforeUserResize)
	SINK_ENTRY(IDC_VSFLEXGRID1, 0xa, OnBeforeScroll)
	SINK_ENTRY(IDC_VSFLEXGRID1, 0x16, OnStartEdit)
END_SINK_MAP()

	VOID __stdcall OnBeforeUserResize(LONG Row, LONG Col, VARIANT_BOOL * Cancel)
	{
		if (m_dtPick != NULL && m_dtPick.IsWindowVisible())
			*Cancel = TRUE;
	}

	VOID __stdcall OnBeforeScroll(LONG OldTopRow, LONG OldLeftCol, LONG NewTopRow, LONG NewLeftCol, VARIANT_BOOL * Cancel)
	{
		if (m_dtPick != NULL && m_dtPick.IsWindowVisible())
			*Cancel = TRUE;
	}

	VOID __stdcall OnStartEdit(LONG Row, LONG Col, VARIANT_BOOL * Cancel)
	{
		// if this is not a date column, allow default behavior
		if (m_spGrid->GetColDataType(Col) != flexDTDate) return;
        
        // we'll handle the editing ourselves
        *Cancel = true;
        
        // initialize value, save original in tag in case user hits escape
		m_strSaveDate = m_spGrid->Text;
		m_dtPick.SetSystemTime(GDT_VALID, GetSystemTime(m_strSaveDate));
        
		// calculate position for time picker control
		RECT rc;
		SetRect(&rc, XTwipsToPixels(m_spGrid->CellLeft),  YTwipsToPixels(m_spGrid->CellTop),
					 XTwipsToPixels(m_spGrid->CellWidth), YTwipsToPixels(m_spGrid->CellHeight));
		rc.right  += rc.left;
		rc.bottom += rc.top;

        // position date picker control over cell
		m_dtPick.SetParent((HWND)m_spGrid->hWnd);
        m_dtPick.MoveWindow(&rc);
        
        // show and activate date picker control
        m_dtPick.ShowWindow(SW_NORMAL);
        m_dtPick.SetFocus();
        
        // make it drop down the calendar
		m_dtPick.SendMessage(WM_KEYDOWN, VK_F4);
	}


/////////////////////////////////////////////////////////////////////////////
//
// Helpers
//
/////////////////////////////////////////////////////////////////////////////

	LPSYSTEMTIME GetSystemTime(_bstr_t strDate)
	{
		static SYSTEMTIME t;
		memset(&t, 0, sizeof(t));
		CComVariant v = (BSTR)strDate;
		if (S_OK == v.ChangeType(VT_DATE))
			VariantTimeToSystemTime(V_DATE(&v), &t);
		return &t;
	}

	_bstr_t GetDateString(LPSYSTEMTIME pSysTime)
	{
		CComVariant v;
		V_VT(&v) = VT_DATE;
		SystemTimeToVariantTime(pSysTime, &V_DATE(&v));
		v.ChangeType(VT_BSTR);
		return _bstr_t(V_BSTR(&v));
	}


/////////////////////////////////////////////////////////////////////////////
//
// Member Variables
//
/////////////////////////////////////////////////////////////////////////////
	CAxWindow			m_wndGrid;
	IVSFlexGridPtr		m_spGrid;
	CDateTimePickerCtrl	m_dtPick;
	_bstr_t				m_strSaveDate;

};

#endif //__MAINDLG_H_
