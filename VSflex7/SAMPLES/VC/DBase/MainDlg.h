// MainDlg.h : Declaration of the CMainDlg

#ifndef __MAINDLG_H_
#define __MAINDLG_H_

#include "resource.h"       // main symbols
#include "DetailDlg.h"

/////////////////////////////////////////////////////////////////////////////
// CMainDlg
class CMainDlg : 
	public CAxDialogImpl<CMainDlg>,
	public IDispEventImpl<IDC_FGMAIN, CMainDlg>
{
public:
	CMainDlg()
	{
	}

	~CMainDlg()
	{
	}

	enum { IDD = IDD_MAINDLG };


////////////////////////////////////////////////////////////////////////////////////////
//
// Message Handlers
//
////////////////////////////////////////////////////////////////////////////////////////
BEGIN_MSG_MAP(CMainDlg)
	MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
	MESSAGE_HANDLER(WM_CLOSE, OnClose)
	MESSAGE_HANDLER(WM_SIZE, OnSize)
END_MSG_MAP()

	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		// get handle on label control
		m_wndLabel = GetDlgItem(IDC_HELPLABEL);

		// create imagelists we'll bind to grid columns
		m_imgDBCursor.Create(IDB_CURSOR, 16, 0, RGB(255,255,255));
		m_imgCategories.Create(IDB_CATEGORIES, 16, 0, RGB(255,255,255));

		// get pointer to VSFlexGrid control
		m_wndGrid = GetDlgItem(IDC_FGMAIN);
		m_wndGrid.QueryControl(&m_spGrid);

		// activate control and events
		m_wndGrid.SetFocus();
		AtlAdviseSinkMap(this, TRUE);

	    // define "NorthWind Best Sellers" SQL query (SQL created with Access)
		_variant_t strSQL = 
			"SELECT Products.ProductID AS Product, "
			"       Categories.CategoryID As Cat, "
			"       Suppliers.SupplierID As Supplier, "
			"       Suppliers.SupplierID As Country, "
			"       Suppliers.SupplierID As Contact, "
			"       Sum(CCur([Order Details].[UnitPrice]*[Quantity]*(1-[Discount])/100)*100) AS Revenue "
			"FROM Suppliers INNER JOIN ((Categories INNER JOIN Products ON Categories.CategoryID = Products.CategoryID) INNER JOIN "
			"((Customers INNER JOIN Orders ON Customers.CustomerID = Orders.CustomerID) INNER JOIN [Order Details] ON "
			"Orders.OrderID = [Order Details].OrderID) ON Products.ProductID = [Order Details].ProductID) ON Suppliers.SupplierID = Products.SupplierID "
			"GROUP BY Products.ProductID, Categories.CategoryID, Suppliers.SupplierID, Suppliers.SupplierID, Suppliers.SupplierID "
			"ORDER BY Sum(CCur([Order Details].[UnitPrice]*[Quantity]*(1-[Discount])/100)*100) DESC;";

		// create connection
		HRESULT hr = m_spConn.CreateInstance(__uuidof(Connection));
		if (FAILED(hr)) {
			MessageBox("Failed to create Connection object.\r\nPlease make sure ADO is installed.");
			return 0;
		}

		// create recordset
		_RecordsetPtr rs;
		hr = rs.CreateInstance(__uuidof(Recordset));
		if (FAILED(hr)) {
			MessageBox("Failed to create Recordset object.\r\nPlease make sure ADO is installed.");
			return 0;
		}

		// open connection
		hr = m_spConn->raw_Open(_bstr_t("NorthWind"), _bstr_t(""), _bstr_t(""), -1);
		if (FAILED(hr)) {
			MessageBox("Failed to open connection.\r\nPlease make sure you have a 'NorthWind' ODBC data source installed.");
			return 0;
		}

		// open recordset
		hr = rs->raw_Open(strSQL, _variant_t((IUnknown*)m_spConn),
                          adOpenUnspecified, adLockUnspecified, 0);
		if (FAILED(hr)) {
			MessageBox("Failed to open recordset.\r\nPlease check SQL syntax.");
			return 0;
		}

		// get data source from recordset
		DataSourcePtr dataSource = rs->GetDataSource();
		if (dataSource == NULL) {
			MessageBox("Failed to get DataSource from Recordset object. This shouldn't happen!");
			return 0;
		}

		// bind grid to data source
		m_spGrid->DataMode = flexDMBound;
		m_spGrid->DataSource = dataSource;

		// more initialization on the AfterDataRefresh event
		m_spGrid->AllowSelection = false;
		m_spGrid->Editable = flexEDKbdMouse;
		m_spGrid->AllowUserResizing = flexResizeBoth;
		m_spGrid->AllowUserFreezing = flexFreezeBoth;
		m_spGrid->BackColorFrozen = 14737632;
		m_spGrid->GridLines = flexGridFlatVert;
		m_spGrid->ExtendLastCol = true;
		m_spGrid->ScrollTips = true;
		
		// create the detail window and initialize it (but don't show it yet)
		m_dlgDetail.m_spGridMain = m_spGrid;
		m_dlgDetail.m_spConnMain = m_spConn;
		RECT rc;
		GetClientRect(&rc);
		m_dlgDetail.Create(m_hWnd, rc);

		return 1;  // Let the system set the focus
	}

	LRESULT OnSize(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		// resize controls to fit the form
		RECT rc;
		GetClientRect(&rc);
		InflateRect(&rc, -3, -3);

		// resize controls
		m_wndLabel.MoveWindow(rc.left, rc.top, rc.right, rc.top + 45);
		rc.top = rc.top + 3 + 45;
		m_wndGrid.MoveWindow(&rc);

		// done
		return 0;
	}

	LRESULT OnClose(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		if (m_spConn) m_spConn->Close();
		EndDialog(IDOK);
		return 0;
	}

	
////////////////////////////////////////////////////////////////////////////////////////
//
// Event Handlers
//
////////////////////////////////////////////////////////////////////////////////////////
BEGIN_SINK_MAP(CMainDlg)
	//Make sure the Event Handlers have __stdcall calling convention
	SINK_ENTRY(IDC_FGMAIN, 0x6, OnBeforeRowColChange)
	SINK_ENTRY(IDC_FGMAIN, 0x15, OnBeforeEdit)
	SINK_ENTRY(IDC_FGMAIN, 0x52, OnAfterDataRefresh)
	SINK_ENTRY(IDC_FGMAIN, DISPID_DBLCLICK, OnDblClick)
	SINK_ENTRY(IDC_FGMAIN, 0x23, OnBeforeScrollTip)

END_SINK_MAP()

	VOID __stdcall OnBeforeRowColChange(LONG OldRow, LONG OldCol, LONG NewRow, LONG NewCol, VARIANT_BOOL * Cancel)
	{
		// update DB-like cursor glyph
		m_spGrid->PutTextMatrix(OldRow, 0, _bstr_t(""));
		m_spGrid->PutTextMatrix(NewRow, 0, _bstr_t("0"));
	}

	VOID __stdcall OnBeforeEdit(LONG Row, LONG Col, VARIANT_BOOL * Cancel)
	{
	    // only allow editing the supplier
		if (Col != m_spGrid->GetColIndex(_bstr_t("Supplier")))
			*Cancel = true;
	}

	VOID __stdcall OnAfterDataRefresh()
	{
		// create a new recordset object for building combo lists
		_RecordsetPtr rsList;
		HRESULT hr = rsList.CreateInstance(__uuidof(Recordset));
		if (FAILED(hr)) return;

		// translate product ID into product name
		long lCol;
		_bstr_t strList;
		hr = rsList->raw_Open(_variant_t("Products"), _variant_t((IUnknown*)m_spConn),
                              adOpenUnspecified, adLockUnspecified, 0);
		if (SUCCEEDED(hr)) {
			lCol = m_spGrid->GetColIndex(_bstr_t("Product"));
			strList = m_spGrid->BuildComboList((IDispatch*)rsList, _bstr_t("ProductName"), _bstr_t("ProductID"));
			m_spGrid->PutColComboList(lCol, strList);
			rsList->Close();
		}
    
		// translate supplier ID supplier info (name, contact, phone)
		hr = rsList->raw_Open(_variant_t("Suppliers"), _variant_t((IUnknown*)m_spConn),
                              adOpenUnspecified, adLockUnspecified, 0);
		if (SUCCEEDED(hr)) {

			lCol = m_spGrid->GetColIndex(_bstr_t("Supplier"));
    		strList = m_spGrid->BuildComboList((IDispatch*)rsList, _bstr_t("CompanyName"), _bstr_t("SupplierID"));
	    	m_spGrid->PutColComboList(lCol, strList);

			lCol = m_spGrid->GetColIndex(_bstr_t("Country"));
    		strList = m_spGrid->BuildComboList((IDispatch*)rsList, _bstr_t("Country"), _bstr_t("SupplierID"));
	    	m_spGrid->PutColComboList(lCol, strList);

			lCol = m_spGrid->GetColIndex(_bstr_t("Contact"));
    		strList = m_spGrid->BuildComboList((IDispatch*)rsList, _bstr_t("ContactName, ContactTitle"), _bstr_t("SupplierID"));
	    	m_spGrid->PutColComboList(lCol, strList);

			rsList->Close();
		}
    
		// do an autosize for the mapped columns
		m_spGrid->AutoSize(1, m_spGrid->Cols - 1L);
    
		// use image list to show category pictures
		lCol = m_spGrid->GetColIndex(_bstr_t("Cat"));
        if (lCol > -1) {
    		m_spGrid->PutColImageList(lCol, (LONG)(HIMAGELIST)m_imgCategories);
	    	m_spGrid->PutColAlignment(lCol, flexAlignCenterCenter);
        }
    
		// use another imagelist to show db-like cursor glyph
		m_spGrid->PutColWidth(0, m_spGrid->GetRowHeight(0));
		m_spGrid->PutColImageList(0, (LONG)(HIMAGELIST)m_imgDBCursor);
		m_spGrid->PutColAlignment(0, flexAlignCenterCenter);
		m_spGrid->PutTextMatrix(1, 0, _bstr_t("0"));
    
		// freeze two columns
		m_spGrid->FrozenCols = 2;
	}

	// show orders for this product
	VOID __stdcall OnDblClick()
	{
		// calculate position for detail dialog:
		// start with top left corner relative to grid, in pixels
		POINT pt;
		pt.x = XTwipsToPixels(m_spGrid->GetColWidth(0));
		pt.y = YTwipsToPixels(m_spGrid->CellTop + m_spGrid->GetRowHeight(0));

		// convert to screen coordinates
		m_wndGrid.ClientToScreen(&pt);

		// build rectangle
		RECT rc;
		SetRect(&rc, 0, 0, XTwipsToPixels(5700), YTwipsToPixels(2700));
		OffsetRect(&rc, pt.x + 5, pt.y + 5);

		// show the detail dialog modelessly (it will go away when deactivated)
		SetCursor(LoadCursor(NULL, (LPCTSTR)IDC_WAIT));
		m_dlgDetail.MoveWindow(&rc);
		m_dlgDetail.ShowWindow(SW_NORMAL);
		SetCursor(NULL);
	}

    // set scroll tip text for display:
    // note that since this is a translated field, we need to retrieve
    // the display text using the Cell(flexcpTextDisplay) property.
    // (the text property holds only the customer ID).
    VOID __stdcall OnBeforeScrollTip(LONG Row)
	{
		long lCol = m_spGrid->GetColIndex(_bstr_t("Product"));
		_variant_t var = m_spGrid->GetCell(flexcpTextDisplay, Row, lCol);
        _bstr_t strTip = " Product: " + _bstr_t(V_BSTR(&var)) + " ";
		m_spGrid->ScrollTipText = strTip;
	}


////////////////////////////////////////////////////////////////////////////////////////
//
// Member Variables
//
////////////////////////////////////////////////////////////////////////////////////////
	IVSFlexGridPtr	m_spGrid;
	CAxWindow		m_wndGrid;
	CWindow			m_wndLabel;
	_ConnectionPtr	m_spConn;
	CImageList		m_imgCategories;
	CImageList		m_imgDBCursor;
	CDetailDlg		m_dlgDetail;
};

#endif //__MAINDLG_H_
