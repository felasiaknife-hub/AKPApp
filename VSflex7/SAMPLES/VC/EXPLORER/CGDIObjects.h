/////////////////////////////////////////////////////////////////////////////
// CGDIObjects - helper class for dealing with GDI objects
//

#ifndef _GDIOBJECTSDECLARED
#define _GDIOBJECTSDECLARED

//////////////////////////////////////////////////////////////////////////////////////////////////////////
//

class CGDIBrush
{
// Constructors
public:
    CGDIBrush()  { m_hBrush = NULL; }
    ~CGDIBrush() { Clear(); }

// Operations
	void Clear()
	{
		if (m_hBrush) DeleteObject(m_hBrush);
		m_hBrush = NULL;
	}

	HBRUSH Create(long clr)
	{
		Clear();
		OleTranslateColor(clr, NULL, (COLORREF*)&clr);
		m_hBrush = CreateSolidBrush(clr);
		return m_hBrush;
	}

	HBRUSH GetBrushHandle()
	{
		return m_hBrush;
	}

// Attributes
	HBRUSH m_hBrush;
};

//////////////////////////////////////////////////////////////////////////////////////////////////////////
//

class CGDIFont
{
// Constructors
public:
    CGDIFont()  { m_hFont = NULL; }
    ~CGDIFont() { Clear(); }

// Operations
	void Clear()
	{
		if (m_hFont) DeleteObject(m_hFont);
		m_hFont = NULL;
	}

	HFONT Create(LPCTSTR szFaceName, float fPointSize, BOOL bBold = false, BOOL bItalic = false, BOOL bUnderline = false)
	{
		// clean up
		Clear();

		// create logical font
		LOGFONT lf;
		memset(&lf, 0, sizeof(lf));
		_tcscpy(lf.lfFaceName, szFaceName);
		lf.lfItalic = bItalic;
		lf.lfUnderline = bUnderline;
		lf.lfWeight = (bBold)? FW_BOLD: FW_NORMAL;
		HDC hDC = GetDC(NULL);
		lf.lfHeight = -MulDiv((int)(fPointSize * 10), GetDeviceCaps(hDC, LOGPIXELSY), 720);
		ReleaseDC(NULL, hDC);

		// create GDI font
		m_hFont = CreateFontIndirect(&lf);

		// done
		return m_hFont;
	}

	HFONT GetFontHandle()
	{
		return m_hFont;
	}

// Attributes
	HFONT m_hFont;
};

#endif // _GDIOBJECTSDECLARED
