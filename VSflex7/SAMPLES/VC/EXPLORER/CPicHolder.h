/////////////////////////////////////////////////////////////////////////////
// CPicHolder - helper class for dealing with picture objects
//

#ifndef _PICHOLDERDECLARED
#define _PICHOLDERDECLARED

class CPicHolder
{
// Constructors
public:
    CPicHolder() { m_pPict = NULL; }
    CPicHolder(LPPICTUREDISP pPict) { m_pPict = NULL; SetPictureDispatch(pPict); }
    ~CPicHolder() { Clear(); }

// Operations
	void Clear()
	{
		if (m_pPict) m_pPict->Release();
		m_pPict = NULL;
	}

	HINSTANCE GetInstanceHandle(HINSTANCE hInst = NULL)
	{
		if (hInst) return hInst;
#ifdef _MFC_VER
		return AfxGetResourceHandle();
#endif
#ifdef _ATL_VER
		return _Module.GetResourceInstance();
#endif
		return NULL;
	}

	BOOL HasPicture()
	{
		return m_pPict != NULL;
	}

	// note: xExt and yExt are picture extents in HIMETRIC units
	BOOL CreateFromMetafile(HMETAFILE hmf, int xExt, int yExt, BOOL bTransferOwnership = false)
    {
	    Clear();
	    PICTDESC pdesc;
	    pdesc.cbSizeofstruct = sizeof(pdesc);
	    pdesc.picType = PICTYPE_METAFILE;
	    pdesc.wmf.hmeta = hmf;
	    pdesc.wmf.xExt = xExt;
	    pdesc.wmf.yExt = yExt;
	    return SUCCEEDED(OleCreatePictureIndirect(&pdesc, IID_IPicture, 
                         bTransferOwnership, (LPVOID*)&m_pPict));
    }

	BOOL CreateFromBitmap(HBITMAP hbm, HPALETTE hpal = NULL, BOOL bTransferOwnership = false)
	{
		Clear();
		PICTDESC pdesc;
		pdesc.cbSizeofstruct = sizeof(pdesc);
		pdesc.picType = PICTYPE_BITMAP;
		pdesc.bmp.hbitmap = hbm;
		pdesc.bmp.hpal = hpal;
		return SUCCEEDED(OleCreatePictureIndirect(&pdesc, IID_IPicture,
						 bTransferOwnership, (LPVOID*)&m_pPict));
	}

	BOOL CreateFromIcon(HICON hicon, BOOL bTransferOwnership = false)
	{
		Clear();
		PICTDESC pdesc;
		pdesc.cbSizeofstruct = sizeof(pdesc);
		pdesc.picType = PICTYPE_ICON;
		pdesc.icon.hicon = hicon;
		return SUCCEEDED(OleCreatePictureIndirect(&pdesc, IID_IPicture,
						 bTransferOwnership, (LPVOID*)&m_pPict));
	}

	BOOL CreateFromIcon(UINT nID, HINSTANCE hInst = NULL)
	{
		// use LoadImage instead of LoadIcon to get icons of arbitrary size
		HICON hIcon = (HICON)LoadImage(GetInstanceHandle(hInst), MAKEINTRESOURCE(nID), IMAGE_ICON, 0, 0, 0);
		if (!hIcon) return false;
		return CreateFromIcon(hIcon, true);
	}

	BOOL CreateFromBitmap(UINT nID, HINSTANCE hInst = NULL)
	{
		HBITMAP hBmp = LoadBitmap(GetInstanceHandle(hInst), MAKEINTRESOURCE(nID));
		if (!hBmp) return false;
		return CreateFromBitmap(hBmp, NULL, true);
	}

	BOOL CreateFromClipboard()
	{
		Clear();

        // get ready to work
		HANDLE h;
		BOOL bRet = false;
		OpenClipboard(NULL);

		// try to get a metafile
		h = GetClipboardData(CF_METAFILEPICT);
		if (h) {
			LPMETAFILEPICT pm = (LPMETAFILEPICT)GlobalLock(h);
			bRet = CreateFromMetafile(pm->hMF, pm->xExt, pm->yExt);
			GlobalUnlock(h);
		}

		// try to get a bitmap
		if (!bRet) {
			h = GetClipboardData(CF_BITMAP);
			if (h) bRet = CreateFromBitmap((HBITMAP)h);
		}

		// try for other formats...

		// all done, cleanup and return
		CloseClipboard();
		return bRet;
	}

    BOOL CreateFromFile(LPTSTR szFileName)
    {
		Clear();

	    // load file into memory
        HGLOBAL hGlobal = FileToGlobalHandle(szFileName);
	    if (!hGlobal) return false;

	    // open stream on memory handle
	    IStream*      pStream   = NULL;
	    IPictureDisp* pPictDisp = NULL;
	    if (S_OK == CreateStreamOnHGlobal(hGlobal, false, &pStream))
	    {
		    // load picture from stream
		    OleLoadPicture(pStream, 0, false, IID_IPicture, (void**)&m_pPict);
		    pStream->Release();
	    }

	    // clean up
	    GlobalFree(hGlobal);

	    // done
	    return (m_pPict != NULL);
    }

	short GetType()
    {
	    short sPicType = (short)PICTYPE_UNINITIALIZED;
	    if (m_pPict != NULL) m_pPict->get_Type(&sPicType);
	    return sPicType;
    }
        
	LPPICTUREDISP GetPictureDispatch(BOOL bAddRef = true)
    {
	    LPPICTUREDISP pPictDisp = NULL;
	    if (m_pPict) m_pPict->QueryInterface(IID_IPictureDisp, (LPVOID*)&pPictDisp);
		if (pPictDisp && !bAddRef) pPictDisp->Release();
	    return pPictDisp;
    }
        
	void SetPictureDispatch(LPPICTUREDISP pDisp)
    {
	    LPPICTURE pPict = NULL;
	    Clear();
	    if (pDisp && SUCCEEDED(pDisp->QueryInterface(IID_IPicture, (LPVOID*)&pPict)))
		    m_pPict = pPict;
    }

    void Render(HDC hDC, LPRECT rcRender, LPRECT rcBounds)
    {
	    if (!m_pPict) return;
        long hmWidth;
		long hmHeight;
		m_pPict->get_Width(&hmWidth);
		m_pPict->get_Height(&hmHeight);
		m_pPict->Render(hDC, rcRender->left, rcRender->top,
			rcRender->right - rcRender->left, rcRender->bottom - rcRender->top,
            0, hmHeight-1, hmWidth, -hmHeight, rcBounds);
	}

// Attributes
	LPPICTURE m_pPict;

private:

    HGLOBAL FileToGlobalHandle(LPTSTR szFileName)
    {
	    // open file assuming szFileName is fully qualified 
	    HFILE f = _lopen(szFileName, OF_READ);
        
		// if failed, look in app directory
		if (HFILE_ERROR == f) {

			// build new file name
			TCHAR szBuff[512];
			GetModuleFileName(GetInstanceHandle(), szBuff, 512);
			LPTSTR sz = _tcsrchr(szBuff, '\\');
			if (sz) sz[1] = 0;
			_tcscat(szBuff, szFileName);

			// try again
			f = _lopen(szBuff, OF_READ);
		
			// if still failed, look one level up from app directory
			// this is good while debugging because the EXE is usually in a project subdir)
			if (HFILE_ERROR == f) {

				// build new file name
				sz = _tcsrchr(szBuff, '\\');
				if (sz) *sz = 0;
				sz = _tcsrchr(szBuff, '\\');
				if (sz) sz[1] = 0;
				_tcscat(szBuff, szFileName);

				// try again
				f = _lopen(szBuff, OF_READ);
			}
		}

		// still failed? bail out
		if (HFILE_ERROR == f) return NULL;

	    // allocate buffer
	    long len = _llseek(f, 0, FILE_END);
	    HGLOBAL hGlobal = GlobalAlloc(GHND, len);

	    // read file into buffer
	    if (hGlobal) {
		    LPTSTR sz = (LPTSTR)GlobalLock(hGlobal);
    	    _llseek(f, 0, FILE_BEGIN);
		    _lread(f, sz, len);
		    GlobalUnlock(hGlobal);
	    }

	    // close file
        _lclose(f);

	    // return handle (caller must free it)
	    return hGlobal;
    }
};

#endif // _PICHOLDERDECLARED
