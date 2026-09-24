/////////////////////////////////////////////////////////////////////////////
// CFileList - helper class for listing files/directories
//

#ifndef _FILELISTDECLARED
#define _FILELISTDECLARED

//////////////////////////////////////////////////////////////////////////////////////////////////////////
//

static int __cdecl qsortCompare(const void *elem1, const void *elem2 ) 
{
	LPWIN32_FIND_DATA pfd1 = (LPWIN32_FIND_DATA)elem1;
	LPWIN32_FIND_DATA pfd2 = (LPWIN32_FIND_DATA)elem2;
	return _tcsicmp(pfd1->cFileName, pfd2->cFileName);
}

class CFileList
{
private:
// Attributes
	_bstr_t				m_strPath;
	int					m_nSize, m_nAlloc;
	LPWIN32_FIND_DATA	m_fd;

// Constructors
public:
    CFileList()
	{
		m_fd = NULL;
		m_nSize = m_nAlloc = 0;
	}
    ~CFileList()
	{
		Clear();
	}

// Operations
	void Clear()
	{
		if (m_fd) free(m_fd);
		m_fd = NULL;
		m_nSize = m_nAlloc = 0;
	}

	void GetFiles(_bstr_t strPath)
	{
		SetPath(strPath);
	}

	void GetSubDirs(_bstr_t strPath)
	{
		SetPath(strPath, true);
	}

	_bstr_t GetFileName(long i)
	{
		if (i < 0 || i >= m_nSize) return _bstr_t("");
		return _bstr_t(m_fd[i].cFileName);
	}

	long GetFileSize(long i)
	{
		if (i < 0 || i >= m_nSize) return 0;
		return (m_fd[i].nFileSizeHigh * MAXDWORD) + m_fd[i].nFileSizeLow;
	}

	_bstr_t GetFileDate(long i)
	{
		if (i < 0 || i >= m_nSize) return _bstr_t("");
		TCHAR sz[20];
		SYSTEMTIME sysTime;
		FileTimeToSystemTime(&(m_fd[i].ftLastWriteTime), &sysTime);
		GetDateFormat(LOCALE_SYSTEM_DEFAULT, DATE_SHORTDATE, &sysTime, NULL, sz, 20);
		return _bstr_t(sz);
	}

	_bstr_t GetFullPath(long i)
	{
		if (i < 0 || i >= m_nSize) return _bstr_t("");
		_bstr_t str(m_strPath);
		str += m_fd[i].cFileName;
		return str;
	}

	BOOL HasSubDirs(long i)
	{
		if (i < 0 || i >= m_nSize) return false;
		CFileList fl;
		_bstr_t str = m_strPath;
		str += m_fd[i].cFileName;
		fl.SetPath(str, true, true);
		return (fl.GetSize() > 0);
	}

	long GetSize()
	{
		return m_nSize;
	}

private:
	void SetPath(_bstr_t strPath, BOOL bWantDirs = false, BOOL bSingle = false)
	{
		// clear existing list
		Clear();

		// convert path into usable string
		USES_CONVERSION;
		TCHAR sz[512];
		_tcscpy(sz, W2T(strPath));
		int iLen = lstrlen(sz);
		if (iLen && sz[iLen - 1] != '\\') _tcscat(sz, _T("\\"));

		// save path and append wildcard
		m_strPath = sz;
		_tcscat(sz, _T("*.*"));

		// scan path
		WIN32_FIND_DATA fd;
		HANDLE h = FindFirstFile(sz, &fd);
		if (h != INVALID_HANDLE_VALUE) {
			for (BOOL bMore = true; bMore; bMore = FindNextFile(h, &fd)) {

				// skip junk
				if (*fd.cFileName == '.') continue;

				// keep if it is what we're looking for
				BOOL bIsDir = (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY);
				if ((bIsDir && bWantDirs) || (!bIsDir && !bWantDirs)) {
					AddItem(&fd);
					if (bSingle) break;
				}
			}
			FindClose(h);
		}

		// sort list
		if (m_nSize > 1)
			qsort(m_fd, m_nSize, sizeof(WIN32_FIND_DATA), qsortCompare);
	}

	BOOL AddItem(LPWIN32_FIND_DATA pfd)
	{
		// expand vector if we have to
		if (m_nAlloc <= m_nSize) {
			m_nAlloc += 100;
			m_fd = (LPWIN32_FIND_DATA)realloc(m_fd, sizeof(WIN32_FIND_DATA) * m_nAlloc);
			if (!m_fd) {
				Clear();
				return false;
			}
		}

		// save data in a new slot
		memcpy(&(m_fd[m_nSize]), pfd, sizeof(WIN32_FIND_DATA));
		m_nSize++;

		// done
		return true;
	}
};

#endif _FILELISTDECLARED
