/////////////////////////////////////////////////////////////////////////////
// TheData - this class contains the sample Link and Node data
//

#ifndef _THEDATADECLARED
#define _THEDATADECLARED

typedef struct tagNODE
{
	double x, y, z;		// node coordinates
	TCHAR  szName[30];	// node name
} NODE;

typedef struct tagLINK
{
	int i, j;			// connectivity
	TCHAR  szName[30];	// link name
	int iSec;			// section type
    double u, v, w;		// section orientation
} LINK;

class CTheData
{
enum { NUM_NODES = 5000 };
enum { NUM_LINKS = 10000 };

// Constructors
public:
    CTheData()  { InitData(); }
    ~CTheData() { }

// Operations
	int GetNodeCount()
	{
		return NUM_NODES;
	}

	int GetLinkCount()
	{
		return NUM_LINKS;
	}

	void InitData()
	{
		// initialize node data
		for (int i = 0; i < NUM_NODES; i++) {
			m_Nodes[i].x = rand() % 700;
			m_Nodes[i].y = rand() % 700;
			m_Nodes[i].z = rand() % 700;
			wsprintf(m_Nodes[i].szName, "Node %d", i);
        }
    
		// initialize link data
		for (i = 0; i < NUM_LINKS; i++) {
			do {
				m_Links[i].i = rand() % NUM_NODES;
				m_Links[i].j = rand() % NUM_NODES;
			} while (m_Links[i].i == m_Links[i].j);
			wsprintf(m_Links[i].szName, "Link %d", i);
			m_Links[i].iSec = 1000 + (rand() % 1000);
            m_Links[i].u = (rand() % 1000) / 1000.0;
            m_Links[i].v = (rand() % 1000) / 1000.0;
            m_Links[i].w = (rand() % 1000) / 1000.0;
		}
	}

// Attributes
	NODE m_Nodes[NUM_NODES];
	LINK m_Links[NUM_LINKS];
};

#endif // _THEDATADECLARED
