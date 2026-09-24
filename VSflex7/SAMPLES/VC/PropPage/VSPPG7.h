/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Thu Jan 06 10:01:57 2000
 */
/* Compiler settings for E:\Ocx\RandD\vsflex7\Samples\VC\PROPPAGE\VSPPG7.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __VSPPG7_h__
#define __VSPPG7_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __ICategorizeProperties_FWD_DEFINED__
#define __ICategorizeProperties_FWD_DEFINED__
typedef interface ICategorizeProperties ICategorizeProperties;
#endif 	/* __ICategorizeProperties_FWD_DEFINED__ */


#ifndef __VSPPG_FWD_DEFINED__
#define __VSPPG_FWD_DEFINED__

#ifdef __cplusplus
typedef class VSPPG VSPPG;
#else
typedef struct VSPPG VSPPG;
#endif /* __cplusplus */

#endif 	/* __VSPPG_FWD_DEFINED__ */


#ifndef __VSPPGFlex_FWD_DEFINED__
#define __VSPPGFlex_FWD_DEFINED__

#ifdef __cplusplus
typedef class VSPPGFlex VSPPGFlex;
#else
typedef struct VSPPGFlex VSPPGFlex;
#endif /* __cplusplus */

#endif 	/* __VSPPGFlex_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 


#ifndef __VSPPG7Lib_LIBRARY_DEFINED__
#define __VSPPG7Lib_LIBRARY_DEFINED__

/* library VSPPG7Lib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_VSPPG7Lib;

#ifndef __ICategorizeProperties_INTERFACE_DEFINED__
#define __ICategorizeProperties_INTERFACE_DEFINED__

/* interface ICategorizeProperties */
/* [unique][uuid][local][object] */ 

typedef /* [public] */ int PROPCAT;

#define	PROPCAT_Nil	( -1 )

#define	PROPCAT_Misc	( -2 )

#define	PROPCAT_Font	( -3 )

#define	PROPCAT_Position	( -4 )

#define	PROPCAT_Appearance	( -5 )

#define	PROPCAT_Behavior	( -6 )

#define	PROPCAT_Data	( -7 )

#define	PROPCAT_List	( -8 )

#define	PROPCAT_Text	( -9 )

#define	PROPCAT_Scale	( -10 )

#define	PROPCAT_DDE	( -11 )


EXTERN_C const IID IID_ICategorizeProperties;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("4d07fc10-f931-11ce-b001-00aa006884e5")
    ICategorizeProperties : public IUnknown
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE MapPropertyToCategory( 
            /* [in] */ DISPID dispid,
            /* [out] */ PROPCAT __RPC_FAR *ppropcat) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetCategoryName( 
            /* [in] */ PROPCAT propcat,
            /* [in] */ LCID lcidcat,
            /* [out] */ BSTR __RPC_FAR *pbstrname) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICategorizePropertiesVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICategorizeProperties __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICategorizeProperties __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICategorizeProperties __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *MapPropertyToCategory )( 
            ICategorizeProperties __RPC_FAR * This,
            /* [in] */ DISPID dispid,
            /* [out] */ PROPCAT __RPC_FAR *ppropcat);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetCategoryName )( 
            ICategorizeProperties __RPC_FAR * This,
            /* [in] */ PROPCAT propcat,
            /* [in] */ LCID lcidcat,
            /* [out] */ BSTR __RPC_FAR *pbstrname);
        
        END_INTERFACE
    } ICategorizePropertiesVtbl;

    interface ICategorizeProperties
    {
        CONST_VTBL struct ICategorizePropertiesVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICategorizeProperties_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICategorizeProperties_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICategorizeProperties_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICategorizeProperties_MapPropertyToCategory(This,dispid,ppropcat)	\
    (This)->lpVtbl -> MapPropertyToCategory(This,dispid,ppropcat)

#define ICategorizeProperties_GetCategoryName(This,propcat,lcidcat,pbstrname)	\
    (This)->lpVtbl -> GetCategoryName(This,propcat,lcidcat,pbstrname)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT STDMETHODCALLTYPE ICategorizeProperties_MapPropertyToCategory_Proxy( 
    ICategorizeProperties __RPC_FAR * This,
    /* [in] */ DISPID dispid,
    /* [out] */ PROPCAT __RPC_FAR *ppropcat);


void __RPC_STUB ICategorizeProperties_MapPropertyToCategory_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICategorizeProperties_GetCategoryName_Proxy( 
    ICategorizeProperties __RPC_FAR * This,
    /* [in] */ PROPCAT propcat,
    /* [in] */ LCID lcidcat,
    /* [out] */ BSTR __RPC_FAR *pbstrname);


void __RPC_STUB ICategorizeProperties_GetCategoryName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICategorizeProperties_INTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_VSPPG;

#ifdef __cplusplus

class DECLSPEC_UUID("2C4CDB4D-6162-11D3-BD96-B76E8682527F")
VSPPG;
#endif

EXTERN_C const CLSID CLSID_VSPPGFlex;

#ifdef __cplusplus

class DECLSPEC_UUID("2C4CDB4E-6162-11D3-BD96-B76E8682527F")
VSPPGFlex;
#endif
#endif /* __VSPPG7Lib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
