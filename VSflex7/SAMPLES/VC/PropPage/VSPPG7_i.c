/* this file contains the actual definitions of */
/* the IIDs and CLSIDs */

/* link this file in with the server and any clients */


/* File created by MIDL compiler version 5.01.0164 */
/* at Thu Jan 06 10:01:57 2000
 */
/* Compiler settings for E:\Ocx\RandD\vsflex7\Samples\VC\PROPPAGE\VSPPG7.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )
#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

const IID LIBID_VSPPG7Lib = {0x2C4CDB41,0x6162,0x11D3,{0xBD,0x96,0xB7,0x6E,0x86,0x82,0x52,0x7F}};


const IID IID_ICategorizeProperties = {0x4d07fc10,0xf931,0x11ce,{0xb0,0x01,0x00,0xaa,0x00,0x68,0x84,0xe5}};


const CLSID CLSID_VSPPG = {0x2C4CDB4D,0x6162,0x11D3,{0xBD,0x96,0xB7,0x6E,0x86,0x82,0x52,0x7F}};


const CLSID CLSID_VSPPGFlex = {0x2C4CDB4E,0x6162,0x11D3,{0xBD,0x96,0xB7,0x6E,0x86,0x82,0x52,0x7F}};


#ifdef __cplusplus
}
#endif

