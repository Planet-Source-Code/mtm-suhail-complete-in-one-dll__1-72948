//	mtm_supp.h : supplement header file for mtm.dll
//	Author: DLL to Lib version 3.00
//	Date: Monday, June 29, 2009
//	Description: The declaration of the mtm.dll's entry-point function.
//	Prototype: BOOL WINAPI xxx_DllMain(HINSTANCE hinstance, DWORD fdwReason, LPVOID lpvReserved);
//	Parameters: 
//		hinstance
//		  Handle to current instance of the application. Use AfxGetInstanceHandle()
//		  to get the instance handle if your project has MFC support.
//		fdwReason
//		  Specifies a flag indicating why the entry-point function is being called.
//		lpvReserved 
//		  Specifies further aspects of DLL initialization and cleanup. Should always
//		  be set to NULL;
//	Comment: Please see the help document for detail information about the entry-point 
//		 function
//	Homepage: http://www.binary-soft.com
//	Technical Support: support@binary-soft.com
/////////////////////////////////////////////////////////////////////

#if !defined(D2L_MTM_SUPP_H__46877657_1906_3998_46F3_219A43AA2672__INCLUDED_)
#define D2L_MTM_SUPP_H__46877657_1906_3998_46F3_219A43AA2672__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifdef __cplusplus
extern "C" {
#endif


#include <windows.h>

#include <objbase.h>

/* The following function provides COM support to the static lib converted from mtm.dll. 
You can call it to check if there are any existing references to COM objects managed by the lib. */

STDAPI  MTM_DllCanUnloadNow(void);

/* The following function provides COM support to the static lib converted from mtm.dll. 
You can call it to retrieve the class object from the lib. */

STDAPI  MTM_DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID FAR* ppv);

/* The following function provides COM support to the static lib converted from mtm.dll. 
Usually it is not necessary to call it directly. */

STDAPI MTM_DllRegisterServer(void);

/* The following function provides COM support to the static lib converted from mtm.dll. 
Usually it is not necessary to call it directly. */

STDAPI MTM_DllUnregisterServer(void);

/* This is mtm.dll's entry-point function. You should call it to do necessary
 initialization and finalization. */

BOOL WINAPI MTM_DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved);


#ifdef __cplusplus
}
#endif

#endif // !defined(D2L_MTM_SUPP_H__46877657_1906_3998_46F3_219A43AA2672__INCLUDED_)