#include "StdAfx.h"
#include "workthread.h"

UINT  ThreadFunc(LPVOID  pParm)
{
	threadInfo *pInfo=(threadInfo*)pParm;
	HWND hWnd = pInfo->hWnd;


	 return 0;
}