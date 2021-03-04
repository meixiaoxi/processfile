#ifndef __WORK_THREAD_HEAD
#define __WORK_THREAD_HEAD

struct threadInfo
{
	HWND	hWnd;//主窗口句柄，用于消息的发送
	CWinThread *pThread;
	char	num[20];
	char	testid[33];
	char    station;
	char	workmode; // I2C or uart
	char	portnum;  // 串口号
	char	fulluartlog;
	char	checkresult[2];
	char	testdata[8];
	char    isExcelAppCreate;
	char	isSaveFileCreate;
	CString opCode;
	CString workStationCode;
	CString shiftCode;
	CString testPersonCode;
	CString testStartTime;
	CString testEndTime;
	CString testResult;
};


UINT  ThreadFunc(LPVOID  pParm);
#endif