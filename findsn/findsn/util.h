#ifndef __UTIL_H__
#define __UTIL_H__
#include <stdio.h>

#define INI_FILE_NAME	"./init.ini"

class CUtil
{
public:
	static void hexstr2buf(char *str, char *buf);
	static __int64 str2hex(const char *ch);  /* 字符串转16进制数 */
	static BOOL DrawBmpFromFile(HWND hwnd, int ctrlId, CString strpicname);
	static void lymTraceBuf(char *buffer, int len);
	//static BOOL CheckPass(char *val, char *range);
	static BOOL CheckPass(double value, char *range);
	static void lymSaveLog(char *str,char* path);
	static void lymSaveLogBuf(char *buf, int len, char *out,char* path);
	static void SaveInfoFile(char *buf,char* path);
	static void lymMoveFile(char* path);
protected:
private:
};

#endif