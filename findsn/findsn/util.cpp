
#include "StdAfx.h"
#include "Util.h"


char headFlag = 0;
void CUtil::SaveInfoFile(char *buf,char* path)
{
	CStdioFile file;
	char str_cid[10]={0};
	char str_mid[33]={0};
	CString str;
	file.Open(path,CFile::modeCreate | CFile::modeNoTruncate | CFile::modeWrite);

	if(file == NULL)
		return;

	file.SeekToEnd();

	CTime tm=CTime::GetCurrentTime();

	if(headFlag == 0)
	{
		file.WriteString("DATE,TIME,SN\n");
		headFlag = 1;
	}

	strncpy(str_cid,buf,9);

	str_cid[9] = 0;

	str.Format(_T("%d/%d/%d")_T(",%d:%d:%d")_T(",%s")_T("\n"),
			tm.GetYear(),tm.GetMonth(),tm.GetDay(),tm.GetHour(),tm.GetMinute(),tm.GetSecond(),str_cid);
	file.WriteString(str);

file.Close();

}

void CUtil::lymSaveLog(char *str,char* path)
{
	int debug = GetPrivateProfileInt("App", "Debug", 0, INI_FILE_NAME);
//	if (debug == 0)
//		return;

	FILE *fp = fopen(path, "ab");
	fwrite(str, strlen(str), 1, fp);
	fclose(fp);
}