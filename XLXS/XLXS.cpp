// XLXS.cpp : 此文件包含 "main" 函数。程序执行将在此处开始并结束。
//

#include "pch.h"
#include <iostream>
#include <windows.h>
#include "libxl.h"
#include <vector>
#include <time.h>
#include <corecrt_io.h>
#include <corecrt_wstring.h>
#include <algorithm>
#include "threadpool.h"

using namespace libxl;

struct ExcelData
{
	CellType type;
	double num;
	const wchar_t* str;
};

struct FileData
{
	std::string pach;
	uint64_t filesize;
};

// wchar_t to string
void Wchar_tToString(std::string& szDst, wchar_t *wchar)
{
	wchar_t * wText = wchar;
	DWORD dwNum = WideCharToMultiByte(CP_OEMCP, NULL, wText, -1, NULL, 0, NULL, FALSE);// WideCharToMultiByteµƒ‘À”√
	char *psText; // psTextŒ™char*µƒ¡Ÿ ± ˝◊È£¨◊˜Œ™∏≥÷µ∏¯std::stringµƒ÷–º‰±‰¡ø
	psText = new char[dwNum];
	WideCharToMultiByte(CP_OEMCP, NULL, wText, -1, psText, dwNum, NULL, FALSE);// WideCharToMultiByteµƒ‘Ÿ¥Œ‘À”√
	szDst = psText;// std::string∏≥÷µ
	delete[]psText;// psTextµƒ«Â≥˝
}

//Converting a WChar string to a Ansi string   
std::string WChar2Ansi(LPCWSTR pwszSrc)
{
	int nLen = WideCharToMultiByte(CP_ACP, 0, pwszSrc, -1, NULL, 0, NULL, NULL);
	if (nLen <= 0) return std::string("");
	char* pszDst = new char[nLen];
	if (NULL == pszDst) return std::string("");
	WideCharToMultiByte(CP_ACP, 0, pwszSrc, -1, pszDst, nLen, NULL, NULL);
	pszDst[nLen - 1] = 0;
	std::string strTemp(pszDst);
	delete[] pszDst;
	return strTemp;
}
std::string ws2s(std::wstring& inputws)
{
	return WChar2Ansi(inputws.c_str());
}

//Converting a Ansi string to WChar string   
std::wstring Ansi2WChar(LPCSTR pszSrc, int nLen)
{
	int nSize = MultiByteToWideChar(CP_ACP, 0, (LPCSTR)pszSrc, nLen, 0, 0);
	if (nSize <= 0) return NULL;
	WCHAR *pwszDst = new WCHAR[nSize + 1];
	if (NULL == pwszDst) return NULL;
	MultiByteToWideChar(CP_ACP, 0, (LPCSTR)pszSrc, nLen, pwszDst, nSize);
	pwszDst[nSize] = 0;
	if (pwszDst[0] == 0xFEFF) // skip Oxfeff   
		for (int i = 0; i < nSize; i++)
			pwszDst[i] = pwszDst[i + 1];
	std::wstring wcharString(pwszDst);
	delete pwszDst;
	return wcharString;
}
std::wstring s2ws(const std::string& s)
{
	return Ansi2WChar(s.c_str(), s.size());
}

void split(const std::string &s, const std::string &seperator, std::vector<std::string>& result) 
{
	typedef std::string::size_type string_size;
	string_size i = 0;

	while (i != s.size()) {
		//找到字符串中首个不等于分隔符的字母；
		int flag = 0;
		while (i != s.size() && flag == 0) {
			flag = 1;
			for (string_size x = 0; x < seperator.size(); ++x)
				if (s[i] == seperator[x]) {
					++i;
					flag = 0;
					break;
				}
		}

		//找到又一个分隔符，将两个分隔符之间的字符串取出；
		flag = 0;
		string_size j = i;
		while (j != s.size() && flag == 0) {
			for (string_size x = 0; x < seperator.size(); ++x)
				if (s[j] == seperator[x]) {
					flag = 1;
					break;
				}
			if (flag == 0)
				++j;
		}
		if (i != j) {
			result.push_back(s.substr(i, j - i));
			i = j;
		}
	}
}


uint32_t getRandomInField(uint32_t begin, uint32_t end)
{
	int nField = end - begin + 1;

	if (nField <= 0) 
	{
		return begin;
	}

	int nValue = rand() % nField;

	nValue += begin;

	return nValue;
}

uint32_t getRandomInFieldEx(uint32_t begin, uint32_t end)
{
	static int nGlobalRandIndex = 0;
	static int vtGlobalRankValue[10000];
	static bool bInit = false;
	srand((unsigned)time(NULL) + nGlobalRandIndex);
	if (bInit == false)
	{
		bInit = true;
		nGlobalRandIndex = 0;
		int nCurIndex;
		int nTemp;

		for (int j = 0; j < 10000; j++)
		{
			vtGlobalRankValue[j] = j + 1;
		}

		for (int i = 0; i < 10000; i++)
		{
			nCurIndex = rand() % (i + 1);
			if (nCurIndex != i)
			{
				nTemp = vtGlobalRankValue[i];
				vtGlobalRankValue[i] = vtGlobalRankValue[nCurIndex];
				vtGlobalRankValue[nCurIndex] = nTemp;
			}
		}
	}

	return  begin + vtGlobalRankValue[(nGlobalRandIndex++) % 10000] % (end - begin);
}


int randSelectQuestion(int from, int to)
{
	srand((unsigned)time(NULL));
	int span = 0;
	if (to > from)
	{
		span = to - from;
		return from + rand() % span;
	}
	else
	{
		span = from - to;
		return to + rand() % span;
	}
}

std::string rand_str(const int len)
{
	std::string str;
	srand(time(NULL));
	int i;
	for (i = 0; i < len; ++i)
	{
		switch ((rand() % 3))
		{
		case 1:
			str += 'A' + getRandomInFieldEx(0,26);
			break;
		case 2:
			str += 'a' + getRandomInFieldEx(0, 26);
			break;
		default:
			str += '0' + getRandomInFieldEx(0, 10);
			break;
		}
	}
	str += '\0';
	return str;
}

uint64_t GetFileSize(std::string& patch)
{
	HANDLE handle = CreateFile(s2ws(patch).c_str(), FILE_READ_EA,
		FILE_SHARE_READ, 0, OPEN_EXISTING, 0, 0);
	if (handle != INVALID_HANDLE_VALUE)
	{
		uint64_t size = GetFileSize(handle, NULL);
		//std::cout << "CreateFile:  " << size/1024 << std::endl;
		CloseHandle(handle);
		return size;
	}
	return 0;
}

void getAllFiles(std::string path, std::vector<FileData>& files)
{
	// 文件句柄
	long long hFile = 0;
	// 文件信息
	struct _finddata_t fileinfo;
	std::string p;
	if ((hFile = _findfirst(p.assign(path).append("\\*").c_str(), &fileinfo)) != -1)
		// "\\*"是指读取文件夹下的所有类型的文件，若想读取特定类型的文件，以png为例，则用“\\*.png”
	{
		do
		{
			//如果是目录,迭代之  
			//如果不是,加入列表  
			if ((fileinfo.attrib &  _A_SUBDIR))
			{
				
			}
			else
			{
				FileData filedata;
				filedata.pach = path + "\\" + fileinfo.name;
				filedata.filesize = GetFileSize(filedata.pach);
				files.push_back(filedata);
			}
		} while (_findnext(hFile, &fileinfo) == 0);
		_findclose(hFile);
	}
}



void Init(std::string pach);



uint32_t InitExcel(std::string& file, uint32_t multiple = 0);
void CreateNewExcel(std::string& file, const wchar_t* sheetname, std::vector<std::vector<ExcelData>>&ExcelDataVec, uint32_t h, uint32_t l , uint32_t multiple);
void WriteData(Sheet* sheet, uint32_t index ,std::vector<ExcelData>& data);
void Print(Sheet* sheet, uint32_t h, uint32_t l);


void Init(std::string pach)
{
	std::cout << "输入线程数 : ";
	uint32_t threadnum = 4;
	std::cin >> threadnum;

	std::cout << "输入数据新增数据倍数 : ";
	uint32_t multiple = 1;
	std::cin >> multiple;

	ThreadPool pool(threadnum);
	std::vector< std::future<uint32_t> > results;
	std::wcout.imbue(std::locale("chs"));
	std::vector<FileData> files;
	getAllFiles(pach, files);
	std::sort(files.begin(), files.end(), [](FileData &a, FileData &b) {return a.filesize > b.filesize; });

	for(auto& it : files)
	{
		results.emplace_back(
			pool.enqueue([&] {
			return InitExcel(it.pach, multiple);
		})
		);
	}

	uint64_t total = 0;
	for (auto && result : results)
	{
		total += result.get();
	}
	std::cout << " 插入数据量 : " << total << std::endl;

}

void Print(Sheet* sheet, uint32_t h, uint32_t l)
{
	auto celltype = sheet->cellType(h, l);
	if (celltype == libxl::CELLTYPE_STRING)
	{
		const wchar_t* s = sheet->readStr(h, l);
		std::wcout << s << std::endl;
	}
	if (celltype == libxl::CELLTYPE_NUMBER)
	{
		std::wcout << sheet->readNum(h, l) << std::endl;
	}
}

void WriteData(Sheet* sheet, uint32_t index, std::vector<ExcelData>& data)
{
	for (auto i = 0; i < data.size(); i++)
	{
		switch (data[i].type)
		{
		case libxl::CELLTYPE_STRING:
		{
			sheet->writeStr(i, index, data[i].str);
		}
		break;
		case libxl::CELLTYPE_NUMBER:
		{
			sheet->writeNum(i, index, data[i].num);
		}
		break;
		default:
			break;
		}
	}
}

void CreateNewExcel(std::string& file,const wchar_t* sheetname, std::vector<std::vector<ExcelData>>&ExcelDataVec, uint32_t h, uint32_t l, uint32_t multiple)
{
	Book* book = xlCreateBook();
	book->setKey(L"GCCG", L"windows-282123090cc0e6036db16b60a1o3p0h9");
	Sheet* sheet = book->addSheet(sheetname);
	std::vector<uint32_t>randtemp;
	uint32_t ranknum = ExcelDataVec.size();
	for (auto i = 0; i < ranknum; i++)
	{
		if (i > 0)
		{
			randtemp.push_back(i);
		}
	}
	WriteData(sheet, 0, ExcelDataVec[0]);
	
	uint32_t index = 0;
	while (!randtemp.empty())
	{
		index++;
		uint32_t tempindex = getRandomInFieldEx(0, randtemp.size());
		WriteData(sheet, index, ExcelDataVec[randtemp[tempindex]]);
		randtemp.erase(randtemp.begin() + tempindex);
	}

	for (uint32_t i = l; i < (l + l * multiple); i++)
	{
		for (auto k = 0; k < h; k++)
		{
			if (k == 1)
			{
				sheet->writeStr(k, i, s2ws("s_"+rand_str(20)).c_str());
			}
			else
			{
				sheet->writeStr(k, i, s2ws(rand_str(20)).c_str());
			}
		}
	}

	book->save(s2ws("des\\"+ file).c_str());
	//::ShellExecute(NULL, L"open", L"example.xls", NULL, NULL, SW_SHOW);
	book->release();
}

uint32_t InitExcel(std::string& file,uint32_t multiple)
{
	clock_t start = clock();		//程序开始计时
	//std::cout << " start " << file.c_str() <<"  threadid : "<< std::this_thread::get_id() << std::endl;
	Book* book = xlCreateBook();
	book->setKey(L"GCCG", L"windows-282123090cc0e6036db16b60a1o3p0h9");
	book->load(s2ws(file).c_str());
	Sheet* sheet = book->getSheet(0); 
	//std::wcout << sheet->name() << std::endl;
	//std::cout << sheet->lastRow() << std::endl;
	//std::cout << sheet->lastCol() << std::endl;
	
	const wchar_t* s = sheet->readStr(1, 0);
	if (s == nullptr || wcscmp(s, L"i_id"))
	{
		std::cout << "error excel : "<< file.c_str() << std::endl;
	}

	uint32_t l = sheet->lastCol();
	for (int c = sheet->firstCol(); c < sheet->lastCol(); c++)
	{
		//Print(sheet, 1, c);

		auto celltype = sheet->cellType(1, c);
		if (celltype != libxl::CELLTYPE_STRING)
		{
			l = c;
			//std::cout << " 有效列 : " << l << std::endl;
			break;
		}
	}

	uint32_t h = sheet->lastRow();
	for (int r = sheet->firstRow(); r < sheet->lastRow(); r++)
	{
		auto celltype = sheet->cellType(r, 0);
		if (celltype != libxl::CELLTYPE_NUMBER && celltype != libxl::CELLTYPE_STRING && r > 1)
		{
			h = r;
			//std::cout << " 有效行 : " << h << std::endl;
			//Print(sheet, r, 0);
			break;
		}
	}

	std::vector<std::vector<ExcelData>>ExcelDataVec;
	for (int c = sheet->firstCol(); c < l; c++)
	{
		std::vector<ExcelData > colDataVec;
		for (int r = sheet->firstRow(); r < h; r++)
		{
			auto celltype = sheet->cellType(r, c);

			switch (celltype)
			{
			case libxl::CELLTYPE_NUMBER:
			{
				double d = sheet->readNum(r, c);
				ExcelData data;
				data.type = celltype;
				data.num = d;
				colDataVec.push_back(data);
			}
				break;
			case libxl::CELLTYPE_STRING:
			{
				const wchar_t* s = sheet->readStr(r, c);
				ExcelData data;
				data.type = celltype;
				data.str = s;
				colDataVec.push_back(data);
			}
				break;

			
			break;
			default:
			{
				ExcelData data;
				data.type = libxl::CELLTYPE_STRING;
				data.str = L" ";
				colDataVec.push_back(data);
				//std::cout << "error cellType : " << celltype<<"    file : " <<file.c_str()  <<"            r : "<< r <<"             c : " <<c<<std::endl;
				//system("pause");
			}
				break;
			}
		}
		ExcelDataVec.push_back(colDataVec);
	}
	std::vector<std::string>result;
	split(file,"\\", result);
	CreateNewExcel(result[1], sheet->name(),ExcelDataVec, h, l, multiple);
	book->release();

	clock_t end = clock();		//程序结束用时
	double endtime = (double)(end - start) / CLOCKS_PER_SEC;
	std::cout  <<" used time:" << endtime <<" s " << result[1].c_str() <<std::endl;		//s为单位

	return h * (multiple * l);
}

int main()
{
	clock_t start = clock();		
	Init("src");
	clock_t end = clock();		
	double endtime = (double)(end - start) / CLOCKS_PER_SEC;
	std::cout <<"总耗时:" << endtime << " s" << std::endl;	

	system("pause");
	return 0;
}

