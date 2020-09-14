#include "stdafx.h"
#include <Windows.h>

#define ERR_LEN 1024
#define TEXT_LEN 32767

#ifdef _WIN64
#define DLL_PATH _T("..\\..\\Release\\excel-win-64.dll")
#else
#define DLL_PATH _T("..\\..\\Release\\excel-win-32.dll")
#endif

#define  XLSX_FILE "..\\..\\Release\\1.xlsx"
#define  IMG_FILE "..\\..\\Release\\1.jpg"

int main()
{
	HMODULE module = LoadLibrary(DLL_PATH);
	if (module == NULL)
	{
		printf("Load excel.dll failed\n");
		return -1;
	}

	char output[ERR_LEN];
	memset(output, 0, ERR_LEN);

	//NewFile
	//typedef void(*NewFileFunc)();
	//NewFileFunc NewFile;
	//NewFile = (NewFileFunc)GetProcAddress(module, "NewFile");
	//NewFile();

	//OpenFile
	typedef void(*OpenFileFunc)(char*, char*);
	OpenFileFunc OpenFile;
	OpenFile = (OpenFileFunc)GetProcAddress(module, "OpenFile");
	OpenFile(XLSX_FILE, output);
	printf(output);

	//NewSheet
	//typedef int(*NewSheetFunc)(char*);
	//NewSheetFunc NewSheet;
	//NewSheet = (NewSheetFunc)GetProcAddress(module, "NewSheet");
	//int sheetIndex = NewSheet("new_sheet");
	//printf("sheetIndex %d", sheetIndex);

	//GetCellValue
	typedef int(*GetCellValueFunc)(char*, char*, char*, char*);
	GetCellValueFunc GetCellValue;
	GetCellValue = (GetCellValueFunc)GetProcAddress(module, "GetCellValue");
	char cellValue[TEXT_LEN];
	memset(cellValue, 0, TEXT_LEN);
	GetCellValue("Sheet1", "A1", cellValue, output);
	printf(cellValue);
	printf(output);

	//SetCellValue
	typedef int(*SetCellValueFunc)(char*, char*, char*, char*);
	SetCellValueFunc SetCellValue;
	SetCellValue = (SetCellValueFunc)GetProcAddress(module, "SetCellValue");
	SetCellValue("Sheet1", "A2", "Hello world.", output);
	printf(output);

	//AddPicture
	//typedef int(*AddPictureFunc)(char*, char*, char*, char*, char*);
	//AddPictureFunc AddPicture;
	//AddPicture = (AddPictureFunc)GetProcAddress(module, "AddPicture");
	//AddPicture("Sheet1", "B4", IMG_FILE, "{\"x_scale\": 0.5, \"y_scale\": 0.5}", output);
	//printf(output);

	//GetActiveSheetIndex
	typedef int(*GetActiveSheetIndexFunc)();
	GetActiveSheetIndexFunc GetActiveSheetIndex;
	GetActiveSheetIndex = (GetActiveSheetIndexFunc)GetProcAddress(module, "GetActiveSheetIndex");
	printf("index %d", GetActiveSheetIndex());

	//SetActiveSheet
	typedef void(*SetActiveSheetFunc)(int);
	SetActiveSheetFunc SetActiveSheet;
	SetActiveSheet = (SetActiveSheetFunc)GetProcAddress(module, "SetActiveSheet");
	SetActiveSheet(0);
	printf("index %d", GetActiveSheetIndex());

	//GetSheetName
	typedef void(*GetSheetNameFunc)(int, char*);
	GetSheetNameFunc GetSheetName;
	GetSheetName = (GetSheetNameFunc)GetProcAddress(module, "GetSheetName");
	char sheetName[TEXT_LEN];
	memset(sheetName, 0, TEXT_LEN);
	GetSheetName(0, sheetName);
	printf(sheetName);

	//SetSheetName
	//typedef void(*SetSheetNameFunc)(char*, char*);
	//SetSheetNameFunc SetSheetName;
	//SetSheetName = (SetSheetNameFunc)GetProcAddress(module, "SetSheetName");
	//SetSheetName("Sheet2", "SheetB");

	//GetSheetIndex
	typedef int(*GetSheetIndexFunc)(char*);
	GetSheetIndexFunc GetSheetIndex;
	GetSheetIndex = (GetSheetIndexFunc)GetProcAddress(module, "GetSheetIndex");
	printf("index %d", GetSheetIndex("Sheet3"));

	//GetSheetCount
	typedef int(*GetSheetCountFunc)();
	GetSheetCountFunc GetSheetCount;
	GetSheetCount = (GetSheetCountFunc)GetProcAddress(module, "GetSheetCount");
	printf("sheet count %d", GetSheetCount());

	//DeleteSheet
	//typedef int(*DeleteSheetFunc)(char*);
	//DeleteSheetFunc DeleteSheet;
	//DeleteSheet = (DeleteSheetFunc)GetProcAddress(module, "DeleteSheet");
	//DeleteSheet("Sheet3");

	//CopySheet
	//typedef void(*CopySheetFunc)(int, int, char*);
	//CopySheetFunc CopySheet;
	//CopySheet = (CopySheetFunc)GetProcAddress(module, "CopySheet");
	//CopySheet(0, 1, output);
	//printf(output);

	//GetRowCount
	typedef int(*GetRowCountFunc)(char*, char*);
	GetRowCountFunc GetRowCount;
	GetRowCount = (GetRowCountFunc)GetProcAddress(module, "GetRowCount");
	printf("row count %d", GetRowCount("Sheet1", output));
	printf(output);

	//GetColumnCount
	typedef int(*GetColumnCountFunc)(char*, int, char*);
	GetColumnCountFunc GetColumnCount;
	GetColumnCount = (GetColumnCountFunc)GetProcAddress(module, "GetColumnCount");
	printf("column count %d", GetColumnCount("Sheet1", 1, output));
	printf(output);

	//SetRowHeight
	typedef int(*SetRowHeightFunc)(char*, int, double, char*);
	SetRowHeightFunc SetRowHeight;
	SetRowHeight = (SetRowHeightFunc)GetProcAddress(module, "SetRowHeight");
	SetRowHeight("Sheet1", 1, 100.1, output);
	printf(output);

	//GetRowHeight
	typedef double(*GetRowHeightFunc)(char*, int, char*);
	GetRowHeightFunc GetRowHeight;
	GetRowHeight = (GetRowHeightFunc)GetProcAddress(module, "GetRowHeight");
	printf("row height %f", GetRowHeight("Sheet1", 1, output));
	printf(output);

	//RemoveRow
	typedef void(*RemoveRowFunc)(char*, int, char*);
	GetRowHeightFunc RemoveRow;
	RemoveRow = (GetRowHeightFunc)GetProcAddress(module, "RemoveRow");
	RemoveRow("Sheet1", 2, output);
	printf(output);

	//InsertRow
	typedef void(*InsertRowFunc)(char*, int, char*);
	InsertRowFunc InsertRow;
	InsertRow = (InsertRowFunc)GetProcAddress(module, "InsertRow");
	InsertRow("Sheet1", 2, output);
	printf(output);

	//DuplicateRowTo
	typedef void(*DuplicateRowToFunc)(char*, int, int, char*);
	DuplicateRowToFunc DuplicateRowTo;
	DuplicateRowTo = (DuplicateRowToFunc)GetProcAddress(module, "DuplicateRowTo");
	DuplicateRowTo("Sheet1", 1, 2, output);
	printf(output);

	//SetCellInt
	typedef void(*SetCellIntFunc)(char*, char*, int, char*);
	SetCellIntFunc SetCellInt;
	SetCellInt = (SetCellIntFunc)GetProcAddress(module, "SetCellInt");
	SetCellInt("Sheet1", "A1", 33, output);
	printf(output);

	//SetCellBool
	typedef void(*SetCellBoolFunc)(char*, char*, int, char*);
	SetCellBoolFunc SetCellBool;
	SetCellBool = (SetCellBoolFunc)GetProcAddress(module, "SetCellBool");
	SetCellBool("Sheet1", "A2", 1, output);
	printf(output);

	//SetCellFloat
	typedef void(*SetCellFloatFunc)(char*, char*, double, int, int, char*);
	SetCellFloatFunc SetCellFloat;
	SetCellFloat = (SetCellFloatFunc)GetProcAddress(module, "SetCellFloat");
	SetCellFloat("Sheet1", "A3", 0.123456, 2, 32, output);
	printf(output);

	//SetCellStr
	typedef void(*SetCellStrFunc)(char*, char*, char*, char*);
	SetCellStrFunc SetCellStr;
	SetCellStr = (SetCellStrFunc)GetProcAddress(module, "SetCellStr");
	SetCellStr("Sheet1", "B1", "bye", output);
	printf(output);

	//GetCellFormula
	typedef void(*GetCellFormulaFunc)(char*, char*, char*, char*);
	GetCellFormulaFunc GetCellFormula;
	GetCellFormula = (GetCellFormulaFunc)GetProcAddress(module, "GetCellFormula");
	char formula[TEXT_LEN];
	memset(formula, 0, TEXT_LEN);
	GetCellFormula("Sheet1", "C1", formula, output);
	printf(output);

	//test faild//GetCellHyperLink
	//typedef void(*GetCellHyperLinkFunc)(char*, char*, char*, char*);
	//GetCellHyperLinkFunc GetCellHyperLink;
	//GetCellHyperLink = (GetCellHyperLinkFunc)GetProcAddress(module, "GetCellHyperLink");
	//char link[TEXT_LEN];
	//memset(link, 0, TEXT_LEN);
	//GetCellHyperLink("Sheet1", "C2", link, output);
	//printf(output);

	//Save
	typedef void(*SaveFunc)(char*);
	SaveFunc Save;
	Save = (SaveFunc)GetProcAddress(module, "Save");
	Save(output);
	printf(output);


	//SaveAs
	//typedef void(*SaveAsFunc)(char*, char*);
	//SaveAsFunc SaveAs;
	//SaveAs = (SaveAsFunc)GetProcAddress(module, "SaveAs");
	//SaveAs("..\\..\\Release\\2.xlsx", output);
	//printf(output);

	FreeLibrary(module);
    return 0;
}

