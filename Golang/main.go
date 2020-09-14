package main
//doc: https://xuri.me/excelize/

/*
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "common.h"
*/
import "C"
//befor import "C" cannot be empty line
import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
)

//export NewFile
func NewFile() {
	xlFile = excelize.NewFile()
}

//export OpenFile
func OpenFile(sheetName, ret_error *C.char){
	xlFile, err = excelize.OpenFile(C.GoString(sheetName))//cann't use :=
	if err != nil {
		C.strcpy(ret_error, C.CString(err.Error()))
	}
}

//export NewSheet
func NewSheet(sheetName *C.char) C.int{
	if xlFile == nil {
		return C.int(-1)
	}
	index := xlFile.NewSheet(C.GoString(sheetName))
	return C.int(index)
}

//export GetCellValue
func GetCellValue(sheet, axis, value, ret_error *C.char) {
	if xlFile == nil {
		C.strcpy(ret_error, C.CString("No Excel Object"))
		return
	}

	str, err := xlFile.GetCellValue(C.GoString(sheet), C.GoString(axis))
	if err != nil {
		C.strcpy(ret_error, C.CString(err.Error()))
	}else{
		C.strcpy(value, C.CString(str))
	}
}

//export SetCellValue
func SetCellValue(sheet, axis, value, ret_error *C.char) {
	if xlFile == nil {
		C.strcpy(ret_error, C.CString("No Excel Object"))
		return
	}

	err := xlFile.SetCellValue(C.GoString(sheet), C.GoString(axis), C.GoString(value))
	if err != nil {
		C.strcpy(ret_error, C.CString(err.Error()))
	}
}

//export AddPicture
func AddPicture(sheet, cell, picture, format, ret_error *C.char) {
	if xlFile == nil {
		C.strcpy(ret_error, C.CString("No Excel Object"))
		return
	}

	err := xlFile.AddPicture(C.GoString(sheet), C.GoString(cell), C.GoString(picture), C.GoString(format))
	if err != nil {
		C.strcpy(ret_error, C.CString(err.Error()))
	}
}

//export GetActiveSheetIndex
func GetActiveSheetIndex()(index C.int) {
	if xlFile == nil {
		return
	}

	return C.int(xlFile.GetActiveSheetIndex())
}

//export SetActiveSheet
func SetActiveSheet(index C.int) {
	if xlFile == nil {
		return
	}

	xlFile.SetActiveSheet(int(index))
}

//export GetSheetName
func GetSheetName(index C.int, name *C.char) {
	if xlFile == nil {
		return
	}

	C.strcpy(name, C.CString(xlFile.GetSheetName(int(index))))
}

//export SetSheetName
func SetSheetName(oldName, newName *C.char) {
	if xlFile == nil {
		return
	}

	xlFile.SetSheetName(C.GoString(oldName), C.GoString(newName))
}

//export GetSheetIndex
func GetSheetIndex(name *C.char) C.int {
	if xlFile == nil {
		return -1
	}

	return C.int(xlFile.GetSheetIndex(C.GoString(name)))
}

//export GetSheetCount
func GetSheetCount() C.int {
	if xlFile == nil {
		return -1
	}
	sheet_list := xlFile.GetSheetList()
	return C.int(len(sheet_list))
}

//export DeleteSheet
func DeleteSheet(name *C.char) {
	if xlFile == nil {
		return
	}

	xlFile.DeleteSheet(C.GoString(name))
}

//export CopySheet
func CopySheet(from, to C.int, ret_error *C.char) {
	if xlFile == nil {
		return
	}

	if err := xlFile.CopySheet(int(from), int(to)); err != nil {
		fmt.Println(err.Error())
		C.strcpy(ret_error, C.CString(err.Error()))
	}
}

//export GetRowCount
func GetRowCount(sheet, ret_error *C.char) C.int {
	if xlFile == nil {
		C.strcpy(ret_error, C.CString("No Excel Object"))
		return -1
	}
	rows, err := xlFile.GetRows(C.GoString(sheet))
	if err != nil {
		fmt.Println(err.Error())
		C.strcpy(ret_error, C.CString(err.Error()))
		return -2
	}else{
		return C.int(len(rows))
	}
}

//export GetColumnCount
func GetColumnCount(sheet *C.char, row_index C.int, ret_error *C.char) C.int {
	if xlFile == nil {
		C.strcpy(ret_error, C.CString("No Excel Object"))
		return -1
	}
	rows, err := xlFile.GetRows(C.GoString(sheet))
	if err != nil {
		fmt.Println(err.Error())
		C.strcpy(ret_error, C.CString(err.Error()))
		return -2
	}else{
		if int(row_index) + 1 > len(rows) {
			C.strcpy(ret_error, C.CString("over rows index"))
			return -3
		}else{
			return C.int(len(rows[row_index]))
		}
	}
}

//export SetRowHeight
func SetRowHeight(sheet *C.char, row_index C.int, height C.double, ret_error *C.char) {
	if xlFile == nil {
		C.strcpy(ret_error, C.CString("No Excel Object"))
		return
	}
	err := xlFile.SetRowHeight(C.GoString(sheet), int(row_index), float64(height))
	if err != nil {
		fmt.Println(err.Error())
		C.strcpy(ret_error, C.CString(err.Error()))
		return
	}
}

//export GetRowHeight
func GetRowHeight(sheet *C.char, row C.int, ret_error *C.char) C.double {
	if xlFile == nil {
		C.strcpy(ret_error, C.CString("No Excel Object"))
		return 0.0
	}
	f, err := xlFile.GetRowHeight(C.GoString(sheet), int(row))
	if err != nil {
		fmt.Println(err.Error())
		C.strcpy(ret_error, C.CString(err.Error()))
		return 0.0
	}else{
		C.print_d(C.double(f))
		return C.double(f)
	}
}

//export RemoveRow
func RemoveRow(sheet *C.char, row C.int, ret_error *C.char) {
	if xlFile == nil {
		C.strcpy(ret_error, C.CString("No Excel Object"))
	}
	err := xlFile.RemoveRow(C.GoString(sheet), int(row))
	if err != nil {
		fmt.Println(err.Error())
		C.strcpy(ret_error, C.CString(err.Error()))
	}
}

//export InsertRow
func InsertRow(sheet *C.char, row C.int, ret_error *C.char) {
	if xlFile == nil {
		C.strcpy(ret_error, C.CString("No Excel Object"))
	}
	err := xlFile.InsertRow(C.GoString(sheet), int(row))
	if err != nil {
		fmt.Println(err.Error())
		C.strcpy(ret_error, C.CString(err.Error()))
	}
}

//export DuplicateRowTo
func DuplicateRowTo(sheet *C.char, row, row2 C.int, ret_error *C.char) {
	if xlFile == nil {
		C.strcpy(ret_error, C.CString("No Excel Object"))
	}
	err := xlFile.DuplicateRowTo(C.GoString(sheet), int(row), int(row2))
	if err != nil {
		fmt.Println(err.Error())
		C.strcpy(ret_error, C.CString(err.Error()))
	}
}

//export SetCellInt
func SetCellInt(sheet, axis *C.char, value C.int, ret_error *C.char) {
	if xlFile == nil {
		C.strcpy(ret_error, C.CString("No Excel Object"))
	}
	err := xlFile.SetCellInt(C.GoString(sheet), C.GoString(axis), int(value))
	if err != nil {
		fmt.Println(err.Error())
		C.strcpy(ret_error, C.CString(err.Error()))
	}
}

//export SetCellBool
func SetCellBool(sheet, axis *C.char, value C.int, ret_error *C.char) {
	if xlFile == nil {
		C.strcpy(ret_error, C.CString("No Excel Object"))
	}
	var b bool
	if value == 0 {
		b = false
	}else{
		b = true
	}
	err := xlFile.SetCellBool(C.GoString(sheet), C.GoString(axis), b)
	if err != nil {
		fmt.Println(err.Error())
		C.strcpy(ret_error, C.CString(err.Error()))
	}
}

//export SetCellFloat
func SetCellFloat(sheet, axis *C.char, value C.double, prec, bitSize C.int, ret_error *C.char) {
	if xlFile == nil {
		C.strcpy(ret_error, C.CString("No Excel Object"))
	}

	err := xlFile.SetCellFloat(C.GoString(sheet), C.GoString(axis), float64(value), int(prec), int(bitSize))
	if err != nil {
		fmt.Println(err.Error())
		C.strcpy(ret_error, C.CString(err.Error()))
	}
}

//export SetCellStr
func SetCellStr(sheet, axis, value, ret_error *C.char) {
	if xlFile == nil {
		C.strcpy(ret_error, C.CString("No Excel Object"))
	}

	err := xlFile.SetCellStr(C.GoString(sheet), C.GoString(axis), C.GoString(value))
	if err != nil {
		fmt.Println(err.Error())
		C.strcpy(ret_error, C.CString(err.Error()))
	}
}

//export SetCellDefault
func SetCellDefault(sheet, axis, value, ret_error *C.char) {
	if xlFile == nil {
		C.strcpy(ret_error, C.CString("No Excel Object"))
	}

	err := xlFile.SetCellDefault(C.GoString(sheet), C.GoString(axis), C.GoString(value))
	if err != nil {
		fmt.Println(err.Error())
		C.strcpy(ret_error, C.CString(err.Error()))
	}
}

//export GetCellFormula
func GetCellFormula(sheet, axis, formula *C.char, ret_error *C.char) {
	if xlFile == nil {
		C.strcpy(ret_error, C.CString("No Excel Object"))
	}

	str, err := xlFile.GetCellFormula(C.GoString(sheet), C.GoString(axis))
	if err != nil {
		fmt.Println(err.Error())
		C.strcpy(ret_error, C.CString(err.Error()))
	}else{
		C.strcpy(formula, C.CString(str))
	}
}

//test faild//export GetCellHyperLink
//func GetCellHyperLink(sheet, axis, link, ret_error *C.char) {
//	if xlFile == nil {
//		C.strcpy(ret_error, C.CString("No Excel Object"))
//	}
//
//	_, str, err := xlFile.GetCellHyperLink(C.GoString(sheet), C.GoString(axis))
//	fmt.Println("link", str)
//	if err != nil {
//		fmt.Println(err.Error())
//		C.strcpy(ret_error, C.CString(err.Error()))
//	}else{
//		fmt.Println(str)
//		C.strcpy(link, C.CString(str))
//	}
//}

//export Save
func Save(ret_error *C.char) {
	if xlFile == nil {
		C.strcpy(ret_error, C.CString("No Excel Object"))
		return
	}

	if err := xlFile.Save(); err != nil {
		fmt.Println(err.Error())
		C.strcpy(ret_error, C.CString(err.Error()))
	}
}

//export SaveAs
func SaveAs(fileName, ret_error *C.char) {
	if xlFile == nil {
		C.strcpy(ret_error, C.CString("No Excel Object"))
		return
	}

	if err := xlFile.SaveAs(C.GoString(fileName)); err != nil {
		fmt.Println(err.Error())
		C.strcpy(ret_error, C.CString(err.Error()))
	}
}

var xlFile *excelize.File
var err error
func main() {
	// Need a main function to make CGO compile package as C shared library
}
