using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEditor;

using NPOI.XSSF;            // xlsx
using NPOI.XSSF.UserModel;
using NPOI.HSSF;            // xls
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using UnityEngine.Events;

namespace Base.Excel
{
    public static class ExcelImporter
    {
        private const string READ_MENU_PATH = "Influsion/Excel/ReadPath";
        private const string WRITE_MENU_PATH = "Influsion/Excel/WritePath";
        private const string IMPORT_MENU_PATH = "Influsion/Excel/Import";

        [MenuItem(READ_MENU_PATH, priority = 0)]
        public static bool SaveReadFilePath()
        {    
            var path = EditorUtility.OpenFilePanelWithFilters("엑셀 파일 선택", Application.dataPath, new[] { "xls", "xlsx" });

            if (!ExcelImportPath.IsInvalidReadFilePath(path))
            {
                ExcelImportPath.SaveReadFilePath(path);

                return true;
            }

            return false;
        }

        [MenuItem(WRITE_MENU_PATH, priority = 1)]
        public static bool SaveWriteFolderPath()
        {
            var path = EditorUtility.SaveFolderPanel("저장경로 선택", Application.dataPath, string.Empty);

            if (!ExcelImportPath.IsInvalidWriteFolderPath(path))
            {
                ExcelImportPath.SaveWriteFolderPath(path);

                return true;
            }

            return false;
        }

        [MenuItem(IMPORT_MENU_PATH, priority = 2, validate = true)]
        public static bool IsValidExcelFile()
        {
            Menu.SetChecked(READ_MENU_PATH, !ExcelImportPath.IsInvalidReadFilePath());
            Menu.SetChecked(WRITE_MENU_PATH, !ExcelImportPath.IsInvalidWriteFolderPath());

            return Menu.GetChecked(READ_MENU_PATH) && Menu.GetChecked(WRITE_MENU_PATH);
        }

        [MenuItem(IMPORT_MENU_PATH, priority = 2)]
        public static void Import()
        {
            var read_path = ExcelImportPath.GetReadFilePath();
            var write_path = ExcelImportPath.GetWriteFolderPath();

            ReadExcel(read_path, (work_book) =>
            {
                for (int i = 0; i < work_book.NumberOfSheets; ++i)
                {
                    var sheet = work_book.GetSheetAt(i);
                    var sheet_name = sheet.SheetName;

                    if (ExcelImportUtility.IsInvalidSheet(sheet_name))
                        continue;

                    LitJson.JsonData table = new LitJson.JsonData();

                    List<string> header_list = new List<string>();
                   
                    for (int j = sheet.FirstRowNum; j < sheet.LastRowNum + 1; ++j)
                    {
                        var row = sheet.GetRow(j);

                        LitJson.JsonData data = new LitJson.JsonData();

                        for (int k = row.FirstCellNum; k < row.LastCellNum; ++k)
                        {
                            var cell = row.GetCell(k);
                            var value = string.Empty;

                            switch (cell.CellType)
                            {
                                case CellType.Boolean: value = cell.BooleanCellValue.ToString(); break;
                                case CellType.Numeric: value = cell.NumericCellValue.ToString(); break;
                                default: value = cell.StringCellValue; break;
                            }

                            if (ExcelImportUtility.IsHeader(value))
                            {
                                header_list.Add(ExcelImportUtility.GetHeader(value));
                            }
                            else if (k < header_list.Count)
                            {
                                data[header_list[k]] = value;
                            }
                        }

                        table.Add(data);
                    }

                    var save_path = $"{write_path}/{sheet_name}.json";
                    
                    File.WriteAllText(save_path, table.ToJson(), System.Text.Encoding.UTF8);
                }

                AssetDatabase.Refresh();
            });
        }

        private static void ReadExcel(string path, UnityAction<IWorkbook> onRead = null)
        {
            try
            {
                using (FileStream stream = File.Open(path, FileMode.Open, FileAccess.ReadWrite))
                {
                    IWorkbook workbook = null;

                    if (Path.GetExtension(path) == ".xls")
                    {
                        workbook = new HSSFWorkbook(stream);
                    }
                    else
                    {
                        workbook = new XSSFWorkbook(stream);
                    }

                    if (null != workbook && null != onRead)
                    {
                        onRead(workbook);
                    }
                }
            }
            catch (Exception e)
            {
                Debug.Log($"Read Excel Exception = {e.Message}");
            }
        }
    }
}