using System.Collections;
using System.Collections.Generic;
using System.IO;
using UnityEditor;
using UnityEngine;

namespace Base.Excel
{
    public static class ExcelImportPath
    {
        private const string READ_KEY = "EXCEL_READ_PATH";
        private const string WRITE_KEY = "EXCEL_WRITE_PATH";

        public static void SaveReadFilePath(string path)
        {
            PlayerPrefs.SetString(READ_KEY, path);
            PlayerPrefs.Save();
        }

        public static void SaveWriteFolderPath(string path)
        {
            PlayerPrefs.SetString(WRITE_KEY, path);
            PlayerPrefs.Save();
        }

        public static string GetReadFilePath()
        {
            return PlayerPrefs.GetString(READ_KEY, string.Empty);
        }

        public static string GetWriteFolderPath()
        {
            return PlayerPrefs.GetString(WRITE_KEY, string.Empty);
        }

        public static bool IsInvalidReadFilePath()
        {
            var path = GetReadFilePath();

            if (IsInvalidReadFilePath(path))
            {
                return true;
            }

            if (!File.Exists(path))
            {
                return true;
            }

            return false;
        }

        public static bool IsInvalidReadFilePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return true;

            var extension = System.IO.Path.GetExtension(path);

            if (string.IsNullOrWhiteSpace(extension))
                return true;

            if (!(".xls" == extension || ".xlsx" == extension))
                return true;

            return false;
        }

        public static bool IsInvalidWriteFolderPath()
        {
            var path = GetWriteFolderPath();

            if (IsInvalidWriteFolderPath(path))
            {
                return true;
            }

            if (!AssetDatabase.IsValidFolder(path))
            {
                return true;
            }

            return false;
        }

        public static bool IsInvalidWriteFolderPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return true;

            var extension = System.IO.Path.GetExtension(path);

            if (!string.IsNullOrWhiteSpace(extension))
                return true;

            return false;
        }
    }
}