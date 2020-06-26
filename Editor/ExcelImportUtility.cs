using System.Collections;
using System.Collections.Generic;
using System.IO;
using UnityEngine;

namespace Base.Excel
{
    public static class ExcelImportUtility
    {
        public static bool IsHeader(string value)
        {
            if (string.IsNullOrEmpty(value))
                return false;

            if (value.Length < 3)
                return false;

            return value[0] == '[' && value[value.Length - 1] == ']';
        }

        public static string GetHeader(string value)
        {
            value = value.Remove(0, 1);
            value = value.Remove(value.Length - 1, 1);

            return value;
        }

        public static bool IsInvalidSheet(string sheet_name)
        {
            return false;
        }

        public static string ToRelativePath(string path)
        {
            var index = path.IndexOf("Assets");

            return path.Remove(0, index);
        }

        public static string ToAbsoultePath(string path)
        {
            path = Path.GetFullPath(path);

            return path.Replace("\\", "/");
        }
    }
}
