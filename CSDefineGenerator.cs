﻿using System;
using System.IO;
using System.Data;
using System.Text;
using System.Collections.Generic;

namespace excel2json {
    /// <summary>
    /// 根据表头，生成C#类定义数据结构
    /// 表头使用三行定义：字段名称、字段类型、注释
    /// </summary>
    class CSDefineGenerator {
        struct FieldDef {
            public string name;
            public string type;
            public string comment;
        }

        string mCode;
        string sheetName;

        public string code {
            get {
                return this.mCode;
            }
        }

        public CSDefineGenerator(string excelName, DataTable sheet) {
            //-- First Row as Column Name
            if (sheet.Rows.Count < 2)
                return;
            // 首字母大写
            
            var sheetNamestrs = sheet.TableName.Split('_');
            for (int i = 0; i < sheetNamestrs.Length; i++)
            {
                sheetName += FirstLetterToUpper(sheetNamestrs[i]);
            }

            List<FieldDef> m_fieldList = new List<FieldDef>();
            DataRow typeRow = sheet.Rows[0];
            DataRow commentRow = sheet.Rows[1];

            foreach (DataColumn column in sheet.Columns) {
                FieldDef field;
                field.name = column.ToString();
                field.type = typeRow[column].ToString();
                field.comment = commentRow[column].ToString();

                m_fieldList.Add(field);
            }

            //-- 创建代码字符串
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("using System.Collections.Generic;");
            sb.AppendLine("namespace Hall");
            sb.AppendLine("{");
            sb.AppendLine("/// <summary>");
            sb.AppendFormat("/// Auto Generated Code By {0}.xlsx", excelName);
            sb.AppendLine();
            sb.AppendLine("/// </summary>");
            sb.AppendFormat("public class {0}\r\n{{", sheetName);
            sb.AppendLine();

            foreach (FieldDef field in m_fieldList) {
                var fieldType = field.type;
                if (fieldType.ToLower().Contains("list"))
                {
                    fieldType = fieldType.Replace("list", "List");
                }
                sb.AppendFormat("\tpublic {0} {1}; // {2}", fieldType, field.name, field.comment);
                sb.AppendLine();
            }

            sb.Append('}');
            sb.AppendLine();
            sb.Append('}');
            mCode = sb.ToString();
        }

        public void SaveToFile(string filePath, Encoding encoding) {
            //-- 保存文件
            using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write)) {
                using (TextWriter writer = new StreamWriter(file, encoding))
                    writer.Write(mCode);
            }
        }

        public string FirstLetterToUpper(string str)
        {
            if (str == null)
                return null;

            if (str.Length > 1)
                return char.ToUpper(str[0]) + str.Substring(1);

            return str.ToUpper();
        }
    }
}
