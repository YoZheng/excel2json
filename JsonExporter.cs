using System;
using System.IO;
using System.Data;
using System.Text;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace excel2json {
    /// <summary>
    /// 将DataTable对象，转换成JSON string，并保存到文件中
    /// </summary>
    class JsonExporter {
        string mContext = "";
        private int head;

        public string context {
            get {
                return mContext;
            }
        }

        /// <summary>
        /// 构造函数：完成内部数据创建
        /// </summary>
        /// <param name="excel">ExcelLoader Object</param>
        public JsonExporter(ExcelLoader excel, bool lowcase, bool exportArray, string dateFormat, int header) {
            head = header;
            List<DataTable> validSheets = new List<DataTable>();
            List<DataRow> sheetsTypeRow = new List<DataRow>();
            for (int i = 0; i < excel.Sheets.Count; i++) {
                DataTable sheet = excel.Sheets[i];

                if (sheet.Columns.Count > 0 && sheet.Rows.Count > header)
                    validSheets.Add(sheet);
                if(sheet.Rows.Count > 0)
                    sheetsTypeRow.Add(sheet.Rows[0]);
            }

            var jsonSettings = new JsonSerializerSettings {
                DateFormatString = dateFormat,
                Formatting = Formatting.Indented
            };

            if (validSheets.Count == 1) {   // single sheet

                //-- convert to object
                object sheetValue = convertSheet(validSheets[0], exportArray, lowcase, sheetsTypeRow[0]);

                //-- convert to json string
                mContext = JsonConvert.SerializeObject(sheetValue, jsonSettings);
            }
            else { // mutiple sheet

                Dictionary<string, object> data = new Dictionary<string, object>();
                for (int i = 0; i < validSheets.Count; i++)
                {
                    object sheetValue = convertSheet(validSheets[i], exportArray, lowcase, sheetsTypeRow[i]);
                    data.Add(validSheets[i].TableName, sheetValue);
                }

                //foreach (var sheet in validSheets)
                //{
                //    object sheetValue = convertSheet(sheet, exportArray, lowcase);
                //    data.Add(sheet.TableName, sheetValue);
                //}

                //-- convert to json string
                mContext = JsonConvert.SerializeObject(data, jsonSettings);
            }
        }

        private object convertSheet(DataTable sheet, bool exportArray, bool lowcase, DataRow sheetsTypeRow) {
            if (exportArray)
                return convertSheetToArray(sheet, lowcase, sheetsTypeRow);
            else
                return convertSheetToDict(sheet, lowcase);
        }

        private object convertSheetToArray(DataTable sheet, bool lowcase, DataRow sheetsTypeRow) {
            List<object> values = new List<object>();

            int firstDataRow = head;
            for (int i = firstDataRow; i < sheet.Rows.Count; i++) {
                DataRow row = sheet.Rows[i];

                values.Add(
                    convertRowToDict(sheet, row, lowcase, firstDataRow, sheetsTypeRow)
                    );
            }

            return values;
        }

        /// <summary>
        /// 以第一列为ID，转换成ID->Object的字典对象
        /// </summary>
        private object convertSheetToDict(DataTable sheet, bool lowcase) {
            Dictionary<string, object> importData =
                new Dictionary<string, object>();

            int firstDataRow = head;
            for (int i = firstDataRow; i < sheet.Rows.Count; i++) {
                DataRow row = sheet.Rows[i];
                string ID = row[sheet.Columns[0]].ToString();
                if (ID.Length <= 0)
                    ID = string.Format("row_{0}", i);

                var rowObject = convertRowToDict(sheet, row, lowcase, firstDataRow);
                rowObject[ID] = ID;
                importData[ID] = rowObject;
            }

            return importData;
        }

        /// <summary>
        /// 把一行数据转换成一个对象，每一列是一个属性
        /// </summary>
        private Dictionary<string, object> convertRowToDict(DataTable sheet, DataRow row, bool lowcase, int firstDataRow, DataRow sheetsTypeRow = null) {
            var rowData = new Dictionary<string, object>();
            int col = 0;
            foreach (DataColumn column in sheet.Columns) {
                object value = row[column];
                //Console.WriteLine("convertRowToDict1: " + sheetsTypeRow[column].ToString());
                if (sheetsTypeRow != null && sheetsTypeRow[column].ToString().ToLower().Contains("list"))
                {
                    var listValue = "[" + value.ToString().Replace(';', ',') + "]";
                    var json = JsonConvert.DeserializeObject(listValue, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
                    value = json;
                    //Console.WriteLine("convertRowToDict1 DeserializeObject: " + json.ToString());
                }


                if (value.GetType() == typeof(System.DBNull)) {
                    value = getColumnDefault(sheet, column, firstDataRow);
                }
                else if (value.GetType() == typeof(double)) { // 去掉数值字段的“.0”
                    double num = (double)value;
                    if ((int)num == num)
                        value = (int)num;
                }

                string fieldName = column.ToString();
                // 表头自动转换成小写
                if (lowcase)
                    fieldName = fieldName.ToLower();

                if (string.IsNullOrEmpty(fieldName))
                    fieldName = string.Format("col_{0}", col);

                rowData[fieldName] = value;
                col++;
            }

            return rowData;
        }

        /// <summary>
        /// 对于表格中的空值，找到一列中的非空值，并构造一个同类型的默认值
        /// </summary>
        private object getColumnDefault(DataTable sheet, DataColumn column, int firstDataRow) {
            for (int i = firstDataRow; i < sheet.Rows.Count; i++) {
                object value = sheet.Rows[i][column];
                Type valueType = value.GetType();
                if (valueType != typeof(System.DBNull)) {
                    if (valueType.IsValueType)
                        return Activator.CreateInstance(valueType);
                    break;
                }
            }
            return "";
        }

        /// <summary>
        /// 将内部数据转换成Json文本，并保存至文件
        /// </summary>
        /// <param name="jsonPath">输出文件路径</param>
        public void SaveToFile(string filePath, Encoding encoding) {
            //-- 保存文件
            using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write)) {
                using (TextWriter writer = new StreamWriter(file, encoding))
                    writer.Write(mContext);
            }
        }
    }
}
