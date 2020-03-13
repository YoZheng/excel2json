using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excel2json
{
    public static class ConvertTool
    {
        public static string FirstLetterToUpper(string str)
        {
            if (str == null)
                return null;

            var convertName = "";
            var sheetNamestrs = str.Split('_');
            for (int i = 0; i < sheetNamestrs.Length; i++)
            {
                if (str.Length > 1)
                {
                    convertName += char.ToUpper(sheetNamestrs[i][0]) + sheetNamestrs[i].Substring(1);
                }
                else
                {
                    convertName += sheetNamestrs[i].ToUpper();
                }
                    
            }

            return convertName;
        }
    }
}
