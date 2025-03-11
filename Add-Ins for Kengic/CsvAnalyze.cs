using System;
using System.Collections.Generic;
using System.IO;

namespace Kengic
{
    public static class CsvAnalyze
    {
        /// <summary>
        /// 根据元素名称获取值
        /// </summary>
        /// <param name="values"></param>
        /// <param name="attributeName"></param>
        /// <returns></returns>
        private static string GetAttribute(this string[] values ,string attributeName)
        {
            switch (attributeName)
            {
                case "Identification":
                    if (values.Length >= 3)
                    {
                        // 获取第一列数据并保留后4位
                        string firstColumn = values[0];
                
                        if (firstColumn.Length > 4)
                        {
                            firstColumn = firstColumn.Substring(firstColumn.Length - 4);
                        }

                        // 将16进制数转化为10进制数
                        int decimalValue = Convert.ToInt32(firstColumn, 16);
                        return decimalValue.ToString();
                    }
                    break;
                case "Alarm text":
                    if (values.Length >= 3)
                    {
                        // 获取第三列数据
                        string thirdColumn = values[2];
                        return thirdColumn;
                    }
                    break;
            }

            return null;
        }
        
        /// <summary>
        /// 解析导出的ProDiag.csv文件
        /// </summary>
        /// <param name="streamReader">csv文件流</param>
        /// <param name="fileName">csv文件名称</param>
        /// <returns>获取Identification和Alarm text</returns>
        public static List<ProDiagInfo> Analyze(this StreamReader streamReader, string fileName)
        {
            var data = new List<ProDiagInfo>();
            
            //获取ProDiag文件的语言
            string language = null;
            
            // 找到最后一个下划线的位置
            int lastUnderscoreIndex = fileName.LastIndexOf('_');

            // 找到 .csv 之前的位置
            int dotCsvIndex = fileName.LastIndexOf(".csv", StringComparison.Ordinal);

            // 提取最后一个下划线和 .csv 之间的字符串
            if (lastUnderscoreIndex != -1 && dotCsvIndex != -1 && lastUnderscoreIndex < dotCsvIndex)
            {
                language = fileName.Substring(lastUnderscoreIndex + 1, dotCsvIndex - lastUnderscoreIndex - 1);
            }

            // 读取并忽略表头
            streamReader.ReadLine();
            
            while (!streamReader.EndOfStream)
            {
                var    line           = streamReader.ReadLine();
                var    values         = line?.Split(';');
                string identification = values.GetAttribute("Identification");
                string alarmText      = values.GetAttribute("Alarm text");
                data.Add(new ProDiagInfo{ Identification = identification, AlarmText = alarmText, Language = language});
            }
            return data;
        }
    }
    
    /// <summary>
    /// 诊断信息
    /// </summary>
    public class ProDiagInfo
    {
        /// <summary>
        /// 报警文本编号
        /// </summary>
        public string Identification { get; set; }
        /// <summary>
        /// 报警文本
        /// </summary>
        public string AlarmText     { get; set; }
        
        /// <summary>
        /// 语言
        /// </summary>
        public string Language      { get; set; }
    }
}