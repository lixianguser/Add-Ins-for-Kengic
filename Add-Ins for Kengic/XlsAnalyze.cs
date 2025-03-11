using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Kengic
{
    public static class XlsAnalyze
    {
        /// <summary>触摸屏报警文本的xlsx文件</summary>
        /// <param name="language">报警文本语言，支持中文zh-CN或英文en-US</param>
        /// <returns></returns>
        public static IWorkbook HmiAlarmText(string language)
        {
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
            IRow         row          = xssfWorkbook.CreateSheet("DiscreteAlarms").CreateRow(0);
            row.CreateCell(0).SetCellValue("ID");
            row.CreateCell(1).SetCellValue("Name");
            row.CreateCell(2).SetCellValue("Alarm text [" + language + "], Alarm text");
            row.CreateCell(3).SetCellValue("Alarm text");
            row.CreateCell(4).SetCellValue("Class");
            row.CreateCell(5).SetCellValue("Trigger tag");
            row.CreateCell(6).SetCellValue("Trigger bit");
            return xssfWorkbook;
        }
        
        /// <summary>
        /// 设置报警文本
        /// </summary>
        /// <param name="language">语言</param>
        /// <param name="workbook"></param>
        /// <param name="supervisionInfo"></param>
        /// <param name="increase">递增量</param>
        /// <param name="triggerTag"></param>
        public static void SetText(this IWorkbook workbook, 
            SupervisionInfo supervisionInfo, string language, int increase,string triggerTag)
        {
            //写入英文报警文本
            ISheet sheet = workbook.GetSheet("DiscreteAlarms");
            IRow   row   = sheet.CreateRow(increase);
            row.SetValue("ID",increase.ToString());
            row.SetValue("Name",supervisionInfo.StateStruct);
            switch (language)
            {
                case "en-US":
                    row.SetValue($"Alarm text [{language}], Alarm text", supervisionInfo.AlarmTextEn);
                    break;
                case "zh-CN":
                    row.SetValue($"Alarm text [{language}], Alarm text", supervisionInfo.AlarmTextZh);
                    break;
            }
            row.SetValue("Class",supervisionInfo.BlockTypeSupervisionNumber);
            row.SetValue("Trigger tag",triggerTag);
            string[] parts = supervisionInfo.Offset.Split('.');
            if (parts.Length > 1)
            {
                string result = parts[1];
                row.SetValue("Trigger bit",result);
            }
        }
        
        /// <summary>
        /// 触摸屏报警文本的xlsx文件设定值
        /// </summary>
        /// <param name="row">当前行</param>
        /// <param name="tagCell">目标列</param>
        /// <param name="value">设定值</param>
        private static void SetValue(this IRow row, string tagCell, string value)
        {
            //填充设定值
            switch (tagCell)
            {
                case "ID":
                    row.CreateCell(0).SetCellValue(value);
                    break;
                case "Name":
                    row.CreateCell(1).SetCellValue(value);
                    break;
                case "Alarm text [zh-CN], Alarm text":
                    row.CreateCell(2).SetCellValue(value);
                    break;
                case "Alarm text [en-US], Alarm text":
                    row.CreateCell(2).SetCellValue(value);
                    break;
                case "Alarm text":
                    row.CreateCell(3).SetCellValue(value);
                    break;
                case "Class":
                    row.CreateCell(4).SetCellValue(value);
                    break;
                case "Trigger tag":
                    row.CreateCell(5).SetCellValue(value);
                    break;
                case "Trigger bit":
                    row.CreateCell(6).SetCellValue(value);
                    break;
            }
        }
        
        /// <summary>
        /// 保存工作表
        /// </summary>
        /// <param name="workbook">工作表</param>
        /// <param name="filePath">保存文件路径</param>
        public static void Save(this IWorkbook workbook, string filePath)
        {
            // 保存文件到指定路径
            using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fileStream);
            }
        }
    }
}