using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Siemens.Engineering;
using Siemens.Engineering.Hmi;
using Siemens.Engineering.Hmi.Communication;
using Siemens.Engineering.Hmi.Cycle;
using Siemens.Engineering.Hmi.Globalization;
using Siemens.Engineering.Hmi.RuntimeScripting;
using Siemens.Engineering.Hmi.Screen;
using Siemens.Engineering.Hmi.Tag;
using Siemens.Engineering.Hmi.TextGraphicList;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;
using Siemens.Engineering.SW.ExternalSources;
using Siemens.Engineering.SW.Tags;
using Siemens.Engineering.SW.Types;
using Siemens.Engineering.SW.WatchAndForceTables;
using Screen = Siemens.Engineering.Hmi.Screen.Screen;
using MessageBox = Siemens.Engineering.AddIn.MessageBox;
using Microsoft.Win32;

namespace Kengic
{
    public static class Command
    {
        /// <summary>
        /// 展示异常详细信息🔎
        /// </summary>
        /// <param name="messageBox">博途Messagebox API</param>
        /// <param name="exception">系统抛出的异常信息</param>
        public static void ShowException(this MessageBox messageBox, EngineeringException exception)
        {
            messageBox.ShowNotification(NotificationIcon.Error, "异常",
                exception.MessageData.Text, exception.Message);
        }

        /// <summary>
        /// 展示异常详细信息🔎
        /// </summary>
        /// <param name="messageBox">博途Messagebox API</param>
        /// <param name="exception">系统抛出的异常信息</param>
        public static void ShowException(this MessageBox messageBox, Exception exception)
        {
            messageBox.ShowNotification(NotificationIcon.Error, "异常",
                exception.Message, exception.StackTrace);
        }

        /// <summary>
        /// 获取V20版本是否Update3及以上版本程序块可以导出成SIMATIC SD文件
        /// </summary>
        /// <returns></returns>
        public static bool CanExportAsDocuments()
        {
            string keyPath       = @"HKEY_LOCAL_MACHINE\SOFTWARE\Siemens\Automation\_InstalledSW\TIAP20\TIA_Opns";
            string versionString = (string)Registry.GetValue(keyPath, "VersionString", null);

            if (versionString != null)
            {
                // 正则提取最后一段版本号
                Match match = Regex.Match(versionString, @"V\d+\.\d+\.\d+\.(\d+)");
                if (match.Success && int.TryParse(match.Groups[1].Value, out int buildNumber))
                {
                    if (buildNumber >= 3)
                    {
                        return true;
                    }
                }
                else
                {
                    return false;
                }
            }

            return false;
        }

        /// <summary>
        /// 导出程序为xml文件📃
        /// </summary>
        /// <param name="exportItem">导出项</param>
        /// <param name="exportPath">导出路径</param>
        /// <returns>导出xml文件的路径</returns>
        public static void ExportInfo(this IEngineeringObject exportItem, string exportPath)
        {
            //定义导出选项
            const ExportOptions exportOption = ExportOptions.WithDefaults | ExportOptions.WithReadOnly;

            switch (exportItem)
            {
                //导出程序块
                case PlcBlock item:
                {
                    if (item.ProgrammingLanguage == ProgrammingLanguage.ProDiag_OB)
                        return;

                    if (item.ProgrammingLanguage == ProgrammingLanguage.ProDiag)
                    {
                        DirectoryInfo directoryInfo =
                            new DirectoryInfo(Path.GetDirectoryName(exportPath) ?? string.Empty);
                        ((FB)item).ExportProDIAGInfo(directoryInfo);
                        return;
                    }

                    if (item.IsConsistent) //编译的结果
                    {
                        //删除已存在的文件
                        if (File.Exists(exportPath))
                        {
                            File.Delete(exportPath);
                        }

                        //导出程序块
                        item.Export(new FileInfo(exportPath), exportOption);
                    }
                    else
                    {
                        throw new ArgumentException("目标未编译");
                    }

                    break;
                }
                case PlcTagTable _:
                case PlcType _:
                case ScreenOverview _:
                case ScreenGlobalElements _:
                case Screen _:
                case TagTable _:
                case Connection _:
                case GraphicList _:
                case TextList _:
                case Cycle _:
                case MultiLingualGraphic _:
                case ScreenTemplate _:
                case VBScript _:
                case ScreenPopup _:
                case ScreenSlidein _:
                case PlcWatchTable _:
                {
                    //删除已存在的文件
                    if (File.Exists(exportPath))
                    {
                        File.Delete(exportPath);
                    }

                    //导出程序块
                    var parameter = new Dictionary<Type, object>
                    {
                        { typeof(FileInfo), new FileInfo(exportPath) },
                        { typeof(ExportOptions), exportOption }
                    };
                    exportItem.Invoke("Export", parameter);
                    break;
                }
                case PlcExternalSource _:

                    break;
            }

            Console.WriteLine($"导出完成：{exportItem}");
        }

        /// <summary>
        /// 递归的方式导出指定instanceOfName的FB块
        /// </summary>
        public static void ExportInfo(this PlcBlockGroup blockGroup, string instanceOfName, string filePath)
        {
            foreach (PlcBlock plcBlock in blockGroup.Blocks)
            {
                if (plcBlock.Name == instanceOfName)
                {
                    ExportInfo(plcBlock, filePath);
                }
            }

            foreach (PlcBlockUserGroup subBlockGroup in blockGroup.Groups)
            {
                subBlockGroup.ExportInfo(instanceOfName, filePath);
            }
        }

        /// <summary>
        /// 导出带诊断的实例DB和ProDiagFB
        /// </summary>
        /// <param name="blockGroup">程序块</param>
        /// <param name="exportPath">导出文件夹</param>
        /// <param name="recursion">使用递归方法</param>
        public static void ExportInfo(this PlcBlockGroup blockGroup, string exportPath, bool recursion)
        {
            //使用递归方法
            if (!recursion)
            {
                return;
            }

            //导出有Supervisions的InstanceDB
            foreach (PlcBlock plcBlock in blockGroup.Blocks)
            {
                if (plcBlock.ProgrammingLanguage == ProgrammingLanguage.DB)
                {
                    foreach (EngineeringAttributeInfo info in plcBlock.GetAttributeInfos())
                    {
                        //获取Supervisions
                        if (info.Name == "Supervisions")
                        {
                            //导出InstanceDB
                            if (plcBlock.GetAttribute("Supervisions") != null)
                            {
                                string filePath = $@"{exportPath}\{plcBlock.Name}.xml";
                                plcBlock.ExportInfo(filePath);
                            }
                        }
                    }
                }

                //获取ProDiagFB
                if (plcBlock.ProgrammingLanguage == ProgrammingLanguage.ProDiag)
                {
                    string[] filePaths =
                    {
                        $@"{exportPath}\{plcBlock.Name}_en-US.csv",
                        $@"{exportPath}\{plcBlock.Name}_zh-CN.csv",
                    };

                    // //删除文件
                    // foreach (var file in filePaths)
                    // {
                    //     if (File.Exists(file))
                    //     {
                    //         File.Delete(file);
                    //     }
                    // }


                    //导出ProDiagFB为.csv文件
                    plcBlock.ExportInfo(filePaths[0]);
                }
            }

            //递归查询所有
            foreach (PlcBlockUserGroup subBlockGroup in blockGroup.Groups)
            {
                subBlockGroup.ExportInfo(exportPath, true);
            }
        }

        /// <summary>
        /// 将程序块导出为文档
        /// </summary>
        /// <param name="plcBlock">程序块</param>
        /// <param name="exportPath">导出路径</param>
        public static void ExportInfoDoc(this PlcBlock plcBlock,string exportPath)
        {
            if (plcBlock.ProgrammingLanguage == ProgrammingLanguage.ProDiag_OB)
                return;

            if (plcBlock.ProgrammingLanguage == ProgrammingLanguage.ProDiag)
                return;
            
            //获取文件路径
            string directory = Path.GetDirectoryName(exportPath);
            if (directory == null)
                return;
            
            //获取文件无后缀名称
            string fileName = Path.GetFileNameWithoutExtension(exportPath);
            
            //删除已存在的文件
            string s7dcl = Path.Combine(directory, $"{fileName}.s7dcl");
            if (File.Exists(s7dcl))
            {
                File.Delete(s7dcl);
            }
            
            //删除已存在的文件
            string s7res = Path.Combine(directory, $"{fileName}.s7res");
            if (File.Exists(s7res))
            {
                File.Delete(s7res);
            }

            if (plcBlock.IsConsistent) //编译的结果
            {
                plcBlock.ExportAsDocuments(new DirectoryInfo(directory), fileName);
            }
        }
        
        /// <summary>
        /// 将 UDT 导出为文档
        /// </summary>
        /// <param name="plcType">用户数据类型</param>
        /// <param name="exportPath">导出路径</param>
        public static void ExportInfoDoc(this PlcType plcType,string exportPath)
        {
            //获取文件路径
            string directory = Path.GetDirectoryName(exportPath);
            if (directory == null)
                return;
            
            //获取文件无后缀名称
            string fileName = Path.GetFileNameWithoutExtension(exportPath);
            
            //删除已存在的文件
            string s7dcl = Path.Combine(directory, $"{fileName}.s7dcl");
            if (File.Exists(s7dcl))
            {
                File.Delete(s7dcl);
            }
            
            //删除已存在的文件
            string s7res = Path.Combine(directory, $"{fileName}.s7res");
            if (File.Exists(s7res))
            {
                File.Delete(s7res);
            }

            if (plcType.IsConsistent) //编译的结果
            {
                plcType.ExportAsDocuments(new DirectoryInfo(directory), fileName);
            }
        }

        /// <summary>
        /// 递归遍历所有 PlcBlockGroup 并返回 List&lt;Block&gt;
        /// </summary>
        /// <param name="blockGroup"></param>
        /// <returns></returns>
        public static List<ProDiagFB> GetProDiagFB(PlcBlockGroup blockGroup)
        {
            var result = new List<ProDiagFB>();

            if (blockGroup == null) return result;

            // 查找当前组内的所有 ProDiag 语言的 PlcBlock
            var proDiagBlocks = blockGroup.Blocks
                .Where(block => block.ProgrammingLanguage == ProgrammingLanguage.ProDiag);

            foreach (PlcBlock block in proDiagBlocks)
            {
                result.Add(new ProDiagFB { Name = block.Name });
            }

            // 递归遍历所有子组
            foreach (PlcBlockUserGroup subGroup in blockGroup.Groups)
            {
                result.AddRange(GetProDiagFB(subGroup));
            }

            return result;
        }

        /// <summary>
        /// 导入xml文件📃
        /// </summary>
        /// <param name="destination">程序块导入目标位置</param>
        /// <param name="filePath">导入xml文件的路径</param>
        public static void ImportInfo(this IEngineeringCompositionOrObject destination, string filePath)
        {
            FileInfo            fileInfo     = new FileInfo(filePath);
            const ImportOptions importOption = ImportOptions.Override;
            filePath = fileInfo.FullName;

            switch (destination)
            {
                case CycleComposition _:
                case ConnectionComposition _:
                case MultiLingualGraphicComposition _:
                case GraphicListComposition _:
                case TextListComposition _:
                {
                    var parameter = new Dictionary<Type, object>
                    {
                        { typeof(string), filePath },
                        { typeof(ImportOptions), importOption }
                    };
                    //导入xml文件
                    ((IEngineeringComposition)destination).Invoke("Import", parameter);
                    break;
                }
                case PlcBlockGroup group when Path.GetExtension(filePath).Equals(".xml"):
                    group.Blocks.Import(fileInfo, importOption);
                    break;
                case PlcBlockGroup group:
                {
                    IEngineeringObject currentDestination = group;
                    while (!(currentDestination is PlcSoftware))
                    {
                        currentDestination = currentDestination.Parent;
                    }

                    PlcExternalSourceComposition col = (currentDestination as PlcSoftware).ExternalSourceGroup
                        .ExternalSources;

                    string sourceName = Path.GetRandomFileName();
                    sourceName = Path.ChangeExtension(sourceName, ".src");
                    PlcExternalSource src = col.CreateFromFile(sourceName, filePath);
                    src.GenerateBlocksFromSource();
                    src.Delete();
                    break;
                }
                case PlcTagTableGroup group:
                    group.TagTables.Import(fileInfo, importOption);
                    break;
                case PlcTypeGroup group:
                    group.Types.Import(fileInfo, importOption);
                    break;
                case PlcExternalSourceGroup group:
                {
                    PlcExternalSource temp = group.ExternalSources.Find(Path.GetFileName(filePath));
                    temp?.Delete();
                    group.ExternalSources.CreateFromFile(Path.GetFileName(filePath), filePath);
                    break;
                }
                case TagFolder folder:
                    folder.TagTables.Import(fileInfo, importOption);
                    break;
                case ScreenFolder folder:
                    folder.Screens.Import(fileInfo, importOption);
                    break;
                case ScreenTemplateFolder folder:
                    folder.ScreenTemplates.Import(fileInfo, importOption);
                    break;
                case ScreenPopupFolder folder:
                    folder.ScreenPopups.Import(fileInfo, importOption);
                    break;
                case ScreenSlideinSystemFolder folder:
                    folder.ScreenSlideins.Import(fileInfo, importOption);
                    break;
                case VBScriptFolder folder:
                    folder.VBScripts.Import(fileInfo, importOption);
                    break;
                case ScreenGlobalElements _:
                    (destination.Parent as HmiTarget)?.ImportScreenGlobalElements(fileInfo, importOption);
                    break;
                case ScreenOverview _:
                    (destination.Parent as HmiTarget)?.ImportScreenOverview(fileInfo, importOption);
                    break;
                case PlcWatchAndForceTableGroup folder:
                    folder.WatchTables.Import(fileInfo, importOption);
                    break;
            }

            Console.WriteLine($"导入完成：{filePath}");
        }

        /// <summary>
        /// 获取项目实例
        /// </summary>
        /// <param name="tiaPortal"></param>
        /// <returns>ProjectBase</returns>
        public static ProjectBase GetProjectBase(this TiaPortal tiaPortal)
        {
            ProjectBase projectBase;

            if (tiaPortal.LocalSessions.Any())
            {
                //多用户本地会话
                projectBase = tiaPortal.LocalSessions
                    .FirstOrDefault(s => s.Project != null && s.Project.IsPrimary)?.Project;
            }
            else
            {
                //本地项目
                projectBase = tiaPortal.Projects.FirstOrDefault(p => p.IsPrimary);
            }

            return projectBase;
        }

        /// <summary>
        /// 通过项目树选中项获取PlcSoftware🌲
        /// </summary>
        /// <param name="menuSelectionProvider"></param>
        /// <returns></returns>
        public static PlcSoftware GetPlcSoftware(this IEnumerable<object> menuSelectionProvider)
        {
            //获取PlcSoftware
            PlcSoftware plcSoftware = null;

            foreach (IEngineeringObject engineeringObject in menuSelectionProvider)
            {
                IEngineeringObject parent = engineeringObject.Parent;
                while (!(parent is PlcSoftware))
                {
                    parent = parent.Parent;
                }

                plcSoftware = parent as PlcSoftware;
            }

            return plcSoftware;
        }

        /// <summary>
        /// 查询触摸屏目标
        /// </summary>
        /// <param name="device">触摸屏设备</param>
        /// <returns></returns>
        public static HmiTarget GetHmiTarget(this Device device)
        {
            DeviceItemComposition deviceItemComposition = device.DeviceItems;
            foreach (DeviceItem deviceItem in deviceItemComposition)
            {
                SoftwareContainer softwareContainer = deviceItem.GetService<SoftwareContainer>();
                if (softwareContainer != null)
                {
                    Software  softwareBase = softwareContainer.Software;
                    HmiTarget hmiTarget    = softwareBase as HmiTarget;
                    return hmiTarget;
                }
            }

            return null;
        }

        /// <summary>
        /// 通过设备项查询PLC目标
        /// </summary>
        /// <param name="deviceItem">PLC设备项</param>
        /// <returns></returns>
        public static PlcSoftware GetPlcSoftware(this DeviceItem deviceItem)
        {
            SoftwareContainer softwareContainer = deviceItem.GetService<SoftwareContainer>();
            if (softwareContainer != null)
            {
                Software    softwareBase = softwareContainer.Software;
                PlcSoftware plcSoftware  = softwareBase as PlcSoftware;
                return plcSoftware;
            }

            return null;
        }

        /// <summary>
        /// 获取项目中所有设备
        /// </summary>
        /// <param name="project">项目</param>
        /// <returns></returns>
        private static IEnumerable<Device> AllDevices(this ProjectBase project)
        {
            foreach (Device device in project.Devices)
                yield return device;
            foreach (var group in project.DeviceGroups)
            {
                foreach (var device in GetDevicesFromGroupRecursive(group))
                {
                    yield return device;
                }
            }
            foreach (Device device in project.UngroupedDevicesGroup.Devices)
                yield return device;
        }
        
        /// <summary>
        /// 遍历用户组中获取设备
        /// </summary>
        /// <param name="group"></param>
        /// <returns></returns>
        private static IEnumerable<Device> GetDevicesFromGroupRecursive(DeviceUserGroup group)
        {
            foreach (var device in group.Devices)
                yield return device;

            foreach (var subGroup in group.Groups)
            {
                foreach (var device in GetDevicesFromGroupRecursive(subGroup))
                    yield return device;
            }
        }

        /// <summary>
        /// 获取所有触摸屏设备信息
        /// </summary>
        /// <param name="projectBase">项目</param>
        /// <returns>设备名称Name、设备Device类</returns>
        public static List<DeviceInfo> GetDeviceInfos(this ProjectBase projectBase)
        {
            //获取所有设备，并把名称和Device写入到数据
            var devices = new List<DeviceInfo>();
            foreach (Device device in projectBase.AllDevices())
            {
                if (device.DeviceItems[0].GetAttribute("TypeIdentifier").ToString().Contains(":6AV2"))
                {
                    //DeviceInfo deviceInfo = new DeviceInfo { Name = device.Name, Device = device };
                    DeviceInfo deviceInfo = new DeviceInfo { Name = device.Name };
                    devices.Add(deviceInfo);
                }
            }

            return devices;
        }

        /// <summary>
        /// 根据设备名称查找设备
        /// </summary>
        /// <param name="projectBase">项目</param>
        /// <param name="deviceName">设备名称</param>
        /// <returns>设备Device</returns>
        public static Device FindDeviceByName(this ProjectBase projectBase, string deviceName)
        {
            //获取所有设备，通过设备名称查找设备，并返回设备。
            foreach (Device device in projectBase.AllDevices())
            {
                if (device.DeviceItems[0].GetAttribute("Name").ToString().Contains(deviceName))
                {
                    return device;
                }
            }

            return null;
        }


        //------

        /// <summary>
        /// 删除文件夹及其内容
        /// </summary>
        /// <param name="targetDir">目标路径</param>
        public static void DeleteDirectoryAndContents(this string targetDir)
        {
            if (!Directory.Exists(targetDir))
            {
                throw new DirectoryNotFoundException($"目录 {targetDir} 不存在。");
            }

            string[] files = Directory.GetFiles(targetDir);
            string[] dirs  = Directory.GetDirectories(targetDir);

            // 先删除所有文件
            foreach (string file in files)
            {
                File.Delete(file);
            }

            // 然后递归删除所有子目录
            foreach (string dir in dirs)
            {
                DeleteDirectoryAndContents(dir);
            }

            // 最后，删除目录本身
            Directory.Delete(targetDir, false);
        }

        /// <summary>
        /// 字符串处理将"/"替换为"_"
        /// </summary>
        /// <param name="name">程序块名称</param>
        /// <returns></returns>
        public static string Replace(this string name)
        {
            //查询名称是否包含“/”，如果包含替换更“_”
            while (name.Contains("/"))
            {
                name = name.Replace("/", "_");
            }

            return name;
        }
    }
}