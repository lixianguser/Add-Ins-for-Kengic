using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using NPOI.SS.UserModel;
using Siemens.Engineering;
using Siemens.Engineering.AddIn;
using Siemens.Engineering.AddIn.Menu;
using Siemens.Engineering.Hmi;
using Siemens.Engineering.Hmi.Tag;
using Siemens.Engineering.HW;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;
using MessageBox = Siemens.Engineering.AddIn.MessageBox;

namespace Kengic
{
    public class AddIn : ContextMenuAddIn
    {
        private readonly TiaPortal   _tiaPortal;

        public AddIn(TiaPortal tiaPortal) : base("Add-Ins")
        {
            _tiaPortal  = tiaPortal;
        }
        
        protected override void BuildContextMenuItems(ContextMenuAddInRoot addInRoot)
        {
            //导出
            Submenu export = addInRoot.Items.AddSubmenu("导出");
            export.Items.AddActionItem<IEngineeringObject>("Xml数据", Export_OnClick);
            export.Items.AddActionItem<PlcBlock>("SIMATIC SD文件", Export_OnClick);
            export.Items.AddActionItem<GlobalDB>("WCS接口", Export_OnClick);
            export.Items.AddActionItem<InstanceDB>("SCADA接口", Export_OnClick);
            
            //导入
            Submenu import = addInRoot.Items.AddSubmenu("导入");
            import.Items.AddActionItem<IEngineeringObject>("Xml数据", Import_OnClick);
            import.Items.AddActionItem<PlcBlockGroup>("SIMATIC SD文件", Import_OnClick);
            
            //扩展工具
            Submenu tools = addInRoot.Items.AddSubmenu("工具");
            tools.Items.AddActionItem<PlcBlock>("自动编号", Number_OnClick);
            tools.Items.AddActionItem<DeviceItem>("创建触摸屏报警", Alarm_OnClick);
            Submenu proDiag = tools.Items.AddSubmenu("指定ProDiagFB");
            //proDiag.Items.AddActionItem<InstanceDB>("指定ProDiag FB1", proDiag_OnClick);
            //proDiag.Items.AddActionItem<InstanceDB>("指定ProDiag FB2", proDiag_OnClick);
            proDiag.Items.AddActionItem<InstanceDB>("指定ProDiagFB",
                (menuSelectionProvider) => ProDiag_OnClick(menuSelectionProvider, "1"));
            proDiag.Items.AddActionItem<InstanceDB>("删除指定的ProDiag",
                menuSelectionProvider => ProDiag_OnClick(menuSelectionProvider, ""));

        }

        /// <summary>
        /// 导出程序块
        /// </summary>
        /// <param name="menuSelectionProvider"></param>
        /// <returns>.xml</returns>
        private void Export_OnClick(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
#if DEBUG
            System.Diagnostics.Debugger.Launch();
#endif
            //反馈API
            FeedbackContext feedback = _tiaPortal.GetFeedbackContext();
            feedback.Log(NotificationIcon.Information,"<导出Xml数据>");
            
            //MessageBox API
            MessageBox messageBox = _tiaPortal.GetMessageBox();
            
            try
            {
                //创建独占窗口
                using (ExclusiveAccess exclusiveAccess = _tiaPortal.ExclusiveAccess("导出Xml数据"))
                {
                    //选择.xml存放的位置
                    FolderBrowserDialog dialog = new FolderBrowserDialog();
                    dialog.Description = "请选择导出文件的保存路径";

                    if (dialog.ShowDialog(new Form() 
                            { TopMost = true, WindowState = FormWindowState.Maximized }) == DialogResult.OK)
                    {
                        feedback.Log(NotificationIcon.Information, $"打开窗口获取.xml文件的保存路径:{dialog.SelectedPath}");

                        string files = "";
                    
                        foreach (IEngineeringObject iEngineeringObject in menuSelectionProvider.GetSelection())
                        {
                            //监听独占窗口取消按钮
                            if (exclusiveAccess.IsCancellationRequested)
                            {
                                feedback.Log(NotificationIcon.Error, $"监听独占窗口取消");
                                break;
                            }
                        
                            //查询名称是否包含“/”，如果包含替换更“_”
                            string name = iEngineeringObject.GetAttribute("Name").ToString().Replace();

                            //获取程序块类型
                            string type = iEngineeringObject.GetType().ToString().Split('.').Last();
                            
                            //获取程序块版本号
                            string version = iEngineeringObject.GetAttribute("HeaderVersion").ToString();
            
                            //创建文件路径
                            string path     = dialog.SelectedPath;
                            string filePath = Path.Combine(path, $"{type}+{name}+v{version}.xml");
            
                            //导出程序块
                            exclusiveAccess.Text = $"正在导出： [{type}]{name}v{version}";
                            iEngineeringObject.ExportInfo(filePath);
                            feedback.Log(NotificationIcon.Information, $"已导出:{filePath}");
                        
                            //设置导出的文件路径
                            files += filePath + "\r\n";
                        }
                    
                        //导出完成
                        messageBox.ShowNotification(NotificationIcon.Success, "导出完成", "目标文件:", files);
                        feedback.Log(NotificationIcon.Success,"<导出Xml数据>完成");
                    }
                }
            }
            catch (EngineeringException ex)
            {
                messageBox.ShowException(ex);
                feedback.Log(NotificationIcon.Error,ex.Message);
                throw;
            }
            catch (Exception ex)
            {
                messageBox.ShowException(ex);
                feedback.Log(NotificationIcon.Error,ex.Message);
                throw;
            }
        }
        
        /// <summary>
        /// 导出程序块为SIMATIC SD文件
        /// </summary>
        /// <param name="menuSelectionProvider"></param>
        /// <returns>.xml</returns>
        private void Export_OnClick(MenuSelectionProvider<PlcBlock> menuSelectionProvider)
        {
#if DEBUG
            System.Diagnostics.Debugger.Launch();
#endif
            //反馈API
            FeedbackContext feedback = _tiaPortal.GetFeedbackContext();
            feedback.Log(NotificationIcon.Information,"<将程序块导出为文档>");
            
            //MessageBox API
            MessageBox messageBox = _tiaPortal.GetMessageBox();
            
            try
            {
                if (!Command.CanExportAsDocuments())
                {
                    feedback.Log(NotificationIcon.Error,"博途版本不支持，仅支持V20 Update3及以上版本");
                    return;
                }
                
                //创建独占窗口
                using (ExclusiveAccess exclusiveAccess = _tiaPortal.ExclusiveAccess("将程序块导出为文档"))
                {
                    //选择.xml存放的位置
                    FolderBrowserDialog dialog = new FolderBrowserDialog();
                    dialog.Description = "请选择导出文件的保存路径";

                    if (dialog.ShowDialog(new Form() 
                            { TopMost = true, WindowState = FormWindowState.Maximized }) == DialogResult.OK)
                    {
                        feedback.Log(NotificationIcon.Information, $"打开窗口获取文档文件的保存路径:{dialog.SelectedPath}");

                        string files = "";
                    
                        foreach (PlcBlock plcBlock in menuSelectionProvider.GetSelection())
                        {
                            //监听独占窗口取消按钮
                            if (exclusiveAccess.IsCancellationRequested)
                            {
                                feedback.Log(NotificationIcon.Error, $"监听独占窗口取消");
                                break;
                            }
                        
                            //查询名称是否包含“/”，如果包含替换更“_”
                            string name = plcBlock.GetAttribute("Name").ToString().Replace();

                            //获取程序块类型
                            string type = plcBlock.GetType().ToString().Split('.').Last();
                            
                            //获取程序块版本号
                            string version = plcBlock.GetAttribute("HeaderVersion").ToString();
            
                            //创建文件路径
                            string path     = dialog.SelectedPath;
                            string filePath = Path.Combine(path, $"{type}+{name}+v{version}.s7dcl");
            
                            //导出程序块
                            exclusiveAccess.Text = $"正在导出： [{type}]{name}v{version}";
                            plcBlock.ExportInfo(filePath);
                            
                            feedback.Log(NotificationIcon.Information, $"已导出");
                            
                            //设置导出的文件路径
                            files += filePath + "\r\n";
                        }
                    
                        //导出完成
                        messageBox.ShowNotification(NotificationIcon.Success, "导出完成", "目标文件:", files);
                        feedback.Log(NotificationIcon.Success,"<将程序块导出为文档>完成");
                    }
                }
            }
            catch (EngineeringException ex)
            {
                messageBox.ShowException(ex);
                feedback.Log(NotificationIcon.Error,ex.Message);
                throw;
            }
            catch (Exception ex)
            {
                messageBox.ShowException(ex);
                feedback.Log(NotificationIcon.Error,ex.Message);
                throw;
            }
        }
        
        /// <summary>
        /// 导出WCS接口
        /// </summary>
        /// <param name="menuSelectionProvider"></param>
        /// <returns>.csv</returns>
        private void Export_OnClick(MenuSelectionProvider<GlobalDB> menuSelectionProvider)
        {
#if DEBUG
            System.Diagnostics.Debugger.Launch();
#endif
            //反馈API
            FeedbackContext feedback  = _tiaPortal.GetFeedbackContext();
            feedback.Log(NotificationIcon.Information,"<导出WCS接口>");
            
            //MessageBox API
            MessageBox messageBox = _tiaPortal.GetMessageBox();
            
            string directory = null;//导出.xml的文件路径

            try
            {
                // 打开窗口获取.csv文件保存位置
                SaveFileDialog dialog = new SaveFileDialog
                {
                    Filter = "逗号分隔值|*.csv",
                    Title  = "请选择保存文件的路径"
                };
                
                if (dialog.ShowDialog(new Form()
                        { TopMost = true, WindowState = FormWindowState.Maximized }) == DialogResult.OK)
                {
                    //获取.csv文件路径
                    string path = dialog.FileName;//.csv文件保存路径
                    if (path == null)
                    {
                        throw new FileNotFoundException("文件保存路径未知错误");
                    }
                    feedback.Log(NotificationIcon.Information, $"打开窗口获取.csv文件的保存路径:{path}");
                    
                    //创建导出DB.xml的文件夹路径
                    string dir = Path.GetDirectoryName(path);//获取.csv文件所在文件夹路径
                    if (dir == null)
                    {
                        throw new FileNotFoundException("导出文件路径未知错误");
                    }
                    directory = Path.Combine(dir, ".temp"); //设置导出DB.xml的文件夹路径
                    if (Directory.Exists(directory) == false)
                    {
                        Directory.CreateDirectory(directory);
                    }
                    feedback.Log(NotificationIcon.Information,$"创建导出DB.xml的文件夹路径:{directory}");
                    
                    //创建独占窗口
                    using (ExclusiveAccess exclusiveAccess = _tiaPortal.ExclusiveAccess("导出WCS接口"))
                    {
                        //创建StreamWriter
                        StreamWriter streamWriter = new StreamWriter(path);
                        //创建标题行
                        streamWriter.WriteLine("\"输送机编号\",\"变量名\",\"类型\",\"DB\",\"开始地址\",\"DB名称\"");
                        
                        foreach (GlobalDB globalDb in menuSelectionProvider.GetSelection())
                        {
                            //监听独占窗口取消按钮
                            if (exclusiveAccess.IsCancellationRequested)
                            {
                                feedback.Log(NotificationIcon.Error, $"监听独占窗口取消");
                                break;
                            }
                            
                            //获取DB的信息
                            string name                = globalDb.Name;// 获取 GlobalDB 名称
                            string number              = globalDb.Number.ToString();//获取编号
                            string programmingLanguage = globalDb.ProgrammingLanguage.ToString();//获取程序语言
                            
                            if (Directory.Exists(directory))
                            {
                                //创建文件路径
                                string filePath = Path.Combine(directory, $"{name.Replace()}.xml");
            
                                //导出程序块
                                globalDb.ExportInfo(filePath);
                                feedback.Log(NotificationIcon.Information,$"已导出：{filePath}");
                                
                                //写入StreamWriter
                                WCS wcs = new WCS
                                {
                                    FilePath       = filePath,
                                    ProgrammingLanguage = programmingLanguage,
                                    Number              = number,
                                    InstanceName        = name,
                                    StreamWriter        = streamWriter
                                };
                                
                                wcs.Run();
                                
                                feedback.Log(NotificationIcon.Information,$"已处理：{name}[{programmingLanguage}{number}]");
                            }
                        }
                        
                        streamWriter.Close();
                    }
                    
                    //导出完成
                    messageBox.ShowNotification(NotificationIcon.Success,"导出完成",$"目标文件:{path}");
                    feedback.Log(NotificationIcon.Success,"<导出WCS接口>完成");
                }
            }
            catch (EngineeringException ex)
            {
                messageBox.ShowException(ex);
                feedback.Log(NotificationIcon.Error,ex.Message);
                throw;
            }
            catch (Exception ex)
            {
                messageBox.ShowException(ex);
                feedback.Log(NotificationIcon.Error,ex.Message);
                throw;
            }
            finally
            {
                directory.DeleteDirectoryAndContents();//删除导出的.xml文件的文件夹
                feedback.Log(NotificationIcon.Information,$"已删除：{directory}文件夹");
            }
        }
        
        /// <summary>
        /// 导出SCADA接口
        /// </summary>
        /// <param name="menuSelectionProvider"></param>
        /// <returns>.csv</returns>
        private void Export_OnClick(MenuSelectionProvider<InstanceDB> menuSelectionProvider)
        {
#if DEBUG
            System.Diagnostics.Debugger.Launch();
#endif
            //反馈API
            FeedbackContext feedback = _tiaPortal.GetFeedbackContext();
            feedback.Log(NotificationIcon.Information,"<导出SCADA接口>");
            
            //MessageBox API
            MessageBox messageBox = _tiaPortal.GetMessageBox();
            
            string directory = null;//导出.xml的文件路径

            try
            {
                // 打开窗口获取.csv文件保存位置
                SaveFileDialog dialog = new SaveFileDialog
                {
                    Filter = "逗号分隔值|*.csv",
                    Title  = "请选择保存文件的路径"
                };

                if (dialog.ShowDialog(new Form()
                        { TopMost = true, WindowState = FormWindowState.Maximized }) == DialogResult.OK)
                {
                    //获取.csv文件路径
                    string path = dialog.FileName; //.csv文件保存路径
                    if (path == null)
                    {
                        throw new FileNotFoundException("文件保存路径未知错误");
                    }

                    feedback.Log(NotificationIcon.Information, $"打开窗口获取.csv文件的保存路径:{path}");
                    
                    //创建导出DB.xml的文件夹路径
                    string dir = Path.GetDirectoryName(path);//获取.csv文件所在文件夹路径
                    if (dir == null)
                    {
                        throw new FileNotFoundException("导出文件路径未知错误");
                    }
                    directory = Path.Combine(dir, ".temp"); //设置导出程序块的.xml文件夹路径
                    if (Directory.Exists(directory) == false)
                    {
                        Directory.CreateDirectory(directory);
                    }
                    feedback.Log(NotificationIcon.Information,$"创建导出程序块的.xml文件夹路径:{directory}");
                    
                    // 独占窗口
                    using (ExclusiveAccess exclusiveAccess = _tiaPortal.ExclusiveAccess("导出中……"))
                    {
                        //创建StreamWriter
                        StreamWriter streamWriter = new StreamWriter(path);
                        //创建标题行
                        streamWriter.WriteLine("\"地址\",\"数据类别0状态1错误2警告\",\"报警文本en-US\",\"报警文本zh-CN\"");

                        foreach (InstanceDB instanceDb in menuSelectionProvider.GetSelection())
                        {
                            //监听独占窗口取消按钮
                            if (exclusiveAccess.IsCancellationRequested)
                            {
                                feedback.Log(NotificationIcon.Error, $"监听独占窗口取消");
                                break;
                            }
                            
                            //获取DB的信息
                            string name                = instanceDb.Name;// 获取 InstanceDB 名称
                            string instanceOfName      = instanceDb.InstanceOfName;//获取 InstanceDB 实例的名称
                            string number              = instanceDb.Number.ToString();//获取编号
                            string programmingLanguage = instanceDb.ProgrammingLanguage.ToString();//获取程序语言

                            if (Directory.Exists(directory))
                            {
                                exclusiveAccess.Text = $"导出中:{name}";
                                //创建InstanceDB文件路径
                                string iDBFilePath = Path.Combine(directory, $"{name.Replace()}.xml");
            
                                //导出DB程序块
                                instanceDb.ExportInfo(iDBFilePath);
                                feedback.Log(NotificationIcon.Information,$"已导出：{iDBFilePath}");
                                
                                
                                //创建FB文件路径
                                string fBFilePath = Path.Combine(directory, $"{instanceOfName.Replace()}.xml");

                                //判断文件夹是否已包含FB块的xml文件
                                string[] fileNames = Directory.GetFiles(directory, Path.GetFileName(fBFilePath));
                                if (fileNames.Length < 1)
                                {
                                    exclusiveAccess.Text = $"导出中:{instanceOfName}";
                                    //导出FB程序块
                                    PlcSoftware plcSoftware = menuSelectionProvider.GetSelection().GetPlcSoftware();
                                    PlcBlockGroup plcBlockGroups = plcSoftware.BlockGroup;
                                    plcBlockGroups.ExportInfo(instanceOfName,fBFilePath);
                                    feedback.Log(NotificationIcon.Information,$"已导出：{fBFilePath}");
                                }
                                
                                //写入StreamWriter
                                SCADA scada = new SCADA
                                {
                                    FBXmlFilePath       = fBFilePath,
                                    IDbXmlFilePath      = iDBFilePath,
                                    InstanceName        = name,
                                    Number              = number,
                                    ProgrammingLanguage = programmingLanguage,
                                    InstanceOfName      = instanceOfName,
                                    StreamWriter        = streamWriter
                                };

                                scada.Run();
                                
                                feedback.Log(NotificationIcon.Information,$"已处理：{name}[{programmingLanguage}{number}]");
                            }
                        }
                        
                        streamWriter.Close();
                    }
                    
                    //导出完成
                    messageBox.ShowNotification(NotificationIcon.Success,"导出完成",$"目标文件:{path}");
                    feedback.Log(NotificationIcon.Success,"<导出SCADA接口>完成");
                }
            }
            catch (EngineeringException ex)
            {
                messageBox.ShowException(ex);
                throw;
            }
            catch (Exception ex)
            {
                messageBox.ShowException(ex);
                feedback.Log(NotificationIcon.Error,ex.Message);
                throw;
            }
            finally
            {
                directory.DeleteDirectoryAndContents();//删除导出的.xml文件的文件夹
                feedback.Log(NotificationIcon.Information,$"已删除：{directory}文件夹");
            }
        }
        
        /// <summary>
        /// 程序块导入
        /// </summary>
        /// <param name="menuSelectionProvider"></param>
        private void Import_OnClick(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
#if DEBUG
            System.Diagnostics.Debugger.Launch();
#endif
            //反馈API
            FeedbackContext feedback  = _tiaPortal.GetFeedbackContext();
            feedback.Log(NotificationIcon.Information,"<导入Xml数据>");
            
            //MessageBox API
            MessageBox messageBox = _tiaPortal.GetMessageBox();
            
            try
            { 
                //创建独占窗口
                using (ExclusiveAccess exclusiveAccess = _tiaPortal.ExclusiveAccess("导入Xml文件"))
                {
                    //获取项目实例
                    ProjectBase projectBase = _tiaPortal.GetProjectBase();
                    
                    //如果项目实例为空，抛出异常
                    if (projectBase == null)
                    {
                        throw new EngineeringObjectDisposedException("无法获取项目实例");
                    }

                    //选择.xml文件
                    OpenFileDialog dialog = new OpenFileDialog
                    {
                        Multiselect = true,
                        Filter      = "xml File(*.xml)| *.xml",
                        Title = "请选择要导入的文件(可多选)"
                    };

                    //打开文件选择窗体
                    if (dialog.ShowDialog(new Form()
                            { TopMost = true, WindowState = FormWindowState.Maximized }) == DialogResult.OK
                        && !string.IsNullOrEmpty(dialog.FileName))
                    {
                        //创建撤回撤销事务
                        using (Transaction transaction = exclusiveAccess.Transaction(projectBase, "导入Xml数据"))
                        {
                            foreach (IEngineeringCompositionOrObject iEngineeringCompositionOrObject in
                                     menuSelectionProvider.GetSelection())
                            {
                                //监听独占窗口取消按钮
                                if (exclusiveAccess.IsCancellationRequested)
                                {
                                    feedback.Log(NotificationIcon.Error, $"监听独占窗口取消");
                                    return;
                                }

                                //轮询批量导入xml文件
                                foreach (string fileName in dialog.FileNames)
                                {
                                    exclusiveAccess.Text = $"正在导入： {fileName}";
                                    iEngineeringCompositionOrObject.ImportInfo(fileName);
                                    feedback.Log(NotificationIcon.Information, $"已导入:{fileName}");
                                }
                            }
                            
                            //创建回滚
                            if (transaction.CanCommit)
                            {
                                transaction.CommitOnDispose();
                                feedback.Log(NotificationIcon.Information, "已创建回滚");
                            }
                        }
                    }
                    
                    //导入完成
                    messageBox.ShowNotification(NotificationIcon.Success, "完成","导入Xml数据完成");
                    feedback.Log(NotificationIcon.Success, $"<导入Xml数据>完成");
                }
                
                
            }
            catch (EngineeringException ex)
            {
                messageBox.ShowException(ex);
                feedback.Log(NotificationIcon.Error,ex.Message);
                throw;
            }
            catch (Exception ex)
            {
                messageBox.ShowException(ex);
                feedback.Log(NotificationIcon.Error,ex.Message);
                throw;
            }
        }
        
        /// <summary>
        /// 程序块导入
        /// </summary>
        /// <param name="menuSelectionProvider"></param>
        private void Import_OnClick(MenuSelectionProvider<PlcBlockGroup> menuSelectionProvider)
        {
#if DEBUG
            System.Diagnostics.Debugger.Launch();
#endif
            //反馈API
            FeedbackContext feedback  = _tiaPortal.GetFeedbackContext();
            feedback.Log(NotificationIcon.Information,"<从文档导入程序块>");
            
            //MessageBox API
            MessageBox messageBox = _tiaPortal.GetMessageBox();
            
            try
            { 
                if (!Command.CanExportAsDocuments())
                {
                    feedback.Log(NotificationIcon.Error,"博途版本不支持，仅支持V20 Update3及以上版本");
                    return;
                }
                
                //创建独占窗口
                using (ExclusiveAccess exclusiveAccess = _tiaPortal.ExclusiveAccess("导入文件"))
                {
                    //获取项目实例
                    ProjectBase projectBase = _tiaPortal.GetProjectBase();
                    
                    //如果项目实例为空，抛出异常
                    if (projectBase == null)
                    {
                        throw new EngineeringObjectDisposedException("无法获取项目实例");
                    }

                    //选择.xml文件
                    OpenFileDialog dialog = new OpenFileDialog
                    {
                        Multiselect = true,
                        Filter      = "s7dcl File(*.s7dcl)| *.s7dcl",
                        Title = "请选择要导入的文件(可多选)"
                    };

                    //打开文件选择窗体
                    if (dialog.ShowDialog(new Form()
                            { TopMost = true, WindowState = FormWindowState.Maximized }) == DialogResult.OK
                        && !string.IsNullOrEmpty(dialog.FileName))
                    {
                        //创建撤回撤销事务
                        using (Transaction transaction = exclusiveAccess.Transaction(projectBase, "从文档导入程序块"))
                        {
                            foreach (PlcBlockGroup plcBlockGroup in
                                     menuSelectionProvider.GetSelection())
                            {
                                //监听独占窗口取消按钮
                                if (exclusiveAccess.IsCancellationRequested)
                                {
                                    feedback.Log(NotificationIcon.Error, $"监听独占窗口取消");
                                    return;
                                }

                                //轮询批量导入xml文件
                                foreach (string fileName in dialog.FileNames)
                                {
                                    exclusiveAccess.Text = $"正在导入： {fileName}";
                                    //iEngineeringCompositionOrObject.ImportInfo(fileName);
                                    //获取文件路径
                                    string directory = Path.GetDirectoryName(fileName);
                                    if (directory == null)
                                        return;
                                    //获取文件无后缀名称
                                    string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);

                                    plcBlockGroup.Blocks.ImportFromDocuments(new DirectoryInfo(directory), fileNameWithoutExtension,
                                        ImportDocumentOptions.Override);
                                    
                                    feedback.Log(NotificationIcon.Information, $"已导入:{fileName}");
                                }
                            }
                            
                            //创建回滚
                            if (transaction.CanCommit)
                            {
                                transaction.CommitOnDispose();
                                feedback.Log(NotificationIcon.Information, "已创建回滚");
                            }
                        }
                    }
                    
                    //导入完成
                    messageBox.ShowNotification(NotificationIcon.Success, "完成","从文档导入程序块完成");
                    feedback.Log(NotificationIcon.Success, $"从文档导入程序块完成");
                }
                
                
            }
            catch (EngineeringException ex)
            {
                messageBox.ShowException(ex);
                feedback.Log(NotificationIcon.Error,ex.Message);
                throw;
            }
            catch (Exception ex)
            {
                messageBox.ShowException(ex);
                feedback.Log(NotificationIcon.Error,ex.Message);
                throw;
            }
        }
        
        /// <summary>
        /// 自动编号
        /// </summary>
        /// <param name="menuSelectionProvider"></param>
        private void Number_OnClick(MenuSelectionProvider<PlcBlock> menuSelectionProvider)
        {
#if DEBUG
            System.Diagnostics.Debugger.Launch();
#endif
            //反馈API
            FeedbackContext feedback  = _tiaPortal.GetFeedbackContext();
            feedback.Log(NotificationIcon.Information,"<自动编号>");
            
            //MessageBox API
            MessageBox messageBox = _tiaPortal.GetMessageBox();

            try
            { 
                //创建独占窗口
                using (ExclusiveAccess exclusiveAccess = _tiaPortal.ExclusiveAccess("自动编号"))
                {
                    //打开自动编号窗口
                    NumberForm numberForm = new NumberForm();
                    if (numberForm.ShowDialog() != DialogResult.OK)
                        return;
                    //获取输入信息
                    int startingNumber = numberForm.StartingNumber;
                    int increment      = numberForm.Increment;
                    feedback.Log(NotificationIcon.Information, 
                        $"获取输入信息:起始编号{startingNumber},递增值{increment}");

                    ProjectBase projectBase = _tiaPortal.GetProjectBase();
                    
                    //创建撤回撤销事务
                    using (Transaction transaction = exclusiveAccess.Transaction(projectBase, 
                               "自动编号"))
                    {
                        foreach (PlcBlock plcBlock in menuSelectionProvider.GetSelection())
                        {
                            //监听独占窗口取消按钮
                            if (exclusiveAccess.IsCancellationRequested)
                            {
                                feedback.Log(NotificationIcon.Error, $"监听独占窗口取消");
                                return; 
                            }
                            exclusiveAccess.Text = "正在编号:" + plcBlock.Name + "设置为 " + startingNumber;
                            
                            //设置为手动编号
                            if (plcBlock.AutoNumber)
                            {
                                plcBlock.AutoNumber = false;
                            }
                            
                            //设置编号
                            if (plcBlock.Number != startingNumber)
                            {
                                plcBlock.Number = startingNumber;
                            }

                            feedback.Log(NotificationIcon.Information,
                                $"已完成：{plcBlock.Name}[{plcBlock.ProgrammingLanguage}{plcBlock.Number}]");
                            startingNumber += increment;
                        }

                        if (transaction.CanCommit)
                        {
                            transaction.CommitOnDispose();
                            feedback.Log(NotificationIcon.Information, "已创建回滚");
                        }
                    }
                    
                    feedback.Log(NotificationIcon.Success,"<自动编号>完成");
                }
            }
            catch (EngineeringException ex)
            {
                messageBox.ShowException(ex);
                feedback.Log(NotificationIcon.Error,ex.Message);
                throw;
            }
            catch (Exception ex)
            {
                messageBox.ShowException(ex);
                feedback.Log(NotificationIcon.Error,ex.Message);
                throw;
            }
        }
        
        /// <summary>
        /// 创建报警文本
        /// </summary>
        /// <param name="menuSelectionProvider"></param>
        private void Alarm_OnClick(MenuSelectionProvider<DeviceItem> menuSelectionProvider)
        {
#if DEBUG
            Debugger.Launch();
#endif
            
            //反馈API
            FeedbackContext feedback  = _tiaPortal.GetFeedbackContext();
            feedback.Log(NotificationIcon.Information,"<创建触摸屏报警>");
            
            //MessageBox API
            MessageBox messageBox = _tiaPortal.GetMessageBox();
            
            string directory = null;//导出.xml的文件路径

            try
            {
                //提示🔔
                const string warning = "请确认以下数据:\r\n"
                                       + "1:PLC中包含ProDiagFB块，并确认相关实例DB已指定ProDiag。\r\n"
                                       + "2:PLC已完全编译，并且无错误。\r\n"
                                       + "3:检查'PLC监控和报警'中'监控定义'，所有的定义均包含对应中英文报警文本。\r\n"
                                       + "按下'确认'键继续使用工具。";
                ConfirmationResult warningResult = messageBox.ShowConfirmation(ConfirmationIcon.General, "提醒", warning,
                    ConfirmationChoices.Yes | ConfirmationChoices.No, ConfirmationResult.Yes);

                if (warningResult != ConfirmationResult.Yes)
                {
                    feedback.Log(NotificationIcon.Error, "提示窗口返回No");
                    return;
                }

                using (ExclusiveAccess exclusiveAccess = _tiaPortal.ExclusiveAccess("创建触摸屏报警"))
                {
                    //获取项目
                    ProjectBase projectBase   = _tiaPortal.GetProjectBase();
                    
                    using (Transaction transaction =
                           exclusiveAccess.Transaction(projectBase, "导入\"自动生成报警变量表\""))
                    {
                        //获取所有设备，并把名称和Device写入到数据
                        var deviceInfos = projectBase.GetDeviceInfos();

                        //新建触摸屏选择窗体
                        AlarmForm mainForm = new AlarmForm();
                        mainForm.SetDevices(deviceInfos); //填充数据

                        if (mainForm.ShowDialog() != DialogResult.OK)
                        {
                            return;
                        }
                        
                        //选择保存文件的位置
                        FolderBrowserDialog dialog = new FolderBrowserDialog();
                        dialog.Description = "请选择导出文件的保存路径";

                        if (dialog.ShowDialog(new Form { TopMost = true, WindowState = FormWindowState.Maximized }) == DialogResult.OK)
                        {
                            //获取保存文件的位置
                            string path     = dialog.SelectedPath;
                            directory = Path.Combine(path, ".temp");
                            string englishPath = $@"{path}\HMIAlarms-en_US.xlsx";
                            string chinesePath = $@"{path}\HMIAlarms-zh_CN.xlsx";
                            
                            feedback.Log(NotificationIcon.Information, $"获取保存文件的位置:{path}");

                            foreach (DeviceItem deviceItem in menuSelectionProvider.GetSelection())
                            {
                                //监听独占窗口取消按钮
                                if (exclusiveAccess.IsCancellationRequested)
                                {
                                    feedback.Log(NotificationIcon.Error, "监听独占窗口取消");
                                    return;
                                }

                                #region 导出InstanceDB和ProDiag

                                //获取PLC目标
                                PlcSoftware plcSoftware = deviceItem.GetPlcSoftware();
                            
                                //获取程序块组
                                PlcBlockGroup plcBlockGroup = plcSoftware.BlockGroup;
                                
                                //导出InstanceDB和ProDiag
                                plcBlockGroup.ExportInfo(directory,true);
                                feedback.Log(NotificationIcon.Information, $"已导出InstanceDB和ProDiag保存到{directory}");

                                #endregion

                                #region 获取报警文本
                                
                                //定义报警文本列表
                                var proDiagInfos = new List<ProDiagInfo>();
                                
                                //查询文件夹路径是否存在
                                if (Directory.Exists(directory))
                                {
                                    //获取文件夹中所有的.csv文件
                                    string[] files         = Directory.GetFiles(directory, "*.csv");
                                    
                                    //筛选出文件名包含"en-US"或"zh-CN"的.csv文件
                                    var      filteredFiles = files.Where(file =>
                                        Path.GetFileName(file).Contains("en-US") ||
                                        Path.GetFileName(file).Contains("zh-CN"));
                                    
                                    //获取报警文本，并且添加到LIST<ProDiagInfos>
                                    foreach (string file in filteredFiles)
                                    {
                                        using (StreamReader reader = new StreamReader(file))
                                        {
                                            string fileName = Path.GetFileName(file);
                                            var    data     = reader.Analyze(fileName);
                                            proDiagInfos.AddRange(data);
                                        }
                                    }
                                    feedback.Log(NotificationIcon.Information, "获取报警文本，并且添加到LIST<ProDiagInfos>");
                                }
                                
                                #endregion

                                #region 获取Supervision接口信息
                                
                                //定义Supervision接口信息列表
                                var supervisionInfos = new List<SupervisionInfo>();

                                //查询文件夹路径是否存在
                                if (Directory.Exists(directory))
                                {
                                    //获取文件夹中所有的.xml文件
                                    string[] files = Directory.GetFiles(directory, "*.xml");

                                    //获取Supervision接口信息，并写入到List<SupervisionInfo>
                                    foreach (string file in files)
                                    {
                                        supervisionInfos.AddRange(file.Analyze(proDiagInfos));
                                    }
                                    feedback.Log(NotificationIcon.Information, "获取Supervision接口信息，并写入到List<SupervisionInfo>");

                                }

                                #endregion
                                
                                #region 获取HMI的Connection名称 
                                
                                //获取HMI目标
                                //Device    hmiDevice = mainForm.device;
                                string selectedName = mainForm.SelectName;
                                Device hmiDevice    = projectBase.FindDeviceByName(selectedName);

                                HmiTarget hmiTarget = hmiDevice.GetHmiTarget();
                                
                                //导出默认变量表
                                TagTable defaultTagTable           = hmiTarget.TagFolder.DefaultTagTable;
                                string   defaultTagTableExportPath = $@"{directory}\{defaultTagTable.Name}.xml";
                                defaultTagTable.ExportInfo(defaultTagTableExportPath);

                                //获取Connection名称
                                XmlDocument xmlDocument = new XmlDocument();
                                xmlDocument.Load(defaultTagTableExportPath);
                                string connection = xmlDocument.GetName("Connection");
                                
                                feedback.Log(NotificationIcon.Information, $"获取Connection名称:{connection}");


                                #endregion

                                #region 创建自动生成报警变量表

                                //定义变量表.xml文件
                                XmlDocument tagTable = XmlAnalyze.TagTable();
                                tagTable.SetAttribute("Name","自动生成报警变量表");
                                    
                                //定义变量.xml文件
                                XmlDocument tag = XmlAnalyze.Tag();

                                //定义触摸屏报警文本.xlsx文件
                                IWorkbook hmiAlarmTextEnglish = XlsAnalyze.HmiAlarmText("en-US");
                                IWorkbook hmiAlarmTextChinese = XlsAnalyze.HmiAlarmText("zh-CN");

                                #endregion

                                #region 处理报警文本和变量

                                //定义报警文本和变量处理的临时变量
                                int    tagCount = 1; //变量的名称递增
                                int    tagID    = 1; //变量的ID递增
                                int    textID   = 1; //报警文本ID递增
                                string old      = null;
                                string tagName  = null;
                                
                                //写入Hmi变量表.xml文件
                                foreach (SupervisionInfo xmlInfo in supervisionInfos)
                                {
                                    if (old != xmlInfo.DB_Name)
                                    {
                                        //设置ID
                                        tag.SetId(tagID); 
                                        
                                        //计算和设置地址%DB1.DBW1
                                        string logicalAddress = xmlInfo.Offset.GetLogicalAddress(xmlInfo.DB_Number);
                                        tag.SetAttribute("LogicalAddress", logicalAddress);
                                        
                                        //计算和设置名称P1001_O_1
                                        tagName = $"{xmlInfo.DB_Name}_O_1";
                                        tag.SetAttribute("Name", tagName);
                                        
                                        //设置连接点信息
                                        tag.SetValue("Connection", connection);
                                        
                                        //tagTable的XML插入一个变量
                                        tagTable.Insert(tag);

                                        //递增
                                        old = xmlInfo.DB_Name;
                                        tagID++;
                    
                                        string[] parts = xmlInfo.Offset.Split('.');
                                        if (parts.Length >0)
                                        {
                                            tagCount = Convert.ToInt32(parts[1]) + 1;
                                        }
                                        else
                                        {
                                            tagCount = 1;
                                        }
                                    }
                                    
                                    //诊断接口超过8时
                                    if (tagCount > 8 && tagCount % 8 == 1)
                                    {
                                        //设置ID
                                        tag.SetId(tagID); 
                                        
                                        //计算和设置地址%DB1.DBW1
                                        string logicalAddress = xmlInfo.Offset.GetLogicalAddress(xmlInfo.DB_Number);
                                        tag.SetAttribute("LogicalAddress", logicalAddress);
                                        
                                        //计算和设置名称P1001_O_1
                                        tagName = $"{xmlInfo.DB_Name}_O_{(tagCount / 8) + (tagCount % 8)}";
                                        tag.SetAttribute("Name", tagName);
                                        
                                        //设置连接点信息
                                        tag.SetValue("Connection", connection);
                                        
                                        //tagTable的XML插入一个变量
                                        tagTable.Insert(tag);

                                        //递增
                                        tagID++;
                                    }
                                    
                                    if (old == xmlInfo.DB_Name)
                                    {
                                        tagCount++;
                                    }
                                    
                                    //写入报警文本
                                    hmiAlarmTextEnglish.SetText(xmlInfo, "en-US", textID, tagName);
                                    hmiAlarmTextChinese.SetText(xmlInfo, "zh-CN", textID, tagName);

                                    textID++;
                                }
                                feedback.Log(NotificationIcon.Information, "已写入Hmi变量表.xml文件");

                                #endregion

                                #region 导入HMI变量表

                                string exportPath = $@"{directory}\{tagTable.GetAttribute("Name")}.xml";
                                tagTable.Save(exportPath);
                                hmiTarget.TagFolder.ImportInfo(exportPath);
                                feedback.Log(NotificationIcon.Information, $"已导入Hmi变量表{tagTable.GetAttribute("Name")}");

                                #endregion

                                #region 保存报警文本表格

                                //保存报警文本表格（英文）
                                hmiAlarmTextEnglish.Save(englishPath);
                                feedback.Log(NotificationIcon.Information, $"已保存报警文本表格（英文）{englishPath}");

                                //保存报警文本表格（中文）
                                hmiAlarmTextChinese.Save(chinesePath);
                                feedback.Log(NotificationIcon.Information, $"已保存报警文本表格（中文）{chinesePath}");

                                #endregion
                            }

                            messageBox.ShowNotification(NotificationIcon.Success, "创建触摸屏报警完成",
                                $"已导入触摸屏变量表[自动生成报警变量表] \n 已导出报警文本表格：\n {englishPath} \n {chinesePath}");
                            
                            feedback.Log(NotificationIcon.Success, "<创建触摸屏报警>完成");

                            //创建回滚
                            if (transaction.CanCommit)
                            {
                                transaction.CommitOnDispose();
                                feedback.Log(NotificationIcon.Information, "已创建回滚");
                            }
                        }
                    }
                }
            }
            catch (EngineeringException ex)
            {
                messageBox.ShowException(ex);
                feedback.Log(NotificationIcon.Error,ex.Message);
                throw;
            }
            catch (Exception ex)
            {
                messageBox.ShowException(ex);
                feedback.Log(NotificationIcon.Error,ex.Message);
                throw;
            }
            finally
            {
                if (directory != null)
                {
                    directory.DeleteDirectoryAndContents();//删除导出的.xml文件的文件夹
                    feedback.Log(NotificationIcon.Information,$"已删除：{directory}文件夹");
                }
            }
        }

        /// <summary>
        /// 指定ProDiagFB
        /// </summary>
        /// <param name="menuSelectionProvider"></param>
        /// <param name="item"></param>
        private void ProDiag_OnClick(MenuSelectionProvider<InstanceDB> menuSelectionProvider, string item)
        {
#if DEBUG
            System.Diagnostics.Debugger.Launch();
#endif
            //反馈API
            FeedbackContext feedback  = _tiaPortal.GetFeedbackContext();
            feedback.Log(NotificationIcon.Information,"<指定ProDiagFB>");
            
            //MessageBox API
            MessageBox messageBox = _tiaPortal.GetMessageBox();
            
            try
            {
                //创建独占窗口
                using (ExclusiveAccess exclusiveAccess = _tiaPortal.ExclusiveAccess("指定ProDiagFB"))
                {
                    //获取项目实例
                    ProjectBase projectBase = _tiaPortal.GetProjectBase();

                    //如果项目实例为空，抛出异常
                    if (projectBase == null)
                    {
                        throw new EngineeringObjectDisposedException("无法获取项目实例");
                    }

                    using (Transaction transaction = exclusiveAccess.Transaction(projectBase, "指定ProDiagFB"))
                    {

                        //新建ProDiagFB选择窗体
                        ProDiagForm proDiagForm = new ProDiagForm();

                        if (item =="1")
                        {
                            //获取所有ProDiagFB，并把Name和PlcBlock写入到数据
                            PlcSoftware   plcSoftware    = menuSelectionProvider.GetSelection().GetPlcSoftware();
                            PlcBlockGroup plcBlockGroups = plcSoftware.BlockGroup;
                            var           proDiagFbs   = Command.GetProDiagFB(plcBlockGroups);
                        
                            proDiagForm.SetDevices(proDiagFbs); //填充数据

                            if (proDiagForm.ShowDialog() != DialogResult.OK)
                            {
                                feedback.Log(NotificationIcon.Error, $"用户窗口取消");
                                return;
                            }
                            feedback.Log(NotificationIcon.Information, $"获取输入信息:{proDiagForm.Block}");
                        }
                        
                        //轮询选中的实例块
                        foreach (InstanceDB instanceDb in menuSelectionProvider.GetSelection())
                        {
                            //监听独占窗口取消按钮
                            if (exclusiveAccess.IsCancellationRequested)
                            {
                                feedback.Log(NotificationIcon.Error, $"监听独占窗口取消");
                                return;
                            }

                            //获取当前AssignedProDiagFB
                            var    AssignedProDiagFB = instanceDb.GetAttribute("AssignedProDiagFB");
                            string value = string.Empty;
                            
                            if (AssignedProDiagFB != null)
                            {
                                value = AssignedProDiagFB.ToString();
                            }

                            //删除AssignedProDiagFB
                            if (item == "")
                            {
                                if (value != item)
                                {
                                    exclusiveAccess.Text = instanceDb.Name + "删除指定FB监控";
                                    instanceDb.SetAttribute("AssignedProDiagFB", item);
                                    feedback.Log(NotificationIcon.Information,
                                        $"删除指定FB监控:{instanceDb.Name}[{instanceDb.ProgrammingLanguage}{instanceDb.Number}]");
                                }

                                break;
                            }

                            //指定AssignedProDiagFB
                            if (value != proDiagForm.Block) //ProDiag_FB1
                            {
                                exclusiveAccess.Text = instanceDb.Name + $"指定FB监控为{proDiagForm.Block}";
                                instanceDb.SetAttribute("AssignedProDiagFB", proDiagForm.Block);
                                feedback.Log(NotificationIcon.Information,
                                    $"指定FB监控:{instanceDb.Name}[{instanceDb.ProgrammingLanguage}{instanceDb.Number}]-{proDiagForm.Block}");
                            }
                        }

                        //创建回滚
                        if (transaction.CanCommit)
                        {
                            transaction.CommitOnDispose();
                            feedback.Log(NotificationIcon.Information, "已创建回滚");
                        }
                    }

                    //导入完成
                    messageBox.ShowNotification(NotificationIcon.Success, "完成","指定ProDiagFB完成");
                    feedback.Log(NotificationIcon.Success, $"<指定ProDiagFB>完成");
                }
            }
            catch (EngineeringException ex)
            {
                messageBox.ShowException(ex);
                feedback.Log(NotificationIcon.Error,ex.Message);
                throw;
            }
            catch (Exception ex)
            {
                messageBox.ShowException(ex);
                feedback.Log(NotificationIcon.Error,ex.Message);
                throw;
            }
        }
    }
}