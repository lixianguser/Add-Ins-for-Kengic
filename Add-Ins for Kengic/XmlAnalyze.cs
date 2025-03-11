using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;

namespace Kengic
{
    public static class XmlAnalyze
    {
        #region 写入Xml数据的方法

         /// <summary>
        /// 生成触摸屏空变量表的Xml文档
        /// </summary>
        /// <returns></returns>
        public static XmlDocument TagTable()
        {
            string str = $"<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n" +
                         $"<Document>\r\n  " +
                         $"<Engineering version=\"V18\" />\r\n  " +
                         $"<Hmi.Tag.TagTable ID=\"0\">\r\n    " +
                         $"<AttributeList>\r\n      " +
                         $"<Name>TagTable</Name>\r\n    " +
                         $"</AttributeList>\r\n    " +
                         $"<ObjectList>\r\n    " +
                         $"</ObjectList>\r\n  " +
                         $"</Hmi.Tag.TagTable>\r\n" +
                         $"</Document>";
            
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.LoadXml(str);

            return xmlDocument;
        }
        
        /// <summary>
        /// 生成触摸屏单个变量的Xml文档
        /// </summary>
        /// <returns></returns>
        public static XmlDocument Tag()
        {
            string str =
                $"<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n" +
                $"<Hmi.Tag.Tag ID=\"1\" CompositionName=\"Tags\">\r\n        " +
                $"<AttributeList>\r\n          " +
                $"<AcquisitionTriggerMode>Continuous</AcquisitionTriggerMode>\r\n          " +
                $"<AddressAccessMode>Absolute</AddressAccessMode>\r\n          " +
                $"<Coding>Binary</Coding>\r\n          " +
                $"<ConfirmationType>None</ConfirmationType>\r\n          " +
                $"<GmpRelevant>false</GmpRelevant>\r\n          " +
                $"<JobNumber>0</JobNumber>\r\n          " +
                $"<Length>2</Length>\r\n          " +
                $"<LinearScaling>false</LinearScaling>\r\n          " +
                $"<LogicalAddress>%DB22.DBW2</LogicalAddress>\r\n          " +
                $"<MandatoryCommenting>false</MandatoryCommenting>\r\n          " +
                $"<Name>3P1</Name>\r\n          " +
                $"<Persistency>false</Persistency>\r\n          " +
                $"<QualityCode>false</QualityCode>\r\n          " +
                $"<ScalingHmiHigh>100</ScalingHmiHigh>\r\n          " +
                $"<ScalingHmiLow>0</ScalingHmiLow>\r\n          " +
                $"<ScalingPlcHigh>10</ScalingPlcHigh>\r\n          " +
                $"<ScalingPlcLow>0</ScalingPlcLow>\r\n          " +
                $"<StartValue />\r\n          " +
                $"<SubstituteValue />\r\n          " +
                $"<SubstituteValueUsage>None</SubstituteValueUsage>\r\n          " +
                $"<Synchronization>false</Synchronization>\r\n          " +
                $"<UpdateMode>ProjectWide</UpdateMode>\r\n          " +
                $"<UseMultiplexing>false</UseMultiplexing>\r\n        " +
                $"</AttributeList>\r\n        " +
                $"<LinkList>\r\n          " +
                $"<AcquisitionCycle TargetID=\"@OpenLink\">\r\n            " +
                $"<Name>1 s</Name>\r\n          " +
                $"</AcquisitionCycle>\r\n          " +
                $"<Connection TargetID=\"@OpenLink\">\r\n            " +
                $"<Name>RC11_HMI_1</Name>\r\n          " +
                $"</Connection>\r\n          " +
                $"<DataType TargetID=\"@OpenLink\">\r\n            " +
                $"<Name>Int</Name>\r\n          " +
                $"</DataType>\r\n          " +
                $"<HmiDataType TargetID=\"@OpenLink\">\r\n            " +
                $"<Name>Int</Name>\r\n          " +
                $"</HmiDataType>\r\n        " +
                $"</LinkList>\r\n      " +
                $"</Hmi.Tag.Tag>";

            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.LoadXml(str);

            return xmlDocument;
        }
        
        /// <summary>
        /// 设定ID值
        /// </summary>
        /// <param name="xmlDocument">Xml文档</param>
        /// <param name="id">设定的ID值</param>
        public static void SetId(this XmlDocument xmlDocument, int id)
        {
            XmlAttribute xmlAttribute = xmlDocument.SelectSingleNode("Hmi.Tag.Tag")?.Attributes?["ID"];
            if (xmlAttribute != null) xmlAttribute.Value = id.ToString();
        }
        
        /// <summary>
        /// 设定HMI连接信息
        /// </summary>
        /// <param name="xmlDocument">Xml文档</param>
        /// <param name="linkName">查找连接节点</param>
        /// <param name="setValue">设定值</param>
        public static void SetValue(this XmlDocument xmlDocument, string linkName, string setValue)
        {
            // 查找 AttributeList 内指定名称的节点
            XmlNode xmlNode = xmlDocument?.SelectSingleNode($"//LinkList/{linkName}/Name");
            if (xmlNode != null)
            {
                //Console.WriteLine("Name value: " + xmlNode.InnerText);
                xmlNode.InnerText = setValue;
            }
            else
            {
                Console.WriteLine($"{linkName} node not found.");
            }
        }
        
        /// <summary>
        /// 变量表中插入单个变量
        /// </summary>
        /// <param name="xmlDocument">变量表的Xml文档</param>
        /// <param name="tagXmlDocument">单个变量的Xml文档</param>
        /// <returns></returns>
        public static XmlDocument Insert(this XmlDocument xmlDocument, XmlDocument tagXmlDocument)
        {
            // 查找 ObjectList 节点
            XmlNode objectListNode = xmlDocument.SelectSingleNode("//ObjectList");
            // 导入 Hmi.Tag.Tag 文档的根节点
            if (tagXmlDocument.DocumentElement != null)
            {
                XmlNode importedNode = xmlDocument.ImportNode(tagXmlDocument.DocumentElement, true);

                // 将新的 Hmi.Tag.Tag 节点插入到 ObjectList 节点中
                if (objectListNode != null) objectListNode.AppendChild(importedNode);
            }

            return xmlDocument;
        }

        #endregion

        #region 获取Xml数据的方法
        
        /// <summary>
        /// Xml文档的第一个Sections的命名空间
        /// </summary>
        /// <remarks>xmlns="http://www.siemens.com/automation/Openness/SW/Interface/v5"</remarks>
        private static XmlNamespaceManager nsmgrSections;
        
        /// <summary>
        /// Xml文档的第一个BlockInstSupervisionGroups的命名空间
        /// </summary>
        /// <remarks>xmlns="http://www.siemens.com/automation/Openness/SW/Interface/v5"</remarks>
        private static XmlNamespaceManager nsmgrSupervision;
        
        /// <summary>
        /// 获取指定请求节点获取命名空间并添加到命名空间
        /// </summary>
        /// <param name="xmlDocument">Xml文档</param>
        /// <param name="TagName">有效的请求节点：Sections、BlockInstSupervisionGroups</param>
        private static void GetXmlns(this XmlDocument xmlDocument , string TagName)
        {
            XmlNode xmlNode;

            switch (TagName)
            {
                case "Sections":
                    xmlNode       = xmlDocument.GetElementsByTagName("Sections")[0];//获取"Sections"节点
                    nsmgrSections = new XmlNamespaceManager(xmlDocument.NameTable);
                    if (xmlNode?.NamespaceURI != null) nsmgrSections.AddNamespace("itf", xmlNode.NamespaceURI);
                    return;
                case "BlockInstSupervisionGroups":
                    xmlNode = xmlDocument.GetElementsByTagName("BlockInstSupervisionGroups")[0];//获取"BlockInstSupervisionGroups"节点
                    nsmgrSupervision = new XmlNamespaceManager(xmlDocument.NameTable);
                    if (xmlNode?.NamespaceURI != null) nsmgrSupervision.AddNamespace("sv", xmlNode.NamespaceURI);
                    return;
            }
        }
        
        /// <summary>
        /// 获取节点名称的值
        /// </summary>
        /// <param name="xmlDocument">Xml文档</param>
        /// <param name="nodeName">要查找的节点名称</param>
        /// <returns></returns>
        public static string GetName(this XmlDocument xmlDocument,string nodeName)
        {
            switch (nodeName)
            {
                case "Connection":
                    // 查找 LinkList 内指定名称的节点
                    XmlNode xmlNode = xmlDocument?.SelectSingleNode($"//LinkList/{nodeName}/Name");
                    if (xmlNode != null)
                    {
                        //Console.WriteLine("Name value: " + xmlNode.InnerText);
                        return xmlNode.InnerText;
                    }
                    Console.WriteLine($"{nodeName} node not found.");
                    return null;
            }
            
            return null;
        }
        
        /// <summary>
        /// 设置属性的值
        /// </summary>
        /// <param name="xmlDocument">Xml文档</param>
        /// <param name="attributeName">查找属性</param>
        /// <param name="setValue">设定值</param>
        public static void SetAttribute(this XmlDocument xmlDocument,string attributeName,string setValue)
        {
            // 查找 AttributeList 内指定名称的节点
            XmlNode xmlNode = xmlDocument?.SelectSingleNode($"//AttributeList/{attributeName}");
            if (xmlNode != null)
            {
                //Console.WriteLine("Name value: " + xmlNode.InnerText);
                xmlNode.InnerText = setValue;
            }
            else
            {
                Console.WriteLine($"{attributeName} node not found.");
            }
        }

        /// <summary>
        /// 通过Xml文档获取属性的值
        /// </summary>
        /// <param name="xmlDocument">Xml文档</param>
        /// <param name="attributeName">查找属性</param>
        /// <returns></returns>
        public static string GetAttribute(this XmlDocument xmlDocument,string attributeName)
        {
            // 查找 AttributeList 内指定名称的节点
            XmlNode xmlNode = xmlDocument?.SelectSingleNode($"//AttributeList/{attributeName}");
            if (xmlNode != null)
            {
                //Console.WriteLine("Name value: " + xmlNode.InnerText);
                return xmlNode.InnerText;
            }
            Console.WriteLine($"{attributeName} node not found.");
            return null;
        }

        /// <summary>
        /// 通过节点获取指定请求值的值
        /// </summary>
        /// <param name="node">接口的节点</param>
        /// <param name="request">有效的请求值：Name、Offset、Datatype、IsStruct、IsUdt</param>
        /// <returns>字符串类型</returns>
        /// <exception cref="ArgumentException">无效的请求值.</exception>
        private static string GetValue(this XmlNode node , string request)
        {
            XmlNode      xmlNode;
            XmlAttribute xmlAttribute;
            XmlNodeList  xmlNodeList;
            
            switch (request)
            {
                case "Name":
                    xmlAttribute = node.Attributes?["Name"];
                    return xmlAttribute != null ? xmlAttribute.Value : string.Empty;
                case "Offset":
                    xmlNode = node.SelectSingleNode("itf:AttributeList/itf:IntegerAttribute[@Name='Offset']", nsmgrSections);
                    return xmlNode != null ? xmlNode.InnerText : string.Empty;
                case "Datatype":
                    xmlAttribute = node.Attributes?["Datatype"];
                    return xmlAttribute != null ? xmlAttribute.Value : string.Empty;
                case "IsStruct":
                    xmlNodeList = node.SelectNodes("itf:Member", nsmgrSections);
                    return xmlNodeList != null && xmlNodeList.Count >0 ? "true" : "false";
                case "IsUdt":
                    xmlNode = node.SelectSingleNode("itf:Sections", nsmgrSections);
                    return xmlNode != null ? "true" : "false";
                case "StartValue":
                    xmlNodeList = node.ChildNodes;
                    foreach (XmlNode members in xmlNodeList)
                    {
                        if (members.Name == "StartValue")
                        {
                            XmlNode startValue = node.SelectSingleNode("itf:StartValue", nsmgrSections);
                            return startValue != null ? startValue.InnerText : string.Empty;
                        }
                    }
                    return string.Empty;
                default:
                    throw new ArgumentException("无效的请求值.", nameof(request));
            }
        }
        
        /// <summary>
        /// 获取监控节点列表
        /// </summary>
        /// <returns></returns>
        private static XmlNodeList GetBlockInstSupervisionNodes(this XmlDocument xmlDocument)
        {
            xmlDocument.GetXmlns("BlockInstSupervisionGroups");
            // 获取所有 BlockInstSupervision 节点
            XmlNodeList supervisionNodes = xmlDocument.SelectNodes("//sv:BlockInstSupervision", nsmgrSupervision);
            return supervisionNodes;
        }
        
        /// <summary>
        /// 获取Supervision信息
        /// </summary>
        /// <param name="supervisionNode">诊断信息节点</param>
        /// <param name="attribute">目标属性:Number、StateStruct、BlockTypeSupervisionNumber</param>
        /// <returns></returns>
        private static string GetBlockInstSupervision(this XmlNode supervisionNode,string attribute)
        {
            switch (attribute)
            {
                case "Number":
                    XmlNode xmlNode = supervisionNode.SelectSingleNode("sv:Number", nsmgrSupervision);
                    return xmlNode != null ? xmlNode.InnerText : string.Empty;
                case "StateStruct":
                    string name = supervisionNode.SelectSingleNode("sv:StateStruct", nsmgrSupervision)?
                        .Attributes?["Name"]?.Value;
                    return name;
                case "BlockTypeSupervisionNumber":
                    string type = supervisionNode.SelectSingleNode("sv:BlockTypeSupervisionNumber", nsmgrSupervision)?.InnerText;
                    return type;
            }

            return null;
        }
        
        /// <summary>
        /// 通过接口名称获取接口节点
        /// </summary>
        /// <param name="interfaceName">接口名称</param>
        /// <param name="xmlDocument"></param>
        /// <returns></returns>
        private static XmlNode GetInterfaceMember(this XmlDocument xmlDocument ,string interfaceName)
        {
            xmlDocument.GetXmlns("Sections");
            XmlNode xmlNode = xmlDocument.SelectSingleNode($"//itf:Member[@Name='{interfaceName}']", nsmgrSections);
            return xmlNode;
        }
        
        #endregion

        #region 辅助计算方法

        /// <summary>计算接口的偏移量</summary>
        /// <param name="offset">接口的偏移量</param>
        /// <returns>40.1</returns>
        private static string CalOffset(this int offset)
        {
            int    num  = offset / 8;
            string str1 = num.ToString();
            num = offset % 8;
            string str2 = num.ToString();
            return str1 + "." + str2;
        }
        
        /// <summary>
        /// 处理诊断名称，获取接口名称
        /// </summary>
        /// <param name="name">诊断名称</param>
        /// <returns></returns>
        private static string GetSupervisionInterface(string name)
        {
            const string pattern = @"#(.*?)_O_";
            Match        match   = Regex.Match(name, pattern);
            if (match.Success)
            {
                string str = match.Groups[1].Value;
                return str;
            }

            return null;
        }
        
        /// <summary>
        /// 获取单个变量的逻辑地址
        /// </summary>
        /// <param name="offset"></param>
        /// <param name="dBNumber"></param>
        /// <returns></returns>
        public static string GetLogicalAddress(this string offset, string dBNumber)
        {
            // 查找 "." 的位置
            int dotIndex = offset.IndexOf('.');
            // 截取 "." 之前的部分
            string result    = offset.Substring(0, dotIndex);
            int    intResult = int.Parse(result) - 1;
            
            if (result == "0")
            {
                intResult = 0;
            }
            
            string str       = $"%DB{dBNumber}.DBW{intResult}";
            return str;
        }

        #endregion

        #region 处理Xml文件
        
        /// <summary>
        /// 处理Xml文件获得相关数据
        /// </summary>
        /// <param name="file">Xml文件路径</param>
        /// <param name="proDiag">诊断数据</param>
        /// <returns></returns>
        public static List<SupervisionInfo> Analyze(this string file, List<ProDiagInfo> proDiag)
        {
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(file);
            var supervisionInfos = new List<SupervisionInfo>();

            //获取所有监控的节点
            XmlNodeList blockInstSupervisionNodes = xmlDocument.GetBlockInstSupervisionNodes();
            if (blockInstSupervisionNodes ==  null)
            {
                throw new NullReferenceException("无法获取BlockInstSupervisionNodes");
            }
            
            foreach (XmlNode blockInstSupervisionNode in blockInstSupervisionNodes)
            {
                //DB的名称:P3102
                string name   = xmlDocument.GetAttribute("Name"); 
                //DB的编号:3102
                string dB_Number = xmlDocument.GetAttribute("Number"); 
                string stateStruct =
                    blockInstSupervisionNode
                        .GetBlockInstSupervision("StateStruct"); //监控的名称:P3102.#Y_DV_M01_Started_Man_O_1
                string number = blockInstSupervisionNode.GetBlockInstSupervision("Number"); //监控的编号:2755
                string blockTypeSupervisionNumber =
                    blockInstSupervisionNode.GetBlockInstSupervision("BlockTypeSupervisionNumber"); //监控的类型:1
                //获取监控的接口节点
                string  itf     = GetSupervisionInterface(stateStruct);
                XmlNode xmlNode = xmlDocument.GetInterfaceMember(itf);
                string  offset  = int.Parse(xmlNode.GetValue("Offset")).CalOffset(); //监控的偏移量:178.0
                
                //获取起始值
                string count             = Regex.Match(itf, @"\d+").Value;
                string startValue        = string.Empty;
                if (count.Length < 2)
                {
                    count = string.Empty;
                }
                XmlNode convNumber = xmlDocument.GetInterfaceMember($"X_ConveyorNumber{count}");
                if (convNumber != null)
                {
                    startValue = convNumber.GetValue("StartValue");
                }

                // 添加数据
                SupervisionInfo xmlInfo = new SupervisionInfo
                {
                    StateStruct = stateStruct,
                    Number      = number,
                    Offset      = offset,
                    DB_Name     = name,
                    DB_Number   = dB_Number
                };
                //添加报警类型
                switch (blockTypeSupervisionNumber)
                {
                    case "1":
                        xmlInfo.BlockTypeSupervisionNumber = "Alarms";
                        break;
                    case "2":
                        xmlInfo.BlockTypeSupervisionNumber = "Events";
                        break;
                    case "3":
                        break;
                }

                //添加报警文本
                // 获取对应的所有英文 AlarmText
                var englishAlarmTexts = proDiag
                    .Where(p => p.Identification == number 
                                && p.Language == "en-US" 
                                && p.AlarmText.Contains(name))
                    .Select(p => p.AlarmText)
                    .ToList();

                string longestAlarmText = string.Empty;

                foreach (var item in englishAlarmTexts.Where(item => item.Length > longestAlarmText.Length))
                {
                    longestAlarmText = item;
                    //设备号的起始值替换P@4%5u@或P@4%4d@
                    longestAlarmText = Regex.Replace(item, @"@[^@]*@", startValue);
                } 
                
                // 最终将最长的字符串赋值给xmlInfo.AlarmTextEn
                xmlInfo.AlarmTextEn = longestAlarmText;

                // 获取对应的所有中文 AlarmText
                var chineseAlarmTexts = proDiag
                    .Where(p => p.Identification == number 
                                && p.Language == "zh-CN" 
                                && p.AlarmText.Contains(name))
                    .Select(p => p.AlarmText)
                    .ToList();
                
                longestAlarmText = string.Empty;

                foreach (var item in chineseAlarmTexts.Where(item => item.Length > longestAlarmText.Length))
                {
                    longestAlarmText = item;
                    //设备号的起始值替换P@4%5u@或P@4%4d@
                    longestAlarmText = Regex.Replace(item, @"@[^@]*@", startValue);
                } 

                // 最终将最长的字符串赋值给xmlInfo.AlarmTextZh
                xmlInfo.AlarmTextZh = longestAlarmText;

                supervisionInfos.Add(xmlInfo);
            }

            //按偏移量顺序排序
            supervisionInfos.Sort((x, y) => double.Parse(x.Offset).CompareTo(double.Parse(y.Offset)));

            return supervisionInfos;
        }
        
        #endregion
    }
    
    /// <summary>
    /// SupervisionInfo诊断接口的信息
    /// </summary>
    public class SupervisionInfo
    {
        /// <summary>
        /// 诊断名称
        /// </summary>
        public string StateStruct { get; set; }

        /// <summary>
        /// 诊断编号
        /// </summary>
        public string Number { get; set; }

        /// <summary>
        /// 诊断类型
        /// </summary>
        public string BlockTypeSupervisionNumber { get; set; }

        /// <summary>
        /// 报警文本-英文
        /// </summary>
        public string AlarmTextEn { get; set; }

        /// <summary>
        /// 报警文本-中文
        /// </summary>
        public string AlarmTextZh { get; set; }

        /// <summary>
        /// 偏移量
        /// </summary>
        public string Offset { get; set; }

        /// <summary>
        /// DB块名称
        /// </summary>
        public string DB_Name { get; set; }

        /// <summary>
        /// DB块编号
        /// </summary>
        public string DB_Number { get; set; }
    }
}