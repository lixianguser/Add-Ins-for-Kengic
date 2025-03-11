using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;

namespace Kengic
{
    public class SCADA
    {
        #region 程序接口
        
        /// <summary>
        /// 获取Xml的命名空间
        /// </summary>
        private XmlNamespaceManager _xmlns;

        /// <summary>
        /// 获取实例的DB名称
        /// </summary>
        public string InstanceName;

        /// <summary>
        /// 获取实例的FB名称
        /// </summary>
        public string InstanceOfName;

        /// <summary>
        /// 获取DB的编号
        /// </summary>
        public string Number;

        /// <summary>
        /// 获取程序语言类型
        /// </summary>
        public string ProgrammingLanguage;

        /// <summary>
        /// 获取InstanceDB Xml文件路径
        /// </summary>
        public string IDbXmlFilePath;

        /// <summary>
        /// 获取FB Xml文件路径
        /// </summary>
        public string FBXmlFilePath;

        /// <summary>
        /// 写入数据流
        /// </summary>
        public StreamWriter StreamWriter;

        /// <summary>
        /// 获取监控列表
        /// </summary>
        private readonly List<string[]> _supervisions = new List<string[]>();
        
        #endregion

        public void Run()
        {
            //读取Xml文件
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(IDbXmlFilePath);
            //获取"Sections"节点
            XmlNode sections = xmlDocument.GetElementsByTagName("Sections")[0];
            if (sections ==null )
            {
                throw new NullReferenceException("DB获取Sections节点失败");
            }
            //获取"BlockInstSupervisionGroups"节点
            XmlNode bISupervisionGroups = xmlDocument.GetElementsByTagName("BlockInstSupervisionGroups")[0];
            if (bISupervisionGroups ==null )
            {
                throw new NullReferenceException("DB获取BlockInstSupervisionGroups节点失败");
            }
            //从Sections和BlockInstSupervisionGroups节点获取NameSpace
            _xmlns = GetXmlns(xmlDocument, sections.NamespaceURI, bISupervisionGroups.NamespaceURI, "");

            //获取FB的监控属性
            xmlDocument.Load(FBXmlFilePath);
            //获取"BlockTypeSupervisions"节点
            XmlNode bTSupervisions = xmlDocument.GetElementsByTagName("BlockTypeSupervisions")[0];
            if (bTSupervisions ==null )
            {
                throw new NullReferenceException("BlockTypeSupervisions");
            }
            //从Sections、BlockInstSupervisionGroups、BlockTypeSupervisions节点获取NameSpace
            _xmlns = GetXmlns(xmlDocument, sections.NamespaceURI, bISupervisionGroups.NamespaceURI, bTSupervisions.NamespaceURI);
            //获取诊断变量和报警文本
            foreach (XmlNode supervisionGroup in bISupervisionGroups)
            {
                foreach (XmlNode supervision in supervisionGroup)
                {
                    string[] item = new string[4];
                    item[0] = GetName(GetSupervision(supervision));//DB监控定义
                    item[1] = GetSupervisionType(supervision);//DB监控类型
                    foreach (XmlNode bTSupervision in bTSupervisions)
                    {
                        if (GetName(GetSupervision(supervision)).Contains(GetName(GetSupervisedOperand(bTSupervision))))
                        {
                            item[2] = GetMultiLanguageText(bTSupervision, "en-US");//获取英文报警文本
                            item[3] = GetMultiLanguageText(bTSupervision, "zh-CN");//获取中文报警文本
                        }
                    }
                    _supervisions.Add(item);
                }
            }

            //获取变量和偏移量
            foreach (XmlNode members in sections)
            {
                foreach (XmlNode member in members)
                {
                    string offset;
                    switch (GetName(member))
                    {
                        //获取运行状态变量
                        case "Q_M01_Forward": //正转运行状态
                        case "Q_M01_Reverse": //反转运行状态
                        //提升机、顶升机
                        case "Q_M02_Forward": 
                        case "Q_M02_Reverse": 
                        //拆盘机
                        case "Q_M01_Raise":
                        case "Q_M01_Lower":
                        case "Q_M02_ExtendLeftFork":
                        case "Q_M02_RetractLeftFork":
                        case "Q_M03_ExtendRightFork":
                        case "Q_M03_RetractRightFork":
                            offset = CalOffset(GetOffset(member));
                            SetSacdaTag(offset, GetName(member),"X");
                            break;
                        case "R_DV_PalletStackCount": //叠盘机托盘数量
                            offset = CalOffset(GetOffset(member));
                            SetSacdaTag(offset, "R_DV_PalletStackCount","INT");
                            break;
                        case "Z_Data"://托盘数据
                            foreach (XmlNode memberOfZ_Data in GetSection(member))
                            {
                                switch (GetName(memberOfZ_Data))
                                {
                                    case "Position01":
                                        foreach (XmlNode memberOfPosition01 in GetSection(memberOfZ_Data))
                                        {
                                            offset = CalOffset(GetOffset(member) + GetOffset(memberOfZ_Data) +
                                                               GetOffset(memberOfPosition01));
                                            string input =
                                                $"{GetName(member)}_{GetName(memberOfZ_Data)}_{GetName(memberOfPosition01)}";
                                            switch (GetName(memberOfPosition01))
                                            {
                                                
                                                case "Origin": //获取原地址变量
                                                case "Destination": //获取目的地变量
                                                    SetSacdaTag(offset, input,"UINT");
                                                    break;
                                                case "PalletID":
                                                    //获取条码的字符长度
                                                    foreach (XmlNode item in GetSection(memberOfPosition01))
                                                    {
                                                        string barcodeType = item.Attributes.GetNamedItem("Datatype").Value;
                                                        SetSacdaTag(offset, GetBarcodeLength(barcodeType), input,"STRING");
                                                    }
                                                    break;
                                                //托盘外检拒绝信息 
                                                case "RejectCode":
                                                    foreach (XmlNode item in GetSection(memberOfPosition01))
                                                    {
                                                        offset = CalOffset(GetOffset(member) + GetOffset(memberOfZ_Data) +
                                                                           GetOffset(memberOfPosition01) + GetOffset(item));
                                                        string reject =
                                                            $"{GetName(member)}_{GetName(memberOfZ_Data)}_{GetName(item)}";
                                                        SetSacdaTag(offset, reject,"X");
                                                    }
                                                    break;
                                            }
                                        }
                                        break;
                                    case "Position02":
                                        // 判断InstanceOfName实例决定是否需要占位2数据
                                        if (!InstanceOfName.Contains("_2P"))
                                            break;
                                        foreach (XmlNode memberOfPosition02 in GetSection(memberOfZ_Data))
                                        {
                                            offset = CalOffset(GetOffset(member) + GetOffset(memberOfZ_Data) +
                                                               GetOffset(memberOfPosition02));
                                            string input =
                                                $"{GetName(member)}_{GetName(memberOfZ_Data)}_{GetName(memberOfPosition02)}";
                                            switch (GetName(memberOfPosition02))
                                            {
                                                case "Origin": //获取原地址变量
                                                case "Destination": //获取目的地变量
                                                    SetSacdaTag(offset, input,"UINT");
                                                    break;
                                                case "PalletID":
                                                    //获取条码的字符长度
                                                    foreach (XmlNode item in GetSection(memberOfPosition02))
                                                    {
                                                        string barcodeType = item.Attributes.GetNamedItem("Datatype").Value;
                                                        SetSacdaTag(offset, GetBarcodeLength(barcodeType), input,"STRING");
                                                    }
                                                    break;
                                                //托盘外检拒绝信息 
                                                case "RejectCode":
                                                    foreach (XmlNode item in GetSection(memberOfPosition02))
                                                    {
                                                        offset = CalOffset(GetOffset(member) + GetOffset(memberOfZ_Data) +
                                                                           GetOffset(memberOfPosition02) + GetOffset(item));
                                                        string reject =
                                                            $"{GetName(member)}_{GetName(memberOfZ_Data)}_{GetName(item)}";
                                                        SetSacdaTag(offset, reject,"X");
                                                    }
                                                    break;
                                            }
                                        }
                                        break;
                                }
                            }
                            break;
                        case "Z_Status"://输送机状态
                            foreach (XmlNode memberOfZ_Status in GetSection(member))
                            {
                                offset = CalOffset(GetOffset(member) + GetOffset(memberOfZ_Status));
                                switch (GetName(memberOfZ_Status))
                                {
                                    case "Fault": //报警
                                    case "Occupied01": //占位1状态
                                    case "AutoModeEnabled"://自动模式
                                    case "JogSelected"://点动模式
                                        SetSacdaTag(offset, GetName(memberOfZ_Status),"X");
                                        break;
                                    case "Occupied02": //占位2状态
                                        // 判断InstanceOfName实例决定是否需要占位2数据
                                        if (!InstanceOfName.Contains("_2P") & !InstanceOfName.Contains("_PStackUni_"))
                                            break;
                                        SetSacdaTag(offset, GetName(memberOfZ_Status),"X");
                                        break;
                                }
                            }
                            break;
                    }
                    
                    //获取诊断变量和报警文本
                    foreach (string[] supervision in _supervisions.Where(supervision => supervision[0].Contains(GetName(member))))
                    {
                        offset = CalOffset(GetOffset(member));
                        SetSacdaTag(offset, GetName(member), supervision[1], supervision[2], supervision[3]);
                    }
                }
            }
        }

        /// <summary>
        /// 获取命名空间
        /// </summary>
        /// <param name="xmlDocument"></param>
        /// <param name="namespaceURI1">Sections命名空间</param>
        /// <param name="namespaceURI2">BlockInstSupervisionGroups命名空间</param>
        /// <param name="namespaceURI3">BlockTypeSupervisions命名空间</param>
        /// <returns>xmlns="http://www.siemens.com/automation/Openness/SW/Interface/v5"
        /// xmlns="http://www.siemens.com/automation/Openness/SW/BlockInstanceSupervisions/v3"</returns>
        private static XmlNamespaceManager GetXmlns(XmlDocument xmlDocument, string namespaceURI1, string namespaceURI2, string namespaceURI3)
        {
            XmlNamespaceManager xmlns = new XmlNamespaceManager(xmlDocument.NameTable);
            xmlns.AddNamespace("x", namespaceURI1);
            xmlns.AddNamespace("y", namespaceURI2);
            xmlns.AddNamespace("z", namespaceURI3);
            return xmlns;
        }

        /// <summary>
        /// 获取偏移量节点
        /// </summary>
        /// <param name="xmlNode"></param>
        /// <returns>321</returns>
        private int GetOffset(XmlNode xmlNode)
        {
            XmlNode integerAttribute = xmlNode.SelectSingleNode("./x:AttributeList", _xmlns)?.FirstChild.FirstChild;
            //获取偏移量
            if (integerAttribute != null)
            {
                int offset = Convert.ToInt16(integerAttribute.Value);
                return offset;
            }

            return 0;
        }

        /// <summary>
        /// 获取监控节点
        /// </summary>
        /// <param name="xmlNode"></param>
        /// <returns></returns>
        private XmlNode GetSupervision(XmlNode xmlNode)
        {
            XmlNode stateStruct = xmlNode.SelectSingleNode("./y:StateStruct", _xmlns);
            return stateStruct;
        }

        /// <summary>
        /// 获取监控类别节点
        /// </summary>
        /// <param name="xmlNode"></param>
        /// <returns></returns>
        private string GetSupervisionType(XmlNode xmlNode)
        {
            string str = xmlNode.SelectSingleNode("./y:BlockTypeSupervisionNumber", _xmlns)?.FirstChild.Value;
            return str;
        }

        /// <summary>
        /// 获取FB的监控定义
        /// </summary>
        /// <param name="xmlNode"></param>
        /// <returns></returns>
        private XmlNode GetSupervisedOperand(XmlNode xmlNode)
        {
            XmlNode supervisedOperand = xmlNode.SelectSingleNode("./z:SupervisedOperand", _xmlns);
            return supervisedOperand;
        }

        /// <summary>
        /// 获取FB的监控的文本定义
        /// </summary>
        /// <param name="xmlNode"></param>
        /// <param name="lang"></param>
        /// <returns></returns>
        private string GetMultiLanguageText(XmlNode xmlNode, string lang)
        {
            XmlNode specificField     = xmlNode.SelectSingleNode("./z:SpecificField", _xmlns);
            XmlNode specificFieldText = specificField?.SelectSingleNode("./z:SpecificFieldText", _xmlns);
            if (specificFieldText != null)
            {
                XmlNode multiLanguageText = specificFieldText.SelectSingleNode($"./z:MultiLanguageText[@Lang=\"{lang}\"]", _xmlns);
                //如果报警文本没有配置中文/英文，修复报错。
                string str = "";

                if (multiLanguageText == null)
                {
                    return str;
                }
                
                str = multiLanguageText.FirstChild.Value;
                return str;
            }

            return null;
        }

        /// <summary>
        /// 计算偏移量
        /// </summary>
        /// <param name="getOffset"></param>
        /// <returns>40.1</returns>
        private static string CalOffset(int getOffset)
        {
            //计算偏移量
            string consult = (getOffset / 8).ToString(); //商，"40"
            string rem = (getOffset % 8).ToString(); //余数，"1"
            string offset = consult + "." + rem; //偏移量结果：40.1

            return offset;
        }

        /// <summary>
        /// 获取Section
        /// </summary>
        /// <param name="xmlNode"></param>
        /// <returns></returns>
        private XmlNode GetSection(XmlNode xmlNode)
        {
            XmlNode section = xmlNode.SelectSingleNode("./x:Sections", _xmlns)?.SelectSingleNode("./x:Section", _xmlns);
            return section;
        }

        /// <summary>
        /// 获取节点的元素为Name的值
        /// </summary>
        /// <param name="xmlNode"></param>
        /// <returns>Name="Q_M01_Forward"</returns>
        private static string GetName(XmlNode xmlNode)
        {
            if (xmlNode.Attributes != null)
            {
                string name = xmlNode.Attributes.GetNamedItem("Name").Value;
                return name;
            }

            return null;
        }

        /// <summary>
        /// 获取条码长度
        /// </summary>
        /// <param name="input"></param>
        /// <returns>20</returns>
        private static string GetBarcodeLength(string input)
        {

            //string input = "Array[1..20] of Char";
            // 查找 "。。" 的位置
            if (input != null)
            {
                int rangeIndex = input.IndexOf("..", StringComparison.Ordinal);

                if (rangeIndex != -1)
                {
                    // 向前搜索数字的边界
                    int startIndex = rangeIndex - 1;
                    while (startIndex >= 0 && char.IsDigit(input[startIndex]))
                    {
                        startIndex--;
                    }
                    startIndex++;

                    // 向后搜索数字的边界
                    int endIndex = rangeIndex + 2;
                    while (endIndex < input.Length && char.IsDigit(input[endIndex]))
                    {
                        endIndex++;
                    }
                    endIndex--;

                    // 提取数字
                    string beforeRange = input.Substring(startIndex, rangeIndex - startIndex);
                    string afterRange  = input.Substring(rangeIndex + 2, endIndex - rangeIndex - 1);

                    int beforeNumber = int.Parse(beforeRange);
                    int afterNumber  = int.Parse(afterRange);

                    // 计算差值
                    string length = (afterNumber - beforeNumber + 1).ToString();

                    return length;
                }
                
                return string.Empty;
            }

            return null;
        }

        /// <summary>
        /// 设置SCADA标签值-Bool
        /// </summary>
        /// <param name="offset">偏移量</param>
        /// <param name="input">输入参数：拼接的变量名称</param>
        /// <param name="dataType">数据类型</param>
        /// <returns>"DB304,X40.0;P1005_Q_M01_Forward","0"</returns>
        private void SetSacdaTag(string offset, string input, string dataType)
        {
            // BOOL：DB编号,X偏移量;标签
            // INT：DB编号,INT偏移量;标签
            // UINT：DB编号,UINT偏移量;标签
            string str = $"\"{ProgrammingLanguage}{Number},{dataType}{offset};{InstanceName}_{input}\",\"0\"";
            StreamWriter.WriteLine(str);
            Console.WriteLine(str);
        }

        /// <summary>
        /// 设置SCADA标签值-字符串
        /// </summary>
        /// <param name="offset">偏移量</param>
        /// <param name="barcodeLength">条码的长度</param>
        /// <param name="input">输入参数：拼接的变量名称</param>
        /// <param name="dataType">数据类型</param>
        /// <returns>"DB304,STRING178.20;P1005_Z_Data_Position01_PalletID","0"</returns>
        private void SetSacdaTag(string offset, string barcodeLength, string input, string dataType)
        {
            // String：DB编号,STRING偏移量.字符长度;标签
            string str =
                $"\"{ProgrammingLanguage}{Number},{dataType}{offset.Replace(".0", "")}.{barcodeLength};{InstanceName}_{input}\",\"0\"";
            StreamWriter.WriteLine(str);
            Console.WriteLine(str);
        }

        /// <summary>
        /// 设置SCADA标签值-监控报警
        /// </summary>
        /// <param name="offset">偏移量</param>
        /// <param name="input">输入参数：拼接的变量名称</param>
        /// <param name="msgType">报警类型</param>
        /// <param name="textUs">报警文本-英文</param>
        /// <param name="textZh">报警文本-中文</param>
        /// <returns>"DB304,X122.1;P1005_Y_DV_JogSelected","1","zh-CN","en-US"</returns>
        private void SetSacdaTag(string offset, string input, string msgType, string textUs, string textZh)
        {
            // BOOL：DB编号,X偏移量;标签
            string str =
                $"\"{ProgrammingLanguage}{Number},X{offset};{InstanceName}_{input}\",\"{msgType}\",\"{textUs}\",\"{textZh}\"";
            StreamWriter.WriteLine(str);
            Console.WriteLine(str);
        }
    }
}