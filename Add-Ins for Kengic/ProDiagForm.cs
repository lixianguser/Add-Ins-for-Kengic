using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Kengic
{
    public partial class ProDiagForm : Form
    {
        private List<ProDiagFB> _FB;

        public string Block;

        public ProDiagForm()
        {
            InitializeComponent();
            // 使窗体置顶显示
            TopMost = true;
            // 使窗体居中显示
            StartPosition = FormStartPosition.CenterScreen;
        }
        
        // private void ProDiagForm_Load(object sender, EventArgs e)
        // {
        //     Focus();  // 窗体加载时获取焦点
        // }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            Activate();  // 窗体显示时激活并聚焦
        }

        public void SetDevices(List<ProDiagFB> devices)
        {
            this._FB = devices;
            BindDevicesToListBox();  // 在设置完 _FB 后再绑定
        }

        private void BindDevicesToListBox()
        {
            // 绑定设备列表到ListBox控件
            listBoxDevices.DataSource = null;  // 先清空数据源
            listBoxDevices.DataSource = _FB;
        }

        private void buttonSelect_Click(object sender, EventArgs e)
        {
            // 获取选中的设备并显示信息

            if (listBoxDevices.SelectedItem is ProDiagFB selectedProDiagInfo)
            {
                Block = selectedProDiagInfo.Name;
            }
        }
    }

    public class ProDiagFB
    {
        public string Name   { get; set; }

        public override string ToString()
        {
            return Name;
        }
    }
}