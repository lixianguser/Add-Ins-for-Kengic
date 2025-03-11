using Siemens.Engineering.HW;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Kengic
{
    public partial class AlarmForm : Form
    {
        private List<DeviceInfo> devices;

        public Device device;

        public AlarmForm()
        {
            InitializeComponent();
            // 使窗体置顶显示
            TopMost = true;
            // 使窗体居中显示
            StartPosition = FormStartPosition.CenterScreen;
        }
        
        // private void MainForm_Load(object sender, EventArgs e)
        // {
        //     Focus();  // 窗体加载时获取焦点
        // }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            Activate();  // 窗体显示时激活并聚焦
        }

        public void SetDevices(List<DeviceInfo> devices)
        {
            this.devices = devices;
            BindDevicesToListBox();  // 在设置完 devices 后再绑定
        }

        private void BindDevicesToListBox()
        {
            // 绑定设备列表到ListBox控件
            listBoxDevices.DataSource = null;  // 先清空数据源
            listBoxDevices.DataSource = devices;
        }

        private void buttonSelect_Click(object sender, EventArgs e)
        {
            // 获取选中的设备并显示信息
            if (listBoxDevices.SelectedItem is DeviceInfo selectedDeviceInfo) device = selectedDeviceInfo.Device;
        }
    }

    public class DeviceInfo
    {
        public string Name   { get; set; }
        public Device Device { get; set; }

        public override string ToString()
        {
            return Name;
        }
    }
}