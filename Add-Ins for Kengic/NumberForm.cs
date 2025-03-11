using System;
using System.Windows.Forms;

namespace Kengic
{
    public partial class NumberForm : Form
    {
        public NumberForm()
        {
            InitializeComponent();
        }

        public int StartingNumber => (int)numericUpDownStarting.Value;//定义变量，起始号
        public int Increment => (int)numericUpDownIncrement.Value;//定义变量，递增

        private Timer _timer;//定义定时器

        //加载窗体时触发事件
        private void NumberForm_Load(object sender, EventArgs e)
        {
            _timer = new Timer();
            _timer.Interval = 100;//定时时间
            _timer.Tick += timer_Tick;//定义定时器结束触发事件
            _timer.Start();
        }

        //失去焦点时触发事件
        private void NumberForm_Deactivate(object sender, EventArgs e)
        {
            _timer = new Timer();
            _timer.Interval = 100;//定时时间
            _timer.Tick += timer_Tick;//定义定时器结束触发事件
            _timer.Start();
        }

        //定时器结束触发事件
        private void timer_Tick(object sender, EventArgs e)
        {
            Activate();
            Focus();
            _timer.Stop();
        }

        //输入数字时触发事件
        private void numericUpDownStarting_Enter(object sender, EventArgs e)
        {
            numericUpDownStarting.Select(0, numericUpDownStarting.Text.Length);
        }

        private void numericUpDownIncrement_Enter(object sender, EventArgs e)
        {
            numericUpDownIncrement.Select(0, numericUpDownIncrement.Text.Length);
        }
    }
}
