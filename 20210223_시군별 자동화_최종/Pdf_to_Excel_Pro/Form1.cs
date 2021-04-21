using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace Pdf_to_Excel_Pro
{
    public partial class Form1 : Form
    {
        int keyword_cnt = 0;

        public static Form1 myForm;
        public Form1()
        {
            InitializeComponent();
            myForm = this;
        }



        private void button1_Click(object sender, EventArgs e)
        {


            if (comboBox1.SelectedItem == null)
            {
                MessageBox.Show("시군명을 선택해주세요.", "Error");
                return;
            }
            else if (keyword_add == false)
            {
                MessageBox.Show("키워드를 추가해주세요.", "Error");
                return;
            }
            else if (textBox2.Text.Equals(""))
            {
                MessageBox.Show("파일을 추가해주세요.", "Error");
                return;
            }
            else if (textBox3.Text.Equals(""))
            {
                MessageBox.Show("저장 경로를 지정해주세요.", "Error");
                return;
            }


            string[] keyword_arr = new string[keyword_cnt];
            for (int i = 0; i < keyword_cnt; i++)
            {
                if (DText[i].Text.Equals(""))
                {
                    MessageBox.Show("키워드를 모두 입력해주세요.", "Error");
                    return;
                }
                keyword_arr[i] = DText[i].Text;
                //label1.Text += keyword_arr[i];
            }



            Main_Class mc = new Main_Class();



            List<TestC> result_list = new List<TestC>();
            TargetList TL1 = new TargetList();
            try
            {
                result_list = mc.MainF(keyword_arr, ref TL1);
                for (int i = 0; i < result_list.Count; i++)
                {
                    result_list[i].num = i + 1;
                    result_list[i].print();
                }
                Excel_Class excel_Ex = new Excel_Class();
                excel_Ex.Excel_Export(result_list, result_list.Count, TL1);
            }
            catch
            {
                MessageBox.Show("파일 경로 오류", "Error");
                return;
            }

        }





        List<TextBox> DText = new List<TextBox>();
        public List<Label> DLabel = new List<Label>();
        List<Button> DButton = new List<Button>();
        bool keyword_add = false;
        int y = 210;
        int x_cnt = 0;
        private void button2_Click(object sender, EventArgs e)
        {

            if (keyword_cnt % 3 == 0)
            {
                x_cnt = 0;
                y += 30;
            }


            DText.Add(new TextBox());
            DText[keyword_cnt].Location = new Point(x_cnt * 150 + 15, y);
            DText[keyword_cnt].Size = new Size(70, 10);
            DText[keyword_cnt].Name = "txtBox" + keyword_cnt.ToString();
            DText[keyword_cnt].Tag = keyword_cnt;
            this.Controls.Add(DText[keyword_cnt]);


            DLabel.Add(new Label());
            DLabel[keyword_cnt].Location = new Point(x_cnt * 150 + 75 + 15, y + 3);
            DLabel[keyword_cnt].Name = "lbl" + keyword_cnt.ToString();
            DLabel[keyword_cnt].Size = new Size(30, 20);
            DLabel[keyword_cnt].Text = "0";
            DLabel[keyword_cnt].Tag = keyword_cnt;
            this.Controls.Add(DLabel[keyword_cnt]);


            DButton.Add(new Button());
            DButton[keyword_cnt].Location = new Point(x_cnt * 150 + 120, y);
            DButton[keyword_cnt].Size = new Size(40, 23);
            DButton[keyword_cnt].Name = "btn" + keyword_cnt.ToString();
            DButton[keyword_cnt].Text = "삭제";
            DButton[keyword_cnt].Click += new EventHandler(DButton_Click);
            DButton[keyword_cnt].Tag = keyword_cnt;
            this.Controls.Add(DButton[keyword_cnt]);


            x_cnt++;
            keyword_add = true;
            keyword_cnt++;
        }

        private void DButton_Click(object sender, EventArgs e)
        {
            Button this_btn = (Button)sender;

            for (int i = 0; i < keyword_cnt; i++)
            {
                Console.WriteLine("DText[i].Tag: " + DText[i].Tag);
                if (this_btn.Tag.Equals(DText[i].Tag))
                {
                    Console.WriteLine("chk");
                    this.Controls.Remove(DText[i]);
                    this.Controls.Remove(DLabel[i]);
                    this.Controls.Remove(DButton[i]);
                    DText.RemoveAt(i);
                    DLabel.RemoveAt(i);
                    DButton.RemoveAt(i);
                    break;
                }
            }

            keyword_cnt--;

            if (keyword_cnt == 0)
                keyword_add = false;


        }


        private void button3_Click(object sender, EventArgs e)
        {
            textBox2.Clear();
            string file_path = null;
            if (textBox2.Text == null)
                openFileDialog1.InitialDirectory = "C:\\";
            else
                openFileDialog1.InitialDirectory = textBox2.Text;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                file_path = openFileDialog1.FileName;
                textBox2.Text = file_path;
            }


        }


        private void button5_Click(object sender, EventArgs e)
        {
            textBox3.Clear();
            string file_path = null;
            saveFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                file_path = saveFileDialog1.FileName;
                textBox3.Text = file_path;
            }
        }

        private void button3_MouseHover(object sender, EventArgs e)
        {
            this.toolTip1.AutoPopDelay = 15000;
            this.toolTip1.InitialDelay = 100;
            this.toolTip1.ReshowDelay = 500;
            this.toolTip1.ShowAlways = true;
            this.toolTip1.SetToolTip(this.button3, "하나의 시군에 대한 '세출예산서' \n병합본을 등록해주세요.");   //파일 찾기

        }

        private void label7_MouseHover(object sender, EventArgs e)
        {
            this.toolTip1.AutoPopDelay = 15000;
            this.toolTip1.InitialDelay = 100;
            this.toolTip1.ReshowDelay = 500;
            this.toolTip1.ShowAlways = true;
            this.toolTip1.SetToolTip(this.label7, "공란으로 두면 전체 페이지를 조사합니다.");
        }

        private void button5_MouseHover(object sender, EventArgs e)
        {
            this.toolTip1.AutoPopDelay = 15000;
            this.toolTip1.InitialDelay = 100;
            this.toolTip1.ReshowDelay = 500;
            this.toolTip1.ShowAlways = true;
            this.toolTip1.SetToolTip(this.button5, "파일명까지 입력해주세요.");
        }
    }
}
