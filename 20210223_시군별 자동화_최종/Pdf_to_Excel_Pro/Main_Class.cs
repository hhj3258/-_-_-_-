using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Pdf_to_Excel_Pro
{
    class TestC
    {
        //   0   1    2     3      4      5      6
        // 공백/넘버/구분/시군청/부서명/사업명/예산액

        public int num = 0;
        public string key = "";
        public string sigun = "";
        public string buseo = "";
        public string saup = "";
        public string pay = "";

        public void print()
        {
            Console.Write(num + " / ");
            Console.Write(key + " / ");
            Console.Write(sigun + " / ");
            Console.Write(buseo + " / ");
            Console.Write(saup + " / ");
            Console.WriteLine(pay);
        }

    }

    class TargetList
    {
        public List<int> page_text_length = new List<int>();
        public List<int> target_index_list = new List<int>();
    }

    class Main_Class
    {
        public List<TestC> MainF(string[] target_name, ref TargetList TL)
        {
            TL = new TargetList();

            PdfReader pdfTest = null;
            List<TestC> return_temp = null;
            try
            {
                string path = Form1.myForm.textBox2.Text;
                pdfTest = new PdfReader(path);
            }
            catch
            {
                Console.WriteLine("파일경로 오류");
                return return_temp;
            }

            ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();

            string pdf_Text = "";
            Form1.myForm.label5.Text = pdfTest.NumberOfPages + "p";


            int start_p = 1;
            int end_p = pdfTest.NumberOfPages;
            if (!Form1.myForm.textBox4.Text.Equals("") && !Form1.myForm.textBox5.Text.Equals(""))
            {
                start_p = Convert.ToInt32(Form1.myForm.textBox4.Text);
                end_p = Convert.ToInt32(Form1.myForm.textBox5.Text);
            }


            for (int i = start_p; i <= end_p; i++)//pdfTest.NumberOfPages // 665 667
            {
                string pdf_text_temp = PdfTextExtractor.GetTextFromPage(pdfTest, i);

                int temp_length = 100;
                for (int j = 0; j < temp_length; j++)
                {

                    if (pdf_text_temp.Length < 5)
                        break;



                    if ((j < pdf_text_temp.Length) && pdf_text_temp[j] != ' ')
                    {
                        if (((j + 1) < pdf_text_temp.Length) && pdf_text_temp[j + 1] != ' ')
                        {
                            temp_length = pdf_text_temp.Length - 1;
                            pdf_text_temp = pdf_text_temp.Insert(j + 1, " ");
                            j++;
                        }
                    }

                }

                pdf_Text += pdf_text_temp + "\n";
                TL.page_text_length.Add(pdf_Text.Length - 1);    //리스트 0번째에 pdf 1p의 텍스트 길이가 담김(반복)


                Form1.myForm.label1.Text = i + "p";
                Application.DoEvents();
            }


            List<TestC> testc_list = new List<TestC>();


            for (int h = 0; h < target_name.Length; h++)
            {
                string target = target_name[h];
                string post_target = target;
                for (int i = 1; ; i = i + 2)    //해당 타겟에 공백 삽입
                {
                    target = target.Insert(i, " ");

                    if (i == target.Length - 1)
                    {
                        break;
                    }
                }


                bool target_bool = pdf_Text.Contains(target);
                string buseo = "부 서 :";
                string target_result = "";
                Pdf_Class PdfClass = new Pdf_Class();
                List<string> buseo_list = new List<string>();

                int target_cnt = -1;    //타겟이 몇개 존재하는지 카운트
                if (target_bool)    //타겟이 존재할 때만 통과
                {
                    int target_index = 0;
                    int buseo_index = -1;
                    int temp1 = -1;
                    for (int j = 0; ; j++)
                    {
                        string buseo_temp = "";
                        target_cnt++;

                        if (j == 0)
                        {
                            target_index = pdf_Text.IndexOf(target);
                            buseo_index = pdf_Text.IndexOf(buseo);
                            if (buseo_index > target_index) buseo_index = -1;

                            if (buseo_index < target_index)
                            {
                                while (temp1 < target_index)
                                {
                                    temp1 = pdf_Text.IndexOf(buseo, temp1 + 1);
                                    if (temp1 > target_index || temp1 == -1)
                                        break;
                                }
                                buseo_index = temp1;
                            }
                        }
                        else
                        {
                            target_index = pdf_Text.IndexOf(target, target_index + 1);

                            while (temp1 < target_index)
                            {
                                buseo_index = temp1;
                                temp1 = pdf_Text.IndexOf(buseo, temp1 + 1);
                                if (temp1 > target_index || temp1 == -1)
                                    break;
                            }
                        }

                        TL.target_index_list.Add(target_index);


                        for (int i = buseo_index; buseo_index != -1; i++)
                        {
                            if (pdf_Text[i] == '\n')
                                break;
                            if (i > buseo_index + 5)
                                buseo_temp += pdf_Text[i];
                        }

                        buseo_list.Add(buseo_temp);



                        if (target_index == -1)     //더이상 타겟이 존재하지 않으면 종료
                            break;

                        target_result = PdfClass.PdfTextPrint(target_index, pdf_Text, target_result);

                    }//outter for
                }//outter if


                for (int i = 0; i < buseo_list.Count; i++)
                {
                    buseo_list[i] = buseo_list[i].Replace(" ", "");
                }

                int result_length = 0;
                for (int i = 0; i < target_result.Length; i++)
                {
                    if (target_result[i] == '\n')
                    {
                        result_length++;
                    }
                }

                //   0   1    2     3      4       5     6
                // 공백/넘버/구분/시군청/부서명/사업명/예산액


                string[,] result = new string[result_length, 7];//result_length
                int result_cnt = 0;
                int empty_cnt = 0;
                string[,] final_result = new string[result_length - empty_cnt, 7];

                buseo_list.Add("temp");


                final_result = CleanUp(result, final_result, result_length, post_target,
                    buseo_list, target_result, result_cnt, ref empty_cnt, ref testc_list);


            }

            pdfTest.Close();
            return testc_list;
        }


        int form_cnt = 0;
        private string[,] CleanUp(string[,] result, string[,] final_result, int result_length, string post_target,
            List<string> buseo_list, string target_result, int result_cnt, ref int empty_cnt, ref List<TestC> testc_list)
        {
            string[] result_temp = new string[result.Length];
            int[] split_cnt = new int[result.Length];
            for (int i = 0; i < result_length; i++)
            {

                result[i, 1] = (i + 1).ToString();
                result[i, 2] = post_target;
                result[i, 3] = "테스트";
                result[i, 4] = buseo_list[i];
                result[i, 5] = "";
                result_temp[i] = "";


                int jcnt = result_cnt;
                for (int j = result_cnt; target_result[j] != '\n'; j++)
                {
                    result_cnt++;

                    if (target_result[j] == ' ' || target_result[j] == ',')
                        continue;

                    result_temp[i] += target_result[j];
                }


                bool split_bool = false;
                for (int j = result_temp[i].Length - 1; j >= 0; j--)
                {
                    if (result_temp[i][j] > 47 && result_temp[i][j] < 58 && split_bool == false)
                        split_cnt[i]++;
                    else
                        split_bool = true;
                }


                int for_cnt = 0;
                for (int j = 0; j < result_temp[i].Length; j++)
                {
                    if (split_cnt[i] == 0)
                    {
                        result[i, 5] = result_temp[i];
                        break;
                    }

                    if (for_cnt < (result_temp[i].Length - split_cnt[i]))
                    {
                        result[i, 5] += result_temp[i][j];
                    }
                    else
                        result[i, 6] += result_temp[i][j];

                    for_cnt++;
                }

                if (buseo_list[i].Equals(""))
                {
                    empty_cnt++;
                    result[i, 6] = "";
                    result[i, 5] = "";
                }

                if (result[i, 6] == null)
                {
                    empty_cnt++;

                    result[i, 6] = "";
                    result[i, 5] = "";
                    i--;
                    result_cnt--;
                }

                result_cnt++;
            }

            int temp_cnt = 0;
            for (int i = 0; i < (result_length - empty_cnt); i++)
            {

                while (result[temp_cnt, 5].Equals(""))
                {
                    temp_cnt++;
                }

                final_result[i, 1] = (i + 1).ToString();

                Form1.myForm.DLabel[form_cnt].Text = final_result[i, 1];

                final_result[i, 2] = post_target;
                final_result[i, 3] = "테스트";
                final_result[i, 4] = buseo_list[temp_cnt];
                final_result[i, 5] = result[temp_cnt, 5];
                final_result[i, 6] = result[temp_cnt, 6];


                TestC testc2 = new TestC();
                testc2.key = post_target;
                testc2.sigun = Form1.myForm.comboBox1.SelectedItem.ToString();
                testc2.buseo = buseo_list[temp_cnt];
                testc2.saup = result[temp_cnt, 5];
                testc2.pay = result[temp_cnt, 6];


                testc_list.Add(testc2);

                temp_cnt++;
            }
            form_cnt++;
            return final_result;
        }


    }
}
