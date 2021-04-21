using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.IO;
using iTextSharp.text;
using 시군별;

namespace PdfToText
{
    class Program
    {
        static void Main(string[] args)
        {

            //string content = ExtractTextFromPdf("C:/Users/impacsys/Desktop/시군별 자동화/2021_강원도_세출.pdf");
            //Console.WriteLine(content);


            /*
            ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
            string extractedText = PdfTextExtractor.GetTextFromPage(pdfTest, 1, strategy);
            byte[] utf8Bytes = Encoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(extractedText));
            string pageText = Encoding.UTF8.GetString(utf8Bytes);
            Console.WriteLine(pageText);
            */

            //string x=PdfTextExtractor.GetTextFromPage(pdfTest, 2);
            //bool x_bool = x.Contains("세 출 예 산 사 업 명 세 서");



            /*
            StringBuilder result = new StringBuilder();
            for (int i = 1; i <= pdfTest.NumberOfPages; i++)
            {
                result.Append(PdfTextExtractor.GetTextFromPage(pdfTest, i, new iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy())).Append("\r\n\r\n");
                Console.WriteLine("---" + i + "---");
            }
            */


            //bool x_bool = result.ToString().Contains("도 청 홈 페 이 지 컨 설 팅 추 진");
            //int index = result;

            //result.Append(PdfTextExtractor.GetTextFromPage(pdfTest, 1, new iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy())).Append("\r\n\r\n");

            
            

            PdfReader pdfTest = new PdfReader("C:/Users/impacsys/Desktop/시군별 자동화 테스트/시군별 자동화/2021_sechul_0100.pdf");
            Console.WriteLine("Last_Page: " + pdfTest.NumberOfPages + "\n");

            ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
            string pdf_Text="";
            for (int i = 1; i <= pdfTest.NumberOfPages; i++)
            {
                pdf_Text = PdfTextExtractor.GetTextFromPage(pdfTest, i, strategy);
                //Console.WriteLine(pdf_Text);
                Console.WriteLine("---" + i + "---\n");
            }
            

            Console.WriteLine("-----------------------------------------------------------------------------------------------");

            string target = "홈페이지";
            
            for(int i=1; ; i=i+2)
            {
                target=target.Insert( i , " ");

                if (i == target.Length-1)
                {
                    break;
                }
            }
            Console.WriteLine("키워드: "+target);
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            
            
            bool target_bool = pdf_Text.Contains(target);
            int target_cnt = -1;
            if (target_bool)
            {


                int target_index = pdf_Text.IndexOf(target, pdf_Text.IndexOf(target));
                
                for (int j = 0; ; j++)
                {
                    target_cnt++;
                    if (j == 0)
                        target_index = pdf_Text.IndexOf(target, target_index);
                    else
                        target_index = pdf_Text.IndexOf(target, target_index + 1);

                    if (target_index == -1)
                        break;
                    //Console.WriteLine("index: "+target_index);
                    //Console.WriteLine(pdf_Text.IndexOf(target)+j);
                    //if (target_bool)
                    //{
                        string temp = "";
                        Console.WriteLine("포함\n");
                        Console.WriteLine(target_index);

                        for (int i = target_index; ; i--)
                        {
                            if (pdf_Text[i] != '\n')
                            {
                                temp += pdf_Text[i];
                            }

                            if (pdf_Text[i] == '\n')
                                break;

                        }

                        for (int i = temp.Length - 1; i > 0; i--)
                            Console.Write(temp[i]);

                        for (int i = target_index; ; i++)
                        {
                            int cnt = 0;
                            char x = pdf_Text[i];
                            if (x == ',')
                            {
                                for (int k = i; ; k++)
                                {
                                    cnt++;
                                    x = pdf_Text[k];
                                    Console.Write(x);
                                    if (cnt == 7)
                                    {
                                        //Console.Write(".");
                                        break;
                                    }

                                }
                                break;
                            }

                            Console.Write(x);


                            if (x == '\n')
                                break;
                        }

                        Console.WriteLine();

                    //}
                    /*
                    else
                    {
                        Console.WriteLine("미포함\n");
                        break;
                    }
                    */
                }
            }

            Console.WriteLine("target_cnt: " + target_cnt);

            //string testt = PdfTextExtractor.GetTextFromPage(pdfTest, 270, strategy);
            //Console.WriteLine(testt);
            pdfTest.Close();
        }



        static string ExtractTextFromPdf(string pdfFile)
        {
            StringBuilder result = new StringBuilder();
            using (Stream newpdfStream = new FileStream(pdfFile, FileMode.Open, FileAccess.Read))
            {
                PdfReader pdfReader = new PdfReader(newpdfStream);

                for (int i = 1; i <= pdfReader.NumberOfPages; i++)
                {
                    result.Append(PdfTextExtractor.GetTextFromPage(pdfReader, i, new iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy())).Append("\r\n\r\n");
                }
            }

            return result.ToString();
        }
    }
}