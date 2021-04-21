using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Pdf_to_Excel_Pro
{

    class Excel_Class
    {
        static Excel.Application excelApp = null;
        static Excel.Workbook workBook = null;
        static Excel.Worksheet workSheet = null;

        public void Excel_Export(List<TestC> result_list, int result_length, TargetList TL)
        {
            try
            {


                string file_name = "";
                string desktopPath = Form1.myForm.textBox3.Text;  // 경로
                //string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string path = Path.Combine(desktopPath, file_name);                              // 엑셀 파일 저장 경로

                excelApp = new Excel.Application();                             // 엑셀 어플리케이션 생성
                workBook = excelApp.Workbooks.Add();                            // 워크북 추가
                workSheet = workBook.Worksheets.get_Item(1) as Excel.Worksheet; // 엑셀 첫번째 워크시트 가져오기

                //workSheet.Cells[x,y] -> x를 바꾸면 세로, y를 바꾸면 가로


                Form1.myForm.label1.Text = "엑셀 변환중...";

                workSheet.Cells[1, 1] = "지역No";
                workSheet.Cells[1, 2] = "구분";
                workSheet.Cells[1, 3] = "시군청";
                workSheet.Cells[1, 4] = "부서명";
                workSheet.Cells[1, 5] = "세부사업명";
                workSheet.Cells[1, 6] = "예산액";



                int a = TL.page_text_length[0];    //페이지마다 텍스트 길이가 쌓임
                int b = TL.target_index_list[0];    //타겟인덱스의 리스트

                int lbl_cnt = 0;
                int lbl_text = Convert.ToInt32(Form1.myForm.DLabel[lbl_cnt].Text);



                for (int i = 0; i < result_length; i++)
                {
                    Application.DoEvents();

                    workSheet.Cells[i + 2, 1] = "=ROW()-1";
                    workSheet.Cells[i + 2, 2] = result_list[i].key;
                    workSheet.Cells[i + 2, 3] = result_list[i].sigun;
                    workSheet.Cells[i + 2, 4] = result_list[i].buseo;
                    workSheet.Cells[i + 2, 5] = result_list[i].saup;
                    workSheet.Cells[i + 2, 6] = result_list[i].pay;

                }


                workSheet.Columns.AutoFit();                                    // 열 너비 자동 맞춤
                workBook.SaveAs(path, Excel.XlFileFormat.xlWorkbookDefault);    // 엑셀 파일 저장

                Form1.myForm.label1.Text = "완료";

                workBook.Close(true);
                excelApp.Quit();
            }
            catch
            {
                Console.WriteLine("오류");
                Form1.myForm.label1.Text = "오류";
            }
            finally
            {
                ReleaseObject(workSheet);
                ReleaseObject(workBook);
                ReleaseObject(excelApp);
            }
        }

        /// <summary>
        /// 액셀 객체 해제 메소드
        /// </summary>
        /// <param name="obj"></param>
        static void ReleaseObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);  // 액셀 객체 해제
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();   // 가비지 수집
            }
        }
    }
}
