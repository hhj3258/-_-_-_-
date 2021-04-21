using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_test
{
    class Dog
    {
        public string name;     // 개 이름
        public string breeds;   // 개 종류
        public string sex;      // 개 성별
    }

    class Program
    {
        static Excel.Application excelApp = null;
        static Excel.Workbook workBook = null;
        static Excel.Worksheet workSheet = null;

        static void Main(string[] args)
        {
            try
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);  // 바탕화면 경로
                string path = Path.Combine(desktopPath, "출력본.xlsx");                              // 엑셀 파일 저장 경로

                excelApp = new Excel.Application();                             // 엑셀 어플리케이션 생성
                workBook = excelApp.Workbooks.Add();                            // 워크북 추가
                workSheet = workBook.Worksheets.get_Item(1) as Excel.Worksheet; // 엑셀 첫번째 워크시트 가져오기

                workSheet.Cells[1, 1] = "이름";
                workSheet.Cells[1, 2] = "종류";
                workSheet.Cells[1, 3] = "성별";

                // 엑셀에 저장할 개 데이터
                List<Dog> dogs = new List<Dog>();
                Dog dd = new Dog();
                dd.name = "test";
                dd.breeds = "test2";
                dd.sex = "test3";
                dogs.Add(dd);
                dogs.Add(new Dog() { name = "곰이", breeds = "시바", sex = "남" });
                dogs.Add(new Dog() { name = "두부", breeds = "포메라니안", sex = "여" });
                dogs.Add(new Dog() { name = "뭉치", breeds = "말티즈", sex = "남" });

                for (int i = 0; i < dogs.Count; i++)
                {
                    Dog dog = dogs[i];

                    // 셀에 데이터 입력
                    workSheet.Cells[2 + i, 1] = dog.name;
                    workSheet.Cells[2 + i, 2] = dog.breeds;
                    workSheet.Cells[2 + i, 3] = dog.sex;
                }

                workSheet.Columns.AutoFit();                                    // 열 너비 자동 맞춤
                workBook.SaveAs(path, Excel.XlFileFormat.xlWorkbookDefault);    // 엑셀 파일 저장
                workBook.Close(true);
                excelApp.Quit();
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