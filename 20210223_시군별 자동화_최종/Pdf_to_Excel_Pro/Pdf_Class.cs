namespace Pdf_to_Excel_Pro
{
    class Pdf_Class
    {

        public string PdfTextPrint(int target_index, string pdf_Text, string result_text)
        {
            string temp = "";

            for (int j = target_index; ; j--)       //해당 타겟 앞부분 출력, '\n' 문자열전까지 출력
            {
                if (pdf_Text[j] == '\n')
                    break;

                temp += pdf_Text[j];
            }

            for (int j = temp.Length - 1; j > 0; j--)   //타겟 앞부분 문자열이 거꾸로 배열에 받아졌으므로 다시 거꾸로
                result_text += temp[j];

            int i = 0;

            for (i = target_index; ; i++)   // ' , ' 이후 7문자 출력후 종료 || '\n' 을 읽어들이면 종료
            {
                int cnt = 0;
                result_text += pdf_Text[i];

                if (pdf_Text[i] == ',')       // 만약 현재 문자가 , 라면
                {
                    bool btemp = false;
                    int k = i + 1;

                    for (; cnt < 6; k++)   // cnt가 6보다 작을 때까지만, 즉 공백포함 6 문자 출력 후 종료
                    {
                        cnt++;  // 단순히 cnt + 1
                        if (pdf_Text[k + 2] == ',') //만약 k+2(공백+다음 문자)가 , 라면 cnt=0 초기화 후 다시 공백포함 6문자 출력
                        {
                            cnt = 0;
                            btemp = true;
                        }

                        result_text += pdf_Text[k];
                    }

                    if (btemp)
                    {
                        result_text += pdf_Text[k++];
                        result_text += pdf_Text[k++];

                    }

                    if (!(pdf_Text[k + 1] >= 48 && pdf_Text[k + 1] <= 57))  //다음문자가 숫자가 아니면
                    {
                        int while_k = k + 1;
                        while (true)
                        {
                            if (pdf_Text[while_k] == '\n')
                                break;
                            result_text += pdf_Text[while_k];
                            while_k++;
                        }
                    }


                    if (pdf_Text[cnt] != '\n')
                    {
                        result_text += '\n';
                    }

                    break;
                }


                if (pdf_Text[i] == '\n')  //해당 타겟의 한 줄을 모두 읽어들이면 종료
                {

                    break;
                }
            }

            string temp2 = "";
            for (int j = target_index; pdf_Text[j] != '\n'; j++)
            {
                temp2 += pdf_Text[j];
            }

            if (!(temp2[temp2.Length - 2] >= 48 && temp2[temp2.Length - 2] <= 57) && temp2[temp2.Length - 1] == ' ')
            {
                result_text = result_text.Substring(0, result_text.Length - 1);
                while (pdf_Text[i + 1] != '\n')
                {
                    result_text += pdf_Text[i + 1];
                    i++;
                }
                result_text += '\n';
            }

            return result_text;
        }
    }
}
