# 시군별 세출예산 정리 자동화 프로그램

## 프로그램 설명
> C# Windows Forms 앱(.NetFrameWork)으로 작성하여 구현한 데이터 파싱 프로그램  
> PDF 서드파티 라이브러리 iTextSharp, 액셀 라이브러리 Microsoft.Office.Interop.Excel 사용  
> iTextSharp 라이브러리: PDF 를 C# 환경에서 다룰 수 있게 해주는 도구
> Microsoft Excel 15.0 Object 라이브러리: EXCEL을 다룰 수 있게 해주는 도구
>> 1. 간단하게 PDF의 대부분의 문자열을 읽어들인 후, 규칙성을 찾아줌
>> 2. 규칙성을 찾아서 정돈된 문자열을 [키워드/시군청/부서명/예산명/예산액] 으로 각각 나누어서 엑셀의 각 셀에 입력
>> 3. Windows Forms 앱(.NET Framework)으로 사용자 인터페이스와 애플리케이션 작성

## 파싱할 PDF 파일 형식
> http://www.provin.gangwon.kr/gw/portal/sub06_06_07_19 해당 URL의 사업명세서 PDF 파일 형식 참조
------------
![화면 캡처 2021-04-21 191521](https://user-images.githubusercontent.com/70702088/115537622-f6f64300-a2d5-11eb-8825-50529b2541f2.png)



