
using HwpObjectLib;
using Microsoft.Win32;
using System;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using static System.Runtime.InteropServices.JavaScript.JSType;

[assembly: SupportedOSPlatform("windows")]
internal class Program
{
    private static void Main(string[] args)
    {
        Console.WriteLine("Hello, World!");
        try
        {/* 한줄주석 ctrl + k + c  , ctrl + k + u
            string Str_M_Path = @"C:\hwp_보안모듈_Automation\보안모듈(Automation)\FilePathCheckerModuleExample.dll";
            RegistryKey key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\HNC\\HwpAutomation\\Modules");
            if (key != null) {
                if (key.GetValue("FilePathCheckerModuleExample") == null)
                {
                    key.SetValue("FilePathCheckerModuleExample",Str_M_Path);
                    Console.WriteLine("아하하핳");
                }
            }
            else
            {
                Console.WriteLine($"{key} Modules을 찾을 수 없습니다.");
            }
            */
            //엑셀경로와 범위 설정
            string Ex_filePath = @"D:\C샵hwp연습용\2023_Csarp_hwp\2023_Csarp_vs\ch1\ConsoleApp1\ConsoleApp1\TempData\수상자목록.xlsx";
            int sheetIndex = 0;

            // ExcelReader 클래스의 ReadExcel 메서드를 사용하여 데이터를 읽어옵니다.
            DataTable dt = ExcelReader.ReadExcel(Ex_filePath, sheetIndex);
            //var dt
            //var rows = dt.AsEnumerable().Select(r => string.Join("\t", r.ItemArray));
            //var result = string.Join(Environment.NewLine, rows);

            // 출력 datatable
            string result = (from row in dt.AsEnumerable()
                             select string.Join("\t", row.ItemArray)
                            ).Aggregate((current, next) => current + Environment.NewLine + next);

            Console.WriteLine(result);

            string processName = "Hwp";
            string FilePath = @"D:\C샵hwp연습용\2023_Csarp_hwp\2023_Csarp_vs\ch1\ConsoleApp1\ConsoleApp1\TempData\h25_certificate_of_award.hwp";
            string FilePath2 = @"D:\C샵hwp연습용\2023_Csarp_hwp\2023_Csarp_vs\ch1\ConsoleApp1\ConsoleApp1\TempData\h25_certificate_of_award완료.hwp";
            Console.WriteLine(FilePath);

            //레지스트리에 경로 설정해주기
            string Str_M_Path = @"C:\hwp_보안모듈_Automation\보안모듈(Automation)\FilePathCheckerModuleExample.dll";
            string HNCRootkey = @"HKEY_Current_User\SOFTWARE\HNC\HwpAutomation\Modules";

            Registry.SetValue(HNCRootkey, "FilePathCheckerModuleExample", Str_M_Path);

            // HwpObject 객체 생성
            IHwpObject hwp = new HwpObject();
            //보안모듈적용
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample");
            //화면보이기
            hwp.XHwpWindows.Active_XHwpWindow.Visible = true;
            //화면열기
            hwp.Open(FilePath, "", "");
            //전체화면
            hwp.Run("FrameFullScreen");
            //필드 리스트로 출력하기 2가 갯수 
            string Str_Fields = hwp.GetFieldList(0,"");
            Console.WriteLine("필드리스트내용 : " + Str_Fields);
            string[] A_str_field = Str_Fields.Split('\x02');
            Console.WriteLine("출력 A_str_field : " + string.Join(",", A_str_field));
            //List<string> HFieldList = new List<string>(str_filed);
            //foreach (string a in HFieldList)
            //{/
            //    Console.WriteLine(a);
            //}

            //문서 현재쪽 전체선택 후 복사해서 마지막쪽으로 가서 붙여넣기
            //hwp.MovePos(0, 0, 0);
            hwp.Run("SelectAll");
            hwp.Run("Copy");
            hwp.MovePos(3,0,0);
            foreach (int i in Enumerable.Range(0,dt.Rows.Count-1))
            {
                //Console.WriteLine(dt.Rows[i+1][0].ToString());
                
                hwp.Run("Paste");
                hwp.MovePos(3, 0, 0);
            }
            //dt의 내용을 밑에 for문에서 하지말고 linq로 교체하기
            var outdt = dt.Clone();
            outdt = (from row in dt.AsEnumerable()
                     where !row.IsNull("시상일")
                     let parsedDate = DateTime.ParseExact(row["시상일"].ToString(), "yyyy년MM월dd일", null)
                     let date1 = parsedDate.Date.ToString("yyyy. MM. dd")
                     select outdt.Rows.Add(
                         row.ItemArray.Take(4).Concat(new object[]{date1}).Concat(row.ItemArray.Skip(5).Take(1)).ToArray())
                        ).CopyToDataTable();

            result = (from row in outdt.AsEnumerable()
                             select string.Join("\t", row.ItemArray)
                            ).Aggregate((current, next) => current + Environment.NewLine + next);

            Console.WriteLine(result);

            // 필드에 내용 바꿔서 넣어주기
            foreach (int i in Enumerable.Range(0, outdt.Rows.Count))
            {
                foreach (string field in A_str_field)
                {
                    //Console.WriteLine(field);
                    //if (field.Equals("시상일"))
                    //{
                        //Console.WriteLine(dt.Rows[i][field].ToString());
                        //DateTime parsedDate = DateTime.ParseExact(dt.Rows[i][field].ToString(), "yyyy년MM월dd일", null);
                        //dt.Rows[i][field] = parsedDate.Date.ToString("yyyy년 MM월 dd일");
                        //Console.WriteLine(dt.Rows[i][field].ToString());
                    //}
                    hwp.PutFieldText($"{field}{{{{{i}}}}}", outdt.Rows[i][field].ToString());

                }
            }


            //hwp.PutFieldText("주최자" + "{{0}}", dt.Rows[0]["주최자"].ToString());
            //hwp.PutFieldText("주최자", "카카오");

            hwp.SaveAs(FilePath2,"","");
            // Hwp 종료
            //hwp.Quit();
            //killProcess(processName);

            // COM 객체 해제
            Marshal.ReleaseComObject(hwp);
            Console.WriteLine("객체해제");
            //Marshal.FinalReleaseComObject(hwp);
            //hwp = null;


       }
        catch (Exception ex) {
            Console.WriteLine("예외 발생");
            Console.WriteLine(ex.Message);
        }

    }

    private static void killProcess(string processName)
    {
        //프로세스 배열 가져옴
        //Process[] processes = Process.GetProcessesByName(processName);

        // 첫 번째로 발견된 프로세스를 가져옵니다.
        Process process = Process.GetProcessesByName(processName).FirstOrDefault();

        // 프로세스가 존재하는 경우 종료합니다.
        if (process != null)
        {
            process.Kill();
            Console.WriteLine($"{processName} 프로세스가 종료되었습니다.");
        }
        else
        {
            Console.WriteLine($"{processName} 프로세스가 찾을 수 없습니다.");
        }
    }






}