﻿
using HwpObjectLib;
using Microsoft.Win32;
using System;
using System.Data;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Excel = Microsoft.Office.Interop.Excel;

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
            string rangeAdd = "A1:F8";


            Excel.Range range = ExcelReader.ReadExcelFile(Ex_filePath, rangeAdd);
            
            string processName = "Hwp";
            string FilePath = @"D:\C샵hwp연습용\2023_Csarp_hwp\2023_Csarp_vs\ch1\ConsoleApp1\ConsoleApp1\TempData\h25_certificate_of_award.hwp";
            //string FilePath2 = @"D:\C샵hwp연습용\2023_Csarp_hwp\2023_Csarp_vs\ch1\ConsoleApp1\ConsoleApp1\TempData\h25_certificate_of_award2.hwp";
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
            //필드 리스트로 출력하기
            string Str_Fields = hwp.GetFieldList(1,"");
            Console.WriteLine("필드리스트내용 : " + Str_Fields);
            string[] str_filed = Str_Fields.Split('\x02');
            Console.WriteLine(string.Join(",", str_filed));

            //List<string> HFieldList = new List<string>(str_filed);
            //foreach (string a in HFieldList)
            //{
            //    Console.WriteLine(a);
            //}

            //문서 현재쪽 전체선택 후 복사해서 마지막쪽으로 가서 붙여넣기
            hwp.Run("SelectAll");
            hwp.Run("Copy");
            hwp.MovePos(3,0,0);
            hwp.Run("Paste");

            //hwp.PutFieldText("주최자", "카카오");
            

            //hwp.SaveAs(FilePath2);
            // Hwp 종료
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