//using System;
//using System.IO;
using HwpObjectLib;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

internal class Program
{
    private static void Main(string[] args)
    {
        Console.WriteLine("Hello, World!");
        try
        {
            string processName = "Hwp";
            string FilePath = @"D:\C샵hwp연습용\2023_Csarp_hwp\2023_Csarp_vs\ch1\ConsoleApp1\ConsoleApp1\TempData\h25_certificate_of_award.hwp";
            //string FilePath2 = @"D:\C샵hwp연습용\2023_Csarp_hwp\2023_Csarp_vs\ch1\ConsoleApp1\ConsoleApp1\TempData\h25_certificate_of_award2.hwp";
            Console.WriteLine(FilePath);
            // HwpObject 객체 생성
            IHwpObject hwp = new HwpObject();

            hwp.XHwpWindows.Active_XHwpWindow.Visible = true;

            hwp.Open(FilePath, "", "");

            //hwp.PutFieldText("주최자", "카카오");


            //hwp.SaveAs(FilePath2);
            // Hwp 종료
            killProcess(processName);
            
            // COM 객체 해제 Marshal.ReleaseComObject(hwp);
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