//using System;
//using System.IO;
using HwpObjectLib;
using System.Runtime.InteropServices;

internal class Program
{
    private static void Main(string[] args)
    {
        Console.WriteLine("Hello, World!");
        try
        {
            string FilePath = @"D:\C샵hwp연습용\2023_Csarp_hwp\2023_Csarp_vs\ch1\ConsoleApp1\ConsoleApp1\TempData\h25_certificate_of_award.hwp";
            //string FilePath2 = @"D:\C샵hwp연습용\2023_Csarp_hwp\2023_Csarp_vs\ch1\ConsoleApp1\ConsoleApp1\TempData\h25_certificate_of_award2.hwp";
            Console.WriteLine(FilePath);
            // HwpObject 객체 생성
            IHwpObject hwp = new HwpObject();

            hwp.XHwpWindows.Active_XHwpWindow.Visible = true;

            hwp.Open(FilePath, "", "");

            //hwp.PutFieldText("주최자", "카카오");

            //hwp.SaveAs(FilePath2);
            // HwpCtrl 종료
            // COM 객체 해제 Marshal.ReleaseComObject(hwp);
            //Marshal.FinalReleaseComObject(hwp);
            //hwp = null;
       }
        catch (Exception ex) {
            Console.WriteLine("예외 발생");
            Console.WriteLine(ex.Message);
        }

    }
}