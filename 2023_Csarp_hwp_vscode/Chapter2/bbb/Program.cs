// using System.Reflection;
// public class MyLibrary
// {
// 	[System.Runtime.InteropServices.DllImport(@"C:\hwp_보안모듈_Automation\HwpObjectLib\Interop.HwpObjectLib.dll")]
// 	public static extern void IHwpObject();
//     public static extern void HwpObject();
// }
internal class Program
{
    private static void Main(string[] args)

    {

        // 어셈블리 파일의 전체 경로를 지정
            string assemblyPath = @"D:\C샵hwp연습용\2023_Csarp_hwp\2023_Csarp_hwp_vscode\Chapter2\aaa\obj\Debug\net8.0\Interop.HwpObjectLib.dll";

            // 어셈블리를 동적으로 로드
            //System.Reflection.Assembly.LoadFrom(assemblyPath);
            //MyLibrary.IHwpObject hwp = new MyLibrary.HwpObject();

        Console.WriteLine("Hello, World!");
    }
}