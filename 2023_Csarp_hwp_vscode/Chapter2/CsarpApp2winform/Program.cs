using System.Windows.Forms;
using System.Drawing;
//using System.Runtime.Intrinsics.X86;
//using System.Runtime.CompilerServices;

class Sample
{
    public static void Main()
    {
        Console.WriteLine("Hello, World!");

        Form f = new Form();
        f.Text = "샘플";

        Label l = new Label();
        l.Width = 100;
        l.Text = "C#입니다.";
        l.Parent = f;

        // PictureBox p = new PictureBox();
        // string imPath = Path.Combine(Application.StartupPath, "Downloads", "우주사진.jpg");
        // p.Image = Image.FromFile(imPath);
        // p.Top = 50;
        // p.Left = 50;
        // p.Parent = f;

        Application.Run(f);
        //dotnet 버전변경해서 다시 해보기
    }
}