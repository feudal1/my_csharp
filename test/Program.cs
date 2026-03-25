using tools;
using cad_tools;
using System.Windows.Forms;

namespace test
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            draw_divider.process_subfolders_with_divider();
        }
    }
}