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
            var swApp = Connect.run();
            var swModel = swApp!.IActiveDoc2;
            TopologyLabeler.LabelCurrentPart(swModel, wlIterations: 1);
        }
    }
}