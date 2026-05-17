using System.Text;

namespace ExcelMacroRunner;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        ApplicationConfiguration.Initialize();

        using var mutex = new Mutex(true, @"Global\ExcelMacroRunner_SingleInstance", out var createdNew);
        if (!createdNew)
        {
            MessageBox.Show(
                "O ExcelMacroRunner já está em execução nesta máquina.",
                "Instância já aberta",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            return;
        }

        Application.Run(new MainForm());
    }
}
