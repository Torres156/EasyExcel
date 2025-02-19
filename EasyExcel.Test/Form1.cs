namespace EasyExcel.Test;

public partial class Form1 : Form
{
    public Form1()
    {
        InitializeComponent();
    }

    private void button1_Click(object sender, EventArgs e)
    {
        var dialog = new OpenFileDialog();
        dialog.Filter = "Excel Files (*.xlsx)|*.xlsx";

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            var startTimer = System.Diagnostics.Stopwatch.StartNew();
            using var reader = new ExcelReader(dialog.OpenFile());
            var rows = reader.GetRows();
            startTimer.Stop();
            
            // foreach (var r in rows)
            // {
            //     Console.WriteLine(string.Join(" | ", r.Values));
            // }
            Console.WriteLine(startTimer.Elapsed);
        }
    }
}