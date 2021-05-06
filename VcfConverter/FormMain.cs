using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using VcfConverter.Classes;
using VcfConverter.Properties;

namespace VcfConverter
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
        }

        private void Log(string msg)
        {
            textBoxLog.Text += msg + Environment.NewLine;
        }
        private void Log(Exception exception)
        {
            textBoxLog.Text += exception.Message + Environment.NewLine;
        }
        private void SetConvertType()
        {
            var extension = Path.GetExtension(openFileDialog1.FileName);
            switch (extension)
            {
                case ".xlsx":
                    buttonStartConverting.Text = "Convert To VCF";
                    buttonStartConverting.Enabled = true;
                    pictureBox1.Image = Resources.excel_vcf;
                    break;
                case ".vcf":
                    buttonStartConverting.Text = "Convert To Excel";
                    buttonStartConverting.Enabled = true;
                    pictureBox1.Image = Resources.vcf_excel;
                    break;
                default:
                    Log("File extension is not supported !");
                    buttonStartConverting.Text = "....";
                    buttonStartConverting.Enabled = false;
                    pictureBox1.Image = null;
                    break;
            }
        }

        private void buttonSelectVcfFile_Click(object sender, EventArgs e)
        {
            var dialogResult = openFileDialog1.ShowDialog();
            if (dialogResult == DialogResult.OK)
                textBoxSrcFilePath.Text = openFileDialog1.FileName;
        }
        private void textBoxSrcFilePath_TextChanged(object sender, EventArgs e)
        {
            SetConvertType();
        }
        private void buttonStartConverting_Click(object sender, EventArgs e)
        {
            if (!buttonStartConverting.Enabled) return;
            try
            {
                var stopWatch = Stopwatch.StartNew();
                buttonStartConverting.Enabled = false;
                Log($"Converting Start at: {DateTime.Now}");
                var vcfConverter = new Classes.VcfConverter(textBoxSrcFilePath.Text);
                var converterResult = vcfConverter.Convert(Log);

                #region Save to file

                var destDirectoryPath = Path.GetDirectoryName(textBoxSrcFilePath.Text) ?? "";
                var destFileNameWithoutExtension = Path.GetFileNameWithoutExtension(textBoxSrcFilePath.Text) ?? "";
                var destFilePath = "";
                switch (converterResult.FileFormat)
                {
                    case VcfConverterFileFormat.Vcf:
                        destFilePath = Path.Combine(destDirectoryPath, $"{destFileNameWithoutExtension}.vcf");
                        break;
                    case VcfConverterFileFormat.Excel:
                        destFilePath = Path.Combine(destDirectoryPath, $"{destFileNameWithoutExtension}.xlsx");
                        break;
                }
                var fileCounter = 0;
                while (File.Exists(destFilePath))
                {
                    fileCounter++;
                    destFilePath = Path.Combine(destDirectoryPath, $"{destFileNameWithoutExtension}{fileCounter}{Path.GetExtension(destFilePath)}");
                }
                File.WriteAllBytes(destFilePath, converterResult.FileContent);

                #endregion

                Log($"Converting Done at: {stopWatch.ElapsedMilliseconds / 1000} seconds");
                stopWatch.Stop();
                buttonStartConverting.Enabled = true;
            }
            catch (Exception exception)
            {
                Log(exception);
                buttonStartConverting.Enabled = true;
            }
        }
    }
}
