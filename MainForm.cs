using System;
using System.IO;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace SurveyProcessor
{
    public class MainForm : Form
    {
        private string inputFilePath = "";
        private string templateFilePath = "";
        private string outputFilePath = "";
        private string surveyTitle = "";

        private Label lblInputFile;
        private Label lblTemplateFile;
        private Label lblOutputFile;
        private Label lblSurveyTitle;
        private Button btnSelectInput;
        private Button btnSelectTemplate;
        private Button btnSelectOutput;
        private Button btnProcess;
        private TextBox txtInputFile;
        private TextBox txtTemplateFile;
        private TextBox txtOutputFile;
        private TextBox txtSurveyTitle;
        private ProgressBar progressBar;
        private RichTextBox txtLog;

        public MainForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Size = new System.Drawing.Size(600, 550);
            this.Text = "Excel Survey Processor";
            this.StartPosition = FormStartPosition.CenterScreen;

            // Survey title input
            lblSurveyTitle = new Label() { Text = "Survey Title", Location = new System.Drawing.Point(20, 20), Size = new System.Drawing.Size(100, 23) };
            txtSurveyTitle = new TextBox() { Location = new System.Drawing.Point(130, 20), Size = new System.Drawing.Size(300, 23) };

            // Input file selection
            lblInputFile = new Label() { Text = "Input Excel File:", Location = new System.Drawing.Point(20, 60), Size = new System.Drawing.Size(100, 23) };
            txtInputFile = new TextBox() { Location = new System.Drawing.Point(130, 60), Size = new System.Drawing.Size(300, 23), ReadOnly = true };
            btnSelectInput = new Button() { Text = "Browse", Location = new System.Drawing.Point(440, 59), Size = new System.Drawing.Size(85, 35) };
            btnSelectInput.Click += BtnSelectInput_Click;

            // Template file selection
            lblTemplateFile = new Label() { Text = "Template File:", Location = new System.Drawing.Point(20, 100), Size = new System.Drawing.Size(100, 23) };
            txtTemplateFile = new TextBox() { Location = new System.Drawing.Point(130, 100), Size = new System.Drawing.Size(300, 23), ReadOnly = true };
            btnSelectTemplate = new Button() { Text = "Browse", Location = new System.Drawing.Point(440, 99), Size = new System.Drawing.Size(85, 35) };
            btnSelectTemplate.Click += BtnSelectTemplate_Click;

            // Output file selection
            lblOutputFile = new Label() { Text = "Output File:", Location = new System.Drawing.Point(20, 140), Size = new System.Drawing.Size(100, 23) };
            txtOutputFile = new TextBox() { Location = new System.Drawing.Point(130, 140), Size = new System.Drawing.Size(300, 23), ReadOnly = true };
            btnSelectOutput = new Button() { Text = "Browse", Location = new System.Drawing.Point(440, 139), Size = new System.Drawing.Size(85, 35) };
            btnSelectOutput.Click += BtnSelectOutput_Click;

            // Process button
            btnProcess = new Button() { Text = "Process Files", Location = new System.Drawing.Point(250, 180), Size = new System.Drawing.Size(100, 30) };
            btnProcess.Click += BtnProcess_Click;

            // Progress bar
            progressBar = new ProgressBar() { Location = new System.Drawing.Point(20, 220), Size = new System.Drawing.Size(520, 23), Visible = false };

            // Log text box
            txtLog = new RichTextBox() { Location = new System.Drawing.Point(20, 250), Size = new System.Drawing.Size(520, 250), ReadOnly = true };

            // Add controls to form
            this.Controls.AddRange(new Control[] {
                lblSurveyTitle, txtSurveyTitle,
                lblInputFile, txtInputFile, btnSelectInput,
                lblTemplateFile, txtTemplateFile, btnSelectTemplate,
                lblOutputFile, txtOutputFile, btnSelectOutput,
                btnProcess, progressBar, txtLog
            });
        }

        private void BtnSelectInput_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    inputFilePath = openFileDialog.FileName;
                    txtInputFile.Text = inputFilePath;
                    LogMessage($"Input file selected: {Path.GetFileName(inputFilePath)}");
                }
            }
        }

        private void BtnSelectTemplate_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    templateFilePath = openFileDialog.FileName;
                    txtTemplateFile.Text = templateFilePath;
                    LogMessage($"Template file selected: {Path.GetFileName(templateFilePath)}");
                }
            }
        }

        private void BtnSelectOutput_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    outputFilePath = saveFileDialog.FileName;
                    txtOutputFile.Text = outputFilePath;
                    LogMessage($"Output file path set: {Path.GetFileName(outputFilePath)}");
                }
            }
        }

        private void BtnProcess_Click(object sender, EventArgs e)
        {
            surveyTitle = txtSurveyTitle.Text.Trim();

            // Validation
            if (string.IsNullOrEmpty(surveyTitle))
            {
                MessageBox.Show("Please enter a survey title.", "Missing Survey Title", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            if (string.IsNullOrEmpty(inputFilePath) || string.IsNullOrEmpty(templateFilePath) || string.IsNullOrEmpty(outputFilePath))
            {
                MessageBox.Show("Please select all required files.", "Missing Files", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ProcessFiles();
        }

        private void ProcessFiles()
        {
            try
            {
                progressBar.Visible = true;
                progressBar.Value = 0;
                btnProcess.Enabled = false;

                LogMessage("Starting file processing...");

                // Load the input Excel file
                progressBar.Value = 20;
                using (var inputWorkbook = new XLWorkbook(inputFilePath))
                {
                    LogMessage("Input file loaded successfully.");

                    // Load the template file
                    progressBar.Value = 40;
                    using (var templateWorkbook = new XLWorkbook(templateFilePath))
                    {
                        LogMessage("Template file loaded successfully.");

                        // Get the worksheets (assuming first worksheet for now)
                        var inputWorksheet = inputWorkbook.Worksheet(1);
                        var templateWorksheet = templateWorkbook.Worksheet(3);

                        progressBar.Value = 60;

                        // Perform your specific operations here
                        PerformDataOperations(inputWorksheet, templateWorksheet);

                        progressBar.Value = 80;

                        // Save the modified template as the output file
                        templateWorkbook.SaveAs(outputFilePath);
                        LogMessage($"Output file saved: {Path.GetFileName(outputFilePath)}");

                        progressBar.Value = 100;
                        LogMessage("Processing completed successfully!");

                        MessageBox.Show("File processing completed successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error: {ex.Message}");
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                progressBar.Visible = false;
                btnProcess.Enabled = true;
            }
        }

        private void PerformDataOperations(IXLWorksheet inputSheet, IXLWorksheet templateSheet)
        {
            // This is where you'll implement your specific logic
            // Example operations:

            // Read values from input sheet
            var cellA1Value = inputSheet.Cell("A1").GetValue<string>();
            var cellB1Value = inputSheet.Cell("B1").GetValue<string>();
            LogMessage($"Read from input: A1={cellA1Value}, B1={cellB1Value}");

            // Fill template based on input values
            // Example: If input A1 contains "Name", put it in template C1
            if (!string.IsNullOrEmpty(cellA1Value))
            {
                templateSheet.Cell("C1").Value = cellA1Value;
                LogMessage($"Set template C1 to: {cellA1Value}");
            }

            // Example: Copy data from input B column to template D column
            for (int row = 1; row <= 10; row++) // Adjust range as needed
            {
                var sourceValue = inputSheet.Cell(row, 2).GetValue<string>(); // Column B
                if (!string.IsNullOrEmpty(sourceValue))
                {
                    templateSheet.Cell(row, 4).Value = sourceValue; // Column D
                }
            }

            // Example: Conditional logic
            var statusValue = inputSheet.Cell("E1").GetValue<string>();
            if (statusValue == "Active")
            {
                templateSheet.Cell("F1").Value = "Processed";
                templateSheet.Cell("F1").Style.Fill.BackgroundColor = XLColor.Green;
            }
            else
            {
                templateSheet.Cell("F1").Value = "Pending";
                templateSheet.Cell("F1").Style.Fill.BackgroundColor = XLColor.Yellow;
            }

            LogMessage("Data operations completed.");

            // TODO: Replace this section with your specific logic
            // You can:
            // - Read values: inputSheet.Cell("A1").GetValue<string>()
            // - Write values: templateSheet.Cell("B2").Value = "Your Value"
            // - Loop through ranges: inputSheet.Range("A1:Z100")
            // - Apply formatting: cell.Style.Fill.BackgroundColor = XLColor.Blue
            // - Use formulas: templateSheet.Cell("C1").FormulaA1 = "=A1+B1"
        }

        private void LogMessage(string message)
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            txtLog.AppendText($"[{timestamp}] {message}\n");
            txtLog.ScrollToCaret();
            Application.DoEvents(); // Allow UI to update
        }
    }
}