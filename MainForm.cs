using System;
using System.IO;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace SurveyProcessor
{
    public class MainForm : Form
    {
        enum ProcessingMode { Description, Consent, Question }

        // Preferences
        private char inputColChar = 'C'; // Column in input sheet to read content from
        private int inputCol;
        private char templateColChar = 'H'; // Column in template sheet to insert content
        private int templateCol;

        // New parameters
        private string qID = "";    // Short Question ID Prefix
        private string catID = "";  // Long Catalogue ID Prefix

        // User input inits
        private string inputFilePath = "";
        private string templateFilePath = "";
        private string outputFolderPath = "";
        private string surveyTitle = "";

        // UI Props
        private Label lblInputFile;
        private Label lblTemplateFile;
        private Label lblOutputFolder;
        private Label lblSurveyTitle;

        // New UI Props
        private Label lblInputCol;
        private Label lblQID;
        private Label lblCatID;
        private TextBox txtInputFile;
        private TextBox txtTemplateFile;
        private TextBox txtOutputFolder;
        private TextBox txtSurveyTitle;
        private TextBox txtInputCol;
        private TextBox txtQID;
        private TextBox txtCatID;
        private Button btnSelectInput;
        private Button btnSelectTemplate;
        private Button btnSelectOutput;
        private Button btnProcess;
        private ProgressBar progressBar;
        private RichTextBox txtLog;

        public MainForm()
        {
            inputCol = char.ToUpper(inputColChar) - 64;
            templateCol = char.ToUpper(templateColChar) - 64;

            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Size = new System.Drawing.Size(600, 650); // Increased height for new fields
            this.Text = "Excel Survey Processor";
            this.StartPosition = FormStartPosition.CenterScreen;

            int y = 20;

            // Survey title input
            lblSurveyTitle = new Label() { Text = "Survey Title", Location = new System.Drawing.Point(20, y), Size = new System.Drawing.Size(120, 23) };
            txtSurveyTitle = new TextBox() { Location = new System.Drawing.Point(150, y), Size = new System.Drawing.Size(300, 23) };
            y += 40;

            // Input Survey Content Column
            lblInputCol = new Label()
            {
                Text = "Input Survey Content Column",
                Location = new System.Drawing.Point(20, y),
                Size = new System.Drawing.Size(200, 23) // Increased width for full label
            };
            txtInputCol = new TextBox()
            {
                Location = new System.Drawing.Point(230, y), // Move textbox to the right of the label
                Size = new System.Drawing.Size(40, 23),
                MaxLength = 1,
                Text = inputColChar.ToString()
            };
            txtInputCol.Text = inputColChar.ToString();
            txtInputCol.TextChanged += (s, e) =>
            {
                if (!string.IsNullOrEmpty(txtInputCol.Text))
                {
                    inputColChar = txtInputCol.Text[0];
                    inputCol = char.ToUpper(inputColChar) - 64;
                }
            };
            y += 40;

            // Short Question ID Prefix
            lblQID = new Label()
            {
                Text = "Short Question ID Prefix",
                Location = new System.Drawing.Point(20, y),
                Size = new System.Drawing.Size(200, 23)
            };
            txtQID = new TextBox()
            {
                Location = new System.Drawing.Point(230, y),
                Size = new System.Drawing.Size(100, 23)
            };
            txtQID.TextChanged += (s, e) => { qID = txtQID.Text; };
            y += 40;

            // Long Catalogue ID Prefix
            lblCatID = new Label()
            {
                Text = "Long Catalogue ID Prefix",
                Location = new System.Drawing.Point(20, y),
                Size = new System.Drawing.Size(200, 23)
            };
            txtCatID = new TextBox()
            {
                Location = new System.Drawing.Point(230, y),
                Size = new System.Drawing.Size(100, 23)
            };
            txtCatID.TextChanged += (s, e) => { catID = txtCatID.Text; };
            y += 40;

            // Add tooltips for the new labels
            var toolTip = new ToolTip();
            toolTip.SetToolTip(lblInputCol, "The column in the spreadsheet the program should read all content from.");
            toolTip.SetToolTip(lblQID, "The prefix of the Short Question IDs (eg. MMSMQ).");
            toolTip.SetToolTip(lblCatID, "The prefix of the long catalogue ID. Probably the short question ID, with the year appended (eg. MMSMQ25).");

            // Input file selection
            lblInputFile = new Label() { Text = "Input Excel File:", Location = new System.Drawing.Point(20, y), Size = new System.Drawing.Size(100, 23) };
            txtInputFile = new TextBox() { Location = new System.Drawing.Point(130, y), Size = new System.Drawing.Size(300, 23), ReadOnly = true };
            btnSelectInput = new Button() { Text = "Browse", Location = new System.Drawing.Point(440, y - 1), Size = new System.Drawing.Size(85, 35) };
            btnSelectInput.Click += BtnSelectInput_Click;
            y += 40;

            // Template file selection
            lblTemplateFile = new Label() { Text = "Template File:", Location = new System.Drawing.Point(20, y), Size = new System.Drawing.Size(100, 23) };
            txtTemplateFile = new TextBox() { Location = new System.Drawing.Point(130, y), Size = new System.Drawing.Size(300, 23), ReadOnly = true };
            btnSelectTemplate = new Button() { Text = "Browse", Location = new System.Drawing.Point(440, y - 1), Size = new System.Drawing.Size(85, 35) };
            btnSelectTemplate.Click += BtnSelectTemplate_Click;
            y += 40;

            // Output folder selection
            lblOutputFolder = new Label() { Text = "Output Folder:", Location = new System.Drawing.Point(20, y), Size = new System.Drawing.Size(100, 23) };
            txtOutputFolder = new TextBox() { Location = new System.Drawing.Point(130, y), Size = new System.Drawing.Size(300, 23), ReadOnly = true };
            btnSelectOutput = new Button() { Text = "Browse", Location = new System.Drawing.Point(440, y - 1), Size = new System.Drawing.Size(85, 35) };
            btnSelectOutput.Click += BtnSelectOutput_Click;
            y += 40;

            // Process button
            btnProcess = new Button() { Text = "Process Files", Location = new System.Drawing.Point(250, y), Size = new System.Drawing.Size(100, 30) };
            btnProcess.Click += BtnProcess_Click;
            y += 40;

            // Progress bar
            progressBar = new ProgressBar() { Location = new System.Drawing.Point(20, y), Size = new System.Drawing.Size(520, 23), Visible = false };
            y += 30;

            // Log text box
            txtLog = new RichTextBox() { Location = new System.Drawing.Point(20, y), Size = new System.Drawing.Size(520, 250), ReadOnly = true };

            // Add controls to form
            this.Controls.AddRange(new Control[] {
                lblSurveyTitle, txtSurveyTitle,
                lblInputCol, txtInputCol,
                lblQID, txtQID,
                lblCatID, txtCatID,
                lblInputFile, txtInputFile, btnSelectInput,
                lblTemplateFile, txtTemplateFile, btnSelectTemplate,
                lblOutputFolder, txtOutputFolder, btnSelectOutput,
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
            using (FolderBrowserDialog saveFileDialog = new FolderBrowserDialog())
            {
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    outputFolderPath = saveFileDialog.SelectedPath;
                    txtOutputFolder.Text = outputFolderPath;
                    LogMessage($"Output file path set: {Path.GetFileName(outputFolderPath)}");
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
                return;
            }

            if (string.IsNullOrEmpty(txtInputCol.Text.Trim()))
            {
                MessageBox.Show("Please enter the input survey content column.", "Missing Input Column", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(txtQID.Text.Trim()))
            {
                MessageBox.Show("Please enter the Short Question ID Prefix.", "Missing Question ID Prefix", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(txtCatID.Text.Trim()))
            {
                MessageBox.Show("Please enter the Long Catalogue ID Prefix.", "Missing Catalogue ID Prefix", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(inputFilePath) || string.IsNullOrEmpty(templateFilePath) || string.IsNullOrEmpty(outputFolderPath))
            {
                MessageBox.Show("Please select all required files or folders.", "Missing Files/Folders", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                        string outputFile = $"{outputFolderPath}\\{surveyTitle}_Processed.xlsx";
                        templateWorkbook.SaveAs(outputFile);
                        LogMessage($"Output file saved: {Path.GetFileName(outputFile)}");

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
            // TODO: Replace this section with your specific logic
            // You can:
            // - Read values: inputSheet.Cell("A1").GetValue<string>()
            // - Write values: templateSheet.Cell("B2").Value = "Your Value"
            // - Loop through ranges: inputSheet.Range("A1:Z100")
            // - Apply formatting: cell.Style.Fill.BackgroundColor = XLColor.Blue
            // - Use formulas: templateSheet.Cell("C1").FormulaA1 = "=A1+B1"

            // Title
            var titleCell = templateSheet.Cell(8, templateCol);
            titleCell.Value = surveyTitle;
            titleCell.Style.Font.Bold = true;
            titleCell.Style.Font.Underline = XLFontUnderlineValues.Single;
            titleCell.Style.Font.FontSize = 16d;

            // About the survey
            var aboutCell = templateSheet.Cell(10, templateCol);
            aboutCell.Value = "About the survey";
            aboutCell.Style.Font.Bold = true;

            ProcessingMode mode = ProcessingMode.Description;

            int questionNo = 0;
            var colLength = inputSheet.Column(inputCol).LastCellUsed().Address.RowNumber;
            for (int row = 3; row <= colLength; row++)
            {
                int templateRow = row + 8;
                var cell = inputSheet.Cell(row, inputCol);
                var target = templateSheet.Cell(templateRow, templateCol);

                switch (mode)
                {
                    case ProcessingMode.Description:
                        // Description
                        if (cell.Value.ToString().ToLower() == "consent")
                        {
                            target.Value = "Consent";
                            target.Style.Font.Bold = true;
                            mode = ProcessingMode.Consent;
                            continue;
                        }
                        target.Value = cell.Value;
                        break;

                    case ProcessingMode.Consent:
                        // Consent
                        target.Value = cell.Value;
                        mode = ProcessingMode.Question;
                        break;

                    case ProcessingMode.Question:
                        // Question
                        if (cell.Value.ToString().Contains('?'))
                        {
                            questionNo++;
                            string q = questionNo.ToString("00");
                            target.Value = cell.Value;
                            target.Style.Font.Bold = true;
                            
                            templateSheet.Cell(templateRow, 9).Value = $"{catID}_{q}";
                            templateSheet.Cell(templateRow, 10).Value = $"{qID}_{q}";

                            // TextQuestion
                            if (cell.CellBelow().Value.ToString().Contains("__"))
                            {
                                templateSheet.Cell(templateRow, 24).Value = "TextQuestion";
                                target.CellBelow().Value = cell.CellBelow().Value;
                                row++;
                            }
                            // Multiple tickboxes
                            else if (cell.Value.ToString().ToLower().Contains("select"))
                            {
                                templateSheet.Cell(templateRow, 24).Value = "Tickboxes";
                                
                                // Iterate through options
                            }
                        }

                        break;

                    default:
                        break;
                }
                target.Style.Alignment.WrapText = true;
                templateSheet.Row(row).AdjustToContents();
                templateSheet.Row(row).ClearHeight();
            }

            LogMessage("Data operations completed.");
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