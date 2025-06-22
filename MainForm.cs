using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace SurveyProcessor
{
        public class MainForm : Form
    {
        // Preferences
        private char inputColChar = 'C';
        private int inputCol;

        // User input
        private string inputFilePath = "";
        private string templateFilePath = "";
        private string outputFolderPath = "";
        private string surveyTitle = "";
        private string qID = "";
        private string catID = "";
        
        // Default files
        private string defaultTemplatePath = "";
        private string defaultInputPath = "";

        // UI Controls
        private Label lblInputFile;
        private Label lblTemplateFile;
        private Label lblOutputFolder;
        private Label lblSurveyTitle;
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
        private Button btnResetTemplate;
        private Button btnResetInput;
        private Button btnProcess;
        private ProgressBar progressBar;
        private RichTextBox txtLog;

        public MainForm()
        {
            inputCol = char.ToUpper(inputColChar) - 64;
            
            // Extract default files from embedded resources
            ExtractDefaultFiles();
            
            InitializeComponent();
            Logger.SetLogTextBox(txtLog); // Connect logger to UI
            
            // Set default files
            if (!string.IsNullOrEmpty(defaultInputPath))
            {
                inputFilePath = defaultInputPath;
                txtInputFile.Text = "Example Input";
                Logger.Log("Default input file loaded");
            }
            
            if (!string.IsNullOrEmpty(defaultTemplatePath))
            {
                templateFilePath = defaultTemplatePath;
                txtTemplateFile.Text = "Default Template";
                Logger.Log("Default template loaded");
            }
        }

        private void InitializeComponent()
        {
            this.Size = new System.Drawing.Size(650, 650); // Increased width for Default button
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
                Size = new System.Drawing.Size(200, 23)
            };
            txtInputCol = new TextBox()
            {
                Location = new System.Drawing.Point(230, y),
                Size = new System.Drawing.Size(40, 23),
                MaxLength = 1,
                Text = inputColChar.ToString()
            };
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

            // Add tooltips
            var toolTip = new ToolTip();
            toolTip.SetToolTip(lblInputCol, "The column in the spreadsheet the program should read all content from.");
            toolTip.SetToolTip(lblQID, "The prefix of the Short Question IDs (eg. MMSMQ).");
            toolTip.SetToolTip(lblCatID, "The prefix of the long catalogue ID. Probably the short question ID, with the year appended (eg. MMSMQ25).");

            // Input file selection
            lblInputFile = new Label() { Text = "Input Excel File:", Location = new System.Drawing.Point(20, y), Size = new System.Drawing.Size(100, 23) };
            txtInputFile = new TextBox() { Location = new System.Drawing.Point(130, y), Size = new System.Drawing.Size(300, 23), ReadOnly = true };
            btnSelectInput = new Button() { Text = "Browse", Location = new System.Drawing.Point(440, y - 1), Size = new System.Drawing.Size(85, 35) };
            btnSelectInput.Click += BtnSelectInput_Click;
            
            // Add reset to default input button
            btnResetInput = new Button() { Text = "Default", Location = new System.Drawing.Point(530, y - 1), Size = new System.Drawing.Size(55, 35) };
            btnResetInput.Click += (s, e) => {
                if (!string.IsNullOrEmpty(defaultInputPath))
                {
                    inputFilePath = defaultInputPath;
                    txtInputFile.Text = "Example Input";
                    Logger.Log("Reset to default input file");
                }
            };
            y += 40;

            // Template file selection
            lblTemplateFile = new Label() { Text = "Template File:", Location = new System.Drawing.Point(20, y), Size = new System.Drawing.Size(100, 23) };
            txtTemplateFile = new TextBox() { Location = new System.Drawing.Point(130, y), Size = new System.Drawing.Size(300, 23), ReadOnly = true };
            btnSelectTemplate = new Button() { Text = "Browse", Location = new System.Drawing.Point(440, y - 1), Size = new System.Drawing.Size(85, 35) };
            btnSelectTemplate.Click += BtnSelectTemplate_Click;
            
            // Add reset to default button
            btnResetTemplate = new Button() { Text = "Default", Location = new System.Drawing.Point(530, y - 1), Size = new System.Drawing.Size(55, 35) };
            btnResetTemplate.Click += (s, e) => {
                if (!string.IsNullOrEmpty(defaultTemplatePath))
                {
                    templateFilePath = defaultTemplatePath;
                    txtTemplateFile.Text = "Default Template";
                    Logger.Log("Reset to default template");
                }
            };
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
            progressBar = new ProgressBar() { Location = new System.Drawing.Point(20, y), Size = new System.Drawing.Size(570, 23), Visible = false };
            y += 30;

            // Log text box
            txtLog = new RichTextBox() { Location = new System.Drawing.Point(20, y), Size = new System.Drawing.Size(570, 250), ReadOnly = true };

            // Add controls to form
            this.Controls.AddRange(new Control[] {
                lblSurveyTitle, txtSurveyTitle,
                lblInputCol, txtInputCol,
                lblQID, txtQID,
                lblCatID, txtCatID,
                lblInputFile, txtInputFile, btnSelectInput, btnResetInput,
                lblTemplateFile, txtTemplateFile, btnSelectTemplate, btnResetTemplate,
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
                    txtInputFile.Text = Path.GetFileName(inputFilePath);
                    Logger.Log($"Custom input file selected: {Path.GetFileName(inputFilePath)}");
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
                    txtTemplateFile.Text = Path.GetFileName(templateFilePath);
                    Logger.Log($"Custom template selected: {Path.GetFileName(templateFilePath)}");
                }
            }
        }

        private void BtnSelectOutput_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    outputFolderPath = folderDialog.SelectedPath;
                    txtOutputFolder.Text = outputFolderPath;
                    Logger.Log($"Output folder selected: {outputFolderPath}");
                }
            }
        }

        private void BtnProcess_Click(object sender, EventArgs e)
        {
            surveyTitle = txtSurveyTitle.Text.Trim();
            qID = txtQID.Text.Trim();
            catID = txtCatID.Text.Trim();

            var config = new ProcessingConfig
            {
                SurveyTitle = surveyTitle,
                InputColumn = inputCol,
                QuestionIdPrefix = qID,
                CatalogueIdPrefix = catID,
                InputFile = inputFilePath,
                TemplateFile = templateFilePath,
                OutputFolder = outputFolderPath
            };

            // Validation
            if (!ConfigValidator.ValidateConfig(config, out string errorMessage))
            {
                MessageBox.Show(errorMessage, "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ProcessFiles(config);
        }

        private void ProcessFiles(ProcessingConfig config)
        {
            try
            {
                progressBar.Visible = true;
                progressBar.Value = 0;
                btnProcess.Enabled = false;

                Logger.Log("Starting file processing...");
                progressBar.Value = 20;

                Logger.Log($"Configuration: {config.SurveyTitle}, Column {(char)(config.InputColumn + 64)}, {config.QuestionIdPrefix}/{config.CatalogueIdPrefix}");
                progressBar.Value = 40;

                SurveyProcessor.ProcessSurvey(config);
                progressBar.Value = 100;

                MessageBox.Show("File processing completed successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Logger.Log($"Error: {ex.Message}");
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                progressBar.Visible = false;
                btnProcess.Enabled = true;
            }
        }

        private void ExtractDefaultFiles()
        {
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                
                // Extract default template
                var templateResourceName = "SurveyProcessor.Resources.DefaultTemplate.xlsx";
                using (Stream stream = assembly.GetManifestResourceStream(templateResourceName))
                {
                    if (stream != null)
                    {
                        defaultTemplatePath = Path.Combine(Path.GetTempPath(), "SurveyProcessor_DefaultTemplate.xlsx");
                        
                        using (var fileStream = File.Create(defaultTemplatePath))
                        {
                            stream.CopyTo(fileStream);
                        }
                        
                        Logger.Log("Default template extracted successfully");
                    }
                    else
                    {
                        Logger.Log("Warning: Default template resource not found");
                    }
                }
                
                // Extract default input file
                var inputResourceName = "SurveyProcessor.Resources.ExampleInput.xlsx";
                using (Stream stream = assembly.GetManifestResourceStream(inputResourceName))
                {
                    if (stream != null)
                    {
                        defaultInputPath = Path.Combine(Path.GetTempPath(), "SurveyProcessor_ExampleInput.xlsx");
                        
                        using (var fileStream = File.Create(defaultInputPath))
                        {
                            stream.CopyTo(fileStream);
                        }
                        
                        Logger.Log("Default input file extracted successfully");
                    }
                    else
                    {
                        Logger.Log("Warning: Default input file resource not found");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"Warning: Could not extract default files: {ex.Message}");
            }
        }
    }
}