namespace SurveyProcessor
{
    // Survey processor
    public class SurveyProcessor
    {
        public static void ProcessSurvey(ProcessingConfig config)
        {
            using var inputWorkbook = new XLWorkbook(config.InputFile);
            using var templateWorkbook = new XLWorkbook(config.TemplateFile);
            
            Logger.Log("✓ Files loaded successfully");

            var inputWorksheet = inputWorkbook.Worksheet(1);
            var templateWorksheet = templateWorkbook.Worksheet(Constants.TemplateWorksheetIndex);

            SetupTemplate(templateWorksheet, config);
            
            var questions = SurveyDataExtractor.ExtractQuestions(inputWorksheet, config.InputColumn);
            ProcessQuestions(questions, templateWorksheet, config);

            string outputFile = Path.Combine(config.OutputFolder, $"{config.SurveyTitle}_Processed.xlsx");
            Logger.Log($"Saving output file: {outputFile}");
            templateWorkbook.SaveAs(outputFile);

            Logger.Log("✓ Processing completed successfully!");
        }

        private static void SetupTemplate(IXLWorksheet templateSheet, ProcessingConfig config)
        {
            // Set title
            var titleCell = templateSheet.Cell(Constants.TitleRow, Constants.TemplateColumn);
            titleCell.Value = config.SurveyTitle;
            titleCell.Style.Font.Bold = true;
            titleCell.Style.Font.Underline = XLFontUnderlineValues.Single;
            titleCell.Style.Font.FontSize = 16d;
            Logger.Log($"Set title: {config.SurveyTitle}");

            // Set "About the survey"
            var aboutCell = templateSheet.Cell(Constants.AboutRow, Constants.TemplateColumn);
            aboutCell.Value = Constants.AboutSurveyText;
            aboutCell.Style.Font.Bold = true;
        }

        private static void ProcessQuestions(List<QuestionData> questions, IXLWorksheet templateSheet, ProcessingConfig config)
        {
            Logger.Log("Processing survey data...");
            
            int templateRow = Constants.ProcessingStartRow;
            int questionNumber = 0;
            
            foreach (var question in questions)
            {
                var processor = ProcessorFactory.GetProcessor(question.Type);
                var result = processor.Process(question, templateSheet, ref templateRow, config, ref questionNumber);
                
                // Apply common formatting
                for (int row = templateRow - result.RowsProcessed + 1; row <= templateRow; row++)
                {
                    var cell = templateSheet.Cell(row, Constants.TemplateColumn);
                    cell.Style.Alignment.WrapText = true;
                    templateSheet.Row(row).AdjustToContents();
                    templateSheet.Row(row).ClearHeight();
                }
            }
            
            Logger.Log($"Processing completed: {questionNumber} questions processed");
        }
    }
}