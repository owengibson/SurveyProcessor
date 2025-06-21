namespace SurveyProcessor
{
    public class TickboxesProcessor : IQuestionProcessor
    {
        public ProcessingResult Process(QuestionData question, IXLWorksheet templateSheet, 
                                      ref int templateRow, ProcessingConfig config, ref int questionNumber)
        {
            questionNumber++;
            templateRow++;
            
            var target = templateSheet.Cell(templateRow, Constants.TemplateColumn);
            target.Value = question.Text;
            target.Style.Font.Bold = true;
            target.Style.Alignment.WrapText = true;
            
            string questionId = questionNumber.ToString("00");
            templateSheet.Cell(templateRow, Constants.CatalogueIdColumn).Value = $"{config.CatalogueIdPrefix}_{questionId}";
            templateSheet.Cell(templateRow, Constants.QuestionIdColumn).Value = $"{config.QuestionIdPrefix}_{questionId}";
            templateSheet.Cell(templateRow, Constants.QuestionTypeColumn).Value = Constants.TickboxesType;
            
            int optionNumber = 1;
            foreach (var option in question.Options)
            {
                templateRow++;
                
                templateSheet.Cell(templateRow, Constants.TemplateColumn).Value = option;
                templateSheet.Cell(templateRow, Constants.OptionNumberColumn).Value = optionNumber;
                
                optionNumber++;
            }
            
            templateRow++;
            
            Logger.Log($"Tickboxes Question {questionNumber}: {question.Text.Substring(0, Math.Min(60, question.Text.Length))}... ({question.Options.Count} options)");
            
            return new ProcessingResult { RowsProcessed = 2 + question.Options.Count };
        }
    }
}