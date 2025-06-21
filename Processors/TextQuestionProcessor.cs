namespace SurveyProcessor
{
    public class TextQuestionProcessor : IQuestionProcessor
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
            templateSheet.Cell(templateRow, Constants.QuestionTypeColumn).Value = Constants.TextQuestionType;
            
            if (question.Options.Any())
            {
                templateRow++;
                templateSheet.Cell(templateRow, Constants.TemplateColumn).Value = question.Options[0];
            }
            
            templateRow++;
            
            Logger.Log($"Text Question {questionNumber}: {question.Text.Substring(0, Math.Min(60, question.Text.Length))}...");
            
            return new ProcessingResult { RowsProcessed = 3 };
        }
    }
}