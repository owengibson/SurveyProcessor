namespace SurveyProcessor
{
    public class InstructionsProcessor : IQuestionProcessor
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
            
            templateSheet.Cell(templateRow, Constants.QuestionTypeColumn).Value = Constants.InstructionsType;
            
            templateRow += 2;
            
            int subQuestionNumber = 1;
            foreach (var subQuestion in question.SubQuestions)
            {
                templateSheet.Cell(templateRow, Constants.TemplateColumn).Value = subQuestion;
                templateSheet.Cell(templateRow, Constants.TemplateColumn).Style.Font.Bold = true;
                templateSheet.Cell(templateRow, Constants.CatalogueIdColumn).Value = $"{config.CatalogueIdPrefix}_{questionNumber}_{subQuestionNumber:00}";
                templateSheet.Cell(templateRow, Constants.QuestionIdColumn).Value = $"{config.QuestionIdPrefix}_{questionNumber}_{subQuestionNumber:00}";
                templateSheet.Cell(templateRow, Constants.QuestionTypeColumn).Value = Constants.TickboxesType;
                
                templateRow++;
                
                int optionNumber = 1;
                foreach (var option in question.Options)
                {
                    templateSheet.Cell(templateRow, Constants.TemplateColumn).Value = option;
                    templateSheet.Cell(templateRow, Constants.OptionNumberColumn).Value = optionNumber;
                    templateRow++;
                    optionNumber++;
                }
                
                subQuestionNumber++;
                if (subQuestionNumber <= question.SubQuestions.Count)
                {
                    templateRow++;
                }
            }
            
            templateRow++;
            
            Logger.Log($"Instructions Question {questionNumber}: {question.SubQuestions.Count} sub-questions, {question.Options.Count} options each");
            
            int totalRowsUsed = 3 + (question.SubQuestions.Count * (1 + question.Options.Count)) + (question.SubQuestions.Count - 1) + 1;
            return new ProcessingResult { RowsProcessed = totalRowsUsed };
        }
    }
}