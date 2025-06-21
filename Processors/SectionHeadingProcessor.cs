namespace SurveyProcessor
{
    public class SectionHeadingProcessor : IQuestionProcessor
    {
        public ProcessingResult Process(QuestionData question, IXLWorksheet templateSheet, 
                                      ref int templateRow, ProcessingConfig config, ref int questionNumber)
        {
            templateRow++;
            var target = templateSheet.Cell(templateRow, Constants.TemplateColumn);
            target.Value = question.Text.ToUpper();
            target.Style.Font.Bold = true;
            target.Style.Font.Underline = XLFontUnderlineValues.Single;
            templateSheet.Cell(templateRow, Constants.QuestionTypeColumn).Value = Constants.SectionHeadingType;
            
            templateRow++;
            
            return new ProcessingResult { RowsProcessed = 2 };
        }
    }
}