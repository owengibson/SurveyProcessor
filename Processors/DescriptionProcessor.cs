using ClosedXML.Excel;

namespace SurveyProcessor
{
    // Specific processors (same as CLI version)
    public class DescriptionProcessor : IQuestionProcessor
    {
        public ProcessingResult Process(QuestionData question, IXLWorksheet templateSheet, 
                                      ref int templateRow, ProcessingConfig config, ref int questionNumber)
        {
            templateRow++;
            var target = templateSheet.Cell(templateRow, Constants.TemplateColumn);
            target.Value = question.Text;
            target.Style.Alignment.WrapText = true;
            
            templateRow++;
            
            Logger.Log($"Description: {question.Text.Substring(0, Math.Min(50, question.Text.Length))}...");
            
            return new ProcessingResult { RowsProcessed = 2 };
        }
    }
}