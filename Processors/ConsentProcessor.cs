using ClosedXML.Excel;

namespace SurveyProcessor
{
    public class ConsentProcessor : IQuestionProcessor
    {
        public ProcessingResult Process(QuestionData question, IXLWorksheet templateSheet, 
                                      ref int templateRow, ProcessingConfig config, ref int questionNumber)
        {
            templateRow++;
            var target = templateSheet.Cell(templateRow, Constants.TemplateColumn);
            target.Value = "Consent";
            target.Style.Font.Bold = true;
            
            Logger.Log("Switched to Consent mode");
            
            return new ProcessingResult { RowsProcessed = 2, NewMode = ProcessingMode.Consent };
        }
    }
}