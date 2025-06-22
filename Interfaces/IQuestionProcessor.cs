using ClosedXML.Excel;
namespace SurveyProcessor
{
    // Question processor interface
    public interface IQuestionProcessor
    {
        ProcessingResult Process(QuestionData question, IXLWorksheet templateSheet, 
                               ref int templateRow, ProcessingConfig config, ref int questionNumber);
    }
}