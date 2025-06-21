namespace SurveyProcessor
{
    public class QuestionData
    {
        public string Text { get; set; }
        public QuestionType Type { get; set; }
        public List<string> Options { get; set; } = new();
        public List<string> SubQuestions { get; set; } = new();
        public int SourceRow { get; set; }
        public bool IsBold { get; set; }
        public bool IsUnderlined { get; set; }
    }
}