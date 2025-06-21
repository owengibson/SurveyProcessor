namespace SurveyProcessor
{
    // Configuration class
    public class ProcessingConfig
    {
        public string SurveyTitle { get; set; }
        public int InputColumn { get; set; }
        public string QuestionIdPrefix { get; set; }
        public string CatalogueIdPrefix { get; set; }
        public string InputFile { get; set; }
        public string TemplateFile { get; set; }
        public string OutputFolder { get; set; }
    }
}