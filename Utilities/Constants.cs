namespace SurveyProcessor
{
    // Configuration and constants 
    public static class Constants
    {
        public const int TemplateWorksheetIndex = 3;
        public const int TemplateColumn = 8; // Column H
        public const int TitleRow = 8;
        public const int AboutRow = 10;
        public const int DataStartRow = 3;
        public const int ProcessingStartRow = 10;
        
        public const int CatalogueIdColumn = 9;
        public const int QuestionIdColumn = 10;
        public const int OptionNumberColumn = 11;
        public const int QuestionTypeColumn = 24;
        
        public const string ConsentTrigger = "consent";
        public const string AboutSurveyText = "About the survey";
        public const string SelectKeyword = "select";
        public const string InstructionsType = "Instructions";
        public const string TextQuestionType = "TextQuestion";
        public const string TickboxesType = "Tickboxes";
        public const string SectionHeadingType = "Section Heading";
    }
}