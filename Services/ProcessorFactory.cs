namespace SurveyProcessor
{
    // Processor factory
    public static class ProcessorFactory
    {
        private static readonly Dictionary<QuestionType, IQuestionProcessor> Processors = new()
        {
            { QuestionType.Description, new DescriptionProcessor() },
            { QuestionType.Consent, new ConsentProcessor() },
            { QuestionType.SectionHeading, new SectionHeadingProcessor() },
            { QuestionType.TextQuestion, new TextQuestionProcessor() },
            { QuestionType.Tickboxes, new TickboxesProcessor() },
            { QuestionType.MultipleChoice, new MultipleChoiceProcessor() },
            { QuestionType.Instructions, new InstructionsProcessor() }
        };

        public static IQuestionProcessor GetProcessor(QuestionType type)
        {
            return Processors.TryGetValue(type, out var processor) ? processor : Processors[QuestionType.Description];
        }
    }
}