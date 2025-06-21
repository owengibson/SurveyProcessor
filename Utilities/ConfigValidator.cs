namespace SurveyProcessor
{
    // Validator utility
    public static class ConfigValidator
    {
        public static bool ValidateConfig(ProcessingConfig config, out string errorMessage)
        {
            errorMessage = null;

            if (string.IsNullOrWhiteSpace(config.SurveyTitle))
            {
                errorMessage = "Survey title cannot be empty";
                return false;
            }

            if (config.InputColumn < 1 || config.InputColumn > 26)
            {
                errorMessage = "Input column must be between 1 and 26 (A-Z)";
                return false;
            }

            if (!File.Exists(config.InputFile))
            {
                errorMessage = $"Input file not found: {config.InputFile}";
                return false;
            }

            if (!File.Exists(config.TemplateFile))
            {
                errorMessage = $"Template file not found: {config.TemplateFile}";
                return false;
            }

            if (!Directory.Exists(config.OutputFolder))
            {
                errorMessage = $"Output folder not found: {config.OutputFolder}";
                return false;
            }

            return true;
        }
    }
}