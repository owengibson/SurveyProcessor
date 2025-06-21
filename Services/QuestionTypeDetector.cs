namespace SurveyProcessor
{
    // Question type detector
    public class QuestionTypeDetector
    {
        public static QuestionType DetectType(IXLCell cell, IXLCell nextCell, ProcessingMode currentMode, IXLWorksheet inputSheet, int inputColumn, int currentRow, int maxRow)
        {
            string cellValue = cell.Value.ToString()?.Trim() ?? "";
            string nextValue = nextCell?.Value.ToString()?.Trim() ?? "";

            if (currentMode == ProcessingMode.Description && 
                cellValue.Equals(Constants.ConsentTrigger, StringComparison.OrdinalIgnoreCase))
            {
                return QuestionType.Consent;
            }

            if (cell.Style.Font.Underline != XLFontUnderlineValues.None)
            {
                return QuestionType.SectionHeading;
            }

            if (cellValue.Contains('?'))
            {
                // Check for TextQuestion: next cell has "__" AND the cell after that is empty
                if (nextValue.Contains("__"))
                {
                    // Check if the cell after next is empty (indicating this is a text input, not an option)
                    var cellAfterNext = currentRow + 2 <= maxRow ? inputSheet.Cell(currentRow + 2, inputColumn) : null;
                    string cellAfterNextValue = cellAfterNext?.Value.ToString()?.Trim() ?? "";
                    
                    if (string.IsNullOrWhiteSpace(cellAfterNextValue))
                    {
                        return QuestionType.TextQuestion;
                    }
                }
                
                if (nextCell?.Style.Font.Bold == true)
                {
                    return QuestionType.Instructions;
                }
                
                if (cellValue.ToLower().Contains(Constants.SelectKeyword))
                {
                    return QuestionType.MultipleChoice;
                }
                
                return QuestionType.Tickboxes;
            }

            return QuestionType.Description;
        }
    }
}