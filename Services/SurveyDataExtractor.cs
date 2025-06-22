using ClosedXML.Excel;

namespace SurveyProcessor
{
    // Data extractor
    public class SurveyDataExtractor
    {
        public static List<QuestionData> ExtractQuestions(IXLWorksheet inputSheet, int inputColumn)
        {
            var questions = new List<QuestionData>();
            var lastUsedCell = inputSheet.Column(inputColumn).LastCellUsed();
            
            if (lastUsedCell == null)
            {
                Logger.Log("WARNING: No data found in the specified column");
                return questions;
            }

            int colLength = lastUsedCell.Address.RowNumber;
            Logger.Log($"Processing {colLength - Constants.DataStartRow + 1} rows starting from row {Constants.DataStartRow}");

            ProcessingMode mode = ProcessingMode.Description;
            
            for (int row = Constants.DataStartRow; row <= colLength; row++)
            {
                var cell = inputSheet.Cell(row, inputColumn);
                var nextCell = row < colLength ? inputSheet.Cell(row + 1, inputColumn) : null;
                
                string cellValue = cell.Value.ToString()?.Trim() ?? "";
                if (string.IsNullOrWhiteSpace(cellValue))
                    continue;

                var questionType = QuestionTypeDetector.DetectType(cell, nextCell, mode, inputSheet, inputColumn, row, colLength);
                var questionData = new QuestionData
                {
                    Text = cellValue,
                    Type = questionType,
                    SourceRow = row,
                    IsBold = cell.Style.Font.Bold,
                    IsUnderlined = cell.Style.Font.Underline != XLFontUnderlineValues.None
                };

                // Handle special cases that require looking ahead
                if (questionType == QuestionType.TextQuestion && nextCell != null)
                {
                    questionData.Options.Add(nextCell.Value.ToString() ?? "");
                    row++; // Skip the next row
                }
                else if (questionType == QuestionType.MultipleChoice || questionType == QuestionType.Tickboxes)
                {
                    // Collect options for both multiple choice and regular tickboxes
                    for (int optionRow = row + 1; optionRow <= colLength; optionRow++)
                    {
                        var optionCell = inputSheet.Cell(optionRow, inputColumn);
                        string optionValue = optionCell.Value.ToString()?.Trim();
                        
                        if (string.IsNullOrWhiteSpace(optionValue))
                            break;
                            
                        questionData.Options.Add(optionValue);
                        row = optionRow;
                    }
                }
                else if (questionType == QuestionType.Instructions)
                {
                    // Collect sub-questions and options
                    for (int subRow = row + 1; subRow <= colLength; subRow++)
                    {
                        var subCell = inputSheet.Cell(subRow, inputColumn);
                        string subValue = subCell.Value.ToString()?.Trim();
                        
                        if (string.IsNullOrWhiteSpace(subValue))
                            break;
                            
                        if (subCell.Style.Font.Bold)
                        {
                            questionData.SubQuestions.Add(subValue);
                        }
                        else
                        {
                            questionData.Options.Add(subValue);
                        }
                        
                        row = subRow;
                    }
                }

                questions.Add(questionData);

                // Update mode based on question type
                if (questionType == QuestionType.Consent)
                {
                    mode = ProcessingMode.Consent;
                }
                else if (mode == ProcessingMode.Consent)
                {
                    mode = ProcessingMode.Question;
                }
            }

            return questions;
        }
    }
}