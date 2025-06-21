# Survey Processor

A powerful tool for converting survey content from Excel spreadsheets into structured survey templates with automatic question type detection and formatting. Made for use by NCLS.

## Table of Contents

-   [Overview](#overview)
-   [Question Type Detection](#question-type-detection)
-   [Input Survey Formatting](#input-survey-formatting)
-   [Question Types & Examples](#question-types--examples)
-   [Output Format](#output-format)
-   [Architecture](#architecture)
-   [Usage](#usage)

## Overview

The Survey Processor automatically detects different types of survey content and formats them appropriately in a template. It processes survey content from a specified column in an Excel file and generates a structured output with proper question IDs, types, and formatting.

### Key Features

-   **Automatic Question Type Detection**: Identifies different question types based on content and formatting
-   **Structured Output**: Generates question IDs, catalogue IDs, and type classifications
-   **Multiple Question Types**: Supports descriptions, consent forms, text questions, multiple choice, tickboxes, and more
-   **Cross-Platform**: Available as both CLI (Mac/Linux/Windows) and WinForms (Windows) versions

## Question Type Detection

The application processes your survey content in three main phases:

1. **Description Phase**: Everything before the "consent" trigger
2. **Consent Phase**: Single row after "consent" is detected
3. **Question Phase**: Everything after consent, with automatic question type detection

### Processing Modes

```
Input Survey Flow:
┌─────────────────┐
│   Description   │ ← General survey information
│   Content       │
├─────────────────┤
│    "consent"    │ ← Trigger word (case insensitive)
├─────────────────┤
│  Consent Text   │ ← One row of consent information
├─────────────────┤
│   Questions     │ ← Automatically typed questions
│   & Responses   │
└─────────────────┘
```

## Input Survey Formatting

### Basic Rules

1. **Single Column Input**: All content should be in one column (default: Column C)
2. **Sequential Processing**: Content is processed row by row from top to bottom
3. **Trigger Words**: "consent" switches from description to consent mode
4. **Question Detection**: Questions are identified by containing a "?" character
5. **Empty Rows**: Empty rows are skipped automatically

### Column Structure

```
Row | Column C Content
----|------------------
1   | Survey description text
2   | More description
3   | consent                    ← Trigger word
4   | Consent form text
5   | What is your name?         ← Text question
6   | ___________________
7   | What is your age?          ← Multiple choice
8   | Under 18
9   | 18-25
10  | 26-35
11  | Over 35
```

## Question Types & Examples

### 1. Description Content

**Detection**: Any content before "consent" trigger word

**Input Format**:

```
This survey aims to understand customer satisfaction.
Please take 5-10 minutes to complete all questions.
Your responses will remain confidential.
```

**Output**:

-   Formatted as regular text in template
-   No question IDs assigned
-   Word wrapping enabled

---

### 2. Consent

**Detection**: Exact word "consent" (case insensitive)

**Input Format**:

```
consent
By participating in this survey, you agree to our terms and conditions.
```

**Output**:

-   "Consent" header (bold)
-   Consent text below
-   Switches processing mode to questions

---

### 3. Section Headings

**Detection**: Text with underline formatting in Excel

**Input Format**:

-   Apply underline formatting to the cell in Excel
-   Example: `DEMOGRAPHIC INFORMATION` (underlined)

**Output**:

-   Converted to uppercase
-   Bold and underlined formatting
-   Type: "Section Heading"
-   No question ID assigned

---

### 4. Text Questions

**Detection**:

-   Contains "?" character
-   Next row contains "\_\_" (underscores)
-   Row after that is empty

**Input Format**:

```
What is your full name?
_____________________

What is your email address?
_____________________

```

**Output**:

-   Question gets ID: `PREFIX_01`, `PREFIX_02`, etc.
-   Type: "TextQuestion"
-   Input placeholder included

---

### 5. Tickboxes (Single Choice)

**Detection**:

-   Contains "?" character
-   Followed by options (non-empty rows)
-   Does NOT contain "select" keyword

**Input Format**:

```
What is your preferred contact method?
Email
Phone
Text message
Post

```

**Output**:

-   Question ID: `CATPREFIX_01`
-   Type: "Tickboxes"
-   Each option gets option number (1, 2, 3, 4)
-   Options do NOT get individual IDs

---

### 6. Multiple Choice

**Detection**:

-   Contains "?" character
-   Contains "select" keyword (case insensitive)
-   Followed by options

**Input Format**:

```
Please select all that apply - which social media do you use?
Facebook
Instagram
Twitter
LinkedIn
TikTok

```

**Output**:

-   Question ID: `CATPREFIX_02`
-   Type: "Tickboxes"
-   Each option gets individual ID: `CATPREFIX_02a`, `CATPREFIX_02b`, etc.
-   Option numbers: 1, 2, 3, 4, 5

---

### 7. Instructions (Complex Questions)

**Detection**:

-   Contains "?" character
-   Next row has **bold** formatting
-   Multiple bold rows followed by regular options

**Input Format**:

```
Please rate each service on a scale of 1-5?
**Customer Service**        ← Bold formatting
**Product Quality**         ← Bold formatting
**Delivery Speed**          ← Bold formatting
Excellent
Good
Average
Poor
Very Poor

```

**Output**:

-   Main question type: "Instructions"
-   Sub-questions get IDs: `CATPREFIX_03_01`, `CATPREFIX_03_02`, `CATPREFIX_03_03`
-   Each sub-question type: "Tickboxes"
-   Same options applied to each sub-question

## Output Format

### Question Identification

| Question Type   | Catalogue ID Format | Question ID Format | Type Column     |
| --------------- | ------------------- | ------------------ | --------------- |
| Text Question   | `CATPREFIX_01`      | `QPREFIX_01`       | TextQuestion    |
| Tickboxes       | `CATPREFIX_02`      | `QPREFIX_02`       | Tickboxes       |
| Multiple Choice | `CATPREFIX_03a`     | `QPREFIX_03a`      | Tickboxes       |
| Instructions    | `CATPREFIX_04_01`   | `QPREFIX_04_01`    | Tickboxes       |
| Section Heading | (none)              | (none)             | Section Heading |

### Template Columns

The processor writes to specific columns in your template:

-   **Column H (8)**: Question/content text
-   **Column I (9)**: Catalogue ID
-   **Column J (10)**: Question ID
-   **Column K (11)**: Option number
-   **Column X (24)**: Question type

## Architecture

The application is built with a modular architecture using these key components:

### Core Classes

```
📁 Models/
├── ProcessingConfig.cs       - Configuration settings
├── QuestionData.cs          - Question data structure
└── ProcessingResult.cs      - Processing results

📁 Enums/
├── QuestionType.cs          - Question type enumeration
└── ProcessingMode.cs        - Processing mode states

📁 Interfaces/
└── IQuestionProcessor.cs    - Question processor interface

📁 Services/
├── QuestionTypeDetector.cs  - Detects question types
├── SurveyDataExtractor.cs   - Extracts data from Excel
├── SurveyProcessor.cs       - Main processing logic
└── ProcessorFactory.cs      - Creates appropriate processors

📁 Processors/
├── DescriptionProcessor.cs
├── ConsentProcessor.cs
├── SectionHeadingProcessor.cs
├── TextQuestionProcessor.cs
├── TickboxesProcessor.cs
├── MultipleChoiceProcessor.cs
└── InstructionsProcessor.cs

📁 Utilities/
├── Constants.cs             - Application constants
├── Logger.cs               - Logging functionality
└── ConfigValidator.cs      - Input validation
```

## Usage

### WinForms Application (Windows)

1. **Launch** the application
2. **Enter** survey details:
    - Survey Title
    - Input Column (A-Z)
    - Question ID Prefix (e.g., "MMSMQ")
    - Catalogue ID Prefix (e.g., "MMSMQ25")
3. **Select** input Excel file
4. **Select** template Excel file
5. **Choose** output folder
6. **Click** "Process Files"

### CLI Application (Mac/Linux/Windows)

```bash
dotnet run "Survey Title" C MMSMQ MMSMQ25 input.xlsx template.xlsx output_folder
```

**Parameters**:

-   `"Survey Title"` - Title for the survey (use quotes if it contains spaces)
-   `C` - Column letter to read from input file
-   `MMSMQ` - Short question ID prefix
-   `MMSMQ25` - Long catalogue ID prefix
-   `input.xlsx` - Source Excel file
-   `template.xlsx` - Template Excel file
-   `output_folder` - Output directory (optional)

## Best Practices

### Input Preparation

1. **Clean Data**: Remove unnecessary empty rows between content
2. **Consistent Formatting**: Use Excel formatting (bold, underline) consistently
3. **Clear Questions**: Make sure questions end with "?"
4. **Logical Flow**: Organize content in the expected sequence (description → consent → questions)
5. **Option Grouping**: Keep question options together with no empty rows between them

### Question Writing

-   **Text Questions**: Follow with underscores `___` on the next row
-   **Multiple Choice**: Include "select" keyword for options that need individual IDs
-   **Instructions**: Use bold formatting for sub-questions
-   **Section Headings**: Use underline formatting in Excel

### Testing

1. **Start Small**: Test with a few questions first
2. **Check Output**: Verify question IDs and types are assigned correctly
3. **Iterate**: Refine input formatting based on output results

## Troubleshooting

### Common Issues

| Problem                | Solution                                          |
| ---------------------- | ------------------------------------------------- |
| Questions not detected | Ensure they contain "?" character                 |
| Wrong question type    | Check formatting (bold/underline) and keywords    |
| Missing options        | Verify no empty rows between question and options |
| No consent detected    | Ensure "consent" is exact word, case insensitive  |

### Validation Errors

The application validates all inputs before processing:

-   ✅ Survey title is not empty
-   ✅ Input column is valid (A-Z)
-   ✅ Files exist and are accessible
-   ✅ Output folder exists

---

_For technical support or feature requests, refer to the application logs for detailed processing information._
