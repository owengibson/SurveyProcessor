namespace SurveyProcessor
{
    // Logger utility (adapted for WinForms)
    public static class Logger
    {
        private static RichTextBox logTextBox;
        
        public static void SetLogTextBox(RichTextBox textBox)
        {
            logTextBox = textBox;
        }
        
        public static void Log(string message)
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            string logMessage = $"[{timestamp}] {message}\n";
            
            if (logTextBox != null)
            {
                logTextBox.AppendText(logMessage);
                logTextBox.ScrollToCaret();
                Application.DoEvents(); // Allow UI to update
            }
        }
    }
}