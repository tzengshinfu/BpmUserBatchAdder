using Excel = NetOffice.ExcelApi;

namespace BpmUserBatchAdder {
    public static partial class ExtensionMethod {
        public static void BeginUpdate(this Excel.Application app) {
            app.ScreenUpdating = false;
            app.DisplayStatusBar = false;
            app.Calculation = Excel.Enums.XlCalculation.xlCalculationManual;
            app.EnableEvents = false;
        }

        public static void EndUpdate(this Excel.Application app) {
            app.ScreenUpdating = true;
            app.DisplayStatusBar = true;
            app.Calculation = Excel.Enums.XlCalculation.xlCalculationAutomatic;
            app.EnableEvents = true;
        }

        public static void AddFailComment(this Excel.Range rng, string text) {
            rng.Interior.ColorIndex = 3;
            rng.AddComment(text);
        }

        public static void AddSuccessComment(this Excel.Range rng, string text) {
            rng.Interior.ColorIndex = 43;
            rng.AddComment(text);
        }
    }
}