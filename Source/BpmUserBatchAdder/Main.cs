using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System;
using System.Configuration;
using System.Data.SqlClient;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = NetOffice.ExcelApi;

namespace BpmUserBatchAdder {
    public class Main : IExcelAddIn {
        public void AutoOpen() {
            Globals.app = new Excel.Application(null, ExcelDnaUtil.Application);
            Globals.app.WorkbookActivateEvent += WorkbookActivateEvent;
            Globals.app.SheetActivateEvent += WorksheetActivateEvent;
            Globals.app.WorkbookBeforeCloseEvent += WorkbookBeforeCloseEvent;
        }

        private void WorkbookBeforeCloseEvent(Excel.Workbook Wb, ref bool Cancel) {
            if (Globals.sheet.Name == "BPM批次新增使用者") {
                Globals.book.Saved = true;
            }
        }

        void WorksheetActivateEvent(NetOffice.COMObject Sh) {
            Globals.sheet = (Excel.Worksheet)Sh;
        }

        void WorkbookActivateEvent(Excel.Workbook Wb) {
            Globals.book = Wb;
            Globals.sheet = (Excel.Worksheet)Wb.ActiveSheet;
        }

        public void AutoClose() {

        }
    }

    [ComVisible(true)]
    public class RibbonController : ExcelRibbon {
        public override string GetCustomUI(string RibbonID) {
            string menu = @"<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
                               <ribbon>
                                   <tabs>
                                       <tab id='tab1' label='{0}'>
                                           <group id='group1' label='{1}'>
                                               <button id='button_NewFile' size='large' label='{3}' onAction='button_NewFile_Click' getImage='GetImage' />
                                               <button id='button_ExcelToDatabase' size='large' label='{2}' onAction='button_ExcelToDatabase_Click' getImage='GetImage' />
                                           </group >
                                       </tab>
                                   </tabs>
                               </ribbon>
                            </customUI>";
            menu = menu.FormatWithArgs("BpmUserBatchAdder", "寫入資料庫", "執行", "開啟新增介面");

            return menu;
        }

        public void button_ExcelToDatabase_Click(IRibbonControl control) {
            try {
                if (Globals.sheet.Name != "BPM批次新增使用者") {
                    MessageBox.Show("請先切換到新增介面(或按[開啟新增介面]按鈕切換)");

                    return;
                }

                if (Globals.sheet.UsedRange.Rows.Count == 1) {
                    MessageBox.Show("無任何資料可新增");

                    return;
                }

                Globals.sheet.UsedRange.ClearComments();

                using (var conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString)) {
                    conn.Open();

                    for (var currentRowIndex = 2; currentRowIndex <= Globals.sheet.UsedRange.Rows.Count; currentRowIndex++) {
                        var transaction = conn.BeginTransaction();
                        var command = new SqlCommand();
                        command.Connection = conn;
                        command.Transaction = transaction;

                        var userIdCell = Globals.sheet.Cells[currentRowIndex, 1];
                        var userNameCell = Globals.sheet.Cells[currentRowIndex, 2];
                        var unitIdCell = Globals.sheet.Cells[currentRowIndex, 3];
                        var functionNameCell = Globals.sheet.Cells[currentRowIndex, 4];
                        var approvalLevelCell = Globals.sheet.Cells[currentRowIndex, 5];
                        var emailAccountCell = Globals.sheet.Cells[currentRowIndex, 6];
                        var resultCell = Globals.sheet.Cells[currentRowIndex, 7];

                        var userId = userIdCell.Value2.HasValue() == true ? userIdCell.Value2.ToString().Trim() : "";
                        var userName = userNameCell.Value2.HasValue() == true ? userNameCell.Value2.ToString().Trim() : "";
                        var unitId = unitIdCell.Value2.HasValue() == true ? unitIdCell.Value2.ToString().Trim() : "";
                        var functionName = functionNameCell.Value2.HasValue() == true ? functionNameCell.Value2.ToString().Trim() : "";
                        var approvalLevel = approvalLevelCell.Value2.HasValue() == true ? approvalLevelCell.Value2.ToString().Trim() : "";
                        var emailAccount = emailAccountCell.Value2.HasValue() == true ? emailAccountCell.Value2.ToString().Trim() : "";

                        if (userId.HasValue() == false) {
                            userIdCell.AddFailComment("此欄位必須有值");
                        }
                        if (userName.HasValue() == false) {
                            userNameCell.AddFailComment("此欄位必須有值");
                        }
                        if (unitId.HasValue() == false) {
                            unitIdCell.AddFailComment("此欄位必須有值");
                        }
                        if (functionName.HasValue() == false) {
                            functionNameCell.AddFailComment("此欄位必須有值");
                        }
                        if (approvalLevel.HasValue() == false) {
                            approvalLevelCell.AddFailComment("此欄位必須有值");
                        }
                        if (emailAccount.HasValue() == false) {
                            emailAccountCell.AddFailComment("此欄位必須有值");
                        }
                        if (userId.HasValue() == false
                            || userName.HasValue() == false
                            || unitId.HasValue() == false
                            || functionName.HasValue() == false
                            || approvalLevel.HasValue() == false
                            || emailAccount.HasValue() == false) {
                            continue;
                        }

                        var userOId = Guid.NewGuid().ToString("N");
                        var employeeOId = Guid.NewGuid().ToString("N");
                        var functionsOId = Guid.NewGuid().ToString("N");

                        var workflowServerOIdResult = Database.GetDataTable(command, "SELECT OID FROM WorkflowServer WHERE isDefault = 1;");
                        if (workflowServerOIdResult.Rows.Count == 0) {
                            unitIdCell.AddFailComment(unitId + "找不到工作主機OID");
                            resultCell.AddFailComment("失敗");

                            continue;
                        }
                        var workflowServerOId = workflowServerOIdResult.Rows[0][0].ToString();

                        var organizationUnitOIdResult = Database.GetDataTable(command, "SELECT OID FROM OrganizationUnit WHERE id = '{0}';".FormatWithArgs(unitId));
                        if (organizationUnitOIdResult.Rows.Count == 0) {
                            unitIdCell.AddFailComment(unitId + "找不到部門OID");
                            resultCell.AddFailComment("失敗");

                            continue;
                        }
                        var organizationUnitOId = organizationUnitOIdResult.Rows[0][0].ToString();

                        var organizationOIdResult = Database.GetDataTable(command, "SELECT organizationOID FROM OrganizationUnit WHERE id = '{0}';".FormatWithArgs(unitId));
                        if (organizationOIdResult.Rows.Count == 0) {
                            unitIdCell.AddFailComment(unitId + "找不到組織OID");
                            resultCell.AddFailComment("失敗");

                            continue;
                        }
                        var organizationOId = organizationOIdResult.Rows[0][0].ToString();

                        var calendarOIdResult = Database.GetDataTable(command, "SELECT OID FROM WorkCalendar where containerOID = '{0}';".FormatWithArgs(organizationOId));
                        if (calendarOIdResult.Rows.Count == 0) {
                            unitIdCell.AddFailComment(unitId + "找不到行事曆OID");
                            resultCell.AddFailComment("失敗");

                            continue;
                        }
                        var calendarOId = calendarOIdResult.Rows[0][0].ToString();

                        var approvalLevelOIdResult = Database.GetDataTable(command, "SELECT OID FROM FunctionLevel WHERE organizationOID = '{1}' AND functionLevelName LIKE '{0}%';".FormatWithArgs(approvalLevel, organizationOId));
                        if (approvalLevelOIdResult.Rows.Count == 0) {
                            approvalLevelCell.AddFailComment(approvalLevel + "找不到OID");
                            resultCell.AddFailComment("失敗");

                            continue;
                        }
                        var approvalLevelOId = approvalLevelOIdResult.Rows[0][0].ToString();

                        var functionDefinitionOIdResult = Database.GetDataTable(command, "SELECT OID FROM FunctionDefinition WHERE organizationOID = '{1}' AND functionDefinitionName = '{0}';".FormatWithArgs(functionName, organizationOId));
                        if (functionDefinitionOIdResult.Rows.Count == 0) {
                            functionNameCell.AddFailComment(functionName + "找不到OID");
                            resultCell.AddFailComment("失敗");

                            continue;
                        }
                        var functionDefinitionOId = functionDefinitionOIdResult.Rows[0][0].ToString();

                        try {
                            var insertUsersSql = @"
                        INSERT INTO Users VALUES (
                            '{0}',
                            '{1}', --員工工號
                            '{2}', --姓名
                            1,
                            'mlCCiMVj8+lN5SjYg0g3bp2WzdA=', --密碼(用系統預設值)
                            NULL,
                            '{3}',
                            'DEFAULT', --系統預設認證方式:LDAP
                            '{4}@usuntek.com', --信箱
                            NULL,
                            '{5}', --工作主機:wfs1
                            1,
                            NULL,
                            NULL,
                            0,
                            0,
                            '{6}', --LDAP帳號
                            NULL,
                            'zh_TW',
                            NULL,
                            1,
                            2
                        );
                        ".FormatWithArgs(userOId, userId, userName, calendarOId, emailAccount, workflowServerOId, userId);

                            var isInsertUsersOk = Database.RunSql(command, insertUsersSql);

                            var insertEmployeeSql = @"
                        INSERT INTO Employee VALUES (
                            '{0}'
                            ,'{1}'
                            ,'{2}'
                            ,'{3}'
                            ,1
                            ,NULL
                        );
                        ".FormatWithArgs(employeeOId, userId, organizationOId, userOId);

                            var isInsertEmployeeOk = Database.RunSql(command, insertEmployeeSql);

                            var insertFunctionsSql = @"
                        INSERT INTO Functions VALUES (
                            '{0}'
                            ,1
                            ,'{1}'
                            ,'{2}'
                            ,'{3}'
                            ,'{4}'
                            ,NULL
                            ,1 --為主部門
                        );
                        ".FormatWithArgs(functionsOId, approvalLevelOId, functionDefinitionOId, userOId, organizationUnitOId);

                            var isInsertFunctionsOk = Database.RunSql(command, insertFunctionsSql);

                            if (isInsertUsersOk == false || isInsertEmployeeOk == false || isInsertFunctionsOk == false) {
                                transaction.Rollback();

                                resultCell.AddFailComment("失敗");
                            }
                            else {
                                transaction.Commit();

                                resultCell.AddSuccessComment("成功");
                            }
                        }
                        catch (Exception ex) {
                            transaction.Rollback();

                            resultCell.AddFailComment(ex.Message);
                        }
                    }

                    MessageBox.Show("執行完成");
                }
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        public void button_NewFile_Click(IRibbonControl control) {
            Globals.app.Workbooks.Add();
            Globals.sheet.Cells[1, 1].Value2 = "工號";
            Globals.sheet.Cells[1, 1].Interior.ColorIndex = 15;
            Globals.sheet.Cells[1, 2].Value2 = "姓名";
            Globals.sheet.Cells[1, 2].Interior.ColorIndex = 15;
            Globals.sheet.Cells[1, 3].Value2 = "部門代號";
            Globals.sheet.Cells[1, 3].Interior.ColorIndex = 15;
            Globals.sheet.Cells[1, 4].Value2 = "職稱";
            Globals.sheet.Cells[1, 4].Interior.ColorIndex = 15;
            Globals.sheet.Cells[1, 5].Value2 = "核決權限";
            Globals.sheet.Cells[1, 5].Interior.ColorIndex = 15;
            Globals.sheet.Cells[1, 6].Value2 = "e-mail帳號";
            Globals.sheet.Cells[1, 6].Interior.ColorIndex = 15;
            Globals.sheet.Cells[1, 7].Value2 = "新增結果";
            Globals.sheet.Cells[1, 7].Interior.ColorIndex = 15;
            Globals.sheet.Name = "BPM批次新增使用者";
            Globals.sheet.ScrollArea = "$A2:$F1048576";
            Globals.app.ActiveWindow.SplitColumn = 0;
            Globals.app.ActiveWindow.SplitRow = 1;
            Globals.app.ActiveWindow.FreezePanes = true;
            Globals.sheet.EnableSelection = Excel.Enums.XlEnableSelection.xlUnlockedCells;
            Globals.sheet.Cells[2, 1].Select();
        }

        public Bitmap GetImage(IRibbonControl control) {
            switch (control.Id) {
                case "button_ExcelToDatabase":
                    return new Bitmap(BpmUserBatchAdder.Properties.Resources.ExcelToDatabase);

                case "button_NewFile":
                    return new Bitmap(BpmUserBatchAdder.Properties.Resources.NewFile);

                default:
                    return null;
            }
        }
    }
}