package kaitou.office.excel;

import kaitou.office.excel.common.SysCode;
import kaitou.office.excel.domain.Application;
import kaitou.office.excel.util.ExcelUtils;
import kaitou.office.excel.util.PropertyUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.joda.time.DateTime;

import java.io.File;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

import static kaitou.office.excel.util.Transformer.transform2Xls;

/**
 * 测试.
 * User: 赵立伟
 * Date: 2015/1/10
 * Time: 21:20
 */
public class Demo {

    public static void main(String[] args) throws Exception {
        String workspacePath = PropertyUtils.getPath("workspace");
        File workspace = new File(workspacePath);
        File[] appFiles = workspace.listFiles();
        String completePath = PropertyUtils.getPath("complete");
        List<Object[]> rowDataList = new ArrayList<Object[]>();
        List<Object[]> sheetDataList = new ArrayList<Object[]>();
        File log = new File(PropertyUtils.getPath("output") + PropertyUtils.getValue("log_name"));
        DecimalFormat df = new DecimalFormat("00000");
        DateTime now = new DateTime();
        String templateName = PropertyUtils.getValue("template_name");
        List<Application> applications = new ArrayList<Application>();
        for (int i = 0; i < appFiles.length; i++) {
            File appFile = appFiles[i];
            Workbook workbook = ExcelUtils.create(appFile);
            int numberOfSheets = workbook.getNumberOfSheets();
            for (int j = 0; j < numberOfSheets; j++) {
                Sheet sheet = workbook.getSheetAt(j);
                Application application = new Application();
                application.fill(sheet);
                SysCode.Code code = SysCode.Code.getCode(application.getModels().substring(0, 1));
                String lastWarrantyCard = ExcelUtils.getLastRowCellStrValue(log, code.getCardNoPref(), 3, SysCode.CellType.STRING);
                long warrantyCardIndex = 0;
                if (lastWarrantyCard != null && !"".equals(lastWarrantyCard.trim())) {
                    warrantyCardIndex = Long.valueOf(lastWarrantyCard.substring(5));
                }
                application.setWarrantyCard(code.getCardNoPref() + "-A" + df.format(++warrantyCardIndex));
                application.setApplyDate(now.toString("yyyy/MM/dd"));
                DateTime installedDate = application.convertInstalledDate();
                DateTime endDate = installedDate.plusDays(364);
                if (endDate.isBeforeNow()) {
                    application.setStatus("过保");
                }
                application.setEndDate(endDate.toString("yyyy/MM/dd"));
                application.setAllModels(code.getModels());
                application.setInitData(application.getInitData() + code.getReadUnit());
//                System.out.println(application.toString());
                applications.add(application);
                rowDataList.add(application.getAllRowData());
                sheetDataList.add(application.getRowData());
                ExcelUtils.add2Sheet(log, code.getCardNoPref(), sheetDataList);
                sheetDataList.clear();
            }
            appFile.renameTo(new File(completePath + appFile.getName()));
        }
        ExcelUtils.add2Sheet(log, "汇总", rowDataList);
        transform2Xls(applications, templateName);
    }
}
