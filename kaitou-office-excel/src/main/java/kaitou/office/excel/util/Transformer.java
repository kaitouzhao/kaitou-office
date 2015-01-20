package kaitou.office.excel.util;

import kaitou.office.excel.domain.Application;
import net.sf.jxls.transformer.XLSTransformer;
import org.apache.poi.ss.usermodel.Workbook;
import org.joda.time.DateTime;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static kaitou.office.excel.util.PropertyUtils.getPath;

/**
 * 转换器.
 * User: 赵立伟
 * Date: 2015/1/11
 * Time: 10:28
 */
public abstract class Transformer {
    /**
     * 根据模板与数据转换excel
     *
     * @param applications 申请集合
     * @param templateName 模板名
     */
    public static void transform2Xls(List<Application> applications, String templateName) throws Exception {
        String templatePath = getPath("template") + templateName;
//        String targetPath = getPath("output") + new DateTime().toString("yyyy_MM_dd") + ".xls";
//        multipleTransform(applications, templatePath, targetPath);
        Workbook wb;
        for (Application application : applications) {
            String targetPath = getPath("output") + application.getNewSheetName() + ".xls";
            wb = transform2Workbook(application, templatePath);
            OutputStream os = new BufferedOutputStream(new FileOutputStream(targetPath));
            wb.write(os);
            os.close();
        }
    }

    /**
     * 单个转换
     *
     * @param application  申请
     * @param templatePath 模板
     * @return 表
     */
    private static Workbook transform2Workbook(Application application, String templatePath) throws Exception {
        List<String> sheetNames = new ArrayList<String>();
        List<Map<String, Object>> fieldMap = new ArrayList<Map<String, Object>>();
        sheetNames.add(application.getNewSheetName());
        fieldMap.add(application.field2Map());
        File templateFile = new File(templatePath);
        InputStream is = new BufferedInputStream(new FileInputStream(templateFile));
        XLSTransformer transformer = new XLSTransformer();
        return transformer.transformMultipleSheetsList(is, fieldMap, sheetNames, "results_JxLsC_", new HashMap(), 0);
    }

    /**
     * 多个转换
     *
     * @param applications 申请集合
     * @param templatePath 模板路径
     * @param targetPath   目标路径
     */
    private static void multipleTransform(List<Application> applications, String templatePath, String targetPath) throws Exception {
        List<Map<String, Object>> fieldMaps = new ArrayList<Map<String, Object>>();
        for (Application application : applications) {
            fieldMaps.add(application.field2Map());
        }
        Map<String, Object> beans = new HashMap<String, Object>();
        beans.put("results_JxLsC_", applications);
        XLSTransformer transformer = new XLSTransformer();
        transformer.groupCollection("results_JxLsC_.newSheetName");
        transformer.transformXLS(templatePath, beans, targetPath);
    }
}
