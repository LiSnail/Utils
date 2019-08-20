import com.lisnail.util.ExcelUtil;
import org.junit.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ExcelUtilTest {

    @Test
    public void Test() throws IOException {
        String[] headers1 = {"序号", "日期", "成功报送情况", "成功报送情况", "成功报送情况", "项目异常情况", "项目异常情况", "项目异常情况", "项目异常情况", "项目异常情况", "项目异常情况"
                , "项目状态变更异常情况", "项目状态变更异常情况", "项目状态变更异常情况", "项目状态变更异常情况", "项目状态变更异常情况", "项目状态变更异常情况", "备注"};
        String[] headnum1 = {"0,0,0,0", "0,0,0,0", "0,0,2,4", "0,0,5,10", "0,0,11,16", "0,0,0,0"};
        String[] headers2 = {"", "", "报送请求总数(项目+项目状态变更)", "项目总数", "状态变更总数", "失败请求总数", "失败项目总数", "重复报送项目个数", "重复报送项目次数", "补报项目数", "补报情况说明", "失败请求总数", "失败状态变更总数", "重复报送状态变更个数", "重复状态变更次数", "补报的状态变更数", "补报情况说明", ""};
        String[] headnum2 = {"0,1,0,0", "0,1,1,1", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,1,17,17"};
//        String[] headerWords = {
//                "id",
//                "date",
//                "requestCount",
//                "projectCount",
//                "statusChangeCount",
//                "requestFailCount",
//                "projectFailCount",
//                "repeatReportCount",
//                "repeatReportProjectCount",
//                "supplementReportCount",
//                "supplementDetails",
//                "changeRequestCount",
//                "changeStatusCount",
//                "changeRepeatCount",
//                "changStausCount",
//                "changeSupplementCount",
//                "changeSupplementDetails",
//                "remark"
//        };

        ExcelUtil<List<String>> excelUtil = new ExcelUtil<List<String>>();
        //表头
        List<String[]> headers = new ArrayList();
        headers.add(headers1);
        headers.add(headers2);
        List<String[]> headnums = new ArrayList();
        headnums.add(headnum1);
        headnums.add(headnum2);
//        List<ExcelVo> resul = new ArrayList<ExcelVo>();
        List<List<String>> resul = new ArrayList<List<String>>();
        for (int i = 0; i < 4; i++) {
            List<String> list = new ArrayList<String>();
            list.add(String.valueOf(i));
            resul.add(list);

//            ExcelVo excelVo = new ExcelVo();
//            excelVo.setId(i+1);
//            excelVo.setDate(new Date());
//            excelVo.setRequestCount(3);
//            excelVo.setProjectCount(4);
//            excelVo.setStatusChangeCount(5);
//            excelVo.setRequestFailCount(6);
//            excelVo.setProjectFailCount(7);
//            excelVo.setRepeatReportCount(8);
//            excelVo.setRepeatReportProjectCount(9);
//            excelVo.setSupplementReportCount(10);
//            excelVo.setSupplementDetails(11);
//            excelVo.setChangeRequestCount(12);
//            excelVo.setChangeStatusCount(13);
//            excelVo.setChangeRepeatCount(14);
//            excelVo.setChangStausCount(15);
//            excelVo.setChangeSupplementCount(16);
//            excelVo.setChangeSupplementDetails(17);
//            resul.add(excelVo);
        }
        InputStream inputStream = excelUtil.excelContent(4, "上报情况", headers, headnums, resul, null);
        excelUtil.export(inputStream, "E:\\", "每日上报情况");

    }
}