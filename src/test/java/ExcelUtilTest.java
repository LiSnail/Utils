import com.lisnail.excelutil.ExcelUtil;
import org.junit.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public class ExcelUtilTest {

    @Test
    public void Test() throws IOException {

        /* 根据传入的表格数据，生成表格。但是需要注意以下几点。
        1.假如你的类型传的类型不是基本类型，那么 headerWords这个参数就填null。
        如果传入的类型是包装类型，那headerWords传入的就是包装类的字段。注意表头的字段顺序要和 hearderWords的顺序一致。
        2.合并单元格。注意复杂的表头生成，你传入的headers的顺序要和 headerNum的顺序一样。
        */
        String[] headers1 = {"序号", "日期", "成功报送情况", "成功报送情况", "成功报送情况", "项目异常情况", "项目异常情况", "项目异常情况", "项目异常情况", "项目异常情况", "项目异常情况"
                , "项目状态变更异常情况", "项目状态变更异常情况", "项目状态变更异常情况", "项目状态变更异常情况", "项目状态变更异常情况", "项目状态变更异常情况", "备注"};
        String[] headnum1 = {"0,0,0,0", "0,0,0,0", "0,0,2,4", "0,0,5,10", "0,0,11,16", "0,0,0,0"};

        String[] headers2 = {"", "", "报送请求总数(项目+项目状态变更)", "项目总数", "状态变更总数", "失败请求总数", "失败项目总数", "重复报送项目个数", "重复报送项目次数", "补报项目数", "补报情况说明", "失败请求总数", "失败状态变更总数", "重复报送状态变更个数", "重复状态变更次数", "补报的状态变更数", "补报情况说明", ""};
        String[] headnum2 = {"0,1,0,0", "0,1,1,1", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,0,0,0", "0,1,17,17"};

        ExcelUtil<String> excelUtil = new ExcelUtil<String>();
        //String[] headerWords = {"id", "date", "requestCount", "projectCount", "statusChangeCount",
        //        "requestFailCount", "projectFailCount", "repeatReportCount", "repeatReportProjectCount",
        //        "supplementReportCount", "supplementDetails", "changeRequestCount", "changeStatusCount",
        //        "changeRepeatCount", "changStausCount", "changeSupplementCount", "changeSupplementDetails",
        //        "remark"
        //};

        //表头

        List<String[]> headers = new ArrayList();
        headers.add(headers1);
        headers.add(headers2);
        List<String[]> headnums = new ArrayList();
        headnums.add(headnum1);
        headnums.add(headnum2);
        //List<ExcelVo> resul = new ArrayList<ExcelVo>();
        List<String> resul = new ArrayList<String>();
        for (int i = 0; i < 4; i++) {

            resul.add(String.valueOf(i));

            //ExcelVo excelVo = new ExcelVo();
            //excelVo.setId(i + 1);
            //excelVo.setDate(new Date());
            //excelVo.setRequestCount(3);
            //excelVo.setProjectCount(4);
            //excelVo.setStatusChangeCount(5);
            //excelVo.setRequestFailCount(6);
            //excelVo.setProjectFailCount(7);
            //excelVo.setRepeatReportCount(8);
            //excelVo.setRepeatReportProjectCount(9);
            //excelVo.setSupplementReportCount(10);
            //excelVo.setSupplementDetails(11);
            //excelVo.setChangeRequestCount(12);
            //excelVo.setChangeStatusCount(13);
            //excelVo.setChangeRepeatCount(14);
            //excelVo.setChangStausCount(15);
            //excelVo.setChangeSupplementCount(16);
            //excelVo.setChangeSupplementDetails(17);
            //resul.add(excelVo);
        }
        InputStream inputStream = excelUtil.excelContent(4, "上报情况", headers, headnums, resul, null);
        excelUtil.export(inputStream, "E:\\", "每日上报情况");

    }
}