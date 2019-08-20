package com.lisnail.util;

import lombok.Getter;
import lombok.Setter;
import java.util.Date;

@Getter
@Setter
public class ExcelVo {
    private Integer id;
    private Date date;
    private Integer requestCount;
    private Integer projectCount;
    private Integer statusChangeCount;
    private Integer requestFailCount;
    private Integer projectFailCount;
    private Integer repeatReportCount;
    private Integer repeatReportProjectCount;
    private Integer supplementReportCount;
    private Integer supplementDetails;
    private Integer changeRequestCount;
    private Integer changeStatusCount;
    private Integer changeRepeatCount;
    private Integer changStausCount;
    private Integer changeSupplementCount;
    private Integer changeSupplementDetails;
    private String remark;
}
