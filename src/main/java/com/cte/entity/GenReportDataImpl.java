/*
package com.cte.entity;

import com.mocard.api.GenReportData;
import com.mocard.common.enums.CardStatus;
import com.mocard.common.enums.TransType;
import com.mocard.common.utils.ExcelUtils;
import com.mocard.db.entity.*;
import com.mocard.db.service.*;
import com.mocard.db.vo.*;
import org.apache.commons.lang3.EnumUtils;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.RandomStringUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.commons.lang3.time.DateUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

*/
/**
 * Created by User on 2017/12/8.
 *//*

@Component("genReportDataImpl")
public class GenReportDataImpl implements GenReportData {
    private static final Logger logger = LogManager.getLogger();

    @Value("${manual_flag}")
    private String manualFlag;

    @Value("${begin_date}")
    private String beginDate;

    @Value("${end_date}")
    private String endDate;

    @Value("${report_path}")
    private String reportPath;

    @Autowired
    private TJobLogService tJobLogService;
    @Autowired
    private TMerchantService tMerchantService;
    @Autowired
    private TBillFlowService tBillFlowService;
    @Autowired
    private TInstructionDayStatisticsService tInstructionDayStatisticsService;
    @Autowired
    private TInstructionMonthStatisticsService tInstructionMonthStatisticsService;
    @Autowired
    private TInstructionLogService tInstructionLogService;
    @Autowired
    private TCardService tCardService;
    @Autowired
    private TMerchantReportService tMerchantReportService;

    @Override
    public String genDayReportData() {
        logger.info("genDayReportData begin......");
        TJobLog tJobLog = new TJobLog();

        String today = DateFormatUtils.format(new Date(), "yyyyMMddHHmmss");
        String randomNum = RandomStringUtils.randomNumeric(3);
        tJobLog.setJobSerial(today + randomNum);
        tJobLog.setJobType("SYS_CUT");
        tJobLog.setJobBurstType("SCHEDULED");
        tJobLog.setJobStatus("00");
        //1.登记事务流水
        tJobLogService.getMapper().insert(tJobLog);
        //2.查询商户表
        TMerchant tMerchant = new TMerchant();
        List<TMerchant> tMerchantList = tMerchantService.getMapper().select(tMerchant);
        //给订单表查询条件赋值
        TBillFlow tBillFlow = new TBillFlow();
        if ("1".equals(manualFlag)) {
            tBillFlow.setBillDateBegin(beginDate);
            tBillFlow.setBillDateEnd(endDate);
        } else {
            String yesterday = DateFormatUtils.format(DateUtils.addDays(new Date(), -1), "yyyyMMdd");
            tBillFlow.setBillDateBegin(yesterday);
            tBillFlow.setBillDateEnd(yesterday);
        }
        //循环获取商户号
        for (TMerchant tMerchant1 : tMerchantList) {
            logger.info("start {} merchant day statistics....", tMerchant1.getMerchantId());
            tBillFlow.setTableName("t_bill_flow_m" + tMerchant1.getMerchantId());
            List<TBillFlow> tBillFlowList = tBillFlowService.getMapper().selectGroup(tBillFlow);

            for (TBillFlow tBillFlow1 : tBillFlowList) {
                TInstructionDayStatistics tInstructionDayStatistics = new TInstructionDayStatistics();
                tInstructionDayStatistics.setTableName("t_instruction_day_statistics_m" + tMerchant1.getMerchantId());
                tInstructionDayStatistics.setMerchantId(tMerchant1.getMerchantId());
                tInstructionDayStatistics.setPosition(tBillFlow1.getBillPosition());
                tInstructionDayStatistics.setStuffId(tBillFlow1.getBillStuff());
                tInstructionDayStatistics.setBusinessType(tBillFlow1.getBillSourceType());

                if (tBillFlow1.getBillSubject() != null || tBillFlow1.getBillSubject() != "") {
                    tInstructionDayStatistics.setBusinessCount(Integer.valueOf(tBillFlow1.getBillSubject()));
                } else {
                    tInstructionDayStatistics.setBusinessCount(0);
                }
                if (tBillFlow1.getBillSubjectCount() != null) {
                    tInstructionDayStatistics.setBusinessSubjectCount(tBillFlow1.getBillSubjectCount());
                } else {
                    tInstructionDayStatistics.setBusinessSubjectCount(0);
                }

                if (tBillFlow1.getReceivableAmount() != null) {
                    tInstructionDayStatistics.setReceivableAmount(tBillFlow1.getReceivableAmount());
                } else {
                    tInstructionDayStatistics.setReceivableAmount(new BigDecimal(0.00));
                }
                if (tBillFlow1.getIncomeAmount() != null) {
                    tInstructionDayStatistics.setIncomeAmount(tBillFlow1.getIncomeAmount());
                } else {
                    tInstructionDayStatistics.setIncomeAmount(new BigDecimal(0.00));
                }
                if (tBillFlow1.getPayableAmount() != null) {
                    tInstructionDayStatistics.setPayableAmount(tBillFlow1.getPayableAmount());
                } else {
                    tInstructionDayStatistics.setPayableAmount(new BigDecimal(0.00));
                }
                if (tBillFlow1.getExpensesAmount() != null) {
                    tInstructionDayStatistics.setExpensesAmount(tBillFlow1.getExpensesAmount());
                } else {
                    tInstructionDayStatistics.setExpensesAmount(new BigDecimal(0.00));
                }
                if (tBillFlow1.getReceivableFeeAmount() != null) {
                    tInstructionDayStatistics.setReceivableFeeAmount(tBillFlow1.getReceivableFeeAmount());
                } else {
                    tInstructionDayStatistics.setReceivableFeeAmount(new BigDecimal(0.00));
                }
                if (tBillFlow1.getDiscountAmount() != null) {
                    tInstructionDayStatistics.setDiscountAmount(tBillFlow1.getDiscountAmount());
                } else {
                    tInstructionDayStatistics.setDiscountAmount(new BigDecimal(0.00));
                }

                tInstructionDayStatistics.setBillDate(tBillFlow1.getBillDate());

                try {
                    tInstructionDayStatistics.setLastBillDate(DateFormatUtils.format(DateUtils.addDays(new SimpleDateFormat("yyyyMMdd").parse(tBillFlow1.getBillDate()), -1), "yyyyMMdd"));
                } catch (ParseException e) {
                    e.printStackTrace();
                }

                tInstructionDayStatistics.setJobSerial(tJobLog.getJobSerial());
                tInstructionDayStatisticsService.getMapper().insertDynamic(tInstructionDayStatistics);
            }
        }

        //3.更新事务流水
        tJobLog.setJobStatus("01");
        tJobLogService.getMapper().updateByPrimaryKey(tJobLog);
        logger.info("genDayReportData finish......");

        //4.生成日报表文件
        genDayReportFile();

        return "交易成功";
    }

//    @Override
//    public String genDayReportFile(){
//        logger.info("genDayReportFile begin......");
//        TJobLog tJobLog = new TJobLog();
//
//        String today = DateFormatUtils.format(new Date(),"yyyyMMddHHmmss");
//        String randomNum = RandomStringUtils.randomNumeric(3);
//        tJobLog.setJobSerial(today + randomNum);
//        tJobLog.setJobType("MER_REPORT");tJobLog.setJobBurstType("SCHEDULED");tJobLog.setJobStatus("00");
//        //1.登记事务流水
//        tJobLogService.getMapper().insert(tJobLog);
//        //2.查询商户表
//        TMerchant tMerchant = new TMerchant();
//        List<TMerchant> tMerchantList = tMerchantService.getMapper().select(tMerchant);
//        //给日统计表查询条件赋值
//        TInstructionDayStatistics tInstructionDayStatistics = new TInstructionDayStatistics();
//        if ("1".equals(manualFlag)){
//            tInstructionDayStatistics.setBillMonthBegin(beginDate);
//            tInstructionDayStatistics.setBillMonthEnd(endDate);
//        }else {
//            String lastDate = DateFormatUtils.format(DateUtils.addDays(new Date(),-1),"yyyyMMdd");
//            tInstructionDayStatistics.setBillMonthBegin(lastDate);
//            tInstructionDayStatistics.setBillMonthEnd(lastDate);
//        }
//        //循环获取商户号
//        for(TMerchant tMerchant1:tMerchantList){
//            logger.info("start {} merchant excel file....",tMerchant1.getMerchantId());
//
//            ExcelUtils excelUtils = null;
//            String filePath = reportPath + "/" + tMerchant1.getMerchantId() + "/";
//            File file = new File(filePath);
//            String[] content ;
//            if(!file.exists() && !file.isDirectory())
//            {
//                file.mkdirs();
//            }
//            //开始赋值
//            String fileName = filePath + "商户日统计报表_M" + tMerchant1.getMerchantId() + "_" + beginDate + ".xls";
//            try {
//                excelUtils = new ExcelUtils(fileName);
//            } catch (FileNotFoundException e) {
//                logger.error("创建文件{}失败,{}",fileName,e);
//                return "创建文件失败";
//            }
//
//            //1.生成第一个sheet 商户日统计表(门店)
//            String title = "售卡/充值/消费/退款/赎回统计表(按门店统计)";
//            String[] queryName = new String[4];
//            queryName[0] = "商户编号 " + tMerchant1.getMerchantId();
//            queryName[1] = "商户名称 " + tMerchant1.getMerchantChineseName();
//            queryName[2] = "起始时间 " + tInstructionDayStatistics.getBillMonthBegin();
//            queryName[3] = "结束时间 " + tInstructionDayStatistics.getBillMonthEnd();
//            String[] headers = new String[] { "序号", "门店编码", "门店名称", "线下售卡笔数", "线下售卡金额",
//            "线上售卡笔数", "线上售卡金额", "线下充值笔数", "线下充值金额", "线上充值笔数", "线上充值金额",
//            "线下消费笔数", "线下消费金额", "线上消费笔数", "线上消费金额", "退款笔数", "退款金额",
//            "赎回笔数", "赎回金额", "冻结笔数", "冻结金额", "解冻笔数", "解冻金额"};
//
//            List<Object[]> contents = new ArrayList<>();
//            int rowId = 0;
//
//            tInstructionDayStatistics.setTableName("t_instruction_day_statistics_m" + tMerchant1.getMerchantId() +
//                    " a left join t_store_m" + tMerchant1.getMerchantId() + " b on a.position=b.store_id");
//            List<ReportDayPositionOfMerchant> reportDayPositionOfMerchantList= tInstructionDayStatisticsService.getMapper().selectGroupFile(tInstructionDayStatistics);
//
//            ReportDayPositionOfMerchant reportDayPositionOfMerchantTotal = new ReportDayPositionOfMerchant();
//            reportDayPositionOfMerchantTotal.setSaleCount(0);
//            reportDayPositionOfMerchantTotal.setSaleAmount(new BigDecimal(0.00));
//            reportDayPositionOfMerchantTotal.setPurchaseCount(0);
//            reportDayPositionOfMerchantTotal.setPurchaseAmount(new BigDecimal(0.00));
//            reportDayPositionOfMerchantTotal.setRechargeCount(0);
//            reportDayPositionOfMerchantTotal.setRechargeAmount(new BigDecimal(0.00));
//            reportDayPositionOfMerchantTotal.setRechargeOnlineCount(0);
//            reportDayPositionOfMerchantTotal.setRechargeOnlineAmount(new BigDecimal(0.00));
//            reportDayPositionOfMerchantTotal.setConsumeCount(0);
//            reportDayPositionOfMerchantTotal.setConsumeAmount(new BigDecimal(0.00));
//            reportDayPositionOfMerchantTotal.setConsumeOnlineCount(0);
//            reportDayPositionOfMerchantTotal.setConsumeOnlineAmount(new BigDecimal(0.00));
//            reportDayPositionOfMerchantTotal.setRefundCount(0);
//            reportDayPositionOfMerchantTotal.setRefundAmount(new BigDecimal(0.00));
//            reportDayPositionOfMerchantTotal.setRedeemCount(0);
//            reportDayPositionOfMerchantTotal.setRedeemAmount(new BigDecimal(0.00));
//            reportDayPositionOfMerchantTotal.setFreezeCount(0);
//            reportDayPositionOfMerchantTotal.setFreezeAmount(new BigDecimal(0.00));
//            reportDayPositionOfMerchantTotal.setThawCount(0);
//            reportDayPositionOfMerchantTotal.setThawAmount(new BigDecimal(0.00));
//
//            for(ReportDayPositionOfMerchant reportDayPositionOfMerchant:reportDayPositionOfMerchantList){
//                content = new String[headers.length];
//                rowId = rowId + 1;
//                content[0] = String.valueOf(rowId);
//                content[1] = reportDayPositionOfMerchant.getPosition();
//                content[2] = reportDayPositionOfMerchant.getName();
//                content[3] = reportDayPositionOfMerchant.getSaleCount().toString();
//                content[4] = reportDayPositionOfMerchant.getSaleAmount().setScale(2).toString();
//                content[5] = reportDayPositionOfMerchant.getPurchaseCount().toString();
//                content[6] = reportDayPositionOfMerchant.getPurchaseAmount().setScale(2).toString();
//                content[7] = reportDayPositionOfMerchant.getRechargeCount().toString();
//                content[8] = reportDayPositionOfMerchant.getRechargeAmount().setScale(2).toString();
//                content[9] = reportDayPositionOfMerchant.getRechargeOnlineCount().toString();
//                content[10] = reportDayPositionOfMerchant.getRechargeOnlineAmount().setScale(2).toString();
//                content[11] = reportDayPositionOfMerchant.getConsumeCount().toString();
//                content[12] = reportDayPositionOfMerchant.getConsumeAmount().setScale(2).toString();
//                content[13] = reportDayPositionOfMerchant.getConsumeOnlineCount().toString();
//                content[14] = reportDayPositionOfMerchant.getConsumeOnlineAmount().setScale(2).toString();
//                content[15] = reportDayPositionOfMerchant.getRefundCount().toString();
//                content[16] = reportDayPositionOfMerchant.getRefundAmount().setScale(2).toString();
//                content[17] = reportDayPositionOfMerchant.getRedeemCount().toString();
//                content[18] = reportDayPositionOfMerchant.getRedeemAmount().setScale(2).toString();
//                content[19] = reportDayPositionOfMerchant.getFreezeCount().toString();
//                content[20] = reportDayPositionOfMerchant.getFreezeAmount().setScale(2).toString();
//                content[21] = reportDayPositionOfMerchant.getThawCount().toString();
//                content[22] = reportDayPositionOfMerchant.getThawAmount().setScale(2).toString();
//
//                reportDayPositionOfMerchantTotal.setSaleCount(reportDayPositionOfMerchantTotal.getSaleCount() + reportDayPositionOfMerchant.getSaleCount());
//                reportDayPositionOfMerchantTotal.setSaleAmount(reportDayPositionOfMerchantTotal.getSaleAmount().add(reportDayPositionOfMerchant.getSaleAmount()));
//                reportDayPositionOfMerchantTotal.setPurchaseCount(reportDayPositionOfMerchantTotal.getPurchaseCount() + reportDayPositionOfMerchant.getPurchaseCount());
//                reportDayPositionOfMerchantTotal.setPurchaseAmount(reportDayPositionOfMerchantTotal.getPurchaseAmount().add(reportDayPositionOfMerchant.getPurchaseAmount()));
//                reportDayPositionOfMerchantTotal.setRechargeCount(reportDayPositionOfMerchantTotal.getRechargeCount() + reportDayPositionOfMerchant.getRechargeCount());
//                reportDayPositionOfMerchantTotal.setRechargeAmount(reportDayPositionOfMerchantTotal.getRechargeAmount().add(reportDayPositionOfMerchant.getRechargeAmount()));
//                reportDayPositionOfMerchantTotal.setRechargeOnlineCount(reportDayPositionOfMerchantTotal.getRechargeOnlineCount() + reportDayPositionOfMerchant.getRechargeOnlineCount());
//                reportDayPositionOfMerchantTotal.setRechargeOnlineAmount(reportDayPositionOfMerchantTotal.getRechargeOnlineAmount().add(reportDayPositionOfMerchant.getRechargeOnlineAmount()));
//                reportDayPositionOfMerchantTotal.setConsumeCount(reportDayPositionOfMerchantTotal.getConsumeCount() + reportDayPositionOfMerchant.getConsumeCount());
//                reportDayPositionOfMerchantTotal.setConsumeAmount(reportDayPositionOfMerchantTotal.getConsumeAmount().add(reportDayPositionOfMerchant.getConsumeAmount()));
//                reportDayPositionOfMerchantTotal.setConsumeOnlineCount(reportDayPositionOfMerchantTotal.getConsumeOnlineCount() + reportDayPositionOfMerchant.getConsumeOnlineCount());
//                reportDayPositionOfMerchantTotal.setConsumeOnlineAmount(reportDayPositionOfMerchantTotal.getConsumeOnlineAmount().add(reportDayPositionOfMerchant.getConsumeOnlineAmount()));
//                reportDayPositionOfMerchantTotal.setRefundCount(reportDayPositionOfMerchantTotal.getRefundCount() + reportDayPositionOfMerchant.getRefundCount());
//                reportDayPositionOfMerchantTotal.setRefundAmount(reportDayPositionOfMerchantTotal.getRefundAmount().add(reportDayPositionOfMerchant.getRefundAmount()));
//                reportDayPositionOfMerchantTotal.setRedeemCount(reportDayPositionOfMerchantTotal.getRedeemCount() + reportDayPositionOfMerchant.getRedeemCount());
//                reportDayPositionOfMerchantTotal.setRedeemAmount(reportDayPositionOfMerchantTotal.getRedeemAmount().add(reportDayPositionOfMerchant.getRedeemAmount()));
//                reportDayPositionOfMerchantTotal.setFreezeCount(reportDayPositionOfMerchantTotal.getFreezeCount() + reportDayPositionOfMerchant.getFreezeCount());
//                reportDayPositionOfMerchantTotal.setFreezeAmount(reportDayPositionOfMerchantTotal.getFreezeAmount().add(reportDayPositionOfMerchant.getFreezeAmount()));
//                reportDayPositionOfMerchantTotal.setThawCount(reportDayPositionOfMerchantTotal.getThawCount() + reportDayPositionOfMerchant.getThawCount());
//                reportDayPositionOfMerchantTotal.setThawAmount(reportDayPositionOfMerchantTotal.getThawAmount().add(reportDayPositionOfMerchant.getThawAmount()));
//                contents.add(content);
//            }
//            //如果没有记录及不需要合计
//            if(contents.size() > 0) {
//                content = new String[headers.length];
//                content[2] = "合计";
//                content[3] = reportDayPositionOfMerchantTotal.getSaleCount().toString();
//                content[4] = reportDayPositionOfMerchantTotal.getSaleAmount().setScale(2).toString();
//                content[5] = reportDayPositionOfMerchantTotal.getPurchaseCount().toString();
//                content[6] = reportDayPositionOfMerchantTotal.getPurchaseAmount().setScale(2).toString();
//                content[7] = reportDayPositionOfMerchantTotal.getRechargeCount().toString();
//                content[8] = reportDayPositionOfMerchantTotal.getRechargeAmount().setScale(2).toString();
//                content[9] = reportDayPositionOfMerchantTotal.getRechargeOnlineCount().toString();
//                content[10] = reportDayPositionOfMerchantTotal.getRechargeOnlineAmount().setScale(2).toString();
//                content[11] = reportDayPositionOfMerchantTotal.getConsumeCount().toString();
//                content[12] = reportDayPositionOfMerchantTotal.getConsumeAmount().setScale(2).toString();
//                content[13] = reportDayPositionOfMerchantTotal.getConsumeOnlineCount().toString();
//                content[14] = reportDayPositionOfMerchantTotal.getConsumeOnlineAmount().setScale(2).toString();
//                content[15] = reportDayPositionOfMerchantTotal.getRefundCount().toString();
//                content[16] = reportDayPositionOfMerchantTotal.getRefundAmount().setScale(2).toString();
//                content[17] = reportDayPositionOfMerchantTotal.getRedeemCount().toString();
//                content[18] = reportDayPositionOfMerchantTotal.getRedeemAmount().setScale(2).toString();
//                content[19] = reportDayPositionOfMerchantTotal.getFreezeCount().toString();
//                content[20] = reportDayPositionOfMerchantTotal.getFreezeAmount().setScale(2).toString();
//                content[21] = reportDayPositionOfMerchantTotal.getThawCount().toString();
//                content[22] = reportDayPositionOfMerchantTotal.getThawAmount().setScale(2).toString();
//                contents.add(content);
//            }
//
//            try {
//                excelUtils.createSheet("商户日统计表(门店)",title,queryName,headers,contents);
//            } catch (Exception e) {
//                logger.error("商户日统计表(门店)Sheet失败,{}",e);
//            }
//            //2.创建商户统计表（终端）sheet
//            title = "消费/退款(按终端统计)";
//            queryName = new String[4];
//            queryName[0] = "商户编号 " + tMerchant1.getMerchantId();
//            queryName[1] = "商户名称 " + tMerchant1.getMerchantChineseName();
//            headers = new String[] { "序号", "门店编码", "门店名称", "终端编码", "终端名称",
//                    "线下消费笔数", "线下消费金额", "线上消费笔数", "线上消费金额", "退款笔数", "退款金额"};
//
//            contents = new ArrayList<>();
//            rowId = 0;
//
//            QueryCondition queryCondition = new QueryCondition();
//            queryCondition.setMerchantId(tMerchant1.getMerchantId());
//            queryCondition.setBeginTime(StringUtils.substring(tInstructionDayStatistics.getBillMonthBegin(),0,4) +
//                    "-" + StringUtils.substring(tInstructionDayStatistics.getBillMonthBegin(),4,6) + "-" +
//                    StringUtils.substring(tInstructionDayStatistics.getBillMonthBegin(),6,8) + " 00:00:00");
//            queryCondition.setEndTime(StringUtils.substring(tInstructionDayStatistics.getBillMonthEnd(),0,4) +
//                    "-" + StringUtils.substring(tInstructionDayStatistics.getBillMonthEnd(),4,6) + "-"
//                    + StringUtils.substring(tInstructionDayStatistics.getBillMonthEnd(),6,8) + " 23:59:59");
//            queryName[2] = "起始时间 " + queryCondition.getBeginTime();
//            queryName[3] = "结束时间 " + queryCondition.getEndTime();
//            List<TermMerchantStatics> termMerchantStaticsList= tInstructionLogService.getMapper().selectGroup(queryCondition);
//
//            //合计数据初始化
//            TermMerchantStatics termMerchantStaticsTotal = new TermMerchantStatics();
//            termMerchantStaticsTotal.setConsumeCount(0);
//            termMerchantStaticsTotal.setConsumeAmount(new BigDecimal(0.00));
//            termMerchantStaticsTotal.setConsumeOnlineCount(0);
//            termMerchantStaticsTotal.setConsumeOnlineAmount(new BigDecimal(0.00));
//            termMerchantStaticsTotal.setRefundCount(0);
//            termMerchantStaticsTotal.setRefundAmount(new BigDecimal(0.00));
//            //明细项
//            for(TermMerchantStatics termMerchantStatics:termMerchantStaticsList){
//                content = new String[headers.length];
//                rowId = rowId + 1;
//                content[0] = String.valueOf(rowId);
//                content[1] = termMerchantStatics.getStoreId();
//                content[2] = termMerchantStatics.getStoreName();
//                content[3] = termMerchantStatics.getTermId();
//                content[4] = termMerchantStatics.getTermName();
//                content[5] = termMerchantStatics.getConsumeCount().toString();
//                termMerchantStaticsTotal.setConsumeCount(termMerchantStaticsTotal.getConsumeCount() + termMerchantStatics.getConsumeCount());
//                content[6] = termMerchantStatics.getConsumeAmount().setScale(2).toString();
//                termMerchantStaticsTotal.setConsumeAmount(termMerchantStaticsTotal.getConsumeAmount().add(termMerchantStatics.getConsumeAmount()));
//                content[7] = termMerchantStatics.getConsumeOnlineCount().toString();
//                termMerchantStaticsTotal.setConsumeOnlineCount(termMerchantStaticsTotal.getConsumeOnlineCount() + termMerchantStatics.getConsumeOnlineCount());
//                content[8] = termMerchantStatics.getConsumeOnlineAmount().setScale(2).toString();
//                termMerchantStaticsTotal.setConsumeOnlineAmount(termMerchantStaticsTotal.getConsumeOnlineAmount().add(termMerchantStatics.getConsumeOnlineAmount()));
//                content[9] = termMerchantStatics.getRefundCount().toString();
//                termMerchantStaticsTotal.setRefundCount(termMerchantStaticsTotal.getRefundCount() + termMerchantStatics.getRefundCount());
//                content[10] = termMerchantStatics.getRefundAmount().setScale(2).toString();
//                termMerchantStaticsTotal.setRefundAmount(termMerchantStaticsTotal.getRefundAmount().add(termMerchantStatics.getRefundAmount()));
//                contents.add(content);
//            }
//
//            //如果没有记录及不需要合计
//            if(contents.size() > 0) {
//                content = new String[headers.length];
//                content[4] = "合计";
//                content[5] = termMerchantStaticsTotal.getConsumeCount().toString();
//                content[6] = termMerchantStaticsTotal.getConsumeAmount().setScale(2).toString();
//                content[7] = termMerchantStaticsTotal.getConsumeOnlineCount().toString();
//                content[8] = termMerchantStaticsTotal.getConsumeOnlineAmount().setScale(2).toString();
//                content[9] = termMerchantStaticsTotal.getRefundCount().toString();
//                content[10] = termMerchantStaticsTotal.getRefundAmount().setScale(2).toString();
//                contents.add(content);
//            }
//
//            try {
//                excelUtils.createSheet("商户日统计表(终端)",title,queryName,headers,contents);
//            } catch (Exception e) {
//                logger.error("商户日统计表(终端)Sheet失败,{}",e);
//            }
//
//            //3.创建卡产品统计表sheet
//            title = "卡产品统计报表";
//            queryName = new String[3];
//            queryName[0] = "商户编号 " + tMerchant1.getMerchantId();
//            queryName[1] = "商户名称 " + tMerchant1.getMerchantChineseName();
//            queryName[2] = "统计时间 " + "全部";
//            headers = new String[] { "序号", "卡名称", "卡标识码", "卡类型", "卡bin号",
//                    "总卡量", "制卡中卡量", "现卡量", "待客户激活", "已激活卡量", "已冻结卡量","已销户卡量"};
//
//            contents = new ArrayList<>();
//            rowId = 0;
//
//            queryCondition = new QueryCondition();
//            queryCondition.setMerchantId(tMerchant1.getMerchantId());
//
//            List<CardProductStatics> cardProductStaticsList= tCardService.getMapper().selectGroup(queryCondition);
//            CardProductStatics cardProductStaticsTotal = new CardProductStatics();
//            cardProductStaticsTotal.setCardTotalCount(0);
//            cardProductStaticsTotal.setCardA0Count(0);
//            cardProductStaticsTotal.setCardA1Count(0);
//            cardProductStaticsTotal.setCardA2Count(0);
//            cardProductStaticsTotal.setCardB0Count(0);
//            cardProductStaticsTotal.setCardB1Count(0);
//            cardProductStaticsTotal.setCardC0Count(0);
//
//            for(CardProductStatics cardProductStatics:cardProductStaticsList){
//                content = new String[headers.length];
//                rowId = rowId + 1;
//                content[0] = String.valueOf(rowId);
//                content[1] = cardProductStatics.getCardName();
//                content[2] = cardProductStatics.getCardType();
//                content[3] = cardProductStatics.getCardTypeName();
//                content[4] = cardProductStatics.getCardBin();
//                content[5] = cardProductStatics.getCardTotalCount().toString();
//                content[6] = cardProductStatics.getCardA0Count().toString();
//                content[7] = cardProductStatics.getCardA1Count().toString();
//                content[8] = cardProductStatics.getCardA2Count().toString();
//                content[9] = cardProductStatics.getCardB0Count().toString();
//                content[10] = cardProductStatics.getCardB1Count().toString();
//                content[11] = cardProductStatics.getCardC0Count().toString();
//
//                cardProductStaticsTotal.setCardTotalCount(cardProductStaticsTotal.getCardTotalCount() + cardProductStatics.getCardTotalCount());
//                cardProductStaticsTotal.setCardA0Count(cardProductStaticsTotal.getCardA0Count() + cardProductStatics.getCardA0Count());
//                cardProductStaticsTotal.setCardA1Count(cardProductStaticsTotal.getCardA1Count() + cardProductStatics.getCardA1Count());
//                cardProductStaticsTotal.setCardA2Count(cardProductStaticsTotal.getCardA2Count() + cardProductStatics.getCardA2Count());
//                cardProductStaticsTotal.setCardB0Count(cardProductStaticsTotal.getCardB0Count() + cardProductStatics.getCardB0Count());
//                cardProductStaticsTotal.setCardB1Count(cardProductStaticsTotal.getCardB1Count() + cardProductStatics.getCardB1Count());
//                cardProductStaticsTotal.setCardC0Count(cardProductStaticsTotal.getCardC0Count() + cardProductStatics.getCardC0Count());
//                contents.add(content);
//            }
//            //如果没有记录及不需要合计
//            if(contents.size() > 0) {
//                content = new String[headers.length];
//                content[4] = "合计";
//                content[5] = cardProductStaticsTotal.getCardTotalCount().toString();
//                content[6] = cardProductStaticsTotal.getCardA0Count().toString();
//                content[7] = cardProductStaticsTotal.getCardA1Count().toString();
//                content[8] = cardProductStaticsTotal.getCardA2Count().toString();
//                content[9] = cardProductStaticsTotal.getCardB0Count().toString();
//                content[10] = cardProductStaticsTotal.getCardB1Count().toString();
//                content[11] = cardProductStaticsTotal.getCardC0Count().toString();
//                contents.add(content);
//            }
//            try {
//                excelUtils.createSheet("卡产品统计表",title,queryName,headers,contents);
//            } catch (Exception e) {
//                logger.error("商户卡产品统计表Sheet失败,{}",e);
//            }
//
//            try {
//                excelUtils.excelFinal();
//            } catch (IOException e) {
//                logger.error("文件写入失败,{}",e);
//                return "文件写入失败";
//            }
//            //3.登记商户报表流水表
//            TMerchantReport tMerchantReport = new TMerchantReport();
//            tMerchantReport.setMerchantId(tMerchant1.getMerchantId());
//            tMerchantReport.setReportType("MD");
//            tMerchantReport.setReportBillDate(tInstructionDayStatistics.getBillMonthBegin());
//            tMerchantReport.setCutDayTime("00:00");
//            tMerchantReport.setReportFileGenerated("1");
//            tMerchantReport.setJobSerial(tJobLog.getJobSerial());
//            tMerchantReport.setRecordCreateTime(DateFormatUtils.format(new Date(),"yyyy-MM-dd hh:mm:ss"));
//            tMerchantReportService.getMapper().insert(tMerchantReport);
//        }
//
//        //4.更新事务流水
//        tJobLog.setJobStatus("01");
//        tJobLogService.getMapper().updateByPrimaryKey(tJobLog);
//        logger.info("genDayReportFile finish......");
//        return "交易成功";
//    }

    */
/**
     * 生成商户日报表
     *
     * @return
     *//*


    public String genDayReportFile() {
        logger.info("genDayReportFile begin......");
        TJobLog tJobLog = new TJobLog();

        String today = DateFormatUtils.format(new Date(), "yyyyMMddHHmmss");
        String randomNum = RandomStringUtils.randomNumeric(3);
        tJobLog.setJobSerial(today + randomNum);
        tJobLog.setJobType("MER_REPORT");
        tJobLog.setJobBurstType("SCHEDULED");
        tJobLog.setJobStatus("00");
        //1.登记事务流水
        tJobLogService.getMapper().insert(tJobLog);
        //2.查询商户表
        TMerchant tMerchant = new TMerchant();
        List<TMerchant> tMerchantList = tMerchantService.getMapper().select(tMerchant);
        //查询条件赋值
        QueryCondition queryCondition = new QueryCondition();
        if ("1".equals(manualFlag)) {
            queryCondition.setBeginTime(beginDate);
            queryCondition.setEndTime(endDate);
        } else {
            String lastDate = DateFormatUtils.format(DateUtils.addDays(new Date(), -1), "yyyyMMdd");
            queryCondition.setBeginTime(lastDate);
            queryCondition.setEndTime(lastDate);
        }
        //循环获取商户号
        for (TMerchant tMerchant1 : tMerchantList) {
            logger.info("start {} merchant excel file....", tMerchant1.getMerchantId());
            queryCondition.setMerchantId(tMerchant1.getMerchantId());
            ExcelUtils excelUtils = null;
            String filePath = reportPath + "/" + tMerchant1.getMerchantId() + "/";
            File file = new File(filePath);
            String[] content;
            if (!file.exists() && !file.isDirectory()) {
                file.mkdirs();
            }
            //开始赋值
            String fileName = filePath + "商户日流水报表_M" + tMerchant1.getMerchantId() + "_" + queryCondition.getBeginTime() + ".xls";
            try {
                excelUtils = new ExcelUtils(fileName);
            } catch (FileNotFoundException e) {
                logger.error("创建文件{}失败,{}", fileName, e);
                return "创建文件失败";
            }

            //1.生成第一个sheet 售卡/充值流水表
            String title = "售卡/充值流水";
            String[] queryName = new String[4];
            queryName[0] = "商户编号 " + tMerchant1.getMerchantId();
            queryName[1] = "商户名称 " + tMerchant1.getMerchantChineseName();
            queryName[2] = "起始时间 " + queryCondition.getBeginTime();
            queryName[3] = "结束时间 " + queryCondition.getEndTime();
            String[] headers = new String[]{"序号", "门店编码", "门店名称", "交易流水号",
                    "卡名称", "卡类型", "售卡/充值", "起始卡号", "卡数量", "线下充值金额(元)", "线上充值金额(元)",
                    "实收金额(元)", "手续费(元)", "优惠金额(元)"};


            List<Object[]> contents = new ArrayList<>();
            int rowId = 0;

            List<SaleRechargeCardDetail> saleRechargeCardDetailList = tBillFlowService.getMapper().selectSaleRechargeDetail(queryCondition);

            SaleRechargeCardDetail saleRechargeCardDetailTotal = new SaleRechargeCardDetail();
            saleRechargeCardDetailTotal.setCardCnt(0);
            saleRechargeCardDetailTotal.setOffAmount(new BigDecimal(0.00));
            saleRechargeCardDetailTotal.setOnAmount(new BigDecimal(0.00));
            saleRechargeCardDetailTotal.setRealAmount(new BigDecimal(0.00));
            saleRechargeCardDetailTotal.setFeeAmount(new BigDecimal(0.00));
            saleRechargeCardDetailTotal.setDisAmount(new BigDecimal(0.00));

            for (SaleRechargeCardDetail saleRechargeCardDetail : saleRechargeCardDetailList) {
                content = new String[headers.length];
                rowId = rowId + 1;
                content[0] = String.valueOf(rowId);
                content[1] = saleRechargeCardDetail.getStoreId();
                content[2] = saleRechargeCardDetail.getStoreName();
                content[3] = saleRechargeCardDetail.getTranSeq();
                content[4] = saleRechargeCardDetail.getCardName();
                content[5] = saleRechargeCardDetail.getCardTypeName();
                content[6] = EnumUtils.getEnum(TransType.class, saleRechargeCardDetail.getTranType()).getValue();
                String firstCardNo = saleRechargeCardDetail.getCardNo().split(",")[0];
                content[7] = firstCardNo.substring(0, 4) + " **** **** " + firstCardNo.substring(firstCardNo.length() - 4, firstCardNo.length());
                content[8] = saleRechargeCardDetail.getCardCnt().toString();
                content[9] = saleRechargeCardDetail.getOffAmount().setScale(2).toString();
                content[10] = saleRechargeCardDetail.getOnAmount().setScale(2).toString();
                content[11] = saleRechargeCardDetail.getRealAmount().setScale(2).toString();
                content[12] = saleRechargeCardDetail.getFeeAmount().setScale(2).toString();
                content[13] = saleRechargeCardDetail.getDisAmount().setScale(2).toString();

                saleRechargeCardDetailTotal.setCardCnt(saleRechargeCardDetailTotal.getCardCnt() + saleRechargeCardDetail.getCardCnt());
                saleRechargeCardDetailTotal.setOffAmount(saleRechargeCardDetailTotal.getOffAmount().add(saleRechargeCardDetail.getOffAmount()));
                saleRechargeCardDetailTotal.setOnAmount(saleRechargeCardDetailTotal.getOnAmount().add(saleRechargeCardDetail.getOnAmount()));
                saleRechargeCardDetailTotal.setRealAmount(saleRechargeCardDetailTotal.getRealAmount().add(saleRechargeCardDetail.getRealAmount()));
                saleRechargeCardDetailTotal.setFeeAmount(saleRechargeCardDetailTotal.getFeeAmount().add(saleRechargeCardDetail.getFeeAmount()));
                saleRechargeCardDetailTotal.setDisAmount(saleRechargeCardDetailTotal.getDisAmount().add(saleRechargeCardDetail.getDisAmount()));

                contents.add(content);
            }
            //如果没有记录及不需要合计
            if (contents.size() > 0) {
                content = new String[headers.length];
                content[7] = "合计";
                content[8] = saleRechargeCardDetailTotal.getCardCnt().toString();
                content[9] = saleRechargeCardDetailTotal.getOffAmount().setScale(2).toString();
                content[10] = saleRechargeCardDetailTotal.getOnAmount().setScale(2).toString();
                content[11] = saleRechargeCardDetailTotal.getRealAmount().setScale(2).toString();
                content[12] = saleRechargeCardDetailTotal.getFeeAmount().setScale(2).toString();
                content[13] = saleRechargeCardDetailTotal.getDisAmount().setScale(2).toString();
                contents.add(content);
            }

            try {
                excelUtils.createSheet("售卡_充值流水表", title, queryName, headers, contents);
            } catch (Exception e) {
                logger.error("售卡_充值流水表Sheet失败,{}", e);
            }
            //2.创建消费_退款流水表sheet
            title = "消费_退款流水表";
            queryName = new String[4];
            queryName[0] = "商户编号 " + tMerchant1.getMerchantId();
            queryName[1] = "商户名称 " + tMerchant1.getMerchantChineseName();
            queryName[2] = "起始时间 " + queryCondition.getBeginTime();
            queryName[3] = "结束时间 " + queryCondition.getEndTime();
            headers = new String[]{"序号", "门店编码", "门店名称", "终端编码", "终端名称",
                    "交易流水号", "卡名称", "卡类型", "卡号", "消费/退货", "实际金额(元)"};

            contents = new ArrayList<>();
            rowId = 0;

            List<TermMerchantDetail> termMerchantDetailList = tInstructionLogService.getMapper().selectConsumeRefundDetail(queryCondition);

            TermMerchantDetail termMerchantDetailTotal = new TermMerchantDetail();
            termMerchantDetailTotal.setTransAmount(new BigDecimal(0.00));

            for (TermMerchantDetail termMerchantDetail : termMerchantDetailList) {
                content = new String[headers.length];
                rowId = rowId + 1;
                content[0] = String.valueOf(rowId);
                content[1] = termMerchantDetail.getStoreId();
                content[2] = termMerchantDetail.getStoreName();
                content[3] = termMerchantDetail.getTermId();
                content[4] = termMerchantDetail.getTermName();
                content[5] = termMerchantDetail.getTranNo();
                content[6] = termMerchantDetail.getCardName();
                content[7] = termMerchantDetail.getCardType();
                String firstCardNo = termMerchantDetail.getCardNo();
                if (firstCardNo != null && firstCardNo != "") {
                    content[8] = firstCardNo.substring(0, 4) + " **** **** " + firstCardNo.substring(firstCardNo.length() - 4, firstCardNo.length());
                }
                content[9] = EnumUtils.getEnum(TransType.class, termMerchantDetail.getTransType()).getValue();
                if (termMerchantDetail.getTransAmount() != null) {
                    content[10] = termMerchantDetail.getTransAmount().toString();
                    termMerchantDetailTotal.setTransAmount(termMerchantDetailTotal.getTransAmount().add(termMerchantDetail.getTransAmount()));
                }

                contents.add(content);
            }
            //如果没有记录及不需要合计
            if (contents.size() > 0) {
                content = new String[headers.length];
                content[9] = "合计";
                content[10] = termMerchantDetailTotal.getTransAmount().setScale(2).toString();

                contents.add(content);
            }

            try {
                excelUtils.createSheet("消费_退款流水表", title, queryName, headers, contents);
            } catch (Exception e) {
                logger.error("消费_退款流水表Sheet失败,{}", e);
            }

            //3.卡余额列表sheet
            title = "卡余额列表";
            queryName = new String[3];
            queryName[0] = "商户编号 " + tMerchant1.getMerchantId();
            queryName[1] = "商户名称 " + tMerchant1.getMerchantChineseName();
            queryName[2] = "统计时间 " + "全部";
            headers = new String[]{"序号", "卡标识码", "卡名称", "卡类型", "卡bin号",
                    "卡号", "卡余额(元)", "卡状态", "卡有效期"};

            contents = new ArrayList<>();
            rowId = 0;

            List<TCard> tCardList = tCardService.getMapper().selectCardDetail(queryCondition);
            TCard tCardTotal = new TCard();
            tCardTotal.setFaceBalance(new BigDecimal(0.00));

            for (TCard tCard : tCardList) {
                content = new String[headers.length];
                rowId = rowId + 1;
                content[0] = String.valueOf(rowId);
                content[1] = tCard.getCardUniqueStr();
                content[2] = tCard.getCardName();
                content[3] = tCard.getCardType();
                content[4] = tCard.getCardBin();
                String cardNo = tCard.getCardNo();
                if (cardNo != null) {
                    content[5] = cardNo.substring(0, 4) + " **** **** " + cardNo.substring(cardNo.length() - 4, cardNo.length());
                }
                if (tCard.getFaceBalance() != null) {
                    content[6] = tCard.getFaceBalance().toString();
                    tCardTotal.setFaceBalance(tCardTotal.getFaceBalance().add(tCard.getFaceBalance()));
                }
                if (tCard.getCardStatus() != null) {
                    content[7] = EnumUtils.getEnum(CardStatus.class, tCard.getCardStatus()).getValue();
                }
                content[8] = tCard.getExpiryDate();

                contents.add(content);
            }
            //如果没有记录及不需要合计
            if (contents.size() > 0) {
                content = new String[headers.length];
                content[5] = "合计";
                content[6] = tCardTotal.getFaceBalance().toString();
                contents.add(content);
            }
            try {
                excelUtils.createSheet("卡余额表", title, queryName, headers, contents);
            } catch (Exception e) {
                logger.error("卡余额表Sheet失败,{}", e);
            }

            try {
                excelUtils.excelFinal();
            } catch (IOException e) {
                logger.error("文件写入失败,{}", e);
                return "文件写入失败";
            }

            //开始赋值
            fileName = filePath + "商户日统计报表_M" + tMerchant1.getMerchantId() + "_" + queryCondition.getBeginTime() + ".xls";
            try {
                excelUtils = new ExcelUtils(fileName);
            } catch (FileNotFoundException e) {
                logger.error("创建文件{}失败,{}", fileName, e);
                return "创建文件失败";
            }

            //4.生成第一个sheet 商户日统计表(门店)
            title = "售卡/充值/消费/退款/赎回统计表(按门店统计)";
            queryName = new String[4];
            queryName[0] = "商户编号 " + tMerchant1.getMerchantId();
            queryName[1] = "商户名称 " + tMerchant1.getMerchantChineseName();
            queryName[2] = "起始时间 " + queryCondition.getBeginTime();
            queryName[3] = "结束时间 " + queryCondition.getEndTime();
            headers = new String[]{"序号", "门店编码", "门店名称", "线下售卡笔数", "线下售卡金额",
                    "线上售卡笔数", "线上售卡金额", "线下充值笔数", "线下充值金额", "线上充值笔数", "线上充值金额",
                    "线下消费笔数", "线下消费金额", "线上消费笔数", "线上消费金额", "退款笔数", "退款金额",
                    "赎回笔数", "赎回金额", "冻结笔数", "冻结金额", "解冻笔数", "解冻金额"};

            contents = new ArrayList<>();
            rowId = 0;

            List<ReportDayPositionOfMerchant> reportDayPositionOfMerchantList = tInstructionDayStatisticsService.getMapper().selectGroupFile(queryCondition);

            ReportDayPositionOfMerchant reportDayPositionOfMerchantTotal = new ReportDayPositionOfMerchant();
            reportDayPositionOfMerchantTotal.setSaleCount(0);
            reportDayPositionOfMerchantTotal.setSaleAmount(new BigDecimal(0.00));
            reportDayPositionOfMerchantTotal.setPurchaseCount(0);
            reportDayPositionOfMerchantTotal.setPurchaseAmount(new BigDecimal(0.00));
            reportDayPositionOfMerchantTotal.setRechargeCount(0);
            reportDayPositionOfMerchantTotal.setRechargeAmount(new BigDecimal(0.00));
            reportDayPositionOfMerchantTotal.setRechargeOnlineCount(0);
            reportDayPositionOfMerchantTotal.setRechargeOnlineAmount(new BigDecimal(0.00));
            reportDayPositionOfMerchantTotal.setConsumeCount(0);
            reportDayPositionOfMerchantTotal.setConsumeAmount(new BigDecimal(0.00));
            reportDayPositionOfMerchantTotal.setConsumeOnlineCount(0);
            reportDayPositionOfMerchantTotal.setConsumeOnlineAmount(new BigDecimal(0.00));
            reportDayPositionOfMerchantTotal.setRefundCount(0);
            reportDayPositionOfMerchantTotal.setRefundAmount(new BigDecimal(0.00));
            reportDayPositionOfMerchantTotal.setRedeemCount(0);
            reportDayPositionOfMerchantTotal.setRedeemAmount(new BigDecimal(0.00));
            reportDayPositionOfMerchantTotal.setFreezeCount(0);
            reportDayPositionOfMerchantTotal.setFreezeAmount(new BigDecimal(0.00));
            reportDayPositionOfMerchantTotal.setThawCount(0);
            reportDayPositionOfMerchantTotal.setThawAmount(new BigDecimal(0.00));

            for (ReportDayPositionOfMerchant reportDayPositionOfMerchant : reportDayPositionOfMerchantList) {
                content = new String[headers.length];
                rowId = rowId + 1;
                content[0] = String.valueOf(rowId);
                content[1] = reportDayPositionOfMerchant.getPosition();
                content[2] = reportDayPositionOfMerchant.getName();
                content[3] = reportDayPositionOfMerchant.getSaleCount().toString();
                content[4] = reportDayPositionOfMerchant.getSaleAmount().setScale(2).toString();
                content[5] = reportDayPositionOfMerchant.getPurchaseCount().toString();
                content[6] = reportDayPositionOfMerchant.getPurchaseAmount().setScale(2).toString();
                content[7] = reportDayPositionOfMerchant.getRechargeCount().toString();
                content[8] = reportDayPositionOfMerchant.getRechargeAmount().setScale(2).toString();
                content[9] = reportDayPositionOfMerchant.getRechargeOnlineCount().toString();
                content[10] = reportDayPositionOfMerchant.getRechargeOnlineAmount().setScale(2).toString();
                content[11] = reportDayPositionOfMerchant.getConsumeCount().toString();
                content[12] = reportDayPositionOfMerchant.getConsumeAmount().setScale(2).toString();
                content[13] = reportDayPositionOfMerchant.getConsumeOnlineCount().toString();
                content[14] = reportDayPositionOfMerchant.getConsumeOnlineAmount().setScale(2).toString();
                content[15] = reportDayPositionOfMerchant.getRefundCount().toString();
                content[16] = reportDayPositionOfMerchant.getRefundAmount().setScale(2).toString();
                content[17] = reportDayPositionOfMerchant.getRedeemCount().toString();
                content[18] = reportDayPositionOfMerchant.getRedeemAmount().setScale(2).toString();
                content[19] = reportDayPositionOfMerchant.getFreezeCount().toString();
                content[20] = reportDayPositionOfMerchant.getFreezeAmount().setScale(2).toString();
                content[21] = reportDayPositionOfMerchant.getThawCount().toString();
                content[22] = reportDayPositionOfMerchant.getThawAmount().setScale(2).toString();

                reportDayPositionOfMerchantTotal.setSaleCount(reportDayPositionOfMerchantTotal.getSaleCount() + reportDayPositionOfMerchant.getSaleCount());
                reportDayPositionOfMerchantTotal.setSaleAmount(reportDayPositionOfMerchantTotal.getSaleAmount().add(reportDayPositionOfMerchant.getSaleAmount()));
                reportDayPositionOfMerchantTotal.setPurchaseCount(reportDayPositionOfMerchantTotal.getPurchaseCount() + reportDayPositionOfMerchant.getPurchaseCount());
                reportDayPositionOfMerchantTotal.setPurchaseAmount(reportDayPositionOfMerchantTotal.getPurchaseAmount().add(reportDayPositionOfMerchant.getPurchaseAmount()));
                reportDayPositionOfMerchantTotal.setRechargeCount(reportDayPositionOfMerchantTotal.getRechargeCount() + reportDayPositionOfMerchant.getRechargeCount());
                reportDayPositionOfMerchantTotal.setRechargeAmount(reportDayPositionOfMerchantTotal.getRechargeAmount().add(reportDayPositionOfMerchant.getRechargeAmount()));
                reportDayPositionOfMerchantTotal.setRechargeOnlineCount(reportDayPositionOfMerchantTotal.getRechargeOnlineCount() + reportDayPositionOfMerchant.getRechargeOnlineCount());
                reportDayPositionOfMerchantTotal.setRechargeOnlineAmount(reportDayPositionOfMerchantTotal.getRechargeOnlineAmount().add(reportDayPositionOfMerchant.getRechargeOnlineAmount()));
                reportDayPositionOfMerchantTotal.setConsumeCount(reportDayPositionOfMerchantTotal.getConsumeCount() + reportDayPositionOfMerchant.getConsumeCount());
                reportDayPositionOfMerchantTotal.setConsumeAmount(reportDayPositionOfMerchantTotal.getConsumeAmount().add(reportDayPositionOfMerchant.getConsumeAmount()));
                reportDayPositionOfMerchantTotal.setConsumeOnlineCount(reportDayPositionOfMerchantTotal.getConsumeOnlineCount() + reportDayPositionOfMerchant.getConsumeOnlineCount());
                reportDayPositionOfMerchantTotal.setConsumeOnlineAmount(reportDayPositionOfMerchantTotal.getConsumeOnlineAmount().add(reportDayPositionOfMerchant.getConsumeOnlineAmount()));
                reportDayPositionOfMerchantTotal.setRefundCount(reportDayPositionOfMerchantTotal.getRefundCount() + reportDayPositionOfMerchant.getRefundCount());
                reportDayPositionOfMerchantTotal.setRefundAmount(reportDayPositionOfMerchantTotal.getRefundAmount().add(reportDayPositionOfMerchant.getRefundAmount()));
                reportDayPositionOfMerchantTotal.setRedeemCount(reportDayPositionOfMerchantTotal.getRedeemCount() + reportDayPositionOfMerchant.getRedeemCount());
                reportDayPositionOfMerchantTotal.setRedeemAmount(reportDayPositionOfMerchantTotal.getRedeemAmount().add(reportDayPositionOfMerchant.getRedeemAmount()));
                reportDayPositionOfMerchantTotal.setFreezeCount(reportDayPositionOfMerchantTotal.getFreezeCount() + reportDayPositionOfMerchant.getFreezeCount());
                reportDayPositionOfMerchantTotal.setFreezeAmount(reportDayPositionOfMerchantTotal.getFreezeAmount().add(reportDayPositionOfMerchant.getFreezeAmount()));
                reportDayPositionOfMerchantTotal.setThawCount(reportDayPositionOfMerchantTotal.getThawCount() + reportDayPositionOfMerchant.getThawCount());
                reportDayPositionOfMerchantTotal.setThawAmount(reportDayPositionOfMerchantTotal.getThawAmount().add(reportDayPositionOfMerchant.getThawAmount()));
                contents.add(content);
            }
            //如果没有记录及不需要合计
            if (contents.size() > 0) {
                content = new String[headers.length];
                content[2] = "合计";
                content[3] = reportDayPositionOfMerchantTotal.getSaleCount().toString();
                content[4] = reportDayPositionOfMerchantTotal.getSaleAmount().setScale(2).toString();
                content[5] = reportDayPositionOfMerchantTotal.getPurchaseCount().toString();
                content[6] = reportDayPositionOfMerchantTotal.getPurchaseAmount().setScale(2).toString();
                content[7] = reportDayPositionOfMerchantTotal.getRechargeCount().toString();
                content[8] = reportDayPositionOfMerchantTotal.getRechargeAmount().setScale(2).toString();
                content[9] = reportDayPositionOfMerchantTotal.getRechargeOnlineCount().toString();
                content[10] = reportDayPositionOfMerchantTotal.getRechargeOnlineAmount().setScale(2).toString();
                content[11] = reportDayPositionOfMerchantTotal.getConsumeCount().toString();
                content[12] = reportDayPositionOfMerchantTotal.getConsumeAmount().setScale(2).toString();
                content[13] = reportDayPositionOfMerchantTotal.getConsumeOnlineCount().toString();
                content[14] = reportDayPositionOfMerchantTotal.getConsumeOnlineAmount().setScale(2).toString();
                content[15] = reportDayPositionOfMerchantTotal.getRefundCount().toString();
                content[16] = reportDayPositionOfMerchantTotal.getRefundAmount().setScale(2).toString();
                content[17] = reportDayPositionOfMerchantTotal.getRedeemCount().toString();
                content[18] = reportDayPositionOfMerchantTotal.getRedeemAmount().setScale(2).toString();
                content[19] = reportDayPositionOfMerchantTotal.getFreezeCount().toString();
                content[20] = reportDayPositionOfMerchantTotal.getFreezeAmount().setScale(2).toString();
                content[21] = reportDayPositionOfMerchantTotal.getThawCount().toString();
                content[22] = reportDayPositionOfMerchantTotal.getThawAmount().setScale(2).toString();
                contents.add(content);
            }

            try {
                excelUtils.createSheet("商户日统计表(门店)", title, queryName, headers, contents);
            } catch (Exception e) {
                logger.error("商户日统计表(门店)Sheet失败,{}", e);
            }

            //5.创建商户统计表（终端）sheet
            title = "消费/退款(按终端统计)";
            queryName = new String[4];
            queryName[0] = "商户编号 " + tMerchant1.getMerchantId();
            queryName[1] = "商户名称 " + tMerchant1.getMerchantChineseName();
            headers = new String[]{"序号", "门店编码", "门店名称", "终端编码", "终端名称",
                    "线下消费笔数", "线下消费金额", "线上消费笔数", "线上消费金额", "退款笔数", "退款金额"};

            contents = new ArrayList<>();
            rowId = 0;
            QueryCondition queryCondition1 = new QueryCondition();
            queryCondition1.setMerchantId(queryCondition.getMerchantId());

            queryCondition1.setBeginTime(StringUtils.substring(queryCondition.getBeginTime(), 0, 4) +
                    "-" + StringUtils.substring(queryCondition.getBeginTime(), 4, 6) + "-" +
                    StringUtils.substring(queryCondition.getBeginTime(), 6, 8) + " 00:00:00");
            queryCondition1.setEndTime(StringUtils.substring(queryCondition.getEndTime(), 0, 4) +
                    "-" + StringUtils.substring(queryCondition.getEndTime(), 4, 6) + "-"
                    + StringUtils.substring(queryCondition.getEndTime(), 6, 8) + " 23:59:59");

            queryName[2] = "起始时间 " + queryCondition.getBeginTime();
            queryName[3] = "结束时间 " + queryCondition.getEndTime();

            List<TermMerchantStatics> termMerchantStaticsList = tInstructionLogService.getMapper().selectGroup(queryCondition1);

            //合计数据初始化
            TermMerchantStatics termMerchantStaticsTotal = new TermMerchantStatics();
            termMerchantStaticsTotal.setConsumeCount(0);
            termMerchantStaticsTotal.setConsumeAmount(new BigDecimal(0.00));
            termMerchantStaticsTotal.setConsumeOnlineCount(0);
            termMerchantStaticsTotal.setConsumeOnlineAmount(new BigDecimal(0.00));
            termMerchantStaticsTotal.setRefundCount(0);
            termMerchantStaticsTotal.setRefundAmount(new BigDecimal(0.00));

            //明细项
            for (TermMerchantStatics termMerchantStatics : termMerchantStaticsList) {
                content = new String[headers.length];
                rowId = rowId + 1;
                content[0] = String.valueOf(rowId);
                content[1] = termMerchantStatics.getStoreId();
                content[2] = termMerchantStatics.getStoreName();
                content[3] = termMerchantStatics.getTermId();
                content[4] = termMerchantStatics.getTermName();
                content[5] = termMerchantStatics.getConsumeCount().toString();
                termMerchantStaticsTotal.setConsumeCount(termMerchantStaticsTotal.getConsumeCount() + termMerchantStatics.getConsumeCount());
                content[6] = termMerchantStatics.getConsumeAmount().setScale(2).toString();
                termMerchantStaticsTotal.setConsumeAmount(termMerchantStaticsTotal.getConsumeAmount().add(termMerchantStatics.getConsumeAmount()));
                content[7] = termMerchantStatics.getConsumeOnlineCount().toString();
                termMerchantStaticsTotal.setConsumeOnlineCount(termMerchantStaticsTotal.getConsumeOnlineCount() + termMerchantStatics.getConsumeOnlineCount());
                content[8] = termMerchantStatics.getConsumeOnlineAmount().setScale(2).toString();
                termMerchantStaticsTotal.setConsumeOnlineAmount(termMerchantStaticsTotal.getConsumeOnlineAmount().add(termMerchantStatics.getConsumeOnlineAmount()));
                content[9] = termMerchantStatics.getRefundCount().toString();
                termMerchantStaticsTotal.setRefundCount(termMerchantStaticsTotal.getRefundCount() + termMerchantStatics.getRefundCount());
                content[10] = termMerchantStatics.getRefundAmount().setScale(2).toString();
                termMerchantStaticsTotal.setRefundAmount(termMerchantStaticsTotal.getRefundAmount().add(termMerchantStatics.getRefundAmount()));
                contents.add(content);
            }

            //如果没有记录及不需要合计
            if (contents.size() > 0) {
                content = new String[headers.length];
                content[4] = "合计";
                content[5] = termMerchantStaticsTotal.getConsumeCount().toString();
                content[6] = termMerchantStaticsTotal.getConsumeAmount().setScale(2).toString();
                content[7] = termMerchantStaticsTotal.getConsumeOnlineCount().toString();
                content[8] = termMerchantStaticsTotal.getConsumeOnlineAmount().setScale(2).toString();
                content[9] = termMerchantStaticsTotal.getRefundCount().toString();
                content[10] = termMerchantStaticsTotal.getRefundAmount().setScale(2).toString();
                contents.add(content);
            }

            try {
                excelUtils.createSheet("商户日统计表(终端)", title, queryName, headers, contents);
            } catch (Exception e) {
                logger.error("商户日统计表(终端)Sheet失败,{}", e);
            }

            //6.创建卡产品统计表sheet
            title = "卡产品统计报表";
            queryName = new String[3];
            queryName[0] = "商户编号 " + tMerchant1.getMerchantId();
            queryName[1] = "商户名称 " + tMerchant1.getMerchantChineseName();
            queryName[2] = "统计时间 " + "全部";
            headers = new String[]{"序号", "卡名称", "卡标识码", "卡类型", "卡bin号",
                    "总卡量", "制卡中卡量", "现卡量", "待客户激活", "已激活卡量", "已冻结卡量", "已销户卡量"};

            contents = new ArrayList<>();
            rowId = 0;

            List<CardProductStatics> cardProductStaticsList = tCardService.getMapper().selectGroup(queryCondition);
            CardProductStatics cardProductStaticsTotal = new CardProductStatics();
            cardProductStaticsTotal.setCardTotalCount(0);
            cardProductStaticsTotal.setCardA0Count(0);
            cardProductStaticsTotal.setCardA1Count(0);
            cardProductStaticsTotal.setCardA2Count(0);
            cardProductStaticsTotal.setCardB0Count(0);
            cardProductStaticsTotal.setCardB1Count(0);
            cardProductStaticsTotal.setCardC0Count(0);

            for (CardProductStatics cardProductStatics : cardProductStaticsList) {
                content = new String[headers.length];
                rowId = rowId + 1;
                content[0] = String.valueOf(rowId);
                content[1] = cardProductStatics.getCardName();
                content[2] = cardProductStatics.getCardType();
                content[3] = cardProductStatics.getCardTypeName();
                content[4] = cardProductStatics.getCardBin();
                content[5] = cardProductStatics.getCardTotalCount().toString();
                content[6] = cardProductStatics.getCardA0Count().toString();
                content[7] = cardProductStatics.getCardA1Count().toString();
                content[8] = cardProductStatics.getCardA2Count().toString();
                content[9] = cardProductStatics.getCardB0Count().toString();
                content[10] = cardProductStatics.getCardB1Count().toString();
                content[11] = cardProductStatics.getCardC0Count().toString();

                cardProductStaticsTotal.setCardTotalCount(cardProductStaticsTotal.getCardTotalCount() + cardProductStatics.getCardTotalCount());
                cardProductStaticsTotal.setCardA0Count(cardProductStaticsTotal.getCardA0Count() + cardProductStatics.getCardA0Count());
                cardProductStaticsTotal.setCardA1Count(cardProductStaticsTotal.getCardA1Count() + cardProductStatics.getCardA1Count());
                cardProductStaticsTotal.setCardA2Count(cardProductStaticsTotal.getCardA2Count() + cardProductStatics.getCardA2Count());
                cardProductStaticsTotal.setCardB0Count(cardProductStaticsTotal.getCardB0Count() + cardProductStatics.getCardB0Count());
                cardProductStaticsTotal.setCardB1Count(cardProductStaticsTotal.getCardB1Count() + cardProductStatics.getCardB1Count());
                cardProductStaticsTotal.setCardC0Count(cardProductStaticsTotal.getCardC0Count() + cardProductStatics.getCardC0Count());
                contents.add(content);
            }
            //如果没有记录及不需要合计
            if (contents.size() > 0) {
                content = new String[headers.length];
                content[4] = "合计";
                content[5] = cardProductStaticsTotal.getCardTotalCount().toString();
                content[6] = cardProductStaticsTotal.getCardA0Count().toString();
                content[7] = cardProductStaticsTotal.getCardA1Count().toString();
                content[8] = cardProductStaticsTotal.getCardA2Count().toString();
                content[9] = cardProductStaticsTotal.getCardB0Count().toString();
                content[10] = cardProductStaticsTotal.getCardB1Count().toString();
                content[11] = cardProductStaticsTotal.getCardC0Count().toString();
                contents.add(content);
            }
            try {
                excelUtils.createSheet("卡产品统计表", title, queryName, headers, contents);
            } catch (Exception e) {
                logger.error("商户卡产品统计表Sheet失败,{}", e);
            }

            try {
                excelUtils.excelFinal();
            } catch (IOException e) {
                logger.error("文件写入失败,{}", e);
                return "文件写入失败";
            }

            //7.登记商户报表流水表
            TMerchantReport tMerchantReport = new TMerchantReport();
            tMerchantReport.setMerchantId(tMerchant1.getMerchantId());
            tMerchantReport.setReportType("MD");
            tMerchantReport.setReportBillDate(queryCondition.getBeginTime());
            tMerchantReport.setCutDayTime("00:00");
            tMerchantReport.setReportFileGenerated("1");
            tMerchantReport.setJobSerial(tJobLog.getJobSerial());
            tMerchantReport.setRecordCreateTime(DateFormatUtils.format(new Date(), "yyyy-MM-dd hh:mm:ss"));
            tMerchantReportService.getMapper().insert(tMerchantReport);
        }

        //8.更新事务流水
        tJobLog.setJobStatus("01");
        tJobLogService.getMapper().updateByPrimaryKey(tJobLog);
        logger.info("genDayReportFile finish......");
        return "交易成功";
    }

    @Override
    public String genMonthReportData() {
        logger.info("genMonthReportData begin......");
        TJobLog tJobLog = new TJobLog();

        String today = DateFormatUtils.format(new Date(), "yyyyMMddHHmmss");
        String randomNum = RandomStringUtils.randomNumeric(3);
        tJobLog.setJobSerial(today + randomNum);
        tJobLog.setJobType("SYS_CUT");
        tJobLog.setJobBurstType("SCHEDULED");
        tJobLog.setJobStatus("00");
        //1.登记事务流水
        tJobLogService.getMapper().insert(tJobLog);
        //2.查询商户表
        TMerchant tMerchant = new TMerchant();
        List<TMerchant> tMerchantList = tMerchantService.getMapper().select(tMerchant);

        //查询条件赋值
        QueryCondition queryCondition = new QueryCondition();
        if ("1".equals(manualFlag)) {
            queryCondition.setBeginTime(beginDate.substring(0, 6));
            queryCondition.setEndTime(endDate.substring(0, 6));
        } else {
            String lastMonth = DateFormatUtils.format(DateUtils.addMonths(new Date(), -1), "yyyyMM");
            queryCondition.setBeginTime(lastMonth);
            queryCondition.setEndTime(lastMonth);
        }

        //循环获取商户号
        for (TMerchant tMerchant1 : tMerchantList) {
            logger.info("start {} merchant month statistics....", tMerchant1.getMerchantId());
            queryCondition.setMerchantId(tMerchant1.getMerchantId());

            List<TInstructionDayStatistics> tInstructionDayStatisticsList = tInstructionDayStatisticsService.getMapper().selectGroup(queryCondition);

            for (TInstructionDayStatistics tInstructionDayStatistics1 : tInstructionDayStatisticsList) {
                TInstructionMonthStatistics tInstructionMonthStatistics = new TInstructionMonthStatistics();
                tInstructionMonthStatistics.setTableName("t_instruction_month_statistics_m" + tMerchant1.getMerchantId());
                tInstructionMonthStatistics.setMerchantId(tMerchant1.getMerchantId());
                tInstructionMonthStatistics.setPosition(tInstructionDayStatistics1.getPosition());
                tInstructionMonthStatistics.setStuffId(tInstructionDayStatistics1.getStuffId());
                tInstructionMonthStatistics.setBusinessType(tInstructionDayStatistics1.getBusinessType());

                if (tInstructionDayStatistics1.getBusinessCount() != null) {
                    tInstructionMonthStatistics.setBusinessCount(tInstructionDayStatistics1.getBusinessCount());
                } else {
                    tInstructionMonthStatistics.setBusinessCount(0);
                }
                if (tInstructionDayStatistics1.getBusinessSubjectCount() != null) {
                    tInstructionMonthStatistics.setBusinessSubjectCount(tInstructionDayStatistics1.getBusinessSubjectCount());
                } else {
                    tInstructionMonthStatistics.setBusinessSubjectCount(0);
                }

                if (tInstructionDayStatistics1.getReceivableAmount() != null) {
                    tInstructionMonthStatistics.setReceivableAmount(tInstructionDayStatistics1.getReceivableAmount());
                } else {
                    tInstructionMonthStatistics.setReceivableAmount(new BigDecimal(0.00));
                }
                if (tInstructionDayStatistics1.getIncomeAmount() != null) {
                    tInstructionMonthStatistics.setIncomeAmount(tInstructionDayStatistics1.getIncomeAmount());
                } else {
                    tInstructionMonthStatistics.setIncomeAmount(new BigDecimal(0.00));
                }
                if (tInstructionDayStatistics1.getPayableAmount() != null) {
                    tInstructionMonthStatistics.setPayableAmount(tInstructionDayStatistics1.getPayableAmount());
                } else {
                    tInstructionMonthStatistics.setPayableAmount(new BigDecimal(0.00));
                }
                if (tInstructionDayStatistics1.getExpensesAmount() != null) {
                    tInstructionMonthStatistics.setExpensesAmount(tInstructionDayStatistics1.getExpensesAmount());
                } else {
                    tInstructionMonthStatistics.setExpensesAmount(new BigDecimal(0.00));
                }
                if (tInstructionDayStatistics1.getReceivableFeeAmount() != null) {
                    tInstructionMonthStatistics.setReceivableFeeAmount(tInstructionDayStatistics1.getReceivableFeeAmount());
                } else {
                    tInstructionMonthStatistics.setReceivableFeeAmount(new BigDecimal(0.00));
                }
                if (tInstructionDayStatistics1.getDiscountAmount() != null) {
                    tInstructionMonthStatistics.setDiscountAmount(tInstructionDayStatistics1.getDiscountAmount());
                } else {
                    tInstructionMonthStatistics.setDiscountAmount(new BigDecimal(0.00));
                }

                tInstructionMonthStatistics.setBillMonth(tInstructionDayStatistics1.getBillDate());

                try {
                    tInstructionMonthStatistics.setLastBillMonth(DateFormatUtils.format(DateUtils.addMonths(new SimpleDateFormat("yyyyMM").parse(tInstructionDayStatistics1.getBillDate()), -1), "yyyyMM"));
                } catch (ParseException e) {
                    e.printStackTrace();
                }

                tInstructionMonthStatistics.setJobSerial(tJobLog.getJobSerial());
                tInstructionMonthStatisticsService.getMapper().insertDynamic(tInstructionMonthStatistics);
            }
        }

        //3.更新事务流水
        tJobLog.setJobStatus("01");
        tJobLogService.getMapper().updateByPrimaryKey(tJobLog);
        logger.info("genMonthReportData finish......");
        return "交易成功";
    }
}
*/
