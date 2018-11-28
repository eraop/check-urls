package com.eraop;

import com.eraop.model.RowModel;
import com.zx.lib.javamail.EumMailPriority;
import com.zx.lib.javamail.MailAccountInfo;
import com.zx.lib.javamail.MailMessageInfo;
import com.zx.lib.javamail.MailSender;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.mail.internet.InternetAddress;
import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @author weiyi
 */
public class App {
    private static boolean flag = true;

    public static void main(String[] args) {
        // 监测文件的完整路径 excel文件路径
        String filePath = System.getProperty("path");
        // filePath = "C:\\Users\\weiyi\\Desktop\\check urls\\test.xlsx";
        System.out.println(filePath);
        // 每条数据监测间隔时间（例：1000~3000  即1秒~3秒）
        String interval = System.getProperty("interval");
        System.out.println(interval);
        // 启动定时抓取(按天抓取)
        String timer = System.getProperty("timer");
        System.out.println(timer);
        // 开始抓取日期
        String startTimeStr = System.getProperty("startTime");
        System.out.println(startTimeStr);

        // 设置默认值
        interval = StringUtils.isNotEmpty(interval) ? interval : "1000-3000";
        timer = (StringUtils.isNotEmpty(timer) && timer.toLowerCase().equals("y")) ? "y" : "n";

        SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
        Date startTime = new Date();
        try {
            if (StringUtils.isNotEmpty(startTimeStr)) {
                startTime = df.parse(startTimeStr);
            }
        } catch (ParseException e) {
            System.out.println("日期格式不正确，已经默认使用当时时间");
        }
        // 获得开始时间
        startTime = handleDate(startTime);
        System.out.println(startTime.toString());

        if (StringUtils.isNotEmpty(filePath)) {
            if (interval.contains("-")) {
                int start = Integer.parseInt(interval.split("-")[0]);
                int end = Integer.parseInt(interval.split("-")[1]);
                // 将开始时间作为上轮监测时间
                Date checkRoundTime = startTime;
                if ("y".equals(timer.toLowerCase())) {
                    //开启定时 如果 开始时间 <= 当前时间 则抓取
                    // if (startTime.getTime() <= handleDate(new Date()).getTime()) {
                    //     handle(filePath, start, end + 1);
                    //     // 将当前时间作为上轮监测时间
                    //     checkRoundTime = handleDate(new Date());
                    // }
                    while (true) {
                        // 当前日期(天) >= 监测时间（天）
                        System.out.println(checkRoundTime.toString());
                        System.out.println(handleDate(new Date()).toString());
                        if (handleDate(new Date()).getTime() >= checkRoundTime.getTime()) {
                            // 将当前时间加一天作为监测时间
                            checkRoundTime = handleDate(new Date(), 1);
                            handle(filePath, start, end + 1);
                            System.out.println("该轮监测任务结束" + new Date().toString());
                        } else {
                            try {
                                System.out.println("轮询监视中......" + new Date().toString());
                                Thread.sleep(3600 * 1000);
                            } catch (InterruptedException e) {
                                e.printStackTrace();
                                System.out.println("线程异常退出......" + new Date().toString());
                            }
                        }
                    }
                } else {
                    handle(filePath, start, end + 1);
                }
                System.out.println("监测任务结束");
            } else {
                System.out.println("时间间隔错误");
            }

        } else {
            System.out.println("文件路径不能为空");
        }
    }

    private static void handle(String excelPath, int start, int end) {
        String logPath = "";
        String logName = "";
        Random random = new Random();
        List<RowModel> list = new ArrayList<>();
        try {
            Date date = new Date();
            SimpleDateFormat df = new SimpleDateFormat("yyyy_MM_dd");
            String time = df.format(date);
            System.out.println("开始验证" + date.toString());
            File excel = new File(excelPath);
            logName = excel.getName() + " - " + time + ".log";
            logPath = excel.getParent() + "/logs/" + logName;
            clearInfoForFile(logPath);
            //判断文件是否存在
            if (excel.isFile() && excel.exists()) {

                //.是特殊字符，需要转义！！！！！
                String[] split = excel.getName().split("\\.");
                Workbook workbook;
                //根据文件后缀（xls/xlsx）进行判断
                FileInputStream fis = new FileInputStream(excel);
                if ("xls".equals(split[1])) {
                    //文件流对象
                    workbook = new HSSFWorkbook(fis);
                } else if ("xlsx".equals(split[1])) {
                    workbook = new XSSFWorkbook(fis);
                } else {
                    System.out.println("文件类型错误!");
                    return;
                }
                fis.close();

                //开始解析
                //读取sheet 0
                Sheet sheet = workbook.getSheetAt(0);
                //第一行是列名，所以不读
                int firstRowIndex = sheet.getFirstRowNum() + 1;
                int lastRowIndex = sheet.getLastRowNum();
                //遍历行
                for (int rIndex = firstRowIndex; rIndex <= lastRowIndex; rIndex++) {
                    Row row = sheet.getRow(rIndex);
                    RowModel rowModel = new RowModel();
                    if (row != null) {
                        //遍历列
                        if (row.getCell(0) != null && StringUtils.isNotEmpty(row.getCell(0).toString())) {
                            if (row.getCell(0) != null) {
                                row.getCell(0).setCellType(Cell.CELL_TYPE_STRING);
                                rowModel.setId(row.getCell(0).toString());
                            }
                            if (row.getCell(1) != null) {
                                rowModel.setUserid(row.getCell(1).toString());
                            }
                            if (row.getCell(2) != null) {
                                rowModel.setUsername(row.getCell(2).toString());
                            }
                            if (row.getCell(3) != null) {

                                rowModel.setUrl(row.getCell(3).toString());
                            }
                            if (row.getCell(4) != null) {
                                if (!row.getCell(4).toString().equals("是") && !row.getCell(4).toString().equals("推文链接不正确")) {
                                    if (rowModel.getUrl().trim().startsWith("http")) {
                                        int code = check(rowModel.getUrl().trim());
                                        if (code == 200) {
                                            rowModel.setResult("否");
                                        } else if (code == 404) {
                                            rowModel.setResult("是");
                                        } else {
                                            rowModel.setResult("未知");
                                        }
                                    } else {
                                        rowModel.setResult("推文链接不正确");
                                    }
                                    row.getCell(4).setCellValue(rowModel.getResult());
                                }
                            }
                            write(logPath, rowModel);
                            try {
                                int sleepTime = random.nextInt(end - start) + start;
                                Thread.sleep(sleepTime);
                            } catch (InterruptedException e) {
                                e.printStackTrace();
                            }
                        }
                    }
                }
                System.out.println("结束验证" + date.toString());
                // 发送邮件通知
                List<String> attachments = new ArrayList<>();
                attachments.add(logPath);
                // sendMailTimer("今日Twitter数据监测结果（" + logName + "）", attachments);
                sendMail("今日Twitter数据监测结果111111（" + logName + "）", attachments);
            } else {
                System.out.println("找不到指定的文件");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    /**
     * 指定时间发送邮件
     *
     * @param title
     * @param attachments
     */
//     private static void sendMailTimer(final String title, final List<String> attachments) {
//         Calendar calendar = Calendar.getInstance();
//         /*
//          * 指定触发的时间  9:30
//          */
//         calendar.set(Calendar.HOUR_OF_DAY, 9);
//         calendar.set(Calendar.MINUTE, 30);
//         calendar.set(Calendar.SECOND, 0);
//         Date time = calendar.getTime();
//         Timer timer = new Timer();
//         timer.schedule(new TimerTask() {
//             @Override
//             public void run() {
//                 System.out.println("邮件开始发送");
//                 sendMail(title, attachments);
//             }
//         }, time);
//
//     }
    private static void sendMail(String title, List<String> attachments) {
        try {

            MailAccountInfo accountInfo = new MailAccountInfo();
            accountInfo.setDomain("qq.com");
            accountInfo.setHost("smtp.exmail.qq.com");
            accountInfo.setPort(465);
            accountInfo.setLoginid("weiyi@safefw.com");
            accountInfo.setPassword("");
            accountInfo.setProtocol("smtp");
            accountInfo.setSsl(true);
            MailMessageInfo msg = new MailMessageInfo();

            msg.setFrom(new InternetAddress("weiyi@safefw.com"));
            List<InternetAddress> list = new ArrayList<>();
            list.add(new InternetAddress("jiyuning@cnzxsoft.com"));
            list.add(new InternetAddress("2210019617@qq.com"));
            list.add(new InternetAddress("weiyi@safefw.com"));
            list.add(new InternetAddress("493214262@qq.com"));
            msg.setReceives(list);
            msg.setSubject(title);
            msg.setContent(title);
            msg.setPriority(EumMailPriority.HIGH);

            msg.setAttachments(attachments);
            msg.setSentDate(new Date());

            MailSender sender = new MailSender(accountInfo, true, 30000);
            sender.connect();
            sender.sendEmail(msg);
            System.out.println("邮件发送成功");
        } catch (Exception e) {
            System.out.println("邮件发送失败");
        }
    }

    private String getValue(Cell hssfCell) {
        if (hssfCell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
            // 返回布尔类型的值
            return String.valueOf(hssfCell.getBooleanCellValue());
        } else if (hssfCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
            // 返回数值类型的值
            Object inputValue = null;
            Long longVal = Math.round(hssfCell.getNumericCellValue());
            Double doubleVal = hssfCell.getNumericCellValue();
            if (Double.parseDouble(longVal + ".0") == doubleVal) {
                inputValue = longVal;
            } else {
                inputValue = doubleVal;
            }
            DecimalFormat df = new DecimalFormat("#");
            return String.valueOf(df.format(inputValue));
        } else {
            // 返回字符串类型的值
            return String.valueOf(hssfCell.getStringCellValue());
        }
    }

    private static void clearInfoForFile(String fileName) {
        File file = new File(fileName);
        File p = new File(file.getParent());
        try {
            if (!p.exists()) {
                p.mkdirs();
            }
            if (!file.exists()) {
                file.createNewFile();
            }

            FileWriter fileWriter = new FileWriter(file);
            fileWriter.write("序号;推特用户ID;用户名;推文链接;是否已删除（只填是或否）;监测时间" + "\r\n");
            fileWriter.flush();
            fileWriter.close();
        } catch (Exception e) {
            System.out.println("初始化日志文件异常");
        }
    }

    public static void write(String fileName, RowModel model) {
        //使用相对路径，日志文件在项目内
        //获得当前系统时间
        Date date = new Date();
        SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        String time = df.format(date);
        File file = new File(fileName);
        Writer out = null;
        try {
            //使用字符流
            out = new FileWriter(file, true);
            //注意反斜杠的方向，/r是回车，/n是换行
            out.write(model.getId() + ";" + model.getUserid() + ";" + model.getUsername() + ";" + model.getUrl() + ";" + model.getResult() + ";" + time + "\r\n");
            out.flush();
            out.close();
            System.out.println(model.getId());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static int check(String address) {
        int status = 0;
        try {
            URL urlObj = new URL(address);
            HttpURLConnection oc = (HttpURLConnection) urlObj.openConnection();
            oc.setUseCaches(false);
            oc.setConnectTimeout(3000);
            status = oc.getResponseCode();
            // 200是请求地址顺利连通。。
        } catch (Exception ignored) {
        }
        return status;

    }

    public static String scanner(String tip) {
        Scanner scanner = new Scanner(System.in);
        scanner.useDelimiter("\n");
        StringBuilder help = new StringBuilder();
        help.append(tip);
        System.out.println(help.toString());
        if (scanner.hasNext()) {
            String ipt = scanner.next();
            if (StringUtils.isNotEmpty(ipt)) {
                return ipt;
            }
        }
        return "";
    }

    /**
     * 处理时间
     *
     * @param date
     * @return
     */
    private static Date handleDate(Date date) {
        return handleDate(date, 0);
    }

    private static Date handleDate(Date date, int addDay) {
        Calendar cal1 = Calendar.getInstance();
        cal1.setTime(date);

        cal1.add(Calendar.DAY_OF_MONTH, addDay);
        // 将时分秒,毫秒域清零
        cal1.set(Calendar.HOUR_OF_DAY, 0);
        cal1.set(Calendar.MINUTE, 0);
        cal1.set(Calendar.SECOND, 0);
        cal1.set(Calendar.MILLISECOND, 0);
        return cal1.getTime();
    }


}



