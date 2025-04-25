package org.qiugul.accounting;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jetbrains.annotations.NotNull;
import org.qiugul.common.NumberUtils;
import org.qiugul.common.StringUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.FileVisitResult;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.attribute.BasicFileAttributes;
import java.text.SimpleDateFormat;
import java.util.*;

public class XlsxCleaner {

    private final static String INCOME_HEADER = "交易类型,日期,分类,子分类,收入账户,金额,成员,商家,项目,备注";

    private final static String OUTCOME_HEADER = "交易类型,日期,分类,子分类,支出账户,金额,成员,商家,项目,备注";


    public static void main(String[] args) throws IOException {
        XlsxCleaner xlsxCleaner = new XlsxCleaner();
        List<String> filePathList = new ArrayList<>();
        filePathList.add("C:\\tmp\\origin");
        xlsxCleaner.process(filePathList, "C:\\tmp\\output1.xlsx");
        // xlsxCleaner.process("D:\\xhren\\data\\smallpdf1.xlsx", "D:\\xhren\\data\\output1.xlsx");
    }


    public void f(String folderPath) throws IOException {
        Path start = Paths.get(folderPath);

        Files.walkFileTree(start, new SimpleFileVisitor<Path>() {
            @Override
            public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) {
                System.out.println("文件: " + file);
                return FileVisitResult.CONTINUE;
            }

            @Override
            public FileVisitResult preVisitDirectory(Path dir, BasicFileAttributes attrs) {
                System.out.println("进入目录: " + dir);
                return FileVisitResult.CONTINUE;
            }

            @Override
            public FileVisitResult visitFileFailed(Path file, IOException exc) {
                System.err.println("访问文件失败: " + file);
                return FileVisitResult.CONTINUE;
            }
        });
    }


    /**
     * 食品酒水：早午晚餐，烟酒茶，水果零食
     * 行车交通：公共交通，打车租车，私家车费用
     * 居家物业：日常用品，水电煤气，房租，物业管理，维修保养，
     * 交流通讯：手机费，上网费，邮寄费
     * 衣服饰品：衣服裤子，鞋帽包包，化妆饰品
     * 休闲娱乐：运动健身，交际聚会，休闲玩乐，宠物宝贝，旅游度假
     * 医疗保健：药品费，保健费，美容费，治疗费
     * 学习进修：数码装备，培训进修
     * 人情往来：送礼请客，孝敬长辈，还人钱物，慈善捐助
     * 金融保险：银行手续，投资亏损，按揭还款，消费税收，利息支出，赔偿罚款
     * 其他杂项：其他支出，意外损失，烂账损失
     *
     * @param filePathList
     * @param outputPath
     * @throws IOException
     */
    public void process(List<String> filePathList, String outputPath) throws IOException {

        List<List<String>> rows = new ArrayList<>();
        for (String filePath : filePathList) {
            List<List<String>> res = readXlsx(filePath);
            rows.addAll(res);
        }


        XSSFWorkbook outWorkbook = new XSSFWorkbook();
        Sheet income = createSheet(outWorkbook, "收入", INCOME_HEADER);
        Sheet outcome = createSheet(outWorkbook, "支出", OUTCOME_HEADER);
        Sheet transfer = createSheet(outWorkbook, "转账", "交易类型,日期,转出账户,转入账户,金额,成员,商家,项目,备注");

        int incomeRowCnt = 1;
        int outcomeRowCnt = 1;
        int transferRowCnt = 1;
        String headerStr = "Date,Currency,Transaction,Balance,TransactionType,CounterParty";
        Map<String, Integer> oldCol2Idx = generateHeaderMap(headerStr);
        for (int i = 0; i < rows.size(); i++) {
            List<String> oldRow = rows.get(i);

            String transaction = oldRow.get(oldCol2Idx.get("Transaction"));
            float transactionValue = NumberUtils.toFloat(transaction, 0);
            String transactionType = oldRow.get(oldCol2Idx.get("TransactionType"));
            String txType = getType(transactionValue, transactionType);
            Sheet curSheet = null;
            int curRowCnt = 1;
            if (Objects.equals("Transfer", txType)) {
                curSheet = transfer;
                curRowCnt = transferRowCnt++;
            } else if (Objects.equals("Income", txType)) {
                curSheet = income;
                curRowCnt = incomeRowCnt++;
            } else if (Objects.equals("Outcome", txType)) {
                curSheet = outcome;
                curRowCnt = outcomeRowCnt++;
            }
            if (curSheet == null) {
                throw new RuntimeException("no sheet");
            }

            Row newRow = curSheet.createRow(curRowCnt);

            if (Objects.equals("Transfer", txType)) {
                processTransfer(oldCol2Idx, newRow, oldRow);
            } else if (Objects.equals("Income", txType)) {
                processIncome(oldCol2Idx, newRow, oldRow);
            } else if (Objects.equals("Outcome", txType)) {
                processOutcome(oldCol2Idx, newRow, oldRow);
            }

        }

        try (FileOutputStream out = new FileOutputStream(outputPath)) {
            outWorkbook.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static List<List<String>> readXlsx(String filePath) throws IOException {
        Path start = Paths.get(filePath);

        List<List<String>> res = new ArrayList<>();

        Files.walkFileTree(start, new SimpleFileVisitor<Path>() {
            @Override
            public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) {
                res.addAll(doReadXlsx(file.toString()));
                return FileVisitResult.CONTINUE;
            }

            @Override
            public FileVisitResult preVisitDirectory(Path dir, BasicFileAttributes attrs) throws IOException {
                return FileVisitResult.CONTINUE;
            }

            @Override
            public FileVisitResult visitFileFailed(Path file, IOException exc) {
                System.err.println("访问文件失败: " + file);
                return FileVisitResult.CONTINUE;
            }
        });

        return res;
    }

    private static List<List<String>> doReadXlsx(String filePath) {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        List<List<String>> rows = new ArrayList<>();

        // 读取Excel文件
        try (FileInputStream fis = new FileInputStream(filePath); Workbook workbook = new XSSFWorkbook(fis)) {
// 遍历每个Sheet
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
// 单sheet, 不用处理
                Sheet sheet = workbook.getSheetAt(i);
                if (sheet == null) {
                    continue;
                }
// 遍历每一行
                for (int j = 0; j <= sheet.getLastRowNum(); j++) {
                    Row row = sheet.getRow(j);
                    if (!isValidRow(row)) {
                        continue;
                    }
// 遍历每一列
                    List<String> curRow = readRowAsStringList(row, sdf);
                    rows.add(curRow);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return rows;
    }

    private void processTransfer(Map<String, Integer> oldCol2Idx, Row newRow, List<String> oldRow) {
        Map<String, Integer> newCol2IdxMap =
                generateHeaderMap("交易类型,日期,转出账户,转入账户,金额,成员,商家,项目,备注");
        boolean negative = false;
        for (Map.Entry<String, Integer> entry : oldCol2Idx.entrySet()) {
            String oldCol = entry.getKey();
            Integer oldIdx = entry.getValue();
            // Date,Currency,Transaction,Balance,TransactionType,CounterParty
            if (Objects.equals(oldCol, "Date")) {
                String newSheetRow = "日期";
                Cell cell = newRow.createCell(newCol2IdxMap.get(newSheetRow));
                cell.setCellValue(oldRow.get(oldIdx));
            } else if (Objects.equals(oldCol, "Transaction")) {
                String newSheetRow = "金额";
                Cell cell = newRow.createCell(newCol2IdxMap.get(newSheetRow));
                String str = oldRow.get(oldIdx);
                if (str.startsWith("-")) {
                    str = str.substring(1);
                    negative = true;
                }
                cell.setCellValue(str);
            } else if (Objects.equals(oldCol, "CounterParty")) {
                String newSheetRow = "备注";
                Cell cell = newRow.createCell(newCol2IdxMap.get(newSheetRow));
                String str = oldRow.get(oldIdx);
                cell.setCellValue(str);
            }
        }

// 账户
        String from = "";
        String to = "";
        String other = getOtherAccount(oldRow.get(oldCol2Idx.get("CounterParty")));
        if (negative) {
            from = "招商银行";
            to = other;
        } else {
            to = "招商银行";
            from = other;
        }
        Cell cell = newRow.createCell(newCol2IdxMap.get("转出账户"));
        cell.setCellValue(from);
        cell = newRow.createCell(newCol2IdxMap.get("转入账户"));
        cell.setCellValue(to);

        cell = newRow.createCell(newCol2IdxMap.get("备注"));
        cell.setCellValue(oldCol2Idx.get("CounterParty"));


        cell = newRow.createCell(newCol2IdxMap.get("交易类型"));
        cell.setCellValue("转账");
    }

    private void processIncome(Map<String, Integer> oldCol2Idx, Row newRow, List<String> oldRow) {
        Map<String, Integer> newCol2IdxMap = generateHeaderMap(INCOME_HEADER);

        // String type = Const.EMPTY;
        // String date = Const.EMPTY;
        // String category = Const.EMPTY;
        // String subCategory = Const.EMPTY;
        // String to = Const.EMPTY;
        // String amount = Const.EMPTY;
        // String member = Const.EMPTY;
        // String seller = Const.EMPTY;
        // String project = Const.EMPTY;
        // String remark = Const.EMPTY;


        for (Map.Entry<String, Integer> entry : oldCol2Idx.entrySet()) {
            String oldCol = entry.getKey();
            Integer oldIdx = entry.getValue();
            // Date,Currency,Transaction,Balance,TransactionType,CounterParty
            if (Objects.equals(oldCol, "Date")) {
                String newSheetRow = "日期";
                Cell cell = newRow.createCell(newCol2IdxMap.get(newSheetRow));
                cell.setCellValue(oldRow.get(oldIdx));
            } else if (Objects.equals(oldCol, "Transaction")) {
                String newSheetRow = "金额";
                Cell cell = newRow.createCell(newCol2IdxMap.get(newSheetRow));
                String str = oldRow.get(oldIdx);
                cell.setCellValue(str);
            } else if (Objects.equals(oldCol, "CounterParty")) {
                String newSheetRow = "备注";
                Cell cell = newRow.createCell(newCol2IdxMap.get(newSheetRow));
                String str = oldRow.get(oldIdx);
                cell.setCellValue(str);
            }
        }

// 账户
        String to = "招商银行";
        String from = getOtherAccount(oldRow.get(oldCol2Idx.get("CounterParty")));

        Cell cell = newRow.createCell(newCol2IdxMap.get("商家"));
        cell.setCellValue(from);
        cell = newRow.createCell(newCol2IdxMap.get("收入账户"));
        cell.setCellValue(to);

        cell = newRow.createCell(newCol2IdxMap.get("备注"));
        cell.setCellValue(oldCol2Idx.get("CounterParty"));


        cell = newRow.createCell(newCol2IdxMap.get("交易类型"));
        cell.setCellValue("收入");

        String type = "职业收入";
        String subType = "工资收入";
        cell = newRow.createCell(newCol2IdxMap.get("分类"));
        cell.setCellValue(type);
        cell = newRow.createCell(newCol2IdxMap.get("子分类"));
        cell.setCellValue(subType);


    }

    private void processOutcome(Map<String, Integer> oldCol2Idx, Row newRow, List<String> oldRow) {
        Map<String, Integer> newCol2IdxMap = generateHeaderMap(OUTCOME_HEADER);


        for (Map.Entry<String, Integer> entry : oldCol2Idx.entrySet()) {
            String oldCol = entry.getKey();
            Integer oldIdx = entry.getValue();
            // Date,Currency,Transaction,Balance,TransactionType,CounterParty
            if (Objects.equals(oldCol, "Date")) {
                String newSheetRow = "日期";
                Cell cell = newRow.createCell(newCol2IdxMap.get(newSheetRow));
                cell.setCellValue(oldRow.get(oldIdx));
            } else if (Objects.equals(oldCol, "Transaction")) {
                String newSheetRow = "金额";
                Cell cell = newRow.createCell(newCol2IdxMap.get(newSheetRow));
                String str = oldRow.get(oldIdx);
                cell.setCellValue(str);
            } else if (Objects.equals(oldCol, "CounterParty")) {
                String newSheetRow = "备注";
                Cell cell = newRow.createCell(newCol2IdxMap.get(newSheetRow));
                String str = oldRow.get(oldIdx);
                cell.setCellValue(str);
            }
        }

// 账户
        String from = "招商银行";
        String to = getOtherAccount(oldRow.get(oldCol2Idx.get("CounterParty")));

        Cell cell = newRow.createCell(newCol2IdxMap.get("商家"));
        cell.setCellValue(to);
        cell = newRow.createCell(newCol2IdxMap.get("支出账户"));
        cell.setCellValue(from);

        cell = newRow.createCell(newCol2IdxMap.get("备注"));
        cell.setCellValue(oldCol2Idx.get("CounterParty"));


        cell = newRow.createCell(newCol2IdxMap.get("交易类型"));
        cell.setCellValue("支出");
        String type = "食品酒水";
        String subType = "早午晚餐";
        cell = newRow.createCell(newCol2IdxMap.get("分类"));
        cell.setCellValue(type);
        cell = newRow.createCell(newCol2IdxMap.get("子分类"));
        cell.setCellValue(subType);
    }

    @NotNull
    private static List<String> readRowAsStringList(Row row, SimpleDateFormat sdf) {
        List<String> curRow = new ArrayList<>();
        for (int k = 0; k < row.getLastCellNum(); k++) {
            Cell cell = row.getCell(k);
            String res = readCellValueAsString(cell, sdf);
            curRow.add(res);
        }
        return curRow;
    }

    private static String readCellValueAsString(Cell cell, SimpleDateFormat sdf) {
        if (cell == null) {
            return "cellNull";
        }
        String res = "valueNull";
        if (cell.getCellType().equals(CellType.STRING)) {
            res = cell.getStringCellValue();
        } else if (cell.getCellType().equals(CellType.NUMERIC)) {
            if (DateUtil.isCellDateFormatted(cell)) {
                Date dateCellValue = cell.getDateCellValue();
                res = sdf.format(dateCellValue);
            } else {
                res = String.valueOf(cell.getNumericCellValue());
            }
        }
        return res;
    }

    private String getOtherAccount(String counterParty) {
// ,华宝证券,东方财富,微信钱包,支付宝余额,南阳银行,兴业银行,现金
        String value = "小荷包,交通卡,招商证券,朝朝宝,阿里云,华宝证券,携程,哈啰,优衣库,蛙三疯,霸碗,陈香贵,天空之城,盖饭湘,楚褚热干面,7-11,东方财富,支付宝余额,南阳商业银行,兴业银行,"
                + "现金,太平洋证券,鹰角,盒马,木鸢网络,网银在线,微信,上海市电力,同花顺,乡村基,美团,董俊宏";
        String[] keywords = StringUtils.split(value, ",");

        for (String keyword : keywords) {
            if (counterParty.contains(keyword)) {
                return keyword;
            }
        }
        if (counterParty.contains("任晓辉")) {
            if (counterParty.contains("622908213240682313")) {
                return "兴业银行";
            } else if (counterParty.contains("6235553020000803010")) {
                return "南阳商业银行";
            } else if (counterParty.contains("6225766640537066")) {
                return "招商信用卡";
            }
        }

        String s = "待报解预算收入:国库,华泰证券（上海）资产管理有限公司 433874017351:基金,上海和闵房产有限公司:乐贤居,（上云10主机）:基金,基金:基金,理财:基金,太平洋健康保险:太平洋健康保险,金城速汇:金城速汇,陈荣:陈荣,跨境支付退款:跨境支付退款";
        Map<String, String> key2Account = new TreeMap<>(Comparator.naturalOrder());
        String[] split = StringUtils.split(s, ",");
        for (String item : split) {
            String[] each = StringUtils.split(item, ":");
            key2Account.put(each[0], each[1]);
        }

        for (Map.Entry<String, String> entry : key2Account.entrySet()) {
            if (counterParty.contains(entry.getKey())) {
                return entry.getValue();
            }
        }

        return counterParty;

    }

    @NotNull
    private static Map<String, Integer> generateHeaderMap(String headerStr) {
        String[] headers = StringUtils.split(headerStr, ",");
        Map<String, Integer> header2Idx = new HashMap<>();
        for (int i = 0; i < headers.length; i++) {
            header2Idx.put(headers[i], i);
        }
        return header2Idx;
    }

    private String getType(float transactionValue, String transactionType) {
        Map<String, List<String>> type2KeyworkMap = getTypeMap();

        String[] matchPriority = {"Transfer", "Income", "Outcome"};
        for (String type : matchPriority) {
            List<String> keywords = type2KeyworkMap.get(type);
            if (keywords == null) {
                continue;
            }
            boolean contains = keywords.stream().anyMatch(transactionType::contains);
            if (contains) {
                return type;
            }
        }
        if (transactionValue > 0) {
            return "Income";
        } else {
            return "Outcome";
        }

// String type;
        // List<String> transferTypeList = getTransferTypeList();
        // boolean contains = transferTypeList.stream().anyMatch(transactionType::contains);
        // if (contains) {
        // type = "Transfer";
        // } else if (transactionValue > 0) {
        // type = "Income";
        // } else {
        // type = "Outcome";
        // }
        // return type;
    }

    private List<String> getTransferTypeList() {
        List<String> list = new ArrayList<>();
        list.add("转账");
        list.add("汇款");
        list.add("基金");
        list.add("朝朝宝");
        list.add("赎回");
        return list;
    }

    private Map<String, List<String>> getTypeMap() {
        String value = "Transfer:转账|汇款|基金|朝朝宝|赎回";
        Map<String, List<String>> map = new HashMap<>();
        for (String each : StringUtils.split(value, ",")) {
            String[] pair = StringUtils.split(each, ":");
            if (pair.length != 2) {
                continue;
            }
            String type = pair[0];
            String[] keywords = StringUtils.split(pair[1], "|");
            map.computeIfAbsent(type, k -> new ArrayList<>()).addAll(Arrays.asList(keywords));
        }
        return map;
    }

    private static Sheet createSheet(XSSFWorkbook outWorkbook, String sheetName, String headerStr) {
        Sheet sheet = outWorkbook.createSheet(sheetName);
        Row header = sheet.createRow(0);
        String[] headerList = StringUtils.split(headerStr, ",");
        for (int i = 0; i < headerList.length; i++) {
            Cell cell = header.createCell(i);
            cell.setCellValue(headerList[i]);
        }
        return sheet;
    }

    private static boolean isValidRow(Row row) {
        try {
            if (row == null) {
                return false;
            }
            boolean validRow = false;
            short lastCellNum = row.getLastCellNum();
            if (lastCellNum >= 1) {
                Cell cell = row.getCell(0);
                if (cell == null) {
                    return false;
                }
                return cell.getCellType().equals(CellType.NUMERIC) && DateUtil.isCellDateFormatted(cell);
            }
            return validRow;
        } catch (Exception e) {
            return false;
        }

    }
}