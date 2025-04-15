package org.qiugul.accounting;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jetbrains.annotations.NotNull;
import org.qiugul.common.NumberUtils;
import org.qiugul.common.StringUtils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Pattern;

public class XlsxCleaner {

    public static void main(String[] args) {
        XlsxCleaner xlsxCleaner = new XlsxCleaner();
        xlsxCleaner.process("C:\\tmp\\smallpdf1.xlsx", "C:\\tmp\\output1.xlsx");
    }


    public void process(String filePath, String outputPath) {
        String headerStr = "Date,Currency,Transaction,Balance,TransactionType,CounterParty";
        Map<String, Integer> oldCol2Idx = generateHeaderMap(headerStr);
        List<List<String>> rows = new ArrayList<>();
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        // 读取Excel文件
        try (FileInputStream fis = new FileInputStream(filePath); Workbook workbook = new XSSFWorkbook(fis)) {

            Pattern p = Pattern.compile("^\\d{4}/(?:1[0-2]|[1-9])/(?:[12][0-9]|3[01]|[1-9])$");
            // 遍历每个Sheet
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                // 遍历每一行
                for (int j = 0; j <= sheet.getLastRowNum(); j++) {
                    Row row = sheet.getRow(j);
                    if (row != null) {
                        boolean validRow = isValidRow(p, row);
                        if (validRow) {
                            // 遍历每一列
                            List<String> curRow = new ArrayList<>();
                            for (int k = 0; k < row.getLastCellNum(); k++) {
                                Cell cell = row.getCell(k);
                                if (cell != null) {
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

                                    curRow.add(res);
                                } else {
                                    curRow.add("cellNull");
                                }
                            }
                            rows.add(curRow);
                        }
                    }
                }
            }


        } catch (IOException e) {
            e.printStackTrace();
        }
        XSSFWorkbook outWorkbook = new XSSFWorkbook();
        Sheet income = createSheet(outWorkbook, "收入", "交易类型,日期,分类,子分类,收入账户,金额,成员,商家,项目,备注");
        Sheet outcome = createSheet(outWorkbook, "支出", "交易类型,日期,分类,子分类,支出账户,金额,成员,商家,项目,备注");
        Sheet transfer = createSheet(outWorkbook, "转账", "交易类型,日期,转出账户,转入账户,金额,成员,商家,项目,备注");

        int incomeRowCnt = 1;
        int outcomeRowCnt = 1;
        int transferRowCnt = 1;
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
                Map<String, Integer> newCol2IdxMap = generateHeaderMap("交易类型,日期,转出账户,转入账户,金额,成员,商家,项目,备注");
                boolean negative = false;
                for (Map.Entry<String, Integer> entry : oldCol2Idx.entrySet()) {
                    String oldCol = entry.getKey();
                    Integer oldIdx = entry.getValue();
                    // Date,Currency,Transaction,Balance,TransactionType,CounterParty
                    if (Objects.equals(oldCol, "Date")){
                        String newSheetRow = "日期";
                        Cell cell = newRow.createCell(newCol2IdxMap.get(newSheetRow));
                        cell.setCellValue(oldRow.get(oldIdx));
                    }else if (Objects.equals(oldCol, "Transaction")){
                        String newSheetRow = "金额";
                        Cell cell = newRow.createCell(newCol2IdxMap.get(newSheetRow));
                        String str = oldRow.get(oldIdx);
                        if (str.startsWith("-")){
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

            } else if (Objects.equals("Income", txType)) {
                curSheet = income;
            } else if (Objects.equals("Outcome", txType)) {
                curSheet = outcome;
            }

        }

        try (FileOutputStream out = new FileOutputStream(outputPath)) {
            outWorkbook.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private String getOtherAccount(String counterParty) {
        // ,华宝证券,东方财富,微信钱包,支付宝余额,南阳银行,兴业银行,现金
        String value = "小荷包,交通卡,招商证券,朝朝宝,阿里云,华宝证券,携程,哈啰,优衣库,蛙三疯,霸碗,陈香贵,天空之城,微信红包,盖饭湘,楚褚热干面,7-11,东方财富,微信钱包,支付宝余额,南阳银行,兴业银行,现金";
        String[] keywords = StringUtils.split(value, ",");

        for (String keyword : keywords) {
            if (counterParty.contains(keyword)){
                return keyword;
            }
        }
        if (counterParty.contains("任晓辉")){
            if (counterParty.contains("622908213240682313")){
                return "兴业银行";
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
        String type;
        List<String> transferTypeList = getTransferTypeList();
        boolean contains = transferTypeList.stream().anyMatch(transactionType::contains);
        if (contains) {
            type = "Transfer";
        } else if (transactionValue > 0) {
            type = "Income";
        } else {
            type = "Outcome";
        }
        return type;
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

    private static boolean isValidRow(Pattern p, Row row) {
        boolean validRow = false;
        if (row == null) {
            return false;
        }
        short lastCellNum = row.getLastCellNum();
        if (lastCellNum >= 1) {
            Cell cell = row.getCell(0);
            if (cell == null) {
                return false;
            }
            if (!cell.getCellType().equals(CellType.NUMERIC) || !DateUtil.isCellDateFormatted(cell)) {
                return false;
            }
            return true;
        }
        return validRow;
    }
}