import java.io.*;
import java.util.*;
import java.util.regex.*;
import javax.swing.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class TxtToExelConverter {
    public static void main(String[] args) {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("텍스트 파일을 선택하세요");
        int result = fileChooser.showOpenDialog(null);

        if (result != JFileChooser.APPROVE_OPTION) {
            JOptionPane.showMessageDialog(null, "입력 파일을 선택하지 않았습니다.");
            return;
        }

        File inputFile = fileChooser.getSelectedFile();
        String outputName = JOptionPane.showInputDialog(null, "생성할 엑셀 파일명을 입력하세요 (예: result.xlsx)", "ConvertResult.xlsx");

        if (outputName == null || outputName.trim().isEmpty()) {
            JOptionPane.showMessageDialog(null, "출력 파일명을 입력하지 않았습니다.");
            return;
        }

        if (!outputName.endsWith(".xlsx")) {
            outputName += ".xlsx";
        }

        File outputFile = new File(inputFile.getParentFile(), outputName);

        BufferedReader reader = null;
        FileOutputStream out = null;
        Workbook workbook = null;

        try {
            reader = new BufferedReader(new InputStreamReader(new FileInputStream(inputFile), "UTF-8"));
            workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Txt");

            CellStyle borderStyle = workbook.createCellStyle();
            borderStyle.setBorderTop(BorderStyle.THIN);
            borderStyle.setBorderBottom(BorderStyle.THIN);
            borderStyle.setBorderLeft(BorderStyle.THIN);
            borderStyle.setBorderRight(BorderStyle.THIN);

            int rowIndex = 0;

            Pattern timestampPattern = Pattern.compile("^\\[(\\d{4}-\\d{2}-\\d{2} \\d{2}:\\d{2}:\\d{2}\\.\\d+)"); // 시간
            Pattern dataPattern = Pattern.compile("\\[(.*?)\\]"); // 모든 대괄호 데이터
            Pattern lListPattern = Pattern.compile("^\\s*L\\[\\d+\\](\\[\\])?"); // 리스트 패턴

            List<String> dataBuffer = new ArrayList<>();
            String rawLine;

            while ((rawLine = reader.readLine()) != null) {
                String trimmedLine = rawLine.trim();
                int indentLevel = countIndentLevel(rawLine);

                Matcher timeMatcher = timestampPattern.matcher(trimmedLine);
                if (timeMatcher.find()) {
                    if (!dataBuffer.isEmpty()) {
                        Row row = sheet.createRow(rowIndex++);
                        for (int i = 0; i < dataBuffer.size(); i++) {
                            Cell cell = row.createCell(i);
                            cell.setCellValue(dataBuffer.get(i));
                            cell.setCellStyle(borderStyle);
                        }
                        dataBuffer.clear();
                    }
                    dataBuffer.add(timeMatcher.group(1)); // timestamp
                    continue;
                }

                Matcher lMatcher = lListPattern.matcher(trimmedLine);
                if (lMatcher.find() && indentLevel == 1) {
                    dataBuffer.add("List item");
                    continue;
                }

                if (indentLevel == 1) {
                    Matcher dataMatcher = dataPattern.matcher(trimmedLine);
                    while (dataMatcher.find()) {
                        dataBuffer.add(dataMatcher.group(1));
                    }
                }
            }

            if (!dataBuffer.isEmpty()) {
                Row row = sheet.createRow(rowIndex++);
                for (int i = 0; i < dataBuffer.size(); i++) {
                    Cell cell = row.createCell(i);
                    cell.setCellValue(dataBuffer.get(i));
                    cell.setCellStyle(borderStyle);
                }
            }

            out = new FileOutputStream(outputFile);
            workbook.write(out);
            JOptionPane.showMessageDialog(null, "✅ 엑셀 파일이 생성되었습니다:\n" + outputFile.getAbsolutePath());

        } catch (Exception e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null, "오류 발생: " + e.getMessage());
        } finally {
            try { if (reader != null) reader.close(); } catch (IOException ignored) {}
            try { if (out != null) out.close(); } catch (IOException ignored) {}
            try { if (workbook != null) workbook.close(); } catch (IOException ignored) {}
        }
    }

    private static int countIndentLevel(String line) {
        int count = 0;
        for (int i = 0; i < line.length(); i++) {
            char ch = line.charAt(i);
            if (ch == '\t') {
                count += 4;
            } else if (ch == ' ') {
                count++;
            } else {
                break;
            }
        }
        return count / 4;
    }
}
