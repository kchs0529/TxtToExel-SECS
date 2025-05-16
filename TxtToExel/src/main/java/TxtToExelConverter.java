import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.*;
import java.util.*;
import java.util.regex.*;

public class TxtToExelConverter {
    public static void main(String[] args) {
        // 1. 입력 파일 선택
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("FDC 텍스트 파일을 선택하세요");
        int result = fileChooser.showOpenDialog(null);

        if (result != JFileChooser.APPROVE_OPTION) {
            JOptionPane.showMessageDialog(null, "입력 파일을 선택하지 않았습니다.");
            return;
        }

        File inputFile = fileChooser.getSelectedFile();

        // 2. 출력 파일 이름 입력
        String outputName = JOptionPane.showInputDialog(null, "생성할 엑셀 파일명을 입력하세요 (예: result.xlsx)", "FDC_Parsed.xlsx");

        if (outputName == null || outputName.trim().isEmpty()) {
            JOptionPane.showMessageDialog(null, "출력 파일명을 입력하지 않았습니다.");
            return;
        }

        // .xlsx 확장자 보장
        if (!outputName.endsWith(".xlsx")) {
            outputName += ".xlsx";
        }

        File outputFile = new File(inputFile.getParentFile(), outputName); // 같은 폴더에 저장

        // 3. 변환 처리 시작
        BufferedReader reader = null;
        FileOutputStream out = null;
        Workbook workbook = null;

        try {
            reader = new BufferedReader(new InputStreamReader(new FileInputStream(inputFile), "UTF-8"));
            workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("FDC Data");

            Pattern timestampPattern = Pattern.compile("^\\[(\\d{4}-\\d{2}-\\d{2} \\d{2}:\\d{2}:\\d{2}\\.\\d+).*\\]$");
            Pattern dataPattern = Pattern.compile("\\[(.*?)\\]");

            String line;
            List<String> dataBuffer = new ArrayList<String>();
            int rowIndex = 0;

            while ((line = reader.readLine()) != null) {
                // L[119][] 같은 줄은 무시
                if (line.trim().startsWith("L[119]")) continue;

                Matcher timestampMatcher = timestampPattern.matcher(line);
                if (timestampMatcher.find()) {
                    if (!dataBuffer.isEmpty()) {
                        Row row = sheet.createRow(rowIndex++);
                        for (int col = 0; col < dataBuffer.size(); col++) {
                            row.createCell(col).setCellValue(dataBuffer.get(col));
                        }
                        dataBuffer.clear();
                    }
                    dataBuffer.add(timestampMatcher.group(1));
                } else {
                    Matcher dataMatcher = dataPattern.matcher(line);
                    while (dataMatcher.find()) {
                        String value = dataMatcher.group(1);
                        dataBuffer.add(value);
                    }
                }
            }

            // 마지막 줄 저장
            if (!dataBuffer.isEmpty()) {
                Row row = sheet.createRow(rowIndex++);
                for (int col = 0; col < dataBuffer.size(); col++) {
                    row.createCell(col).setCellValue(dataBuffer.get(col));
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
}
