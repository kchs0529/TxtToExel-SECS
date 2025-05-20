import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TxtToExelConverter {
    public static void main(String[] args) {
    	//JFileChooser 유니코드들은 한글이 깨지지 않게 처리하기 위함
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

            // 3줄 헤더 설정
            Row headerRow0 = sheet.createRow(0);
            Row headerRow1 = sheet.createRow(1);
            Row headerRow2 = sheet.createRow(2);

         // 헤더 추가
            String[] headers = new String[] {
                "Item Name", "OrientAngleOffset", "AutomatedAngleOffset", "XTiltHomeOffsetEC", "CarrierCompleteAlertTime", "ProcessStep", "WaferStartTime", "WaferEndTime", "WaferElapsedTime", "CarrierStartTime", "CarrierEndTime", "CarrierElapsedTime", "StepStartTime", "StepEndTime", "StepElapsedTime", "AMU", "CryoP2Temperature", "CryoP3Temperature", "CryoP4Temperature", "CryoP5Temperature", "DilutionGasFlow", "DIWaterTemperature", "DopantGasFlow", "ExtractionCurrent", "ExtractionSuppressionCurr", "ExtractionSuppressionVolt", "ExtractionVoltage", "FinalBeamEnergy", "FloodgunArcCurrent", "FloodgunGasFlow", "GasBottle1Pressure", "GasBottle2Pressure", "GasBottle3Pressure", "GasBottle4Pressure", "GasBottle5Pressure", "IG1Pressure", "IG2Pressure", "IG3Pressure", "ImplantAngle", "ImplantCurrent", "ImplantMapCurrent", "ImplantPercentComplete", "MainWaterTemperature", "OrientAngle", "ResolvingAperaturePos", "SetupCupBeamCurrent", "Sigma", "SourceBeamCurrent", "SourceMagnetCurrent", "SourceTuneTime", "Specie", "ThetaAxis", "TotalDose", "TotalTuneTime", "VaporizerTemperature", "WaferCooling", "WaferCurrentPass", "WaferCurrentStep", "WaferSteps", "WaferTotalPasses", "WaferTotalPassNumber", "Yaxis", "Zaxis", "AccelCurrent", "AccelSupprVolt", "AccelVoltage", "AnalyzerField1", "AnalyzerField2", "ArcCurrent1", "ArcCurrent2", "ArcVoltage1", "ArcVoltage2", "BeamEnergy", "Charge", "DecelVoltage", "FilamentCurrent1", "FilamentCurrent2", "FloodGunFilCurrent", "FocusCurrent", "IonMass", "MaxDose", "MeanDose", "MinDose", "NumberOfGlitches", "DIResistivity", "MFC1TotalFlow", "MFC2TotalFlow", "MFC3TotalFlow", "MFC4TotalFlow", "MFC5TotalFlow", "MFC6TotalFlow", "PFGGasTotalFlow", "DoseMapVelocity", "PFGHotTime", "PFGWarmTime", "PFGXeGasPressure", "WaferOnOrientor", "ChillerTemperature", "ChillerResistivity", "DoseTrimActual", "ImplantChamberPressure", "ImplantCurrentProcessStep", "ImplantTotalDose", "DoseStep1", "ImplantAngleStep1", "OrientAngleStep1", "AccelControllerVoltageProgramValue", "AccelSuppControllerVoltProgramValue", "IG1Filament1Life", "IG1Filament2Life", "IG2Filament1Life", "IG2Filament2Life", "IG3Filament1Life", "IG3Filament2Life", "IG4Filament1Life", "IG4Filament2Life", "IG5Filament1Life", "IG5Filament2Life", "IG7Filament1Life", "IG7Filament2Life"
            };
            
            // 첫 번째 줄: 항목 이름
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow0.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(borderStyle);
            }

            // 두 번째 줄: SVID
            String[] svids = { "SVID", "40", "41", "70", "94", "199", "221", "222", "223", "224", "225", "226", "227", "228", "229", "5010", "5180", "5190", "5200", "5210", "5230", "5240", "5250", "5270", "5280", "5290", "5300", "5320", "5330", "5340", "5350", "5360", "5370", "5380", "5390", "5400", "5410", "5420", "5430", "5440", "5470", "5480", "6050", "6060", "6190", "6750", "6760", "6800", "6810", "6820", "6830", "6840", "6850", "6860", "6900", "6910", "6930", "6940", "6950", "6960", "6970", "6980", "6990", "7500", "7510", "7520", "7530", "7540", "7550", "7560", "7570", "7580", "7590", "7600", "7630", "7640", "7650", "7680", "7690", "7710", "7720", "7730", "7740", "7750", "10160", "12500", "12510", "12520", "12530", "12540", "12550", "12560", "12580", "12590", "12600", "12630", "12970", "15270", "15280", "15300", "15340", "15350", "15360", "16980", "16990", "17000", "17030", "17040", "31040", "31050", "31060", "31070", "31080", "31090", "31100", "31110", "31120", "31130", "31140", "31150" };
            for (int i = 0; i < svids.length; i++) {
            	Cell cell = headerRow1.createCell(i);
                cell.setCellValue(svids[i]);
                cell.setCellStyle(borderStyle);
            }
            
            // 셀 스타일 정의 (줄바꿈 + 중앙 정렬)
            CellStyle cornerStyle = workbook.createCellStyle();
            cornerStyle.setWrapText(true);
            cornerStyle.setAlignment(HorizontalAlignment.CENTER);
            cornerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            // (2,0) 셀 = 3번째 줄 첫 번째 셀
            Cell diagonalCell = headerRow2.createCell(0);
            diagonalCell.setCellStyle(cornerStyle);
            diagonalCell.setCellValue("Unit\nTime");

            // 도형 대각선 + 셀 크기 조정
            sheet.setColumnWidth(0, 15 * 256);
            headerRow2.setHeightInPoints(40);

            XSSFDrawing drawing = ((XSSFSheet) sheet).createDrawingPatriarch();
            XSSFClientAnchor anchor = new XSSFClientAnchor(
                0, 0, 1023, 255,
                0, 2, 1, 3
            );
            anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);

            XSSFSimpleShape line = drawing.createSimpleShape(anchor);
            line.setShapeType(ShapeTypes.LINE);
            line.setLineWidth(1.0);

            // 세 번째 줄: 단위
            String[] units = {"deg", "deg", "deg", "sec", "", "cs", "cs", "s", "cs", "cs", "s", "cs", "cs", "s", "amu", "K", "K", "K", "K", "sccm", "C", "sccm", "mA", "mA", "kV", "kV", "keV", "", "sccm", "", "", "", "", "", "Torr", "Torr", "Torr", "deg", "", "", "%", "C", "cnts", "mm", "", "%", "","", "min","", "mm", "ions/cm2", "min", "C", "Torr", "cnts", "cnts", "cnts", "cnts", "cnts", "mm", "mm", "mA", "kV", "kV", "kGauss", "kGauss", "", "", "V", "V", "", "", "kV", "", "","", "mA", "amu", "%", "%", "%", "", "MOhmcm", "ml", "ml", "ml", "ml", "ml", "ml", "ml", "cm/s", "min", "min", "psig", "cnts", "C", "MOhmcm", "%", "Torr", "cnts", "ions/cm2", "ions/cm2", "deg", "deg", "kV", "kV", "day", "day", "day", "day", "day", "day", "day", "day", "day", "day", "day", "day" };
            for (int i = 0; i < units.length; i++) {
            	Cell cell = headerRow2.createCell(i+1);
                cell.setCellValue(units[i]);
                cell.setCellStyle(borderStyle);
            }

            // 데이터는 4번째 줄부터
            int rowIndex = 3;


//            Pattern timestampPattern = Pattern.compile("^\\[(\\d{4}-\\d{2}-\\d{2} \\d{2}:\\d{2}:\\d{2}\\.\\d+)\\s+S1F4V11\\]");
//            Pattern dataPattern = Pattern.compile("\\[(.*?)\\]");
//            Pattern lStreamPattern = Pattern.compile("^L\\[\\d+\\]\\[\\]");
//            Pattern lListPattern = Pattern.compile("^\\s*L\\[\\d+\\]");
            
            Pattern timestampPattern = Pattern.compile("^\\[(\\d{4}-\\d{2}-\\d{2} \\d{2}:\\d{2}:\\d{2}\\.\\d+)\\s+S1F4V11\\]");
            Pattern dataPattern = Pattern.compile("\\[(.*?)\\]");
            Pattern lListPattern = Pattern.compile("^\\s*L\\[\\d+\\](\\[\\])?");

            List dataBuffer = new ArrayList();
            String rawLine;
            int lCount = 0;
            boolean afterSecondL = false;
            boolean allowSingleIndent = false;

            while ((rawLine = reader.readLine()) != null) {
                String trimmedLine = rawLine.trim();

                Matcher timeMatcher = timestampPattern.matcher(trimmedLine);
                if (timeMatcher.find()) {
                    if (!dataBuffer.isEmpty()) {
                        Row row = sheet.createRow(rowIndex++);
                        for (int i = 0; i < dataBuffer.size(); i++) {
                            Cell cell = row.createCell(i);
                            cell.setCellValue((String) dataBuffer.get(i));
                            cell.setCellStyle(borderStyle);
                        }
                        dataBuffer.clear();
                    }
                    dataBuffer.add(timeMatcher.group(1));
                    lCount = 0;
                    afterSecondL = false;
                    allowSingleIndent = false;
                    continue;
                }

                Matcher lMatcher = lListPattern.matcher(trimmedLine);
                if (lMatcher.find()) {
                    lCount++;
                    if (lCount == 1) {
                        // 첫 번째 L[...] 또는 L[...][]는 무시
                        continue;
                    } else {
                    	// 두 번째  L[...] 또는 L[...][] 부터는 List item이 들어가도록 변경
                        dataBuffer.add("List item");
                        afterSecondL = true;
                        allowSingleIndent = true;
                        continue;
                    }
                }

                int indentLevel = countIndentLevel(rawLine);

                //1번 들여쓰기까지는 데이터 기록, 2번부터는 기록x
                if (afterSecondL) {
                    if (indentLevel == 1 && allowSingleIndent) {
                        // 한 번 들여쓰기까지만 허용
                    } else if (indentLevel >= 2) {
                        continue;
                    } else {
                        allowSingleIndent = false;
                    }
                }

                Matcher dataMatcher = dataPattern.matcher(trimmedLine);
                while (dataMatcher.find()) {
                    String value = dataMatcher.group(1);
                    dataBuffer.add(value);
                }
            }

            if (!dataBuffer.isEmpty()) {
                Row row = sheet.createRow(rowIndex++);
                for (int i = 0; i < dataBuffer.size(); i++) {
                    row.createCell(i).setCellValue((String) dataBuffer.get(i));
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