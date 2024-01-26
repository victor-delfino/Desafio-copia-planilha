import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class EstudanteGrade {

    private static final String FILE_PATH = "C:\\Users\\Claudia\\Downloads\\Cópia de Engenharia de Software - Desafio [Victor Hugo].xlsx";  // Substitua pelo caminho do seu arquivo

    public static void main(String[] args) {
        try {
            FileInputStream fileInputStream = new FileInputStream(FILE_PATH);
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            processStudentGrades(workbook);

            FileOutputStream fileOutputStream = new FileOutputStream("c:\\Users\\Claudia\\Downloads\\Cópia de Engenharia de Software - Desafio [Victor Hugo]Resultado.xlsx");  // Nome do arquivo de saída
            workbook.write(fileOutputStream);

            fileInputStream.close();
            fileOutputStream.close();
            workbook.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void processStudentGrades(Workbook workbook) {
        Sheet sheet = workbook.getSheetAt(0);  // Assumindo que os dados estão na primeira planilha

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            double p1 = row.getCell(1).getNumericCellValue();
            double p2 = row.getCell(2).getNumericCellValue();
            double p3 = row.getCell(3).getNumericCellValue();
            int totalClasses = (int) row.getCell(4).getNumericCellValue();
            int attendedClasses = (int) row.getCell(5).getNumericCellValue();

            double average = (p1 + p2 + p3) / 3;
            int totalFails = totalClasses / 4;  // 25% de faltas

            if (attendedClasses < totalFails) {
                setStudentSituation(row, "Reprovado por Falta");
            } else {
                if (average < 5) {
                    setStudentSituation(row, "Reprovado por Nota");
                } else if (average < 7) {
                    double naf = calculateNaf(average);
                    setStudentSituationAndNaf(row, "Exame Final", naf);
                } else {
                    setStudentSituationAndNaf(row, "Aprovado", 0);
                }
            }
        }
    }

    private static void setStudentSituation(Row row, String situation) {
        Cell situationCell = row.createCell(6);
        situationCell.setCellValue(situation);

        Cell nafCell = row.createCell(7);
        nafCell.setCellValue(0);  // Nota para Aprovação Final
    }

    private static void setStudentSituationAndNaf(Row row, String situation, double naf) {
        Cell situationCell = row.createCell(6);
        situationCell.setCellValue(situation);

        Cell nafCell = row.createCell(7);
        nafCell.setCellValue(Math.ceil(naf));  // Arredondar para o próximo número inteiro (aumentar) se necessário
    }

    private static double calculateNaf(double average) {
        return Math.max(5, 2 * 7 - average);
    }
}
