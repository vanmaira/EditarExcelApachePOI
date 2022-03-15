package com.project.EditarExcelApachePOI;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

import java.io.*;

public class EditExcel {


    private static final String fileName = "src/main/resources/def.xls";

    public static void main(String[] args) throws IOException {

        try {
            HSSFWorkbook workbook;
            try (FileInputStream file = new FileInputStream(new File(fileName))) {
                workbook = new HSSFWorkbook(file);
                HSSFSheet dados1 = workbook.getSheetAt(0);
                HSSFSheet dados2 = workbook.getSheetAt(1);

                Row dados1Row = dados1.getRow(4);
                dados1Row.getCell((short) 1).setCellValue(123);

                Row dados1Row2 = dados1.getRow(5);
                dados1Row2.getCell((short) 1).setCellValue(456);

              /*  Row dados2Row1 = dados2.getRow(3);
                dados2Row1.getCell((short) 1).setCellValue(789);

                Row dados2Row2 = dados2.getRow(4);
                dados2Row2.getCell((short) 1).setCellValue(1022);*/

                file.close();

                FileOutputStream outFile = new FileOutputStream(new File(fileName));
                workbook.write(outFile);

                outFile.close();
                System.out.println("Arquivo Excel editado com sucesso!");
            }

        } catch (FileNotFoundException e) {
            System.out.println("Arquivo Excel não encontrado!");
        } catch (IOException e) {
            System.out.println("Erro na edição do arquivo!");
        }
    }
}