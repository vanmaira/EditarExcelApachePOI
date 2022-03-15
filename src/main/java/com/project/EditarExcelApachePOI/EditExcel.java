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

                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                HSSFSheet sheet = workbook.getSheetAt(0);
                int lastRowNo = sheet.getLastRowNum();

                for (int rownum = 2; rownum <= lastRowNo; rownum++) {
                }

                Row dados1Row = dados1.getRow(3);
                dados1Row.getCell((short) 1).setCellValue("VANESSAr");

                Row dados1Row2 = dados1.getRow(4);
                dados1Row2.getCell((short) 1).setCellValue("VANESSA");

                Row dados2Row = dados1.getRow(6);
                dados2Row.getCell((short) 1).setCellValue("NILSINHO");

                Row dados2Row4 = dados1.getRow(7);
                dados2Row4.getCell((short) 1).setCellValue("FELIPE");
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