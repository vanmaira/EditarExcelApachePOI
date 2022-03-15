package com.project.EditarExcelApachePOI;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;

public class BckFile {
    public static void main(String[] args) {

        File sourceFile=new File("src/main/resources/template.xls");
        File destinationFile=new File("src/main/resources/def.xls");
        try{
            Files.copy(sourceFile.toPath(), destinationFile.toPath());
        } catch (IOException e){
            e.printStackTrace();
        }
    }
}
