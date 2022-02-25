package com.study.exceluploader;

import com.fasterxml.jackson.databind.exc.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.*;

@Component
public class ExcelUtil {

    public String getCellValue(Cell cell) {

        String value = "";

        if(cell == null){
            return value;
        }

        switch (cell.getCellType()) {
            case 1:
                value = cell.getStringCellValue();
                break;
            case 0:
                value = (int) cell.getNumericCellValue() + "";
                break;
            case 2:
                value = null;
                break;
            default:
                break;
        }
        return value;
    }

    public List<Map<String, Object>> getListData(MultipartFile file) {

        List<Map<String, Object>> mapList = new ArrayList<Map<String,Object>>();

        try {
            OPCPackage opcPackage = OPCPackage.open(file.getInputStream());
            XSSFWorkbook workbook = new XSSFWorkbook(opcPackage);

            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> iteratorRow = sheet.iterator();

            int rowIdx = 0;
            ArrayList<String> keyList = new ArrayList<>();

            while (iteratorRow.hasNext()) {
                Row currentRow = iteratorRow.next();

                ArrayList<String> cells = new ArrayList<>();
                Iterator<Cell> cellsInRow = currentRow.cellIterator();
                int cellIdx = 0;
                Map<String, Object> map = new HashMap<String, Object>();
                while (cellsInRow.hasNext()) {

                    Cell currentCell = cellsInRow.next();
                    cells.add(getCellValue(currentCell));

                    if (rowIdx > 0) {
                        map.put(keyList.get(cellIdx), cells.get(cellIdx));
                    }
                    cellIdx++;
                }
                if (rowIdx == 0) {
                    keyList = cells;
                } else mapList.add(map);
                rowIdx++;
            }

        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException | org.apache.poi.openxml4j.exceptions.InvalidFormatException e) {
            e.printStackTrace();
        }

        return mapList;
    }
}
