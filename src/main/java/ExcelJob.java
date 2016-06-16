import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by hyar on 6/16/2016.
 */
class ExcelJob {

    private List<CellObject> cellObjects;

    ExcelJob() {
        cellObjects = new ArrayList<CellObject>();
    }

    boolean loadExcelSheet(File workbookFile, int sheetIndex) {
        try {
            // todo: check if it's readonly
            Workbook wb = WorkbookFactory.create(workbookFile);
            Sheet sheet = wb.getSheetAt(sheetIndex);
            readExcelSheet(sheet);
            return true;

        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
    }

    private void readExcelSheet(Sheet sheet) {
        for (Row row : sheet) {
            makeTestEntries(row);
        }
    }

    private void makeTestEntries(Row row) {
        CellObject cellObject;

        if ((row.getCell(0) != null) && !row.getCell(0).toString().equals("")) {
            cellObject = new CellObject();

            cellObject.outCome = row.getCell(0).toString();
            cellObject.rerun = row.getCell(1).toString();
            cellObject.testCaseTitle = row.getCell(2).toString();
            cellObjects.add(cellObject);
        }
    }

    List<CellObject> getPairedCredentialsMap() {
        return cellObjects;
    }

}