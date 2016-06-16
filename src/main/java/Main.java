import java.io.File;

public class Main {
    public static final String FILE_LOCATION = "C:\\Users\\hyar\\Desktop\\tests.xlsx";

    public static void main(String[] args) {
        File file = new File(FILE_LOCATION);
        ExcelJob excelJob = new ExcelJob();
        boolean isLoaded = excelJob.loadExcelSheet(file, 1);
        System.out.println("done");
    }
}
