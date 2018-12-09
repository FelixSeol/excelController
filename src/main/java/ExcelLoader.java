import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Collections;
import java.util.List;
import java.util.Random;

public class ExcelLoader {
    public static final String FILE_IN_NAME = "C:\\Users\\felixseol\\Downloads\\wordbook.xlsx";
    public static final String FILE_OUT_NAME = "C:\\Users\\felixseol\\Downloads\\wordbookout.xlsx";
    public static final File FILE_IN = new File(FILE_IN_NAME);
    public static final File FILE_OUT = new File(FILE_OUT_NAME);
    public static final int FIRSTPAGE = 1;
    public static final int SECONDPAGE = 32;
    private static String TARGET_TEST_SHEET_NAME = "시험";

    private static FileInputStream io;
    public static XSSFWorkbook workbook;

    public static void TestBuild(){
        WordtestMaker maker1 = new WordtestMaker(9, 9, 30);
        maker1.setWordDB();
        testsheetBuild(maker1, FIRSTPAGE);

        WordtestMaker maker2 = new WordtestMaker(1, 9, 30);
        maker2.setWordDB();
        testsheetBuild(maker2, SECONDPAGE);
    }

    public static void testsheetBuild(WordtestMaker maker, int startIdxPerPage){
        List<List> wordList = maker.getWordList();
        try {
            FileOutputStream out = new FileOutputStream(ExcelLoader.FILE_OUT);
            XSSFSheet targetSheet = ExcelLoader.workbook.getSheet(TARGET_TEST_SHEET_NAME);

            long seed = System.nanoTime();
            Collections.shuffle(wordList, new Random(seed));
            for (int i = 0; i < maker.TARGET_TEST_WORD_NUM; i++) {
                // 시험지 i행
                XSSFRow row = targetSheet.getRow(startIdxPerPage + i);
                Cell cell = row.getCell(2);

                // in Memory wordlist, 셔플 되어있다.
                List word = wordList.get(i);
                // word (히라가나, 음, 한자)를 섞어준다.
                Collections.shuffle(word, new Random(seed));
                // 그 중 첫번째 값을 넣어줌.
                cell.setCellValue(String.valueOf(word.get(0)));
                System.out.println(cell.getStringCellValue());
            }

            ExcelLoader.workbook.write(out);
            out.close();
        }catch (Exception e ){
            e.printStackTrace();
        }
    }
    public static void main (String[] args) {
        try {
            io = new FileInputStream(new File(FILE_IN_NAME));
            workbook = new XSSFWorkbook(io);
            ExcelLoader.TestBuild();
            io.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
