import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.omg.IOP.TAG_RMI_CUSTOM_MAX_STREAM_FORMAT;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

public class WordtestMaker {
    private static int TARGET_TEST_FIRST_PAGE_CELL_ROW_NUM = 1;
    public static int TARGET_TEST_COVERAGE_MIN;
    public static int TARGET_TEST_COVERAGE_MAX;
    public static int TARGET_TEST_WORD_NUM;
    public List<List> wordList = new ArrayList<>();

    public WordtestMaker (int min, int max, int wordNum){
        TARGET_TEST_COVERAGE_MIN = min;
        TARGET_TEST_COVERAGE_MAX = max;
        TARGET_TEST_WORD_NUM = wordNum;
    }

    public List<List> getWordList() {
        return wordList;
    }

    public void setWordDB(){
        XSSFSheet databaseSheet = ExcelLoader.workbook.getSheet("DB");
        Iterator<Row> rowIterator = databaseSheet.iterator();
        try {
            while (rowIterator.hasNext()) {
                XSSFRow row = (XSSFRow) rowIterator.next();
                // 주차 index
                int week = (int) row.getCell(0).getNumericCellValue();
                if (week < TARGET_TEST_COVERAGE_MIN )
                    continue;
                else if( week > TARGET_TEST_COVERAGE_MAX )
                    break;

                Iterator<Cell> cellIterator = row.cellIterator();
                List<Object> value = new ArrayList<Object>();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType()) {
                        case NUMERIC:
                            value.add((int)cell.getNumericCellValue());
                            break;

                        case STRING:
                            value.add(cell.getStringCellValue());
                            break;

                        case BLANK:
                            break;
                        default:
                            //System.out.print(cell.getCellType());
                            break;
                    }
                }
                // 첫 번째 column (주차) 삭제
                value.remove(0);
                wordList.add(value);
            }
        } catch (NullPointerException e){
        }
    }


}
