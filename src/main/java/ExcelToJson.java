import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.codehaus.jackson.map.ObjectMapper;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.json.simple.JSONObject;

public class ExcelToJson {
    public static void main(String[] args) throws IOException {

        // Step 1: Read Excel File into Java List Objects
        List students = readExcelFile("F:\\Excel2Json\\src\\main\\StudentRecords.xlsx");

        // Step 2: Write Java List Objects to JSON File
        writeObjects2JsonFile(students, "F:\\Excel2Json\\src\\main\\Output.json");

        System.out.println("Done");
    }
    private static List readExcelFile(String filePath)
    {
        try
        {
            FileInputStream excelFile = new FileInputStream(new File(filePath));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheetAt(0);
            Iterator rows = sheet.iterator();
            List lstStudents = new ArrayList();
            int rowNumber = 0;
            while (rows.hasNext())
            {
                Row currentRow = (Row) rows.next();
                // skip header
                if(rowNumber == 0)
                {
                    rowNumber++;
                    continue;
                }

                Iterator cellsInRow = currentRow.iterator();
                Student stud = new Student();
                int cellIndex = 0;
                while (cellsInRow.hasNext())
                {
                    Cell currentCell = (Cell) cellsInRow.next();

                    if(cellIndex==0)
                    {
                        // Student Name
                        stud.setName(currentCell.getStringCellValue());
                    } else if(cellIndex==1)
                    {
                        // Student Age
                        stud.setAge((int) currentCell.getNumericCellValue());
                    } else if(cellIndex==2)
                    {
                        // Student TotalMarks
                        stud.setTotalMarks((int) currentCell.getNumericCellValue());
                    }
                    cellIndex++;
                }

                lstStudents.add(stud);
            }

            workbook.close();

            return lstStudents;
        } catch (IOException e)
        {
            throw new RuntimeException("FAIL! -> message = " + e.getMessage());
        }
    }
    private static void writeObjects2JsonFile(List student, String pathFile) throws IOException {

        ObjectMapper mapper = new ObjectMapper();

        String jsonString = "";

        try {

            jsonString = mapper.writeValue("Student.json",student);
            //mapper.writerWithDefaultPrettyPrinter().writeValueAsString(student);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
