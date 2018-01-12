import com.agile.api.*;
import com.agile.px.EventActionResult;
import com.agile.px.IEventAction;
import com.agile.px.IEventInfo;
import com.agile.px.IObjectEventInfo;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

public class updateBOMbyBatch implements IEventAction{
    @Override
    public EventActionResult doAction(IAgileSession session, INode Node, IEventInfo req) {

        String filepath = "C:\\ExcelFile\\OldAndNew.xlsx";
        try {
            IObjectEventInfo info = (IObjectEventInfo)req;
            IDataObject obj = info.getDataObject();
            IChange change = (IChange)obj;

            ArrayList OldNum_list = readExcel(filepath,session);
            System.out.println("Old Number List: "+OldNum_list);

        } catch (APIException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }


        return null;
    }

    private static ArrayList readExcel(String path,IAgileSession session) throws IOException {
        ArrayList<String> OldNum = new ArrayList<>();

        FileInputStream inp = new FileInputStream(path);
        XSSFWorkbook wb = new XSSFWorkbook(inp);                //讀取Excel
        XSSFSheet sheet = wb.getSheetAt(0);             //讀取wb內的頁面
        XSSFRow row = sheet.getRow(0);               //讀取頁面0的第一行
        int rowlength = sheet.getPhysicalNumberOfRows();       // number of row
        int columnlength = row.getPhysicalNumberOfCells();     // number of column

        for (int i=1;i<rowlength;i++){
            row = sheet.getRow(i);
            for(int j =0;j<2;j++){
                String excelCell = row.getCell(j)+"";
                OldNum.add(excelCell);
                System.out.println("Part Num: "+ excelCell);
            }
        }

        return OldNum;
    }

}
