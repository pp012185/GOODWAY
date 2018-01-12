import com.agile.api.*;
import com.agile.px.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;

public class checkExcelFile implements IEventAction{
    @Override
    public EventActionResult doAction(IAgileSession session, INode Node, IEventInfo req) {

        System.out.println("------Start------");
        String filepath = "C:\\ExcelFile\\OldAndNew.xlsx";
        try {
            IObjectEventInfo info = (IObjectEventInfo)req;
            IDataObject obj = info.getDataObject();
            IChange change = (IChange)obj;


            System.out.println("Download");
            downloadExcelFile(session,change);
            boolean A =  readExcel(filepath,session);
            System.out.println("Result: "+ A);

            return new EventActionResult(req,new ActionResult(0,"Success: "));
        } catch (Exception e) {
            e.printStackTrace();
            return new EventActionResult(req,new ActionResult(ActionResult.EXCEPTION,new Exception("Excel error")));
        }
//        catch (IOException e) {
//            e.printStackTrace();
//            return new EventActionResult(req,new ActionResult(ActionResult.EXCEPTION,new Exception("Excel error")));
//        }



    }

    private static boolean readExcel(String path,IAgileSession session) throws IOException, APIException {
        boolean allinSystem=false;

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
                System.out.println("Part Num: "+ excelCell);
                IItem item = (IItem) session.getObject(IItem.OBJECT_TYPE, excelCell);
                System.out.println(item.getValue(ItemConstants.ATT_TITLE_BLOCK_NUMBER));
            }
        }
        allinSystem = true;
        return allinSystem;
    }

    private void downloadExcelFile(IAgileSession session,IChange change) throws APIException, IOException {
        InputStream ins = null;
        try {
            // 找檔案的table
            ITable attach_tb = change.getTable(ChangeConstants.TABLE_ATTACHMENTS);
            Iterator it = attach_tb.iterator();
            if (it.hasNext()) {
                IRow row = (IRow) it.next();
                ins = ((IAttachmentFile) row).getFile();   // 丟到inputstream
            }
        } catch (RuntimeException getFileEx) {
            /*
             * 紀錄RuntimeException Message
             */
            System.out.println("\t" + getFileEx.toString());
            System.out.println(getFileEx.toString());
            // For Error Mail Notification
            String getFileRuntimeExceptionMsg = getFileEx.toString();
            throw getFileEx;
        } catch (Exception ioEx) {
            System.out.println("\t Non-RuntimeException: " + ioEx.toString());
            System.out.println("Non-RuntimeException: " + ioEx.toString());
            throw ioEx;
        }
        System.out.println("\t Get File Input Stream.");
        System.out.println("Get File Input Stream.");

        /*
         * 指定文件附件檔案的傳送路徑(<PLM_TEMP_FILE目錄路徑>\<文件號>_<版本>)
         */

        String localFileFolderPath="C:\\ExcelFile";
        String fileName ="OldAndNew.xlsx";
        File file = new File(localFileFolderPath );
        String filePath = file.getPath();
        System.out.println("\t Set File Path.");
        System.out.println("Set File Path.");

        /*
         * 輸出檔案到中介檔案庫
         */
        // fileName要包含副檔名

        FileOutputStream fos = new FileOutputStream(filePath + "\\" + fileName);
        System.out.println("\t File Output Stream Ready.");
        System.out.println("File Output Stream Ready.");
        byte[] b = new byte[2048];
        int off = 0;
        int len = 0;
        while ((len = ins.read(b)) != -1) {
            fos.write(b, off, len);
        }
        System.out.println("\t Copy File Done.");
        System.out.println("Copy File Done.");

    }


}
