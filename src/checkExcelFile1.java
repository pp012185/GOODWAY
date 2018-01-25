import com.agile.api.*;
import com.agile.px.*;
import com.anselm.plm.utilobj.Ini;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;

public class checkExcelFile1 implements IEventAction,ICustomAction{
    static final Ini ini = new Ini("C:\\Agile\\Config.ini");
    static final int NumOfexcelrow = Integer.valueOf(ini.getValue("parameter","NumOfexcelrow")) ;
    static final int NumOfBatch = Integer.valueOf(ini.getValue("parameter","NumOfBatch"));
    static final String localpath = ini.getValue("path","localpath");
    static final String fileName = ini.getValue("Name","fileName");
    @Override
    public EventActionResult doAction(IAgileSession session, INode Node, IEventInfo req) {


        System.out.println("------Start------");

        try {
            IObjectEventInfo info = (IObjectEventInfo)req;
            IDataObject obj = info.getDataObject();
            IChange change = (IChange)obj;
            String filepath = localpath+"\\"+change.getName().toString()+"_"+fileName;

            System.out.println("Download");
            downloadExcelFile(session,change,localpath,fileName);

            boolean A =  readExcel(filepath,session);

            // Get the Part Category cell
            ICell cell = change.getCell(ChangeConstants.ATT_PAGE_THREE_LIST01);
            // Get available list values for Part Category
            IAgileList values = cell.getAvailableValues();
            // Set the value to Electrical
            values.setSelection(new Object[] { "Yes" });
            cell.setValue(values);

            System.out.println("Result: "+ A);



            return new EventActionResult(req,new ActionResult(ActionResult.STRING,"Success: "));
        } catch (Exception e) {
            e.printStackTrace();
            return new EventActionResult(req,new ActionResult(ActionResult.EXCEPTION,new Exception("Excel error")));
        }
//        catch (IOException e) {
//            e.printStackTrace();
//            return new EventActionResult(req,new ActionResult(ActionResult.EXCEPTION,new Exception("Excel error")));
//        }



    }

    private static boolean readExcel(String path,IAgileSession session) throws IOException, APIException ,Exception{
        boolean allinSystem=false;
        System.out.println("read path: "+path);
        FileInputStream inp = new FileInputStream(path);
        XSSFWorkbook wb = new XSSFWorkbook(inp);                //讀取Excel
        XSSFSheet sheet = wb.getSheetAt(0);             //讀取wb內的頁面
        XSSFRow row = sheet.getRow(0);               //讀取頁面0的第一行
        int rowlength = sheet.getPhysicalNumberOfRows();       // number of row
        int columnlength = row.getPhysicalNumberOfCells();     // number of column

        System.out.println("excel row: "+ rowlength);
        System.out.println("config row: "+NumOfexcelrow);

        if((rowlength-1)==NumOfexcelrow){
            for (int i=1;i<(NumOfexcelrow+1);i++){
                row = sheet.getRow(i);
                for(int j =0;j<2;j++){
                    String excelCell = row.getCell(j)+"";
                    System.out.println("Part Num: "+ excelCell);
                    IItem item = (IItem) session.getObject(IItem.OBJECT_TYPE, excelCell);
                    System.out.println(item.getValue(ItemConstants.ATT_TITLE_BLOCK_NUMBER));
                }
            }
            inp.close();
            allinSystem = true;
        }else {
            inp.close();
            allinSystem = false;
        throw new Exception("Excel Error");
        }

        return allinSystem;
    }

    private void downloadExcelFile(IAgileSession session,IChange change,String path,String name) throws Exception {
        InputStream ins = null;
        String changeName = change.getName().toString();
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

        String localFileFolderPath = path;
        String fileName = changeName+"_"+name;
        File file = new File(localFileFolderPath);
        String filePath = file.getPath();
        System.out.println("filePath:"+filePath);
        System.out.println("fileName:"+fileName);
        System.out.println("\t Set File Path.");
        System.out.println("Set File Path.");

        /*
         * 輸出檔案到中介檔案庫
         */
        // fileName要包含副檔名

        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(filePath + "\\" +fileName);
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


        } catch (IOException e) {
            e.printStackTrace();
            throw new Exception();


        } finally {
            if(fos!=null) fos.close();
            if(ins!=null) ins.close();
        }

    }


    @Override
    public ActionResult doAction(IAgileSession session, INode iNode, IDataObject obj) {

        System.out.println("------Start------");
        try {
            String result = "";
            IChange change = (IChange)obj;

            String filepath = localpath+"\\"+change.getName().toString()+"_"+fileName;

            System.out.println("Download");
            downloadExcelFile(session,change,localpath,fileName);

            boolean A =  readExcel(filepath,session);
            System.out.println("Result: "+ A);

            if(A){
                // Get the Part Category cell
                ICell cell = change.getCell(ChangeConstants.ATT_PAGE_THREE_LIST01);
                // Get available list values for Part Category
                IAgileList values = cell.getAvailableValues();
                // Set the value to Electrical
                values.setSelection(new Object[] { "Yes" });
                cell.setValue(values);
                result ="Success";
            }else {
                result ="Excel error";
            }



            return new ActionResult(0,"Status: "+result);
        } catch (Exception e) {
            e.printStackTrace();
            return new ActionResult(ActionResult.EXCEPTION,new Exception("Excel error"));
        }
//        catch (IOException e) {
//            e.printStackTrace();
//            return new EventActionResult(req,new ActionResult(ActionResult.EXCEPTION,new Exception("Excel error")));
//        }

    }
}
