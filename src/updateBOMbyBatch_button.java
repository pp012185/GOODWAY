import com.agile.api.*;
import com.agile.px.ActionResult;
import com.agile.px.ICustomAction;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class updateBOMbyBatch_button implements ICustomAction{
    @Override
    public ActionResult doAction(IAgileSession session, INode Node, IDataObject obj) {

        String filepath = "C:\\ExcelFile\\OldAndNew.xlsx";
        try {
            System.out.println("------Start------");
            IChange change = (IChange) obj;
            ArrayList WhereUsedList = getWhereUsedList(filepath,session);          // build Where Used list
            System.out.println("Where Used List: "+WhereUsedList);
            AddToChange(change,WhereUsedList);
            //Iterator it = WhereUsedList.iterator();
            //while (it.hasNext()){

            //}


            System.out.println("--------End--------");
        } catch (IOException e) {
            e.printStackTrace();
        } catch (APIException e) {
            e.printStackTrace();
        }


        return new ActionResult(0,"update BOM button success!");
    }


    /***********************************************************************
        * 讀Excel, 建立OldNum_list, 丟到buildWhereUsedList method, 得到WhereUsed_list
        * *********************************************************************/
    private static ArrayList getWhereUsedList(String path,IAgileSession session) throws IOException, APIException {
        ArrayList<String> OldNum_list = new ArrayList<>();

        FileInputStream inp = new FileInputStream(path);
        XSSFWorkbook wb = new XSSFWorkbook(inp);                //讀取Excel
        XSSFSheet sheet = wb.getSheetAt(0);             //讀取wb內的頁面
        XSSFRow row = sheet.getRow(0);               //讀取頁面0的第一行
        int rowlength = sheet.getPhysicalNumberOfRows();       // number of row
        int columnlength = row.getPhysicalNumberOfCells();     // number of column

        for (int i=1;i<rowlength;i++){
            row = sheet.getRow(i);
            int j = 0;
            String excelCell = row.getCell(j)+"";
            OldNum_list.add(excelCell);
            System.out.println("Excel Old Num: "+ excelCell);
        }
        ArrayList<?> WhereUsed_list = buildWhereUsedList(OldNum_list,session);


        return WhereUsed_list;
    }
    private static ArrayList buildWhereUsedList(ArrayList OldNum_list,IAgileSession session) throws APIException {
        ArrayList<IItem> WhereUsed_list = new ArrayList<>();
        Iterator it1 = OldNum_list.iterator();
        while (it1.hasNext()){
            String OldNum = it1.next().toString();
            System.out.println("OldNum: "+OldNum);
            IItem olditem = (IItem) session.getObject(IItem.OBJECT_TYPE, OldNum);
            ITable WhereUsed_tb=olditem.getTable(ItemConstants.TABLE_WHEREUSED);
            Iterator it2 = WhereUsed_tb.iterator();
            while(it2.hasNext()){
                IRow row =(IRow) it2.next();
                IItem oldparentitem = (IItem)row.getReferent();
                System.out.println("Parent item number:"+oldparentitem.getName());
                if(!WhereUsed_list.contains(oldparentitem)) WhereUsed_list.add(oldparentitem);
            }
        }
        return WhereUsed_list;
    }

    /***********************************************************************
        * 把 WhereUsed_list 的item 加入表單的Affected item  => 之後要加到次表單內
        * *********************************************************************/
    private static void AddToChange(IChange change, ArrayList WhereUsed_list) throws APIException {

        ITable affecteditem_tb = change.getTable(ChangeConstants.TABLE_AFFECTEDITEMS);
        Iterator it = WhereUsed_list.iterator();
        while(it.hasNext()) {
            IItem item = (IItem)it.next();
            IRow affectedrow = affecteditem_tb.createRow(item);
        }
    }



    /***********************************************************************
        * 開立次表單，和其workflow
        * *********************************************************************/


}
