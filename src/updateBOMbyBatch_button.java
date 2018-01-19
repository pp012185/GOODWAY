import com.agile.api.*;
import com.agile.px.ActionResult;
import com.agile.px.ICustomAction;
import com.anselm.plm.utilobj.Ini;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.reflect.Array;
import java.util.*;

public class updateBOMbyBatch_button implements ICustomAction{


    @Override
    public ActionResult doAction(IAgileSession session, INode Node, IDataObject obj) {

        Ini ini = new Ini("D:\\Agile\\Agile936\\config.ini");
        int NumOfexcelrow = Integer.valueOf(ini.getValue("parameter","NumOfexcelrow")) ;
        int NumOfBatch = Integer.valueOf(ini.getValue("parameter","NumOfBatch"));
        String localpath = ini.getValue("path","localpath");
        String fileName = ini.getValue("Name","fileName");

        String filepath = localpath+"\\"+fileName ;
        int batchSize = NumOfBatch;
        try {
            System.out.println("------Start------");
            IChange change = (IChange) obj;

            // build Where Used list
            ArrayList WhereUsedList = getWhereUsedList(filepath,session);
            System.out.println("Where Used List: "+WhereUsedList);

            // create mapping
            HashMap map = getMapping(filepath);
            System.out.println("Map: "+ map);

            System.out.println("WhereUsedList length: "+ WhereUsedList.size());
            int WhereUsedListlength = WhereUsedList.size();
            int NumOfsubChange = (WhereUsedListlength/batchSize) +1;
            System.out.println("NumOfsubChange: "+NumOfsubChange);
            int j =0;


            for(int i=0;i<NumOfsubChange;i++){

                // create sub change
                IChange change2 = CreateSubChange(session);

                // parent item add to sub change affected item tab
                AddToChange(change2,WhereUsedList,j,batchSize);
                j=j+batchSize;

                //改affected item 的redline table
                updateRedlineBOM(change2,session,map);

                // add sub change to main change relationship tab
                ITable reltionship_tb = obj.getTable(ChangeConstants.TABLE_RELATIONSHIPS);
                reltionship_tb.createRow(change2);
            }

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
                System.out.println("Lifecycle phase: "+oldparentitem.getValue(ItemConstants.ATT_TITLE_BLOCK_LIFECYCLE_PHASE).toString());
                String phase = oldparentitem.getValue(ItemConstants.ATT_TITLE_BLOCK_LIFECYCLE_PHASE).toString();
                //System.out.println(phase.equals("EOL") || phase.equals("Pre EOL"));
                if(!WhereUsed_list.contains(oldparentitem)  &&  !(phase.equals("EOL") || phase.equals("Pre EOL"))  ) WhereUsed_list.add(oldparentitem);
            }
        }
        return WhereUsed_list;
    }


    /***********************************************************************
        * 建立新舊料號的HashMap
        * *********************************************************************/
    private static HashMap getMapping(String path) throws IOException {
        Map<String,String> map = new HashMap<>();

        FileInputStream inp = new FileInputStream(path);
        XSSFWorkbook wb = new XSSFWorkbook(inp);                //讀取Excel
        XSSFSheet sheet = wb.getSheetAt(0);             //讀取wb內的頁面
        XSSFRow row = sheet.getRow(0);               //讀取頁面0的第一行
        int rowlength = sheet.getPhysicalNumberOfRows();       // number of row
        int columnlength = row.getPhysicalNumberOfCells();     // number of column

        for (int i=1;i<rowlength;i++){
            row = sheet.getRow(i);
            String oldNumCell = row.getCell(0)+"";
            String newNumCell = row.getCell(1)+"";
            map.put(oldNumCell,newNumCell);
        }

        return (HashMap) map;
    }


    /***********************************************************************
        * 把 WhereUsed_list 的item 加入表單的Affected item
        * *********************************************************************/
    private static void AddToChange(IChange change, ArrayList WhereUsed_list,int num,int bsize) throws APIException {

        ITable affecteditem_tb = change.getTable(ChangeConstants.TABLE_AFFECTEDITEMS);
        System.out.println("1");
        for(int i=0;i<bsize;i++){
            IItem item = (IItem) WhereUsed_list.get(num);
            IRow affectedrow = affecteditem_tb.createRow(item);
            num=num+1;
            if(num == WhereUsed_list.size()) break;
        }

        //Iterator it = WhereUsed_list.iterator();
        //while(it.hasNext()) {
        //    IItem item = (IItem)it.next();
        //    IRow affectedrow = affecteditem_tb.createRow(item);
        //}
    }


    /***********************************************************************
        * 開立次表單，和其workflow
        * *********************************************************************/
    private static IChange CreateSubChange(IAgileSession session) throws APIException {
        IChange change = null;

        //Define a variable for the subclass
        IAdmin admin = session.getAdminInstance();
        IAgileClass classSubCO = null;
        String atoNextNumber="";
        //Get the subclass ID
        IAgileClass[] classes =  admin.getAgileClasses(IAdmin.CONCRETE);
        for (int i = 0; i < classes.length; i++) {
            // System.out.println("Class Name: "+classes[i].getName().toString());
            if (classes[i].getName().equals("SubCO")) {
                IAutoNumber an = classes[i].getAutoNumberSources()[0];
                atoNextNumber = an.getNextNumber();
                classSubCO = classes[i];
                break;
            }
        }
        //Create a SubCO object
        System.out.println("subNUM:"+atoNextNumber);

        if (classSubCO != null) change = (IChange) session.createObject(classSubCO, atoNextNumber);

        IWorkflow[] wfs = change.getWorkflows();            // 得到全部workflow的選項

        IWorkflow workflow = null;
        for (int i = 0; i < wfs.length; i++) {
            // System.out.println("workflow: "+wfs[i].toString() );
            if (wfs[i].getName().equals("SubCO"))             // 選要使用的workflow名稱
                workflow = wfs[i];
        }
        change.setWorkflow(workflow);


        return change;
    }


    /***************************************************************************************
        * 改Redline BOM table，加入新料號item，將舊料號的欄位填入新料號的欄位，刪除舊料號 item  => 之後要改成客戶的欄位
        * *************************************************************************************/
    private static void updateRedlineBOM(IChange change, IAgileSession session, HashMap map) throws APIException {

        // get affected table
        ITable Affected_tb =  change.getTable(ChangeConstants.TABLE_AFFECTEDITEMS);
        Iterator it = Affected_tb.iterator();
        while (it.hasNext()){
            // get affected item
            IRow row = (IRow) it.next();
            System.out.println("affected item: "+row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_ITEM_NUMBER).toString());
            System.out.println("Old REV: "+row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_OLD_REV).toString());
            String NewRev = getNewRev(row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_OLD_REV).toString());
            row.setValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV,NewRev);
            System.out.println("New REV: "+row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV).toString());
            String AffectedItemNum = row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_ITEM_NUMBER).toString();
            IItem AffectedItem =(IItem) session.getObject(IItem.OBJECT_TYPE,AffectedItemNum);
            // get redline BOM
            AffectedItem.setRevision(change.getName());
            ITable RedlineBOM_tb = AffectedItem.getTable(ItemConstants.TABLE_REDLINEBOM);
            Iterator it2 = RedlineBOM_tb.iterator();
            // temp list -> because ConcurrentModificationException
            ArrayList temp_list = new ArrayList<>();
            while(it2.hasNext()) {
                IRow row2 = (IRow) it2.next();
                String oldNum =row2.getValue(1011).toString();
                if(map.get(oldNum)==null) continue;
                temp_list.add(row2);
            }
            // System.out.println(temp_list);
            Iterator it3 = temp_list.iterator();
            while (it3.hasNext()){
                IRow row3 = (IRow) it3.next();
                //System.out.println(row3.getValues());
                // get new item
                String oldNum =row3.getValue(1011).toString();
                if(map.get(oldNum)==null) continue;
                IItem newitem = (IItem) session.getObject(IItem.OBJECT_TYPE,map.get(oldNum));
                // add new row, set value
                IRow RedlineRow = RedlineBOM_tb.createRow(newitem);
                RedlineRow.setValue(12508,row3.getValue(12508));
                RedlineRow.setValue(12509,row3.getValue(12509));
            }
            // remove old row
            RedlineBOM_tb.removeAll(temp_list);
            }
        }
    private static String getNewRev(String oldRev){

        String NewRev = "";
        String a = oldRev.substring(0,1);
        int b = Integer.parseInt(oldRev.substring(1,2));
        String c = oldRev.substring(2,3);
        int d = Integer.parseInt(oldRev.substring(3,4));

        if (d==9){
            b=b+1;
            d=0;
        }else {
            d=d+1;
        }

        NewRev = a+b+c+d;
        return NewRev;
    }

    }






