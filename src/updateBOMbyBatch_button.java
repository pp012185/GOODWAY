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

    Ini ini = new Ini("C:\\Agile\\Config.ini");
    int NumOfexcelrow = Integer.valueOf(ini.getValue("parameter","NumOfexcelrow")) ;
    int NumOfBatch = Integer.valueOf(ini.getValue("parameter","NumOfBatch"));
    String localpath = ini.getValue("path","localpath");
    String fileName = ini.getValue("Name","fileName");
    @Override
    public ActionResult doAction(IAgileSession session, INode Node, IDataObject obj) {


        System.out.println("Node AIPNAME: "+Node.getAPIName());
//        try {
//            System.out.println("------Start------");
//            IChange change = (IChange) obj;
//            String filepath = localpath+"\\"+change.getName().toString()+"_"+fileName;
//
//            // create mapping
//            HashMap map = getMapping(filepath);
//            System.out.println("Map: "+ map);
//
//            // update
//            ITable reltionship_tb = obj.getTable(ChangeConstants.TABLE_RELATIONSHIPS);
//            Iterator it = reltionship_tb.iterator();
//            while (it.hasNext()){
//                System.out.println("找relation table");
//                IRow row = (IRow) it.next();
//                IDataObject obj2 = row.getReferent();
//                // 應該不用superclass? 判斷是否為BOM替換單的APIName?
//                //System.out.println("判斷: "+obj2.getAgileClass().getSuperClass().getAPIName().toLowerCase());
//                if (obj2.getAgileClass().getSuperClass().getAPIName().toLowerCase().equals("changeordersclass")) {
//                    System.out.println("是change order");
//                    IChange change2 = (IChange) obj2;
//                    if(change2.getStatus().equals("申請人")) {
//                        System.out.println("是申請人");
//                        //改affected item 的redline table
//                        updateRedlineBOM(change2,session,map);
//                        //次表單進下一站
//                        changeStatus(change2);
//                    }
//                }
//            }
//
//            // update flag
//            ITable relation_tb = change.getTable(ChangeConstants.TABLE_RELATIONSHIPS);
//            Iterator it_relation = relation_tb.iterator();
//            Boolean a=false;
//            while (it_relation.hasNext()){
//                IRow r =(IRow) it_relation.next();
//                IChange c =(IChange) session.getObject(IChange.OBJECT_TYPE,r.getValue(ChangeConstants.ATT_RELATIONSHIPS_NAME).toString());
//                //System.out.println("SUB: "+ c.getStatus());
//                if ( c.getStatus().toString().equals("申請人")){
//                    a=false;
//                    break;
//                }else if( c.getStatus().toString().equals("發行")){
//                    a=true;
//                }
//            }
//            if(a){
//                // Get the Part Category cell
//                ICell cell = change.getCell(ChangeConstants.ATT_PAGE_THREE_LIST02);
//                // Get available list values for Part Category
//                IAgileList values = cell.getAvailableValues();
//                // Set the value to Electrical
//                values.setSelection(new Object[] { "Yes" });
//                cell.setValue(values);
//            }
//
//
//            System.out.println("--------End--------");
//        } catch (IOException e) {
//            e.printStackTrace();
//        } catch (APIException e) {
//            e.printStackTrace();
//        }


        return new ActionResult(0,"update BOM button success!");
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
            //System.out.println("Site"+row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_SITES));
            if(row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_SITES).toString().equals("")) continue;
            String tempsite =row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_SITES).toString();

            String AffectedItemNum = row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_ITEM_NUMBER).toString();
            IItem AffectedItem =(IItem) session.getObject(IItem.OBJECT_TYPE,AffectedItemNum);
            //System.out.println("掃完一次BOM");
            // get redline BOM
            AffectedItem.setRevision(change.getName());
            AffectedItem.setManufacturingSite(tempsite);
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
            //System.out.println("新增BOM並改欄位");
            // System.out.println(temp_list);
            HashMap<String,String[]> tempmap = new HashMap<>();
            Iterator it3 = temp_list.iterator();
            while (it3.hasNext()){
                String[] array = new String[2];
                IRow row3 = (IRow) it3.next();
                //System.out.println(row3.getValues());
                // get new item
                String oldNum =row3.getValue(1011).toString();
                if(map.get(oldNum)==null) continue;
                IItem newitem = (IItem) session.getObject(IItem.OBJECT_TYPE,map.get(oldNum));
                // add new row, set value
                IRow RedlineRow = RedlineBOM_tb.createRow(newitem);
                HashMap map2 = new HashMap();

                //RedlineRow.setValue(1012,row3.getValue(1012));      // 序號
                array[0]=row3.getValue(1012).toString();
                map2.put(1637,row3.getValue(1637));      // 製造單位
                map2.put(12508,row3.getValue(12508));    // 底數
                map2.put(1035,row3.getValue(1035));      // 組成用量
                map2.put(2177,row3.getValue(2177));      // 計算值
                //RedlineRow.setValue(1019,row3.getValue(1019));      // 插件位置
                if(!row3.getValue(1019).toString().isEmpty())array[1]=row3.getValue(1019).toString();
                map2.put(2175,row3.getValue(2175));      // 優先順序
                map2.put(2176,row3.getValue(2176));      // 代替群組
                map2.put(1638,row3.getValue(1638));      // 標準成本計算
                tempmap.put(newitem.getName(),array);
                //RedlineRow.setValue(12508,row3.getValue(12508));
                //RedlineRow.setValue(12509,row3.getValue(12509));
                RedlineRow.setValues(map2);
            }
            //System.out.println("刪除BOM");
            // remove old row
            RedlineBOM_tb.removeAll(temp_list);

            //System.out.println("加序號和插件位置");
            ITable RedlineBOM_tb2 = AffectedItem.getTable(ItemConstants.TABLE_REDLINEBOM);
            Iterator it4 = RedlineBOM_tb2.iterator();
            while (it4.hasNext()){
                IRow row4 = (IRow) it4.next();
                String[] a = tempmap.get(row4.getValue(1011));
                if(a != null)
                {
                    row4.setValue(1012,a[0]);      // 序號
                    row4.setValue(1019,a[1]);
                }
            }

            //System.out.println("改版本");
            //System.out.println("affected item: "+row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_ITEM_NUMBER).toString());
            //System.out.println("Old REV: "+row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_OLD_REV).toString());
            String NewRev = getNewRev(row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_OLD_REV).toString());
            if(NewRev.equals("")) ;
            else row.setValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV,NewRev);
            //System.out.println("New REV: "+row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV).toString());
        }
        System.out.println("Complete BOM替換單: "+ change.getName().toString());
    }
    private static String getNewRev(String oldRev){

        String NewRev = "";
        if (oldRev.equals("")){
            NewRev = "";
        }else {
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
        }

        return NewRev;
    }

    /***************************************************************************************
     * 次表單進入下一個站別
     * *************************************************************************************/
    private static void changeStatus(IChange change) throws APIException {
        // Get the next status of the change
        IStatus nextStatus = change.getDefaultNextStatus();
        // Get default approvers for the next status
        ISignoffReviewer[] defaultApprovers =change.getReviewers(nextStatus, WorkflowConstants.USER_APPROVER);
        List<ISignoffReviewer> approverList =Arrays.asList(defaultApprovers);
        // Get default observers for the next status
        ISignoffReviewer[] defaultObservers =change.getReviewers(nextStatus, WorkflowConstants.USER_OBSERVER);
        List<ISignoffReviewer> observerList =Arrays.asList(defaultObservers);
        // Change to the next status
        change.changeStatus(nextStatus, false, "", false, false, null, approverList,observerList, null, false);
    }




}






