import java.util.*;

import com.agile.api.*;
import com.anselm.plm.utilobj.Ini;
import com.anselm.plm.utilobj.LogIt;
import org.w3.x2000.x09.xmldsig.ObjectType;

import javax.swing.plaf.synth.SynthOptionPaneUI;

public class TEST {

    static LogIt log = new LogIt();

    private static String getNewRev(String oldRev) {

        String NewRev = "";
        String a = oldRev.substring(0, 1);
        int b = Integer.parseInt(oldRev.substring(1, 2));
        String c = oldRev.substring(2, 3);
        int d = Integer.parseInt(oldRev.substring(3, 4));

        if (d == 9) {
            b = b + 1;
            d = 0;
        } else {
            d = d + 1;
        }

        NewRev = a + b + c + d;
        return NewRev;
    }

    static IAgileSession connect() {
        IAgileSession session = null;
        try {
            System.setProperty("disable.agile.sessionID.generation", "true");
            HashMap params = new HashMap();
            params.put(AgileSessionFactory.USERNAME, "admin");
            params.put(AgileSessionFactory.PASSWORD, "agile933");
            AgileSessionFactory factory;
            factory = AgileSessionFactory.getInstance("http://srv-plm-test.goodway.local:7001/Agile");
            session = factory.createSession(params);
            log.log("成功登入Agile");
        } catch (APIException e) {
            log.log(e);
            log.log("連接Agile時發生錯誤，請檢查config是否輸入錯誤");
            e.printStackTrace();
            System.exit(1);
        }

        return session;
    }


    public static void main(String[] args) {
        IAgileSession session = connect();
        IChange change = null;

          Ini ini = new Ini("C:\\Agile\\Config.ini");
          int NumOfexcelrow = Integer.valueOf(ini.getValue("parameter","NumOfexcelrow")) ;
          String localpath = ini.getValue("path","localpath");
          String fileName = ini.getValue("Name","fileName");

        System.out.println("config row: "+ NumOfexcelrow);


        /*
        IAgileSession session = connect();
        IChange change = null;
        try {
            change = (IChange) session.getObject(IChange.OBJECT_TYPE, "SUB000010");
            IItem item = (IItem) session.getObject(IItem.OBJECT_TYPE,"JHE9020X10001AS-H");
            ITable table = change.getTable(ChangeConstants.TABLE_AFFECTEDITEMS);
            System.out.println(item.getName());
            table.add(item);

            ITable Affected_tb =  change.getTable(ChangeConstants.TABLE_AFFECTEDITEMS);
            Iterator it = Affected_tb.iterator();
            while (it.hasNext()) {
            IRow row = (IRow) it.next();
            System.out.println("Site:"+row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_SITES)+":");
            System.out.println("Site?"+row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_SITES).toString().equals(""));

                if (row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_SITES).toString().equals("")) continue;
                System.out.println("affected item: " + row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_ITEM_NUMBER).toString());
                System.out.println("Old REV: " + row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_OLD_REV).toString());
                String NewRev = getNewRev(row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_OLD_REV).toString());
                row.setValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV, NewRev);
                System.out.println("New REV: " + row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV).toString());
                String AffectedItemNum = row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_ITEM_NUMBER).toString();
                IItem AffectedItem = (IItem) session.getObject(IItem.OBJECT_TYPE, AffectedItemNum);
                // get redline BOM
                AffectedItem.setRevision(change.getName());
                AffectedItem.setManufacturingSite("USA Hub");
                ITable RedlineBOM_tb = AffectedItem.getTable(ItemConstants.TABLE_REDLINEBOM);
                Iterator it2 = RedlineBOM_tb.iterator();

            }


        } catch (APIException e) {
            e.printStackTrace();
        }


        //Ini ini = new Ini("D:\\pp\\config.ini");
        //String string = ini.getValue("agile","a");
        //System.out.println(string);
/*
        Ini ini2 = new Ini("C:\\Users\\pp\\IdeaProjects\\UpdateBOM\\config.ini");
        int NumOfexcelrow = Integer.valueOf(ini2.getValue("parameter","NumOfexcelrow")) ;
        int NumOfBatch = Integer.valueOf(ini2.getValue("parameter","NumOfBatch"));
        String localpath = ini2.getValue("path","localpath");
        String fileName = ini2.getValue("Name","fileName");
        System.out.println(localpath+"\\"+fileName);
        System.out.println("C:\\ExcelFile\\OldAndNew.xlsx");
        System.out.println(NumOfexcelrow);
        System.out.println(NumOfBatch);


        /*
        String a = "V1.0";
        String b = "A1.5";
        String c = "V1.9";
        String d = "A2.9";

        System.out.println(a+" -> "+ getNewRev(a));
        System.out.println(b+" -> "+ getNewRev(b));
        System.out.println(c+" -> "+ getNewRev(c));
        System.out.println(d+" -> "+ getNewRev(d));


        System.out.println(5/2);
        System.out.println(4199/200);

        ArrayList lst = new ArrayList();
        lst.add("A");
        lst.add("B");
        lst.add("C");
        System.out.println(lst.get(0));
        System.out.println(lst.size()==3);


        Map map = new HashMap();
        map.put("1","A");
        map.put("2","B");
        map.put("3","C");
        System.out.println(map);

        System.out.println(map.get("3"));
        System.out.println(map.get("4"));
        System.out.println(map.get("4")==null);

        // ConcurrentModificationException
        List<String> list = new ArrayList<String>();
        list.add("A");
        list.add("B");
        int i =0;
        Iterator it = list.iterator();
        while (it.hasNext()){
            String s = (String) it.next();
            list.add("C");
            i=i+1;
            System.out.println("sdsds");
            System.out.println("i= "+i);
            System.out.println("qwewqeqwe");
        }
        */


    }
}

