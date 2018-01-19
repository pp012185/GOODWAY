import java.util.*;
import com.anselm.plm.utilobj.Ini;

import javax.swing.plaf.synth.SynthOptionPaneUI;

public class TEST {


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

    public static void main(String[] args) {

        //Ini ini = new Ini("D:\\pp\\config.ini");
        //String string = ini.getValue("agile","a");
        //System.out.println(string);

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
