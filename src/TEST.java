import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class TEST {


    public static void main(String[] args) {
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

    }


}
