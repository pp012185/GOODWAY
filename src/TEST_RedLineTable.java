import com.agile.api.*;
import com.agile.px.ActionResult;
import com.agile.px.ICustomAction;

import java.util.Iterator;

public class TEST_RedLineTable implements ICustomAction{
    @Override
    public ActionResult doAction(IAgileSession session, INode iNode, IDataObject obj) {

        IChange change = (IChange) obj;
        try {
            ITable Affected_tb =  change.getTable(ChangeConstants.TABLE_AFFECTEDITEMS);
            Iterator it = Affected_tb.iterator();
            if (it.hasNext()){
                IRow row = (IRow) it.next();
                System.out.println("affected item: "+row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_ITEM_NUMBER).toString());
                String AffectedItemNum = row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_ITEM_NUMBER).toString();
                IItem AffectedItem =(IItem) session.getObject(IItem.OBJECT_TYPE,AffectedItemNum);
                IItem item = (IItem) session.getObject(IItem.OBJECT_TYPE,"P00009");
                AffectedItem.setRevision(change.getName());
                ITable RedlineBOM_tb = AffectedItem.getTable(ItemConstants.TABLE_REDLINEBOM);
                // get old row
                Iterator it2 = RedlineBOM_tb.iterator();
                IRow row2=null;
                if(it2.hasNext()) row2 = (IRow) it2.next();
                // add new row, set value
                IRow RedlineRow = RedlineBOM_tb.createRow(item);
                RedlineRow.setValue(12508,row2.getValue(12508));
                RedlineRow.setValue(12509,row2.getValue(12509));
                // remove old row
                RedlineBOM_tb.removeRow(row2);

            }

        } catch (APIException e) {
            e.printStackTrace();
        }

        return new ActionResult(0,"⚠⚠⚠ TEST RedLine Table success!");
    }
}
