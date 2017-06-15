package se.redpill.poiUtility;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 *
 * @author andreas.arvidsson@redpill-linpro.com
 */
public class SheetPage implements Iterable<SheetRow> {

    private final Sheet sheet;

    private final Map<String, Integer> colNames = new HashMap<String, Integer>();

    private final int rowCount, colCount;

    public SheetPage(Sheet sheet) {
        this.sheet = sheet;
        rowCount = sheet.getPhysicalNumberOfRows();
        colCount = calcColumnCount();
        calcColumnNames();
    }

    public int getRowCount() {
        return rowCount;
    }

    public int getColCount() {
        return colCount;
    }

    public SheetRow getRow(int index) {
        return new SheetRow(sheet.getRow(index), colCount);
    }

    public int getColIndex(String colName) {
        return colNames.get(colName);
    }

    private int calcColumnCount() {
        int count = 0;
        for (int i = 0; i < 10 || i < rowCount; i++) {
            if (sheet.getRow(i).getPhysicalNumberOfCells() > count) {
                count = sheet.getRow(i).getPhysicalNumberOfCells();
            }
        }
        return count;
    }

    private void calcColumnNames() {
        Row row = sheet.getRow(0);
        Cell cell;
        for (int i = 0; i < row.getPhysicalNumberOfCells(); i++) {
            cell = row.getCell(i);
            if (cell != null) {
                colNames.put(cell.toString(), i);
            }
        }
    }

    @Override
    public Iterator<SheetRow> iterator() {
        return new MyIterator();
    }

    class MyIterator implements Iterator<SheetRow> {

        private int index = 0;

        @Override
        public boolean hasNext() {
            return index < rowCount;
        }

        @Override
        public SheetRow next() {
            return getRow(index++);
        }

        @Override
        public void remove() {
            throw new UnsupportedOperationException("not supported yet");
        }
    }

}
