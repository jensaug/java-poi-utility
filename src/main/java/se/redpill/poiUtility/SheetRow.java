package se.redpill.poiUtility;

import java.util.Iterator;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;

/**
 *
 * @author andreas.arvidsson@redpill-linpro.com
 */
public class SheetRow implements Iterable<String> {

    private final Row row;
    private final int size;

    private final static DataFormatter FMT = new DataFormatter();

    public SheetRow(Row row, int size) {
        this.row = row;
        this.size = size;
    }

    public String getCell(int index) {
        return FMT.formatCellValue(row.getCell(index));
    }

    public int getCellAsInt(int index) {
        return (int) row.getCell(index).getNumericCellValue();
    }

    public double getCellAsDouble(int index) {
        return row.getCell(index).getNumericCellValue();
    }

    public boolean getCellAsBool(int index) {
        return row.getCell(index).getBooleanCellValue();
    }

    @Override
    public Iterator<String> iterator() {
        return new MyIterator();
    }

    class MyIterator implements Iterator<String> {

        private int index = 0;

        @Override
        public boolean hasNext() {
            return index < size;
        }

        @Override
        public String next() {
            return getCell(index++);
        }

        @Override
        public void remove() {
            throw new UnsupportedOperationException("not supported yet");
        }
    }

}
