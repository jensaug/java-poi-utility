package se.redpill.poiUtility;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PushbackInputStream;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author andreas.arvidsson@redpill-linpro.com
 */
public class Spreadsheet implements Iterable<SheetPage> {

    private Workbook wb;
    private int size;

    public Spreadsheet(InputStream is) throws IOException, InvalidFormatException {
        InputStream pbIS = new PushbackInputStream(is, 10);
        //HSSF (.xls)
        if (POIFSFileSystem.hasPOIFSHeader(pbIS)) {
            POIFSFileSystem fs = new POIFSFileSystem(is);
            wb = new HSSFWorkbook(fs);
        }
        //XSSF (.xlsx)
        else {
            OPCPackage opc = OPCPackage.open(pbIS);
            wb = new XSSFWorkbook(opc);
        }
        size = wb.getNumberOfSheets();
    }

    public Spreadsheet(File file) throws IOException, InvalidFormatException {
        this(new FileInputStream(file));
    }

    public Spreadsheet(String path) throws IOException, InvalidFormatException {
        this(new File(path));
    }

    public SheetPage getPage(int pageIndex) {
        return new SheetPage(wb.getSheetAt(pageIndex));
    }

    @Override
    public Iterator<SheetPage> iterator() {
        return new MyIterator();
    }

    class MyIterator implements Iterator<SheetPage> {

        private int index = 0;

        @Override
        public boolean hasNext() {
            return index < size;
        }

        @Override
        public SheetPage next() {
            return getPage(index++);
        }

        @Override
        public void remove() {
            throw new UnsupportedOperationException("not supported yet");
        }
    }

}
