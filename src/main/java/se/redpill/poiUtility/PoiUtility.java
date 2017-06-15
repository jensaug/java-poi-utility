package se.redpill.poiUtility;

import java.io.BufferedReader;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.PrintStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XLSX2CSV;
import org.xml.sax.SAXException;

/**
 * Utilities for the stupid Apache POI API
 * @author jens.augustsson@redpill-linpro.com
 *
 */
public class PoiUtility {

	private BufferedReader bufferedReader;
	private List<String> headers = new ArrayList<String>();
	private Character separator;
	private String[] lineSplits;
	
	/**
	 * If this constructor is used, non-static (stateful) methods can be used
	 * @param inputStream
	 * @param sheetName
	 * @param minColumns
	 * @param separator - defaults to tab
	 * @param firstLineIsHeader
	 */
	public PoiUtility(InputStream inputStream, String sheetName, int minColumns, Character separator, boolean firstLineIsHeader) {
		try {
			this.separator = separator != null ? separator : '\t';
			this.bufferedReader  = PoiUtility.getCsvBufferedReader(inputStream, sheetName, minColumns, this.separator);
			if (firstLineIsHeader) {
				String[] splits = bufferedReader.readLine().split(this.separator.toString());
				for (int i = 0; i < splits.length; i++) {
					splits[i] = getString(splits[i]);
				}
				this.headers = Arrays.asList(splits);
			}
		} catch (IOException e) {
			throw new RuntimeException("Failed to construct PoiUtility for sheet " + sheetName, e);
		}
	}

//Public nonStatic
	
	public boolean readLine() {
		try {
			String line = this.bufferedReader.readLine();
			if (line != null) {
				this.lineSplits = line.split(separator.toString());
				return !line.trim().equals("");
			} else {
				this.bufferedReader.close();
				return false;
			}
		} catch (IOException e) {
			throw new RuntimeException("Could not readLine", e);
		}
	}

	public String getCell(int columnIndex) {
		return getString(this.lineSplits[columnIndex]);
	}
	
	public String getCellAsString(String header) {
		int idx = getColumnIndex(header);
		return getString(this.lineSplits[idx]);
	}

	public Integer getCellAsInteger(String header) {
		int idx = getColumnIndex(header);
		return getInteger(this.lineSplits[idx]);
	}	
	
	public Long getCellAsLong(String header) {
		int idx = getColumnIndex(header);
		return getLong(this.lineSplits[idx]);
	}

	public Date getCellAsDate(String header, SimpleDateFormat sdf, boolean nullifyNonParseable) {
		Date date = null;
		int idx = -1;
		try {
			idx = getColumnIndex(header);
			String split = this.lineSplits[idx];
			if (split != null && split.trim().length() > 0) {
				if(split.startsWith("\""))
					split = (String) split.subSequence(1,(split.length() -1));
				if (sdf != null) {
					date = sdf.parse(split);
				} else {
					date = new SimpleDateFormat().parse(split);
				}								
			}
		} catch (ParseException e) {
			if (!nullifyNonParseable) {
				throw new RuntimeException("Could not parse date from data " + this.lineSplits[idx]);				
			}
		}
		return date;
	}
	
	public Boolean getCellAsBoolean(String header, String trueValue, String falseValue) {
		int idx = getColumnIndex(header);
		return getBoolean(this.lineSplits[idx], trueValue, falseValue);		
	}
	
// Public static	
	
	public static String cleanCases(String data) {
		if (data != null && data.length() > 1) {
			//Data exists, remove double quotes
			return data.substring(0, 1).toUpperCase() + data.substring(1).toLowerCase();
		}
		return data;
	}

	public static String getString(String data) {
		if (data != null && data.startsWith("\"")) {
			//Data exists, remove double quotes
			data = data.substring(1, data.length()-1);
		}
		if (data == null || data.trim().length() == 0) {
			return null;
		}
		return data;
	}

	public static Integer getInteger(String data) {
		if (data != null && data.length() > 1) {
			try {
				//Should be an Integer
				return Integer.parseInt(data);		
			} catch (NumberFormatException e) {
				//Can be a Float
				return ((Float) Float.parseFloat(data)).intValue();			
			}
		}
		throw new RuntimeException("Data cannot be parsed to Integer: " + data);
	}

	public static Integer getIntegerFromInt(String data) {
			//Should be an int
		try {
			return Integer.valueOf(data);
		} catch (NumberFormatException e) {
			e.printStackTrace();
			throw new RuntimeException("Data cannot be parsed to Integer: " + data);
		}
	}

	public static Long getLong(String data) {
		if (data != null && data.length() > 1) {
			try {
				//Should be an Integer
				return Long.parseLong(data);		
			} catch (NumberFormatException e) {
				//Can be a Float
				return ((Float) Float.parseFloat(data)).longValue();			
			}
		}
		throw new RuntimeException("Data cannot be parsed to Long: " + data);
	}

	/**
	 * 
	 * @param data
	 * @param trueValue - this string or 1 or true will return TRUE
	 * @param falseValue - this string or 0 or false will return FALSE
	 * @return
	 */
	public static Boolean getBoolean(String data, String trueValue, String falseValue) {
		Boolean b = null;
		data = getString(data);
		if (data != null) {
			if ((trueValue != null && trueValue.equals(data)) || "1".equals(data) || "true".equals(data.toLowerCase())) {
				b = true;
			}
			if (b == null && ((falseValue != null && falseValue.equals(data)) || "0".equals(data) || "false".equals(data.toLowerCase())) ) {
				b = false;
			}
		}			
		if (b != null) {
			return b;
		} else {
			throw new RuntimeException("Data cannot be parsed to Boolean: " + data);			
		}
	}
	
	public static Date getDateOrNull(String data, SimpleDateFormat sdf) {
		Date date = null;
		try {
			date = sdf.parse(data);
		} catch (ParseException e) {
			;//Ignore badly formatted date
		}
		return date;
	}
	/**
	 * Utility for Apache POI XLSX2CSV, enabling streaming rows from a specific sheet to a line-based BufferedReader
	 * @param inputStream
	 * @param sheetName
	 * @param minColumns
	 * @param separator
	 * @return
	 */
	public static BufferedReader getCsvBufferedReader(InputStream inputStream, String sheetName, int minColumns, char separator) {
		BufferedReader br = null; 
		try {
			OPCPackage pkg = OPCPackage.open(inputStream);
			ByteArrayOutputStream baos = new ByteArrayOutputStream();
			PrintStream ps = new PrintStream(baos);
			XLSX2CSV xlsx2csv = new XLSX2CSV(pkg, ps, minColumns);
			xlsx2csv.setSeparator(separator);
			xlsx2csv.process();
			
			ByteArrayInputStream bais = new ByteArrayInputStream(baos.toByteArray());
			br = new BufferedReader(new InputStreamReader(bais));
			String line = br.readLine(); //First line is empty	
			boolean foundSheet = false;
			while (line != null && !foundSheet) {
				if (line != null && line.startsWith(sheetName)) {
					foundSheet = true;
				} else {
					line = br.readLine();
				}
			}
			if (!foundSheet) {
				throw new RuntimeException("Could not find sheet " + sheetName + " in uploaded file");
			}
//			bais.close();
//			ps.close();
//			baos.close();
//			pkg.close();			
			
		} catch (InvalidFormatException e) {
			throw new RuntimeException(e);
		} catch (IOException e) {
			throw new RuntimeException(e);
		} catch (OpenXML4JException e) {
			throw new RuntimeException(e);
		} catch (ParserConfigurationException e) {
			throw new RuntimeException(e);
		} catch (SAXException e) {
			throw new RuntimeException(e);
		}		
		return br;
	}

//Privates
	
	private int getColumnIndex(String header) {
		int idx = this.headers.indexOf(header);		
		if (idx > -1) {
			return idx;
		} else {
			throw new RuntimeException("Could not find header " + header);
		}
	}
	
}
