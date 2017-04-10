import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.text.SimpleDateFormat;
import java.util.Date;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

public class excel2js {

	private static String Path = "C:\\Users\\Administrator\\Desktop\\data.xls";
	/**
	 * @param args
	 */
	public static void main(String[] args) {
		String rootDir = Path.substring(0,Path.lastIndexOf("\\"));
		String newDir = rootDir + "\\javascripts";
		String fireName = Path.substring(Path.lastIndexOf("\\") + 1, Path.lastIndexOf("."));
		jxl.Workbook readwb = null;
		try {
			// 构建Workbook对象, 只读Workbook对象
			// 直接从本地文件创建Workbook
			File file =new File(newDir);
			if (!file.exists()) {
				file.mkdir();
			}
			
			InputStream instream = new FileInputStream(Path);
			readwb = Workbook.getWorkbook(instream);
			// Sheet的下标是从0开始
			// 获取第一张Sheet表
			for (int k = 0; k < readwb.getNumberOfSheets(); k++) {
				Sheet readsheet = readwb.getSheet(k);
				int rsColumns = readsheet.getColumns();
				int rsRows = readsheet.getRows();
				String outPath = newDir + "\\" + fireName+ "_" + readsheet.getName() + ".js";
				file = new File(outPath);
				try {
					FileOutputStream fos = new FileOutputStream(file);
					fos.write(convert2Byte("/////////////////////////////////////////////////////////////////////////////\r\n"));
					fos.write(convert2Byte("// Introduce:本类来自excel数据的转换，由Haxido、rz20编写的Java工具自动完成。\r\n"));
					fos.write(convert2Byte("// Copyright: 杭州歆享科技有限公司版本所有\r\n"));
					fos.write(convert2Byte("// Author: Haxido(韩相东)、rz20(蒋卓航)\r\n"));
					fos.write(convert2Byte("// Version: 1.0.0 " + getTime() + "\r\n"));
					fos.write(convert2Byte("/////////////////////////////////////////////////////////////////////////////\r\n\r\n"));
					fos.write(convert2Byte("if(sgdb == null){var sgdb = {};}\r\n\r\n"));
					
					fos.write(convert2Byte("/////////////////////////////////////////////////////////////////////////////\r\n"));
					fos.write(convert2Byte("// " + readsheet.getName() + " Helper\r\n"));
					int startJ = readsheet.getCell(0, 1).getContents().equals("id") ? 1 : 0;//自动将ID做为主键来处理
					for (int i = startJ; i < rsColumns; i++) {
						String tag1 = readsheet.getCell(i, 0).getContents();
						String tag2 = readsheet.getCell(i, 1).getContents();
						fos.write(convert2Byte("// " + tag2 + ":" + tag1 + "\r\n"));
					}
					fos.write(convert2Byte("/////////////////////////////////////////////////////////////////////////////\r\n\n"));
					fos.write(convert2Byte("if(sgdb." + readsheet.getName() + " == null){ sgdb." + readsheet.getName() + " = {\r\n\n"));					
					for (int i = 2; i < rsRows; i++) {
						writeTab(fos);
						if (startJ > 0) {
							fos.write(convert2Byte(readsheet.getCell(0, i).getContents() + " : {"));
						} else {
							fos.write(convert2Byte((i - 1) + " : {"));
						}
						
						for (int j = startJ; j < rsColumns; j++) {
							Cell cell = readsheet.getCell(j, i);
							String tag = readsheet.getCell(j, 1).getContents();
							String data = cell.getContents();
							if (data == null) {
								data = "";
							}
							fos.write(convert2Byte(tag  + " : "));
							try {
								Double.valueOf(data);
								fos.write(convert2Byte(data));
							} catch (Exception e) {
								if (data.equals("true") || data.equals("false")) {
									fos.write(convert2Byte(data));
								} else {
									fos.write(convert2Byte("\"" + data + "\""));
								}
							}
							if (j < rsColumns - 1) {
								fos.write(convert2Byte(","));
							}
						}
						fos.write(convert2Byte("},\r\n"));
					}
					fos.write(convert2Byte("  };\r\n}"));
					fos.close();
					System.out.println("成功输出文件：" + outPath);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
			
			// // 利用已经创建的Excel工作薄,创建新的可写入的Excel工作薄
			// jxl.write.WritableWorkbook wwb = Workbook.createWorkbook(new
			// File(
			// "F:/红楼人物1.xls"), readwb);
			// // 读取第一张工作表
			// jxl.write.WritableSheet ws = wwb.getSheet(0);
			// // 获得第一个单元格对象
			// jxl.write.WritableCell wc = ws.getWritableCell(0, 0);
			// // 判断单元格的类型, 做出相应的转化
			// if (wc.getType() == CellType.LABEL)
			// {
			// Label l = (Label) wc;
			// l.setString("新姓名");
			// }
			// // 写入Excel对象
			// wwb.write();
			// wwb.close();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			readwb.close();
		}
	}
	
	private static byte[] convert2Byte(String str)
			throws UnsupportedEncodingException {
		return str.getBytes("utf-8");
	}
	private static void writeTab(FileOutputStream fos) throws IOException {
		fos.write("\t".getBytes());
	}
	private static String getTime() {
		Date date = new Date(System.currentTimeMillis());
		SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		return df.format(date);
	}
}
