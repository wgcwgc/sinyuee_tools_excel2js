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

public class excel2js_ext {

	private static String Path = "C:\\Users\\Administrator\\Desktop\\data.xls";
	/**
	 * @param args
	 */
	public static void main(String[] args) {
		Path = args[0];
		String rootDir = Path.substring(0,Path.lastIndexOf("\\"));
		String newDir = rootDir + "\\javascripts";
		String fireName = Path.substring(Path.lastIndexOf("\\") + 1, Path.lastIndexOf("."));
		jxl.Workbook readwb = null;
		try {
			// ����Workbook����, ֻ��Workbook����
			// ֱ�Ӵӱ����ļ�����Workbook
			File file =new File(newDir);
			if (!file.exists()) {
				file.mkdir();
			}
			
			InputStream instream = new FileInputStream(Path);
			readwb = Workbook.getWorkbook(instream);
			// Sheet���±��Ǵ�0��ʼ
			// ��ȡ��һ��Sheet��
			for (int k = 0; k < readwb.getNumberOfSheets(); k++) {
				Sheet readsheet = readwb.getSheet(k);
				int rsColumns = readsheet.getColumns();
				//ȥ������
				while(readsheet.getCell(rsColumns - 1, 1).getContents().equals("")){
					rsColumns --;
				}
				int rsRows = readsheet.getRows();
				String outPath = newDir + "\\" + fireName+ "_" + readsheet.getName() + ".js";
				file = new File(outPath);
				try {
					FileOutputStream fos = new FileOutputStream(file);
					fos.write(convert2Byte("/////////////////////////////////////////////////////////////////////////////\r\n"));
					fos.write(convert2Byte("// Introduce:��������excel���ݵ�ת������Haxido��rz20��д��Java�����Զ���ɡ�\r\n"));
					fos.write(convert2Byte("// Copyright: �������Ƽ����޹�˾�汾����\r\n"));
					fos.write(convert2Byte("// Author: Haxido(���ණ)��rz20(��׿��)\r\n"));
					fos.write(convert2Byte("// Version: 1.0.0 " + getTime() + "\r\n"));
					fos.write(convert2Byte("/////////////////////////////////////////////////////////////////////////////\r\n\r\n"));
					fos.write(convert2Byte("if(sgdb == null){var sgdb = {};}\r\n\r\n"));
					
					fos.write(convert2Byte("/////////////////////////////////////////////////////////////////////////////\r\n"));
					fos.write(convert2Byte("// " + readsheet.getName() + " Helper\r\n"));
					int startJ = readsheet.getCell(0, 1).getContents().equals("id") ? 1 : 0;//�Զ���ID��Ϊ����������
					for (int i = 0; i < rsColumns; i++) {
						String tag1 = readsheet.getCell(i, 0).getContents();
						String tag2 = readsheet.getCell(i, 1).getContents();
						fos.write(convert2Byte("// " + tag2 + ":" + tag1 + "\r\n"));
					}
					fos.write(convert2Byte("/////////////////////////////////////////////////////////////////////////////\r\n\n"));
					fos.write(convert2Byte("if(sgdb." + readsheet.getName() + " == null){ sgdb." + readsheet.getName() + " = {\r\n"));					
					String id = "";
					int id_number = 0;
					if(readsheet.getCell(1,  1).getContents().contains("id_")){
						for (int i = 2; i < rsRows; i++) {
							//���ֿ�������
							if(readsheet.getCell(0, i).getContents().equals("")){
								continue;
							}
							
							writeTab(fos);
							
							//����Ϊע����������
//							if(readsheet.getCell(1, i).getContents().equals("0")){
//								fos.write(convert2Byte("//" + readsheet.getCell(2, i).getContents() + "\n"));
//								continue;
//							}
							
							if (startJ > 0) {
								if(id.equals(readsheet.getCell(0, i).getContents())){//������id��ͬ
									writeTab(fos);
									fos.write(convert2Byte(readsheet.getCell(1, i).getContents() + ":{"));
								}
								else{//������id��ͬ
									if(readsheet.getCell(1, i).getContents().equals("0")){
										if(id_number > 0){ //���ǵ�һ������ս�
											fos.write(convert2Byte("},\n\t"));
										}
										
										fos.write(convert2Byte("//" + readsheet.getCell(2, i).getContents() + "\n"));
										continue;
									}
									
									id_number++;
									id = readsheet.getCell(0, i).getContents();
									fos.write(convert2Byte(id + ":{\n\t\t" + readsheet.getCell(1, i).getContents() + ":{"));
								}
								
							} else {
								fos.write(convert2Byte((i - 1) + ":{\n\t\t"));
							}
							
							for (int j = startJ + 1; j < rsColumns; j++) {
								
								Cell cell = readsheet.getCell(j, i);
								String tag = readsheet.getCell(j, 1).getContents();
								String data = cell.getContents();
								if (data == null) {
									data = "";
								}
								fos.write(convert2Byte(tag  + ":"));
								if(tag.length() > 1){
									if(tag.charAt(tag.length() - 1) == '_'){
										data = "\"" + data + "\"";
										fos.write(convert2Byte(data));
									}else{
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
									}
								}
								
								
								if (j < rsColumns - 1) {
									fos.write(convert2Byte(", "));
								}
							}
							fos.write(convert2Byte("},\r\n"));
						}
						fos.write(convert2Byte("\t},\n};}"));
						fos.close();
						System.out.println("�ɹ�����ļ���" + outPath);
					}
					else{
						for (int i = 2; i < rsRows; i++) {
							//���ֿ�������
							if(readsheet.getCell(0, i).getContents().equals("")){
								continue;
							}
							
							writeTab(fos);
							if (startJ > 0) {
								fos.write(convert2Byte(readsheet.getCell(0, i).getContents() + ":{"));
							} else {
								fos.write(convert2Byte((i - 1) + ":{"));
							}
							
							for (int j = startJ; j < rsColumns; j++) {
								Cell cell = readsheet.getCell(j, i);
								String tag = readsheet.getCell(j, 1).getContents();
								String data = cell.getContents();
								if (data == null) {
									data = "";
								}
								fos.write(convert2Byte(tag  + ":"));
								if(tag.length() > 1){
									if(tag.charAt(tag.length() - 1) == '_'){
										data = "\"" + data + "\"";
										fos.write(convert2Byte(data));
									}else{
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
									}
								}
								if (j < rsColumns - 1) {
									fos.write(convert2Byte(", "));
								}
							}
							fos.write(convert2Byte("},\r\n"));
						}
						fos.write(convert2Byte("};}"));
						fos.close();
						System.out.println("�ɹ�����ļ���" + outPath);
					}
					
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
			
			
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
