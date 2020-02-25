package com.execl.help;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 测试main方法
 * @liangjisnhan
 */
public class ExeclHelp {

	private static final String pathName = "C:/Users/KIM/Desktop/work_bos/666.xlsx";
	
	public static void main(String[] args) throws IOException {

		File file = new File(pathName);
		InputStream inputStream = new FileInputStream(file);
		List<ArrayList<String>> list = HostInsertDB.readXlsx(inputStream, file);
		for (int i = 0; i < list.size(); i++) {
			//\u00A0 标识空格   Non-Breaking SPace
			String str = list.get(i).get(0).trim().replace("\u00A0", "");
			Pattern p = Pattern.compile("\\s*|\t|\r|\n");
            Matcher m = p.matcher(str);
            m.replaceAll("");
			System.out.println("'" +str+ "',");
		}
		inputStream.close();

	}
}
