package com.ckl.tool;

import java.io.File;
import java.util.List;

public class ToolMain {

	public static void main(String[] args) throws Exception {
		
		String srcPath = "f:/ckl/ckd/������ѧ-ְҵ��ר-��ת��ѧ¥׮��ʩ����¼���ܱ�.xlsx";
		String dstPath = "f:/ckl/ckd/��׹�ע׮��׮ʩ����¼1.xls";
		
		List<String[]> list = POIUtil.readExcel(new File(srcPath));
		if (list == null || list.isEmpty()) {
			System.err.println("û�ж�ȡ������");
			return;
		}
		
		 
	}

}
