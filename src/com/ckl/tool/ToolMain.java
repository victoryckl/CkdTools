package com.ckl.tool;

import java.io.File;
import java.util.List;

public class ToolMain {

	public static void main(String[] args) throws Exception {
		
		String srcPath = "f:/ckl/ckd/第四中学-职业中专-中转教学楼桩基施工记录汇总表.xlsx";
		String dstPath = "f:/ckl/ckd/冲孔灌注桩单桩施工记录1.xls";
		
		List<String[]> list = POIUtil.readExcel(new File(srcPath));
		if (list == null || list.isEmpty()) {
			System.err.println("没有读取到数据");
			return;
		}
		
		 
	}

}
