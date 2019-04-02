package com.canjun.onenine;

import java.io.File;

public class ChangeName {

	private static String path = "E:\\javaee\\day24-基础加强\\day24-基础加强\\视频";

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		File dir = new File(path);
		File[] files = dir.listFiles();
		for (File file : files) {
			String name = file.getName();
			System.out.println(name);
		   String newName="24_"+name;
			file.renameTo(new File(path+"/"+newName));
			System.out.println(">>newName>>>>"+newName);
		}

	}

}
