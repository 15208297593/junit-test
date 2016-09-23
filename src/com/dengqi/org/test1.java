package com.dengqi.org;

import java.util.ArrayList;
import java.util.List;

import org.junit.Test;

public class test1 {
		public static void main(String[] args) {
			double a = 0.05*10;
			double c = 0.02;
			int b = (int) (a/c-10);
			System.out.println(a/c-10);
			System.out.println(b);
		}
	@Test
	public void test(){
		String str = "";
		String[] strArray = str.split(",");
		List<String> lstr1 = new ArrayList<String>();
		List<String> lstr2 = new ArrayList<String>();
		System.out.println(strArray.length);
		for(int i = 0 ;i<strArray.length;i++){
			if(i%2 == 0){
				continue;
			}else{
			lstr1.add(strArray[i]);
			}
		}
		for(String str3:lstr1){
			String new1 = str3.replace("\"value\":\"", "").replace("\"}", "").replace("]", "");
			lstr2.add(new1);
		}
		System.out.println(lstr2);
		
	}
}
