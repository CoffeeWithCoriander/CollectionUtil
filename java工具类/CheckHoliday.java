package com.tci.util;

import java.io.BufferedReader;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.Date;

public class CheckHoliday {
    //0 上班 1周末 2节假日
    public static String checkDay() {
        //判断今天是否是工作日 周末 还是节假日
        SimpleDateFormat f=new SimpleDateFormat("yyyyMMdd");
        String httpArg=f.format(new Date());
        String httpUrl="http://tool.bitefu.net/jiari";
        BufferedReader reader = null;
        String result = null;
        StringBuffer sbf = new StringBuffer();
        httpUrl = httpUrl + "?d=" + httpArg;
        try {
            URL url = new URL(httpUrl);
            HttpURLConnection connection = (HttpURLConnection) url
                    .openConnection();
            connection.setRequestMethod("GET");
            connection.connect();
            InputStream is = connection.getInputStream();
            reader = new BufferedReader(new InputStreamReader(is, "UTF-8"));
            String strRead = null;
            while ((strRead = reader.readLine()) != null) {
                sbf.append(strRead);
                sbf.append("\r\n");
            }
            reader.close();
            result = sbf.toString();
        } catch (Exception e) {
            e.printStackTrace();
        }
        String[] resultArr = result.split("\"");
        if(resultArr !=null && resultArr.length>2){
        	result = resultArr[1];
        }
		return result;
    }
}
