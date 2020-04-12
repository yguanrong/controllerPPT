package com.znv.utils;

import com.znv.App;
import org.java_websocket.WebSocket;
import org.java_websocket.client.WebSocketClient;
import org.java_websocket.handshake.ServerHandshake;

import javax.swing.*;
import java.io.*;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.Properties;

/**
 * Hello world!
 *
 */
public class WebSocketClientTest {

	private static WebSocketClient client;
	public static void subscribeFss( final JacobPptUtil jacobPptUtil ) {


		//读取 jar包外的配置文件
		String filePath = System.getProperty("user.dir")
				+ "/config/service.properties";
		InputStream inForBiuld = null;
		try {
			inForBiuld = new BufferedInputStream(new FileInputStream(filePath));
		} catch (FileNotFoundException e) {
			App.Waraning(e.getMessage());
			e.printStackTrace();
		}

		//读取 jar包内的配置文件
		// 将配置文件加载到流中
		InputStream inForDev = WebSocketClientTest.class.getClassLoader()
				.getResourceAsStream("service.properties");
		// 创建并加载配置文件
		Properties pro = new Properties();
		try {
			// 加载配置文件流
			pro.load(inForDev);
		} catch (IOException e) {
			App.Waraning(e.getMessage());
			e.printStackTrace();
		}
		// 获取配置文件定义的值
		String url = pro.getProperty("url");

		try {
			client = new WebSocketClient(new URI(url)) {

				@Override
				public void onOpen(ServerHandshake handshakedata) {
					// TODO Auto-generated method stub
					System.out.println("open socket");
					// 订阅32011520006
					// client.send("2&11000020007");
					client.send("2&32011520006");
				}

				@Override
				public void onMessage(String message) {
					// TODO Auto-generated method stub
					System.out.println(message);
					if ("下一页".equals(message)){
						jacobPptUtil.nextPage();
					}
					if ("上一页".equals(message)){
						jacobPptUtil.upPage();
					}
					if ("结束放映".equals(message)){
						jacobPptUtil.stopPpt();
					}
					if ("开始放映".equals(message)){
						try {
							jacobPptUtil.runPpt();
						} catch (Exception e) {
							e.printStackTrace();
						}
					}
					if ("关闭软件".equals(message)){
						try {
							jacobPptUtil.exitPPT();
						} catch (Exception e) {
							e.printStackTrace();
						}
					}
				}

				@Override
				public void onError(Exception ex) {
					// TODO Auto-generated method stub
					System.err.println(ex);
				}

				@Override
				public void onClose(int code, String reason, boolean remote) {
					// TODO Auto-generated method stub
					System.out.println(code+"---close socket---"+reason);
				}

			};
		} catch (URISyntaxException e) {
			App.Waraning(e.getMessage());
//			Logger.L.error(e);
		}

		client.connect();
	}

	public static void sendMessage(String message){
		if (client != null){
			client.send(message);
		}
	}

}
