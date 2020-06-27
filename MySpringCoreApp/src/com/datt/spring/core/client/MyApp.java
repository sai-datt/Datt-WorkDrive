package com.datt.spring.core.client;

import org.springframework.context.annotation.AnnotationConfigApplicationContext;

import com.datt.spring.core.config.ContextConfig;

public class MyApp {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		// ask bean factory to make us available the service

		// Initiate the bean factory
		AnnotationConfigApplicationContext context = 
				new AnnotationConfigApplicationContext(ContextConfig.class);
	}

}
