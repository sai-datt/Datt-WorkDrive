package com.datt.spring.core.service.impl;

import com.datt.spring.core.service.IMessageService;

public class EmailService implements IMessageService {

	@Override
	public String sendMessage(String message, String to) {
		return "Email sent to " + to + "with message:" + message;
	}

}
