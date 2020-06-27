package com.datt.spring.core.service.impl;

import com.datt.spring.core.service.IMessageService;

public class SmsService implements IMessageService {

	@Override
	public String sendMessage(String message, String to) {
		return "Sms sent to " + to + "with message:" + message;
	}

}
