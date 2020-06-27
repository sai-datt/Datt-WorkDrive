package com.datt.spring.core.config;

import org.springframework.context.annotation.ComponentScan;
import org.springframework.context.annotation.Configuration;

//Need to indicate spring to use this class as config class
@Configuration
//Need to specify location of resources
@ComponentScan("com.datt.spring.core.service")
public class ContextConfig {

}
