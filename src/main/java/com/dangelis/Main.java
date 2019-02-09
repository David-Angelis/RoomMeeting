package com.dangelis;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class Main {
    public static void main(String[] args){
    	System.setProperty("javax.net.ssl.trustStore","//Library/Java/JavaVirtualMachines/jdk1.8.0_05.jdk/Contents/Home/jre/lib/security\n" + 
    			"/cacerts"); 
        SpringApplication.run(Main.class,args);

    }
}
