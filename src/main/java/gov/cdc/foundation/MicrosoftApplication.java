package gov.cdc.foundation;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.security.oauth2.config.annotation.web.configuration.EnableResourceServer;

@SpringBootApplication
@EnableResourceServer
public class MicrosoftApplication {

	public static void main(String[] args) {
		SpringApplication.run(MicrosoftApplication.class, args);
	}
}
