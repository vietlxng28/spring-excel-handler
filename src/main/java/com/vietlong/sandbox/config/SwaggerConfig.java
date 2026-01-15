package com.vietlong.sandbox.config;

import io.swagger.v3.oas.models.OpenAPI;
import io.swagger.v3.oas.models.info.Contact;
import io.swagger.v3.oas.models.info.Info;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

@Configuration
public class SwaggerConfig {

  @Bean
  public OpenAPI customOpenAPI() {
    return new OpenAPI()
        .info(new Info()
            .title("Sandbox API")
            .version("1.0")
            .description("API Sandbox")
            .contact(new Contact()
                .name("VietLong")
                .email("vtlong@example.com")));
  }

}
