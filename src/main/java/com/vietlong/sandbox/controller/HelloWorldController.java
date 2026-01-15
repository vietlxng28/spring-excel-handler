package com.vietlong.sandbox.controller;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import io.swagger.v3.oas.annotations.tags.Tag;

@Tag(name = "Hello World", description = "Hello World API")
@RestController
@RequestMapping("/hello-world")
public class HelloWorldController {

  @GetMapping
  public String helloWorld() {
    return "Hello World";
  }

}