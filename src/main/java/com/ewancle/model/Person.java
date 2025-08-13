package com.ewancle.model;

import java.time.LocalDateTime;

public record Person(Long id, String name, String email, int age, LocalDateTime createdAt) {}
