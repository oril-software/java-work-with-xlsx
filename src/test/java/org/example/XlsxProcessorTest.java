package org.example;

import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class XlsxProcessorTest {

	@Test
	void testXlsxProcessing() {
		List<User> users = new ArrayList<>();
		users.add(new User("John", "Doe", 25, "john_doe@longemail.com"));
		users.add(new User("Jane", "Doe", 21, "jane@oril.co"));
		users.add(new User("Test", "User", 35, "test@oril.co"));

		XlsxProcessor xlsxProcessor = new XlsxProcessor();
		File file = xlsxProcessor.createXlsxFile(users);

		assertNotNull(file);
		assertTrue(file.isFile());
		assertEquals("users.xlsx", file.getName());

		List<User> parsedUsers = xlsxProcessor.parseXlsxFile(file);
		assertFalse(parsedUsers.isEmpty());
		assertEquals(users.size(), parsedUsers.size());
		assertEquals(users.get(0).getFirstName(), parsedUsers.get(0).getFirstName());
		assertEquals(users.get(1).getEmail(), parsedUsers.get(1).getEmail());

		assertTrue(file.delete());
	}

}
