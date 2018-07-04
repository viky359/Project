package com.project.parkmobile;

import java.io.IOException;
import org.testng.annotations.Test;
import com.project.Execution.Execute;

public class Driver {
	@Test
	public void DriverExecution() throws InterruptedException, IOException {
		try {
			Execute ex = new Execute();
			ex.startExecution();
		} catch (Exception e) {
			System.out.println(e.getStackTrace());
		}
	}
}
