package com.leadprogram;

import java.io.IOException;

public class FunctionalatiesExecuter extends Launcher {

	public static void main(String[] args) throws InterruptedException, IOException {
		System.out.println("Automation started");
		Launcher.launch();
		LeadProgram.leadProgram();
	}
}
