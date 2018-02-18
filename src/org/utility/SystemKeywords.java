package org.utility;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;

import com.crestech.opkey.plugin.KeywordLibrary;
import com.crestech.opkey.plugin.communication.contracts.functionresult.FunctionResult;
import com.crestech.opkey.plugin.communication.contracts.functionresult.Result;
import com.crestech.opkey.plugin.utility.Validations;
import com.crestech.opkey.plugin.utility.exceptionHandler.ArgumentDataInvalidException;
import com.crestech.opkey.plugin.utility.exceptionHandler.ArgumentDataMissingException;

public class SystemKeywords implements KeywordLibrary {
	public FunctionResult Method_pressKeys(String argument) throws AWTException, ArgumentDataMissingException, ArgumentDataInvalidException {
		Validations.checkDataForBlank(0);
		Validations.checkDataForWhiteSpace(0);
		String Key = null;
		Robot robot = new Robot();
		char s = argument.charAt(0);

		switch (s) {

		case '{':
			Key = argument.substring(1, argument.indexOf("}")).toLowerCase();
			new SystemKeywords().processKeys(Key);
			break;

		case '%':
			Key = argument.substring(2, argument.indexOf("}")).toLowerCase();
			robot.keyPress(KeyEvent.VK_ALT);
			new SystemKeywords().processKeys(Key);
			robot.keyRelease(KeyEvent.VK_ALT);
			break;

		case '^':
			Key = argument.substring(2, argument.indexOf("}")).toLowerCase();
			robot.keyPress(KeyEvent.VK_CONTROL);
			new SystemKeywords().processKeys(Key);
			robot.keyRelease(KeyEvent.VK_CONTROL);
			break;

		case '+':
			Key = argument.substring(2, argument.indexOf("}")).toLowerCase();
			robot.keyPress(KeyEvent.VK_SHIFT);
			new SystemKeywords().processKeys(Key);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;

		default:
			StringSelection stringSelection = new StringSelection(argument);
			Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
			clipboard.setContents(stringSelection, stringSelection);
			robot.keyPress(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_V);
			robot.keyRelease(KeyEvent.VK_V);
			robot.keyRelease(KeyEvent.VK_CONTROL);
			System.out.println("default operation");
		}
		return Result.PASS().setOutput(true).make();

	}

	private void processKeys(String Key) throws AWTException {

		Robot robot = new Robot();
		switch (Key) {
		case "esc":
			robot.keyPress(KeyEvent.VK_ESCAPE);
			robot.keyRelease(KeyEvent.VK_ESCAPE);
			System.out.println("Escape pressed");
			break;
		case "tab":
			robot.keyPress(KeyEvent.VK_TAB);
			robot.keyRelease(KeyEvent.VK_TAB);
			System.out.println("Tab pressed");
			break;
		case "capslock":
			robot.keyPress(KeyEvent.VK_CAPS_LOCK);
			robot.keyRelease(KeyEvent.VK_CAPS_LOCK);
			System.out.println("capslock pressed");
			break;
		case "end":
			robot.keyPress(KeyEvent.VK_END);
			robot.keyRelease(KeyEvent.VK_END);
			System.out.println("end pressed");
			break;
		case "home":
			robot.keyPress(KeyEvent.VK_HOME);
			robot.keyRelease(KeyEvent.VK_HOME);
			System.out.println("Home pressed");
			break;
		case "ins":
		case "insert":
			robot.keyPress(KeyEvent.VK_INSERT);
			robot.keyRelease(KeyEvent.VK_INSERT);
			System.out.println("Insert pressed");
			break;
		case "pgdn":
			robot.keyPress(KeyEvent.VK_PAGE_DOWN);
			robot.keyRelease(KeyEvent.VK_PAGE_DOWN);
			System.out.println("PageDown pressed");
			break;
		case "pgup":
			robot.keyPress(KeyEvent.VK_PAGE_UP);
			robot.keyRelease(KeyEvent.VK_PAGE_UP);
			System.out.println("PageUp pressed");
			break;
		case "del":
		case "delete":
			robot.keyPress(KeyEvent.VK_DELETE);
			robot.keyRelease(KeyEvent.VK_DELETE);
			System.out.println("Delete pressed");
			break;
		case "left":
			robot.keyPress(KeyEvent.VK_LEFT);
			robot.keyRelease(KeyEvent.VK_LEFT);
			System.out.println("Left Key pressed");
			break;
		case "right":
			robot.keyPress(KeyEvent.VK_RIGHT);
			robot.keyRelease(KeyEvent.VK_RIGHT);
			System.out.println("Right pressed");
			break;
		case "up":
			robot.keyPress(KeyEvent.VK_UP);
			robot.keyRelease(KeyEvent.VK_UP);
			System.out.println("UpKey pressed");
			break;
		case "down":
			robot.keyPress(KeyEvent.VK_DOWN);
			robot.keyRelease(KeyEvent.VK_DOWN);
			System.out.println("DownKey pressed");
			break;
		case "numlock":
			robot.keyPress(KeyEvent.VK_NUM_LOCK);
			robot.keyRelease(KeyEvent.VK_NUM_LOCK);
			System.out.println("Numlock pressed");
			break;
		case "prtsc":
			robot.keyPress(KeyEvent.VK_PRINTSCREEN);
			robot.keyRelease(KeyEvent.VK_PRINTSCREEN);
			System.out.println("PrintScreen pressed");
			break;
		case "scrolllock":
			robot.keyPress(KeyEvent.VK_SCROLL_LOCK);
			robot.keyRelease(KeyEvent.VK_SCROLL_LOCK);
			System.out.println("Home pressed");
			break;
		case "f1":
			robot.keyPress(KeyEvent.VK_F1);
			robot.keyRelease(KeyEvent.VK_F1);
			System.out.println("F1 pressed");
			break;
		case "f2":
			robot.keyPress(KeyEvent.VK_F2);
			robot.keyRelease(KeyEvent.VK_F2);
			System.out.println("F2 pressed");
			break;
		case "f3":
			robot.keyPress(KeyEvent.VK_F3);
			robot.keyRelease(KeyEvent.VK_F3);
			System.out.println("F3 pressed");
			break;
		case "f4":
			robot.keyPress(KeyEvent.VK_F4);
			robot.keyRelease(KeyEvent.VK_F4);
			System.out.println("F4 pressed");
			break;
		case "f5":
			robot.keyPress(KeyEvent.VK_F5);
			robot.keyRelease(KeyEvent.VK_F5);
			System.out.println("F5 pressed");
			break;
		case "f6":
			robot.keyPress(KeyEvent.VK_F6);
			robot.keyRelease(KeyEvent.VK_F6);
			System.out.println("F6 pressed");
			break;
		case "f7":
			robot.keyPress(KeyEvent.VK_F7);
			robot.keyRelease(KeyEvent.VK_F7);
			System.out.println("F7 pressed");
			break;
		case "f8":
			robot.keyPress(KeyEvent.VK_F8);
			robot.keyRelease(KeyEvent.VK_F8);
			System.out.println("F8 pressed");
			break;
		case "f9":
			robot.keyPress(KeyEvent.VK_F9);
			robot.keyRelease(KeyEvent.VK_F9);
			System.out.println("F9 pressed");
			break;
		case "f10":
			robot.keyPress(KeyEvent.VK_F10);
			robot.keyRelease(KeyEvent.VK_F10);
			System.out.println("F10 pressed");
			break;
		case "f11":
			robot.keyPress(KeyEvent.VK_F11);
			robot.keyRelease(KeyEvent.VK_F11);
			System.out.println("F11 pressed");
			break;
		case "f12":
			robot.keyPress(KeyEvent.VK_F12);
			robot.keyRelease(KeyEvent.VK_F12);
			System.out.println("F12 pressed");
			break;
		case "f13":
			robot.keyPress(KeyEvent.VK_F13);
			robot.keyRelease(KeyEvent.VK_F13);
			System.out.println("F13 pressed");
			break;
		case "f14":
			robot.keyPress(KeyEvent.VK_F14);
			robot.keyRelease(KeyEvent.VK_F14);
			System.out.println("F14 pressed");
			break;
		case "f15":
			robot.keyPress(KeyEvent.VK_F15);
			robot.keyRelease(KeyEvent.VK_F15);
			System.out.println("F15 pressed");
			break;
		case "f16":
			robot.keyPress(KeyEvent.VK_F16);
			robot.keyRelease(KeyEvent.VK_F16);
			System.out.println("F16 pressed");
			break;
		case "add":
			robot.keyPress(KeyEvent.VK_ADD);
			robot.keyRelease(KeyEvent.VK_ADD);
			System.out.println("Keyboard Add pressed");
			break;
		case "subtract":
			robot.keyPress(KeyEvent.VK_SUBTRACT);
			robot.keyRelease(KeyEvent.VK_SUBTRACT);
			System.out.println("Keyboard Subtract pressed");
			break;
		case "multiply":
			robot.keyPress(KeyEvent.VK_MULTIPLY);
			robot.keyRelease(KeyEvent.VK_MULTIPLY);
			System.out.println("Keyboard Multiply pressed");
			break;
		case "divide":
			robot.keyPress(KeyEvent.VK_DIVIDE);
			robot.keyRelease(KeyEvent.VK_DIVIDE);
			System.out.println("Keyboard Divide pressed");
			break;
		case "break":
			robot.keyPress(KeyEvent.VK_PAUSE);
			robot.keyRelease(KeyEvent.VK_PAUSE);
			System.out.println("BreakPause Key pressed");
			break;
		case "enter":
		case "~":
			robot.keyPress(KeyEvent.VK_ENTER);
			robot.keyRelease(KeyEvent.VK_ENTER);
			System.out.println("Enter Key pressed");
			break;
		case "a":
			robot.keyPress(KeyEvent.VK_A);
			robot.keyRelease(KeyEvent.VK_A);
			System.out.println("A pressed");
			break;
		case "b":
			robot.keyPress(KeyEvent.VK_B);
			robot.keyRelease(KeyEvent.VK_B);
			System.out.println("B pressed");
			break;
		case "c":
			robot.keyPress(KeyEvent.VK_C);
			robot.keyRelease(KeyEvent.VK_C);
			System.out.println("C pressed");
			break;
		case "d":
			robot.keyPress(KeyEvent.VK_D);
			robot.keyRelease(KeyEvent.VK_D);
			System.out.println("D pressed");
			break;
		case "e":
			robot.keyPress(KeyEvent.VK_E);
			robot.keyRelease(KeyEvent.VK_E);
			System.out.println("D pressed");
			break;
		case "f":
			robot.keyPress(KeyEvent.VK_F);
			robot.keyRelease(KeyEvent.VK_F);
			System.out.println("E pressed");
			break;
		case "g":
			robot.keyPress(KeyEvent.VK_G);
			robot.keyRelease(KeyEvent.VK_G);
			System.out.println("F pressed");
			break;
		case "h":
			robot.keyPress(KeyEvent.VK_H);
			robot.keyRelease(KeyEvent.VK_H);
			System.out.println("H pressed");
			break;
		case "i":
			robot.keyPress(KeyEvent.VK_I);
			robot.keyRelease(KeyEvent.VK_I);
			System.out.println("I pressed");
			break;
		case "j":
			robot.keyPress(KeyEvent.VK_J);
			robot.keyRelease(KeyEvent.VK_J);
			System.out.println("J pressed");
			break;
		case "k":
			robot.keyPress(KeyEvent.VK_K);
			robot.keyRelease(KeyEvent.VK_K);
			System.out.println("K pressed");
			break;
		case "l":
			robot.keyPress(KeyEvent.VK_L);
			robot.keyRelease(KeyEvent.VK_L);
			System.out.println("L pressed");
			break;
		case "m":
			robot.keyPress(KeyEvent.VK_M);
			robot.keyRelease(KeyEvent.VK_M);
			System.out.println("M pressed");
			break;
		case "n":
			robot.keyPress(KeyEvent.VK_N);
			robot.keyRelease(KeyEvent.VK_N);
			System.out.println("N pressed");
			break;
		case "o":
			robot.keyPress(KeyEvent.VK_O);
			robot.keyRelease(KeyEvent.VK_O);
			System.out.println("O pressed");
			break;
		case "p":
			robot.keyPress(KeyEvent.VK_P);
			robot.keyRelease(KeyEvent.VK_P);
			System.out.println("P pressed");
			break;
		case "q":
			robot.keyPress(KeyEvent.VK_Q);
			robot.keyRelease(KeyEvent.VK_Q);
			System.out.println("Q pressed");
			break;
		case "r":
			robot.keyPress(KeyEvent.VK_R);
			robot.keyRelease(KeyEvent.VK_R);
			System.out.println("R pressed");
			break;
		case "s":
			robot.keyPress(KeyEvent.VK_S);
			robot.keyRelease(KeyEvent.VK_S);
			System.out.println("S pressed");
			break;
		case "t":
			robot.keyPress(KeyEvent.VK_T);
			robot.keyRelease(KeyEvent.VK_T);
			System.out.println("T pressed");
			break;
		case "u":
			robot.keyPress(KeyEvent.VK_U);
			robot.keyRelease(KeyEvent.VK_U);
			System.out.println("U pressed");
			break;
		case "v":
			robot.keyPress(KeyEvent.VK_V);
			robot.keyRelease(KeyEvent.VK_V);
			System.out.println("V pressed");
			break;
		case "w":
			robot.keyPress(KeyEvent.VK_W);
			robot.keyRelease(KeyEvent.VK_W);
			System.out.println("W pressed");
			break;
		case "x":
			robot.keyPress(KeyEvent.VK_X);
			robot.keyRelease(KeyEvent.VK_X);
			System.out.println("X pressed");
			break;
		case "y":
			robot.keyPress(KeyEvent.VK_Y);
			robot.keyRelease(KeyEvent.VK_Y);
			System.out.println("Y pressed");
			break;
		case "z":
			robot.keyPress(KeyEvent.VK_Z);
			robot.keyRelease(KeyEvent.VK_Z);
			System.out.println("Z pressed");
			break;
		default:
			System.out.println("Invalid Key pressed...");

		}

	}

}
