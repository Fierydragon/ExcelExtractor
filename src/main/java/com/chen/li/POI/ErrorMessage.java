package com.chen.li.POI;

public class ErrorMessage {
	private int line;
	private String message;
	
	public ErrorMessage(int line, String message) {
		super();
		this.line = line;
		this.message = message;
	}
	
	public int getLine() {
		return line;
	}
	public void setLine(int line) {
		this.line = line;
	}
	public String getMessage() {
		return message;
	}
	public void setMessage(String message) {
		this.message = message;
	}
	@Override
	public String toString() {
		return "line: " + line + ", message: " + message;
	}
}
