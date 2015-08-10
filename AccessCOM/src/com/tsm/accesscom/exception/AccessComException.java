package com.tsm.accesscom.exception;

public class AccessComException extends Exception {

    private static final long serialVersionUID = 1L;

    public AccessComException(String message) {
	super(message);
    }

    public AccessComException(Throwable cause) {
	super(cause);
    }

    public AccessComException(String message,
			      Throwable cause) {
	super(message, cause);
    }

}
