package com.transunion.wf.exception;

public class MailActionsException extends RuntimeException {

    private static final long serialVersionUID = 1L;

    public MailActionsException(String message) {
    	super(message);
    }
    
    public MailActionsException(Exception cause, String message) {
    	super(message, cause);
    }
   
}
