package com.tsm.accesscom.api;

import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.tsm.accesscom.exception.AccessComException;

/**
 * Standard implementation for calling Microsoft Com objects. It basically uses Jacob API behind de scenes. See:
 * http://sourceforge.net/projects/jacob-project/
 * 
 * @author Juliano Mohr - juliano@theservicemanager.com
 */
public class Com implements AccessCom {

    /**
     * Name of the program the client wishes to connect. It has to be registered on Windows register.
     */
    private String programId;

    /**
     * Jacob Dispatch object used to make the calls.
     */
    private Dispatch dispatch;

    /**
     * Creates a new connection with the Com object.
     * 
     * @param programId name of the program the client wishes to connect.
     */
    public Com(String programId) {

	if (programId == null) {
	    throw new IllegalArgumentException(
		    "Cannot instantiate object because programId is required. Check the properties file.");
	}

	this.programId = programId;
	this.loadJacobDispatchObject();
    }

    public Dispatch getDispatch() {
	return dispatch;
    }

    public void setDispatch(Dispatch dispatch) {
	this.dispatch = dispatch;
    }

    public String getProgramId() {
	return programId;
    }

    public void setProgramId(String programId) {
	this.programId = programId;
    }

    /**
     * Calls a function of a MS Com object using Jacob API.
     * 
     * @param functionName name of the function.
     * @param parameters parameters required for the function specified.
     * @return Object with response from the function called.
     * @throws com.tsm.accesscom.exception.AccessComException if any error ocurrs during the call.
     */
    public Object callFunction(String functionName,
			       Object... parameters) throws AccessComException {

	return Dispatch.callN(this.getDispatch(), functionName, parameters);
    }

    /**
     * Calls a function of a MS Com object returning a String.
     * 
     * @param functionName name of the function.
     * @param parameters parameters required for the function specified.
     * @return String with response from the function called.
     * @throws com.tsm.accesscom.exception.AccessComException if any error ocurrs during the call.
     */
    public String callStringFunction(String functionName,
				     Object... parameters) throws AccessComException {

	return Dispatch.callN(this.getDispatch(), functionName, parameters).toString();
    }

    /**
     * Calls a function of a MS Com object returning a array of byte.
     * 
     * @param functionName name of the function.
     * @param parameters parameters required for the function specified.
     * @return byte[] with response from the function called.
     * @throws com.tsm.accesscom.exception.AccessComException if any error ocurrs during the call.
     */
    public byte[] callByteArrayFunction(String functionName,
					Object... parameters) throws AccessComException {

	return Dispatch.callN(this.getDispatch(), functionName, parameters).toSafeArray().toByteArray();
    }

    /**
     * Creates a Dispatch object and assigns it to this instance of Com class.
     */
    protected void loadJacobDispatchObject() {

	ActiveXComponent activeXComponent = new ActiveXComponent(this.programId);
	this.dispatch = activeXComponent.getObject();
    }

    /**
     * Reads the properties file with the specified name.
     * 
     * @param fileName full qualified name of the file. Pass only the file's name if file is under src folder.
     * @return java.util.Properties file.
     */
    protected static Properties loadPropertiesFromFile(String fileName) {

	InputStream inputStream = Com.class.getClassLoader().getResourceAsStream(fileName);
	Properties properties = new Properties();
	try {
	    if (inputStream != null) {
		properties.load(inputStream);
	    }
	} catch (IOException e) {
	    e.printStackTrace();
	} finally {
	    try {
		if (inputStream != null) {
		    inputStream.close();
		}
	    } catch (IOException e) {
		e.printStackTrace();
	    }
	}
	return properties;
    }
}
