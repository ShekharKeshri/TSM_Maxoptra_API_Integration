package com.tsm.accesscom.api;

import com.tsm.accesscom.exception.AccessComException;

/**
 * Standard interface for accessing Microsoft Com objects.
 * 
 * @author Juliano Mohr - juliano@theservicemanager.com
 */
public interface AccessCom {

    /**
     * Calls a function of a MS Com object.
     * 
     * @param functionName name of the function.
     * @param parameters parameters required for the function specified.
     * @return Object with response from the function called.
     * @throws com.tsm.accesscom.exception.AccessComException if any error ocurrs during the call.
     */
    public Object callFunction(String functionName,
			       Object... parameters) throws AccessComException;

    /**
     * Calls a function of a MS Com object returning a String.
     * 
     * @param functionName name of the function.
     * @param parameters parameters required for the function specified.
     * @return String with response from the function called.
     * @throws com.tsm.accesscom.exception.AccessComException if any error ocurrs during the call.
     */
    public String callStringFunction(String functionName,
				     Object... parameters) throws AccessComException;

    /**
     * Calls a function of a MS Com object returning a array of byte.
     * 
     * @param functionName name of the function.
     * @param parameters parameters required for the function specified.
     * @return byte[] with response from the function called.
     * @throws com.tsm.accesscom.exception.AccessComException if any error ocurrs during the call.
     */
    public byte[] callByteArrayFunction(String functionName,
					Object... parameters) throws AccessComException;

}
