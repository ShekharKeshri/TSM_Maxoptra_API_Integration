package com.tsm.accesscom.api;

import java.util.Properties;

import com.jacob.com.Dispatch;
import com.tsm.accesscom.exception.AccessComException;

/**
 * Specific implementation for calling functions of TSM application.
 * 
 * @author Juliano Mohr - juliano@theservicemanager.com
 */
public class TsmCom extends Com {

    public static final String ERROR_INITIALIZE_TSM = "Failed to initialize TSM (TSM_Init method).";

    /**
     * Full qualified name of the tsm properties file. If it's under src folder, only the file's name.
     */
    private static final String PROPERTIES_FILE_NAME = "tsm.properties";

    /**
     * Used to keep the properties found in the tsm properties file.
     */
    private static String[] tsmProperties = new String[4];

    /**
     * Indicates if client has initialized TSM or not.
     */
    private Boolean isTSMInitialized = false;

    /**
     * Contains the directory of TSM application. Used for initialization.
     */
    private String tsmDir;

    /**
     * Contains the TSM user's name. Used for initialization.
     */
    private String tsmUser;

    /**
     * Contains the TSM user's password. Used for initialization.
     */
    private String tsmPassword;

    /**
     * Reads the properties file and assigns properties into attributes.
     */
    static {

	Properties properties = Com.loadPropertiesFromFile(PROPERTIES_FILE_NAME);

	tsmProperties[0] = properties.getProperty("tsm.progname");
	tsmProperties[1] = properties.getProperty("tsm.dir");
	tsmProperties[2] = properties.getProperty("tsm.user");
	tsmProperties[3] = properties.getProperty("tsm.passwd");
    }

    /**
     * Creates a new TsmCom object with values from the tsm properties file.
     */
    public TsmCom() {

	super(tsmProperties[0]);

	if (tsmProperties[1] == null || tsmProperties[2] == null || tsmProperties[3] == null) {

	    throw new IllegalArgumentException(
		    "Cannot instantiate object because there is a problem with tsm.properties file.");
	}

	this.tsmDir = tsmProperties[1];
	this.tsmUser = tsmProperties[2];
	this.tsmPassword = tsmProperties[3];
    }

    /**
     * Creates a new TsmCom object with values passed through parameters.
     * 
     * @param programId name of the program in windows register.
     * @param tsmDir directory where TSM is installed.
     * @param tsmUser TSM user's name.
     * @param tsmPassword TSM user's password.
     */
    public TsmCom(String programId,
		  String tsmDir,
		  String tsmUser,
		  String tsmPassword) {
	super(programId);
	this.tsmDir = tsmDir;
	this.tsmUser = tsmUser;
	this.tsmPassword = tsmPassword;
    }

    public String getTsmDir() {
	return tsmDir;
    }

    public void setTsmDir(String tsmDir) {
	this.tsmDir = tsmDir;
    }

    public String getTsmPassword() {
	return tsmPassword;
    }

    public void setTsmPassword(String tsmPassword) {
	this.tsmPassword = tsmPassword;
    }

    public String getTsmUser() {
	return tsmUser;
    }

    public void setTsmUser(String tsmUser) {
	this.tsmUser = tsmUser;
    }

    /**
     * Initiliazes TSM for determined user at determined installation. Runs the TSM_Init method.
     * 
     * @throws com.tsm.accesscom.exception.AccessComException
     */
    public void initializeTSM() throws AccessComException {

	if (!this.isTSMInitialized) {
		//modified by edsel on 6-12-2010 by maria's intructions [BEGIN]
	    //Object response = Dispatch.call(super.getDispatch(), "TSM_Init", tsmUser, tsmPassword, tsmDir, true);		
		Object response = Dispatch.call(super.getDispatch(), "TSM_Init_SS", tsmUser, tsmPassword, tsmDir, "SS");
		//modified by edsel on 6-12-2010 by maria's intructions [END]
	    if (response.toString().equals("true")) {
		this.isTSMInitialized = true;
	    } else {
		Object[] params = {"global.lasterror"};
		String lasterror = "TSM Last Error is: ";
		lasterror = lasterror + Dispatch.callN(super.getDispatch(), "tsm_vfpfunc", params).toString();
		throw new AccessComException(ERROR_INITIALIZE_TSM + " " + lasterror);
	    }
	}
    }

    @Override
    public Object callFunction(String functionName,
			       Object... parameters) throws AccessComException {

	this.initializeTSM();
	return Dispatch.callN(super.getDispatch(), functionName, parameters);
    }

    @Override
    public String callStringFunction(String functionName,
				     Object... parameters) throws AccessComException {

	this.initializeTSM();
	return Dispatch.callN(super.getDispatch(), functionName, parameters).toString();
    }

    @Override
    public byte[] callByteArrayFunction(String functionName,
					Object... parameters) throws AccessComException {

	this.initializeTSM();
	return Dispatch.callN(super.getDispatch(), functionName, parameters).toSafeArray().toByteArray();
    }
}
