package com.tsm.accesscom.client;

import jp.ne.so_net.ga2.no_ji.jcom.IDispatch;
import jp.ne.so_net.ga2.no_ji.jcom.ReleaseManager;

import org.jawin.DispatchPtr;
import org.jawin.win32.COINIT;
import org.jawin.win32.Ole32;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.tsm.accesscom.api.AccessCom;
import com.tsm.accesscom.api.TsmCom;

public class ClientTest extends Thread {

    public ClientTest() {
	
    }
    
    public ClientTest(String name) {
	super(name);
    }
    
    private void doTest() {

	Variant result1 = null;
	String result2 = null;
	byte[] result3 = null;

	AccessCom accessCom_01 = new TsmCom();	
	AccessCom accessCom_02 = new TsmCom();

	try {
	    
	    
//	    accessCom_01.callFunction("TSM_INIT_SS", "CRAIG", "", "c:/develop/tsm/TSM68", "LI", null, null, null, null, "apple");
	    System.out.println(accessCom_01.callStringFunction("tsm_vfpfunc", "set('datasession')"));
	    
//	    accessCom_02.callFunction("TSM_INIT_SS", "ADMIN", "", "c:/develop/tsm/TSM68", "LI", null, null, null, null, "banana");
	    System.out.println(accessCom_02.callStringFunction("tsm_vfpfunc", "set('datasession')"));
	    
//	    accessCom_01.getDispatch().
	    
//	    result1 = (Variant) accessCom.callFunction("tsm_read", "invoices", "total");
//	    result2 = accessCom.callStringFunction("tsm_read", "invoices", "total");
//	    result3 = accessCom.callByteArrayFunction("tsm_read", "invoices", "total");
	    
	} catch (Exception e) {
	    e.printStackTrace();
	}

//	System.out.println(result1);
//	System.out.println(result2);
//	System.out.println(result3);
    }
    
    public void doOtherTest() {
	
	
	Dispatch dispatch_01 = new Dispatch("tsm61.global");
	Object[] params = {"ADMIN", "", "C:/Program Files/TSM", "RM", true, true};
	Variant variant = Dispatch.callN(dispatch_01, "TSM_INIT_SS", params);
	System.out.println(variant.toString());
	Object[] param = {"select * from servcard where servcard.Allocated < 22/07/2008", "A"};
	System.out.println(Dispatch.call(dispatch_01, "TSM_Query", param));
	//( cQuery, lToXML)
	//System.out.println(Dispatch.call(dispatch_01, "tsm_vfpfunc", "global.test").toString());
	//System.out.println(Dispatch.call(dispatch_01, "tsm_vfpfunc", "set('datasession')").toString());
	
//	Dispatch dispatch_02 = new Dispatch("tsm61.global");
//	Object[] params_02 = {"ADMIN", "", "C:/Program Files/TSM", "LI", null, null, null, null, "banana"};
//	Dispatch.callN(dispatch_02, "TSM_INIT_SS", params_02);
//	System.out.println(Dispatch.call(dispatch_01, "tsm_vfpfunc", "global.test").toString());
//	System.out.println(Dispatch.call(dispatch_02, "tsm_vfpfunc", "set('datasession')").toString());
//	System.out.println(Dispatch.call(dispatch_01, "tsm_vfpfunc", "set('datasession')").toString());
//	
//	Dispatch dispatch_03 = new Dispatch("tsm61.global");
//	Object[] params_03 = {"ADMIN", "", "C:/Program Files/TSM", "LI", null, null, null, null, "ORANGE"};
//	Dispatch.callN(dispatch_03, "TSM_INIT_SS", params_03);
//	System.out.println(Dispatch.call(dispatch_01, "tsm_vfpfunc", "global.test").toString());
//	System.out.println(Dispatch.call(dispatch_02, "tsm_vfpfunc", "set('datasession')").toString());
//	System.out.println(Dispatch.call(dispatch_01, "tsm_vfpfunc", "set('datasession')").toString());
//	System.out.println(Dispatch.call(dispatch_03, "tsm_vfpfunc", "set('datasession')").toString());
    }
    
    public void run()  {
	
	System.out.println(getName());
	
	try {
	    doJaWinTest();
	} catch (Exception e) {
	    // TODO Auto-generated catch block
	    e.printStackTrace();
	}
    }

    public void doOneMoreTest() throws Exception {

	ReleaseManager releaseManager = new ReleaseManager();
	IDispatch iDispatch = new IDispatch(releaseManager, "tsm61.global");
	
	Object[] params = {"CRAIG", "", "c:/develop/tsm/TSM68", "LI", null, null, null, null, "apple"};
	iDispatch.invoke("TSM_INIT_SS", params);
	System.out.println(iDispatch.invoke("tsm_vfpfunc", new Object[] {"global.test"}));
	
	ReleaseManager releaseManager_02 = new ReleaseManager();
	IDispatch iDispatch_02 = new IDispatch(releaseManager_02, "tsm61.global");
	
	Object[] params_02 = {"ADMIN", "", "c:/develop/tsm/TSM68", "LI", null, null, null, null, "banana"};
	iDispatch_02.invoke("TSM_INIT_SS", params_02);
	System.out.println(iDispatch.invoke("tsm_vfpfunc", new Object[] {"global.test"}));
    }
    
    public synchronized void doJaWinTest() throws Exception {
	
	Ole32.CoInitialize(COINIT.MULTITHREADED);
	
	DispatchPtr dispatchPtr = new DispatchPtr("tsm61.global");
	
	Object[] params = {getName(), "", "c:/develop/tsm/TSM68", "LI", null, null, null, null, String.valueOf(Math.random() * 100)};
	dispatchPtr.invokeN("TSM_INIT_SS", params);
	System.out.println(dispatchPtr.invoke("tsm_vfpfunc", "global.test"));
	System.out.println(dispatchPtr.invoke("tsm_vfpfunc", "set('datasession')"));
		
	Ole32.CoUninitialize();
//	
//	DispatchPtr dispatchPtr_02 = new DispatchPtr("tsm61.global");
//	
//	Object[] params_02 = {"ADMIN", "", "c:/develop/tsm/TSM68", "LI", null, null, null, null, "banana"};
//	dispatchPtr_02.invokeN("TSM_INIT_SS", params);
//	System.out.println(dispatchPtr.invoke("tsm_vfpfunc", "global.test"));
//	System.out.println(dispatchPtr_02.invoke("tsm_vfpfunc", "global.test"));
	
//	Ole32.CoUninitialize();
	
    }
    
    public static void main(String[] args) throws Exception{

	ClientTest clientTest = new ClientTest();
//	clientTest.doTest();
	clientTest.doOtherTest();
//	clientTest.doOneMoreTest();
//	clientTest.doJaWinTest();
	
//	Thread thread = new ClientTest("ADMIN");
//	thread.start();
//	Thread thread_02 = new ClientTest("CRAIG");
//	thread_02.start();
    }

}
