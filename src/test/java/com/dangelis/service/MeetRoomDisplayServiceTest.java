package com.dangelis.service;

import java.text.ParseException;
import java.util.ArrayList;
import java.util.List;

import org.easymock.EasyMock;
import org.easymock.EasyMockSupport;
import org.easymock.Mock;
import org.junit.Before;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.mockito.internal.util.reflection.Whitebox;
import org.powermock.api.easymock.PowerMock;
import org.powermock.core.classloader.annotations.PrepareForTest;
import org.powermock.modules.junit4.PowerMockRunner;

import com.dangelis.entity.Appointment;
import com.dangelis.exchangeservice.ExchangeServiceImpl;
import com.dangelis.exchangeservice.exception.AppointmentException;

import junit.framework.Assert;

@RunWith(PowerMockRunner.class)
@PrepareForTest(ExchangeServiceImpl.class)
public class MeetRoomDisplayServiceTest extends EasyMockSupport {
	MeetRoomDisplayService intance;
	
	@Before
	public void init() {
		intance=new MeetRoomDisplayService();
		PowerMock.resetAll();
		PowerMock.createStrictMock(ExchangeServiceImpl.class);
		
	}
	
	@Mock
	ExchangeServiceImpl serviceMock;
	
	@Test
	public void test01() throws AppointmentException, ParseException {
		Whitebox.setInternalState(intance, "exchangeService", serviceMock);
		EasyMock.expect(serviceMock.getAllAppointmentsByEmailByDay("123","123","123")).andReturn(new ArrayList<Appointment>());
		PowerMock.replayAll();
		List<Appointment>result=intance.getAllAppointmentsByEmail("123","123","123");
		Assert.assertEquals(new ArrayList<Appointment>().isEmpty(), result.isEmpty());
		PowerMock.verifyAll();
	}
	
	

}
