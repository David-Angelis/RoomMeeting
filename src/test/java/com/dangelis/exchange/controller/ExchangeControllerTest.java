package com.dangelis.exchange.controller;

import java.text.ParseException;
import java.util.ArrayList;
import java.util.List;

import com.dangelis.entity.Room;
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
import com.dangelis.service.MeetRoomDisplayService;

import junit.framework.Assert;

@RunWith(PowerMockRunner.class)
@PrepareForTest(MeetRoomDisplayService.class)
public class ExchangeControllerTest extends EasyMockSupport {
	
	ExchangeController intance;
	
	@Mock
	MeetRoomDisplayService serviceMock;
	
	@Before
	public void init() {
		intance=new ExchangeController();
		PowerMock.resetAll();
		PowerMock.createStrictMock(MeetRoomDisplayService.class);
		
	}
	
	@Test
	public void test001() throws AppointmentException, ParseException {
		Whitebox.setInternalState(intance, "service", serviceMock);
		List<Appointment>list=new ArrayList<Appointment>();
		list.add(new Appointment());
		
		EasyMock.expect(serviceMock.getAllAppointmentsByEmail("123","123","123")).andReturn(list);
		PowerMock.replayAll();
		Room result=intance.getAllAppointmentsByEmail("123","123","123");
		PowerMock.verifyAll();
		Assert.assertEquals(1, result.getList().size());
	}
	
	@Test
	public void test002() throws AppointmentException, ParseException {
		Whitebox.setInternalState(intance, "service", serviceMock);
		List<Appointment>list=new ArrayList<Appointment>();
		list.add(new Appointment());
		
		EasyMock.expect(serviceMock.getAllAppointmentsByEmail("123","123","123")).andThrow(new AppointmentException());
		PowerMock.replayAll();
		Room result;
		result = intance.getAllAppointmentsByEmail("123","123","123");
		PowerMock.verifyAll();
		Assert.assertEquals(0, result.getList().size());
	}

}
