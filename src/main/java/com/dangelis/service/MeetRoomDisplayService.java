package com.dangelis.service;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.dangelis.entity.Appointment;
import com.dangelis.exchangeservice.ExchangeServiceImpl;
import com.dangelis.exchangeservice.exception.AppointmentException;

import java.text.ParseException;
import java.util.List;


@Service
public class MeetRoomDisplayService  {

	  @Autowired
	    ExchangeServiceImpl exchangeService;
	  
	  public List<Appointment> getAllAppointmentsByEmail(String email,String dateIni,String dateFinal) throws AppointmentException, ParseException {
		 return exchangeService.getAllAppointmentsByEmailByDay(email,dateIni,dateFinal);
			
		}
}
