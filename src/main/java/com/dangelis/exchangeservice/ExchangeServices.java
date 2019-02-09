package com.dangelis.exchangeservice;

import java.text.ParseException;
import java.util.List;

import com.dangelis.entity.Appointment;
import com.dangelis.exchangeservice.exception.AppointmentException;

public interface ExchangeServices {

	public List<Appointment> getAllAppointmentsByEmailByDay(String email,String dateInit,String dateFinal) throws AppointmentException, ParseException;
}
