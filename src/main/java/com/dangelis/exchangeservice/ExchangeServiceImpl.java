package com.dangelis.exchangeservice;

import java.net.URI;
import java.net.URISyntaxException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import microsoft.exchange.webservices.data.core.response.AttendeeAvailability;
import microsoft.exchange.webservices.data.property.complex.availability.CalendarEvent;
import org.springframework.stereotype.Repository;

import com.dangelis.entity.Appointment;
import com.dangelis.exchangeservice.exception.AppointmentException;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.availability.AvailabilityData;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.misc.error.ServiceError;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.misc.availability.AttendeeInfo;
import microsoft.exchange.webservices.data.misc.availability.GetUserAvailabilityResults;
import microsoft.exchange.webservices.data.misc.availability.TimeWindow;
@Repository
public class ExchangeServiceImpl implements ExchangeServices {
	
	private static ExchangeServiceImpl exchangeImpl;
	private static ExchangeService service;
	static {

		
       

		service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
		try {
			System.setProperty("javax.net.ssl.trustStore","C:\\Program Files\\Java\\jdk1.8.0_111\\jre\\lib\\security\\cacerts");
			service.setUrl(new URI("https://mail.altran.com/EWS/Exchange.asmx"));
		} catch (URISyntaxException e) {

			e.printStackTrace();
		}
		ExchangeCredentials credentials = new WebCredentials("dangelis", "", "e");
		service.setCredentials(credentials);
	}
	 
	 public static ExchangeServiceImpl getInstance() {
		 return exchangeImpl;
	 }
	 
	 /**
	     * Initialize the Exchange Credentials.
	     * Don't forget to replace the "USRNAME","PWD","DOMAIN_NAME" variables.
	 * @throws URISyntaxException 
	     */
	    public ExchangeServiceImpl(){

	       
	    }	


	

	public List<Appointment> getAllAppointmentsByEmailByDay(String email,String dateIni,String dateFinal) throws AppointmentException, ParseException {
		service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
		try {
			System.setProperty("javax.net.ssl.trustStore","C:\\Program Files\\Java\\jdk1.8.0_111\\jre\\lib\\security\\cacerts");
			service.setUrl(new URI("https://mail.altran.com/EWS/Exchange.asmx"));
		} catch (URISyntaxException e) {

			e.printStackTrace();
		}
		ExchangeCredentials credentials = new WebCredentials("dangelis", "", "");
		service.setCredentials(credentials);
	    	List<Appointment>list=new ArrayList<Appointment>();
	    	List<AttendeeInfo> attendees = new ArrayList<AttendeeInfo>();
	    	attendees.add(new AttendeeInfo(email));
		SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-ddHH:mm:ss");

		Date d1 = format.parse( dateIni);

		Date d2 =format.parse( dateFinal);

		
		try {
			d1 = format.parse(dateIni);
			d2=format.parse(dateFinal);
			Calendar c = Calendar.getInstance();
			c.setTime(d2);
			c.add(Calendar.DATE, 1);  // number of days to add
			d2 =c.getTime();
		} catch (ParseException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		
		GetUserAvailabilityResults results = null;
		try {
			results = service.getUserAvailability(
					attendees,
					new TimeWindow(d1, d2),
					AvailabilityData.FreeBusyAndSuggestions);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

			// Output attendee availability information.
			int attendeeIndex = 0;

			for (AttendeeAvailability attendeeAvailability : results.getAttendeesAvailability()) {
				System.out.println("Availability for " + attendees.get(attendeeIndex));
				if (attendeeAvailability.getErrorCode() == ServiceError.NoError) {
					for (CalendarEvent calendarEvent : attendeeAvailability.getCalendarEvents()) {
						System.out.println("Calendar event");
						System.out.println("  Start time: " + calendarEvent.getStartTime().toString());
						System.out.println("  End time: " + calendarEvent.getEndTime().toString());

						if (calendarEvent.getDetails() != null)
						{
							System.out.println("  Subject: " + calendarEvent.getDetails().getSubject());
							// Output additional properties.
						}
						
						com.dangelis.entity.Appointment appointment=new com.dangelis.entity.Appointment();
						appointment.setSubject(calendarEvent.getDetails().getSubject());
						appointment.setDateStart(dateToLocalDateTime(calendarEvent.getStartTime()));
						appointment.setDateEnd(dateToLocalDateTime(calendarEvent.getEndTime()));
						list.add(appointment);
					}
				}

				attendeeIndex++;
			}
			
		return list;	
		
	}
	
	
	private  LocalDateTime dateToLocalDateTime(Date date) {
	    return LocalDateTime.ofInstant(date.toInstant(), ZoneId.systemDefault());
	}

	

}
