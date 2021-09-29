function myFunction(e) {
var sheet = SpreadsheetApp.getActiveSheet();
var row =  SpreadsheetApp.getActiveSheet().getLastRow();
var name = e.values[1];
var emaillowercase = e.values[3];
var email = emaillowercase.toLowerCase(); 
var currentzip = e.values[4];
var Time = e.values[9];
var StartDate = e.values[8];
var date = new Date(e.values[8]);
var level = e.values[5];
var startTime = new Date( StartDate + " " + Time);
var months = ["January", "February", "march", "April", "May", "June", "July", "August", "September", "October", "November", "December"]; 
var days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
var datevisit = days[date.getDay()] + ", " + date.getDate() + " " + months[date.getMonth()] + " " + date.getFullYear();
var time = e.values[9];
var titleEvent = ("School Visit " + level + " with " + name);
CalendarApp.getCalendarById('c_s5q1v17u6t1o005crn4ac8kj0k@group.calendar.google.com').createEvent(titleEvent, startTime, startTime, {sendNotifications: true}).addGuest(email).addEmailReminder(30);
var subject = "MHIS Visit Appointment";
var body = "Assalamu’alaykum warahmatullahi wabarakatuh,<br><br>Dear Mr/Mrs " + name + " Thank you for filling  Book a School Visit form. <br><br>Here we share with you, the schedule that you have chosen :<br>Date	 	: " + datevisit + "<br>Time		: " + time + "<br><br>You will be contacted through WhatsApp by our Admission Team about the confirmation of your appointment. When the schedule is confirmed, you will receive a reminder via Google Calendar one day before the date and one hour before the time. If you wish to reschedule your visit, please inform us at least 1x24 hours in advance. <br><br>For the health safety of everyone, we would like to remind again about the strict protocols while in the school area:<br><ol><li>Always wear your mask (double mask is a must)</li><li>Check the temperature before going inside (below 37.3)</li><li>Wash hands with hand sanitizer</li><li>Always maintain physical distancing (1.5 - 2 meters)</li></ol><br><br>For further information, you may contact us through +62 813 8908 1220 or +62 821 2546 9320 at office hours from 8 a.m - 3 p.m.<br><br>Looking forward to having a talk with you.<br><br>Wassalamualaykum warahmatullahi wabarakatuh<br><br>Warm regards, <br>Mutiara Harapan Islamic School";
var Geocodingcurrent = UrlFetchApp.fetch("https://maps.googleapis.com/maps/api/geocode/json?address=" + currentzip +"+Indonesia&key=AIzaSyDW0XaRzzDDMgUdPOXvPYrKej__-b6Cby4");
var jsoncurrent = Geocodingcurrent.getContentText();
var datacurrent = JSON.parse(jsoncurrent);
var filtered_kelurahan = datacurrent.results[0].address_components.filter(function(address_component){
    return address_component.types.includes("administrative_area_level_4");
      })
var filtered_kecamatan = datacurrent.results[0].address_components.filter(function(address_component){
    return address_component.types.includes("administrative_area_level_3");
      })
var filtered_kota = datacurrent.results[0].address_components.filter(function(address_component){
    return address_component.types.includes("administrative_area_level_2");
      })
var kota = filtered_kota.length ? filtered_kota[0].long_name: "";
var kecamatan = filtered_kecamatan.length ? filtered_kecamatan[0].long_name: "";
var kelurahan = filtered_kelurahan.length ? filtered_kelurahan[0].long_name: "";
var setkelurahan = sheet.getRange(row,15).setValue(kelurahan);
var setkecamatan = sheet.getRange(row,14).setValue(kecamatan);
var setkota = sheet.getRange(row,13).setValue(kota);
var options = {};
        options.htmlBody = body;
        GmailApp.sendEmail("" +email+ "", 
                   subject,
                   "Please open this email with support HTML",
                   options
                  );   
}
