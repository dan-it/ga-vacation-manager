/*
 * Basic idea from https://productforums.google.com/forum/#!topic/calendar/8ryyCllffqI
 * 
 * with help from https://github.com/ClaudiaJ/gas-calendar-accept/blob/master/Code.gs
 *
 */

function statusFilters(status) {
  if (status instanceof Array) {
    return function(invite) {
      return status.includes(invite.getMyStatus());
    }
  } else {
    return function(invite) {
      return invite.getMyStatus() === status;
    }
  }
}

function processInvites() {
  const calendarId =  'crilomazzo.org_vbs93lemfma90fi59f0j4mrm5o@group.calendar.google.com'; // this needs to be the email address of the calendar you're monitoring
  const invited = "INVITED";
  const accepted = "YES";
  const accept = CalendarApp.GuestStatus.YES;
  const reject = CalendarApp.GuestStatus.NO;
  const rejection = "Richiesta ferie rifiutata"; //subject line for our email to reject a booking
  const future = 2; //Number of years to look ahead
  const defaultMax = 2; //Max number of people allowed on vacation
  const maxDelim = 'MAX: ';
 
  var calendar = CalendarApp.getCalendarById(calendarId);
 
  var start = new Date();
  var end = new Date(start);
  end.setFullYear(end.getFullYear()+future); //move end date by "future" years
  var invites = calendar.getEvents(start, end).filter(statusFilters(CalendarApp.GuestStatus.INVITED)); //find all future invites
  //var invites = calendar.getEvents(start, end, {statusFilters:[CalendarApp.GuestStatus.INVITED]});
  
  for(var i = 0; i < invites.length; i++){
    Logger.log("Processing: "+invites[i].getTitle());
    Logger.log(invites[i].getStartTime()+invites[i].getEndTime());
    //Process conflicts day by day
    for (var checkDate = invites[i].getStartTime() ; checkDate.valueOf() < invites[i].getEndTime().valueOf() && invites[i].getMyStatus()==CalendarApp.GuestStatus.INVITED; checkDate.setDate(checkDate.getDate()+1)) {
      if (checkDate.valueOf() > invites[i].getStartTime().valueOf()) { checkDate.setHours(0,0,0,0); }
      //find max for this day
      var dayMaxEvent = calendar.getEventsForDay(checkDate, {search: maxDelim});
      var dayMaxNum;
      switch (dayMaxEvent.length) {
        case 0 : 
          dayMaxNum = defaultMax -1;
          break;
        case 1 : 
          dayMaxNum = dayMaxEvent[0].getTitle().slice(maxDelim.length) -1;
          break;
        default : //more than one MAX per day.
          dayMaxNum = 0;
      }
      var endCheckDate = invites[i].getEndTime();
      endCheckDate.setSeconds(-1); //move to day before in case ends at midnight.
      var startCheckDate = new Date(checkDate);
      endCheckDate.setHours(0,0,0,0);
      startCheckDate.setHours(0,0,0,0);
      var lastDay = false;
      if (startCheckDate.valueOf() == endCheckDate.valueOf()) {
        endCheckDate = invites[i].getEndTime();
        startCheckDate = checkDate;
        lastDay = true;
      } else {
        endCheckDate.setDate(startCheckDate.getDate()+1)
      }
      startCheckDate = checkDate;
      
      var conflicts = calendar.getEvents(startCheckDate, endCheckDate).filter(statusFilters(CalendarApp.GuestStatus.YES));
        if(conflicts.length>dayMaxNum){
          Logger.log("Found a potential conflict to: " + invites[i].getTitle());
          Logger.log("Creator is: " + invites[i].getCreators());
          var body =
              "La richiesta per il perido da"+
                "<b>"+invites[i].getStartTime()+"</b>"+
                  "<br>a<br><b>"+
                    invites[i].getEndTime()+"</b><br><br>"+
                      "è stata rifiutata automaticamente.";
          MailApp.sendEmail(invites[i].getCreators(), rejection, "", {htmlBody: body})      
          invites[i].setMyStatus(reject);
        } else {
          if (lastDay) {
            Logger.log("No conflict, accepting: " + invites[i].getTitle());
            invites[i].setMyStatus(accept);
            invites[i].setTitle(invites[i].getCreators())
          }
        }
      }
    }
  }
