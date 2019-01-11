var ews = require('ews-javascript-api');

module.exports = function(RED) {

    function nodeRedEWS(config) {
      RED.nodes.createNode(this, config);
      var node = this;
      node.name = config.ews;
      node.email = config.email;
      node.password = config.password;
      node.ewsUri = config.ewsUri;
  
      this.on('input', function (msg) {
        var EWSFlowNode = RED.nodes.getNode(node.name);
      
        // exit if empty credentials
        if (EWSFlowNode == null || EWSFlowNode.credentials == null) {
          node.warn('EWS credentials are missing.');
          return;
        }
        
        //get credentials and URI for Exchange Web Services from config
        var email = EWSFlowNode.credentials.email;
        var password = EWSFlowNode.credentials.password;
        var ewsUri = EWSFlowNode.credentials.ewsUri;
        
        var service = new ews.ExchangeService(ews.ExchangeVersion.Exchange2013);
        service.Credentials = new ews.WebCredentials(email, password); //authorize
        service.Url = new ews.Uri(ewsUri);

        var folderId = new ews.FolderId(ews.WellKnownFolderName.Calendar, new ews.Mailbox(msg.user||email, "SMTP"));  //if null msg.user on input, then get appointments for email address set in config
        
        var view = new ews.CalendarView(ews.DateTime.Now.Date.Add(Number(msg.daynum)||0, "day"), ews.DateTime.Now.Date.Add((Number(msg.daynum)||0)+1, "day"));  // appointments om selected day. 
        service.FindAppointments(folderId, view).then((response) => {
          let appointments = response.Items;

          let appointmentsResult = [];

          for (var i in appointments)
            {
              let appointment = appointments[i];  
              
              let app={};

              app.Subject=appointment.Subject;
              app.DisplayTo=appointment.DisplayTo;
              app.DisplayCc=appointment.DisplayCc;
              app.Location=appointment.Location;
              app.StartDate=appointment.Start.toString();
              app.EndDate=appointment.End.toString();
              app.StartDateFormat=appointment.Start.Format("DD-MM-YYYY");
              app.EndDateFormat=appointment.End.Format("DD-MM-YYYY");
              app.StartTime=appointment.Start.Format("hh:mm:ss");
              app.EndTime=appointment.End.Format("hh:mm:ss");

              appointmentsResult.push(app);
            }
          msg.appointments = appointmentsResult;
          node.send([msg, null]);

}, function (error) {
  node.error(error);
})

});
}
  
    RED.nodes.registerType('nodeRedEWS', nodeRedEWS);
  
    function nodeRedEWSSettings(n) {
      RED.nodes.createNode(this, n);
    }
  
    RED.nodes.registerType('nodeRedEWS-access', nodeRedEWSSettings, {
      credentials: {
        email: {
          type: 'text'
        },
        password: {
          type: 'password'
        },
        ewsUri: {
          type: 'text'
        }
      }
    });
  };