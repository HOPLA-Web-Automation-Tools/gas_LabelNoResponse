var scriptName = "Label No Response Emails"
var userProperties = PropertiesService.getUserProperties();

var unrespondedLabel = userProperties.getProperty("unrespondedLabel") || 'No Response',
    ignoreLabel = userProperties.getProperty("ignoreLabel") || 'Ignore No Response',
    minTime = userProperties.getProperty("minTime") || '5d', // 5 days
    maxTime = userProperties.getProperty("maxTime") || '14d', // 14 days
    noResponse_checkFrequency_HOUR = userProperties.getProperty("noResponse_checkFrequency_HOUR") || 1,
    status = userProperties.getProperty("status") || 'disabled';

user_email = Session.getEffectiveUser().getEmail();

function test() {
   var str = "hello";
  Logger.log('I said ' + str);
}




function doGet(e) {
    if (e.parameter.setup) { //SETUP    
        deleteAllTriggers()


        ScriptApp.newTrigger("main_noresponse").timeBased().atHour(noResponse_checkFrequency_HOUR).everyDays(1).create();
        var content = "<p>" + scriptName + " has been installed on your email " + user_email + ". " +
            "It is currently set to label no-response emails every 1AM.</p>" +
            '<p>You can change these settings by clicking the WAT Suite extension icon or WAT Settings on gmail.</p>';        

        return HtmlService.createHtmlOutput(content);
    } else if (e.parameter.test) {
        var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL)
        return HtmlService.createHtmlOutput(authInfo.getAuthorizationStatus());
    } else if (e.parameter.noresponse_saveVariables) { //SET NO RESPONSE VARIABLES
        userProperties.setProperty("unrespondedLabel", e.parameter.nr_unresponsedlabel || unrespondedLabel);
        userProperties.setProperty("ignoreLabel", e.parameter.nr_ignoreLabel || ignoreLabel);
        userProperties.setProperty("minTime", (e.parameter.nr_minTime + "d") || minTime);
        userProperties.setProperty("maxTime", (e.parameter.nr_maxTime + "d") || maxTime);
        userProperties.setProperty("noResponse_checkFrequency_HOUR", (e.parameter.nr_checkFrequency_HOUR) || noResponse_checkFrequency_HOUR);
        userProperties.setProperty("status", (e.parameter.nr_status) || status);
      

        unrespondedLabel = userProperties.getProperty("unrespondedLabel"),
            ignoreLabel = userProperties.getProperty("ignoreLabel"),
            minTime = userProperties.getProperty("minTime"),
            maxTime = userProperties.getProperty("maxTime"),
            noResponse_checkFrequency_HOUR = userProperties.getProperty("noResponse_checkFrequency_HOUR"),
            status = userProperties.getProperty("status");

        deleteAllTriggers()

        if (e.parameter.nr_status == 'enabled'){
          ScriptApp.newTrigger("main_noresponse").timeBased().atHour(noResponse_checkFrequency_HOUR).everyDays(1).create();
        }        
        //return HtmlService.createHtmlOutput("No-Response settings has been saved.");
        return ContentService.createTextOutput("No-Response settings has been saved.");
        //return ContentService.createTextOutput(
        //e.parameters.prefix + '("Finished")')
        //.setMimeType(ContentService.MimeType.JAVASCRIPT);
    } else if (e.parameter.noresponse_trigger) { //DO IT NOW
        var res = main_noresponse();
        var ret = res ? "no_response has finished." : 'Failed with error ['+e+']. Script will try again in 5mins.';
        return ContentService.createTextOutput(ret);
        
        
    } else if (e.parameter.noresponse_disable) { //DISABLE
        userProperties.setProperty("status","disabled")
        deleteAllTriggers()
        return ContentService.createTextOutput("Triggers has been disabled.");
    }else if (e.parameter.noresponse_enable) { //ENABLE      
        userProperties.setProperty("status","enabled")
        deleteAllTriggers();
        ScriptApp.newTrigger("main_noresponse").timeBased().atHour(noResponse_checkFrequency_HOUR).everyDays(1).create();
        return ContentService.createTextOutput("Triggers has been enabled.");
    } else if (e.parameter.noresponse_getVariables) { //GET VARIABLES
        var triggers = ScriptApp.getProjectTriggers();
        var status;
        if (triggers.length != 1) {
            status = 'disabled';
        } else {
            status = 'enabled';
        }
        resjson = {
            'unrespondedLabel': userProperties.getProperty("unrespondedLabel") || 'No Response',
            'ignoreLabel': userProperties.getProperty("ignoreLabel") || 'Ignore No Response',
            'minTime': userProperties.getProperty("minTime") || '5d',
            'maxTime': userProperties.getProperty("maxTime") || '14d',
            'noResponse_checkFrequency_HOUR': userProperties.getProperty("noResponse_checkFrequency_HOUR") || 1,
            'status': status
        }
        return ContentService.createTextOutput(JSON.stringify(resjson));
    } else if (e.parameter.getTriggers) {
        var sTriggers = ''
        var triggers = ScriptApp.getProjectTriggers();
        for (var i = 0; i < triggers.length; i++) {
            sTriggers += triggers[i].getUniqueId() + ' + ';

        }
        Logger.log(sTriggers);
        return HtmlService.createHtmlOutput(sTriggers);
    } else { //NO PARAMETERS
        // use an externally hosted stylesheet
        var style = '<link href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" rel="stylesheet">';

        // get the query "greeting" parameter and set the default to "Hello"
        var greeting = scriptName;
        // get the query "name" parameter and set the default to "World!"
        var name = "has been installed";

        // create and use a template
        var heading = HtmlService.createTemplate('<h1><?= greeting ?> <?= name ?>!</h1>')

        // set the template variables
        heading.greeting = greeting;
        heading.name = name;

        deleteAllTriggers()

        var content = "<p>" + scriptName + " has been installed on your email " + user_email + ". " +
            "It is currently set to label no-response emails every 1AM.</p>" +
            '<p>You can change these settings by clicking the WAT Suite extension icon or WAT Settings on gmail.</p>';  

        ScriptApp.newTrigger("main_noresponse").timeBased().atHour(noResponse_checkFrequency_HOUR).everyDays(1).create();

        var HTMLOutput = HtmlService.createHtmlOutput();
        HTMLOutput.append(style);
        HTMLOutput.append(heading.evaluate().getContent());
        HTMLOutput.append(content);

        return HTMLOutput
    }



}

function deleteAllTriggers() {
    //DELETE ALL TRIGGERS
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
        ScriptApp.deleteTrigger(triggers[i]);
    }
    //DELETE ALL TRIGGERS***
}



/*
 * This script goes through your Gmail Inbox and finds recent emails where you
 * were the last respondent. It applies a nice label to them, so you can
 * see them in Priority Inbox or do something else.
 *
 * To remove and ignore an email thread, just remove the unrespondedLabel and
 * apply the ignoreLabel.
 *
 * This is most effective when paired with a time-based script trigger.
 *
 * For installation instructions, read this blog post:
 * http://jonathan-kim.com/2013/Gmail-No-Response/
 */

// Mapping of Gmail search time units to milliseconds.
var UNIT_MAPPING = {
    h: 36e5, // Hours
    d: 864e5, // Days
    w: 6048e5, // Weeks
    m: 263e7, // Months
    y: 3156e7 // Years
};

var ADD_LABEL_TO_THREAD_LIMIT = 100;
var REMOVE_LABEL_TO_THREAD_LIMIT = 100;

function main_noresponse() {
  try{
    processUnresponded();
  }catch (e){
    ScriptApp.newTrigger("main_noresponse").timeBased().after(5 * 60 * 1000); //5 mins
    return 0;
  }
  cleanUp();
  return 1;
}

function processUnresponded() {
    // Define all the filters.
    var filters = [
        'is:sent',
        'from:me',
        '-in:chats',
        '(-subject:"unsubscribe" AND -"This message was automatically generated by Gmail.")',
        'older_than:' + minTime,
        'newer_than:' + maxTime
    ];

    var threads = GmailApp.search(filters.join(' ')),
        threadMessages = GmailApp.getMessagesForThreads(threads),
        unrespondedThreads = [],
        minTimeAgo = new Date();

    minTimeAgo.setTime(subtract(minTimeAgo, minTime));

    Logger.log('Processing ' + threads.length + ' threads.');

    // Filter threads where I was the last respondent.
    threadMessages.forEach(function(messages, i) {
        var thread = threads[i],
            lastMessage = messages[messages.length - 1],
            lastFrom = lastMessage.getFrom(),
            lastTo = lastMessage.getTo(), // I don't want to hear about it when I am sender and receiver
            lastMessageIsOld = lastMessage.getDate().getTime() < minTimeAgo.getTime();

        if (isMe(lastFrom) && !isMe(lastTo) && lastMessageIsOld && !threadHasLabel(thread, ignoreLabel)) {
            unrespondedThreads.push(thread);
        }
    })

    // Mark unresponded in bulk.
    markUnresponded(unrespondedThreads);
    Logger.log('Updated ' + unrespondedThreads.length + ' messages.');
}

function subtract(date, timeStr) {
    // Takes a date object and subtracts a Gmail-style time string (e.g. '5d').
    // Returns a new date object.
    var re = /^([0-9]+)([a-zA-Z]+)$/,
        parts = re.exec(timeStr),
        val = parts && parts[1],
        unit = parts && parts[2],
        ms = UNIT_MAPPING[unit];

    return date.getTime() - (val * ms);
}

function isMe(fromAddress) {
    var addresses = getEmailAddresses();
    for (i = 0; i < addresses.length; i++) {
        var address = addresses[i],
            r = RegExp(address, 'i');

        if (r.test(fromAddress)) {
            return true;
        }
    }

    return false;
}

function getEmailAddresses() {
    // Cache email addresses to cut down on API calls.
    if (!this.emails) {
        Logger.log('No cached email addresses. Fetching.');
        var me = Session.getActiveUser().getEmail(),
            emails = GmailApp.getAliases();

        emails.push(me);
        this.emails = emails;
        Logger.log('Found ' + this.emails.length + ' email addresses that belong to you.');
    }
    return this.emails;
}

function threadHasLabel(thread, labelName) {
    var labels = thread.getLabels();

    for (i = 0; i < labels.length; i++) {
        var label = labels[i];

        if (label.getName() == labelName) {
            return true;
        }
    }

    return false;
}

function markUnresponded(threads) {
    var label = getLabel(unrespondedLabel);

    // addToThreads has a limit of 100 threads. Use batching.
    if (threads.length > ADD_LABEL_TO_THREAD_LIMIT) {
        for (var i = 0; i < Math.ceil(threads.length / ADD_LABEL_TO_THREAD_LIMIT); i++) {
            label.addToThreads(threads.slice(100 * i, 100 * (i + 1)));
        }
    } else {
        label.addToThreads(threads);
    }
}

function getLabel(labelName) {
    // Cache the labels.
    this.labels = this.labels || {};
    label = this.labels[labelName];

    if (!label) {
        Logger.log('Could not find cached label "' + labelName + '". Fetching.', this.labels);

        var label = GmailApp.getUserLabelByName(labelName);

        if (label) {
            Logger.log('Label exists.');
        } else {
            Logger.log('Label does not exist. Creating it.');
            label = GmailApp.createLabel(labelName);
        }
        this.labels[labelName] = label;
    }
    return label;
}

function cleanUp() {
    var label = getLabel(unrespondedLabel),
        iLabel = getLabel(ignoreLabel),
        threads = label.getThreads(),
        expiredThreads = [],
        expiredDate = new Date();

    expiredDate.setTime(subtract(expiredDate, maxTime));

    if (!threads.length) {
        Logger.log('No threads with that label');
        return;
    } else {
        Logger.log('Processing ' + threads.length + ' threads.');
    }

    threads.forEach(function(thread) {
        var lastMessageDate = thread.getLastMessageDate();

        // Remove all labels from expired threads.
        if (lastMessageDate.getTime() < expiredDate.getTime()) {
            Logger.log('Thread expired');
            expiredThreads.push(thread);
        } else {
            Logger.log('Thread not expired');
        }
    });

    // removeFromThreads has a limit of 100 threads. Use batching.
    if (expiredThreads.length > REMOVE_LABEL_TO_THREAD_LIMIT) {
        for (var i = 0; i < Math.ceil(expiredThreads.length / REMOVE_LABEL_TO_THREAD_LIMIT); i++) {
            label.removeFromThreads(expiredThreads.slice(100 * i, 100 * (i + 1)));
            iLabel.removeFromThreads(expiredThreads.slice(100 * i, 100 * (i + 1)));
        }
    } else {
        label.removeFromThreads(expiredThreads);
        iLabel.removeFromThreads(expiredThreads);
    }

    Logger.log(expiredThreads.length + ' unresponded messages expired.');
}