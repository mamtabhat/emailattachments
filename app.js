var Imap = require("imap");
var MailParser = require("mailparser").MailParser;
var Promise = require("bluebird");
var fs = require("fs");
var path = require('path');
var base64  = require('base64-stream');
setInterval(function(){
    var imapConfig = {
      user: 'brahmendra@wings4talent.com',
      password: 'Brahmi13@123',
      host: 'outlook.office365.com',
      port: 993,
     tls: true
  };
  
  var imap = new Imap(imapConfig);
  //Promise.promisifyAll(imap);
  
  imap.once("ready", execute);
  imap.once("error", function(err) {
    console.error("Connection error: " + err.stack);
  });
  
  
  imap.connect();
  
  function execute() {
      imap.openBox("Resumes", false, function(err, mailBox) {
          if (err) {
            console.error("Folder opening error",err);
              return;
          }
  
          let delay = 1 * 24 * 3600 * 1000,
          yesterday = new Date();
          yesterday.setTime(Date.now() - delay);
      
  
          imap.search(['UnSeen', ['SINCE', yesterday]], function(err, results) {
              if (err) {
                console.error("Read mail error",err);
                  return;
              }
              console.log("results", results);


              if(results.length > 0){
                var f = imap.fetch(results, { bodies: "" , struct: true });
                 f.on("message", processMessage);
                 f.once("error", function(err) {
                     return Promise.reject(err);
                 });
              }
               
          });
      });
  }
  
  
  function processMessage(msg, seqno) {
      console.log("Processing msg #" + seqno);
       console.log("mag",msg);
       var prefix = '(#' + seqno + ') ';
       var parser = new MailParser();
  
        parser.on("headers", function(headers) {
          console.log("Header: " + JSON.stringify(headers));
       });
       
  
        msg.once('attributes', function(attrs) {
  
  
          var attachments = findAttachmentParts(attrs.struct);
         
  
          for (var i = 0, len=attachments.length ; i < len; ++i) {
              var attachment = attachments[i];
  
              console.log('Fetching attachment %s', attachment.params.name);
  
               
              var f = imap.fetch(attrs.uid , { //do not use imap.seq.fetch here
                  bodies: [attachment.partID],
                  struct: true
                });
                console.log("uid", attrs.uid);
              
                console.log("attachment", attachment);

                console.log("attachment", typeof attachment);
              
                //build function to process attachment message
                f.on('message', buildAttMessageFunction(attachment, attrs.uid));
  
               
  
          }
  
          
  
        });
        
        
  }
  
  function toUpper(thing) { return thing && thing.toUpperCase ? thing.toUpperCase() : thing;}
  
  function findAttachmentParts(struct, attachments) {
      attachments = attachments ||  [];
      for (var i = 0, len = struct.length, r; i < len; ++i) {
        if (Array.isArray(struct[i])) {
          findAttachmentParts(struct[i], attachments);
        } else {
          if (struct[i].disposition && ['ATTACHMENT'].indexOf(toUpper(struct[i].disposition.type)) > -1) {
            attachments.push(struct[i]);
          }
        }
      }
      return attachments;
    }
  
  
    function buildAttMessageFunction(attachment, uid) {
      var filename = attachment.params.name;
      var encoding = attachment.encoding;
      
      return function (msg, seqno) {
        var prefix = '(#' + seqno + ') ';
        
          
        msg.on('body', function(stream, info) {
  
         
  
         
           var writeStream = fs.createWriteStream('./Resumes/'+filename);
   
           // console.log("writeStream", writeStream);
  
           
           
  
          
           if (toUpper(encoding) === 'BASE64') {
           
  
          
             var file = filename;
  
             var str = file.includes("docx");
  
             console.log("str111111", str);
  
          stream.pipe(new base64.Base64Decode()).pipe(writeStream);
  
          
          imap.addFlags(uid, 'seen', function (err) {
        });

            
    //      imap.move(uid, '[Outlook]/Deleted', (err) => {
    //       if (err)
    //           console.log('Move email to trash ', err);
    //  });
  
  
            
                  // mammoth.extractRawText({path: filename})
                  // .then(function(result){
                  //     var text = result.value; // The raw text 
              
                  //     //this prints all the data of docx file
                  //    // console.log(text);
        
                  //     fs.writeFileSync(path.join(__dirname,"Resumes",filename), text,"UTF8");
              
                  // })
                  // .done();
                
                      
               
  
                  //fs.writeFileSync(path.join(__dirname,"Resumes",filename), data,"UTF8");
                  
                 
  
            } 
          
          
      });
              
  
          msg.once('end', function() {
            console.log(prefix + 'Finished email');
            //imap.end();
          });
       
      };
    }
  
  
  }, 60000 * 30);
 