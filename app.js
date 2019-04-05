var express = require('express'),
    app = express(),
    bodyParser = require('body-parser'),
    officeGen = require('officegen'),
    nodemailer = require('nodemailer'),
    async = require ( 'async' ),
    fs = require('fs'),
    path = require('path');

var themeXml = fs.readFileSync ( path.resolve ( __dirname, 'receipt.xml' ), 'utf8' );
var smtpTransport = nodemailer.createTransport({
    service: 'Gmail',
    auth: {
        type: 'OAuth2',
        user: 'chenweinberg@gmail.com',
        clientId: '695987177749-4aod00ughphctfqouqdn98dn372g1n5o.apps.googleusercontent.com',
        clientSecret: 'Ag2EGt01gZG1SNpJEtY6LU9o',
        refreshToken: '1/8x5rt7dWT7PLclRHowjpNVRfP1R164dho6z8ai6HbkA',
        accessToken: 'ya29.GluABXFWmYm5nOxCsoAtWyeu8KNDg3H8nGhOa_RAhpUyZU2fkJVdT4b6dNAYkFCG9DCdSq1XYAazry5QLCtGbzmzFKpRKwQ5y53R59_d4XoUZJ5zjQEp57kVfq8P',
        expires: 3000 
    }
});
app.set('view engine','ejs');
app.use(express.static('public'));
app.use(bodyParser.urlencoded({extended:true}));


app.get('/', function(req, res){
    res.render('home');
});


app.post('/acu-receipt', function(req, res){
    var docx = officeGen ({
        'type': 'docx',
        'orientation': 'portrait',
        'subject': 'Neshama Therapy Receipt',
        'description': 'Neshama Therapy Receipt'
    });
    docx.on('error', function(err){
        console.log(err);
        res.redirect('/');
    });
    var pObj = docx.createP();
    pObj.options.align = 'left';
    pObj.addImage ( path.resolve(__dirname, 'neshamalogo.png' ) );
    pObj.addLineBreak ();
    pObj.addLineBreak ();
    pObj.addText('2 Carlton St. Suite 720 Toronto, ON M5B 1J3', {bold:true});
    pObj.addLineBreak ();
    pObj.addText('Chen Weinberg, R. Ac, RMT', {bold:true});
    pObj.addLineBreak ();
    pObj.addText('Acupuncture Registration Number: 2295');
    pObj.addLineBreak ();
    pObj.addText('College of Traditional Chinese Medicine and Acupuncturists of Ontario');
    pObj.addLineBreak ();
    pObj.addLineBreak ();
    pObj.addHorizontalLine ();
    pObj.addLineBreak ();
    pObj.addLineBreak ();
    pObj.addText('ACUPUNTURE RECEIPT', {color:'07421e', bold: true, font_size: 14 });
    pObj.addLineBreak ();
    pObj.addText('Patient Name: ' + req.body.patient_name, {color:'06a2db', bold: true, font_size: 18 });
    pObj.addLineBreak ();
    pObj.addText('Date: ' + req.body.date, {color:'204903', bold: true});
    pObj.addLineBreak ();
    pObj.addText('Duration of Treatment: ' + req.body.service_length, {color:'204903', bold: true});
    pObj.addLineBreak ();
    pObj.addText('Description Of Service: ' + req.body.service_desc, {color:'204903', bold: true});
    pObj.addLineBreak ();
    pObj.addText('Total Amount: ' + req.body.amount, {color:'204903', bold: true});
    pObj.addLineBreak ();
    pObj.addLineBreak ();
    pObj.addLineBreak ();
    pObj.addText('Sincerely,');
    pObj.addLineBreak ();
    pObj.addText('Chen Weinberg, R. Ac, RMT', {color:'10350f', bold: true});
    pObj.addLineBreak ();
    pObj.addImage ( path.resolve(__dirname, 'signature.JPG' ) );

    var out = fs.createWriteStream ( 'tmp/receipt.docx' );
    out.on ( 'error', function ( err ) {
        return res.render('error', {error: err});
    });
    var mailOptions = {
        from: "Neshama Therapy <chenweinberg@gmail.com>", // sender address
        to: req.body.patient_email, // list of receivers
        subject: "Your Acupuncture Receipt For Today's Treatment " + req.body.date, // Subject line
        text: "Thank You For Choosing Neshama Therapy, Your Receipt Is Attached With This Email. ", // plaintext body
        attachments: [
            {
                filename: 'your-receipt.docx',
                path: __dirname + '/tmp/receipt.docx'
                // content: Buffer.from('/tmp/receipt.docx', 'base64'),
                // contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            }
        ]
    }
    var mailOptions2 = {
        from: "Neshama Therapy <chenweinberg@gmail.com>", // sender address
        to: 'chenweinberg@gmail.com', // list of receivers
        subject: "You Sent An Acupuncture Receipt to " + req.body.patient_email, // Subject line
        text: "You sent an acupuncture receipt to " + req.body.patient_email + ' on ' + req.body.date, // plaintext body
        attachments: [
            {
                filename: 'your-receipt.docx',
                path: __dirname + '/tmp/receipt.docx'
                // content: Buffer.from('/tmp/receipt.docx', 'base64'),
                // contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            }
        ]
    }
    async.parallel ([
        function ( done ) {
            out.on ( 'close', function () {
                smtpTransport.sendMail(mailOptions, function(error, response){
                    if(error){
                        res.render('error', {error: error});
                    }else{
                        res.render('error', {error: 'Successfully sent mail and downloaded file.'})
                    }
                });
                smtpTransport.sendMail(mailOptions2, function(error, response){
                    if(error){
                        res.render('error', {error: error});
                    }else{
                        res.send("Finished sending mail");
                    }
                    smtpTransport.close(); 
                });
                done ( null );
            });
            docx.generate ( out );
        }
    ], function ( err ) {
        if ( err ) {
            res.send( 'Error creating the word document: ' + err );
        }
    });

});


app.post('/rmt-receipt', function(req, res){
    var docx = officeGen ({
        'type': 'docx',
        'orientation': 'portrait',
        'subject': 'Neshama Therapy Receipt',
        'description': 'Neshama Therapy Receipt'
    });
    docx.on('error', function(err){
        console.log('Error: ' + err);
        res.redirect('/');
    });
    var pObj = docx.createP();
    pObj.options.align = 'left';
    pObj.addImage ( path.resolve(__dirname, 'neshamalogo.png' ) );
    pObj.addLineBreak ();
    pObj.addLineBreak ();
    pObj.addText('2 Carlton St. Suite 720 Toronto, ON M5B 1J3', {bold:true});
    pObj.addLineBreak ();
    pObj.addText('Chen Weinberg, R. Ac, RMT', {bold:true});
    pObj.addLineBreak ();
    pObj.addText('RMT Registration number: M658');
    pObj.addLineBreak ();
    pObj.addText('College of Massage Therapists of Ontario');
    pObj.addLineBreak ();
    pObj.addLineBreak ();
    pObj.addHorizontalLine ();
    pObj.addText('RMT RECEIPT', {color:'07421e', bold: true, font_size: 14 });
    pObj.addLineBreak ();
    pObj.addText('Patient Name: ' + req.body.patient_name, {color:'06a2db', bold: true, font_size: 18 });
    pObj.addLineBreak ();
    pObj.addText('Date: ' + req.body.date, {color:'204903', bold: true});
    pObj.addLineBreak ();
    pObj.addText('Duration of Treatment: ' + req.body.service_length, {color:'204903', bold: true});
    pObj.addLineBreak ();
    pObj.addText('Description Of Service: ' + req.body.service_desc, {color:'204903', bold: true});
    pObj.addLineBreak ();
    pObj.addText('Total Amount: ' + req.body.amount, {color:'204903', bold: true});
    pObj.addLineBreak ();
    pObj.addLineBreak ();
    pObj.addLineBreak ();
    pObj.addText('Issued By');
    pObj.addLineBreak ();
    pObj.addText('Chen Weinberg, R. Ac, RMT', {color: '10350f', bold: true});
    pObj.addLineBreak ();
    pObj.addImage ( path.resolve(__dirname, 'signature.JPG' ) );

    var out = fs.createWriteStream ( 'tmp/receipt.docx' );
    out.on ( 'error', function ( err ) {
        res.send( 'Error creating writestream out: \n' + err );
    });
    var mailOptions = {
        from: "Neshama Therapy <chenweinberg@gmail.com>", // sender address
        to: req.body.patient_email, // list of receivers
        subject: "Your RMT Receipt For Today's Treatment " + req.body.date, // Subject line
        text: "Thank You For Choosing Neshama Therapy, Your Receipt Is Attached With This Email. ", // plaintext body
        attachments: [
            {
                filename: 'your-receipt.docx',
                path: __dirname + '/tmp/receipt.docx'
                // content: Buffer.from('/tmp/receipt.docx', 'base64'),
                // contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            }
        ]
    }
    var mailOptions2 = {
        from: "Neshama Therapy <chenweinberg@gmail.com>", // sender address
        to: 'chenweinberg@gmail.com, katie_acting@live.ca', // list of receivers
        subject: "You sent an RMT receipt to " + req.body.patient_email, // Subject line
        text: "You sent an RMT receipt to " + req.body.patient_email + ' on ' + req.body.date, // plaintext body
        attachments: [
            {
                filename: 'your-receipt.docx',
                path: __dirname + '/tmp/receipt.docx'
                // content: Buffer.from('/tmp/receipt.docx', 'base64'),
                // contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            }
        ]
    }
    async.parallel ([
        function ( done ) {
            out.on ( 'close', function () {
                smtpTransport.sendMail(mailOptions, function(error, response){
                    if(error){
                        console.log(error);
                        // res.render('error', {error: error});
                    }else{
                        res.send('completed')
                    }
                });
                smtpTransport.sendMail(mailOptions2, function(error, response){
                    if(error){
                        return res.render('error', {error: error});
                    }else{
                        res.send('completed')
                    }
                });
                smtpTransport.close();
                done ( null );
            });
            docx.generate ( out );
        }
    ], function ( err ) {
        if ( err ) {
            res.send( 'Error creating the word document: ' + err );
        } // Endif.
    });
});


app.get('/download', function(req, res) {
    res.download('./tmp/receipt.docx');
});

app.listen(process.env.PORT || 3000, process.env.IP, function(){
    console.log('<-------- OfficeGenNeshama running on port 3000!!! -------->')
});