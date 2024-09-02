function main() {
    const date = new Date();
    const logTimestamp = date.toISOString().split('Z')[0].replace('T', ' ');
    const logPrefix = logTimestamp + " DEBUG " + logger;

    try {
        if (document) {
            log(`Processing document Node: ${document.nodeRef}`);

            if (document.hasAspect("cm:emailed")) {
                handleEmailedDocument(document);
            }
        } else {
            log("Document is null or undefined.");
        }
    } catch (error) {
        logError("An error occurred", error);
    }
}

function handleEmailedDocument(document) {
    const subjectLine = document.properties['cm:subjectline'];
    log(`Received a Mail Body Document with subject: ${subjectLine}`);

    const [actualSubject, dateSubject] = parseSubjectLine(subjectLine);

    updateDocumentProperties(document, actualSubject, dateSubject);
    updateEmailAddresses(document);

    document.save();
    log("Document saved successfully.");
}

function parseSubjectLine(subjectLine) {
    const [actualSubject, dateSubject] = subjectLine.split("##").map(part => part.trim());
    return [
        cleanText(actualSubject),
        cleanText(dateSubject)
    ];
}

function updateDocumentProperties(document, actualSubject, dateSubject) {
    const cleanedSubject = cleanText(actualSubject);
    
    document.properties['cm:subjectline'] = cleanedSubject;
    document.properties['cm:title'] = cleanedSubject;

    const fileName = `${cleanedSubject}${dateSubject}`;
    document.properties['cm:name'] = cleanFileName(fileName);

    log(`Updated fileName to: ${fileName}`);
}

function updateEmailAddresses(document) {
    const addresseesList = document.properties['cm:addressees'].split("##").map(part => part.trim());
    const allRecipients = cleanText(addresseesList[0]);
    const emailTO = cleanText(addresseesList[1]);
    const emailFrom = cleanText(addresseesList[2]);

    document.properties['cm:addressee'] = emailTO;
    log(`Updated emailTO to: ${emailTO}`);

    document.properties['cm:originator'] = emailFrom;
    log(`Updated emailFrom to: ${emailFrom}`);

    document.properties['cm:addressees'] = allRecipients;
    log(`Updated all recipients to: ${allRecipients}`);
}

function cleanText(text) {
    return text.normalize('NFKD')
        .replace(/[\u2014\u2013\u002d\u2010\u2011\u2012\u2015\uFE58\uFE63\uFF0D\u2E3A\u2E3B\uFFFC\u0300-\u036f]/g, '')
        .trim();
}

function cleanFileName(fileName) {
    return fileName.normalize('NFKD')
        .replace(/[/\\?%*:|"<>]/g, '')
        .trim();
}

function log(message) {
    logger.system.out(logTimestamp + " DEBUG " + logger + " " + message);
}

function logError(message, error) {
    logger.system.out(logTimestamp + " ERROR " + logger + " " + message + ": " + error);
}

main();
